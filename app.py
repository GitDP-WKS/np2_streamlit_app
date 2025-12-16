import re
from io import BytesIO
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="ЭЗС — обращения (2025)", layout="wide")

# ======== НАСТРОЙКИ ========
DEFAULT_SHEET_ID = "1YN_8UtrZMqOTYZHaLzczwkkfocD-sS_wKrlSBmn-S50"
DEFAULT_GIDS: List[str] = [
    "880054222","290665501","1707951068","1280453214","1898471504","1456377749",
    "100006210","1678514560","1664238791","1022163523","824830115","2075524941"
]

# Колонки заводов строго такие:
PLANTS = ["E-Prom", "NSP", "Другое"]  # "Другое" = нет принадлежности к заводу / пусто / не распознано

# ======== ВСПОМОГАТЕЛЬНЫЕ ========
def gsheets_csv_url(sheet_id: str, gid: str) -> str:
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip()).lower()

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df

def pick_col(columns: List[str], candidates: List[str]) -> Optional[str]:
    if not columns:
        return None
    m = {norm(c): c for c in columns}
    for cand in candidates:
        k = norm(cand)
        if k in m:
            return m[k]
    for cand in candidates:
        k = norm(cand)
        for nk, orig in m.items():
            if k and k in nk:
                return orig
    return None

def vendor_to_plant(v: str) -> str:
    """Строго: E-Prom / NSP / Другое."""
    s = norm(v)
    if not s or s == "nan":
        return "Другое"
    # NSP
    if "nsp" in s or "нсп" in s:
        return "NSP"
    # E-Prom
    if "e-prom" in s or "eprom" in s or "e prom" in s or "е-пром" in s or "епром" in s:
        return "E-Prom"
    return "Другое"

def parse_dt_smart(df: pd.DataFrame, col_date: str, col_time: Optional[str]) -> pd.Series:
    if col_time:
        s = df[col_date].astype(str) + " " + df[col_time].astype(str)
    else:
        s = df[col_date].astype(str)
    dt1 = pd.to_datetime(s, dayfirst=True, errors="coerce")
    ok1 = float(dt1.notna().mean())
    if ok1 >= 0.5:
        return dt1
    dt2 = pd.to_datetime(s, dayfirst=False, errors="coerce")
    ok2 = float(dt2.notna().mean())
    return dt2 if ok2 > ok1 else dt1

@st.cache_data(ttl=900, show_spinner=False)
def load_gid(sheet_id: str, gid: str) -> pd.DataFrame:
    url = gsheets_csv_url(sheet_id, gid)
    return pd.read_csv(url, on_bad_lines="skip")

def load_all(sheet_id: str, gids: List[str]) -> Tuple[pd.DataFrame, List[str]]:
    frames: List[pd.DataFrame] = []
    errors: List[str] = []
    for gid in gids:
        try:
            dfi = load_gid(sheet_id, gid)
            dfi["_source_gid"] = gid
            frames.append(dfi)
        except Exception as e:
            errors.append(f"GID {gid}: {e}")
    if not frames:
        return pd.DataFrame(), errors
    out = pd.concat(frames, ignore_index=True, sort=False)
    return out, errors

def add_totals(df: pd.DataFrame, row_name: str = "Итого") -> pd.DataFrame:
    df2 = df.copy()
    df2[row_name] = df2.sum(axis=1)
    total_row = pd.DataFrame(df2.sum(axis=0)).T
    total_row.index = [row_name]
    return pd.concat([df2, total_row], axis=0)

# ======== UI ========
st.title("📊 ЭЗС — обращения (2025)")

with st.sidebar:
    st.header("Источник (Google Sheets)")
    sheet_id = st.text_input("Sheet ID", value=DEFAULT_SHEET_ID)
    gids_text = st.text_area("GID листов (через запятую)", value=",".join(DEFAULT_GIDS), height=100)
    st.caption("Таблица должна быть расшарена: “Anyone with the link → Viewer”.")
    if st.button("🔄 Сбросить кэш"):
        st.cache_data.clear()
        st.toast("Кэш очищен.", icon="✅")

gids = [g.strip() for g in str(gids_text).split(",") if g.strip()]
raw, errors = load_all(sheet_id, gids)

if errors:
    st.warning("Некоторые листы не прочитались — продолжил с теми, что удалось.")
    with st.expander("Ошибки по листам"):
        st.code("\n".join(errors))

if raw.empty:
    st.error("Не удалось прочитать данные ни с одного листа. Проверь доступ и Sheet ID/GID.")
    st.stop()

df = normalize_columns(raw)

# ======== определяем колонки ========
cols = df.columns.tolist()
auto_date = pick_col(cols, ["Дата обращения", "Дата", "Date"])
auto_time = pick_col(cols, ["Время обращения", "Время", "Time"])
auto_reason = pick_col(cols, ["Причина обращения", "Причина", "Problem", "Причина/тема"])
auto_station = pick_col(cols, ["Номер ЭЗС", "ЭЗС", "Station", "Станция"])
auto_vendor = pick_col(cols, ["Производитель станции", "Производитель", "Завод", "Vendor"])
auto_note = pick_col(cols, ["Примечание", "Комментарий", "Note"])

with st.expander("🛠️ Сопоставление колонок (если авто не угадал)"):
    c1, c2 = st.columns(2)
    with c1:
        col_date = st.selectbox("Колонка ДАТА", options=cols, index=(cols.index(auto_date) if auto_date in cols else 0))
        col_time = st.selectbox("Колонка ВРЕМЯ (можно пусто)", options=["— нет —"] + cols,
                                index=(1 + cols.index(auto_time) if auto_time in cols else 0))
        col_reason = st.selectbox("Колонка ПРИЧИНА (как в таблице)", options=cols,
                                  index=(cols.index(auto_reason) if auto_reason in cols else 0))
    with c2:
        col_station = st.selectbox("Колонка НОМЕР ЭЗС (можно пусто)", options=["— нет —"] + cols,
                                   index=(1 + cols.index(auto_station) if auto_station in cols else 0))
        col_vendor = st.selectbox("Колонка ПРОИЗВОДИТЕЛЬ/ЗАВОД", options=["— нет —"] + cols,
                                  index=(1 + cols.index(auto_vendor) if auto_vendor in cols else 0))
        col_note = st.selectbox("Колонка ПРИМЕЧАНИЕ (можно пусто)", options=["— нет —"] + cols,
                                index=(1 + cols.index(auto_note) if auto_note in cols else 0))
    st.dataframe(df.head(10), use_container_width=True)

col_time = None if col_time == "— нет —" else col_time
col_station = None if col_station == "— нет —" else col_station
col_vendor = None if col_vendor == "— нет —" else col_vendor
col_note = None if col_note == "— нет —" else col_note

if not col_date or col_date not in df.columns:
    st.error("Не выбрана колонка даты.")
    st.stop()
if not col_reason or col_reason not in df.columns:
    st.error("Не выбрана колонка причины.")
    st.stop()

df["_dt"] = parse_dt_smart(df, col_date, col_time)
if df["_dt"].isna().all():
    st.error("Не удалось распарсить даты (все значения пустые). Проверь формат даты/времени.")
    st.stop()

# ======== Завод (строго 3 значения) ========
if col_vendor and col_vendor in df.columns:
    df["Завод"] = df[col_vendor].astype(str).apply(vendor_to_plant)
else:
    df["Завод"] = "Другое"

# Диагностика: какие значения реально пришли в колонке производителя/завода
with st.expander("🔎 Диагностика завода"):
    if col_vendor and col_vendor in df.columns:
        vc = df[col_vendor].astype(str).value_counts().head(30).reset_index()
        vc.columns = ["Значение в исходной колонке", "Строк"]
        st.dataframe(vc, use_container_width=True, hide_index=True)
    st.write("После нормализации (Завод):")
    st.dataframe(df["Завод"].value_counts().reset_index().rename(columns={"index":"Завод","Завод":"Строк"}), use_container_width=True, hide_index=True)

# ======== ФИЛЬТРЫ (2025 по умолчанию) ========
st.subheader("Фильтры")
f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.4, 1.2])

df_2025 = df[df["_dt"].dt.year == 2025].copy()
if df_2025.empty:
    st.warning("В данных нет 2025 года (или даты не распарсились как 2025). Показываю всё, что есть.")
    df_2025 = df.copy()

min_d = df_2025["_dt"].min().date()
max_d = df_2025["_dt"].max().date()

with f1:
    period_mode = st.radio("Период", ["Весь 2025", "Месяц", "Диапазон"], horizontal=False)
with f2:
    month = st.selectbox("Месяц", [f"2025-{m:02d}" for m in range(1, 13)], index=0) if period_mode == "Месяц" else None
with f3:
    if period_mode == "Диапазон":
        start_date = st.date_input("С даты", value=min_d)
        end_date = st.date_input("По дату", value=max_d)
    else:
        start_date, end_date = None, None
with f4:
    plant_filter = st.multiselect("Завод", options=PLANTS, default=["E-Prom","NSP"])

fdf = df_2025.copy()
if period_mode == "Месяц" and month:
    fdf = fdf[fdf["_dt"].dt.to_period("M").astype(str) == month]
elif period_mode == "Диапазон" and start_date and end_date:
    start = pd.to_datetime(start_date)
    end = pd.to_datetime(end_date) + pd.Timedelta(days=1)
    fdf = fdf[(fdf["_dt"] >= start) & (fdf["_dt"] < end)]

if plant_filter:
    fdf = fdf[fdf["Завод"].isin(plant_filter)]

# ======== KPI ========
k1, k2, k3, k4 = st.columns(4)
total = int(len(fdf))
uniq_station = int(fdf[col_station].nunique()) if col_station else 0
k1.metric("Обращений", total)
k2.metric("Уникальных ЭЗС", uniq_station if col_station else "—")
k3.metric("E-Prom", int((fdf["Завод"] == "E-Prom").sum()) if total else 0)
k4.metric("NSP", int((fdf["Завод"] == "NSP").sum()) if total else 0)

st.divider()

# ======== 1) Причина x Завод (строго 3 колонки) ========
st.markdown("### Разбивка по причинам × завод (E-Prom / NSP)")
tab = pd.crosstab(
    fdf[col_reason].fillna("—").astype(str),
    fdf["Завод"].fillna("Другое").astype(str),
    dropna=False,
)

# гарантируем строго 3 колонки и порядок
for p in PLANTS:
    if p not in tab.columns:
        tab[p] = 0
tab = tab[PLANTS]
tab = add_totals(tab, row_name="Итого")

view_reason = tab.reset_index().rename(columns={"index":"Причина", col_reason:"Причина"})
st.dataframe(view_reason, use_container_width=True, hide_index=True)

st.divider()

# ======== 2) Помесячно 2025 (по всем листам) ========
st.markdown("### Все обращения по месяцам 2025 (по всем листам)")
all_2025 = df[df["_dt"].dt.year == 2025].copy()
pr = pd.period_range("2025-01", "2025-12", freq="M")
monthly = (
    all_2025.dropna(subset=["_dt"])
           .groupby(all_2025["_dt"].dt.to_period("M"))
           .size()
           .reindex(pr, fill_value=0)
)
monthly_table = pd.DataFrame([monthly.values], columns=[p.strftime("%Y-%m") for p in pr])
monthly_table.insert(0, "Показатель", "Обращения")
st.dataframe(monthly_table, use_container_width=True, hide_index=True)

st.divider()

# ======== 3) ТОП-5 ЭЗС ========
st.markdown("### ТОП-5 ЭЗС по обращениям (по фильтрам)")
top5_view = pd.DataFrame()
if not col_station:
    st.info("Колонка 'Номер ЭЗС' не выбрана — ТОП-5 недоступен.")
else:
    top5 = (
        fdf.groupby(col_station)
           .size()
           .sort_values(ascending=False)
           .head(5)
           .rename("Обращения")
           .reset_index()
    )
    plant_mode = (
        fdf.groupby(col_station)["Завод"]
           .agg(lambda s: s.dropna().astype(str).mode().iloc[0] if len(s.dropna()) else "Другое")
    )
    top5["Завод"] = top5[col_station].map(plant_mode)
    top5_view = top5
    st.dataframe(top5_view, use_container_width=True, hide_index=True)

st.divider()

# ======== Сырые данные ========
st.markdown("### Сырые данные (по фильтрам)")
show_cols = []
for c in [col_date, col_time, col_reason, "Завод", col_station, col_vendor, col_note, "_source_gid"]:
    if c and c in fdf.columns and c not in show_cols:
        show_cols.append(c)

display_df = fdf.sort_values("_dt", ascending=False)
st.dataframe(display_df[show_cols], use_container_width=True, hide_index=True)

# ======== Экспорт ========
st.markdown("### Экспорт")
d1, d2 = st.columns(2)

with d1:
    st.download_button(
        "⬇️ CSV (фильтрованные данные)",
        data=display_df[show_cols].to_csv(index=False).encode("utf-8-sig"),
        file_name="ezs_filtered_2025.csv",
        mime="text/csv",
    )

with d2:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        view_reason.to_excel(w, sheet_name="reason_x_plant", index=False)
        monthly_table.to_excel(w, sheet_name="monthly_2025", index=False)
        if len(top5_view):
            top5_view.to_excel(w, sheet_name="top5", index=False)
        display_df[show_cols].to_excel(w, sheet_name="raw_filtered", index=False)
    st.download_button(
        "⬇️ Excel (сводка)",
        data=bio.getvalue(),
        file_name="ezs_dashboard_2025.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption("Лёгкая версия v2: заводы строго E-Prom/NSP/Другое (пусто/не распознано). Добавлен блок диагностики значений в исходной колонке.")
