import json
import re
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="ЭЗС — дашборд обращений", layout="wide")


# =========================
# Defaults
# =========================
DEFAULT_SHEET_ID = "1YN_8UtrZMqOTYZHaLzczwkkfocD-sS_wKrlSBmn-S50"
DEFAULT_GIDS = "2075524941"  # можно указать несколько через запятую

DEFAULT_THEME_RULES = [
    {"theme": "Мобильное приложение", "keywords": ["мобильн", "прилож"]},
    {"theme": "Запуск сессии / Авторизация", "keywords": ["не запускает", "запуск", "авториза", "сесси"]},
    {"theme": "Прерывание сессии", "keywords": ["прерыван", "самопроизвольн"]},
    {"theme": "Платежи / Баланс", "keywords": ["пополн", "банковск", "баланс", "денеж", "возврат"]},
    {"theme": "Скорость/мощность", "keywords": ["низкая скорость", "медленно", "мощност"]},
    {"theme": "Оффлайн/сеть/доступность", "keywords": ["не в сети", "недоступ", "мониторинг"]},
    {"theme": "Парковка/занято", "keywords": ["парков", "занято", "двс", "пдд"]},
    {"theme": "Коннекторы/кнопка", "keywords": ["коннектор", "аварийн", "кнопк"]},
    {"theme": "Установка ЭЗС", "keywords": ["установк", "территори"]},
]

# Нормализация производителей в "заводы"
DEFAULT_VENDOR_MAP = {
    "ЕПРОМ": ["епром", "eprom", "e-prom", "e prom"],
    "НСП": ["нсп", "nsp"],
}


# =========================
# Helpers (safe / robust)
# =========================
def gsheets_csv_url(sheet_id: str, gid: str) -> str:
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip()).lower()


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df


def pick_col(columns: List[str], candidates: List[str]) -> Optional[str]:
    if not columns:
        return None
    norm_cols = {_norm(c): c for c in columns}
    for cand in candidates:
        key = _norm(cand)
        if key in norm_cols:
            return norm_cols[key]
    # partial match fallback
    for cand in candidates:
        key = _norm(cand)
        for n, orig in norm_cols.items():
            if key and key in n:
                return orig
    return None


def safe_stop(msg: str, details: Optional[str] = None) -> None:
    st.error(msg)
    if details:
        with st.expander("Показать детали"):
            st.code(details)
    st.stop()


def read_table_from_upload(uploaded) -> pd.DataFrame:
    name = (uploaded.name or "").lower()
    data = uploaded.getvalue()
    try:
        if name.endswith(".xlsx") or name.endswith(".xls"):
            return pd.read_excel(BytesIO(data))
        return pd.read_csv(BytesIO(data), on_bad_lines="skip")
    except Exception as e:
        safe_stop("Не смог прочитать загруженный файл.", str(e))
    return pd.DataFrame()  # unreachable


@st.cache_data(ttl=600, show_spinner=False)
def load_from_gsheets(sheet_id: str, gid: str) -> pd.DataFrame:
    url = gsheets_csv_url(sheet_id, gid)
    return pd.read_csv(url, on_bad_lines="skip")


def parse_dt_smart(df: pd.DataFrame, col_date: str, col_time: Optional[str]) -> pd.Series:
    """Пробует dayfirst=True и если почти всё NaT — пробует dayfirst=False."""
    if col_time:
        s = df[col_date].astype(str) + " " + df[col_time].astype(str)
    else:
        s = df[col_date].astype(str)

    dt1 = pd.to_datetime(s, dayfirst=True, errors="coerce")
    ok1 = dt1.notna().mean()

    if ok1 >= 0.5:
        return dt1

    dt2 = pd.to_datetime(s, dayfirst=False, errors="coerce")
    ok2 = dt2.notna().mean()

    return dt2 if ok2 > ok1 else dt1


def classify_theme(text: str, rules: List[Dict]) -> str:
    t = _norm(text)
    for r in rules:
        theme = (r.get("theme", "") or "").strip() or "Без темы"
        kws = r.get("keywords", []) or []
        for k in kws:
            kk = _norm(k)
            if kk and kk in t:
                return theme
    return "Другое"


def add_totals_crosstab(index: pd.Series, columns: pd.Series, total_name: str = "Итого") -> pd.DataFrame:
    idx = index.fillna("—").astype(str)
    col = columns.fillna("—").astype(str)
    ct = pd.crosstab(idx, col, dropna=False)
    ct[total_name] = ct.sum(axis=1)
    total_row = ct.sum(axis=0).to_frame().T
    total_row.index = [total_name]
    out = pd.concat([ct, total_row], axis=0).reset_index().rename(columns={"index": "Строка"})
    return out


def normalize_vendor_to_plant(v: str, vendor_map: Dict[str, List[str]]) -> str:
    s = _norm(v)
    if not s or s == "nan":
        return "—"
    for plant, keys in vendor_map.items():
        for k in keys:
            kk = _norm(k)
            if kk and kk in s:
                return plant
    return "Другое"


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        for name, sdf in sheets.items():
            sdf.to_excel(w, sheet_name=name[:31], index=False)
    return bio.getvalue()


# =========================
# UI: source
# =========================
st.title("📊 ЭЗС — дашборд обращений (v5)")

with st.sidebar:
    st.header("Источник")
    mode = st.radio("Откуда брать данные?", ["Google Sheets (1+ листов)", "Загрузить файл (CSV/Excel)"])

    sheet_id = DEFAULT_SHEET_ID
    gids_text = DEFAULT_GIDS
    uploaded = None

    if mode == "Google Sheets (1+ листов)":
        sheet_id = st.text_input("Google Sheet ID", value=DEFAULT_SHEET_ID)
        gids_text = st.text_input("GID листов (через запятую)", value=DEFAULT_GIDS)
        st.caption("Если данные по месяцам лежат на разных вкладках — добавь все GID сюда.")
        if st.button("🔄 Сбросить кэш"):
            st.cache_data.clear()
            st.toast("Кэш очищен.", icon="✅")
    else:
        uploaded = st.file_uploader("Загрузи CSV или Excel", type=["csv", "xlsx", "xls"])

    st.divider()
    st.header("Тематики (опционально)")
    rules_text = st.text_area(
        "Правила (JSON). Можно править.",
        value=json.dumps(DEFAULT_THEME_RULES, ensure_ascii=False, indent=2),
        height=220,
    )
    try:
        theme_rules = json.loads(rules_text)
        if not isinstance(theme_rules, list):
            raise ValueError("JSON должен быть списком.")
    except Exception as e:
        st.warning(f"JSON правил сломан — использую дефолт. ({e})")
        theme_rules = DEFAULT_THEME_RULES

    st.divider()
    st.header("Заводы (ЕПРОМ/НСП)")
    vendor_map_text = st.text_area(
        "Сопоставление производителя → завод (JSON).",
        value=json.dumps(DEFAULT_VENDOR_MAP, ensure_ascii=False, indent=2),
        height=160,
    )
    try:
        vendor_map = json.loads(vendor_map_text)
        if not isinstance(vendor_map, dict):
            raise ValueError("JSON должен быть словарём.")
    except Exception as e:
        st.warning(f"JSON заводов сломан — использую дефолт. ({e})")
        vendor_map = DEFAULT_VENDOR_MAP


# =========================
# Load (support multi-gid)
# =========================
try:
    if mode == "Google Sheets (1+ листов)":
        gids = [g.strip() for g in str(gids_text).split(",") if g.strip()]
        if not gids:
            safe_stop("Не указаны GID листов.")
        frames = []
        errors = []
        for gid in gids:
            try:
                dfi = load_from_gsheets(sheet_id, gid)
                dfi["_source_gid"] = gid
                frames.append(dfi)
            except Exception as e:
                errors.append(f"GID {gid}: {e}")
        if errors:
            st.warning("Некоторые листы не прочитались — продолжил с теми, что удалось.")
            with st.expander("Список ошибок по листам"):
                st.code("\n".join(errors))
        if not frames:
            safe_stop("Не удалось прочитать ни один лист по GID.", "\n".join(errors) if errors else None)
        raw = pd.concat(frames, ignore_index=True, sort=False)
    else:
        if not uploaded:
            st.info("Загрузи файл, чтобы продолжить.")
            st.stop()
        raw = read_table_from_upload(uploaded)
        raw["_source_gid"] = "upload"
except Exception as e:
    safe_stop(
        "Не смог загрузить данные.",
        "Частые причины:\n"
        "1) Таблица не расшарена (нужен public viewer)\n"
        "2) Неверный Sheet ID / GID\n"
        "3) Временный сбой/лимит Google\n\n"
        f"Ошибка: {e}",
    )

if raw is None or len(raw) == 0:
    safe_stop("Таблица пустая или не прочиталась (0 строк).")

df = normalize_columns(raw)

# =========================
# Column mapping (auto + manual)
# =========================
cols = df.columns.tolist()

auto_date = pick_col(cols, ["Дата обращения", "Дата", "Date"])
auto_time = pick_col(cols, ["Время обращения", "Время", "Time"])
auto_reason = pick_col(cols, ["Причина обращения", "Причина", "Тема обращения"])
auto_station = pick_col(cols, ["Номер ЭЗС", "ЭЗС", "Station", "Станция"])
auto_vendor = pick_col(cols, ["Производитель станции", "Производитель", "Vendor"])
auto_note = pick_col(cols, ["Примечание", "Комментарий", "Note"])
auto_id = pick_col(cols, ["№", "N", "No", "Номер", "ID"])

with st.expander("🛠️ Диагностика и сопоставление колонок"):
    st.write("Если авто-распознавание промахнулось — выбери колонки вручную.")
    c1, c2 = st.columns(2)
    with c1:
        col_date = st.selectbox("Колонка ДАТА", options=cols, index=(cols.index(auto_date) if auto_date in cols else 0))
        col_time = st.selectbox("Колонка ВРЕМЯ (можно пусто)", options=["— нет —"] + cols, index=(1 + cols.index(auto_time) if auto_time in cols else 0))
        col_reason = st.selectbox("Колонка ПРИЧИНА (как в таблице)", options=cols, index=(cols.index(auto_reason) if auto_reason in cols else 0))
    with c2:
        col_station = st.selectbox("Колонка НОМЕР ЭЗС (можно пусто)", options=["— нет —"] + cols, index=(1 + cols.index(auto_station) if auto_station in cols else 0))
        col_vendor = st.selectbox("Колонка ПРОИЗВОДИТЕЛЬ (нужно для заводов)", options=["— нет —"] + cols, index=(1 + cols.index(auto_vendor) if auto_vendor in cols else 0))
        col_note = st.selectbox("Колонка ПРИМЕЧАНИЕ (можно пусто)", options=["— нет —"] + cols, index=(1 + cols.index(auto_note) if auto_note in cols else 0))
    st.caption("Первые 10 строк:")
    st.dataframe(df.head(10), use_container_width=True)

# normalize "— нет —"
col_time = None if col_time == "— нет —" else col_time
col_station = None if col_station == "— нет —" else col_station
col_vendor = None if col_vendor == "— нет —" else col_vendor
col_note = None if col_note == "— нет —" else col_note

if not col_date or col_date not in df.columns:
    safe_stop("Не выбрана корректная колонка даты.")
if not col_reason or col_reason not in df.columns:
    safe_stop("Не выбрана корректная колонка причины (как в таблице).")

# =========================
# Parse datetime + derived
# =========================
df["_dt"] = parse_dt_smart(df, col_date, col_time)

if df["_dt"].isna().all():
    safe_stop(
        "Не удалось распарсить даты/время (все значения стали пустыми).",
        "Проверь формат даты/времени в таблице.\n"
        "Если данные на разных листах — убедись, что формат одинаковый.\n"
        "Совет: в Google Sheets поставь тип столбца 'Дата' и 'Время'."
    )

df["_week_start"] = df["_dt"].dt.to_period("W-MON").dt.start_time
df["_week_label"] = df["_week_start"].dt.strftime("%Y-%m-%d")
df["_month"] = df["_dt"].dt.to_period("M")  # Period[M]
df["_month_label"] = df["_month"].astype(str)

# Завод (ЕПРОМ/НСП)
if col_vendor and col_vendor in df.columns:
    df["Завод"] = df[col_vendor].astype(str).apply(lambda x: normalize_vendor_to_plant(x, vendor_map))
else:
    df["Завод"] = "—"

# Тематика (опционально; НЕ заменяет причину)
df["Тема"] = df[col_reason].astype(str).apply(lambda x: classify_theme(x, theme_rules))

# =========================
# Filters
# =========================
st.subheader("Фильтры")
c1, c2, c3, c4, c5 = st.columns([1.2, 1.2, 1.2, 1.4, 1.2])

latest_week = df["_week_label"].dropna().max()

with c1:
    period_mode = st.radio("Период", ["Последняя неделя", "Выбор недели", "Диапазон дат"], horizontal=False)

with c2:
    if period_mode == "Выбор недели":
        week = st.selectbox("Неделя (понедельник)", sorted(df["_week_label"].dropna().unique())[::-1], index=0)
    else:
        week = latest_week

with c3:
    if period_mode == "Диапазон дат":
        min_dt = df["_dt"].min().date()
        max_dt = df["_dt"].max().date()
        start_date = st.date_input("С даты", value=min_dt)
        end_date = st.date_input("По дату", value=max_dt)
    else:
        start_date, end_date = None, None

with c4:
    plant_filter = st.multiselect(
        "Завод",
        options=sorted(df["Завод"].dropna().unique()),
        default=["ЕПРОМ", "НСП"] if set(["ЕПРОМ", "НСП"]).issubset(set(df["Завод"].unique())) else [],
        placeholder="Все заводы",
    )

with c5:
    reason_filter = st.multiselect(
        "Причина (как в таблице)",
        options=sorted(df[col_reason].dropna().astype(str).unique()),
        default=[],
        placeholder="Все причины",
    )

# apply filters
fdf = df
if period_mode in ("Последняя неделя", "Выбор недели") and week:
    fdf = fdf[fdf["_week_label"] == week]

if period_mode == "Диапазон дат":
    start = pd.to_datetime(start_date)
    end = pd.to_datetime(end_date) + pd.Timedelta(days=1)
    fdf = fdf[(fdf["_dt"] >= start) & (fdf["_dt"] < end)]

if plant_filter:
    fdf = fdf[fdf["Завод"].isin(plant_filter)]

if reason_filter:
    fdf = fdf[fdf[col_reason].astype(str).isin(reason_filter)]

# =========================
# KPIs
# =========================
k1, k2, k3, k4 = st.columns(4)
total = int(len(fdf))
uniq_station = int(fdf[col_station].nunique()) if col_station else 0
k1.metric("Обращений", f"{total}")
k2.metric("Уникальных ЭЗС", f"{uniq_station}" if col_station else "—")
k3.metric("Топ-завод", fdf["Завод"].value_counts().index[0] if total else "—")
k4.metric("Топ-причина", fdf[col_reason].value_counts().index[0] if total else "—")

st.divider()

# =========================
# Correct breakdown: Reason x Plant
# =========================
st.markdown("#### Разбивка по причинам × завод (ЕПРОМ / НСП)")
# rows = как в таблице, columns = завод
reason_plant = pd.crosstab(
    fdf[col_reason].fillna("—").astype(str),
    fdf["Завод"].fillna("—").astype(str),
    dropna=False,
).reset_index().rename(columns={col_reason: "Причина"})

# ensure columns order
for col in ["ЕПРОМ", "НСП", "Другое", "—"]:
    if col in reason_plant.columns:
        pass
st.dataframe(reason_plant, use_container_width=True, hide_index=True)

st.divider()

# =========================
# Optional: Themes (still useful)
# =========================
with st.expander("Дополнительно: сводка по тематикам (авто-классификация)"):
    theme_counts = fdf["Тема"].value_counts().rename_axis("Тема").reset_index(name="Обращения")
    st.dataframe(theme_counts, use_container_width=True, hide_index=True)
    if len(theme_counts):
        st.bar_chart(theme_counts.set_index("Тема")["Обращения"])

# =========================
# Monthly summary 2024-2025 across ALL LOADED DATA
# =========================
st.markdown("#### Все обращения по месяцам (2024–2025) — по всем загруженным листам")
df_2425 = df[df["_dt"].dt.year.isin([2024, 2025])].copy()
month_range = pd.period_range("2024-01", "2025-12", freq="M")

monthly = (
    df_2425.dropna(subset=["_dt"])
          .groupby(df_2425["_dt"].dt.to_period("M"))
          .size()
          .reindex(month_range, fill_value=0)
)

monthly_table = pd.DataFrame([monthly.values], columns=monthly.index.astype(str))
monthly_table.insert(0, "Показатель", "Обращения")
st.dataframe(monthly_table, use_container_width=True, hide_index=True)

st.caption("Если тут нули — значит в выбранных GID нет дат 2024–2025 (или даты не парсятся). Проверь GID и формат даты.")

st.divider()

# =========================
# Top-5 stations (filtered)
# =========================
st.markdown("#### ТОП-5 ЭЗС по обращениям (в выбранном периоде)")
top5 = pd.DataFrame()
if not col_station:
    st.info("Колонка Номер ЭЗС не задана — ТОП-5 недоступен.")
else:
    top5 = (
        fdf.groupby(col_station)
           .size()
           .sort_values(ascending=False)
           .head(5)
           .rename("Обращения")
           .reset_index()
    )
    # добавить завод (самый частый по станции)
    plant_map = (
        fdf.groupby(col_station)["Завод"]
           .agg(lambda s: s.dropna().astype(str).mode().iloc[0] if len(s.dropna()) else "—")
    )
    top5["Завод"] = top5[col_station].map(plant_map)

    st.dataframe(top5, use_container_width=True, hide_index=True)
    st.bar_chart(top5.set_index(col_station)["Обращения"])

st.divider()

# =========================
# Raw data (filtered)
# =========================
st.markdown("#### Сырые данные (после фильтров)")
show_cols: List[str] = []
for c in [auto_id, col_date, col_time, col_reason, "Завод", col_station, col_vendor, col_note, "_source_gid"]:
    if c and c in fdf.columns and c not in show_cols:
        show_cols.append(c)
if not show_cols:
    show_cols = [c for c in fdf.columns if not str(c).startswith("_")]

display_df = fdf.sort_values("_dt", ascending=False) if "_dt" in fdf.columns else fdf
st.dataframe(display_df[show_cols], use_container_width=True, hide_index=True)

# =========================
# Export
# =========================
st.markdown("### Экспорт")
d1, d2 = st.columns([1, 1])

with d1:
    st.download_button(
        "⬇️ CSV (по фильтрам)",
        data=display_df[show_cols].to_csv(index=False).encode("utf-8-sig"),
        file_name="ezs_filtered.csv",
        mime="text/csv",
    )

with d2:
    sheets: Dict[str, pd.DataFrame] = {
        "reason_x_plant_filtered": reason_plant,
        "monthly_2024_2025_all": monthly_table,
        "filtered_raw": display_df[show_cols],
    }
    if len(top5):
        sheets["top5_filtered"] = top5

    st.download_button(
        "⬇️ Excel (сводка)",
        data=to_excel_bytes(sheets),
        file_name="ezs_dashboard_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption("v5: исправлена логика — таблица теперь 'Причина (как в Google) × завод (ЕПРОМ/НСП)'. Добавлена загрузка нескольких вкладок (несколько GID).")
