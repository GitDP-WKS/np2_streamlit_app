import json
import re
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


st.set_page_config(page_title="ЭЗС — дашборд обращений", layout="wide")


DEFAULT_SHEET_ID = "1YN_8UtrZMqOTYZHaLzczwkkfocD-sS_wKrlSBmn-S50"
DEFAULT_GID = "2075524941"

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


def gsheets_csv_url(sheet_id: str, gid: str) -> str:
    # Public export URL (works when sheet is shared "anyone with link" as Viewer)
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip()).lower()


def pick_col(columns: List[str], candidates: List[str]) -> Optional[str]:
    norm_cols = {_norm(c): c for c in columns}
    for cand in candidates:
        key = _norm(cand)
        if key in norm_cols:
            return norm_cols[key]
    # fallback: partial match
    for cand in candidates:
        key = _norm(cand)
        for n, orig in norm_cols.items():
            if key in n:
                return orig
    return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df


def parse_datetime(df: pd.DataFrame) -> Tuple[pd.DataFrame, str, str]:
    df = normalize_columns(df)

    col_date = pick_col(df.columns.tolist(), ["Дата обращения", "Дата", "Date"])
    col_time = pick_col(df.columns.tolist(), ["Время обращения", "Время", "Time", "Время обращения "])

    if col_date is None:
        raise ValueError("Не нашёл колонку с датой. Ожидал 'Дата обращения'.")

    # В некоторых выгрузках время может быть пустым — тогда считаем только дату
    if col_time is None:
        df["_dt"] = pd.to_datetime(df[col_date], dayfirst=True, errors="coerce")
        return df, col_date, ""

    df["_dt"] = pd.to_datetime(
        df[col_date].astype(str) + " " + df[col_time].astype(str),
        dayfirst=True,
        errors="coerce",
    )
    return df, col_date, col_time


def add_week(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # Неделя с началом в понедельник
    df["_week_start"] = df["_dt"].dt.to_period("W-MON").dt.start_time
    df["_week_label"] = df["_week_start"].dt.strftime("%Y-%m-%d")
    return df


def classify_theme(reason: str, rules: List[Dict]) -> str:
    text = _norm(reason)
    for r in rules:
        theme = r.get("theme", "").strip() or "Без темы"
        kws = r.get("keywords", [])
        if any(_norm(k) in text for k in kws if str(k).strip()):
            return theme
    return "Другое"


def apply_themes(df: pd.DataFrame, reason_col: str, rules: List[Dict]) -> pd.DataFrame:
    df = df.copy()
    df["Тема"] = df[reason_col].astype(str).apply(lambda x: classify_theme(x, rules))
    return df


@st.cache_data(ttl=600, show_spinner=False)
def load_data(sheet_id: str, gid: str) -> pd.DataFrame:
    url = gsheets_csv_url(sheet_id, gid)
    df = pd.read_csv(url)
    return df


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            sdf.to_excel(writer, sheet_name=name[:31], index=False)
    return bio.getvalue()


st.title("📊 ЭЗС — дашборд обращений")

with st.sidebar:
    st.header("Источник данных")
    sheet_id = st.text_input("Google Sheet ID", value=DEFAULT_SHEET_ID)
    gid = st.text_input("GID (лист)", value=DEFAULT_GID)
    st.caption("Важно: таблица должна быть расшарена как “Anyone with the link → Viewer”.")

    if st.button("🔄 Обновить кэш"):
        st.cache_data.clear()
        st.toast("Кэш очищен, данные перечитаются заново.", icon="✅")

    st.divider()
    st.header("Тематики")
    rules_text = st.text_area(
        "Правила (JSON). Можно править под себя.",
        value=json.dumps(DEFAULT_THEME_RULES, ensure_ascii=False, indent=2),
        height=260,
    )

    try:
        theme_rules = json.loads(rules_text)
        if not isinstance(theme_rules, list):
            raise ValueError("JSON должен быть списком правил.")
    except Exception as e:
        st.error(f"Ошибка в правилах JSON: {e}")
        theme_rules = DEFAULT_THEME_RULES


# --- Load ---
try:
    raw = load_data(sheet_id, gid)
except Exception as e:
    st.error(
        "Не смог загрузить таблицу. "
        "Проверь доступ (публичный просмотр) и правильность Sheet ID / GID.\n\n"
        f"Текст ошибки: {e}"
    )
    st.stop()

# --- Parse & normalize ---
df, col_date, col_time = parse_datetime(raw)

col_reason = pick_col(df.columns.tolist(), ["Причина обращения", "Причина", "Тема обращения"])
col_station = pick_col(df.columns.tolist(), ["Номер ЭЗС", "ЭЗС", "Station", "Станция"])
col_vendor = pick_col(df.columns.tolist(), ["Производитель станции", "Производитель", "Vendor"])
col_note = pick_col(df.columns.tolist(), ["Примечание", "Комментарий", "Note"])
col_id = pick_col(df.columns.tolist(), ["№", "N", "No", "Номер", "ID"])

if col_reason is None:
    st.error("Не нашёл колонку 'Причина обращения'. Проверь названия колонок на листе.")
    st.stop()

df = add_week(df)
df = apply_themes(df, col_reason, theme_rules)

# --- Filters ---
st.subheader("Фильтры")

min_dt = df["_dt"].min()
max_dt = df["_dt"].max()

c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 1.4])

with c1:
    # default: latest week
    latest_week = df["_week_label"].dropna().max()
    week_mode = st.radio("Период", ["Последняя неделя", "Выбор недели", "Диапазон дат"], horizontal=False)

with c2:
    if week_mode == "Выбор недели":
        week = st.selectbox("Неделя (понедельник)", sorted(df["_week_label"].dropna().unique())[::-1], index=0)
    else:
        week = latest_week

with c3:
    if week_mode == "Диапазон дат":
        start_date = st.date_input("С даты", value=(min_dt.date() if pd.notna(min_dt) else None))
        end_date = st.date_input("По дату", value=(max_dt.date() if pd.notna(max_dt) else None))
    else:
        start_date, end_date = None, None

with c4:
    theme_filter = st.multiselect(
        "Тематики",
        options=sorted(df["Тема"].dropna().unique()),
        default=[],
        placeholder="Все тематики",
    )

# manufacturer filter
if col_vendor is not None:
    vendors = sorted([v for v in df[col_vendor].dropna().unique()])
    vendor_filter = st.multiselect("Производитель", options=vendors, default=[], placeholder="Все производители")
else:
    vendor_filter = []

# Apply filters
fdf = df.copy()

if week_mode in ("Последняя неделя", "Выбор недели") and week:
    fdf = fdf[fdf["_week_label"] == week]

if week_mode == "Диапазон дат" and start_date and end_date:
    # inclusive end date
    fdf = fdf[(fdf["_dt"] >= pd.to_datetime(start_date)) & (fdf["_dt"] < pd.to_datetime(end_date) + pd.Timedelta(days=1))]

if theme_filter:
    fdf = fdf[fdf["Тема"].isin(theme_filter)]

if vendor_filter and col_vendor is not None:
    fdf = fdf[fdf[col_vendor].isin(vendor_filter)]

# --- KPIs ---
k1, k2, k3, k4 = st.columns(4)

total = int(len(fdf))
unique_stations = int(fdf[col_station].nunique()) if col_station is not None else 0
top_theme = (fdf["Тема"].value_counts().index[0] if total else "—")
top_reason = (fdf[col_reason].value_counts().index[0] if total else "—")

k1.metric("Обращений", f"{total}")
k2.metric("Уникальных ЭЗС", f"{unique_stations}" if col_station else "—")
k3.metric("Топ-тематика", top_theme)
k4.metric("Частая причина", top_reason)

st.divider()

# --- Charts & tables ---
left, right = st.columns([1.2, 1])

with left:
    st.markdown("#### Динамика по неделям")
    trend = (
        df.dropna(subset=["_week_start"])
          .groupby("_week_start")
          .size()
          .rename("Обращения")
          .reset_index()
          .sort_values("_week_start")
    )
    if len(trend):
        st.line_chart(trend.set_index("_week_start")["Обращения"])
    else:
        st.info("Недостаточно данных для графика.")

with right:
    st.markdown("#### Разбивка по тематикам (в выбранном периоде)")
    theme_counts = fdf["Тема"].value_counts().rename_axis("Тема").reset_index(name="Обращения")
    st.dataframe(theme_counts, use_container_width=True, hide_index=True)
    if len(theme_counts):
        st.bar_chart(theme_counts.set_index("Тема")["Обращения"])

st.divider()

st.markdown("#### ТОП-5 ЭЗС по обращениям (в выбранном периоде)")
if col_station is None:
    st.warning("Не нашёл колонку 'Номер ЭЗС' — ТОП-5 по станциям недоступен.")
else:
    top5 = (
        fdf.groupby(col_station)
           .size()
           .sort_values(ascending=False)
           .head(5)
           .rename("Обращения")
           .reset_index()
    )
    if col_vendor is not None:
        # добавить производителя (самый частый для станции)
        vendor_map = (
            fdf.groupby(col_station)[col_vendor]
               .agg(lambda s: s.dropna().mode().iloc[0] if len(s.dropna().mode()) else "")
        )
        top5["Производитель"] = top5[col_station].map(vendor_map)

    st.dataframe(top5, use_container_width=True, hide_index=True)

    # Барчарт для топ-5
    st.bar_chart(top5.set_index(col_station)["Обращения"])

st.divider()

st.markdown("#### Сырые данные (после фильтров)")
show_cols = []
for c in [col_id, col_date, col_time, col_reason, "Тема", col_station, col_vendor, col_note]:
    if c and c in fdf.columns and c not in show_cols:
        show_cols.append(c)
if not show_cols:
    show_cols = fdf.columns.tolist()

st.dataframe(fdf[show_cols].sort_values("_dt", ascending=False), use_container_width=True, hide_index=True)

# --- Downloads ---
st.markdown("### Экспорт")
d1, d2 = st.columns([1, 1])

with d1:
    st.download_button(
        "⬇️ Скачать CSV (по фильтрам)",
        data=to_csv_bytes(fdf[show_cols]),
        file_name="ezs_filtered.csv",
        mime="text/csv",
    )

with d2:
    # несколько листов: raw filtered + theme + top5
    sheets = {"filtered": fdf[show_cols]}
    sheets["themes"] = theme_counts
    if col_station is not None:
        sheets["top5"] = top5 if "top5" in locals() else pd.DataFrame()
    st.download_button(
        "⬇️ Скачать Excel (сводка)",
        data=to_excel_bytes(sheets),
        file_name="ezs_dashboard_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption("Если хочешь — добавлю отдельный блок “ТОП причин” и переключатель “месяц/квартал/год”.")
