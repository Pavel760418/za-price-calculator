"""
Веб-приложение AI-агента ценообразования сети супермаркетов «Зеленое Яблоко».

AI-агент рассчитывает цены, наценку и маржу по отношению к конкурентам
в различных сценариях, формирует итоговый Excel-файл и выводит ключевой
BI-дашборд с выводами прямо в браузере.

Запуск локально:
    python -m streamlit run streamlit_app.py

Развёртывание (Streamlit Community Cloud):
    Укажите этот файл (streamlit_app.py) как главный модуль,
    зависимости берутся из requirements.txt.

Приложение оборачивает существующий публичный API
(za_price_calculator.service.ZAPriceCalculator) без изменения логики модуля.
Метрики дашборда считаются в pandas по той же методике, что и в Excel-книге
(наценка, маржа, отклонение от рынка, сценарии С1–С3).
"""
from __future__ import annotations

import io
import logging
import tempfile
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from za_price_calculator.config import COMPETITOR_COLUMNS, THRESHOLDS
from za_price_calculator.exceptions import ZAPriceCalculatorError
from za_price_calculator.io_handlers.loader import load_source_file
from za_price_calculator.service import ZAPriceCalculator

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)

st.set_page_config(
    page_title="Зеленое Яблоко · AI-агент ценообразования",
    page_icon="🍏",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- Названия конкурентов для человекочитаемых подписей ---
COMPETITOR_LABELS = {
    "Цена_7_Континент": "7 Континент",
    "Цена_Европейский": "Европейский",
    "Цена_Тропики": "Тропики",
    "Цена_Мята": "Мята",
    "Цена_Отличный": "Отличный",
    "Цена_Оптовик_Молоток": "Оптовик (Молоток)",
    "Цена_Оптовик_Редукторный": "Оптовик (Редукторный)",
}

APP_VERSION = "2026-07-20.2"

GREEN = "#1a7f37"
GREEN_LIGHT = "#2ea043"
RED = "#c0392b"
AMBER = "#d4a017"
BLUE = "#1f6feb"

# --- Шаблон файла продаж ---
SALES_TEMPLATE_HEADERS = ["Штрихкод", "Кол-во продаж", "Выручка", "Валовая прибыль"]
# Примеры используют реальные штрихкоды из прайса — при загрузке они сопоставятся с товарами.
SALES_TEMPLATE_ROWS = [
    ("4665272570017", 120, 3480, 900),
    ("4602984001637", 85, 8330, 2100),
    ("8690511171614", 40, 25960, 5200),
]


@st.cache_data(show_spinner=False)
def _sales_template_bytes() -> bytes:
    """Формирует Excel-шаблон файла продаж с заголовками и примерами строк."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Продажи"

    header_fill = PatternFill("solid", fgColor="1A7F37")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for ci, title in enumerate(SALES_TEMPLATE_HEADERS, start=1):
        cell = ws.cell(1, ci, title)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for ri, row in enumerate(SALES_TEMPLATE_ROWS, start=2):
        for ci, value in enumerate(row, start=1):
            cell = ws.cell(ri, ci, value)
            if ci == 1:  # штрихкод храним как текст, чтобы не терять ведущие нули
                cell.number_format = "@"

    for ci, width in enumerate((22, 16, 14, 18), start=1):
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# =============================== Оформление =================================
st.markdown(
    """
    <style>
      #MainMenu, footer {visibility: hidden;}
      .block-container {padding-top: 1.6rem; padding-bottom: 2rem; max-width: 1250px;}
      .za-hero {
        background: linear-gradient(120deg, #0f5132 0%, #1a7f37 45%, #2ea043 100%);
        border-radius: 18px; padding: 26px 32px; color: #fff;
        box-shadow: 0 10px 30px rgba(26,127,55,.28); margin-bottom: 8px;
      }
      .za-hero h1 {margin: 0 0 6px 0; font-size: 30px; font-weight: 800; letter-spacing:.2px;}
      .za-hero p {margin: 0; font-size: 15px; opacity: .94; line-height: 1.5; max-width: 900px;}
      .za-badge {
        display:inline-block; background: rgba(255,255,255,.18); backdrop-filter: blur(4px);
        border:1px solid rgba(255,255,255,.35); color:#fff; font-size:12px; font-weight:600;
        padding:4px 12px; border-radius:999px; margin-bottom:14px;
      }
      div[data-testid="stMetric"] {
        background: #ffffff; border: 1px solid #e6e8eb; border-radius: 14px;
        padding: 14px 16px; box-shadow: 0 1px 3px rgba(16,24,40,.06);
      }
      div[data-testid="stMetricLabel"] p {font-size: 13px; color:#57606a; font-weight:600;}
      div[data-testid="stMetricValue"] {font-size: 26px; color:#0f5132;}
      .za-insight {
        border-radius: 12px; padding: 14px 16px; margin-bottom: 10px; font-size: 14.5px;
        line-height: 1.5; border-left: 5px solid; background:#fff; border:1px solid #eaecef;
      }
      .za-insight b {font-weight: 700;}
      .za-ok   {border-left-color:#1a7f37;}
      .za-warn {border-left-color:#d4a017;}
      .za-bad  {border-left-color:#c0392b;}
      .za-info {border-left-color:#1f6feb;}
      .stTabs [data-baseweb="tab-list"] {gap: 6px;}
      .stTabs [data-baseweb="tab"] {font-weight: 600;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="za-hero">
      <div class="za-badge">🤖 AI-агент ценообразования</div>
      <h1>🍏 Зеленое Яблоко</h1>
      <p>Интеллектуальный агент по расчёту <b>цен, наценки и маржи</b> сети супермаркетов
      «Зеленое Яблоко» по отношению к <b>конкурентам</b> в различных
      <b>сценариях</b>. Загрузите прайс-лист — агент сформирует готовый Excel-файл
      и покажет ключевой BI-дашборд с выводами и рекомендациями.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.caption(
    f"🔖 Версия интерфейса: {APP_VERSION} · шаблон продаж + BI-дашборд. "
    "Если вы не видите блок «С чего начать» и кнопку «Скачать шаблон продаж» — "
    "приложение работает на старой версии: перезапустите его (локально — заново "
    "`streamlit run`; на Streamlit Cloud — Manage app → Reboot) и обновите страницу "
    "с очисткой кэша (Ctrl/Cmd+Shift+R)."
)


# =============================== Логика расчётов =============================
@st.cache_data(show_spinner=False)
def _load_source_df(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    """Загружает и нормализует прайс-лист во временном файле (кэшируется по байтам)."""
    suffix = Path(file_name).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=True) as tmp:
        tmp.write(file_bytes)
        tmp.flush()
        return load_source_file(tmp.name)


@st.cache_data(show_spinner=False)
def _build_excel(
    source_bytes: bytes, source_name: str,
    sales_bytes: bytes | None, sales_name: str | None,
) -> bytes:
    """Формирует итоговый Excel-файл через публичный API и возвращает его байты."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        src = tmp_dir / (Path(source_name).stem + (Path(source_name).suffix or ".xlsx"))
        src.write_bytes(source_bytes)
        sales = None
        if sales_bytes is not None:
            sales = tmp_dir / (Path(sales_name).stem + (Path(sales_name).suffix or ".xlsx"))
            sales.write_bytes(sales_bytes)
        out = tmp_dir / "Зеленое_Яблоко_калькулятор.xlsx"
        ZAPriceCalculator().run(
            source_path=str(src),
            output_path=str(out),
            sales_path=str(sales) if sales else None,
        )
        return out.read_bytes()


def _compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Считает наценку, маржу, отклонение от рынка и сигналы по методике модуля."""
    d = df.copy()
    # Приводим цены конкурентов к числам: пустые/нечисловые столбцы иначе остаются
    # object, из-за чего Ср_рынок/Откл_рынок становятся object и ломают nlargest.
    comp = d[COMPETITOR_COLUMNS].apply(pd.to_numeric, errors="coerce")
    d[COMPETITOR_COLUMNS] = comp
    d["Ср_рынок"] = comp.mean(axis=1)
    d["Медиана_рынок"] = comp.median(axis=1)
    d["Мин_конкурент"] = comp.min(axis=1)

    purch = pd.to_numeric(d["Закупочная_цена"], errors="coerce")
    za = pd.to_numeric(d["Розничная_цена_ЗЯ"], errors="coerce")
    d["Закупочная_цена"] = purch
    d["Розничная_цена_ЗЯ"] = za

    d["Валидна"] = purch.notna() & (purch > 0) & za.notna() & (za > 0)
    d["ВП_ед"] = za - purch
    d["Наценка"] = (za - purch) / purch
    d["Маржа"] = (za - purch) / za
    d["Откл_рынок"] = (za - d["Ср_рынок"]) / d["Ср_рынок"]

    def signal(x):
        if pd.isna(x):
            return "Нет данных рынка"
        if x > THRESHOLDS.price_above_market:
            return "Выше рынка"
        if x < THRESHOLDS.price_below_market:
            return "Ниже рынка"
        return "На уровне"

    d["Сигнал"] = d["Откл_рынок"].apply(signal)

    def margin_class(m):
        if pd.isna(m):
            return "—"
        if m > THRESHOLDS.margin_high:
            return "Хорошая (>20%)"
        if m < THRESHOLDS.margin_low:
            return "Низкая (<10%)"
        return "Средняя (10–20%)"

    d["Класс_маржи"] = d["Маржа"].apply(margin_class)

    # Риск при выравнивании до минимального конкурента (сценарий С3)
    m_c3 = (d["Мин_конкурент"] - purch) / d["Мин_конкурент"]
    d["Маржа_С3"] = m_c3

    def risk(m):
        if pd.isna(m):
            return "—"
        if m < THRESHOLDS.risk_critical:
            return "КРИТИЧНО"
        if m < THRESHOLDS.risk_moderate:
            return "УМЕРЕННЫЙ"
        return "ОК"

    d["Риск_С3"] = m_c3.apply(risk)
    return d


def _scenarios(d: pd.DataFrame) -> pd.DataFrame:
    """Сводка по сценариям ценообразования (средние по валидным позициям)."""
    v = d[d["Валидна"]].copy()
    purch = v["Закупочная_цена"]

    def block(name, price, note):
        margin = (price - purch) / price
        vp = price - purch
        return {
            "Сценарий": name,
            "Цена": price.mean(),
            "Маржа": margin.mean(),
            "ВП_ед": vp.mean(),
            "Вывод": note,
        }

    rows = [
        block("Текущая", v["Розничная_цена_ЗЯ"], "Базовый уровень"),
        block("С1 · Средняя рынка", v["Ср_рынок"], "Выравнивание по рынку"),
        block("С2 · Медиана рынка", v["Медиана_рынок"], "Устойчив к выбросам"),
        block("С3 · Мин. конкурент", v["Мин_конкурент"], "Риск потери маржи"),
    ]
    return pd.DataFrame(rows)


def _insight(kind: str, text: str) -> None:
    st.markdown(f'<div class="za-insight za-{kind}">{text}</div>', unsafe_allow_html=True)


def _fmt_pct(x) -> str:
    return "—" if pd.isna(x) else f"{x * 100:.1f}%"


def _fmt_num(x) -> str:
    return "—" if pd.isna(x) else f"{x:,.2f}".replace(",", " ")


# =============================== Панель загрузки =============================
with st.sidebar:
    st.markdown("### 📂 Загрузка данных")
    source_upload = st.file_uploader(
        "Прайс-лист (обязательно)",
        type=["xlsx", "xls", "xlsm"],
        help="Excel с товарами, закупочными и розничными ценами ЗЯ и ценами конкурентов.",
    )
    sales_upload = st.file_uploader(
        "Файл продаж (опционально)",
        type=["xlsx", "xls", "xlsm"],
        help="Колонки: Штрихкод (обязательно), Кол-во продаж, Выручка, Валовая прибыль. "
             "Строки сопоставляются с прайсом по штрихкоду.",
    )
    st.download_button(
        "📥 Скачать шаблон продаж (.xlsx)",
        data=_sales_template_bytes(),
        file_name="шаблон_продаж_ЗЯ.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        help="Готовый шаблон с нужными колонками и примерами строк.",
        key="dl_tmpl_sidebar",
    )
    st.caption(
        "Файл продаж загружается в блок «Файл продаж». Обязательная колонка — "
        "**Штрихкод**; остальные (Кол-во продаж, Выручка, Валовая прибыль) — опциональны. "
        "Конкуренты в прайсе: 7 Континент, Европейский, Тропики, Мята, Отличный, "
        "Оптовик (Молоток), Оптовик (Редукторный)."
    )

if source_upload is None:
    st.info(
        "⬅️ Загрузите прайс-лист на панели слева, чтобы AI-агент рассчитал цены "
        "и показал дашборд с выводами по сценариям."
    )
    st.markdown("#### С чего начать")
    step1, step2 = st.columns(2)
    with step1:
        st.markdown(
            "**1. Прайс-лист (обязательно)** — загрузите на панели слева "
            "в блоке «Прайс-лист». Excel с ценами ЗЯ и конкурентов."
        )
    with step2:
        st.markdown(
            "**2. Продажи (опционально)** — загрузите в блок «Файл продаж». "
            "Обязательная колонка — **Штрихкод**. Ниже можно скачать готовый шаблон:"
        )
        st.download_button(
            "📥 Скачать шаблон продаж (.xlsx)",
            data=_sales_template_bytes(),
            file_name="шаблон_продаж_ЗЯ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Готовый шаблон с нужными колонками и примерами строк.",
            key="dl_tmpl_intro",
        )
    st.stop()

# --- Загрузка и расчёт ---
try:
    with st.spinner("AI-агент анализирует данные..."):
        source_bytes = source_upload.getvalue()
        sales_bytes = sales_upload.getvalue() if sales_upload else None
        df = _load_source_df(source_bytes, source_upload.name)
        metrics = _compute_metrics(df)
        scen = _scenarios(metrics)
        excel_bytes = _build_excel(
            source_bytes, source_upload.name, sales_bytes,
            sales_upload.name if sales_upload else None,
        )
except ZAPriceCalculatorError as exc:
    st.error(f"Ошибка обработки данных: {exc}")
    st.stop()
except Exception as exc:  # noqa: BLE001
    st.error("Непредвиденная ошибка при расчёте.")
    st.exception(exc)
    st.stop()

valid = metrics[metrics["Валидна"]]
n_all = len(metrics)
n_valid = len(valid)

# --- Верхняя строка: статус + скачивание ---
c1, c2 = st.columns([3, 1])
with c1:
    st.success(
        f"Готово! Обработано позиций: **{n_all}**, с полной экономикой: **{n_valid}**. "
        "Итоговый Excel-файл сформирован."
    )
with c2:
    st.download_button(
        "⬇️ Скачать Excel",
        data=excel_bytes,
        file_name="Зеленое_Яблоко_калькулятор.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

tab_dash, tab_scen, tab_data = st.tabs(
    ["📊 Дашборд", "🎯 Сценарии", "📋 Данные и выгрузка"]
)

# =============================== Вкладка: Дашборд ============================
with tab_dash:
    avg_margin = valid["Маржа"].mean()
    avg_markup = valid["Наценка"].mean()
    avg_za = valid["Розничная_цена_ЗЯ"].mean()
    avg_vp = valid["ВП_ед"].mean()
    n_above = int((metrics["Сигнал"] == "Выше рынка").sum())
    n_below = int((metrics["Сигнал"] == "Ниже рынка").sum())
    n_level = int((metrics["Сигнал"] == "На уровне").sum())
    n_crit = int((metrics["Риск_С3"] == "КРИТИЧНО").sum())
    n_good = int((metrics["Класс_маржи"] == "Хорошая (>20%)").sum())
    n_low = int((metrics["Класс_маржи"] == "Низкая (<10%)").sum())

    st.markdown("#### Ключевые показатели")
    k = st.columns(4)
    k[0].metric("Позиций в прайсе", f"{n_all:,}".replace(",", " "))
    k[1].metric("Средняя маржа", _fmt_pct(avg_margin))
    k[2].metric("Средняя наценка", _fmt_pct(avg_markup))
    k[3].metric("Средняя ВП / ед.", _fmt_num(avg_vp))
    k2 = st.columns(4)
    k2[0].metric("Средняя цена ЗЯ", _fmt_num(avg_za))
    k2[1].metric("Выше рынка", f"{n_above}", delta=f"{n_above / n_all * 100:.0f}%" if n_all else None,
                 delta_color="inverse")
    k2[2].metric("Ниже рынка", f"{n_below}", delta=f"{n_below / n_all * 100:.0f}%" if n_all else None)
    k2[3].metric("Критический риск (С3)", f"{n_crit}", delta_color="inverse")

    st.divider()
    g1, g2 = st.columns(2)

    with g1:
        st.markdown("##### Позиционирование по цене относительно рынка")
        sig_df = (
            metrics["Сигнал"].value_counts()
            .rename_axis("Сигнал").reset_index(name="Позиции")
        )
        color_scale = alt.Scale(
            domain=["Ниже рынка", "На уровне", "Выше рынка", "Нет данных рынка"],
            range=[GREEN, BLUE, RED, "#8b949e"],
        )
        chart = (
            alt.Chart(sig_df)
            .mark_bar(cornerRadiusEnd=6)
            .encode(
                x=alt.X("Позиции:Q", title="Кол-во позиций"),
                y=alt.Y("Сигнал:N", sort="-x", title=None),
                color=alt.Color("Сигнал:N", scale=color_scale, legend=None),
                tooltip=["Сигнал", "Позиции"],
            )
            .properties(height=210)
        )
        st.altair_chart(chart, width="stretch")

    with g2:
        st.markdown("##### Распределение по уровню маржи")
        mrg_df = (
            valid["Класс_маржи"].value_counts()
            .rename_axis("Класс").reset_index(name="Позиции")
        )
        mrg_scale = alt.Scale(
            domain=["Низкая (<10%)", "Средняя (10–20%)", "Хорошая (>20%)"],
            range=[RED, AMBER, GREEN],
        )
        chart2 = (
            alt.Chart(mrg_df)
            .mark_bar(cornerRadiusEnd=6)
            .encode(
                x=alt.X("Позиции:Q", title="Кол-во позиций"),
                y=alt.Y("Класс:N", sort="-x", title=None),
                color=alt.Color("Класс:N", scale=mrg_scale, legend=None),
                tooltip=["Класс", "Позиции"],
            )
            .properties(height=210)
        )
        st.altair_chart(chart2, width="stretch")

    st.divider()
    st.markdown("#### 🧠 Выводы AI-агента")

    if n_all:
        share_above = n_above / n_all * 100
        share_below = n_below / n_all * 100
        if n_above:
            _insight(
                "bad",
                f"<b>{n_above}</b> позиций ({share_above:.0f}%) стоят <b>выше рынка</b>. "
                "Рекомендуется рассмотреть снижение цены до сценария С1 (средняя рынка) "
                "или С2 (медиана рынка), чтобы повысить конкурентоспособность.",
            )
        if n_below:
            _insight(
                "ok",
                f"<b>{n_below}</b> позиций ({share_below:.0f}%) стоят <b>ниже рынка</b>. "
                "Есть потенциал повышения цены и маржи без потери конкурентоспособности.",
            )
        if n_crit:
            _insight(
                "warn",
                f"<b>{n_crit}</b> позиций получают <b>критический риск</b> маржи (&lt;5%) "
                "при выравнивании до минимального конкурента (С3) — снижать цену опасно.",
            )
        if n_low:
            _insight(
                "warn",
                f"<b>{n_low}</b> позиций имеют <b>низкую маржу</b> (&lt;10%). "
                "Приоритет: переговоры с поставщиком или корректировка цены.",
            )
        _insight(
            "info",
            f"Средняя маржа по прайсу — <b>{_fmt_pct(avg_margin)}</b>, "
            f"средняя наценка — <b>{_fmt_pct(avg_markup)}</b>. "
            f"Позиций с хорошей маржой (&gt;20%): <b>{n_good}</b> — у них максимальная ценовая гибкость.",
        )

# =============================== Вкладка: Сценарии ==========================
with tab_scen:
    st.markdown("#### Сравнение сценариев ценообразования")
    st.caption(
        "Средние значения по позициям с полной экономикой. Сценарии показывают, "
        "как изменятся цена и маржа при разной стратегии относительно конкурентов."
    )

    show = pd.DataFrame({
        "Сценарий": scen["Сценарий"],
        "Ср. цена": scen["Цена"].apply(_fmt_num),
        "Ср. маржа": scen["Маржа"].apply(_fmt_pct),
        "Ср. ВП/ед": scen["ВП_ед"].apply(_fmt_num),
        "Вывод": scen["Вывод"],
    })
    st.dataframe(show, width="stretch", hide_index=True)

    st.markdown("##### Средняя маржа по сценариям")
    scen_chart = (
        alt.Chart(scen)
        .mark_bar(cornerRadiusEnd=6)
        .encode(
            x=alt.X("Сценарий:N", sort=list(scen["Сценарий"]), title=None,
                    axis=alt.Axis(labelAngle=0, labelLimit=160)),
            y=alt.Y("Маржа:Q", title="Средняя маржа", axis=alt.Axis(format="%")),
            color=alt.Color(
                "Сценарий:N",
                scale=alt.Scale(
                    domain=list(scen["Сценарий"]),
                    range=[BLUE, GREEN_LIGHT, GREEN, RED],
                ),
                legend=None,
            ),
            tooltip=[
                "Сценарий",
                alt.Tooltip("Маржа:Q", format=".1%"),
                alt.Tooltip("ВП_ед:Q", format=".2f"),
            ],
        )
        .properties(height=300)
    )
    st.altair_chart(scen_chart, width="stretch")

    base_m = scen.loc[scen["Сценарий"] == "Текущая", "Маржа"].iloc[0]
    c1_m = scen.loc[scen["Сценарий"] == "С1 · Средняя рынка", "Маржа"].iloc[0]
    c3_m = scen.loc[scen["Сценарий"] == "С3 · Мин. конкурент", "Маржа"].iloc[0]
    _insight(
        "info",
        f"Текущая средняя маржа <b>{_fmt_pct(base_m)}</b>. При выравнивании до средней рынка (С1) "
        f"она составит <b>{_fmt_pct(c1_m)}</b>, а при снижении до минимального конкурента (С3) — "
        f"<b>{_fmt_pct(c3_m)}</b>. Сценарий С2 (медиана) устойчивее к ценовым выбросам конкурентов.",
    )

# =============================== Вкладка: Данные ============================
with tab_data:
    st.markdown("#### Топ-10 позиций по текущей марже")
    tm = valid.nlargest(10, "Маржа")
    top_margin = pd.DataFrame({
        "Наименование": tm["Наименование"].values,
        "Маржа": tm["Маржа"].apply(_fmt_pct).values,
        "ВП/ед": tm["ВП_ед"].apply(_fmt_num).values,
        "Цена ЗЯ": tm["Розничная_цена_ЗЯ"].apply(_fmt_num).values,
    })
    st.dataframe(top_margin, width="stretch", hide_index=True)

    st.markdown("#### Топ-10 позиций выше рынка")
    ta = metrics[metrics["Сигнал"] == "Выше рынка"].nlargest(10, "Откл_рынок")
    if len(ta):
        top_above = pd.DataFrame({
            "Наименование": ta["Наименование"].values,
            "Откл. от рынка": ta["Откл_рынок"].apply(
                lambda x: "—" if pd.isna(x) else f"{x * 100:+.1f}%").values,
            "Цена ЗЯ": ta["Розничная_цена_ЗЯ"].apply(_fmt_num).values,
            "Ср. рынок": ta["Ср_рынок"].apply(_fmt_num).values,
        })
        st.dataframe(top_above, width="stretch", hide_index=True)
    else:
        st.caption("Позиций выше рынка не найдено.")

    st.divider()
    st.markdown("#### Итоговый Excel-файл")
    st.write(
        "Полная книга содержит 6 листов: Инструкция, Исходные данные, Продажи, "
        "Расчеты, Dashboard, Сценарный анализ — со всеми формулами и оформлением."
    )
    st.download_button(
        "⬇️ Скачать итоговый Excel (.xlsx)",
        data=excel_bytes,
        file_name="Зеленое_Яблоко_калькулятор.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
