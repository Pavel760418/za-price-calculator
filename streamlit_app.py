"""
Streamlit-приложение для развёртывания калькулятора цен «Зеленое Яблоко».

Запуск локально:
    streamlit run streamlit_app.py

Развёртывание:
    Streamlit Community Cloud — укажите этот файл (streamlit_app.py) как
    главный модуль, зависимости берутся из requirements.txt.

Приложение оборачивает существующий публичный API
(za_price_calculator.service.ZAPriceCalculator) без изменения логики модуля:
загруженные файлы сохраняются во временные пути, поскольку загрузчик модуля
работает с путями к файлам, а не с файловыми объектами.
"""
from __future__ import annotations

import logging
import tempfile
from pathlib import Path

import streamlit as st

from za_price_calculator.exceptions import ZAPriceCalculatorError
from za_price_calculator.service import ZAPriceCalculator

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)

st.set_page_config(
    page_title="ZA Price Calculator",
    page_icon="🍏",
    layout="centered",
)


def _save_upload(upload, directory: Path) -> Path:
    """Сохраняет загруженный файл во временную директорию, сохраняя расширение."""
    suffix = Path(upload.name).suffix or ".xlsx"
    target = directory / f"{Path(upload.name).stem}{suffix}"
    target.write_bytes(upload.getbuffer())
    return target


def _parse_sheet(raw: str):
    """Пустая строка -> 0 (первый лист); число -> индекс; иначе -> имя листа."""
    raw = (raw or "").strip()
    if raw == "":
        return 0
    return int(raw) if raw.isdigit() else raw


st.title("🍏 ZA Price Calculator")
st.caption(
    "Автоматический расчёт цен, наценки, маржи и сценариев для сети «Зеленое Яблоко». "
    "Загрузите прайс-лист — приложение сформирует готовый Excel-файл с 6 листами."
)

with st.form("calc_form"):
    source_upload = st.file_uploader(
        "Прайс-лист (обязательно)",
        type=["xlsx", "xls", "xlsm"],
        help="Excel-файл с товарами, закупочными и розничными ценами и ценами конкурентов.",
    )
    sales_upload = st.file_uploader(
        "Файл продаж (опционально)",
        type=["xlsx", "xls", "xlsm"],
        help="Колонки: Штрихкод, Кол-во продаж, Выручка, Валовая прибыль.",
    )

    with st.expander("Дополнительные параметры"):
        source_sheet = st.text_input(
            "Лист прайс-листа (имя или индекс)", value="0",
            help="По умолчанию 0 — первый лист.",
        )
        sales_sheet = st.text_input(
            "Лист продаж (имя или индекс)", value="0",
        )

    submitted = st.form_submit_button("Рассчитать", type="primary")

if submitted:
    if source_upload is None:
        st.error("Загрузите обязательный файл прайс-листа (--source).")
        st.stop()

    with st.spinner("Выполняется расчёт..."):
        try:
            with tempfile.TemporaryDirectory() as tmp:
                tmp_dir = Path(tmp)
                source_path = _save_upload(source_upload, tmp_dir)
                sales_path = _save_upload(sales_upload, tmp_dir) if sales_upload else None
                output_path = tmp_dir / "ZA_Price_Calculator_result.xlsx"

                ZAPriceCalculator().run(
                    source_path=str(source_path),
                    output_path=str(output_path),
                    sales_path=str(sales_path) if sales_path else None,
                    source_sheet=_parse_sheet(source_sheet),
                    sales_sheet=_parse_sheet(sales_sheet),
                )

                result_bytes = output_path.read_bytes()

            st.success("Готово! Файл успешно сформирован.")
            st.download_button(
                label="⬇️ Скачать результат (.xlsx)",
                data=result_bytes,
                file_name="ZA_Price_Calculator_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except ZAPriceCalculatorError as exc:
            st.error(f"Ошибка обработки: {exc}")
        except Exception as exc:  # noqa: BLE001
            st.exception(exc)
