"""Построение листа 'Сценарный анализ' - сводная таблица по 5 сценариям и сигналы."""
from __future__ import annotations

from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from za_price_calculator.config import PALETTE, SHEETS
from za_price_calculator.core.sheet_calculations import (
    COL_MARGIN_CUR,
    COL_MARGIN_SIG,
    COL_NEW_RETAIL,
    COL_QTY,
    COL_RETAIL,
    COL_RISK,
    COL_SIGNAL,
    COL_VP_CUR,
    S1,
    S2,
    S3,
    S4,
    S5,
)
from za_price_calculator.styling.styles import (
    HDR_FILL,
    HDR_FONT,
    NUM2,
    PCT,
    align,
    border_thin,
    fill,
    font,
)


def _L(col: int) -> str:
    return get_column_letter(col)


# Строки сценариев: С1/С3/С5 скрыты от пользователя, С2 и С4 видимы.
_HIDDEN_SCENARIO_ROWS = (7, 9, 11)  # С1, С3, С5


def build_scenarios_sheet(ws: Worksheet, n_rows: int) -> None:
    """
    Строит сводный лист сравнения сценариев ценообразования и блок автосигналов.

    :param ws: целевой лист.
    :param n_rows: количество строк данных на листе 'Расчеты' (N).
    """
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PALETTE.purple

    for col, w in (("A", 3), ("B", 32), ("C", 24), ("D", 16), ("E", 16), ("F", 16), ("G", 16), ("H", 30)):
        ws.column_dimensions[col].width = w

    ws.merge_cells("B2:H2")
    t = ws["B2"]
    t.value = "СЦЕНАРНЫЙ АНАЛИЗ - Сравнение сценариев (видимые: С2, С4) | Релиз 3.5"
    t.font = font(bold=True, size=14, color=PALETTE.white)
    t.fill = fill(PALETTE.dark_blue)
    t.alignment = align("center")
    ws.row_dimensions[2].height = 32

    qty_L = _L(COL_QTY)
    new_retail_L = _L(COL_NEW_RETAIL)
    ws.merge_cells("B3:H3")
    s = ws["B3"]
    s.value = (
        f"Кол-во продаж подтягивается с листа Продажи на Расчеты (кол. {qty_L}) по штрихкоду. "
        f"«Сумм.прогн.ВП» = Σ(кол-во × ВП/ед); без продаж — сумма ВП/ед. "
        f"Новая Розничная цена ЗЯ — кол. {new_retail_L}; С4 берёт её автоматически. "
        "С1/С3/С5 скрыты."
    )
    s.font = font(size=9, color="595959")
    s.alignment = align("left", indent=1)
    ws.row_dimensions[3].height = 16

    if n_rows == 0:
        ws.merge_cells("B5:H6")
        c = ws["B5"]
        c.value = "Нет данных для сценарного анализа - загрузите прайс-лист с товарами."
        c.font = font(bold=True, size=12, color=PALETTE.accent_red)
        c.alignment = align("center", wrap=True)
        return

    cs = f"'{SHEETS.calculations}'"
    last = n_rows + 2

    hdr_row = 5
    for ci, h in enumerate(
        ["Сценарий", "Описание", "Ср.новая цена", "Ср.новая маржа", "Ср.ВП/ед", "Сумм.прогн.ВП", "Вывод"], start=2
    ):
        c = ws.cell(hdr_row, ci, h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = align("center", wrap=True)
        c.border = border_thin()
    ws.row_dimensions[hdr_row].height = 22

    def avg_if(col):
        return f'=IFERROR(AVERAGE(IF({cs}!{col}3:{col}{last}<>"",{cs}!{col}3:{col}{last})),"")'

    def avg_if_gt0(col):
        return f'=IFERROR(AVERAGE(IF({cs}!{col}3:{col}{last}>0,{cs}!{col}3:{col}{last})),"")'

    def sum_forecast_vp(vp_per_col: str) -> str:
        """
        Суммарный прогноз ВП = Σ (Кол-во продаж × ВП/ед сценария).
        Если продаж нет — сумма ВП/ед (ориентир без объёма).
        N() приводит пустые формулы "" к 0, без #VALUE!.
        """
        qty_rng = f"{cs}!{qty_L}3:{qty_L}{last}"
        vp_rng = f"{cs}!{vp_per_col}3:{vp_per_col}{last}"
        return (
            f'=IFERROR(IF(SUMPRODUCT(N({qty_rng}))=0,'
            f'SUMPRODUCT(N({vp_rng})),'
            f'SUMPRODUCT(N({qty_rng}),N({vp_rng}))),"")'
        )

    d_col = _L(COL_RETAIL)
    s_col = _L(COL_MARGIN_CUR)
    t_col = _L(COL_VP_CUR)
    s1p, s1m, s1v = _L(S1[0]), _L(S1[2]), _L(S1[3])
    s2p, s2m, s2v = _L(S2[0]), _L(S2[2]), _L(S2[3])
    s3p, s3m, s3v = _L(S3[0]), _L(S3[2]), _L(S3[3])
    s4p, s4m, s4v = _L(S4[0]), _L(S4[2]), _L(S4[3])
    s5p, s5m, s5v = _L(S5[1]), _L(S5[3]), _L(S5[4])

    scen_rows = [
        (6, "Текущая", "Действующая цена ЗЯ",
         avg_if(d_col), avg_if(s_col), avg_if(t_col), sum_forecast_vp(t_col),
         "Базовый уровень", fill(PALETTE.grey_bg)),
        (7, "С1: Средняя рынка", "Цена = средняя по конкурентам",
         avg_if(s1p), avg_if(s1m), avg_if(s1v), sum_forecast_vp(s1v),
         "Выравнивание по рынку", fill("EBF3FB")),
        (8, "С2: Медиана рынка", "Цена = медианная по конкурентам",
         avg_if(s2p), avg_if(s2m), avg_if(s2v), sum_forecast_vp(s2v),
         "Устойчив к выбросам", fill("EBF3FB")),
        (9, "С3: Мин.конкурент", "Цена = минимальный конкурент",
         avg_if(s3p), avg_if(s3m), avg_if(s3v), sum_forecast_vp(s3v),
         "Риск потери маржи", fill(PALETTE.red_bg)),
        (10, "С4: Произв.цена", f"Цена из «Новая Розничная цена ЗЯ» (кол.{new_retail_L})",
         avg_if_gt0(s4p), avg_if(s4m), avg_if(s4v), sum_forecast_vp(s4v),
         "Пользовательский сценарий", fill("F0E6FF")),
        (11, "С5: Целевая маржа", f"Цена по целевой марже (кол.{_L(S5[0])} в Расчетах)",
         avg_if(s5p), avg_if(s5m), avg_if(s5v), sum_forecast_vp(s5v),
         "Маржинальный сценарий", fill(PALETTE.green_bg)),
    ]
    for (row, lbl, desc, f_price, f_mrg, f_vp, f_sum, note, rfill) in scen_rows:
        ws.row_dimensions[row].height = 22
        for ci, val in enumerate([lbl, desc, f_price, f_mrg, f_vp, f_sum, note], start=2):
            c = ws.cell(row, ci, val)
            c.fill = rfill
            c.border = border_thin()
            c.font = font(bold=(ci == 2), size=10)
            c.alignment = align("center") if ci > 3 else align("left", indent=1)
            if ci == 4:
                c.number_format = NUM2
            elif ci == 5:
                c.number_format = PCT
            elif ci in (6, 7):
                c.number_format = NUM2

    for row in _HIDDEN_SCENARIO_ROWS:
        ws.row_dimensions[row].hidden = True

    ws.row_dimensions[13].height = 8
    ws.merge_cells("B14:H14")
    c = ws["B14"]
    c.value = "АВТОМАТИЧЕСКИЕ СИГНАЛЫ"
    c.font = font(bold=True, size=11, color=PALETTE.white)
    c.fill = fill(PALETTE.dark_blue)
    c.alignment = align("center")
    ws.row_dimensions[14].height = 22

    sig, risk, msig = _L(COL_SIGNAL), _L(COL_RISK), _L(COL_MARGIN_SIG)
    signals = [
        ("Позиций с ценой ВЫШЕ рынка", f'=COUNTIF({cs}!{sig}3:{sig}10000,"Выше рынка")',
         "Рассмотреть снижение цены до С2 (медиана рынка)"),
        ("Позиций с ценой НИЖЕ рынка", f'=COUNTIF({cs}!{sig}3:{sig}10000,"Ниже рынка")',
         "Потенциал повышения цены без потери конкурентоспособности"),
        ("Позиций с КРИТИЧНЫМ риском ВП", f'=COUNTIF({cs}!{risk}3:{risk}10000,"КРИТИЧНО")',
         "Маржа при цене мин.конкурента <5% — снижать цену опасно"),
        ("Позиций с УМЕРЕННЫМ риском ВП", f'=COUNTIF({cs}!{risk}3:{risk}10000,"УМЕРЕННЫЙ")',
         "Маржа при цене мин.конкурента 5–15% — мониторинг"),
        ("Позиций с ХОРОШЕЙ маржой (>20%)", f'=COUNTIF({cs}!{msig}3:{msig}10000,"Хорошая(>20%)")',
         "Высокая ценовая гибкость"),
        ("Позиций с НИЗКОЙ маржой (<10%)", f'=COUNTIF({cs}!{msig}3:{msig}10000,"Низкая(<10%)")',
         "Приоритет: переговоры с поставщиком / ценовая корректировка"),
    ]
    for j, (lbl, frm, rec) in enumerate(signals):
        row_j = 15 + j
        ws.row_dimensions[row_j].height = 22
        ws.merge_cells(f"B{row_j}:C{row_j}")
        c = ws.cell(row_j, 2, lbl)
        c.font = font(bold=True, size=10)
        c.fill = fill(PALETTE.grey_bg)
        c.alignment = align("left", indent=1)
        c.border = border_thin()

        c = ws.cell(row_j, 4, frm)
        c.font = font(bold=True, size=14)
        c.number_format = "#,##0"
        c.alignment = align("center")
        c.fill = fill(PALETTE.light_blue)
        c.border = border_thin()

        ws.merge_cells(f"E{row_j}:H{row_j}")
        c = ws.cell(row_j, 5, rec)
        c.font = font(size=9, color="595959")
        c.alignment = align("left", indent=1, wrap=True)
        c.fill = fill(PALETTE.grey_bg)
        c.border = border_thin()

    ws.conditional_formatting.add(
        "D15:D20", CellIsRule(operator="greaterThan", formula=["0"], fill=fill(PALETTE.red_bg))
    )
