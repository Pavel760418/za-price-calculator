"""
Построение листа 'Расчеты' - ядро калькулятора.
Содержит все формулы: анализ рынка, текущую экономику, 5 сценариев ценообразования
и моделирование продаж. Формулы идентичны оригинальному ZA_Price_Calculator.xlsx
с доработками третьего релиза (наценка сейчас / новая розничная / наценка новая).
"""
from __future__ import annotations

from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from za_price_calculator.config import CALC_GROUPS, PALETTE, SHEETS
from za_price_calculator.styling.styles import (
    NUM0,
    NUM2,
    PCT,
    align,
    border_thin,
    fill,
    font,
)

INPUT_FONT = font(size=10, color=PALETTE.input_text)
FORMULA_FONT = font(size=10, color=PALETTE.formula_text)
XREF_FONT = font(size=10, color=PALETTE.xref_text)
INPUT_FILL = fill(PALETTE.input_blue_bg)

# Индексы колонок (1-based). После «Розничная цена ЗЯ» добавлены 3 колонки (+3 сдвиг).
COL_NAME = 1
COL_BARCODE = 2
COL_PURCHASE = 3
COL_RETAIL = 4
COL_MARKUP_NOW = 5
COL_NEW_RETAIL = 6
COL_MARKUP_NEW = 7
COL_SIGNAL = 8
COL_AVG_MKT = 9
COL_MED_MKT = 10
COL_MIN_MKT = 11
COL_MAX_MKT = 12
COL_DEV_AVG = 13
COL_DEV_MED = 14
COL_DEV_MIN = 15
COL_RANK = 16
COL_POS = 17
COL_MARKUP_CUR = 18
COL_MARGIN_CUR = 19
COL_VP_CUR = 20
COL_MARGIN_SIG = 21
COL_RISK = 22
COL_QTY = 23
COL_REV = 24
COL_VP_SUM = 25
COL_MARGIN_TOT = 26

# Сценарии: (price, markup, margin, vp, d_price, d_margin, d_vp)
S1 = (27, 28, 29, 30, 31, 32, 33)
S2 = (34, 35, 36, 37, 38, 39, 40)
S3 = (41, 42, 43, 44, 45, 46, 47)
S4 = (48, 49, 50, 51, 52, 53, 54)
S5 = (55, 56, 57, 58, 59, 60, 61, 62)  # input + price + 5 metrics

# Моделирование продаж
COL_S1_REV, COL_S1_VP = 63, 64
COL_S2_REV, COL_S2_VP = 65, 66
COL_S3_REV, COL_S3_VP = 67, 68
COL_S4_REV, COL_S4_VP = 69, 70
COL_S5_REV, COL_S5_VP = 71, 72
COL_S1_GROW_N, COL_S1_GROW_P = 73, 74
COL_S3_GROW_N, COL_S3_GROW_P = 75, 76

LAST_COL = 76

# Скрыть от пользователя: С1, С3, С5 и связанные колонки моделирования.
# С2 и С4 остаются видимыми. Формулы и зависимости сохраняются.
HIDDEN_COL_RANGES: list[tuple[int, int]] = [
    (S1[0], S1[-1]),
    (S3[0], S3[-1]),
    (S5[0], S5[-1]),
    (COL_S1_REV, COL_S1_VP),
    (COL_S3_REV, COL_S3_VP),
    (COL_S5_REV, COL_S5_VP),
    (COL_S1_GROW_N, COL_S1_GROW_P),
    (COL_S3_GROW_N, COL_S3_GROW_P),
]

COLUMN_DEFS: list[tuple[int, str, int, str, bool]] = [
    (1, "Наименование", 40, "@", False),
    (2, "Штрихкод", 16, "@", False),
    (3, "Закупочная цена", 14, NUM2, False),
    (4, "Розничная цена ЗЯ", 14, NUM2, False),
    (5, "Наценка сейчас", 13, PCT, False),
    (6, "Новая розничная цена ЗЯ", 16, NUM2, True),
    (7, "Наценка новая", 13, PCT, False),
    (8, "Сигнал цены", 16, "@", False),
    (9, "Ср.цена рынка", 14, NUM2, False),
    (10, "Медиана рынка", 14, NUM2, False),
    (11, "Мин.конкурент", 14, NUM2, False),
    (12, "Макс.конкурент", 14, NUM2, False),
    (13, "Откл.от средней", 13, PCT, False),
    (14, "Откл.от медианы", 13, PCT, False),
    (15, "Откл.от минимума", 13, PCT, False),
    (16, "Рейтинг цены", 11, "0", False),
    (17, "Позиц.на рынке", 14, "@", False),
    (18, "Наценка тек.%", 13, PCT, False),
    (19, "Маржа тек.%", 13, PCT, False),
    (20, "ВП/ед тек.", 13, NUM2, False),
    (21, "Сигнал маржи", 14, "@", False),
    (22, "Риск потери ВП", 14, "@", False),
    (23, "Кол-во продаж", 14, NUM0, False),
    (24, "Выручка тек.", 14, NUM2, False),
    (25, "ВП тек.суммарная", 14, NUM2, False),
    (26, "Маржа тек.общая%", 14, PCT, False),
    (27, "С1:Цена", 12, NUM2, False),
    (28, "С1:Наценка%", 11, PCT, False),
    (29, "С1:Маржа%", 11, PCT, False),
    (30, "С1:ВП/ед", 12, NUM2, False),
    (31, "С1:Дельта цены%", 11, PCT, False),
    (32, "С1:Дельта маржи%", 11, PCT, False),
    (33, "С1:Дельта ВП/ед", 12, NUM2, False),
    (34, "С2:Цена", 12, NUM2, False),
    (35, "С2:Наценка%", 11, PCT, False),
    (36, "С2:Маржа%", 11, PCT, False),
    (37, "С2:ВП/ед", 12, NUM2, False),
    (38, "С2:Дельта цены%", 11, PCT, False),
    (39, "С2:Дельта маржи%", 11, PCT, False),
    (40, "С2:Дельта ВП/ед", 12, NUM2, False),
    (41, "С3:Цена", 12, NUM2, False),
    (42, "С3:Наценка%", 11, PCT, False),
    (43, "С3:Маржа%", 11, PCT, False),
    (44, "С3:ВП/ед", 12, NUM2, False),
    (45, "С3:Дельта цены%", 11, PCT, False),
    (46, "С3:Дельта маржи%", 11, PCT, False),
    (47, "С3:Дельта ВП/ед", 12, NUM2, False),
    (48, "С4:Нов.цена[ВВОД]", 14, NUM2, False),
    (49, "С4:Наценка%", 11, PCT, False),
    (50, "С4:Маржа%", 11, PCT, False),
    (51, "С4:ВП/ед", 12, NUM2, False),
    (52, "С4:Дельта цены%", 11, PCT, False),
    (53, "С4:Дельта маржи%", 11, PCT, False),
    (54, "С4:Дельта ВП/ед", 12, NUM2, False),
    (55, "С5:Цел.маржа[ВВОД]", 14, PCT, True),
    (56, "С5:Цена", 12, NUM2, False),
    (57, "С5:Наценка%", 11, PCT, False),
    (58, "С5:Маржа%", 11, PCT, False),
    (59, "С5:ВП/ед", 12, NUM2, False),
    (60, "С5:Дельта цены%", 11, PCT, False),
    (61, "С5:Дельта маржи%", 11, PCT, False),
    (62, "С5:Дельта ВП/ед", 12, NUM2, False),
    (63, "С1:Прогн.выручка", 14, NUM2, False),
    (64, "С1:Прогн.ВП", 14, NUM2, False),
    (65, "С2:Прогн.выручка", 14, NUM2, False),
    (66, "С2:Прогн.ВП", 14, NUM2, False),
    (67, "С3:Прогн.выручка", 14, NUM2, False),
    (68, "С3:Прогн.ВП", 14, NUM2, False),
    (69, "С4:Прогн.выручка", 14, NUM2, False),
    (70, "С4:Прогн.ВП", 14, NUM2, False),
    (71, "С5:Прогн.выручка", 14, NUM2, False),
    (72, "С5:Прогн.ВП", 14, NUM2, False),
    (73, "С1:Треб.рост шт", 13, NUM0, False),
    (74, "С1:Треб.рост %", 12, PCT, False),
    (75, "С3:Треб.рост шт", 13, NUM0, False),
    (76, "С3:Треб.рост %", 12, PCT, False),
]

_CENTER_COLS = {COL_BARCODE, COL_SIGNAL, COL_POS, COL_MARGIN_SIG, COL_RISK}
_XREF_COLS = {COL_NAME, COL_BARCODE, COL_PURCHASE, COL_RETAIL}


def _L(col: int) -> str:
    return get_column_letter(col)


def _write(ws: Worksheet, row: int, col: int, formula, num_fmt: str, is_input: bool = False) -> None:
    cell = ws.cell(row, col, formula)
    cell.border = border_thin()
    cell.number_format = num_fmt
    if col == COL_NAME:
        cell.alignment = align("left", indent=1)
    elif col in _CENTER_COLS:
        cell.alignment = align("center")
    else:
        cell.alignment = align("right")
    if is_input:
        cell.font = INPUT_FONT
        cell.fill = INPUT_FILL
    elif col in _XREF_COLS:
        cell.font = XREF_FONT
    else:
        cell.font = FORMULA_FONT


def _hide_scenario_columns(ws: Worksheet) -> None:
    """Скрывает колонки С1/С3/С5 (и связанные), оставляя видимыми С2 и С4."""
    for c1, c2 in HIDDEN_COL_RANGES:
        for ci in range(c1, c2 + 1):
            ws.column_dimensions[_L(ci)].hidden = True


def build_calculations_sheet(ws: Worksheet, n_rows: int) -> None:
    """
    Строит лист 'Расчеты' с полным набором формул для n_rows строк товаров.

    :param ws: целевой лист openpyxl.
    :param n_rows: количество строк данных на листе 'Исходные данные' (N).
    """
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = PALETTE.orange

    col_grp_fill: dict[int, str] = {}
    for (c1, c2, label, hexcolor) in CALC_GROUPS:
        for ci in range(c1, c2 + 1):
            col_grp_fill[ci] = hexcolor
        if c2 > c1:
            ws.merge_cells(f"{_L(c1)}1:{_L(c2)}1")
        cell = ws.cell(1, c1, label)
        cell.fill = fill(hexcolor)
        cell.font = font(bold=True, size=9, color=PALETTE.white)
        cell.alignment = align("center")
        cell.border = border_thin()
    ws.row_dimensions[1].height = 18

    for (ci, header, width, num_fmt, is_input) in COLUMN_DEFS:
        cell = ws.cell(2, ci, header)
        cell.fill = fill(col_grp_fill.get(ci, PALETTE.dark_blue))
        cell.font = font(bold=True, size=9, color=PALETTE.white)
        cell.alignment = align("center", wrap=True)
        cell.border = border_thin()
        ws.column_dimensions[_L(ci)].width = width
    ws.row_dimensions[2].height = 36

    src = f"'{SHEETS.source}'"
    sales = f"'{SHEETS.sales}'"

    if n_rows == 0:
        _hide_scenario_columns(ws)
        return

    B = _L(COL_BARCODE)
    C, D, F = _L(COL_PURCHASE), _L(COL_RETAIL), _L(COL_NEW_RETAIL)
    I, J, K = _L(COL_AVG_MKT), _L(COL_MED_MKT), _L(COL_MIN_MKT)
    P = _L(COL_RANK)
    S, T = _L(COL_MARGIN_CUR), _L(COL_VP_CUR)
    W, X, Y = _L(COL_QTY), _L(COL_REV), _L(COL_VP_SUM)

    for ri in range(n_rows):
        r = ri + 3
        sr = ri + 2

        _write(ws, r, COL_NAME, f"={src}!A{sr}", "@")
        _write(ws, r, COL_BARCODE, f"={src}!B{sr}", "@")
        _write(ws, r, COL_PURCHASE, f"={src}!C{sr}", NUM2)
        _write(ws, r, COL_RETAIL, f"={src}!D{sr}", NUM2)

        # Наценка сейчас = (Розн. ЗЯ − Закуп.) / Закуп. — та же база, что у «Наценка тек.%»
        _write(ws, r, COL_MARKUP_NOW,
               f'=IF(AND({C}{r}>0,{D}{r}>0),({D}{r}-{C}{r})/{C}{r},"")', PCT)
        # Новая розничная цена ЗЯ — ручной ввод [ВВОД]; С4 ссылается на эту колонку
        _write(ws, r, COL_NEW_RETAIL, None, NUM2, is_input=True)
        # Наценка новая = (Новая розн. − Закуп.) / Закуп.
        _write(ws, r, COL_MARKUP_NEW,
               f'=IF(AND({C}{r}>0,{F}{r}>0),({F}{r}-{C}{r})/{C}{r},"")', PCT)

        _write(ws, r, COL_SIGNAL,
               f'=IF(OR({D}{r}="",{I}{r}=""),"-",'
               f'IF({D}{r}>{I}{r}*1.05,"Выше рынка",'
               f'IF({D}{r}<{I}{r}*0.95,"Ниже рынка","На уровне")))', "@")

        comp = ",".join(f"{src}!{_L(ci)}{sr}" for ci in range(5, 12))
        _write(ws, r, COL_AVG_MKT, f'=IFERROR(AVERAGE({comp}),"")', NUM2)
        _write(ws, r, COL_MED_MKT, f'=IFERROR(MEDIAN({comp}),"")', NUM2)
        _write(ws, r, COL_MIN_MKT, f'=IFERROR(MIN({comp}),"")', NUM2)
        _write(ws, r, COL_MAX_MKT, f'=IFERROR(MAX({comp}),"")', NUM2)

        for ci, ref in ((COL_DEV_AVG, I), (COL_DEV_MED, J), (COL_DEV_MIN, K)):
            _write(ws, r, ci, f'=IF(AND({D}{r}<>"",{ref}{r}<>""),({D}{r}-{ref}{r})/{ref}{r},"")', PCT)

        _write(ws, r, COL_RANK, f'=IFERROR(RANK({D}{r},({comp})),"")', "0")
        _write(ws, r, COL_POS,
               f'=IF({P}{r}="","-",IF({P}{r}=1,"Дороже всех",'
               f'IF({P}{r}=2,"2-й по цене",IF({P}{r}<=4,"Средний","Ниже конкурентов"))))', "@")

        _write(ws, r, COL_MARKUP_CUR, f'=IF(AND({C}{r}>0,{D}{r}>0),({D}{r}-{C}{r})/{C}{r},"")', PCT)
        _write(ws, r, COL_MARGIN_CUR, f'=IF(AND({C}{r}>0,{D}{r}>0),({D}{r}-{C}{r})/{D}{r},"")', PCT)
        _write(ws, r, COL_VP_CUR, f'=IF(AND({C}{r}>0,{D}{r}>0),{D}{r}-{C}{r},"")', NUM2)

        _write(ws, r, COL_MARGIN_SIG,
               f'=IF({S}{r}="","-",IF({S}{r}<0.1,"Низкая(<10%)",'
               f'IF({S}{r}<0.2,"Средняя(10-20%)","Хорошая(>20%)")))', "@")
        _write(ws, r, COL_RISK,
               f'=IF(OR({K}{r}="",{C}{r}=""),"-",IF(({K}{r}-{C}{r})/{K}{r}<0.05,"КРИТИЧНО",'
               f'IF(({K}{r}-{C}{r})/{K}{r}<0.15,"УМЕРЕННЫЙ","НИЗКИЙ")))', "@")

        # Кол-во продаж подтягивается с листа «Продажи» по штрихкоду (при загрузке файла продаж).
        _write(
            ws, r, COL_QTY,
            f'=IFERROR(VLOOKUP({B}{r},{sales}!A:B,2,FALSE),"")',
            NUM0,
        )
        _write(ws, r, COL_REV, f'=IF({W}{r}>0,{W}{r}*{D}{r},"")', NUM2)
        _write(ws, r, COL_VP_SUM, f'=IF({W}{r}>0,{W}{r}*{T}{r},"")', NUM2)
        _write(ws, r, COL_MARGIN_TOT, f'=IF({X}{r}>0,{Y}{r}/{X}{r},"")', PCT)

        def scenario_block(price_col, mkup_col, mrg_col, vp_col, dprc_col, dmrg_col, dvp_col, price_formula):
            if price_formula is not None:
                _write(ws, r, price_col, price_formula, NUM2)
            price_L = _L(price_col)
            mrg_L = _L(mrg_col)
            vp_L = _L(vp_col)
            _write(ws, r, mkup_col, f'=IF(AND({price_L}{r}<>"",{C}{r}>0),({price_L}{r}-{C}{r})/{C}{r},"")', PCT)
            _write(ws, r, mrg_col, f'=IF(AND({price_L}{r}<>"",{price_L}{r}>0),({price_L}{r}-{C}{r})/{price_L}{r},"")', PCT)
            _write(ws, r, vp_col, f'=IF(AND({price_L}{r}<>"",{C}{r}>0),{price_L}{r}-{C}{r},"")', NUM2)
            _write(ws, r, dprc_col, f'=IF(AND({price_L}{r}<>"",{D}{r}>0),({price_L}{r}-{D}{r})/{D}{r},"")', PCT)
            _write(ws, r, dmrg_col,
                   f'=IF(AND({mrg_L}{r}<>"",{S}{r}<>""),({mrg_L}{r}-{S}{r})/ABS(IF({S}{r}=0,1,{S}{r})),"")', PCT)
            _write(ws, r, dvp_col, f'=IF(AND({vp_L}{r}<>"",{T}{r}<>""),{vp_L}{r}-{T}{r},"")', NUM2)

        scenario_block(*S1, f'=IF({I}{r}<>"",{I}{r},"")')
        scenario_block(*S2, f'=IF({J}{r}<>"",{J}{r},"")')
        scenario_block(*S3, f'=IF({K}{r}<>"",{K}{r},"")')

        # С4:Нов.цена[ВВОД] ← «Новая розничная цена ЗЯ» (формула-ссылка)
        _write(ws, r, S4[0], f'=IF({F}{r}<>"",{F}{r},"")', NUM2)
        scenario_block(S4[0], S4[1], S4[2], S4[3], S4[4], S4[5], S4[6], None)

        # С5: целевая маржа [ВВОД] + производные
        s5_in, s5_price = _L(S5[0]), _L(S5[1])
        s5_mrg, s5_vp = _L(S5[3]), _L(S5[4])
        _write(ws, r, S5[0], None, PCT, is_input=True)
        _write(ws, r, S5[1], f'=IF(AND({s5_in}{r}>0,{C}{r}>0),{C}{r}/(1-{s5_in}{r}),"")', NUM2)
        _write(ws, r, S5[2], f'=IF(AND({s5_price}{r}<>"",{C}{r}>0),({s5_price}{r}-{C}{r})/{C}{r},"")', PCT)
        _write(ws, r, S5[3], f'=IF({s5_price}{r}<>"",{s5_in}{r},"")', PCT)
        _write(ws, r, S5[4], f'=IF(AND({s5_price}{r}<>"",{C}{r}>0),{s5_price}{r}-{C}{r},"")', NUM2)
        _write(ws, r, S5[5], f'=IF(AND({s5_price}{r}<>"",{D}{r}>0),({s5_price}{r}-{D}{r})/{D}{r},"")', PCT)
        _write(ws, r, S5[6],
               f'=IF(AND({s5_mrg}{r}<>"",{S}{r}<>""),({s5_mrg}{r}-{S}{r})/ABS(IF({S}{r}=0,1,{S}{r})),"")', PCT)
        _write(ws, r, S5[7], f'=IF(AND({s5_vp}{r}<>"",{T}{r}<>""),{s5_vp}{r}-{T}{r},"")', NUM2)

        s1p, s1vp = _L(S1[0]), _L(S1[3])
        s2p, s2vp = _L(S2[0]), _L(S2[3])
        s3p, s3vp = _L(S3[0]), _L(S3[3])
        s4p, s4vp = _L(S4[0]), _L(S4[3])
        s5p, s5vp_l = _L(S5[1]), _L(S5[4])
        g1n, g3n = _L(COL_S1_GROW_N), _L(COL_S3_GROW_N)

        for ci, expr, nf in (
            (COL_S1_REV, f'=IF(AND({W}{r}>0,{s1p}{r}<>""),{W}{r}*{s1p}{r},"")', NUM2),
            (COL_S1_VP, f'=IF(AND({W}{r}>0,{s1vp}{r}<>""),{W}{r}*{s1vp}{r},"")', NUM2),
            (COL_S2_REV, f'=IF(AND({W}{r}>0,{s2p}{r}<>""),{W}{r}*{s2p}{r},"")', NUM2),
            (COL_S2_VP, f'=IF(AND({W}{r}>0,{s2vp}{r}<>""),{W}{r}*{s2vp}{r},"")', NUM2),
            (COL_S3_REV, f'=IF(AND({W}{r}>0,{s3p}{r}<>""),{W}{r}*{s3p}{r},"")', NUM2),
            (COL_S3_VP, f'=IF(AND({W}{r}>0,{s3vp}{r}<>""),{W}{r}*{s3vp}{r},"")', NUM2),
            (COL_S4_REV, f'=IF(AND({W}{r}>0,{s4p}{r}>0),{W}{r}*{s4p}{r},"")', NUM2),
            (COL_S4_VP, f'=IF(AND({W}{r}>0,{s4vp}{r}<>""),{W}{r}*{s4vp}{r},"")', NUM2),
            (COL_S5_REV, f'=IF(AND({W}{r}>0,{s5p}{r}<>""),{W}{r}*{s5p}{r},"")', NUM2),
            (COL_S5_VP, f'=IF(AND({W}{r}>0,{s5vp_l}{r}<>""),{W}{r}*{s5vp_l}{r},"")', NUM2),
            (COL_S1_GROW_N, f'=IF(AND({W}{r}>0,{Y}{r}>0,{s1vp}{r}<>0),{Y}{r}/{s1vp}{r}-{W}{r},"")', NUM0),
            (COL_S1_GROW_P, f'=IF(AND({g1n}{r}<>"",{W}{r}>0),{g1n}{r}/{W}{r},"")', PCT),
            (COL_S3_GROW_N, f'=IF(AND({W}{r}>0,{Y}{r}>0,{s3vp}{r}<>0),{Y}{r}/{s3vp}{r}-{W}{r},"")', NUM0),
            (COL_S3_GROW_P, f'=IF(AND({g3n}{r}<>"",{W}{r}>0),{g3n}{r}/{W}{r},"")', PCT),
        ):
            _write(ws, r, ci, expr, nf)

    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{_L(LAST_COL)}{n_rows + 2}"

    last_data_row = n_rows + 2
    # Условное форматирование отклонений от средней и медианы — один стиль
    for col_idx in (COL_DEV_AVG, COL_DEV_MED):
        letter = _L(col_idx)
        ws.conditional_formatting.add(
            f"{letter}3:{letter}{last_data_row}",
            ColorScaleRule(start_type="min", start_color="63BE7B",
                            mid_type="num", mid_value=0, mid_color="FFEB84",
                            end_type="max", end_color="F8696B"),
        )
    ws.conditional_formatting.add(
        f"{S}3:{S}{last_data_row}",
        ColorScaleRule(start_type="min", start_color="F8696B",
                        mid_type="percentile", mid_value=50, mid_color="FFEB84",
                        end_type="max", end_color="63BE7B"),
    )
    ws.conditional_formatting.add(
        f"{T}3:{T}{last_data_row}",
        DataBarRule(start_type="min", end_type="max", color="4472C4"),
    )

    _hide_scenario_columns(ws)
