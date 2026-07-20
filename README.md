# ZA Price Calculator — Python модуль

Production-ready модуль для сети супермаркетов «Зеленое Яблоко».
Переносит логику Excel-калькулятора цен/наценки/маржи/сценариев в код на Python.

## Архитектура

```
za_price_calculator/
├── __init__.py
├── config.py                    # Константы: названия колонок, алиасы, цветовая палитра, пороги
├── exceptions.py                 # ZAPriceCalculatorError, FileLoadError, ValidationError, SheetBuildError
├── builder.py                    # Оркестратор: собирает все листы в Workbook, сохраняет файл
├── service.py                    # Публичный API: класс ZAPriceCalculator.run(...)
├── cli.py                        # CLI точка входа: python -m za_price_calculator.cli
│
├── io_handlers/
│   ├── __init__.py
│   └── loader.py                 # load_source_file(), load_sales_file() — чтение и очистка Excel
│
├── styling/
│   ├── __init__.py
│   └── styles.py                 # Font/Fill/Alignment/Border объекты, форматы чисел
│
├── core/
│   ├── __init__.py
│   ├── sheet_instructions.py     # Лист "Инструкция"
│   ├── sheet_source.py           # Лист "Исходные данные" (Excel Table tbl_Src)
│   ├── sheet_sales.py            # Лист "Продажи" (Excel Table tbl_Sales, опционально)
│   ├── sheet_calculations.py     # Лист "Расчеты" — 73 колонки формул (ядро калькулятора)
│   ├── sheet_dashboard.py        # Лист "Dashboard" — KPI-карточки, топ-10, светофоры
│   └── sheet_scenarios.py        # Лист "Сценарный анализ" — сводная таблица + автосигналы
│
└── tests/
    └── __init__.py
```

## Установка зависимостей

```bash
pip install pandas openpyxl
```

## Использование в PyCharm

```python
from za_price_calculator.service import ZAPriceCalculator

calc = ZAPriceCalculator()
result_path = calc.run(
    source_path="прайс_лист.xlsx",       # обязательный файл с товарами и ценами
    output_path="output/ZA_Calc.xlsx",   # куда сохранить готовый калькулятор
    sales_path="продажи.xlsx",           # опционально: файл с реальными продажами
)
print(f"Готово: {result_path}")
```

## Использование через CLI

```bash
python -m za_price_calculator.cli --source прайс.xlsx --output out.xlsx
python -m za_price_calculator.cli --source прайс.xlsx --sales продажи.xlsx --output out.xlsx -v
```

## Требования к входному файлу (--source)

Обязательные колонки (порядок не важен, поддерживаются синонимы):
- Наименование / Название / Товар
- Штрихкод / Штрих-код / Ш/к
- Закупочная цена
- Розничная цена ЗЯ / Зеленое яблоко
- Цены конкурентов: 7 Континент, Европейский, Тропики, Мята, Отличный, Оптовик (молоток), Оптовик (редук)

## Опциональный файл продаж (--sales)

Колонки: Штрихкод, Кол-во продаж, Выручка, Валовая прибыль.
При отсутствии файла калькулятор строит ценовой блок без ошибок — поля продаж остаются пустыми,
доступен ручной ввод количества продаж прямо на листе "Расчеты" (колонка T).

## Результат

Итоговый .xlsx с 6 листами: Инструкция, Исходные данные, Продажи, Расчеты, Dashboard, Сценарный анализ.
Все формулы, условное форматирование, Excel Tables и цветовое оформление идентичны
эталонному ZA_Price_Calculator.xlsx.

## Обработка ошибок

- FileLoadError — файл не найден или неверный формат
- ValidationError — отсутствуют обязательные колонки или нет валидных строк после очистки
- SheetBuildError — ошибка при построении листов книги

Все исключения наследуются от ZAPriceCalculatorError и логируются через модуль logging.
