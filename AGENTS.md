# AGENTS.md

## Cursor Cloud specific instructions

### What this is
`za_price_calculator` is a standalone Python CLI/library (no database or external services). It reads an input price-list `.xlsx` and generates a formatted 6-sheet Excel workbook. `streamlit_app.py` at the repo root is an optional web front-end (for Streamlit Community Cloud) that wraps the same `ZAPriceCalculator` API. Runtime deps: `pandas`, `openpyxl`, `streamlit` (installed by the update script; also listed in `requirements.txt`).

### Streamlit web app
`streamlit_app.py` is a Russian-language AI-agent front-end for the «Зеленое Яблоко» chain: it generates the Excel workbook AND renders an in-app BI dashboard (KPI cards, Altair charts, scenario comparison, auto-generated conclusions). Dashboard metrics are recomputed in pandas mirroring the module's methodology (наценка = ВП/закупка, маржа = ВП/цена, отклонение от средней/медианы рынка, сценарии С2/С4, пороги из `config.THRESHOLDS`). UI label: **Релиз 3.5**.

Run locally: `python -m streamlit run streamlit_app.py` (the `streamlit` console script installs to `~/.local/bin`, which is not on PATH, so prefer `python -m streamlit`). Deploy on Streamlit Community Cloud by pointing it at `streamlit_app.py`; deps come from `requirements.txt`.

The optional sales file uploads via the "Файл продаж" block; required column is `Штрихкод` (others: Кол-во продаж, Выручка, Валовая прибыль), matched to the price list by barcode. On the Excel «Расчеты» sheet, «Кол-во продаж» is filled via `VLOOKUP` from the «Продажи» sheet. A ready-made template lives at `шаблон_продаж_ЗЯ.xlsx` (repo root) and is also downloadable in-app via the sidebar button (`_sales_template_bytes()`); keep the two in sync if you change the headers.

Gotchas when editing `streamlit_app.py`:
- The module loader only accepts file *paths* (checks `Path.exists()`/suffix), so uploads are written to a temp dir before calling the API — keep that pattern.
- Avoid pandas `Styler` (`df.style.format`) — it requires `jinja2`, which is not a dependency; format values to strings and use plain `st.dataframe` instead.
- Altair treats `.` in a field name as nested access, so keep chart column names dot-free (e.g. `Маржа`, not `Ср. маржа`).

### Running (hello-world / end-to-end)
Run from the repo root (the package uses absolute `za_price_calculator.*` imports, so the root must be on `sys.path`):

```bash
python -m za_price_calculator.cli --source sample_source.xlsx --output output/test.xlsx -v
```

`sample_source.xlsx` is a ready-made input in the repo. `--sales` is optional; omitting it leaves sales fields blank (not an error). Programmatic API: `ZAPriceCalculator().run(source_path=..., output_path=..., sales_path=...)` in `za_price_calculator/service.py`.

### Gotchas
- Generated sheet names carry emoji prefixes (e.g. `🔢 Расчеты`, `📈 Dashboard`), so look up sheets by the exact emoji-prefixed name, not the bare Russian word.
- Committed `.pyc` files under `__pycache__/` were built for CPython 3.10; the VM runs Python 3.12, which transparently recompiles from the `.py` sources — ignore the stale 3.10 pyc files.
- `za_price_calculator/tests/` contains only an empty `__init__.py` — there is no automated test suite. Lint/syntax can be checked with `python -m compileall za_price_calculator`.
- Loader warnings like "N строк без закупочной цены" are expected for the sample data (rows with missing prices) and are non-fatal.
