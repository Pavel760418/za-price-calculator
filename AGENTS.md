# AGENTS.md

## Cursor Cloud specific instructions

### What this is
`za_price_calculator` is a standalone Python CLI/library (no web server, database, or external services). It reads an input price-list `.xlsx` and generates a formatted 6-sheet Excel workbook. Runtime deps: `pandas`, `openpyxl` (installed by the update script; also listed in `requirements.txt`).

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
