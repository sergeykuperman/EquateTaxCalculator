# EquateTaxCalculator

Calculates Israeli capital-gains tax (ITO Section 91(b) + Circular 10/2025 Moses rule) for EquatePlus RSU sales and IBKR stock disposals. Produces Form 1325 support workbooks and annual filing summaries in ILS.

## Requirements

- Python 3.12+
- pip / venv

## Setup

```bash
git clone https://github.com/sergeykuperman/EquateTaxCalculator.git
cd EquateTaxCalculator
python3 -m venv venv
source venv/bin/activate          # macOS / Linux
pip install --upgrade pip
pip install -r requirements.txt
```

## Input files (place in repo root)

| File | Source | Description |
|------|--------|-------------|
| `U*_tax_statement.pdf` | IBKR | Annual Activity Statement (Tax) PDF, e.g. `U3484618_20240101_20241231_tax_statement.pdf` |
| `consumption_D.M.YYYY.csv` | EquatePlus | Semicolon-delimited lot CSV with `Acquisition date`, `Consumption`, `Purchase price` |
| `sale_D.M.YYYY.pdf` | EquatePlus | Matching sale summary PDF for each consumption CSV |

---

## Workflow

### Step 1 — Generate authoritative Form 1325 support workbook

```bash
python generate_1325_support.py --year 2024
# or supply the IBKR PDF explicitly:
python generate_1325_support.py --year 2024 --ibkr-pdf U3484618_20240101_20241231_tax_statement.pdf
```

Produces **`form1325_support_2024.xlsx`** with six sheets:

| Sheet | Purpose |
|-------|---------|
| Form 1325 Entry View | Manual transcription view — 10 rows/form + check subtotals |
| Form 1325 Support | Full per-lot audit detail (exact floats + filing-rounded values) |
| Tax Data | Machine-readable, one row per lot, schema v1 (input for Step 2) |
| Rounding & Cross-Check | Aggregation-rule verification + delta vs prior run |
| Metadata | Generation provenance (timestamp, methodology, rounding policy) |
| Sources | One row per source file with SHA-256 hash |

**Verify before proceeding:** open "Rounding & Cross-Check" — all "Aggregation rule satisfied?" must be True.

### Step 1b — Generate Form 1325 PDF (auto-run by Step 1, or separately)

Step 1 automatically generates the PDF after writing the XLSX. To run it separately or regenerate:

```bash
python generate_1325_pdf.py --year 2024
```

Requires `taxpayer.json` in the repo root (git-ignored — never commit):
```json
{"taxpayer_name": "Your Name In Hebrew", "file_number": "123456789"}
```

Or pass identity on the CLI:
```bash
python generate_1325_pdf.py --year 2024 --taxpayer-name "קופרמן סרגיי" --file-number "313985129"
```

Produces **`form1325_<year>_25pct.pdf`** — the official Form 1325 filled with all Entry View values, ready for submission.

Blank official templates live in `templates/1325/` for years 2020–2025.

Other options:
```
--no-signature          skip inserting signature.jpg
--signature-date DATE   override signature date (default: today, DD/MM/YYYY)
--output-dir PATH       output directory (default: .)
--debug                 emit debug PDF with labelled red field-anchor boxes
```

### Step 2 — Generate annual filing summary

```bash
python annual_tax_summary.py --year 2024
# or supply the IBKR PDF explicitly:
python annual_tax_summary.py --year 2024 --ibkr-pdf U3484618_20240101_20241231_tax_statement.pdf
```

Reads **`form1325_support_2024.xlsx`** (Tax Data sheet only — no re-parsing of source files).
Adds dividend, interest, and withholding tax from the IBKR PDF.
Verifies the IBKR PDF hash against the Sources sheet.

Produces **`tax_annual_summary_2024.xlsx`** with:
- **Annual Summary** — all filing values (1301/1322/1324) in ILS
- **Provenance** — SHA-256 of the consumed 1325 workbook

### Step 3 (optional) — Per-sale EquatePlus diagnostic worksheets

```bash
python tax_calculator.py
```

Writes `consumption_<date>_with_calc.xlsx` (Data + Summary sheets) and `tax_summary_<year>.xlsx` for per-sale audit. These outputs are **not** used by the annual summary.

---

## Annual Summary output rows

| Row | Form | Description |
|-----|------|-------------|
| Interest income | 1301 | IBKR SYEP interest, converted via BoI rate |
| Dividend spread by country | 1301 | Per-country ILS totals |
| Dividend total | 1301 | All dividends + payments in lieu |
| Dividend + external income total | 1301 | Dividends + interest |
| Foreign withholding tax | 1324 | Tax paid abroad (positive), from IBKR Cash Report |
| IBKR taxable realized stock gains | 1322 | Moses/Form 1325 gains @ 25% |
| IBKR deductible realized stock losses | 1322 | Moses/Form 1325 losses (positive) @ 25% |
| IBKR gross sale value (stock disposals) | 1322 | Turnover (מחזור מכירות) |
| EquatePlus taxable realized gains | 1322 | Moses/Form 1325 gains @ 25% |
| EquatePlus deductible realized losses | 1322 | Moses/Form 1325 losses (positive) @ 25% |
| EquatePlus gross sale value | 1322 | Turnover |
| TOTAL taxable gains | 1322 | IBKR + EquatePlus combined |
| TOTAL deductible losses | 1322 | IBKR + EquatePlus combined |
| TOTAL gross sale value | 1322 | IBKR + EquatePlus combined |

All FX conversions use the Bank of Israel representative EUR/ILS (or USD/ILS) rate, walking back up to 7 days for weekends and holidays (ITO Section 91(b)).

---

## Tax methodology

**Moses rule (ITO Section 91(b) + Circular 10/2025):** Foreign-currency securities use the BoI FX rate as an inflation index (מדד). Four cases:

| FX | Price | Result |
|----|-------|--------|
| Up | Up | Taxable gain = nominal gain − inflationary component |
| Up | Down | Deductible loss = nominal loss (inflationary component cannot enlarge) |
| Up | Down (nominal gain) | Taxable = 0, Loss = 0 (fully inflationary) |
| Down | Any | Negative inflationary component treated as zero |

**Rounding:** `ROUND_HALF_UP` per Form 1325 row. Annual gains/losses = sum of rounded rows (bottom-up). Annual gross turnover = `round(sum of exact proceeds)` once (top-down).

---

## Tests

```bash
# Unit tests (no source files needed):
pytest test_tax.py -v -m "not integration"

# Integration tests (require 2024 source files in tests/integration/fixtures/2024/):
pytest tests/integration/test_2024_pipeline.py -v

# PDF generation regression test (requires form1325_support_2022.xlsx + 2022 template):
pytest tests/integration/test_2022_pdf.py -v
```

Integration fixture layout (excluded from git — contains personal financial data):
```
tests/integration/fixtures/2024/
  ibkr/   U3484618_20240101_20241231_tax_statement.pdf
  equate/ consumption_*.csv  sale_*.pdf
```

See `tests/integration/fixtures/2024/README.md` for setup instructions.

---

## Module overview

| Module | Role |
|--------|------|
| `tax_utils.py` | `moses_gain_loss()`, `_round_ils()`, `boi_ils()` — single authoritative implementations |
| `ibkr_parser.py` | IBKR PDF parsing; returns raw broker facts + income parsers |
| `equate_parser.py` | EquatePlus CSV+PDF parsing; returns raw broker facts |
| `capital_gains.py` | FX conversion, commission allocation, adjusted cost, Moses application |
| `generate_1325_support.py` | Authoritative filing pipeline → `form1325_support_<year>.xlsx` |
| `generate_1325_pdf.py` | Presentation layer: fills official Form 1325 PDF from Entry View → `form1325_<year>_25pct.pdf` |
| `annual_tax_summary.py` | Reads Tax Data; adds income; → `tax_annual_summary_<year>.xlsx` |
| `tax_calculator.py` | Diagnostic per-sale worksheets (audit only) |
| `compound_calc.py` | Standalone portfolio projection (unrelated to tax calculator) |
