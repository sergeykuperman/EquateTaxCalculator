# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Running

```bash
python tax_calculator.py       # process all consumption_*.csv / sale_*.pdf pairs in the repo root
python compound_calc.py        # standalone portfolio projection → portfolio_projection_5y.xlsx
```

There are no tests and no linter configured.

## Architecture

**`tax_calculator.py`** is the main script. It:

1. Scans the repo root for `consumption_<D.M.YYYY>.csv` files, pairs each with a matching `sale_<D.M.YYYY>.pdf`.
2. **`parse_sale_pdf()`** — extracts sale price (EUR), settlement date, PDF-embedded FX rate, and fees from the PDF using `pdfplumber` + regex.
3. **`boi_eur_ils()`** — fetches the Bank of Israel (BoI) EUR/ILS representative rate for a given date via the BoI SDMX API, walking back up to 7 days for weekends/holidays. Also supports a `live=True` mode via the public BoI exchange-rates API.
4. **`process_pair()`** — for each CSV lot row, fetches the BoI rate at acquisition date (`fx_acq`) and at settlement date (`fx_set`). Computes:
   - `gross_sale_shekel = shares × sale_price_EUR × fx_set`
   - `cost_shekel = shares × purchase_price_EUR × fx_acq`
   - `real_gain_shekel = gross_sale - cost - proportional_fees`
   - Tax = `max(0, total_real_gain) × 25%`
   - Writes a per-sale `consumption_<date>_with_calc.xlsx` with Data + Summary sheets.
5. **`main()`** — after all pairs, if there are ≥2 sales it computes an **annual net gain** (losses offset gains across all sales) and writes a combined `tax_summary_<years>.xlsx`.

**Key tax law context (ITO Section 91(b)):** Foreign-currency securities use the BoI FX rate method instead of CPI. The inflationary component (`cost × (fx_set/fx_acq − 1)`) is tax-exempt; only the real gain above FX movement is taxable. This is why the script uses two BoI rates per lot rather than a CPI index.

**Input files** (must be in repo root):
- `sale_D.M.YYYY.pdf` — EquatePlus sale summary PDF
- `consumption_D.M.YYYY.csv` — semicolon-delimited, comma-decimal CSV with columns including `Acquisition date`, `Consumption` (shares), `Purchase price`

**`compound_calc.py`** is an unrelated one-off script that projects a portfolio balance over 72 months at 12% annual return. It has no connection to the tax calculator.
