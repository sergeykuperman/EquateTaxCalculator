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
4. **`process_pair()`** — for each CSV lot row, fetches the BoI rate at acquisition date (`fx_acq`) and at sale/execution date (`fx_sell`). Applies the Moses/Form 1325 4-case algorithm per lot:
   - `original_cost_shekel = shares × purchase_price_EUR × fx_acq`
   - `gross_sale_shekel = shares × sale_price_EUR × fx_sell` (turnover for Form 1322)
   - `net_sale_shekel = gross_sale_shekel - allocated_sale_fee`
   - `adjusted_cost_shekel = original_cost_shekel × (fx_sell / fx_acq)`
   - Per lot: `taxable_gain_shekel` and `deductible_loss_shekel` via `moses_gain_loss()` (see ITO Section 91(b) + Circular 10/2025)
   - Summary columns per sale: `Total_gross_sale_shekel`, `Total_taxable_gain`, `Total_deductible_loss`, `Total_net_gain`, `Total_tax_to_pay = max(0, net) × 25%`
   - Writes a per-sale `consumption_<date>_with_calc.xlsx` with Data + Summary sheets.
5. **`main()`** — after all pairs, writes a combined `tax_summary_<years>.xlsx` with one row per sale plus a TOTAL row. The TOTAL row sums `Total_taxable_gain` and `Total_deductible_loss` across all sales; tax is re-computed as `max(0, net_annual_gain) × 25%` so losses offset gains before the 25% rate applies.

**Key tax law context (ITO Section 91(b) + Circular 10/2025 Moses rule):** Foreign-currency securities use the BoI FX rate as an index (מדד). When FX rises the inflationary component is exempt; when FX falls the negative inflationary component is treated as zero (cannot enlarge a gain or convert a loss into a deductible amount). The `moses_gain_loss()` function implements the full 4-case algorithm with unit tests runnable via `python tax_calculator.py --test`.

**Input files** (must be in repo root):
- `sale_D.M.YYYY.pdf` — EquatePlus sale summary PDF
- `consumption_D.M.YYYY.csv` — semicolon-delimited, comma-decimal CSV with columns including `Acquisition date`, `Consumption` (shares), `Purchase price`

**`compound_calc.py`** is an unrelated one-off script that projects a portfolio balance over 72 months at 12% annual return. It has no connection to the tax calculator.
