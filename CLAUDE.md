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
# Step 1 — authoritative Form 1325 support workbook (run first)
python generate_1325_support.py --year 2024 [--ibkr-pdf PATH]
# → also auto-generates form1325_2024_25pct.pdf (non-fatal if PDF generation fails)

# Step 1b (optional) — generate / regenerate Form 1325 PDF separately
python generate_1325_pdf.py --year 2024 [--config taxpayer.json]

# Step 2 — annual filing summary (reads Tax Data from step 1)
python annual_tax_summary.py --year 2024 [--ibkr-pdf PATH]

# Step 3 (optional) — per-sale EquatePlus diagnostic worksheets (audit only)
python tax_calculator.py

# Tests
pytest test_tax.py -v -m "not integration"                    # unit tests, no source files needed
pytest tests/integration/test_2024_pipeline.py -v             # integration tests (need fixtures)
pytest tests/integration/test_2022_pdf.py -v                  # PDF generation regression test
```

## Architecture

Single-path capital-gains architecture — one implementation of each algorithm, one authoritative pipeline.

```
IBKR *_tax_statement.pdf          EquatePlus CSVs + PDFs
        ↓                                   ↓
ibkr_parser.py                    equate_parser.py
(raw broker facts only)           (raw broker facts only)
        ↓                                   ↓
              capital_gains.py
              (ONLY place that does: FX conversion, commission allocation,
               adjusted cost, Moses application)
              (uses tax_utils for math/FX primitives)
                          ↓
              generate_1325_support.py          tax_calculator.py
              (authoritative filing pipeline)    (diagnostic per-sale worksheets)
                          ↓                              ↓
              form1325_support_<year>.xlsx    consumption_*_with_calc.xlsx
                ├─ Form 1325 Entry View
                ├─ Form 1325 Support
                ├─ Tax Data          (machine-readable, one row per lot)
                ├─ Rounding & Cross-Check
                ├─ Metadata
                └─ Sources           (one row per source file + SHA-256)
                          ↓
              generate_1325_pdf.py
              (presentation layer only — reads Form 1325 Entry View)
              (templates from templates/1325/, taxpayer identity from taxpayer.json)
                          ↓
              form1325_<year>_25pct.pdf
                          ↓
              annual_tax_summary.py
              (reads Tax Data — no Moses, no source re-parsing)
              (reads IBKR PDF for dividends/interest/WHT only)
                          ↓
              tax_annual_summary_<year>.xlsx
                ├─ Annual Summary (1301/1322/1324 filing rows)
                └─ Provenance    (SHA-256 of consumed 1325 workbook)
```

### Module responsibilities

**`tax_utils.py`** — shared math/FX/tax utilities
- `moses_gain_loss(original_cost_ils, net_sale_ils, fx_buy, fx_sell) → (gain, loss)` — single authoritative Moses implementation
- `_round_ils(value) → int` — ROUND_HALF_UP; single authoritative rounding implementation
- `boi_ils(currency, day) → float` — BoI SDMX API, 7-day fallback, cached
- `TAX_RATE = 0.25`

**`capital_gains.py`** — single calculation layer
- `build_ibkr_lots(raw_lots) → list[dict]` — fetches FX, allocates commission (signed, negative=cost), calls Moses
- `build_equate_lots(raw_lots) → list[dict]` — same for EquatePlus (commission positive=cost)
- No other module may do acquisition-basis→ILS, sale-proceeds→ILS, commission allocation, adjusted-cost, or Moses

**`ibkr_parser.py`** — IBKR PDF parsing + income extraction
- `parse_ibkr_lots_detail(text) → list[dict]` — raw facts; assigns `fill_id`/`lot_id`; adds fill-level totals (`fill_quantity_fc`, `fill_proceeds_fc`, `fill_commission_fc`) for allocation verification
- `fill_id = "IBKR/{date}/{sym}/{exch}/{fill_seq}"` — shared by all lots in one fill
- `lot_id  = "{fill_id}/LOT/{lot_seq}"` — unique per Tax Data row
- Income parsers: `parse_dividends_by_country`, `parse_interest_transactions_ils`, `parse_withholding_tax_ils`, `parse_ticker_tax_country`

**`equate_parser.py`** — EquatePlus CSV+PDF parsing
- `parse_equateplus_lots(year) → list[dict]` — year-filtered; raises RuntimeError on missing sale PDF; validates execution date against filename (±1 day OK, >1 day raises)
- `parse_equateplus_pair(csv_path, sale_pdf) → list[dict]` — single pair, for use by tax_calculator.py
- `sale_id = "EQ/{date}/{ticker}/{sale_seq}"` — sale_seq prevents collision when same ticker sells twice same day
- `lot_id  = "{sale_id}/LOT/{acq_date_str}"` (with `/{seq}` suffix if acq_date repeats)
- Raw lots carry `fill_quantity_fc`, `fill_proceeds_fc`, `fill_commission_fc` for `_verify_allocation()`

**`generate_1325_support.py`** — authoritative filing pipeline
- CLI: `--year YYYY`, `--ibkr-pdf PATH`
- Validates IBKR PDF filename year matches `--year`
- `_verify_allocation()` — genuine independent check: lot qty/proceeds/commission sums vs fill-level totals (both IBKR and EquatePlus)
- Sources sheet: only records files that actually produced Tax Data rows
- Raises on lot_id duplicates (RuntimeError, not assert)

**`annual_tax_summary.py`** — annual filing; reads Tax Data only
- `load_tax_data(year)` — validates schema_version, tax_year, required columns, lot_id uniqueness
- `_verify_ibkr_pdf_hash()` — checks IBKR PDF sha256 against Sources sheet before parsing income
- Aggregates by `(source, tax_rate)`; raises RuntimeError on any rate outside `{0.25}`
- Never re-parses source files for capital gains

**`generate_1325_pdf.py`** — Form 1325 PDF generator (presentation layer only)
- CLI: `--year YYYY`, `--xlsx PATH`, `--template PATH`, `--signature PATH`, `--output-dir PATH`, `--no-signature`, `--signature-date DD/MM/YYYY`, `--taxpayer-name TEXT`, `--file-number TEXT`, `--config PATH`, `--debug`
- Reads `Form 1325 Entry View` exclusively — no tax calculations, no re-parsing of source files
- Templates in `templates/1325/`; calibrated for years 2020–2025; raises for uncalibrated years
- Taxpayer identity: CLI args → `taxpayer.json` (git-ignored) → raise; never inferred from XLSX
- `taxpayer.json` format: `{"taxpayer_name": "קופרמן סרגיי", "file_number": "313985129"}` — normal Unicode Hebrew (BiDi reordering applied internally via `python-bidi`)
- Groups data rows by `(form_no, rate_int_pct)`; one PDF per group; max 10 rows per form
- Called non-fatally by `generate_1325_support.py` after XLSX write
- `generate_pdfs(year, xlsx_path, ...) → list[str]` — public API

**`tax_calculator.py`** — diagnostic per-sale worksheets
- Uses `parse_equateplus_pair()` + `build_equate_lots()` — no duplicate raw-lot construction
- Output is audit-only; not consumed by annual_tax_summary.py

## Tax law context

**ITO Section 91(b) + Circular 10/2025 — Moses rule:** Foreign-currency securities use the BoI FX rate as an inflation index (מדד). Four cases:
1. FX up, price up → taxable gain = nominal gain − inflationary component
2. FX up, price down → deductible loss = nominal loss (inflationary component cannot enlarge a loss)
3. FX up, nominal gain but all inflationary → taxable = 0, loss = 0
4. FX down → negative inflationary component treated as zero (cannot enlarge gain or convert loss)

**Rounding policy:**
- Per Form 1325 row: `_round_ils(exact_value)` — ROUND_HALF_UP
- Annual gains/losses (1322): sum of per-row rounded values (bottom-up)
- Annual gross turnover (1322): `_round_ils(sum of exact proceeds)` — rounded once (top-down)

## Tax Data schema (v1)

One row per lot in Tax Data sheet. Key columns:

| Column | Notes |
|--------|-------|
| `schema_version` | `"1"` |
| `tax_year` | e.g. `"2024"` |
| `source` | `"IBKR"` or `"EquatePlus"` |
| `fill_id` | Shared by all lots from same broker execution |
| `lot_id` | Unique per row — primary key |
| `tax_rate` | `0.25` |
| `gross_sale_proceeds_ils_exact` | No `_filing` column — 1322 uses `_round_ils(sum_exact)` |
| `taxable_gain_ils_filing` | `_round_ils(taxable_gain_ils_exact)` |
| `deductible_loss_ils_filing` | `_round_ils(deductible_loss_ils_exact)` |

## Input files

**IBKR:** `<account>_<YYYYMMDD>_<YYYYMMDD>_tax_statement.pdf` in repo root

**EquatePlus:** `consumption_D.M.YYYY.csv` + `sale_D.M.YYYY.pdf` pairs in repo root
- CSV: semicolon-delimited, comma-decimal, columns: `Acquisition date` (DD Mon YYYY), `Consumption` (shares), `Purchase price` (EUR/share)

## Integration test fixtures

Real 2024 source files live in `tests/integration/fixtures/2024/` (git-ignored — personal financial data).
See `tests/integration/fixtures/2024/README.md` for layout.

Frozen 2024 regression values (verified 2026-08-26):
- IBKR lots: 7, EquatePlus lots: 70
- Gains @ 25%: 129,210 ILS, Losses @ 25%: 1,861 ILS, Gross turnover: 333,313 ILS
