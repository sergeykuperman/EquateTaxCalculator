#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_1325_pdf.py — Presentation layer: places Form 1325 Entry View filing
values onto the official blank Form 1325 PDF template.

No tax calculations are performed here. All values consumed are already-rounded
filing values produced by generate_1325_support.py.

Usage:
    python generate_1325_pdf.py --year 2022 [options]

Options:
    --xlsx PATH            default: form1325_support_<year>.xlsx
    --template PATH        default: from DEFAULT_TEMPLATES[year]
    --signature PATH       default: signature.jpg
    --output-dir PATH      default: .
    --no-signature
    --signature-date DD/MM/YYYY    default: today
    --taxpayer-name TEXT   overrides taxpayer.json
    --file-number TEXT     overrides taxpayer.json
    --config PATH          default: taxpayer.json
    --debug                emit debug PDF with field anchors
"""

import argparse
import json
import math
import os
import datetime
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
import pymupdf

# ─── Template registry ───────────────────────────────────────────────────────
# Keyed by tax year (int). Raises immediately for uncalibrated years.
# Paths are relative to the templates/1325/ directory alongside this module.

_TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "templates", "1325")

DEFAULT_TEMPLATES = {
    2020: os.path.join(_TEMPLATES_DIR, "Service_Pages_Income_tax_annual-report-2020_1325-2020.pdf"),
    2021: os.path.join(_TEMPLATES_DIR, "Service_Pages_Income_tax_annual-report-2021_1325-2021.pdf"),
    2022: os.path.join(_TEMPLATES_DIR, "Service_Pages_Income_tax_annual-report-2022_1325-2022.pdf"),
    2023: os.path.join(_TEMPLATES_DIR, "Service_Pages_Income_tax_annual-report-2023_1325-2023.pdf"),
    2024: os.path.join(_TEMPLATES_DIR, "Service_Pages_Income_tax_annual-report-2024_itc1325-2024.pdf"),
    2025: os.path.join(_TEMPLATES_DIR, "Service_Pages_Income_tax_annual-report-2026_1325-2025.pdf"),
}

# ─── Hebrew font candidates ───────────────────────────────────────────────────

HEBREW_FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/noto/NotoSansHebrew-Regular.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]

# ─── Coordinate map (year 2022) ───────────────────────────────────────────────
# All coordinates from PyMuPDF span["origin"] / span["bbox"] on 1325-2022.pdf.
# PyMuPDF uses top-left origin; insert_text() takes a baseline point.
# insert_text() x = left edge of text baseline.

FORM_1325_2022_COORDS = {
    # ── Header ──────────────────────────────────────────────────────────────
    # "313985129": origin=(305.22,157.75), bbox right=350.26
    "file_number":    {"x_right": 350.26, "y": 157.75, "size": 9},

    # "קופרמן סרגיי": origin=(536.20,158.03), bbox right=541.44
    # RTL; rendered via insert_htmlbox with embedded Hebrew font
    "taxpayer_name":  {"x_right": 541.44, "y": 158.03, "size": 9,
                       "bbox": (358.0, 148.0, 541.44, 162.0)},

    # Yes circle (ןכ): ZapfDingbats 'o' origin=(263.47,158.03)
    # bbox=(263.47,147.82,272.62,159.82)
    # Place 'X' centred inside this bbox
    "foreign_asset_yes_bbox": (263.47, 147.82, 272.62, 159.82),

    # "25" (tax rate %): origin=(117.64,215.41)
    "tax_rate_pct":   {"x0": 117.64, "y": 215.41, "size": 9},

    # ── Transaction rows ─────────────────────────────────────────────────────
    # Row 1 baseline y derived from median of filled span origins:
    #   BRG=269.09, acq_date=268.72, orig_cost=271.09, idx=269.30,
    #   adj_cost=270.30, sale_date=269.51, consid=269.90, gain=269.51
    # Using 269.51 (mode of cleanest fields)
    "row_1_baseline_y": 269.51,
    # Row number label spacing: 1→270.20, 2→290.04, 3→309.88 → spacing=19.84
    "row_height": 19.84,

    # Column anchors
    # security: BRG bbox=(745.70,260.94,765.20,270.98), centre x≈755.45
    # others: x_right = bbox[2] of last char
    "col": {
        "security":         {"x_center": 755.45, "size": 9},
        "acquisition_date": {"x_right": 576.55,  "size": 9},
        "original_cost":    {"x_right": 490.45,  "size": 9},
        "index_factor":     {"x_right": 438.27,  "size": 7},
        "adjusted_cost":    {"x_right": 379.08,  "size": 9},
        "sale_date":        {"x_right": 312.35,  "size": 9},
        "consideration":    {"x_right": 229.67,  "size": 9},
        "nominal_value":    {"x_right": 215.28,  "size": 9},
        "real_gain":        {"x_right": 151.09,  "size": 9},
        "capital_loss":     {"x_right": 102.10,  "size": 9},
    },

    # ── Totals ───────────────────────────────────────────────────────────────
    # Total gain/loss box: drawing rect=(28.3,462.2,347.8,482.4)
    #   gain column x=[102.1–175.9], loss column x=[28.3–102.1]
    # "20,483" (total gain): origin=(123.17,474.86), bbox right=150.69
    "total_gain":  {"x_right": 150.69, "y": 474.86, "size": 9},
    "total_loss":  {"x_right": 101.0,  "y": 474.86, "size": 9},

    # Total sales box: drawing rect=(28.3,488.3,287.4,508.5)
    #   value column x=[28.3–174.1]
    # No value in calibration; baseline = box_top(488.3) + 12.7 = 501.0
    # (same offset as total_gain: 462.2+12.7=474.9 ✓) — VERIFY with --debug
    "total_sales": {"x_right": 173.0, "y": 501.0, "size": 9},  # VERIFY

    # ── Signature ────────────────────────────────────────────────────────────
    # Image: get_image_rects → Rect(183.53,514.32,298.53,552.65)
    "signature": {"x0": 183.53, "y0": 514.32, "x1": 298.53, "y1": 552.65},

    # Date: right of signature box, above signature line (y=552.84)
    # ESTIMATE — verify with --debug
    "signature_date": {"x0": 310.0, "y": 545.0, "size": 9},  # ESTIMATE
}

COORDS_BY_YEAR = {
    2020: FORM_1325_2022_COORDS,
    2021: FORM_1325_2022_COORDS,
    2022: FORM_1325_2022_COORDS,
    2023: FORM_1325_2022_COORDS,
    2024: FORM_1325_2022_COORDS,
    2025: FORM_1325_2022_COORDS,
}


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _find_hebrew_font() -> str:
    for path in HEBREW_FONT_CANDIDATES:
        if os.path.exists(path):
            return path
    raise RuntimeError(
        "No Hebrew-capable font found. Install one of:\n  " +
        "\n  ".join(HEBREW_FONT_CANDIDATES)
    )


def _insert_right(page, x_right, y, text, fontname="helv", fontsize=9,
                  color=(0, 0, 0)):
    w = pymupdf.get_text_length(text, fontname=fontname, fontsize=fontsize)
    page.insert_text(pymupdf.Point(x_right - w, y), text,
                     fontname=fontname, fontsize=fontsize, color=color)


def _insert_left(page, x0, y, text, fontname="helv", fontsize=9,
                 color=(0, 0, 0)):
    page.insert_text(pymupdf.Point(x0, y), text,
                     fontname=fontname, fontsize=fontsize, color=color)


def _insert_center(page, x_center, y, text, fontname="helv", fontsize=9,
                   color=(0, 0, 0)):
    w = pymupdf.get_text_length(text, fontname=fontname, fontsize=fontsize)
    page.insert_text(pymupdf.Point(x_center - w / 2, y), text,
                     fontname=fontname, fontsize=fontsize, color=color)


def _insert_hebrew_rtl(page, bbox_rect, text, font_path, fontsize=9,
                       color=(0, 0, 0)):
    """Insert RTL Hebrew text right-aligned inside a rect using insert_textbox."""
    from bidi.algorithm import get_display
    visual_text = get_display(text)
    font = pymupdf.Font(fontfile=font_path)
    page.insert_font(fontname="_heb", fontbuffer=font.buffer)
    page.insert_textbox(
        pymupdf.Rect(*bbox_rect),
        visual_text,
        fontname="_heb",
        fontsize=fontsize,
        align=pymupdf.TEXT_ALIGN_RIGHT,
        color=color,
    )


def _format_ils(value) -> str:
    if value is None or value == 0 or (isinstance(value, float) and math.isnan(value)):
        return ""
    return f"{int(round(float(value))):,}"


def _format_date(d) -> str:
    if d is None:
        return ""
    if isinstance(d, (datetime.date, datetime.datetime)):
        return d.strftime("%d/%m/%Y")
    s = str(d)
    # ISO format YYYY-MM-DD
    if len(s) == 10 and s[4] == "-":
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s


def _format_index_factor(v) -> str:
    if v is None:
        return ""
    return f"{float(v):.6f}"


def normalise_rate(value) -> int:
    """Accept 0.25 (float) or 25 (int). Return integer percentage. Raises on non-integer result."""
    d = Decimal(str(value))
    pct = d * 100 if d < 1 else d
    if pct != pct.to_integral_value():
        raise ValueError(f"Non-integer tax rate percentage: {value!r}")
    return int(pct)


# ─── XLSX reader ─────────────────────────────────────────────────────────────

def load_entry_view(xlsx_path: str) -> list[dict]:
    """Read Form 1325 Entry View sheet. Returns data rows (skips subtotal rows)."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    if "Form 1325 Entry View" not in wb.sheetnames:
        raise RuntimeError(f"Sheet 'Form 1325 Entry View' not found in {xlsx_path}")
    ws = wb["Form 1325 Entry View"]

    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter, None)
    if header is None:
        raise RuntimeError("Form 1325 Entry View is empty")

    col_map = {str(v).strip(): i for i, v in enumerate(header) if v is not None}

    required = [
        "1325 Form No.", "1325 Row No.", "Security (ticker)",
        "Acquisition date", "Original price ILS",
        "1 + שיעור עליית המדד",
        "Adjusted original price ILS", "Sale date",
        "ערך נקוב במכירה",
        "Net sale consideration ILS", "Real capital gain ILS",
        "Real capital loss ILS", "Tax rate",
    ]
    missing = [c for c in required if c not in col_map]
    if missing:
        raise RuntimeError(f"Missing columns in Entry View: {missing}")

    def _get(row, col):
        idx = col_map.get(col)
        return row[idx] if idx is not None else None

    data_rows = []
    subtotal_rows = []

    for row in rows_iter:
        if all(v is None for v in row):
            continue
        form_no_raw = _get(row, "1325 Form No.")
        # Subtotal rows have a string like "Form 1" in the form_no column
        if isinstance(form_no_raw, str):
            subtotal_rows.append({
                "form_no_label": form_no_raw,
                "real_gain_ils":  _get(row, "Real capital gain ILS"),
                "real_loss_ils":  _get(row, "Real capital loss ILS"),
                "net_sale_ils":   _get(row, "Net sale consideration ILS"),
            })
            continue
        if form_no_raw is None:
            continue

        tax_rate_raw = _get(row, "Tax rate")
        if tax_rate_raw is None:
            continue

        data_rows.append({
            "form_no":          int(form_no_raw),
            "row_no":           _get(row, "1325 Row No."),
            "ticker":           _get(row, "Security (ticker)"),
            "acquisition_date": _get(row, "Acquisition date"),
            "original_cost":    _get(row, "Original price ILS"),
            "fx_ratio":         _get(row, "1 + שיעור עליית המדד"),
            "adjusted_cost":    _get(row, "Adjusted original price ILS"),
            "sale_date":        _get(row, "Sale date"),
            "nominal_value":    _get(row, "ערך נקוב במכירה"),
            "net_sale_ils":     _get(row, "Net sale consideration ILS"),
            "real_gain_ils":    _get(row, "Real capital gain ILS"),
            "real_loss_ils":    _get(row, "Real capital loss ILS"),
            "tax_rate":         tax_rate_raw,
        })

    wb.close()

    # Validate: monetary filing values are whole integers
    for r in data_rows:
        for field in ("original_cost", "adjusted_cost", "net_sale_ils",
                      "real_gain_ils", "real_loss_ils"):
            v = r[field]
            if v is not None and v != 0:
                if float(v) != int(round(float(v))):
                    raise RuntimeError(
                        f"Non-integer filing value in {field}={v!r} "
                        f"(form {r['form_no']}, row {r['row_no']})"
                    )

    return data_rows, subtotal_rows


def load_taxpayer_identity(
    taxpayer_name: str | None = None,
    file_number: str | None = None,
    config_path: str = "taxpayer.json",
) -> tuple[str, str]:
    """
    Priority: explicit args > config file > error.
    Never infers from XLSX or source files.
    """
    if taxpayer_name and file_number:
        return taxpayer_name, file_number

    if os.path.exists(config_path):
        with open(config_path, encoding="utf-8") as f:
            cfg = json.load(f)
        name = taxpayer_name or cfg.get("taxpayer_name", "").strip()
        num  = file_number  or cfg.get("file_number", "").strip()
        if name and num:
            return name, num

    raise RuntimeError(
        "Taxpayer identity not found. Provide --taxpayer-name / --file-number "
        f"or create {config_path} with keys 'taxpayer_name' and 'file_number'."
    )


def group_rows(data_rows: list[dict]) -> dict[tuple[int, int], list[dict]]:
    """Group by (form_no, rate_int_pct). Raises if any group exceeds 10 rows."""
    groups: dict[tuple[int, int], list[dict]] = {}
    for row in data_rows:
        rate_int = normalise_rate(row["tax_rate"])
        key = (row["form_no"], rate_int)
        groups.setdefault(key, []).append(row)
    for key, rows in groups.items():
        if len(rows) > 10:
            raise RuntimeError(
                f"Group {key} has {len(rows)} rows (max 10 per form page)"
            )
    return groups


def compute_totals(rows: list[dict]) -> tuple[int, int, int]:
    """Sum already-rounded filing values from Entry View rows."""
    gain  = sum(int(round(float(r["real_gain_ils"] or 0))) for r in rows)
    loss  = sum(int(round(float(r["real_loss_ils"] or 0))) for r in rows)
    sales = sum(int(round(float(r["net_sale_ils"]  or 0))) for r in rows)
    return gain, loss, sales


# ─── PDF generator ────────────────────────────────────────────────────────────

def generate_form_pdf(
    template_path: str,
    rows: list[dict],
    taxpayer_name: str,
    file_number: str,
    rate_int: int,
    coords: dict,
    signature_path: str = "signature.jpg",
    no_signature: bool = False,
    signature_date: str | None = None,
    output_path: str = "output.pdf",
    debug: bool = False,
    subtotal: dict | None = None,
):
    """Overlay filing values onto one page of the blank template."""
    doc = pymupdf.open(template_path)
    page = doc[0]

    hebrew_font_path = _find_hebrew_font()

    if signature_date is None:
        signature_date = datetime.date.today().strftime("%d/%m/%Y")

    # ── Header ─────────────────────────────────────────────────────────────
    # File number (LTR numeric)
    fn = coords["file_number"]
    _insert_right(page, fn["x_right"], fn["y"], str(file_number),
                  fontsize=fn["size"])

    # Taxpayer name (RTL Hebrew)
    tn = coords["taxpayer_name"]
    _insert_hebrew_rtl(page, tn["bbox"], taxpayer_name,
                       font_path=hebrew_font_path, fontsize=tn["size"])

    # Foreign-asset Yes mark — 'X' centred inside Yes circle bbox
    yb = coords["foreign_asset_yes_bbox"]
    cx = (yb[0] + yb[2]) / 2
    cy_baseline = (yb[1] + yb[3]) / 2 + 3  # +3 to approximate baseline from centre
    _insert_center(page, cx, cy_baseline, "X", fontsize=7)

    # Tax rate %
    tr = coords["tax_rate_pct"]
    _insert_left(page, tr["x0"], tr["y"], str(rate_int), fontsize=tr["size"])

    # ── Transaction rows ───────────────────────────────────────────────────
    row_y0    = coords["row_1_baseline_y"]
    row_h     = coords["row_height"]
    col       = coords["col"]

    for i, row in enumerate(rows):
        y = row_y0 + i * row_h

        # Security ticker — centre aligned
        if row.get("ticker"):
            _insert_center(page, col["security"]["x_center"], y,
                           str(row["ticker"]), fontsize=col["security"]["size"])

        # Acquisition date
        if row.get("acquisition_date"):
            _insert_right(page, col["acquisition_date"]["x_right"], y,
                          _format_date(row["acquisition_date"]),
                          fontsize=col["acquisition_date"]["size"])

        # Original cost
        if row.get("original_cost") is not None:
            v = _format_ils(row["original_cost"])
            if v:
                _insert_right(page, col["original_cost"]["x_right"], y, v,
                              fontsize=col["original_cost"]["size"])

        # Index factor
        if row.get("fx_ratio") is not None:
            _insert_right(page, col["index_factor"]["x_right"], y,
                          _format_index_factor(row["fx_ratio"]),
                          fontsize=col["index_factor"]["size"])

        # Adjusted cost
        if row.get("adjusted_cost") is not None:
            v = _format_ils(row["adjusted_cost"])
            if v:
                _insert_right(page, col["adjusted_cost"]["x_right"], y, v,
                              fontsize=col["adjusted_cost"]["size"])

        # Sale date
        if row.get("sale_date"):
            _insert_right(page, col["sale_date"]["x_right"], y,
                          _format_date(row["sale_date"]),
                          fontsize=col["sale_date"]["size"])

        # Consideration (net sale)
        if row.get("net_sale_ils") is not None:
            v = _format_ils(row["net_sale_ils"])
            if v:
                _insert_right(page, col["consideration"]["x_right"], y, v,
                              fontsize=col["consideration"]["size"])

        # Nominal value — blank for ordinary foreign shares
        if row.get("nominal_value") is not None:
            v = _format_ils(row["nominal_value"])
            if v:
                _insert_right(page, col["nominal_value"]["x_right"], y, v,
                              fontsize=col["nominal_value"]["size"])

        # Real gain
        if row.get("real_gain_ils"):
            v = _format_ils(row["real_gain_ils"])
            if v:
                _insert_right(page, col["real_gain"]["x_right"], y, v,
                              fontsize=col["real_gain"]["size"])

        # Capital loss
        if row.get("real_loss_ils"):
            v = _format_ils(row["real_loss_ils"])
            if v:
                _insert_right(page, col["capital_loss"]["x_right"], y, v,
                              fontsize=col["capital_loss"]["size"])

    # ── Totals ─────────────────────────────────────────────────────────────
    if subtotal:
        total_gain  = int(round(float(subtotal["real_gain_ils"] or 0)))
        total_loss  = int(round(float(subtotal["real_loss_ils"] or 0)))
        total_sales = int(round(float(subtotal["net_sale_ils"]  or 0)))
    else:
        total_gain, total_loss, total_sales = compute_totals(rows)

    tg = coords["total_gain"]
    tl = coords["total_loss"]
    ts = coords["total_sales"]

    if total_gain:
        _insert_right(page, tg["x_right"], tg["y"],
                      _format_ils(total_gain), fontsize=tg["size"])
    if total_loss:
        _insert_right(page, tl["x_right"], tl["y"],
                      _format_ils(total_loss), fontsize=tl["size"])
    if total_sales:
        _insert_right(page, ts["x_right"], ts["y"],
                      _format_ils(total_sales), fontsize=ts["size"])

    # ── Signature ──────────────────────────────────────────────────────────
    if not no_signature and os.path.exists(signature_path):
        sg = coords["signature"]
        page.insert_image(
            pymupdf.Rect(sg["x0"], sg["y0"], sg["x1"], sg["y1"]),
            filename=signature_path,
            keep_proportion=True,
        )

    # Signature date
    sd = coords["signature_date"]
    _insert_left(page, sd["x0"], sd["y"], signature_date, fontsize=sd["size"])

    # ── Debug overlay ──────────────────────────────────────────────────────
    if debug:
        _draw_debug_overlay(page, coords, len(rows))

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()


def _draw_debug_overlay(page, coords, n_rows):
    """Draw labelled red bounding boxes for every field anchor."""
    RED = (1, 0, 0)

    def _box(x0, y0, x1, y1, label):
        r = pymupdf.Rect(x0, y0, x1, y1)
        page.draw_rect(r, color=RED, width=0.5)
        page.insert_text(pymupdf.Point(x0, y0 - 1), label,
                         fontsize=5, color=RED)

    fn = coords["file_number"]
    _box(fn["x_right"] - 50, fn["y"] - 8, fn["x_right"], fn["y"] + 2,
         "file_number")

    tn = coords["taxpayer_name"]
    _box(*tn["bbox"], "taxpayer_name")

    yb = coords["foreign_asset_yes_bbox"]
    _box(*yb, "yes_circle")

    tr = coords["tax_rate_pct"]
    _box(tr["x0"], tr["y"] - 8, tr["x0"] + 20, tr["y"] + 2, "tax_rate")

    row_y0 = coords["row_1_baseline_y"]
    row_h  = coords["row_height"]
    col    = coords["col"]
    for i in range(n_rows):
        y = row_y0 + i * row_h
        for cname, cdata in col.items():
            if "x_right" in cdata:
                xr = cdata["x_right"]
                _box(xr - 40, y - 8, xr, y + 2, f"r{i+1}_{cname[:4]}")
            elif "x_center" in cdata:
                xc = cdata["x_center"]
                _box(xc - 15, y - 8, xc + 15, y + 2, f"r{i+1}_{cname[:4]}")

    tg = coords["total_gain"]
    _box(tg["x_right"] - 50, tg["y"] - 8, tg["x_right"], tg["y"] + 2,
         "total_gain")
    tl = coords["total_loss"]
    _box(tl["x_right"] - 50, tl["y"] - 8, tl["x_right"], tl["y"] + 2,
         "total_loss")
    ts = coords["total_sales"]
    _box(ts["x_right"] - 50, ts["y"] - 8, ts["x_right"], ts["y"] + 2,
         "total_sales(VERIFY)")

    sg = coords["signature"]
    _box(sg["x0"], sg["y0"], sg["x1"], sg["y1"], "signature")

    sd = coords["signature_date"]
    _box(sd["x0"], sd["y"] - 8, sd["x0"] + 60, sd["y"] + 2,
         "sig_date(ESTIMATE)")


# ─── Public API ───────────────────────────────────────────────────────────────

def generate_pdfs(
    year: int,
    xlsx_path: str,
    output_dir: str = ".",
    taxpayer_name: str | None = None,
    file_number: str | None = None,
    config_path: str = "taxpayer.json",
    template_path: str | None = None,
    signature_path: str = "signature.jpg",
    no_signature: bool = False,
    signature_date: str | None = None,
    debug: bool = False,
) -> list[str]:
    """
    Generate Form 1325 PDF(s) for the given year and XLSX workbook.
    Returns list of generated PDF file paths.
    Raises on any error (caller decides whether to treat as fatal).
    """
    year = int(year)
    # Resolve template
    if template_path is None:
        if year not in DEFAULT_TEMPLATES:
            raise RuntimeError(
                f"No calibrated template for year {year}. "
                f"Available: {list(DEFAULT_TEMPLATES.keys())}. "
                f"Use --template to supply one."
            )
        template_path = DEFAULT_TEMPLATES[year]
    if not os.path.exists(template_path):
        raise RuntimeError(f"Template PDF not found: {template_path}")

    # Resolve coordinate map
    if year not in COORDS_BY_YEAR:
        raise RuntimeError(
            f"No coordinate map for year {year}. "
            f"Available: {list(COORDS_BY_YEAR.keys())}."
        )
    coords = COORDS_BY_YEAR[year]

    # Resolve XLSX
    if not os.path.exists(xlsx_path):
        raise RuntimeError(f"XLSX not found: {xlsx_path}")

    # Load data
    data_rows, subtotal_rows = load_entry_view(xlsx_path)
    if not data_rows:
        raise RuntimeError("No data rows found in Form 1325 Entry View")

    # Taxpayer identity
    name, num = load_taxpayer_identity(taxpayer_name, file_number, config_path)

    # Group by (form_no, rate_int_pct)
    groups = group_rows(data_rows)

    # Build subtotal lookup: form_no_int → subtotal dict
    subtotal_lookup: dict[int, dict] = {}
    for s in subtotal_rows:
        label = s["form_no_label"]  # e.g. "Form 1"
        try:
            form_int = int(label.split()[-1])
            subtotal_lookup[form_int] = s
        except (ValueError, IndexError):
            pass

    generated: list[str] = []

    for (form_no, rate_int), rows in sorted(groups.items()):
        rate_str = f"{rate_int}pct"
        if len(groups) == 1 or max(k[0] for k in groups) == 1:
            out_name = f"form1325_{year}_{rate_str}.pdf"
        else:
            out_name = f"form1325_{year}_{rate_str}_part{form_no}.pdf"
        if debug:
            out_name = out_name.replace(".pdf", "_debug.pdf")

        out_path = os.path.join(output_dir, out_name)

        subtotal = subtotal_lookup.get(form_no)

        generate_form_pdf(
            template_path=template_path,
            rows=rows,
            taxpayer_name=name,
            file_number=num,
            rate_int=rate_int,
            coords=coords,
            signature_path=signature_path,
            no_signature=no_signature,
            signature_date=signature_date,
            output_path=out_path,
            debug=debug,
            subtotal=subtotal,
        )

        # Post-generation check: page count matches template
        template_pages = pymupdf.open(template_path).page_count
        output_pages   = pymupdf.open(out_path).page_count
        if output_pages != template_pages:
            raise RuntimeError(
                f"Generated PDF page count ({output_pages}) != "
                f"template ({template_pages}): {out_path}"
            )

        generated.append(out_path)

    return generated


# ─── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Generate official Form 1325 PDF from the support workbook."
    )
    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--xlsx", type=str, default=None)
    parser.add_argument("--template", type=str, default=None)
    parser.add_argument("--signature", type=str, default="signature.jpg")
    parser.add_argument("--output-dir", type=str, default=".")
    parser.add_argument("--no-signature", action="store_true")
    parser.add_argument("--signature-date", type=str, default=None,
                        metavar="DD/MM/YYYY")
    parser.add_argument("--taxpayer-name", type=str, default=None)
    parser.add_argument("--file-number", type=str, default=None)
    parser.add_argument("--config", type=str, default="taxpayer.json")
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()

    xlsx_path = args.xlsx or f"form1325_support_{args.year}.xlsx"

    paths = generate_pdfs(
        year=args.year,
        xlsx_path=xlsx_path,
        output_dir=args.output_dir,
        taxpayer_name=args.taxpayer_name,
        file_number=args.file_number,
        config_path=args.config,
        template_path=args.template,
        signature_path=args.signature,
        no_signature=args.no_signature,
        signature_date=args.signature_date,
        debug=args.debug,
    )
    for p in paths:
        print(f"Generated: {p}")


if __name__ == "__main__":
    main()
