"""
Integration tests for generate_1325_pdf.py — 2022 BRG regression case.

Requires: form1325_support_2022.xlsx in the repo root.
Run from repo root: pytest tests/integration/test_2022_pdf.py -v
"""

import os
import tempfile
import pytest
import pymupdf

# Repo root is three levels up from this file
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
XLSX_2022 = os.path.join(REPO_ROOT, "form1325_support_2022.xlsx")
TEMPLATE_2022 = os.path.join(
    REPO_ROOT, "Service_Pages_Income_tax_annual-report-2022_1325-2022.pdf"
)


def _skip_if_missing():
    if not os.path.exists(XLSX_2022):
        pytest.skip(f"Missing fixture: {XLSX_2022}")
    if not os.path.exists(TEMPLATE_2022):
        pytest.skip(f"Missing template: {TEMPLATE_2022}")


@pytest.fixture(scope="module")
def generated_pdf(tmp_path_factory):
    _skip_if_missing()
    from generate_1325_pdf import generate_pdfs

    out_dir = str(tmp_path_factory.mktemp("pdf_out"))
    paths = generate_pdfs(
        year=2022,
        xlsx_path=XLSX_2022,
        output_dir=out_dir,
        taxpayer_name="קופרמן סרגיי",
        file_number="313985129",
        no_signature=True,
    )
    assert len(paths) == 1, f"Expected 1 PDF, got {len(paths)}: {paths}"
    return paths[0]


class TestPdfGeneration:
    def test_output_file_exists(self, generated_pdf):
        assert os.path.exists(generated_pdf)

    def test_page_count_matches_template(self, generated_pdf):
        template_pages = pymupdf.open(TEMPLATE_2022).page_count
        output_pages = pymupdf.open(generated_pdf).page_count
        assert output_pages == template_pages

    def test_output_filename(self, generated_pdf):
        assert "2022" in generated_pdf
        assert "25pct" in generated_pdf


class TestFieldValues:
    """Verify that all BRG calibration values appear in the generated PDF."""

    @pytest.fixture(autouse=True)
    def _load_spans(self, generated_pdf):
        doc = pymupdf.open(generated_pdf)
        page = doc[0]
        self.spans = []
        for b in page.get_text("dict")["blocks"]:
            if b.get("type") != 0:
                continue
            for line in b.get("lines", []):
                for span in line.get("spans", []):
                    self.spans.append(span)

    def _find(self, text):
        """Return list of spans whose stripped text matches."""
        return [s for s in self.spans if s["text"].strip() == text]

    def test_security_brg(self):
        assert self._find("BRG"), "BRG not found in generated PDF"

    def test_acquisition_date(self):
        assert self._find("09/04/2020"), "Acquisition date not found"

    def test_original_cost(self):
        assert self._find("5,161"), "Original cost not found"

    def test_index_factor(self):
        assert self._find("0.978080"), "Index factor not found"

    def test_adjusted_cost(self):
        assert self._find("5,048"), "Adjusted cost not found"

    def test_sale_date(self):
        assert self._find("10/10/2022"), "Sale date not found"

    def test_consideration(self):
        assert self._find("25,644"), "Consideration not found"

    def test_real_gain(self):
        matches = [s for s in self._find("20,483")]
        assert len(matches) >= 1, "Real gain '20,483' not found"

    def test_total_gain(self):
        matches = [s for s in self._find("20,483")]
        assert len(matches) >= 2, (
            "Expected '20,483' at least twice (row + total), "
            f"found {len(matches)}"
        )

    def test_file_number(self):
        assert self._find("313985129"), "File number not found"

    def test_tax_rate(self):
        assert self._find("25"), "Tax rate (25) not found"


class TestFieldCoordinates:
    """
    Verify that key fields land at approximately the correct coordinates.
    Calibration values from 1325-2022.pdf span["origin"].
    Tolerance: ±2 points.
    """

    TOL = 2.0

    @pytest.fixture(autouse=True)
    def _load_spans(self, generated_pdf):
        doc = pymupdf.open(generated_pdf)
        page = doc[0]
        self.spans = []
        for b in page.get_text("dict")["blocks"]:
            if b.get("type") != 0:
                continue
            for line in b.get("lines", []):
                for span in line.get("spans", []):
                    self.spans.append(span)

    def _find_text(self, text):
        return [s for s in self.spans if s["text"].strip() == text]

    def _assert_near(self, span, exp_x, exp_y, label):
        ox, oy = span["origin"]
        assert abs(ox - exp_x) <= self.TOL, (
            f"{label}: origin x={ox:.2f} expected ≈{exp_x} (±{self.TOL})"
        )
        assert abs(oy - exp_y) <= self.TOL, (
            f"{label}: origin y={oy:.2f} expected ≈{exp_y} (±{self.TOL})"
        )

    def test_brg_coordinate(self):
        spans = self._find_text("BRG")
        assert spans, "BRG not found"
        # Calibration: origin=(745.70, 269.09); we insert at row_1 y=269.51
        self._assert_near(spans[0], 745.70, 269.51, "BRG")

    def test_original_cost_coordinate(self):
        spans = self._find_text("5,161")
        assert spans, "5,161 not found"
        # Calibration: origin=(467.93, 271.09); inserted right-aligned, x may differ slightly
        ox, oy = spans[0]["origin"]
        assert abs(oy - 269.51) <= self.TOL, (
            f"5,161 origin y={oy:.2f} expected ≈269.51"
        )
        # x_right=490.45; text width ≈ 22.5 → origin x ≈ 467.9
        assert abs(ox - 467.93) <= self.TOL, (
            f"5,161 origin x={ox:.2f} expected ≈467.93"
        )

    def test_acquisition_date_coordinate(self):
        spans = self._find_text("09/04/2020")
        assert spans, "09/04/2020 not found"
        ox, oy = spans[0]["origin"]
        assert abs(oy - 269.51) <= self.TOL, (
            f"acq_date origin y={oy:.2f} expected ≈269.51"
        )

    def test_row1_gain_coordinate(self):
        # Find the row-1 gain (lower y value ≈ 269), not the total (y ≈ 474)
        spans = self._find_text("20,483")
        row_spans = [s for s in spans if abs(s["origin"][1] - 269.51) <= self.TOL]
        assert row_spans, "Row-1 gain '20,483' not found near y=269.51"
        self._assert_near(row_spans[0], 123.57, 269.51, "row1_gain")

    def test_total_gain_coordinate(self):
        spans = self._find_text("20,483")
        total_spans = [s for s in spans if abs(s["origin"][1] - 474.86) <= self.TOL]
        assert total_spans, "Total gain '20,483' not found near y=474.86"
        self._assert_near(total_spans[0], 123.17, 474.86, "total_gain")

    def test_file_number_coordinate(self):
        spans = self._find_text("313985129")
        assert spans, "File number not found"
        self._assert_near(spans[0], 305.22, 157.75, "file_number")
