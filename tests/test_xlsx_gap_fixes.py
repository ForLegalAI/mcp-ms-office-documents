"""Round 8 regression tests for three real bugs in the xlsx formula/reference
pipeline. Each bug was reproduced against the live implementation before
fixing.

- Bug 1: cross-sheet function range syntax emitted an invalid double sheet
  prefix (=SUM(Data!B2:Data!B4)) which yields #VALUE!.
- Bug 2: the local B[0]/A[0] "current row reference" notation resolved
  table-relative, silently corrupting every row past the first data row.
- Bug 3: financial mode forced ALL 4-digit numbers to text, not just
  plausible years, breaking SUM in real Excel and dropping number formats.
"""

from __future__ import annotations

import io
import sys
from pathlib import Path
from unittest.mock import patch

import pytest
from openpyxl import load_workbook

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from xlsx_tools.base_xlsx_tool import markdown_to_excel  # noqa: E402

OUTPUT_DIR = Path(__file__).parent / "output" / "xlsx"


@pytest.fixture(scope="module", autouse=True)
def setup_output_dir():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    yield


def _intercept_upload(markdown: str, **kwargs) -> bytes:
    """Run markdown_to_excel with upload patched out; return raw xlsx bytes."""
    captured = {}

    def fake_upload(file_obj, suffix, **kw):
        captured["data"] = file_obj.read()
        return "fake://test.xlsx"

    with patch("xlsx_tools.base_xlsx_tool.upload_file", side_effect=fake_upload):
        markdown_to_excel(markdown, **kwargs)
    return captured["data"]


# ════════════════════════════════════════════════════════════════════════════
# ROUND 8 — three real bugs reproduced and fixed.
# ════════════════════════════════════════════════════════════════════════════


class TestCrossSheetFunctionRangePrefix:
    """Bug 1: =SheetName!T1.SUM(B[0]:B[2]) must emit =SUM(Sheet!B2:B4),
    not the invalid =SUM(Sheet!B2:Sheet!B4) (which yields #VALUE!)."""

    def test_cross_sheet_sum_single_prefix(self):
        from xlsx_tools.helpers import adjust_formula_references
        all_pos = {"Data": {"T1": 1}}
        result = adjust_formula_references(
            "=Data!T1.SUM(B[0]:B[2])", 2, {"T1": 1}, all_pos
        )
        assert result == "=SUM(Data!B2:B4)"

    def test_cross_sheet_average_single_prefix(self):
        from xlsx_tools.helpers import adjust_formula_references
        all_pos = {"Data": {"T1": 1}}
        result = adjust_formula_references(
            "=Data!T1.AVERAGE(B[0]:B[2])", 2, {"T1": 1}, all_pos
        )
        assert result == "=AVERAGE(Data!B2:B4)"

    def test_cross_sheet_sum_end_to_end_recalcs(self):
        """End-to-end: the corrected formula must recalc without #VALUE!."""
        markdown = (
            "## Sheet: Data\n\n"
            "| V |\n|---|\n| 10 |\n| 20 |\n| 30 |\n\n"
            "## Sheet: Summary\n\n"
            "| Total |\n|---|\n| =Data!T1.SUM(A[0]:A[2]) |\n"
        )
        data = _intercept_upload(markdown, recalc=True)
        wb = load_workbook(io.BytesIO(data), data_only=True)
        assert wb["Summary"]["A2"].value == 60

    def test_cross_sheet_range_still_correct(self):
        """The sibling cs_range pattern was already correct — guard against
        a regression when fixing the function pattern."""
        from xlsx_tools.helpers import adjust_formula_references
        all_pos = {"Data": {"T1": 1}}
        result = adjust_formula_references(
            "=Data!T1.B[0]:T1.B[2]", 2, {"T1": 1}, all_pos
        )
        assert result == "=Data!B2:B4"


class TestCurrentRowRelativeRefs:
    """Bug 2: B[0]/A[0] is documented as CURRENT-row-relative but resolved
    table-relative, silently corrupting every row past the first data row."""

    def test_unit_current_row_offset_zero(self):
        from xlsx_tools.helpers import adjust_formula_references
        tp = {"T1": 1}
        for cur in [2, 3, 4, 5]:
            result = adjust_formula_references("=B[0]", cur, table_positions=tp)
            assert result == f"=B{cur}", f"row {cur}: {result}"

    def test_unit_current_row_negative_offset(self):
        from xlsx_tools.helpers import adjust_formula_references
        tp = {"T1": 1}
        # B[-1] on row N → B(N-1); the most common running-total pattern.
        for cur in [2, 3, 4, 5]:
            result = adjust_formula_references("=B[-1]", cur, table_positions=tp)
            assert result == f"=B{cur - 1}", f"row {cur}: {result}"

    def test_unit_current_row_positive_offset(self):
        from xlsx_tools.helpers import adjust_formula_references
        tp = {"T1": 1}
        for cur in [2, 3, 4]:
            result = adjust_formula_references("=B[1]", cur, table_positions=tp)
            assert result == f"=B{cur + 1}", f"row {cur}: {result}"

    def test_unit_range_pattern_uses_current_row(self):
        from xlsx_tools.helpers import adjust_formula_references
        tp = {"T1": 1}
        # =SUM(B[0]:E[0]) at row 4 → =SUM(B4:E4)
        result = adjust_formula_references("=SUM(B[0]:E[0])", 4, table_positions=tp)
        assert result == "=SUM(B4:E4)"
        # Range spanning rows: =SUM(B[-1]:B[1]) at row 3 → =SUM(B2:B4)
        result = adjust_formula_references("=SUM(B[-1]:B[1])", 3, table_positions=tp)
        assert result == "=SUM(B2:B4)"

    def test_table_relative_T1_still_correct(self):
        """The T1.B[n] form is table-relative (offset from first data row)
        and must be unaffected by the current-row fix."""
        from xlsx_tools.helpers import adjust_formula_references
        tp = {"T1": 1}
        for n, expected_row in [(0, 2), (1, 3), (2, 4)]:
            result = adjust_formula_references(f"=T1.B[{n}]", 99, table_positions=tp)
            assert result == f"=B{expected_row}"

    def test_end_to_end_running_total(self):
        """The classic running-total pattern must produce correct cached values."""
        markdown = (
            "| Delta | RunningTotal |\n"
            "|---|---|\n"
            "| 10 | =A[0] |\n"
            "| 20 | =B[-1]+A[0] |\n"
            "| 30 | =B[-1]+A[0] |\n"
        )
        data = _intercept_upload(markdown, recalc=True)
        wb = load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        assert ws["B2"].value == 10
        assert ws["B3"].value == 30
        assert ws["B4"].value == 60

    def test_end_to_end_per_row_multiply(self):
        """=A[0]*2 on each row must hit each row's A, not always row 2."""
        markdown = (
            "| Base | TwiceBase |\n"
            "|---|---|\n"
            "| 5 | =A[0]*2 |\n"
            "| 10 | =A[0]*2 |\n"
            "| 15 | =A[0]*2 |\n"
        )
        data = _intercept_upload(markdown, recalc=True)
        wb = load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        assert ws["B2"].value == 10
        assert ws["B3"].value == 20
        assert ws["B4"].value == 30


class TestYearAsStringRestriction:
    """Bug 3: financial mode forced ALL 4-digit numbers to text, not just
    plausible years. Revenue of 1500 became text '1500'."""

    def test_year_string_in_range(self):
        from xlsx_tools.helpers import _is_year_string
        for y in ["2024", "2025", "2099", "2100", "1900", "2050"]:
            assert _is_year_string(y), f"{y} should be a year"

    def test_year_string_out_of_range(self):
        from xlsx_tools.helpers import _is_year_string
        for n in ["1500", "1899", "2101", "5000", "9999", "1234"]:
            assert not _is_year_string(n), f"{n} should NOT be a year"

    def test_non_four_digit_not_year(self):
        from xlsx_tools.helpers import _is_year_string
        for v in ["2050E", "2024A", "abc", "12345", "123"]:
            assert not _is_year_string(v)

    def test_revenue_4_digit_stays_numeric_in_financial_mode(self):
        """A 4-digit revenue value must remain a number, not become text."""
        markdown = "| Revenue |\n|---|\n| 1500 |\n| 2500 |\n"
        data = _intercept_upload(markdown, financial_modeling=True, recalc=False)
        ws = load_workbook(io.BytesIO(data)).active
        assert isinstance(ws["A2"].value, (int, float))
        assert ws["A2"].value == 1500
        assert ws["A2"].number_format == "#,##0;(#,##0);-"
        assert isinstance(ws["A3"].value, (int, float))

    def test_year_column_stays_text_in_financial_mode(self):
        """Genuine years (2024, 2025) must still be text labels."""
        markdown = "| Year |\n|---|\n| 2024 |\n| 2025 |\n"
        data = _intercept_upload(markdown, financial_modeling=True, recalc=False)
        ws = load_workbook(io.BytesIO(data)).active
        assert ws["A2"].value == "2024"
        assert isinstance(ws["A2"].value, str)
        assert ws["A3"].value == "2025"

    def test_sum_over_4_digit_revenue_correct_in_financial_mode(self):
        """End-to-end: SUM over 4-digit revenue cells must total correctly,
        even when the file is later opened in real Excel (which does NOT
        coerce text-numerics in SUM the way the recalc engine does)."""
        markdown = (
            "| Revenue |\n|---|\n| 1500 |\n| 2500 |\n| 3500 |\n| =SUM(A2:A4) |\n"
        )
        data = _intercept_upload(markdown, financial_modeling=True, recalc=True)
        wb = load_workbook(io.BytesIO(data), data_only=True)
        ws = wb.active
        # All inputs must be numeric (not text) so Excel-native SUM works.
        assert isinstance(ws["A2"].value, (int, float))
        assert isinstance(ws["A3"].value, (int, float))
        assert isinstance(ws["A4"].value, (int, float))
        assert ws["A5"].value == 7500
