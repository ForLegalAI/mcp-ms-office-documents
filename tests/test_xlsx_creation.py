"""Tests for Excel (xlsx) multi-sheet support.

These tests verify that the markdown to Excel conversion handles
the '## Sheet: Name' heading syntax correctly for multi-sheet workbooks.
"""

import sys
from pathlib import Path
from unittest.mock import patch, MagicMock

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from openpyxl import Workbook, load_workbook
import io

# We test the internal parsing logic, mocking the upload step.
from xlsx_tools.base_xlsx_tool import markdown_to_excel


def _create_workbook_from_markdown(markdown_content: str) -> Workbook:
    """Helper that runs markdown_to_excel but intercepts the workbook before upload.

    Patches upload_file to capture the BytesIO and returns a loaded Workbook.
    """
    captured = {}

    def fake_upload(file_obj, suffix, **kwargs):
        captured['data'] = file_obj.read()
        file_obj.seek(0)
        return "https://fake-url/test.xlsx"

    with patch("xlsx_tools.base_xlsx_tool.upload_file", side_effect=fake_upload):
        markdown_to_excel(markdown_content)

    wb = load_workbook(io.BytesIO(captured['data']))
    return wb


# Output directory for test files
OUTPUT_DIR = Path(__file__).parent / "output" / "xlsx"


@pytest.fixture(scope="module", autouse=True)
def setup_output_dir():
    """Create output directory if it doesn't exist."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    yield


class TestMultiSheet:
    """Tests for multi-sheet Excel workbooks via ## Sheet: Name."""

    def test_single_sheet_default_name(self):
        """Markdown without ## Sheet: heading → single sheet named 'Data Report'."""
        markdown = """# Report

| Name | Value |
|------|-------|
| A    | 1     |
"""
        wb = _create_workbook_from_markdown(markdown)
        assert len(wb.sheetnames) == 1
        assert wb.sheetnames[0] == "Data Report"

    def test_two_sheets(self):
        """Markdown with two ## Sheet: headings → two sheets."""
        markdown = """## Sheet: Revenue

| Quarter | Amount |
|---------|--------|
| Q1      | 1000   |
| Q2      | 1200   |

## Sheet: Expenses

| Quarter | Amount |
|---------|--------|
| Q1      | 800    |
| Q2      | 900    |
"""
        wb = _create_workbook_from_markdown(markdown)
        assert len(wb.sheetnames) == 2
        assert "Revenue" in wb.sheetnames
        assert "Expenses" in wb.sheetnames

    def test_sheet_names_correct(self):
        """Verify sheet names are correctly set from headings."""
        markdown = """## Sheet: Summary

| Metric | Value |
|--------|-------|
| Total  | 100   |

## Sheet: Detail Data

| Item | Count |
|------|-------|
| A    | 50    |
| B    | 50    |
"""
        wb = _create_workbook_from_markdown(markdown)
        assert wb.sheetnames[0] == "Summary"
        assert wb.sheetnames[1] == "Detail Data"

    def test_data_on_correct_sheets(self):
        """Verify tables land on the correct sheets."""
        markdown = """## Sheet: Sheet1

| Col1 | Col2 |
|------|------|
| X    | Y    |

## Sheet: Sheet2

| ColA | ColB |
|------|------|
| M    | N    |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws1 = wb["Sheet1"]
        ws2 = wb["Sheet2"]

        # Check data exists on Sheet1 (header row + data row)
        # The exact row depends on spacing; just verify the values exist somewhere
        sheet1_values = []
        for row in ws1.iter_rows(values_only=True):
            sheet1_values.extend([v for v in row if v is not None])
        assert "X" in sheet1_values or "x" in str(sheet1_values).lower()

        sheet2_values = []
        for row in ws2.iter_rows(values_only=True):
            sheet2_values.extend([v for v in row if v is not None])
        assert "M" in sheet2_values or "m" in str(sheet2_values).lower()

    def test_three_sheets(self):
        """Test creating three sheets."""
        markdown = """## Sheet: Alpha

| A |
|---|
| 1 |

## Sheet: Beta

| B |
|---|
| 2 |

## Sheet: Gamma

| C |
|---|
| 3 |
"""
        wb = _create_workbook_from_markdown(markdown)
        assert len(wb.sheetnames) == 3
        assert wb.sheetnames == ["Alpha", "Beta", "Gamma"]

    def test_backwards_compatible_no_sheet_heading(self):
        """Without any ## Sheet: headings, everything goes to 'Data Report'."""
        markdown = """# My Report

| Name | Age |
|------|-----|
| Alice | 30 |
| Bob   | 25 |
"""
        wb = _create_workbook_from_markdown(markdown)
        assert len(wb.sheetnames) == 1
        assert wb.sheetnames[0] == "Data Report"


class TestCrossSheetReferences:
    """Tests for cross-sheet cell references via SheetName!T1.B[0] syntax."""

    def test_simple_cross_sheet_reference(self):
        """A formula on Sheet2 referencing a cell on Sheet1 via SheetName!T1.B[0]."""
        markdown = """## Sheet: Revenue

| Quarter | Amount |
|---------|--------|
| Q1      | 1000   |
| Q2      | 1200   |

## Sheet: Dashboard

| Metric | Value |
|--------|-------|
| Q1 Rev | =Revenue!T1.B[0] |
| Q2 Rev | =Revenue!T1.B[1] |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws = wb["Dashboard"]
        # Revenue table starts at row 1 (first sheet, no heading before table)
        # T1 on Revenue starts at row 1, data row [0] → row 2, data row [1] → row 3
        # So the formulas should be =Revenue!B2 and =Revenue!B3
        q1_cell = ws.cell(row=1, column=2)  # header row of Dashboard T1 is row 1
        q2_cell = ws.cell(row=2, column=2)  # data row 0
        # Data rows of Dashboard are at row 2, row 3
        q1_val = ws.cell(row=2, column=2).value
        q2_val = ws.cell(row=3, column=2).value
        assert q1_val == "=Revenue!B2", f"Expected =Revenue!B2, got {q1_val}"
        assert q2_val == "=Revenue!B3", f"Expected =Revenue!B3, got {q2_val}"

    def test_cross_sheet_reference_with_spaces_in_name(self):
        """Sheet names with spaces get quoted in the Excel formula."""
        markdown = """## Sheet: Sales Data

| Product | Revenue |
|---------|---------|
| Widget  | 5000    |

## Sheet: Summary

| Metric | Value |
|--------|-------|
| Total  | =Sales Data!T1.B[0] |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws = wb["Summary"]
        cell_val = ws.cell(row=2, column=2).value
        assert cell_val == "='Sales Data'!B2", f"Expected ='Sales Data'!B2, got {cell_val}"

    def test_forward_reference_sheet1_refs_sheet2(self):
        """Sheet1 formula references Sheet2 that hasn't been parsed yet (forward reference)."""
        markdown = """## Sheet: Dashboard

| Metric | Value |
|--------|-------|
| Total  | =Details!T1.B[0] |

## Sheet: Details

| Item | Amount |
|------|--------|
| Rent | 3000   |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws = wb["Dashboard"]
        cell_val = ws.cell(row=2, column=2).value
        assert cell_val == "=Details!B2", f"Expected =Details!B2, got {cell_val}"

    def test_cross_sheet_range_reference(self):
        """Cross-sheet range reference SheetName!T1.B[0]:T1.B[2]."""
        markdown = """## Sheet: Data

| Name | Score |
|------|-------|
| A    | 10    |
| B    | 20    |
| C    | 30    |

## Sheet: Summary

| Metric | Value |
|--------|-------|
| Total  | =SUM(Data!T1.B[0]:T1.B[2]) |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws = wb["Summary"]
        cell_val = ws.cell(row=2, column=2).value
        assert cell_val == "=SUM(Data!B2:B4)", f"Expected =SUM(Data!B2:B4), got {cell_val}"

    def test_mixed_local_and_cross_sheet(self):
        """Formula mixing local and cross-sheet references."""
        markdown = """## Sheet: Revenue

| Quarter | Amount |
|---------|--------|
| Q1      | 1000   |

## Sheet: Costs

| Quarter | Amount |
|---------|--------|
| Q1      | 400    |

## Sheet: Profit

| Quarter | Revenue | Cost | Profit |
|---------|---------|------|--------|
| Q1      | =Revenue!T1.B[0] | =Costs!T1.B[0] | =B[0]-C[0] |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws = wb["Profit"]
        rev_val = ws.cell(row=2, column=2).value
        cost_val = ws.cell(row=2, column=3).value
        profit_val = ws.cell(row=2, column=4).value
        assert rev_val == "=Revenue!B2", f"Expected =Revenue!B2, got {rev_val}"
        assert cost_val == "=Costs!B2", f"Expected =Costs!B2, got {cost_val}"
        assert profit_val == "=B2-C2", f"Expected =B2-C2, got {profit_val}"

    def test_cross_sheet_with_header_before_table(self):
        """Cross-sheet reference where the source sheet has a header before the table."""
        markdown = """## Sheet: Source

# Monthly Data

| Month | Value |
|-------|-------|
| Jan   | 100   |
| Feb   | 200   |

## Sheet: Target

| Metric | Value |
|--------|-------|
| Jan    | =Source!T1.B[0] |
| Feb    | =Source!T1.B[1] |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws = wb["Target"]
        # Source: header at row 1, +2 spacing → table starts at row 3
        # T1 data[0] = row 4, data[1] = row 5
        jan_val = ws.cell(row=2, column=2).value
        feb_val = ws.cell(row=3, column=2).value
        assert jan_val == "=Source!B4", f"Expected =Source!B4, got {jan_val}"
        assert feb_val == "=Source!B5", f"Expected =Source!B5, got {feb_val}"


class TestAdjustFormulaReferencesUnit:
    """Unit tests for adjust_formula_references with cross-sheet support."""

    def test_cross_sheet_single_cell(self):
        from xlsx_tools.helpers import adjust_formula_references
        all_positions = {"Revenue": {"T1": 1}}
        result = adjust_formula_references(
            "=Revenue!T1.B[0]", 10, {}, all_positions
        )
        assert result == "=Revenue!B2"

    def test_cross_sheet_quoted_name(self):
        from xlsx_tools.helpers import adjust_formula_references
        all_positions = {"My Sheet": {"T1": 1}}
        result = adjust_formula_references(
            "=My Sheet!T1.A[2]", 10, {}, all_positions
        )
        assert result == "='My Sheet'!A4"

    def test_cross_sheet_range(self):
        from xlsx_tools.helpers import adjust_formula_references
        all_positions = {"Data": {"T1": 1}}
        result = adjust_formula_references(
            "=SUM(Data!T1.B[0]:T1.B[4])", 10, {}, all_positions
        )
        assert result == "=SUM(Data!B2:B6)"

    def test_cross_sheet_function_pattern(self):
        from xlsx_tools.helpers import adjust_formula_references
        all_positions = {"Sales": {"T1": 3}}
        result = adjust_formula_references(
            "=Sales!T1.SUM(B[0]:D[0])", 10, {}, all_positions
        )
        # T1 starts at row 3, data[0] → row 4
        assert result == "=SUM(Sales!B4:Sales!D4)"

    def test_local_reference_still_works(self):
        from xlsx_tools.helpers import adjust_formula_references
        result = adjust_formula_references(
            "=T1.B[0]", 5, {"T1": 1}, {}
        )
        assert result == "=B2"

    def test_mixed_local_and_cross_sheet(self):
        from xlsx_tools.helpers import adjust_formula_references
        all_positions = {"Revenue": {"T1": 1}}
        result = adjust_formula_references(
            "=Revenue!T1.B[0]-B[0]", 5, {"T1": 3}, all_positions
        )
        # Revenue!T1.B[0] → Revenue!B2, B[0] → B4 (current table starts at 3, data[0] = row 4)
        assert result == "=Revenue!B2-B4"


class TestNumberFormats:
    """Tests for cell number_format (percent, thousands separator)."""

    def test_percent_cells_get_percent_format(self):
        """Cells with '50%' should be stored as 0.5 with number_format '0%'."""
        markdown = """| Metric | Rate |
|--------|------|
| Growth | 50%  |
| Margin | 8%   |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws = wb.active
        # Data rows start at row 2 (row 1 = header)
        growth_cell = ws.cell(row=2, column=2)
        margin_cell = ws.cell(row=3, column=2)
        assert growth_cell.value == pytest.approx(0.5)
        assert growth_cell.number_format == '0%'
        assert margin_cell.value == pytest.approx(0.08)
        assert margin_cell.number_format == '0%'

    def test_non_percent_number_no_percent_format(self):
        """A plain decimal like 0.5 (without '%') should NOT get '0%' format."""
        markdown = """| Item  | Value |
|-------|-------|
| Alpha | 0.5   |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws = wb.active
        cell = ws.cell(row=2, column=2)
        assert cell.value == pytest.approx(0.5)
        assert cell.number_format != '0%'

    def test_thousands_separator_format(self):
        """Values >= 1000 should get '#,##0' number format."""
        markdown = """| Item | Amount |
|------|--------|
| Rent | 5000   |
"""
        wb = _create_workbook_from_markdown(markdown)
        ws = wb.active
        cell = ws.cell(row=2, column=2)
        assert cell.value == 5000
        assert cell.number_format == '#,##0'


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])

