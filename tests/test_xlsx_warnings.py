"""The Excel tool's warnings channel (#114).

Excel used to lose content and mis-resolve formulas in silence: a line that is
not a table row is dropped, a formula naming a table or a sheet that does not
exist still resolves — to the wrong cell, or to #REF! once Excel opens the
file — and an over-length formula is stored as text. All of it produced a
success response and a log line the caller never sees.

Each test here names one of those and asserts it comes back as data: a stable
code, a severity, and the sheet and cell (or source line) to look at.
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from xlsx_tools import _markdown_to_excel_buffer
from xlsx_tools import warnings as W
from warning_channel import SEVERITIES, SEVERITY_ERROR, SEVERITY_WARNING

TABLE = "| Item | Value |\n|------|-------|\n| A    | 1     |\n"


def build(markdown, **kwargs):
    """Build *markdown* and return the warnings it produced."""
    buffer, warnings = _markdown_to_excel_buffer(markdown, **kwargs)
    buffer.close()
    return warnings


def codes(warnings):
    return [warning.code for warning in warnings]


def only(warnings, code):
    matched = [w for w in warnings if w.code == code]
    assert matched, f"expected a {code} warning, got {codes(warnings)}"
    return matched[0]


class TestTheCodeTable:

    def test_every_code_has_a_severity(self):
        codes_defined = {
            value for name, value in vars(W).items()
            if name.isupper() and isinstance(value, str) and not name.startswith("SEVERITY")
        }
        assert codes_defined == set(W.WARNING_SEVERITY)

    def test_every_severity_is_one_of_the_three(self):
        assert set(W.WARNING_SEVERITY.values()) <= set(SEVERITIES)


class TestContentThatDidNotArrive:

    def test_a_line_with_nowhere_to_go_is_reported(self):
        warnings = build(f"# Report\n\nThis sentence is not a table.\n\n{TABLE}")

        warning = only(warnings, W.LINE_DROPPED)
        assert warning.severity == SEVERITY_ERROR
        assert warning.location["line"] == 3
        assert warning.location["sheet"] == "Data Report"
        assert "This sentence is not a table." in warning.message

    def test_a_dropped_line_names_the_sheet_it_was_on(self):
        warnings = build(f"## Sheet: Revenue\n{TABLE}\nstray prose\n")

        assert only(warnings, W.LINE_DROPPED).location["sheet"] == "Revenue"

    def test_a_long_dropped_line_is_quoted_in_part(self):
        """Quote enough to find the line, not the whole of it."""
        warnings = build(f"{TABLE}\n" + "x" * 200 + "\n")

        message = only(warnings, W.LINE_DROPPED).message
        assert "…" in message
        assert "x" * 60 in message
        assert "x" * 61 not in message

    def test_a_pipe_line_that_is_not_a_table_is_reported(self):
        """It looks like a table, so parse_table eats it before the
        dropped-line branch can see it — and it is gone from the workbook."""
        warnings = build(f"{TABLE}\n| Item | Qty |\n")

        warning = only(warnings, W.TABLE_INCOMPLETE)
        assert warning.severity == SEVERITY_ERROR
        assert warning.location["line"] == 5
        assert warning.location["sheet"] == "Data Report"
        assert "| Item | Qty |" in warning.message
        assert "separator row" in warning.message

    def test_separator_rows_with_no_header_are_reported(self):
        """The other way parse_table comes back with nothing."""
        warnings = build(f"{TABLE}\n|---|---|\n|---|---|\n")

        assert codes(warnings) == [W.TABLE_INCOMPLETE]

    def test_a_well_formed_table_reports_nothing(self):
        assert build(TABLE) == []

    def test_a_cell_that_cannot_be_written_is_reported(self, monkeypatch):
        import xlsx_tools.helpers as helpers
        monkeypatch.setattr(helpers, "apply_cell_formatting",
                            lambda *a, **k: (_ for _ in ()).throw(ValueError("bad cell")))

        warning = only(build(TABLE), W.CELL_FAILED)
        assert warning.severity == SEVERITY_ERROR
        assert warning.location["sheet"] == "Data Report"
        assert warning.location["cell"] == "A2"     # the first data cell
        assert "bad cell" in warning.message


class TestFormulasThatWillNotCompute:

    def test_a_reference_to_a_table_that_does_not_exist_is_reported(self):
        warnings = build("| Item | Value |\n|------|-------|\n| A | =T9.B[0] |\n")

        warning = only(warnings, W.TABLE_REFERENCE_MISSING)
        assert warning.severity == SEVERITY_ERROR
        assert warning.location["cell"] == "B2"
        assert "T9" in warning.message

    def test_a_reference_to_a_sheet_that_does_not_exist_is_reported(self):
        warnings = build("| Item | Value |\n|------|-------|\n| A | =Nope!T1.B[0] |\n")

        warning = only(warnings, W.SHEET_REFERENCE_MISSING)
        assert warning.severity == SEVERITY_ERROR
        assert "Nope" in warning.message
        assert "#REF!" in warning.message

    def test_an_over_length_formula_stored_as_text_is_reported(self):
        long_formula = "=" + "+".join(["1"] * 5000)
        warnings = build(f"| Item | Value |\n|------|-------|\n| A | {long_formula} |\n")

        warning = only(warnings, W.FORMULA_TOO_LONG)
        assert warning.severity == SEVERITY_ERROR
        assert warning.location["cell"] == "B2"
        assert "stored as text" in warning.message

    def test_a_formula_whose_references_will_not_resolve_is_reported(self, monkeypatch):
        import xlsx_tools.helpers as helpers
        monkeypatch.setattr(helpers, "_resolve_row",
                            lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom")))

        warnings = build("| Item | Value |\n|------|-------|\n| A | =T1.B[0] |\n")

        warning = only(warnings, W.FORMULA_UNRESOLVED)
        assert warning.location["cell"] == "B2"
        assert "boom" in warning.message

    def test_a_circular_reference_is_reported(self):
        warnings = build("| Item | Value |\n|------|-------|\n| A | =B2+1 |\n")

        warning = only(warnings, W.CIRCULAR_REFERENCE)
        assert warning.severity == SEVERITY_ERROR
        assert warning.location["cell"] == "B2"
        assert "resolve the cycle to 0" in warning.message


class TestNamesAndFormatting:

    def test_a_table_without_a_separator_row_is_reported(self):
        """parse_table() does not require the separator, it only skips rows
        that look like one — so the table parses and its first row silently
        becomes the header."""
        warnings = build("| Item | Qty |\n| A    | 1   |\n")

        warning = only(warnings, W.TABLE_SEPARATOR_MISSING)
        assert warning.severity == SEVERITY_WARNING
        assert warning.location["line"] == 1
        assert warning.location["sheet"] == "Data Report"
        assert "first row was used as the header" in warning.message

    def test_the_warning_describes_what_actually_happened(self):
        """The first row really is the header, which is why this is worth
        saying: T1.B[0] counts from the row after it."""
        from openpyxl import load_workbook

        buffer, _ = _markdown_to_excel_buffer(
            "| Item | Qty |\n| A    | 1   |\n| B | =T1.B[0] |\n")
        sheet = load_workbook(buffer).active
        buffer.close()

        assert sheet["A1"].value == "Item"        # header, not data
        assert sheet["A2"].value == "A"           # first data row
        assert sheet["B3"].value == "=B2"         # T1.B[0] → the row after A1

    def test_a_table_with_a_separator_reports_nothing(self):
        assert build("| Item | Qty |\n|------|-----|\n| A | 1 |\n") == []


    def test_a_colliding_sheet_name_is_reported(self):
        warnings = build("## Sheet: Report\n| A |\n|---|\n| 1 |\n\n"
                         "## Sheet: Report\n| B |\n|---|\n| 2 |\n")

        warning = only(warnings, W.SHEET_NAME_COLLISION)
        assert warning.severity == SEVERITY_WARNING
        assert "Report" in warning.message

    def test_an_unusable_sheet_name_is_reported(self, monkeypatch):
        """Excel rejects []*?:/\\ in a title; the parser normally strips them."""
        import xlsx_tools.parser as parser
        monkeypatch.setattr(parser, "_sanitize_sheet_name", lambda name: name)

        warnings = build(f"## Sheet: Bad[Name]\n{TABLE}")

        warning = only(warnings, W.SHEET_NAME_INVALID)
        assert "Bad[Name]" in warning.message
        assert "cross-sheet references" in warning.message

    def test_a_renamed_table_header_is_reported(self):
        warnings = build("| Q1 | Q1 |\n|----|----|\n| 1  | 2  |\n", auto_filter=True)

        warning = only(warnings, W.HEADER_RENAMED)
        assert warning.severity == SEVERITY_WARNING
        assert warning.location["cell"] == "B1"
        assert "Q1_2" in warning.message

    def test_a_blank_table_header_is_reported(self):
        warnings = build("| Q1 |  |\n|----|----|\n| 1  | 2 |\n", auto_filter=True)

        assert "is empty" in only(warnings, W.HEADER_RENAMED).message

    def test_a_malformed_styles_entry_is_reported(self):
        warnings = build(f"<!-- styles: bogus -->\n{TABLE}")

        warning = only(warnings, W.STYLE_ENTRY_INVALID)
        assert warning.severity == SEVERITY_WARNING
        assert "bogus" in warning.message

    def test_an_unresolvable_style_target_is_reported(self):
        warnings = build(f"<!-- styles: 2B=bold -->\n{TABLE}")

        assert "2B" in only(warnings, W.STYLE_ENTRY_INVALID).message

    def test_an_oversized_style_range_is_reported_once(self):
        """It is skipped for one specific reason; say that reason only."""
        warnings = build(f"<!-- styles: A1:Z100000=bold -->\n{TABLE}")

        assert codes(warnings) == [W.STYLE_RANGE_TOO_LARGE]
        assert "10000" in warnings[0].message

    def test_styling_that_could_not_be_applied_is_reported(self, monkeypatch):
        import xlsx_tools.styles as styles
        monkeypatch.setattr(styles, "apply_style_spec",
                            lambda *a, **k: (_ for _ in ()).throw(ValueError("nope")))

        warnings = build(f"<!-- styles: A1=bold -->\n{TABLE}")

        warning = only(warnings, W.STYLE_FAILED)
        assert warning.location["cell"] == "A1"
        assert "nope" in warning.message


class TestTheEntryPoint:

    def test_clean_markdown_produces_a_buffer_and_no_warnings(self):
        buffer, warnings = _markdown_to_excel_buffer(f"# Report\n\n{TABLE}")

        assert buffer.tell() == 0
        assert warnings == []
        buffer.close()

    def test_the_buffer_still_holds_a_workbook_when_something_went_wrong(self):
        """A warning is not an error: the file is delivered either way."""
        buffer, warnings = _markdown_to_excel_buffer(f"{TABLE}\nstray prose\n")

        assert len(buffer.read()) > 0
        assert codes(warnings) == [W.LINE_DROPPED]
        buffer.close()

    def test_warnings_are_records_not_sentences(self):
        warnings = build(f"{TABLE}\nstray prose\n")

        assert warnings[0].as_dict() == {
            "code": W.LINE_DROPPED,
            "severity": SEVERITY_ERROR,
            "message": warnings[0].message,
            "sheet": "Data Report",
            "line": 5,
        }

    def test_a_wall_of_prose_does_not_produce_a_wall_of_warnings(self):
        from warning_channel import DEFAULT_LIMIT, WARNINGS_TRUNCATED

        warnings = build(TABLE + "\n" + "\n".join(f"prose {i}" for i in range(200)))

        assert len(warnings) == DEFAULT_LIMIT + 1
        assert warnings[-1].code == WARNINGS_TRUNCATED


class TestToolBoundary:
    """The structure only matters if it survives to the caller."""

    @staticmethod
    async def call(monkeypatch, **arguments):
        import main
        from fastmcp import Client

        async def fake_upload(file_buffer, extension, file_name, user_context,
                              message, **kwargs):
            return "https://example.invalid/book.xlsx"

        monkeypatch.setattr(main, "upload_and_format_response", fake_upload)
        async with Client(main.mcp) as client:
            return await client.call_tool("create_excel_from_markdown",
                                          arguments, raise_on_error=False)

    async def test_warnings_reach_the_caller_as_objects(self, monkeypatch):
        result = await self.call(monkeypatch,
                                 markdown_content=f"{TABLE}\nstray prose\n")

        assert not result.is_error
        assert result.data["file"] == "https://example.invalid/book.xlsx"
        warning = result.data["warnings"][0]
        assert warning["code"] == W.LINE_DROPPED
        assert warning["severity"] == "error"
        assert warning["sheet"] == "Data Report"
        assert warning["line"] == 5

    async def test_a_clean_workbook_still_returns_a_bare_url(self, monkeypatch):
        result = await self.call(monkeypatch, markdown_content=TABLE)

        assert not result.is_error
        assert result.data == "https://example.invalid/book.xlsx"
