"""A table row without its trailing pipe is still a row, in Word as in Excel (#144).

GitHub-flavoured Markdown makes the trailing ``|`` optional, and models drop
it. Excel already accepted such a row; Word demanded both pipes, so a table
whose last row was ``| 3 | 4`` ended one row early and printed ``| 3 | 4``
under it as prose with nothing reported. A *header* without its trailing pipe
was worse: the table started on the separator, its header read ``---|---``,
and the real header was printed as prose.

Both tools now read rows through ``markdown_table``. What Word's block
detection treats as a table *before* parsing stays strict, so prose such as
``|x| < 5`` keeps flowing into its paragraph and filling a placeholder inline.
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest  # noqa: E402
from docx import Document  # noqa: E402

from docx_tools import _markdown_to_word_buffer  # noqa: E402
from docx_tools import warnings as W  # noqa: E402
from docx_tools.block_elements import parse_table as word_parse_table  # noqa: E402
from docx_tools.dynamic_docx_tools import _replace_placeholders_in_document  # noqa: E402
from docx_tools.markdown_processor import process_markdown_content  # noqa: E402
from docx_tools.patterns import (  # noqa: E402
    contains_block_markdown,
    expand_br_to_block_breaks,
    line_starts_block,
)
from markdown_table import is_separator_cells, is_table_line, split_row  # noqa: E402
from xlsx_tools.base_xlsx_tool import _markdown_to_excel_buffer  # noqa: E402
from xlsx_tools.helpers import parse_table as excel_parse_table  # noqa: E402

ISSUE_TABLE = "| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4\n"


def render(markdown):
    """Render *markdown* into a fresh document; return (doc, warning codes)."""
    doc = Document()
    warnings = W.channel()
    process_markdown_content(doc, markdown, warnings=warnings)
    return doc, [w.code for w in warnings]


def rows_of(table):
    return [[cell.text for cell in row.cells] for row in table.rows]


def prose(doc):
    return [p.text for p in doc.paragraphs if p.text]


# ---------------------------------------------------------------------------
# The shared row grammar
# ---------------------------------------------------------------------------

class TestMarkdownTable:

    @pytest.mark.parametrize("line", ["| a | b |", "| a | b", "|", "||", "|---|---"])
    def test_a_line_starting_with_a_pipe_is_a_table_line(self, line):
        assert is_table_line(line)

    @pytest.mark.parametrize("line", ["", "a | b |", "text", "- | a |"])
    def test_a_line_not_starting_with_a_pipe_is_not(self, line):
        assert not is_table_line(line)

    @pytest.mark.parametrize("line", ["| a | b |", "| a | b", "  | a | b  ", "|a|b"])
    def test_the_trailing_pipe_is_optional_when_splitting(self, line):
        assert split_row(line) == ["a", "b"]

    def test_empty_cells_are_kept(self):
        assert split_row("| a |  | c") == ["a", "", "c"]

    def test_a_lone_pipe_is_a_row_with_no_cells(self):
        assert split_row("|") == []

    def test_text_after_the_last_pipe_is_a_cell(self):
        """Not dropped, as the old Word split did: it is the row's last cell."""
        assert split_row("|---|---|x") == ["---", "---", "x"]

    @pytest.mark.parametrize("cells", [
        ["---", "---"], ["-"], [":--", ":-:", "--:"], ["-", ":---------:"],
    ])
    def test_separator_cells(self, cells):
        assert is_separator_cells(cells)

    @pytest.mark.parametrize("cells", [
        [], [""], ["", ""], ["---", ""], ["---", "x"], ["-:-"], ["::"], ["- -"],
    ])
    def test_not_separator_cells(self, cells):
        assert not is_separator_cells(cells)

    def test_a_separator_with_text_after_it_is_not_one(self):
        assert not is_separator_cells(split_row("|---|---|x"))


# ---------------------------------------------------------------------------
# Word
# ---------------------------------------------------------------------------

class TestWordKeepsTheRow:

    def test_the_issue_table_has_three_rows_and_no_stray_paragraph(self):
        doc, codes = render(ISSUE_TABLE)

        assert len(doc.tables) == 1
        assert rows_of(doc.tables[0]) == [["A", "B"], ["1", "2"], ["3", "4"]]
        assert not any("|" in text for text in prose(doc))
        assert codes == []

    def test_the_issue_table_through_the_buffer_function(self):
        buffer, warnings = _markdown_to_word_buffer("Intro\n\n" + ISSUE_TABLE + "\nAfter\n")

        doc = Document(buffer)
        assert rows_of(doc.tables[0]) == [["A", "B"], ["1", "2"], ["3", "4"]]
        assert "| 3 | 4" not in "\n".join(prose(doc))
        assert warnings == []

    def test_a_header_without_its_trailing_pipe(self):
        """Used to build a table whose header read ``---|---``, print the real
        header as prose and report table_separator_missing."""
        doc, codes = render("| A | B\n|---|---|\n| 1 | 2 |\n")

        assert rows_of(doc.tables[0]) == [["A", "B"], ["1", "2"]]
        assert prose(doc) == []
        assert codes == []

    def test_a_separator_without_its_trailing_pipe(self):
        doc, codes = render("| A | B |\n|:---|---:\n| 1 | 2 |\n")

        assert rows_of(doc.tables[0]) == [["A", "B"], ["1", "2"]]
        assert codes == []

    def test_no_trailing_pipe_anywhere(self):
        doc, codes = render("| A | B\n|---|---\n| 1 | 2\n| 3 | 4\n")

        assert rows_of(doc.tables[0]) == [["A", "B"], ["1", "2"], ["3", "4"]]
        assert prose(doc) == []
        assert codes == []

    def test_a_ragged_row_without_its_trailing_pipe(self):
        doc, codes = render("| A | B | C |\n|---|---|---|\n| 1 | 2\n")

        assert rows_of(doc.tables[0]) == [["A", "B", "C"], ["1", "2", ""]]
        assert prose(doc) == []

    def test_a_lone_pipe_mid_table_stays_a_row(self):
        """Excel's rule: any line starting with ``|`` is a row, an empty one
        here. Pins what both tools already did."""
        doc, _ = render("| A | B |\n|---|---|\n|\n| 1 | 2 |\n")

        assert rows_of(doc.tables[0]) == [["A", "B"], ["", ""], ["1", "2"]]

    def test_a_lone_row_without_trailing_pipe_is_reported(self):
        """Used to be written as prose with nothing said."""
        doc, codes = render("intro\n\n| x\n")

        assert doc.tables == []
        assert prose(doc) == ["intro", "| x"]
        assert codes == [W.TABLE_NOT_RECOGNISED]

    def test_a_br_in_a_row_without_trailing_pipe_stays_in_the_cell(self):
        """The <br> pre-pass leaves table rows alone; a row missing its
        trailing pipe is one too, or ``- y`` would be torn out as a list."""
        doc, codes = render("| A | B |\n|---|---|\n| x<br>- y | z\n")

        assert rows_of(doc.tables[0]) == [["A", "B"], ["x\n- y", "z"]]
        assert codes == []


class TestWordProseWithPipesIsUnchanged:
    """Block detection before parsing reads a pipe line strictly."""

    def test_line_starts_block_is_false_for_pipe_prose(self):
        assert not line_starts_block(["|x| < 5"], 0)
        assert not line_starts_block(["Note:  ", "|x| < 5 holds"], 1)

    def test_contains_block_markdown_is_false_for_pipe_prose(self):
        assert not contains_block_markdown("|x| < 5")
        assert not contains_block_markdown("so |x| < 5 holds")

    def test_a_full_pipe_row_is_still_a_block(self):
        assert line_starts_block(["| a | b |"], 0)
        assert contains_block_markdown("| a | b |")

    def test_a_header_without_trailing_pipe_starts_a_block(self):
        """Another pipe line follows it, so it is a table and must not be
        swallowed into the paragraph above."""
        assert line_starts_block(["| A | B", "|---|---|"], 0)

    def test_a_soft_break_run_still_takes_pipe_prose(self):
        doc, _ = render("Given  \n|x| < 5 holds\n")

        assert prose(doc) == ["Given\n|x| < 5 holds"]

    def test_a_soft_break_run_stops_before_a_table_missing_its_trailing_pipe(self):
        doc, _ = render("Given  \n| A | B\n|---|---|\n| 1 | 2 |\n")

        assert prose(doc) == ["Given"]
        assert rows_of(doc.tables[0]) == [["A", "B"], ["1", "2"]]

    def test_a_br_in_lone_pipe_prose_is_still_promoted(self):
        assert expand_br_to_block_breaks("|x| < 5<br>- item") == "|x| < 5\n- item"

    @pytest.mark.parametrize("template", [
        "The bound {{cond}} holds.",
        "{{cond}}",
    ])
    def test_a_placeholder_value_with_pipes_stays_inline(self, template):
        doc = Document()
        doc.add_paragraph(template)

        _replace_placeholders_in_document(doc, {"cond": "|x| < 5"})

        assert [p.text for p in doc.paragraphs] == [template.replace("{{cond}}", "|x| < 5")]
        assert doc.tables == []


# ---------------------------------------------------------------------------
# Word and Excel agree
# ---------------------------------------------------------------------------

SAME_ROWS = [
    ISSUE_TABLE,
    "| A | B\n|---|---|\n| 1 | 2 |\n",
    "| A | B |\n|---|---\n| 1 | 2 |\n",
    "| A | B\n|---|---\n| 1 | 2\n| 3 | 4\n",
    "| A | B | C |\n|---|---|---|\n| 1 | 2\n",
    "| A | B |\n|---|---|\n|\n| 1 | 2 |\n",
    "| A | B |\n|---|---|x\n| 1 | 2 |\n",
    "| A | B |\n| 1 | 2\n",
]


class TestWordAndExcelAgree:

    @pytest.mark.parametrize("markdown", SAME_ROWS)
    def test_the_two_parsers_give_the_same_rows(self, markdown):
        lines = markdown.split("\n")
        word_rows, _, word_next, word_sep = word_parse_table(lines, 0)
        excel_rows, excel_next = excel_parse_table(lines, 0)

        assert word_rows == list(excel_rows)
        assert word_next == excel_next
        assert word_sep == excel_rows.has_separator

    def test_excel_keeps_the_issue_tables_last_row(self):
        """Pins Excel's existing behaviour: it already kept the row."""
        from openpyxl import load_workbook

        buffer, warnings = _markdown_to_excel_buffer(ISSUE_TABLE)
        ws = load_workbook(buffer).active

        assert [[c.value for c in row] for row in ws.iter_rows()] == [
            ["A", "B"], [1, 2], [3, 4]]
        assert warnings == []

    def test_a_lone_pipe_mid_table_stays_a_row_in_excel(self):
        rows, _ = excel_parse_table("| A | B |\n|---|---|\n|\n| 1 | 2 |".split("\n"), 0)

        assert list(rows) == [["A", "B"], [], ["1", "2"]]
