"""A percent column reads European decimals the way a number column does (#143).

``types: number`` turned ``4,3`` into 4.3 while ``types: percent`` left the
same ``4,3`` as text, and an untyped ``4,3%`` stayed text too: three places
decided "is this a percent" with a bare ``float()``, and none of them knew the
decimal comma. They now share ``_parse_percent_body()``.

The one form that is refused is a lone comma before exactly three digits with
no dot (``0,125``): 0.125 to a European reader and 125 to an English one, so
guessing either is a 1000x error for the other. It stays text, as it did, and
a typed column now says so — once per column, with a count — instead of
handing back a success response with a string where a number was declared.

Every build here goes through ``_markdown_to_excel_buffer()``, the path the
tool handler uses, and reads the saved workbook back.
"""
import io

import pytest
from openpyxl import load_workbook

from xlsx_tools import _markdown_to_excel_buffer
from xlsx_tools import warnings as W


def build(markdown):
    """Build *markdown*; return (first worksheet, warnings)."""
    buffer, warnings = _markdown_to_excel_buffer(markdown)
    try:
        workbook = load_workbook(io.BytesIO(buffer.getvalue()))
    finally:
        buffer.close()
    return workbook.active, list(warnings)


def table(types, *rows):
    """A two-column table (label, value) under a ``types`` directive."""
    lines = [f"<!-- types: {types} -->"] if types else []
    lines += ["| Item | Rate |", "|------|------|"]
    lines += [f"| r{i} | {value} |" for i, value in enumerate(rows, start=1)]
    return "\n".join(lines) + "\n"


def kept_as_text(warnings):
    return [w for w in warnings if w.code == W.TYPED_VALUE_KEPT_AS_TEXT]


class TestTypedPercentColumn:

    def test_a_european_decimal_is_a_percent(self):
        ws, warnings = build(table("text, percent", "4,3"))

        assert ws["B2"].value == pytest.approx(0.043)
        assert ws["B2"].number_format == "0.0%"
        assert kept_as_text(warnings) == []

    def test_dot_grouping_with_a_decimal_comma(self):
        """Decimals are counted on the normalised `1234.56`, not on the raw
        text, where `1.234,56` would have looked like six of them."""
        ws, _ = build(table("text, percent", "1.234,56"))

        assert ws["B2"].value == pytest.approx(12.3456)
        assert ws["B2"].number_format == "0.00%"

    @pytest.mark.parametrize("raw", ["4,3 %", "4,3 %"])
    def test_a_space_before_the_sign(self, raw):
        """Czech typography puts a (often non-breaking) space between the
        number and the `%`."""
        ws, _ = build(table("text, percent", raw))

        assert ws["B2"].value == pytest.approx(0.043)
        assert ws["B2"].number_format == "0.0%"

    def test_a_negative_european_decimal_with_the_sign(self):
        ws, _ = build(table("text, percent", "-4,3%"))

        assert ws["B2"].value == pytest.approx(-0.043)
        assert ws["B2"].number_format == "0.0%"

    @pytest.mark.parametrize("raw,expected", [
        ("0,5", 0.005),
        (",5", 0.005),
        ("4,30", 0.043),
        ("3,14159", 0.0314159),
        ("1,234.5", 12.345),
        ("1,234,567", 12345.67),
    ])
    def test_other_comma_forms_read_as_written(self, raw, expected):
        ws, warnings = build(table("text, percent", raw))

        assert ws["B2"].value == pytest.approx(expected)
        assert kept_as_text(warnings) == []

    def test_a_declared_format_applies_to_a_european_decimal(self):
        ws, _ = build(table("text, percent:0.00%", "4,3"))

        assert ws["B2"].value == pytest.approx(0.043)
        assert ws["B2"].number_format == "0.00%"

    def test_a_formula_learns_its_precision_from_european_literals(self):
        """The format learner accepts exactly what the coercion accepts (#126):
        a `4,3` beside the formula is a percent, so it sets the precision."""
        ws, _ = build(table("text, percent", "4,3", "=B2*2"))

        assert ws["B3"].number_format == "0.0%"

    def test_number_and_percent_columns_agree(self):
        """The same European text, read by both column types, is the same
        number — the percent one a hundredth of it."""
        ws, _ = build(
            "<!-- types: text, number, percent -->\n"
            "| Item | Number | Rate |\n|------|--------|------|\n"
            "| a | 4,3 | 4,3 |\n| b | 1.234,56 | 1.234,56 |\n"
        )

        for row in (2, 3):
            assert ws[f"C{row}"].value == pytest.approx(ws[f"B{row}"].value / 100)


class TestTheAmbiguousForm:

    def test_a_comma_before_three_digits_stays_text_and_says_so(self):
        ws, warnings = build(table("text, percent", "0,125"))

        assert ws["B2"].value == "0,125"
        [warning] = kept_as_text(warnings)
        assert warning.severity == "warning"
        assert warning.location == {"sheet": ws.title, "cell": "B2"}
        assert "0,125" in warning.message
        assert "ambiguous" in warning.message

    def test_the_ambiguous_form_teaches_a_formula_nothing(self):
        ws, _ = build(table("text, percent", "0,125", "=B2*2"))

        assert ws["B3"].number_format == "0%"     # not 0.000%


class TestUntypedPercent:

    def test_a_european_decimal_with_a_sign(self):
        ws, _ = build(table(None, "4,3%"))

        assert ws["B2"].value == pytest.approx(0.043)
        assert ws["B2"].number_format == "0.0%"

    def test_dot_grouping_with_a_decimal_comma(self):
        ws, _ = build(table(None, "1.234,56%"))

        assert ws["B2"].value == pytest.approx(12.3456)
        assert ws["B2"].number_format == "0.00%"

    def test_the_ambiguous_form_stays_text(self):
        ws, warnings = build(table(None, "1,234%"))

        assert ws["B2"].value == "1,234%"
        assert kept_as_text(warnings) == []       # no type was declared

    def test_without_the_sign_it_is_not_a_percent(self):
        """Untyped, the `%` is what makes it a percent; `4,3` alone keeps the
        untyped-number rule, where a lone comma is not guessed at."""
        ws, _ = build(table(None, "4,3"))

        assert ws["B2"].value == "4,3"

    @pytest.mark.parametrize("raw", ["inf%", "nan%", "1_0%"])
    def test_non_finite_and_underscores_are_refused(self, raw):
        ws, _ = build(table(None, raw))

        assert ws["B2"].value == raw


class TestTypedValueKeptAsText:
    """Every typed column, not only percent, reports a value it kept as text."""

    def test_one_warning_per_column_with_the_count(self):
        ws, warnings = build(
            "<!-- types: text, number -->\n"
            "| Item | Amount |\n|------|--------|\n"
            "| a | - |\n| b | 10 |\n| c | n/a |\n| d | — |\n"
        )

        [warning] = kept_as_text(warnings)
        assert warning.location == {"sheet": ws.title, "cell": "B2"}
        assert warning.message.startswith("3 values in column B ('Amount')")
        assert "'-'" in warning.message
        assert "leave the cell empty" in warning.message

    def test_each_column_reports_on_its_own(self):
        _, warnings = build(
            "<!-- types: text, number, percent -->\n"
            "| Item | Amount | Rate |\n|------|--------|------|\n"
            "| a | - | - |\n| b | - | 5 |\n"
        )

        assert [w.location["cell"] for w in kept_as_text(warnings)] == ["B2", "C2"]

    @pytest.mark.parametrize("types,value", [
        ("text, currency:$", "about $5"),
        ("text, date", "someday"),
        ("text, bool", "maybe"),
        ("text, percent", "lots"),
    ])
    def test_every_coercing_type_reports(self, types, value):
        _, warnings = build(table(types, value))

        assert len(kept_as_text(warnings)) == 1

    def test_a_text_column_never_reports(self):
        _, warnings = build(table("text, text", "anything"))

        assert kept_as_text(warnings) == []

    def test_an_empty_typed_cell_never_reports(self):
        _, warnings = build(table("text, number", "", "5"))

        assert kept_as_text(warnings) == []

    def test_a_formula_cell_never_reports(self):
        _, warnings = build(table("text, percent", "5", "=B2*2"))

        assert kept_as_text(warnings) == []


class TestNoNewSilentNumbers:
    """Review of #143: forms master kept as text must not become numbers now.

    Untyped cells declare nothing and cannot warn, so a loose reading there
    is a silent wrong number; a typed one warns, but only if it refuses.
    """

    @pytest.mark.parametrize("raw", ["1,5e3%", "1,234e2%", "10, 20%", "4 3%", "1, 5%",
                                     # two separators that are not digit grouping
                                     "4,3.%", "1,5.0%", "12,34.5%", "1.5,2%", "1,2,345%", "1.,5%"])
    def test_untyped_forms_master_kept_as_text_still_are(self, raw):
        ws, _ = build(table(None, raw))
        assert ws["B2"].value == raw

    @pytest.mark.parametrize("raw", ["1,5e3", "10, 20", "4 3", "4,3.", "1.5,2"])
    def test_typed_forms_are_refused_and_reported(self, raw):
        ws, warnings = build(table("text, percent", raw))
        assert ws["B2"].value == raw
        assert len(kept_as_text(warnings)) == 1

    @pytest.mark.parametrize("raw, expected", [
        ("1 234,56", 12.3456),        # a space grouping thousands
        ("1 234,5", 12.345),     # a non-breaking one
        ("−4,3", -0.043),        # the typographic minus sign
        ("+4,3", 0.043),
        ("1e3", 10.0),                # an exponent with no comma, as on master
    ])
    def test_typographic_and_plain_forms_read(self, raw, expected):
        ws, _ = build(table("text, percent", raw))
        assert ws["B2"].value == pytest.approx(expected)


class TestTheAmbiguityHint:

    def test_the_hint_appears_when_any_kept_value_is_ambiguous(self):
        """The first kept value is `-`; the comma hint must still be there."""
        _, warnings = build(table("text, percent", "-", "0,125"))
        [w] = kept_as_text(warnings)
        assert "2 values" in w.message
        assert "0.125" in w.message, w.message

    def test_two_typed_tables_on_one_sheet_report_separately(self):
        md = table("text, percent", "x") + "\n" + table("text, percent", "y")
        _, warnings = build(md)
        assert len(kept_as_text(warnings)) == 2
