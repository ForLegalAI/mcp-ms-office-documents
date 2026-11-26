"""Number and currency columns read a decimal comma by the table's evidence (#211).

``types: number`` and ``types: currency`` read ``0,125`` as 125: a lone comma
before exactly three digits was always taken for a thousands separator, and
the looser grammar of those columns also turned ``1,5e3`` into 15000,
``10 20`` into 1020 and ``nan`` into an empty cell — all without a word.

They now share the percent column's reader, ``_parse_decimal_text()`` (#143):

- ``0,125``, ``,125`` and ``1234,567`` are decimals everywhere — no locale
  groups thousands behind a zero or in a group of four;
- malformed values stay text and are reported as ``typed_value_kept_as_text``;
- the one truly ambiguous form, ``4,375``, is settled by the table's own
  numbers: European evidence (``2,5``) makes it 4.375, English evidence
  (``1.5``, ``1,234.5``) 4375, both leave it as text with a hint, and none
  reads it as 4375 and says so with ``ambiguous_comma_read_as_thousands``.

Untyped cells keep refusing to guess at a lone comma, and a lead group with a
leading zero (``0,125``, ``01,234``) is no longer taken for grouping there.

Each test says whether it is a **fix** (it fails on the code before #211, for
the reason in its docstring) or a **pin** (behaviour that must not move).
Every build goes through ``_markdown_to_excel_buffer()``, the path the tool
handler uses, and reads the saved workbook back.
"""
import io

import pytest
from openpyxl import Workbook, load_workbook

from xlsx_tools import _markdown_to_excel_buffer
from xlsx_tools import warnings as W
from xlsx_tools.helpers import _apply_column_type, _parse_percent_body

NBSP = chr(0x00A0)        # no-break space
NNBSP = chr(0x202F)       # narrow no-break space
MINUS = chr(0x2212)       # typographic minus sign


def build(markdown):
    """Build *markdown*; return (first worksheet, warnings)."""
    buffer, warnings = _markdown_to_excel_buffer(markdown)
    try:
        workbook = load_workbook(io.BytesIO(buffer.getvalue()))
    finally:
        buffer.close()
    return workbook.active, list(warnings)


def table(types, *rows, header="| Item | Value |"):
    """A table under a ``types`` directive; a row is a value or a tuple of them."""
    lines = [f"<!-- types: {types} -->"] if types else []
    width = header.count("|") - 1
    lines += [header, "|" + "---|" * width]
    for i, row in enumerate(rows, start=1):
        cells = row if isinstance(row, tuple) else (row,)
        lines.append("| " + " | ".join((f"r{i}",) + cells) + " |")
    return "\n".join(lines) + "\n"


def column(ws, letter, rows):
    return [ws[f"{letter}{row}"].value for row in range(2, 2 + rows)]


def of(code, warnings):
    return [w for w in warnings if w.code == code]


def kept(warnings):
    return of(W.TYPED_VALUE_KEPT_AS_TEXT, warnings)


# The new code by its public name, so this file also runs — and fails for the
# right reasons — against the code before #211, which has no such constant.
READ_AS_THOUSANDS = "ambiguous_comma_read_as_thousands"


def guessed(warnings):
    return of(READ_AS_THOUSANDS, warnings)


def test_the_code_is_declared_with_info_severity():
    """Fix: the code is new in #211."""
    assert W.WARNING_SEVERITY.get(READ_AS_THOUSANDS) == "info"


# ── Fixes ─────────────────────────────────────────────────────────────────────


class TestALeadGroupThatCannotBeGrouping:
    """Fix: master read every lone comma before three digits as thousands."""

    @pytest.mark.parametrize("types,raw,expected", [
        ("text, number", "0,125", 0.125),
        ("text, number", "-0,125", -0.125),
        ("text, number", ",125", 0.125),
        ("text, number", "0,001", 0.001),
        ("text, number", "1234,567", 1234.567),
        ("text, currency:€", "0,125 €", 0.125),
        ("text, currency:€", "-0,125 €", -0.125),
        ("text, currency:€", ",125 €", 0.125),
        ("text, currency:€", "0,001 €", 0.001),
        ("text, currency:€", "1234,567 €", 1234.567),
    ])
    def test_is_a_decimal(self, types, raw, expected):
        """Fix: master gave 125, -125, 125, 1 and 1234567."""
        ws, warnings = build(table(types, raw))

        assert ws["B2"].value == pytest.approx(expected)
        assert kept(warnings) == [] and guessed(warnings) == []

    @pytest.mark.parametrize("raw", ["0,125", "-0,125", "0,001",
                                     "0,125,000", "0,125.5", "0.125,5", "0.125.000",
                                     "00,125", "01,234"])
    def test_untyped_it_is_not_digit_grouping(self, raw):
        """Fix: master's untyped grouping took any one-to-three-digit lead
        group, so these became 125, -125, 1, 125000, 125.5, 125.5, 125000, 125
        and 1234. A lone comma is not guessed at untyped, so they stay text."""
        ws, warnings = build(table(None, raw))

        assert ws["B2"].value == raw
        assert warnings == []

    def test_grouping_behind_a_zero_is_refused_in_a_typed_column(self):
        """Fix: master read typed `0,125,000` as 125000, silently."""
        ws, warnings = build(table("text, number", "0,125,000"))

        assert ws["B2"].value == "0,125,000"
        assert len(kept(warnings)) == 1


class TestMalformedValuesAreRefused:
    """Fix: master stripped separators without checking the grammar."""

    @pytest.mark.parametrize("raw,master", [
        ("1,5e3", 15000), ("4,3.", 43), ("1.5,2", 15.2),
        ("10 20", 1020), ("10, 20", 10.2), ("1 e3", 1000), ("1_000", 1000),
    ])
    @pytest.mark.parametrize("types,suffix", [("text, number", ""),
                                              ("text, currency:€", " €")])
    def test_stays_text_and_is_reported(self, raw, master, types, suffix):
        """Fix: each was a silent number on master (*master*)."""
        ws, warnings = build(table(types, raw + suffix))

        assert ws["B2"].value == raw + suffix
        [warning] = kept(warnings)
        assert warning.location == {"sheet": ws.title, "cell": "B2"}

    @pytest.mark.parametrize("raw", ["nan", "inf", "-inf", "Infinity"])
    @pytest.mark.parametrize("types,suffix", [("text, number", ""),
                                              ("text, currency:€", " €")])
    def test_non_finite_values_are_kept_not_emptied(self, raw, types, suffix):
        """Fix: master stored a float NaN/inf, which serialises as an empty
        cell — the value was lost and nothing was reported."""
        ws, warnings = build(table(types, raw + suffix))

        assert ws["B2"].value == raw + suffix
        assert len(kept(warnings)) == 1

    def test_one_warning_per_column_however_many(self):
        """Fix: the column reports once, with the count."""
        ws, warnings = build(table("text, number", "1,5e3", "4,3.", "nan", "5"))

        [warning] = kept(warnings)
        assert warning.message.startswith("3 values in column B ('Value')")


class TestTypographicSpacesAndMinus:
    """Fix: master removed only ASCII spaces, so these stayed text."""

    @pytest.mark.parametrize("types,raw,expected", [
        ("text, number", f"1{NBSP}234,5", 1234.5),
        ("text, number", f"1{NNBSP}234,5", 1234.5),
        ("text, number", f"{MINUS}5", -5),
        ("text, number", f"{MINUS}1{NBSP}234,5", -1234.5),
        ("text, currency:Kč", f"1{NBSP}234,50{NBSP}Kč", 1234.5),
        ("text, currency:Kč", f"1{NNBSP}234,50 Kč", 1234.5),
        ("text, currency:Kč", f"{MINUS}1 234 Kč", -1234),
    ])
    def test_reads(self, types, raw, expected):
        """Fix: kept as text on master."""
        ws, warnings = build(table(types, raw))

        assert ws["B2"].value == pytest.approx(expected)
        assert kept(warnings) == []


class TestEuropeanEvidence:
    """Fix: with a decimal comma in the table, `4,375` is 4.375, not 4375."""

    def test_in_the_column_itself(self):
        """Fix: master read `4,375` as 4375 beside a `2,5` it read as 2.5."""
        ws, warnings = build(table("text, number", "4,375", "2,5"))

        assert column(ws, "B", 2) == [pytest.approx(4.375), pytest.approx(2.5)]
        assert warnings == []

    def test_with_a_declared_number_format(self):
        """Fix: the format is the caller's, the reading is the evidence's."""
        ws, _ = build(table("text, number:0.000", "4,375", "2,5"))

        assert ws["B2"].value == pytest.approx(4.375)
        assert ws["B2"].number_format == "0.000"

    def test_from_another_typed_column_of_the_table(self):
        """Fix: the `Kč` column's `12 500,00` writes a decimal comma, so the
        rate column beside it — with no evidence of its own — is European."""
        ws, warnings = build(table(
            "text, number, currency:Kč", ("4,375", "12 500,00 Kč"),
            header="| Item | Rate | Amount |"))

        assert ws["B2"].value == pytest.approx(4.375)
        assert ws["C2"].value == pytest.approx(12500)
        assert warnings == []

    def test_from_an_untyped_column_of_the_table(self):
        """Fix: an untyped `2,5` stays text itself, but it shows how the
        table writes numbers."""
        ws, warnings = build(table(
            "text, number", ("4,375", "2,5"), header="| Item | Rate | Note |"))

        assert ws["B2"].value == pytest.approx(4.375)
        assert ws["C2"].value == "2,5"
        assert guessed(warnings) == []

    def test_survives_markdown_and_ignores_formulas(self):
        """Fix: the bold `**2,5**` is evidence as the cell loop reads it; the
        formula row is neither evidence nor a value to report."""
        ws, warnings = build(table("text, number", "4,375", "**2,5**", "=B2*2"))

        assert ws["B2"].value == pytest.approx(4.375)
        assert ws["B3"].value == pytest.approx(2.5)
        assert ws["B4"].value == "=B2*2"
        assert warnings == []

    def test_the_mirror_dot_is_thousands(self):
        """Fix: in a table that writes `2,50 Kč`, `1.234 Kč` is 1234 Kč;
        master read it as 1.234."""
        ws, warnings = build(table("text, currency:Kč", "1.234 Kč", "2,50 Kč"))

        assert ws["B2"].value == pytest.approx(1234)
        assert ws["B3"].value == pytest.approx(2.5)
        assert warnings == []

    def test_two_tables_on_one_sheet_take_their_own_evidence(self):
        """Fix: evidence is per table; the European one reads 4.375, the
        English one beside it 4375."""
        european = table("text, number", "4,375", "2,5")
        english = table("text, number", "4,375", "1.5")
        ws, warnings = build(european + "\n" + english)

        assert ws["B2"].value == pytest.approx(4.375)
        # the second table starts two rows below the first one's last row
        second = [row for row in ws.iter_rows(min_row=6, max_col=2, values_only=True)]
        assert second[1][1] == pytest.approx(4375)
        assert warnings == []


class TestNoEvidence:
    """Fix: with nothing to go on, `4,375` is still 4375 — and now says so."""

    def test_a_czech_column_is_read_as_thousands_and_reported_once(self):
        """Fix: master read all three as thousands with no warning."""
        ws, warnings = build(table("text, number", "4,375", "3,125", "1,250"))

        assert column(ws, "B", 3) == [4375, 3125, 1250]
        assert kept(warnings) == []
        [warning] = guessed(warnings)
        assert warning.severity == "info"
        assert warning.location == {"sheet": ws.title, "cell": "B2"}
        assert warning.message.startswith("3 values in column B ('Value')")
        assert "'4,375' at B2, read as 4375" in warning.message
        assert "4.375" in warning.message and "(4375)" in warning.message

    def test_one_value_names_both_readings(self):
        """Fix: no warning on master; the hint names both readings."""
        ws, warnings = build(table("text, currency:$", "($1,234)"))

        assert ws["B2"].value == pytest.approx(-1234)
        [warning] = guessed(warnings)
        assert warning.message.startswith("'($1,234)' in column B ('Value') was read as -1234")
        assert "write 1.234" in warning.message and "write 1234" in warning.message

    def test_formula_and_empty_cells_are_not_counted(self):
        """Fix: one value, so the singular message — the formula and the
        empty cell are neither values nor evidence."""
        ws, warnings = build(table("text, number", "4,375", "=B2*2", ""))

        [warning] = guessed(warnings)
        assert warning.message.startswith("'4,375' in column B")

    def test_each_column_reports_on_its_own(self):
        """Fix: no warning on master; one per table column."""
        _, warnings = build(table("text, number, number", ("1,234", "5,678"),
                                  header="| Item | A | B |"))

        assert [w.location["cell"] for w in guessed(warnings)] == ["B2", "C2"]


class TestMixedEvidence:
    """Fix: a table that writes numbers both ways cannot settle `1,234`."""

    def test_the_ambiguous_value_stays_text_with_a_hint(self):
        """Fix: master read `1,234` as 1234 beside a `2,5` and a `1,234.5`
        that contradict each other."""
        ws, warnings = build(table("text, number", "1,234", "2,5", "1,234.5"))

        assert column(ws, "B", 3) == ["1,234", pytest.approx(2.5), pytest.approx(1234.5)]
        [warning] = kept(warnings)
        assert warning.location == {"sheet": ws.title, "cell": "B2"}
        assert "ambiguous" in warning.message and "both ways" in warning.message
        assert "(1.234)" in warning.message
        assert guessed(warnings) == []


class TestPercent:
    """Fix: a percent takes the same classification (#143 refused `0,125`)."""

    @pytest.mark.parametrize("raw,expected", [
        ("0,125", 0.00125), (",125", 0.00125), ("1234,567", 12.34567),
    ])
    def test_a_lead_group_that_cannot_be_grouping_is_a_decimal(self, raw, expected):
        """Fix: kept as text on master."""
        ws, warnings = build(table("text, percent", raw))

        assert ws["B2"].value == pytest.approx(expected)
        assert ws["B2"].number_format == "0.000%"
        assert warnings == []

    def test_untyped_too(self):
        """Fix: an untyped `0,125%` is the same percent."""
        ws, _ = build(table(None, "0,125%"))

        assert ws["B2"].value == pytest.approx(0.00125)

    def test_the_truly_ambiguous_form_still_stays_text(self):
        """Pin: no evidence pass for a percent — `4,375` is refused even in a
        European table."""
        ws, warnings = build(table("text, percent", "4,375", "2,5"))

        assert ws["B2"].value == "4,375"
        assert len(kept(warnings)) == 1


# ── Pins ──────────────────────────────────────────────────────────────────────


class TestEnglishEvidenceAndDefaults:
    """Pin: English tables read as they always did, and without a warning."""

    def test_comma_grouping_beside_it(self):
        ws, warnings = build(table("text, number", "1,234", "1,234.5"))

        assert column(ws, "B", 2) == [1234, pytest.approx(1234.5)]
        assert warnings == []

    def test_a_dollar_column(self):
        ws, warnings = build(table("text, currency:$", "$1,234", "$12.50"))

        assert column(ws, "B", 2) == [1234, pytest.approx(12.5)]
        assert warnings == []

    def test_a_dot_with_three_decimals_is_a_decimal_without_evidence(self):
        ws, warnings = build(table("text, number", "1.234"))

        assert ws["B2"].value == pytest.approx(1.234)
        assert warnings == []

    def test_apply_column_type_still_reads_thousands_by_default(self):
        """The keyword-only policy defaults to the no-evidence reading, so
        a positional call is unchanged."""
        cell = Workbook().active["A1"]
        _apply_column_type(cell, "1,234", "number")
        assert cell.value == 1234


class TestCurrencyFormsThatAlreadyRead:
    """Pin: the currency shapes master read keep reading the same."""

    @pytest.mark.parametrize("types,raw,expected", [
        ("text, currency:Kč", "1 234,50 Kč", 1234.5),
        ("text, currency:Kč", "Kč 1 234,50", 1234.5),
        ("text, currency:Kč", "1 234,50Kč", 1234.5),
        ("text, currency:Kč", "1 234 567 Kč", 1234567),
        ("text, currency:Kč", "1 234 567,5 Kč", 1234567.5),
        ("text, currency:€", "1.234,50 €", 1234.5),
        ("text, currency:€", "€1,234.50", 1234.5),
        ("text, currency:€", "5 €", 5),
        ("text, currency:€", "-€5", -5),
        ("text, currency:$", "$ 1,234.50", 1234.5),
        ("text, currency:$", "-$1,234.50", -1234.5),
        ("text, currency:$", "$-1,234.50", -1234.5),
        ("text, currency:$", "($1,234.56)", -1234.56),
        ("text, currency:$", "$(1,234.56)", -1234.56),
        ("text, currency:$", "( $1,234.56 )", -1234.56),
        ("text, currency:$", "$ (1 234)", -1234),
        ("text, currency:$", "(1 234,50)", -1234.5),
    ])
    def test_reads_as_before(self, types, raw, expected):
        ws, warnings = build(table(types, raw))

        assert ws["B2"].value == pytest.approx(expected)
        assert kept(warnings) == []

    @pytest.mark.parametrize("types,raw,expected", [
        ("text, currency:€", "- 5 €", -5),
        ("text, currency:€", "€ - 5", -5),
        ("text, currency:€", "€ -5", -5),
        ("text, currency:$", "- $1,234.50", -1234.5),
        ("text, number", "- 5", -5),
        ("text, number", "+ 5", 5),
    ])
    def test_a_sign_apart_from_its_number(self, types, raw, expected):
        ws, warnings = build(table(types, raw))

        assert ws["B2"].value == pytest.approx(expected)
        assert kept(warnings) == []

    @pytest.mark.parametrize("raw", ["EUR 5", "5,- €"])
    def test_what_was_text_stays_text(self, raw):
        ws, _ = build(table("text, currency:€", raw))

        assert ws["B2"].value == raw


class TestUntypedGrouping:
    """Pin: untyped digit grouping that was certain stays as it was."""

    @pytest.mark.parametrize("raw,expected", [
        ("1,234", 1234), ("-1,234", -1234), ("1,234.56", 1234.56),
        ("1.234,56", 1234.56), ("1.234.567", 1234567),
        ("4,375", 4375),       # the documented gap: untyped cannot learn
    ])
    def test_reads(self, raw, expected):
        ws, warnings = build(table(None, raw))

        assert ws["B2"].value == pytest.approx(expected)
        assert warnings == []

    @pytest.mark.parametrize("raw", ["1,5", ",125", "1 234,5"])
    def test_a_lone_comma_is_not_guessed_at(self, raw):
        ws, _ = build(table(None, raw))

        assert ws["B2"].value == raw


# ── Differential: every #143 percent case ─────────────────────────────────────

# Every value tests/test_xlsx_percent_european_143.py hands a percent, with
# what _parse_percent_body() returned for it on master (8d592bd): the
# fraction, rounded, and the normalised text, or None.
MASTER_143 = {
    "4,3 %": (0.043, "4.3"), f"4,3{NBSP}%": (0.043, "4.3"),
    "0,5": (0.005, "0.5"), ",5": (0.005, ".5"), "4,30": (0.043, "4.30"),
    "3,14159": (0.0314159, "3.14159"), "1,234.5": (12.345, "1234.5"),
    "1,234,567": (12345.67, "1234567"),
    "inf%": None, "nan%": None, "1_0%": None, "about $5": None,
    "someday": None, "maybe": None, "lots": None,
    "1,5e3%": None, "1,234e2%": None, "10, 20%": None, "4 3%": None,
    "1, 5%": None, "4,3.%": None, "1,5.0%": None, "12,34.5%": None,
    "1.5,2%": None, "1,2,345%": None, "1.,5%": None,
    "1,5e3": None, "10, 20": None, "4 3": None, "4,3.": None, "1.5,2": None,
    "1 234,56": (12.3456, "1234.56"), f"1{NBSP}234,5": (12.345, "1234.5"),
    f"{MINUS}4,3": (-0.043, "-4.3"), "+4,3": (0.043, "+4.3"),
    "1e3": (10.0, "1e3"), "4,3": (0.043, "4.3"),
    "1.234,56": (12.3456, "1234.56"), "-4,3%": (-0.043, "-4.3"),
    "0,125": None, "4,3%": (0.043, "4.3"), "1.234,56%": (12.3456, "1234.56"),
    "1,234%": None, "anything": None, "": None, "5": (0.05, "5"),
    "-": None, "y": None, "x": None,
}

# The differences #211 intends among those cases, and why.
INTENDED_211 = {
    "0,125": (0.00125, "0.125"),    # a zero lead group is never grouping
}


@pytest.mark.parametrize("raw", sorted(MASTER_143))
def test_every_143_percent_case_reads_as_on_master_but_the_intended(raw):
    """The reader moved from _parse_percent_body into _parse_decimal_text;
    across every #143 case only the listed differences may appear."""
    expected = INTENDED_211.get(raw, MASTER_143[raw])
    parsed = _parse_percent_body(raw.rstrip("%"))
    actual = None if parsed is None else (round(parsed[0], 12), parsed[1])

    assert actual == expected
