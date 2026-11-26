"""Cell-level machinery for the Excel tool.

Everything that happens inside one table once the parser has placed it:
``parse_table`` splits the Markdown rows (the row grammar is the root
``markdown_table`` module, shared with Word); ``resolve_cell`` decides whether a
cell is a formula, percent, number, date or text; ``adjust_formula_references``
rewrites the table-relative formula syntax (``T1.B[0]``, ``B[-1]``,
``Sheet!T1.SUM(...)``) to real cell references; the ``types`` directive
coerces typed columns, whose numbers all go through ``_parse_decimal_text``
and the table's evidence for an ambiguous ``4,375`` (``_comma3_policies``);
``add_table_to_sheet`` writes the table with fonts,
borders, alignment, number formats, column widths, the ``styles`` directive
and an optional Excel Table object.

The substitution order inside ``adjust_formula_references`` is load-bearing;
see docs/development/tools/excel.md ("Formula reference rewriting").
"""
import math
import re
import logging
from dataclasses import dataclass
from datetime import datetime

from dateutil import parser as dateutil_parser
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from markdown_table import is_separator_cells, is_table_line, split_row

from . import warnings as W

logger = logging.getLogger(__name__)

# ── Layout Constants ──────────────────────────────────────────────────────────
TABLE_BOTTOM_SPACING = 2
MIN_COLUMN_WIDTH = 12
MAX_COLUMN_WIDTH = 25
COLUMN_WIDTH_PADDING = 2

# Date formats to try before falling back to dateutil auto-detection.
# Order matters — more specific/common formats first.
# Each entry: (strptime_format, excel_number_format)
DATE_FORMATS: list[tuple[str, str]] = [
    # ISO
    ("%Y-%m-%d", "YYYY-MM-DD"),
    ("%Y-%m-%dT%H:%M:%S", "YYYY-MM-DD HH:MM:SS"),
    ("%Y-%m-%dT%H:%M", "YYYY-MM-DD HH:MM"),
    # European (day first)
    ("%d.%m.%Y", "DD.MM.YYYY"),
    ("%d/%m/%Y", "DD/MM/YYYY"),
    ("%d-%m-%Y", "DD-MM-YYYY"),
    ("%d. %m. %Y", "DD. MM. YYYY"),
    # US (month first)
    ("%m/%d/%Y", "MM/DD/YYYY"),
    # With time
    ("%d.%m.%Y %H:%M", "DD.MM.YYYY HH:MM"),
    ("%d.%m.%Y %H:%M:%S", "DD.MM.YYYY HH:MM:SS"),
    ("%m/%d/%Y %H:%M", "MM/DD/YYYY HH:MM"),
    # Short year
    ("%d.%m.%y", "DD.MM.YY"),
    ("%d/%m/%y", "DD/MM/YY"),
    ("%m/%d/%y", "MM/DD/YY"),
    # Named months
    ("%d %b %Y", "DD MMM YYYY"),
    ("%d %B %Y", "DD MMMM YYYY"),
    ("%b %d, %Y", "MMM DD, YYYY"),
    ("%B %d, %Y", "MMMM DD, YYYY"),
]

# The leading group of real digit grouping: one to three digits, not starting
# with 0. Nobody groups thousands as ``0,125`` or ``01,234``, so a lead group
# with a leading zero is never grouping — in a typed column it marks a decimal
# comma, in an untyped cell it stays text (#211).
_LEAD_GROUP = r'[1-9]\d{0,2}'

# Digit-grouping patterns accepted in columns with no `types` directive.
#
# A lone comma is deliberately NOT enough: "1,5" is 1.5 across most of Europe
# and 15 nowhere, so auto-detection must not guess at it. Only groups of
# exactly three digits are taken:
#     1,234    1,234.56    1,234,567.89      comma grouping, dot decimal
#     1.234,56             1.234.567         dot grouping, comma decimal
# Dot grouping additionally needs a disambiguator — a comma decimal part or a
# second group — because a bare "1.234" is a plain decimal in English and must
# keep parsing as 1.234. The lead group must not start with 0: "0,125" read as
# 125 here, and "0,125,000" as 125000 (#211); both now stay text.
#
# "1,234" itself is still read as 1234 untyped, and so is a Czech "4,375": an
# untyped cell has no column to learn from and nothing to warn with. A column
# of decimal commas declares `types: number`, which reads it with
# _parse_decimal_text and the table's own evidence (add_table_to_sheet).
_GROUPED_COMMA_RE = re.compile(rf'^[+-]?{_LEAD_GROUP}(?:,\d{{3}})+(?:\.\d+)?$')
_GROUPED_DOT_RE = re.compile(
    rf'^[+-]?{_LEAD_GROUP}(?:\.\d{{3}})+,\d+$'      # 1.234,56
    rf'|^[+-]?{_LEAD_GROUP}(?:\.\d{{3}}){{2,}}$'    # 1.234.567
)

# Excel number formats for plain numeric cells (no `types` directive). Only
# applied from 1000 up, where digit grouping is the point; below that the
# General format shows the value as written and is left alone.
THOUSANDS_FORMAT = '#,##0'
THOUSANDS_FORMAT_DECIMALS = '#,##0.00'
_THOUSANDS_THRESHOLD = 1000

# Cap on decimals carried from a percentage's source text into its format.
_PERCENT_MAX_DECIMALS = 4

# ── The numeric grammar of typed columns (#143, #211) ──────────────────────
#
# One reader, _parse_decimal_text, for every column type that holds a number —
# ``number``, ``currency`` and ``percent`` — and for an untyped ``…%``. The
# patterns below are its grammar; the function states the rules in prose.

# What a numeric body may look like once its edges are trimmed and any space
# grouping removed: a sign, digits with dots and commas, and an exponent only
# where there is no comma (``1,5e3`` is refused: read as thousands it gives
# 10x the number). ``nan``, ``inf`` and PEP 515 underscores never match.
_NUMBER_BODY_RE = re.compile(
    r'^[+-]?(?=[\d.,]*\d)[\d.,]+$'
    r'|^[+-]?(?=[\d.]*\d)[\d.]+[eE][+-]?\d+$'
)
# A body with two separators, or several of one, must be genuine digit
# grouping — groups of three after a lead group that does not start with 0 —
# with an optional decimal part in the other separator: ``1,234.5``,
# ``1,234,567``, ``1.234,56``, ``1.234.567``. Without this ``4,3.`` read as 43
# and ``1.5,2`` as 15.2 (#143 review), and ``0,125,000`` as 125000 (#211).
_DIGIT_GROUPED_RE = re.compile(
    _GROUPED_COMMA_RE.pattern
    + rf'|^[+-]?{_LEAD_GROUP}(?:\.\d{{3}})+(?:,\d+)?$'
)
# Indian lakh/crore grouping: a last group of three, every group before it of
# two — ``1,23,456.70``, ``12,34,56,789``. Comma grouping like ``1,234,567``,
# read by typed columns only (``currency:₹`` read it before #211); an untyped
# cell keeps its three-digit grouping, _GROUPED_COMMA_RE.
_INDIAN_GROUPED_RE = re.compile(r'^[+-]?[1-9]\d?(?:,\d{2})+,\d{3}(?:\.\d+)?$')
# Four or more digits, a comma and exactly three: ``1234,567``. A decimal —
# four digits cannot be one thousands group — but not evidence of a decimal
# comma: an English writer who left out a separator writes it too (#211
# review).
_UNGROUPED_COMMA_RE = re.compile(r'^[+-]?\d{4,},\d{3}$')
# Whitespace inside a body is allowed only as digit grouping: ``1 234,56``.
# ``10, 20``, ``10 20`` and ``4 3`` stay text (#143 review, #211).
_SPACE_GROUPED_RE = re.compile(rf'^[+-]?{_LEAD_GROUP}(?: \d{{3}})+(?:[.,]\d+)?$')
# The one truly ambiguous lone comma: a real lead group, a comma, exactly
# three digits and nothing else — ``1,234``, ``4,375``, ``125,500``. Thousands
# grouping in English, a decimal comma in most of Europe, a thousand times
# apart. A lone comma anywhere else is a decimal comma: ``4,3``; ``0,125`` and
# ``,125``, since a zero or empty lead group is never grouping; ``1234,567``,
# since four digits before the comma cannot be one group.
_AMBIGUOUS_COMMA_RE = re.compile(rf'^[+-]?{_LEAD_GROUP},\d{{3}}$')
# Its mirror: ``1.234`` is a decimal in English and 1234 in a European table.
_AMBIGUOUS_DOT_RE = re.compile(rf'^[+-]?{_LEAD_GROUP}\.\d{{3}}$')
# A sign written apart from its number — ``- 5``, or ``€ - 5`` once the symbol
# is gone. Number and currency columns close it up, as they always read it;
# a percent body keeps the #143 rule, under which it stays text.
_SIGN_SPACE_RE = re.compile(r'^([+-]) +(?=[\d.,])')

# Policies for the truly ambiguous ``d,ddd`` (see _parse_decimal_text).
COMMA3_REFUSE = 'refuse'          # stays text: percent, and a mixed table
COMMA3_THOUSANDS = 'thousands'    # 4,375 → 4375: English evidence, or none
COMMA3_DECIMAL = 'decimal'        # 4,375 → 4.375: European evidence
_COMMA3_POLICIES = frozenset({COMMA3_REFUSE, COMMA3_THOUSANDS, COMMA3_DECIMAL})

# What one value proves about how its table writes numbers (_decimal_evidence).
EVIDENCE_EUROPEAN = 'european'    # a decimal comma, or dot grouping
EVIDENCE_ENGLISH = 'english'      # a decimal point, or comma grouping

# Non-breaking, narrow non-breaking and thin spaces, as Czech typography
# writes them before % and between digit groups.
_NUMBER_SPACES = str.maketrans({'\u00a0': ' ', '\u202f': ' ', '\u2009': ' '})
# The typographic minus sign, read as a hyphen-minus.
_MINUS_SIGN = '\u2212'

# Minimum length to even attempt date parsing (avoids matching plain numbers)
_MIN_DATE_LENGTH = 6
# Regex to quickly reject values that clearly can't be dates
_DATE_CANDIDATE_RE = re.compile(r'^\d{1,4}[\.\-/]|^\d{1,2}\s+\w|^\w+\s+\d')


def _try_parse_date(value: str) -> tuple[datetime, str] | None:
    """Attempt to parse a string as a date/datetime.

    Tries explicit formats first (fast, unambiguous), then falls back to
    dateutil for natural language dates.

    Returns (datetime_obj, excel_number_format) or None.
    """
    if len(value) < _MIN_DATE_LENGTH:
        return None
    if not _DATE_CANDIDATE_RE.match(value):
        return None

    # Try explicit formats first (deterministic, no ambiguity)
    for fmt, xl_fmt in DATE_FORMATS:
        try:
            dt = datetime.strptime(value, fmt)
            return dt, xl_fmt
        except ValueError:
            continue

    # Fallback to dateutil (handles many international/natural formats)
    try:
        dt = dateutil_parser.parse(value, dayfirst=True, fuzzy=False)
        # Only accept if the string is sufficiently "date-like" —
        # dateutil can parse things like "1" or "March" alone which we don't want
        if dt and len(value) >= 8:
            # Determine appropriate format based on whether time is present
            if dt.hour or dt.minute or dt.second:
                return dt, "YYYY-MM-DD HH:MM:SS"
            return dt, "YYYY-MM-DD"
    except (ValueError, TypeError, OverflowError):
        pass

    return None


def _normalize_grouped_number(text: str) -> str:
    """Strip digit-group separators when the grouping is unambiguous.

    Returns ``text`` unchanged when the grouping could be read more than one
    way, so the caller's ``float()`` still rejects it and the value stays text.
    """
    if _GROUPED_COMMA_RE.match(text):
        return text.replace(',', '')
    if _GROUPED_DOT_RE.match(text):
        return text.replace('.', '').replace(',', '.')
    return text


def _thousands_format_for(value: float) -> str:
    """Pick the grouped number format for a plain numeric cell.

    Whole numbers take the integer format; anything else keeps two decimals.
    Previously every value from 1000 up got ``#,##0`` unconditionally, which
    silently displayed 1500.75 as ``1,501`` — the stored value was right but
    the number a reader saw was not.
    """
    return THOUSANDS_FORMAT if float(value).is_integer() else THOUSANDS_FORMAT_DECIMALS


# The kinds _lex_decimal sorts a body into.
_KIND_PLAIN = 'plain'                       # 1234, 1 234, 1e3, 1., 5,
_KIND_DOT_DECIMAL = 'dot_decimal'           # 1.5, 0.125, 1234.567
_KIND_COMMA_DECIMAL = 'comma_decimal'       # 4,3, 0,125, ,125
_KIND_UNGROUPED_COMMA = 'ungrouped_comma'   # 1234,567
_KIND_COMMA_GROUPED = 'comma_grouped'       # 1,234.5, 1,234,567, 1,23,456
_KIND_DOT_GROUPED = 'dot_grouped'           # 1.234,56, 1.234.567
_KIND_AMBIGUOUS_COMMA = 'ambiguous_comma'   # 1,234, 4,375
_KIND_AMBIGUOUS_DOT = 'ambiguous_dot'       # 1.234

# What each kind proves about the table it is in. The neutral kinds prove
# nothing and are absent: plain integers (a trailing ``1.`` or ``5,`` writes
# no decimal digit), ``1234,567`` (a decimal either way, but an English
# writer who left out a separator writes it too) and the two ambiguous forms.
_KIND_EVIDENCE = {
    _KIND_COMMA_DECIMAL: EVIDENCE_EUROPEAN,
    _KIND_DOT_GROUPED: EVIDENCE_EUROPEAN,
    _KIND_DOT_DECIMAL: EVIDENCE_ENGLISH,
    _KIND_COMMA_GROUPED: EVIDENCE_ENGLISH,
}
# How a note names what an evidence value shows.
_KIND_SHOWS = {
    _KIND_COMMA_DECIMAL: 'writes a decimal comma',
    _KIND_DOT_GROUPED: 'groups thousands with dots',
    _KIND_DOT_DECIMAL: 'writes a decimal point',
    _KIND_COMMA_GROUPED: 'groups thousands with commas',
}


def _lex_decimal(text: str) -> tuple[str, str] | None:
    """Check *text* against the numeric grammar: ``(compact_text, kind)`` or None.

    *compact_text* has its edges trimmed, its space grouping removed and a
    typographic minus read as ``-``; *kind* is one of the ``_KIND_*`` names.
    The grammar is :func:`_parse_decimal_text`'s, which states it in prose;
    :func:`_decimal_evidence` reads the kind for what it proves.
    """
    text = text.translate(_NUMBER_SPACES).replace(_MINUS_SIGN, '-').strip()
    if ' ' in text or '\t' in text:
        if not _SPACE_GROUPED_RE.match(text):
            return None
        text = text.replace(' ', '')
    if not _NUMBER_BODY_RE.match(text):
        return None
    mantissa, _, exponent = text.replace('E', 'e').partition('e')
    commas, dots = mantissa.count(','), mantissa.count('.')
    if not commas and not dots:
        kind = _KIND_PLAIN
    elif (commas and dots) or commas > 1 or dots > 1:
        if exponent:
            return None                     # 1.234.567e3 is nobody's number
        if _GROUPED_COMMA_RE.match(mantissa) or _INDIAN_GROUPED_RE.match(mantissa):
            kind = _KIND_COMMA_GROUPED
        elif _DIGIT_GROUPED_RE.match(mantissa):
            kind = _KIND_DOT_GROUPED
        else:
            return None
    elif mantissa.endswith((',', '.')):
        kind = _KIND_PLAIN                  # 1. and 5, write no decimal digit
    elif commas:
        if _AMBIGUOUS_COMMA_RE.match(mantissa):
            kind = _KIND_AMBIGUOUS_COMMA
        elif _UNGROUPED_COMMA_RE.match(mantissa):
            kind = _KIND_UNGROUPED_COMMA
        else:
            kind = _KIND_COMMA_DECIMAL
    else:
        kind = (_KIND_AMBIGUOUS_DOT
                if not exponent and _AMBIGUOUS_DOT_RE.match(mantissa)
                else _KIND_DOT_DECIMAL)
    return text, kind


def _parse_decimal_text(text: str, *, comma3: str) -> tuple[float, str] | None:
    """Read a number written with dots, commas or spaces: ``(value, normalised)``.

    The one numeric reader for every column type that holds a number —
    ``number`` and ``currency`` in :func:`_apply_column_type`, ``percent``
    through :func:`_parse_percent_body` — so the three read the same text the
    same way (#143, #211). Callers strip what is theirs first — a currency
    symbol and accounting parentheses, a ``%`` — through
    :func:`_numeric_body`.

    Returns None for anything the grammar refuses, and the cell stays text,
    which a typed column then reports as ``typed_value_kept_as_text``. The
    normalised text is what ``float()`` read; a percent counts its decimals on
    it, so ``1.234,56`` shows two, not the six its raw text implies.

    The grammar:

    - Whitespace at the edges is dropped (``4,3 %``, a Czech non-breaking
      space). Inside, it is allowed only as digit grouping (``1 234,56``);
      ``10 20``, ``10, 20`` and ``4 3`` are refused. Non-breaking, narrow
      non-breaking and thin spaces count as spaces.
    - The typographic minus sign (U+2212) reads as ``-``.
    - An exponent is accepted only without a comma: ``1,5e3`` is refused.
    - Two separators, or several of one, must be real digit grouping — groups
      of three after a lead group that does not start with 0: ``1.234,56``,
      ``1,234.5``, ``1,234,567`` and ``1.234.567`` read; ``4,3.``, ``1.5,2``
      and ``0,125,000`` are refused. Indian lakh grouping, twos before a last
      three (``1,23,456.70``), is comma grouping too.
    - A lone separator at the end writes no decimal digit: ``1.`` and ``5,``
      are the integers 1 and 5.
    - A lone comma is a decimal comma — ``4,3``, ``,5``, ``3,14159`` — and so
      is a lone comma before three digits wherever the digits in front of it
      cannot be a thousands group: ``0,125`` and ``,125`` (nobody groups
      behind a zero), ``1234,567`` (four digits are not one group). Before
      this, ``types: number`` read ``0,125`` as 125 (#211).
    - The one truly ambiguous form is a lone comma after one to three digits
      and before exactly three, with no dot: ``1,234``, ``4,375``. *comma3*
      decides it: ``'refuse'`` returns None (a percent column, and a number
      column whose table writes numbers both ways), ``'thousands'`` reads
      4375, ``'decimal'`` reads 4.375. Under ``'decimal'`` — a table that
      writes European numbers — the mirror ``1.234`` is read as thousands;
      under the other two it is the decimal it is in English.
    - ``nan``, ``inf`` and PEP 515 underscores are refused: the first two
      serialise as an empty cell, the third would turn a product code into a
      number.
    """
    if comma3 not in _COMMA3_POLICIES:
        raise ValueError(f"unknown comma3 policy {comma3!r}")
    lexed = _lex_decimal(text)
    if lexed is None:
        return None
    compact, kind = lexed
    if kind == _KIND_AMBIGUOUS_COMMA and comma3 == COMMA3_REFUSE:
        return None
    normalised = _strip_thousands_separators(compact, comma3=comma3)
    try:
        value = float(normalised)
    except ValueError:
        return None
    if not math.isfinite(value):
        return None
    return value, normalised


def _decimal_evidence(text: str) -> tuple[str, str] | None:
    """What *text* proves about how its table writes numbers: ``(evidence, kind)``.

    *evidence* is ``EVIDENCE_EUROPEAN`` for a value that can only be a
    decimal comma or dot grouping — ``4,3``, ``0,125``, ``1.234,5``,
    ``1.234.567``, ``1 234,5`` — and ``EVIDENCE_ENGLISH`` for a decimal point
    or comma grouping — ``1.5``, ``12.50``, ``0.125``, ``1,234.5``,
    ``1,234,567``, ``1,23,456``; *kind* is the ``_KIND_*`` it was read as,
    which a note names (``_KIND_SHOWS``). None for what proves nothing —
    integers, ``1 234``, a trailing ``1.`` or ``5,``, ``1234,567``, the
    ambiguous ``1,234`` and ``1.234`` themselves — and for anything the
    grammar refuses.
    """
    lexed = _lex_decimal(text)
    if lexed is None or lexed[1] not in _KIND_EVIDENCE:
        return None
    return _KIND_EVIDENCE[lexed[1]], lexed[1]


def _ambiguous_form(text: str) -> tuple[str, str] | None:
    """``(compact, kind)`` if *text* is ``4,375`` or its mirror ``1.234``, else None.

    What the warnings about those forms quote in their hint; *kind* is
    ``_KIND_AMBIGUOUS_COMMA`` or ``_KIND_AMBIGUOUS_DOT``.
    """
    lexed = _lex_decimal(text)
    if lexed is None or lexed[1] not in (_KIND_AMBIGUOUS_COMMA, _KIND_AMBIGUOUS_DOT):
        return None
    return lexed


def _parse_percent_body(body: str) -> tuple[float, str] | None:
    """Read the number in front of a ``%`` as a fraction, or None.

    The one test of "is this a percent", shared by every place that asks it —
    the ``percent`` branch of :func:`_apply_column_type`, Step 3 of
    :func:`resolve_cell` and the format learner
    :func:`_percent_format_for_column` — so the three cannot disagree (#126,
    #143). Callers strip the ``%`` themselves; whether it is required is
    theirs to decide (a typed column takes it optional, an untyped cell needs
    it).

    Returns ``(fraction, normalised_text)``: ``4,3`` → ``(0.043, '4.3')``.
    The number is read by :func:`_parse_decimal_text` under the ``'refuse'``
    policy: ``0,125`` is 0.125 % and ``1234,567`` is 1234.567 % (#211), but
    the truly ambiguous ``4,375`` and ``1,234`` stay text. A percent takes no
    evidence from its table — the number columns do — so it does not guess.
    """
    parsed = _parse_decimal_text(body, comma3=COMMA3_REFUSE)
    if parsed is None:
        return None
    value, normalised = parsed
    return value / 100, normalised


def _percent_decimals(numeric_text: str) -> int:
    """How many decimal places the source text of a percent carries.

    Counts after the first ``.``, so give it the *normalised* text from
    :func:`_parse_percent_body`; a comma is still read as a decimal point for
    callers that pass a bare ``50,5``.
    """
    fraction = numeric_text.strip().replace(',', '.').partition('.')[2].strip()
    return min(len(fraction), _PERCENT_MAX_DECIMALS)


def _percent_format_with(decimals: int) -> str:
    """The Excel percent format showing *decimals* decimal places."""
    return f"0.{'0' * decimals}%" if decimals else '0%'


def _percent_format_for(numeric_text: str) -> str:
    """Build a percent format that preserves the precision of the source text.

    ``50%`` → ``0%``, ``50.5%`` → ``0.0%``, ``50.25%`` → ``0.00%``. The format
    was previously a flat ``0%``, which rendered 50.5% as ``51%`` — a visible
    value the source never contained.
    """
    return _percent_format_with(_percent_decimals(numeric_text))


def _percent_format_for_column(table_data: list[list[str]], col_idx: int) -> str | None:
    """The percent format a column's own literal values imply, or None.

    A formula has no source text to take its precision from, so a computed
    cell in a ``types: percent`` column used to fall back to a flat ``0%`` and
    render 4.3% as ``4%`` — while the literal values beside it each kept their
    own precision through :func:`_percent_format_for`. The loss was silent and
    only hit formula cells (#126).

    The widest precision in the column wins: showing a value with more decimal
    places than its neighbours is a cosmetic surprise, showing it with fewer
    hides part of the number. Returns None when the column has no literal
    percent to learn from, leaving the caller its own fallback.

    A cell counts exactly when :func:`_apply_column_type` would coerce it, and
    the two tests are deliberately identical: :func:`_numeric_body` then
    :func:`_parse_percent_body`, the functions both call. Either half
    drifting apart from the coercion has already caused this both ways —
    ``1,234%`` is ambiguous, so the coercion leaves it as text and it must not
    set the format of the cells that are percents; while a bare ``50.5`` or a
    European ``4,3`` in a percent column *is* coerced (the ``%`` is optional
    there), so it must (#143). Neither the sign nor a comma is a test of its
    own; what the column does with the value is. Decimals are counted on the
    normalised text, so ``1.234,56`` teaches two.
    """
    decimals = None
    for _, text in _data_cell_texts(table_data, col_idx):
        body, _ = _numeric_body(text, 'percent')         # exactly what the
        parsed = _parse_percent_body(body)               # column coercion
        if parsed is None:                               # itself accepts
            continue
        decimals = max(decimals or 0, _percent_decimals(parsed[1]))
    return None if decimals is None else _percent_format_with(decimals)


def _data_cell_texts(table_data: list[list[str]], col_idx: int):
    """Yield ``(row_idx, text)`` for each literal data cell of column *col_idx*.

    The one walk over a column's values that the learners share — the
    percent format learner above and the evidence pass,
    :func:`_comma3_policies` — so neither can count what the coercion skips:
    data rows only (the header is a name, not a value), whole-cell markdown
    stripped as the cell loop strips it (``**2,5**`` is ``2,5``), formulas
    and empty cells skipped (a formula has no source text, an empty cell no
    value), and a ragged row that stops short of the column is passed over.
    """
    for row_idx, row in enumerate(table_data[1:], start=1):
        if col_idx >= len(row):
            continue
        text = _strip_markdown_formatting(row[col_idx])[0].strip()
        if text and not text.startswith('='):
            yield row_idx, text


def _parse_column_alignments(cells: list[str]) -> list[str | None]:
    """Extract column alignments from a separator row's split *cells*.

    Returns a list of alignment strings ('left', 'center', 'right') or None per column.
    This is the same logic used by docx_tools but returns generic strings
    instead of Word-specific enums.
    """
    alignments: list[str | None] = []
    for cell in cells:
        if cell.startswith(':') and cell.endswith(':'):
            alignments.append('center')
        elif cell.endswith(':'):
            alignments.append('right')
        elif cell.startswith(':'):
            alignments.append('left')
        else:
            alignments.append(None)  # auto — will use heuristic
    return alignments


def parse_table(lines: list[str], start_idx: int) -> tuple[list[list[str]] | None, int]:
    """Parse markdown table and return (table_data, next_index).

    The table runs while lines are table rows by
    :func:`markdown_table.is_table_line` — a leading ``|``, the trailing one
    optional — and each row is split by :func:`markdown_table.split_row`; the
    Word tool's ``parse_table()`` uses the same two functions, so the two
    tools keep the same rows (#144).

    Also extracts column alignments from the separator row and attaches them
    as the 'col_alignments' attribute on the returned TableData instance.
    """
    rows: list[list[str]] = []
    i = start_idx

    # Find all consecutive table lines
    while i < len(lines):
        line = lines[i].strip()
        if not is_table_line(line):
            break
        rows.append(split_row(line))
        i += 1

    if len(rows) < 2:  # Need at least header and separator
        return None, i if i > start_idx else start_idx + 1

    # A run of nothing but separator rows is not a table at all: there is no
    # header for them to sit under. Empty table_data is what the caller of this
    # function reads as "no table here", and it reports table_incomplete.
    if all(is_separator_cells(cells) for cells in rows):
        return TableData([], [], has_separator=False), i

    # Parse table data, extracting alignment from the separator row.
    #
    # Markdown has exactly one separator, directly under the header, and a row
    # of dashes anywhere else is data — `| - | - |` is how a caller writes "not
    # applicable in either column". Matching by shape alone dropped such a row
    # from the sheet with nothing said, and widening the match from three
    # dashes to one (#114) made a lone `-` hit it, which is the likelier
    # spelling by far. So position decides: row 1 may be the separator, every
    # other row is content. One dash is enough, as in CommonMark: demanding
    # three wrote a caller's ``|--|--|`` into the sheet as a row of literal
    # dashes and then reported ``table_separator_missing`` (#114).
    table_data: list[list[str]] = []
    col_alignments: list[str | None] = []
    separator_in_place = False
    for idx, cells in enumerate(rows):
        if idx == 1 and is_separator_cells(cells):
            col_alignments = _parse_column_alignments(cells)
            separator_in_place = True
            continue
        table_data.append(cells)

    # Attach alignment info to the table_data list
    table_data_with_align = TableData(table_data, col_alignments,
                                      has_separator=separator_in_place)
    return table_data_with_align, i


class TableData(list):
    """A list subclass that carries column alignment metadata.

    *has_separator* records whether the markdown had a ``|---|---|`` row
    **directly under its first row**, which is the only place markdown gives it
    meaning. :func:`parse_table` does not require one — it only skips the rows
    that look like one, wherever they are — so a table written without it, or
    with it somewhere else, still parses, with its first row taken as the
    header. That is usually what the caller meant, but it is a decision made on
    their behalf, so the parser reports it (#114). It defaults to True: only
    markdown that was really parsed can say otherwise.
    """

    def __init__(self, data: list[list[str]], col_alignments: list[str | None] | None = None,
                 has_separator: bool = True):
        super().__init__(data)
        self.col_alignments: list[str | None] = col_alignments or []
        self.has_separator: bool = has_separator


# ── Cell Resolution ───────────────────────────────────────────────────────────

@dataclass
class CellResult:
    """Resolved cell metadata — all information needed to write a cell to Excel."""
    value: str | int | float | datetime  # The cleaned value to write
    is_formula: bool = False
    is_percent: bool = False
    percent_format: str = ""  # Excel number format for percents (e.g. "0.0%")
    is_date: bool = False
    date_format: str = ""  # Excel number format for dates (e.g. "YYYY-MM-DD")
    bold: bool = False
    italic: bool = False
    monospace: bool = False

    @property
    def formatting_info(self) -> dict[str, bool]:
        """Legacy-compatible formatting dict for apply_cell_formatting()."""
        return {'bold': self.bold, 'italic': self.italic, 'monospace': self.monospace}




def _strip_markdown_formatting(raw_text: str) -> tuple[str, dict[str, bool]]:
    """Strip inline markdown formatting markers from a cell value.

    Returns (clean_text, formatting_dict) where formatting_dict has
    'bold', 'italic', 'monospace' keys.
    """
    clean_text = raw_text.strip()
    formatting = {'bold': False, 'italic': False, 'monospace': False}

    if clean_text.startswith('**') and clean_text.endswith('**') and len(clean_text) > 4:
        clean_text = clean_text[2:-2]
        formatting['bold'] = True
    elif clean_text.startswith('*') and clean_text.endswith('*') and len(clean_text) > 2:
        clean_text = clean_text[1:-1]
        formatting['italic'] = True
    elif clean_text.startswith('`') and clean_text.endswith('`') and len(clean_text) > 2:
        clean_text = clean_text[1:-1]
        formatting['monospace'] = True

    return clean_text, formatting


def resolve_cell(raw_text: str) -> CellResult:
    """Parse a raw markdown cell string into a fully resolved CellResult.

    Combines formatting detection, formula detection, and type conversion
    in a single pass — the unified replacement for the former three-function pipeline
    of parse_cell_formatting → detect_formula_pattern → format_cell_value.
    """
    # Step 1: Strip markdown formatting markers
    clean_text, formatting = _strip_markdown_formatting(raw_text)
    bold = formatting['bold']
    italic = formatting['italic']
    monospace = formatting['monospace']

    # Step 2: Check if it's an explicit formula (= prefix)
    if clean_text.startswith('='):
        return CellResult(
            value=clean_text, is_formula=True,
            bold=bold, italic=italic, monospace=monospace,
        )

    # Step 3: Detect percent and convert to number. The body is read by the
    # same function a `types: percent` column uses, so `4,3%` and `1.234,56%`
    # parse here too, and an ambiguous `1,234%` stays text in both (#143).
    if clean_text.endswith('%'):
        parsed = _parse_percent_body(clean_text[:-1])
        if parsed is not None:
            numeric_val, normalised = parsed
            return CellResult(
                value=numeric_val, is_percent=True,
                percent_format=_percent_format_for(normalised),
                bold=bold, italic=italic, monospace=monospace,
            )
        # Not a valid percent number — fall through

    # Step 4: Try numeric conversion. Digit-group separators are stripped only
    # when the grouping is unambiguous — see _normalize_grouped_number.
    try:
        numeric_val = float(_normalize_grouped_number(clean_text))
    except ValueError:
        pass
    else:
        # float() also accepts Python literals that are not spreadsheet
        # numbers. 'nan'/'inf'/'-Infinity' have no XLSX representation —
        # openpyxl writes them as an empty <v>, so the cell silently arrives
        # BLANK — and PEP 515 underscores mean a product code like '1_000'
        # would quietly become 1000. Both belong in the text branch.
        if math.isfinite(numeric_val) and '_' not in clean_text:
            return CellResult(
                value=numeric_val,
                bold=bold, italic=italic, monospace=monospace,
            )

    # Step 5: Try date detection (after numeric, so "2024" isn't parsed as a date)
    date_result = _try_parse_date(clean_text)
    if date_result:
        dt, xl_fmt = date_result
        return CellResult(
            value=dt, is_date=True, date_format=xl_fmt,
            bold=bold, italic=italic, monospace=monospace,
        )

    # Step 6: Plain text
    return CellResult(
        value=clean_text,
        bold=bold, italic=italic, monospace=monospace,
    )


def apply_cell_formatting(cell, formatting_info: dict[str, bool]) -> None:
    """Apply inline markdown formatting (bold/italic/code) to an Excel cell.

    openpyxl fonts are immutable, so each branch builds a replacement rather
    than mutating. The current font's family is carried over — dropping it
    silently reset the cell to the workbook default, which is invisible today
    (everything is Calibri) but would quietly undo any per-cell or
    workbook-wide font choice layered underneath.
    """
    current_font = cell.font
    if formatting_info['bold']:
        cell.font = Font(name=current_font.name, bold=True,
                         color=current_font.color, size=current_font.size)
    elif formatting_info['italic']:
        cell.font = Font(name=current_font.name, italic=True,
                         color=current_font.color, size=current_font.size)
    elif formatting_info['monospace']:
        cell.font = Font(name='Courier New',
                         color=current_font.color, size=current_font.size)


# ── Formula Reference Resolution ─────────────────────────────────────────────

def _quote_sheet_name(name: str) -> str:
    """Return the sheet name quoted for Excel if it contains spaces or special chars.

    An apostrophe inside the name is doubled, which is how Excel escapes it
    within a quoted sheet reference. Without that, a sheet called ``John's
    Data`` produced ``'John's Data'!B2`` — the quoted section ends at the
    apostrophe and the rest is garbage, so the whole formula is invalid.
    """
    if re.search(r"[^A-Za-z0-9_]", name):
        return "'{}'".format(name.replace("'", "''"))
    return name


def _resolve_row(positions: dict[str, int], table_num: int, offset: int, fallback_row: int,
                 warnings=None, location: dict | None = None) -> int:
    """Resolve a table-relative row reference to an absolute Excel row number.

    Args:
        positions: Table positions dict ({"T1": start_row, ...}) for the target sheet.
        table_num: Table number (1-based).
        offset: Row offset within the table (0 = first data row).
        fallback_row: Row to use if the table isn't found in positions.
        warnings: The build's :class:`~warning_channel.WarningChannel`.
        location: Where to report it — ``{"sheet": …, "cell": …}``.

    Returns:
        The absolute Excel row number.

    A missing table key (e.g. ``T9`` when only 3 tables exist) is logged at
    WARNING level and reported on *warnings*. The formula still resolves —
    using ``fallback_row`` — so the file ships, but the reference almost
    certainly points at the wrong cell, which is otherwise a silent failure the
    caller has no way to notice.
    """
    key = f"T{table_num}"
    base = positions.get(key)
    if base is not None:
        return base + 1 + offset  # +1 to skip header row
    known = ", ".join(sorted(positions.keys())) or "none"
    logger.warning(
        "Formula references %s but no such table exists in the target sheet "
        "(known tables: %s); falling back to the current row. This likely "
        "produces a wrong cell reference — check the table numbering.",
        key, known,
    )
    if warnings is not None:
        warnings.add(
            W.TABLE_REFERENCE_MISSING,
            f"the formula references {key}, but the target sheet has no such "
            f"table (it has: {known}); the reference fell back to the current "
            f"row and almost certainly points at the wrong cell.",
            **(location or {}),
        )
    return fallback_row + offset


def _warn_unknown_sheet(sheet: str, all_sheet_table_positions: dict[str, dict[str, int]],
                        warnings=None, location: dict | None = None) -> None:
    """Report a cross-sheet reference that names a sheet which doesn't exist.

    The formula still resolves (the regex emits a syntactically valid
    cross-sheet reference), but Excel will show ``#REF!`` on open. Surfacing
    the typo during generation — e.g. ``Revenue!T1.B[0]`` when the sheet is
    actually named ``Revenue Model`` — beats letting it fail silently in the
    client.
    """
    if sheet not in all_sheet_table_positions:
        known = ", ".join(sorted(all_sheet_table_positions.keys())) or "none"
        logger.warning(
            "Formula references sheet '%s' which does not exist in the workbook "
            "(known sheets: %s). The generated reference will likely resolve to "
            "#REF! in Excel.",
            sheet, known,
        )
        if warnings is not None:
            warnings.add(
                W.SHEET_REFERENCE_MISSING,
                f"the formula references a sheet named '{sheet}', which the "
                f"workbook does not have (it has: {known}); Excel will show "
                f"#REF! there.",
                **(location or {}),
            )


# A sheet name inside a cross-sheet reference: either the quoted Excel form or
# a bare name. Quoting is what makes a name containing an operator character
# unambiguous — `=A1-P&L!T1.B[0]` cannot be parsed, `=A1-'P&L'!T1.B[0]` can —
# and it is the form Excel itself writes. Doubled apostrophes inside a quoted
# name are an escaped apostrophe.
#
# Deliberately a single capturing group: the three cross-sheet patterns below
# index their groups positionally, so an alternation with two groups would
# shift every later index.
_SHEET_NAME_PATTERN = r"((?:'[^']*(?:''[^']*)*')|[\w\s.]+)"


def _unquote_sheet_name(raw: str) -> str:
    """Normalise a matched sheet name, stripping Excel quoting if present."""
    raw = raw.strip()
    if len(raw) >= 2 and raw.startswith("'") and raw.endswith("'"):
        return raw[1:-1].replace("''", "'")
    return raw


def _make_cell_ref(column: str, row: int, sheet: str | None = None) -> str:
    """Build a cell reference string, optionally with a quoted sheet prefix."""
    if sheet:
        return f"{_quote_sheet_name(sheet)}!{column}{row}"
    return f"{column}{row}"


def adjust_formula_references(
    formula: str,
    current_excel_row: int,
    table_positions: dict[str, int] | None = None,
    all_sheet_table_positions: dict[str, dict[str, int]] | None = None,
    warnings=None,
    location: dict | None = None,
) -> str:
    """Convert row-relative references [offset] and table references T1.B[1] to actual Excel row numbers.

    Also resolves cross-sheet references like ``SheetName!T1.B[0]`` → ``'SheetName'!B2``.

    *warnings* is the build's :class:`~warning_channel.WarningChannel` and
    *location* the ``{"sheet": …, "cell": …}`` of the formula's own cell: a
    reference that resolves to the wrong place is written into a workbook that
    opens fine, so it is only ever caught if it is reported.
    """
    if not formula.startswith('='):
        return formula

    if table_positions is None:
        table_positions = {}
    if all_sheet_table_positions is None:
        all_sheet_table_positions = {}

    logger.debug("Resolving formula: %s (current_row=%d)", formula, current_excel_row)

    def _row(positions: dict[str, int], table_num: int, offset: int) -> int:
        """:func:`_resolve_row` with this formula's fallback row and location."""
        return _resolve_row(positions, table_num, offset, current_excel_row,
                            warnings=warnings, location=location)

    def _check_sheet(sheet: str) -> None:
        _warn_unknown_sheet(sheet, all_sheet_table_positions,
                            warnings=warnings, location=location)

    try:
        # ── Cross-sheet references (must be resolved BEFORE local patterns) ──

        # Cross-sheet function: SheetName!T1.SUM(B[0]:E[0])
        cs_func_pattern = _SHEET_NAME_PATTERN + r"!T(\d+)\.(SUM|AVERAGE|MAX|MIN)\(([A-Z]+)\[([+-]?\d+)\]:([A-Z]+)\[([+-]?\d+)\]\)"

        def _replace_cs_func(match):
            sheet = _unquote_sheet_name(match.group(1))
            table_num = int(match.group(2))
            func_name = match.group(3)
            start_col = match.group(4)
            start_offset = int(match.group(5))
            end_col = match.group(6)
            end_offset = int(match.group(7))
            _check_sheet(sheet)
            pos = all_sheet_table_positions.get(sheet, {})
            sr = _row(pos, table_num, start_offset)
            er = _row(pos, table_num, end_offset)
            qs = _quote_sheet_name(sheet)
            # The sheet prefix belongs on the FIRST endpoint only:
            # =SUM(Data!B2:B4). This is the canonical form Excel itself
            # writes; repeating the prefix on both endpoints is redundant
            # and needlessly divergent, and the quoted variant
            # ('My Sheet'!B2:'My Sheet'!B4) is a shape third-party parsers
            # are much less likely to have been tested against.
            result = f"{func_name}({qs}!{start_col}{sr}:{end_col}{er})"
            logger.debug("  Cross-sheet func: %s → %s", match.group(0), result)
            return result

        formula = re.sub(cs_func_pattern, _replace_cs_func, formula)

        # Cross-sheet range: SheetName!T1.B[0]:T1.E[0]
        cs_range_pattern = _SHEET_NAME_PATTERN + r"!T(\d+)\.([A-Z]+)\[([+-]?\d+)\]:T(\d+)\.([A-Z]+)\[([+-]?\d+)\]"

        def _replace_cs_range(match):
            sheet = _unquote_sheet_name(match.group(1))
            st_num = int(match.group(2))
            start_col = match.group(3)
            start_offset = int(match.group(4))
            et_num = int(match.group(5))
            end_col = match.group(6)
            end_offset = int(match.group(7))
            _check_sheet(sheet)
            pos = all_sheet_table_positions.get(sheet, {})
            sr = _row(pos, st_num, start_offset)
            er = _row(pos, et_num, end_offset)
            qs = _quote_sheet_name(sheet)
            result = f"{qs}!{start_col}{sr}:{end_col}{er}"
            logger.debug("  Cross-sheet range: %s → %s", match.group(0), result)
            return result

        formula = re.sub(cs_range_pattern, _replace_cs_range, formula)

        # Cross-sheet single cell: SheetName!T1.B[0]
        cs_cell_pattern = _SHEET_NAME_PATTERN + r"!T(\d+)\.([A-Z]+)\[([+-]?\d+)\]"

        def _replace_cs_cell(match):
            sheet = _unquote_sheet_name(match.group(1))
            table_num = int(match.group(2))
            column = match.group(3)
            offset = int(match.group(4))
            _check_sheet(sheet)
            pos = all_sheet_table_positions.get(sheet, {})
            actual_row = _row(pos, table_num, offset)
            result = _make_cell_ref(column, actual_row, sheet)
            logger.debug("  Cross-sheet cell: %s → %s", match.group(0), result)
            return result

        formula = re.sub(cs_cell_pattern, _replace_cs_cell, formula)

        # ── Local (same-sheet) references ──
        # NOTE: Range and function patterns must be processed BEFORE single-cell
        # to prevent the single-cell regex from consuming parts of range expressions.

        # Table range references e.g. T1.B[0]:T1.E[0]
        table_range_pattern = r'T(\d+)\.([A-Z]+)\[([+-]?\d+)\]:T(\d+)\.([A-Z]+)\[([+-]?\d+)\]'

        def replace_table_range(match):
            start_table_num = int(match.group(1))
            start_col = match.group(2)
            start_offset = int(match.group(3))
            end_table_num = int(match.group(4))
            end_col = match.group(5)
            end_offset = int(match.group(6))
            start_row = _row(table_positions, start_table_num, start_offset)
            end_row = _row(table_positions, end_table_num, end_offset)
            return f"{start_col}{start_row}:{end_col}{end_row}"

        adjusted = re.sub(table_range_pattern, replace_table_range, formula)

        # Simplified function over table range e.g. T1.SUM(B[0]:E[0])
        table_func_pattern = r'T(\d+)\.(SUM|AVERAGE|MAX|MIN)\(([A-Z]+)\[([+-]?\d+)\]:([A-Z]+)\[([+-]?\d+)\]\)'

        def replace_table_function(match):
            table_num = int(match.group(1))
            func_name = match.group(2)
            start_col = match.group(3)
            start_offset = int(match.group(4))
            end_col = match.group(5)
            end_offset = int(match.group(6))
            start_row = _row(table_positions, table_num, start_offset)
            end_row = _row(table_positions, table_num, end_offset)
            return f"{func_name}({start_col}{start_row}:{end_col}{end_row})"

        adjusted = re.sub(table_func_pattern, replace_table_function, adjusted)

        # Table cell references e.g. T1.B[1] (AFTER range patterns)
        table_pattern = r'T(\d+)\.([A-Z]+)\[([+-]?\d+)\]'

        def replace_table_reference(match):
            table_num = int(match.group(1))
            column = match.group(2)
            offset = int(match.group(3))
            actual_row = _row(table_positions, table_num, offset)
            result = f"{column}{actual_row}"
            logger.debug("  Local table ref: %s → %s", match.group(0), result)
            return result

        adjusted = re.sub(table_pattern, replace_table_reference, adjusted)

        # ── Row-relative references: B[0], A[-1], B[0]:E[0] ──
        #
        # These are offsets from the row the formula LIVES ON: in a formula on
        # row 7, B[0] is B7, B[-1] is B6, B[1] is B8. That is what the tool
        # description documents ("B[0] for current row references") and it is
        # the only reading under which the syntax means anything: resolving
        # B[n] against the table's first data row — as this did previously —
        # makes it an exact duplicate of the T1.B[n] form handled above.
        #
        # The practical consequence of the old behaviour was that every data
        # row of a computed column produced the SAME formula: `=B[0]*C[0]`
        # written down a table resolved to `=B4*C4` on every row, so running
        # totals, per-row products and growth rates were all silently wrong.
        # Point at a fixed cell with T1.B[n]; point at the current row with B[n].

        # Row-relative range e.g. B[0]:E[0] (BEFORE single-cell relative)
        range_pattern = r'([A-Z]+)\[([+-]?\d+)\]:([A-Z]+)\[([+-]?\d+)\]'

        def replace_range(match):
            start_col = match.group(1)
            start_offset = int(match.group(2))
            end_col = match.group(3)
            end_offset = int(match.group(4))
            start_row = current_excel_row + start_offset
            end_row = current_excel_row + end_offset
            return f"{start_col}{start_row}:{end_col}{end_row}"

        adjusted = re.sub(range_pattern, replace_range, adjusted)

        # Handle row-relative references e.g. B[0] (AFTER range pattern)
        rel_pattern = r'([A-Z]+)\[([+-]?\d+)\]'

        def replace_rel(match):
            column = match.group(1)
            offset = int(match.group(2))
            result = f"{column}{current_excel_row + offset}"
            logger.debug("  Relative ref: %s → %s", match.group(0), result)
            return result

        adjusted = re.sub(rel_pattern, replace_rel, adjusted)

        logger.debug("  Resolved formula: %s → %s", formula, adjusted)
        return adjusted

    except Exception as e:
        logger.warning("Failed to adjust formula references for '%s': %s", formula, e)
        if warnings is not None:
            warnings.add(
                W.FORMULA_UNRESOLVED,
                f"the references in '{formula}' could not be resolved ({e}); "
                f"the formula was written unchanged and Excel will not "
                f"understand it.",
                **(location or {}),
            )
        return formula


# ── Directive Helpers ──────────────────────────────────────────────────────────

# Currency symbols → Excel format string
_CURRENCY_FORMATS = {
    '$': '$#,##0.00',
    '€': '#,##0.00 €',
    '£': '£#,##0.00',
    '¥': '¥#,##0',
    'Kč': '#,##0.00 "Kč"',
    'zł': '#,##0.00 "zł"',
    'kr': '#,##0.00 "kr"',
    'CHF': '"CHF" #,##0.00',
    'R$': '"R$" #,##0.00',
    '₹': '₹#,##0.00',
}


# Type keywords that legitimately start a new column spec. Used by
# _parse_types_directive to tell a real column boundary from a comma that
# lives inside an Excel number format (e.g. the ',' in number:#,##0).
_KNOWN_TYPE_KEYWORDS = frozenset(
    {"text", "bool", "currency", "number", "date", "percent"}
)


def _parse_types_directive(value: str) -> list[str | None]:
    """Parse a types directive value like 'text, currency:$, date, bool, number'.

    Returns a list of type specs (or None for unspecified columns).

    Commas separate columns, but Excel number formats themselves contain
    commas (``#,##0``), so a naive ``split(',')`` shreds a literal format like
    ``number:#,##0.00`` into ``number:#`` + ``##0.00`` and shifts every later
    column by one — silent data corruption with no error. We split on commas
    but re-join any fragment that does NOT start a new column spec back onto
    the previous one. A new column spec is either empty (unspecified column)
    or begins with a known type keyword; anything else is a continuation of
    the preceding fragment's format string.
    """
    if not value:
        return []
    specs: list[str | None] = []
    for frag in value.split(','):
        stripped = frag.strip()
        token = stripped.split(':', 1)[0].strip().lower()
        is_new_spec = (stripped == "") or (token in _KNOWN_TYPE_KEYWORDS)
        if is_new_spec or not specs:
            specs.append(stripped or None)
        else:
            # Continuation of a literal format that contained a comma —
            # re-join with the comma that split() consumed.
            prev = specs[-1] or ""
            specs[-1] = f"{prev},{frag}".strip() or None
    return specs


def _strip_thousands_separators(text: str, *, comma3: str = COMMA3_THOUSANDS) -> str:
    """Normalise a numeric string that may carry thousands/decimal separators.

    The normaliser, not the validator: :func:`_parse_decimal_text` checks the
    grammar first and hands this only text it accepts, so a malformed
    ``4,3.`` or ``1,5e3`` never gets here to be misread. Returns a string for
    ``float()``.

    Handles English (``1,234.56``), European (``1.234,56``, ``1.234.567``) and
    a lone comma. Both separators present: the last one is the decimal
    separator. Several commas: thousands. A lone comma is a decimal comma
    (``1,5``, ``0,125``, ``1234,567``) except the truly ambiguous ``1,234``,
    which *comma3* decides — thousands unless it is ``'decimal'``; a caller
    that must not guess refuses it before calling. Under ``'decimal'`` the
    mirror ``1.234`` is thousands too (#211).
    """
    text = text.strip()
    if ',' in text and '.' in text:
        # Both separators present — the LAST one is the decimal separator.
        if text.rfind(',') > text.rfind('.'):
            return text.replace('.', '').replace(',', '.')  # European 1.234,56
        return text.replace(',', '')                        # English  1,234.56
    if text.count(',') > 1:
        return text.replace(',', '')                        # 1,234,567
    if ',' in text:
        if _AMBIGUOUS_COMMA_RE.match(text) and comma3 != COMMA3_DECIMAL:
            return text.replace(',', '')                    # thousands: 1,234
        return text.replace(',', '.')                       # decimal: 1,5  0,125
    if text.count('.') > 1:
        return text.replace('.', '')                        # 1.234.567
    if comma3 == COMMA3_DECIMAL and _AMBIGUOUS_DOT_RE.match(text):
        return text.replace('.', '')                        # European 1.234
    return text


def _currency_symbol(type_spec: str) -> str:
    """The symbol a ``currency`` spec names; ``$`` when it names none."""
    symbol = type_spec.split(':', 1)[1].strip() if ':' in type_spec else ''
    return symbol or '$'


def _is_numeric_type(type_spec: str | None) -> bool:
    """Whether *type_spec* is a ``number``, ``currency`` or ``percent`` column."""
    return bool(type_spec) and type_spec.strip().lower().startswith(
        ('number', 'currency', 'percent'))


def _numeric_body(text: str, type_spec: str) -> tuple[str, bool]:
    """The number a cell of a numeric column holds: ``(body, negative)``.

    The one extraction step in front of :func:`_parse_decimal_text`, shared by
    the coercion in :func:`_apply_column_type` and by the learners that must
    count exactly what it reads — :func:`_percent_format_for_column` and the
    evidence pass :func:`_comma3_policies` (#211). *text* is the cell with its
    whole-cell markdown already stripped.

    - ``percent``: the ``%`` is optional; the rest is left to the grammar,
      which keeps the #143 whitespace rule (``- 4,3`` stays text).
    - ``currency``: the symbol is removed, then accounting parentheses —
      ``($1,234)`` is -1234, flagged by *negative* — each with the whitespace
      around it.
    - ``number`` and ``currency``: a sign written apart from its number is
      closed up (``- 5``, ``€ - 5``, ``- $1,234.50``), as these columns always
      read it. Any other inner whitespace is left for the grammar, which takes
      it only as digit grouping: ``1 234,5`` reads, ``10 20`` does not.
    """
    type_lower = type_spec.strip().lower()
    text = text.strip()
    if type_lower.startswith('percent'):
        return text.rstrip('%'), False
    negative = False
    if type_lower.startswith('currency'):
        text = text.replace(_currency_symbol(type_spec), '').strip()
        # Accounting-style negatives wrap the amount in parentheses:
        # ($1,234) means -1234. float('(1234)') raises, so without this the
        # cell stayed text.
        if text.startswith('(') and text.endswith(')'):
            negative = True
            text = text[1:-1].strip()
    text = text.translate(_NUMBER_SPACES).replace(_MINUS_SIGN, '-').strip()
    return _SIGN_SPACE_RE.sub(r'\1', text), negative


def _apply_column_type(cell, raw_text: str, type_spec: str | None, *,
                       comma3: str = COMMA3_THOUSANDS) -> bool:
    """Apply column type coercion to a cell based on directive.

    Returns True if type was applied (caller should skip default processing),
    False if default processing should continue.

    ``number``, ``currency`` and ``percent`` read their value with
    :func:`_parse_decimal_text`, after :func:`_numeric_body` has taken off
    the symbol, the parentheses or the ``%``. *comma3* is how a ``number`` or
    ``currency`` column reads the truly ambiguous ``4,375``: the table loop
    passes what the table's own evidence implies (:func:`_comma3_policies`,
    #211); the default, ``'thousands'``, is the reading with no evidence. A
    ``percent`` column always refuses that form (:func:`_parse_percent_body`).

    A value the declared type cannot read is kept as text — ``nan`` and
    ``inf`` too, which no longer become an empty cell (#211); ask
    :func:`_kept_as_text` afterwards whether that happened, so the caller can
    report it (#143). The function takes no warnings channel on purpose: the
    report is one per column, not one per cell, and only the table loop sees
    the whole column.
    """
    if not type_spec:
        return False

    clean = raw_text.strip()
    type_lower = type_spec.lower()

    # text — force string, no conversion
    if type_lower == 'text':
        cell.value = clean
        return True

    # bool — map common boolean strings to Excel boolean
    if type_lower == 'bool':
        lower_val = clean.lower()
        if lower_val in ('true', 'yes', '1', 'on'):
            cell.value = True
        elif lower_val in ('false', 'no', '0', 'off'):
            cell.value = False
        else:
            cell.value = clean  # Unrecognized → keep as text
        return True

    # currency:<symbol> — strip the symbol and any accounting parentheses,
    # then read the number as a number column does.
    if type_lower.startswith('currency'):
        symbol = _currency_symbol(type_spec)
        body, negative = _numeric_body(clean, type_spec)
        parsed = _parse_decimal_text(body, comma3=comma3)
        if parsed is None:
            cell.value = clean  # Can't parse → keep as text
        else:
            cell.value = -abs(parsed[0]) if negative else parsed[0]
            cell.number_format = _CURRENCY_FORMATS.get(symbol, f'#,##0.00 "{symbol}"')
        return True

    # number or number:<format> — parse as number with optional format.
    # Was `clean.replace(',', '')`, which turned the European decimal '1,5'
    # into 15 — a silent 10x error — and then read `0,125` as 125 (#211).
    if type_lower.startswith('number'):
        fmt = type_spec.split(':', 1)[1].strip() if ':' in type_spec else None
        body, _ = _numeric_body(clean, type_spec)
        parsed = _parse_decimal_text(body, comma3=comma3)
        if parsed is None:
            cell.value = clean
        else:
            cell.value = parsed[0]
            if fmt:
                cell.number_format = fmt
            elif abs(cell.value) >= _THOUSANDS_THRESHOLD:
                cell.number_format = _thousands_format_for(cell.value)
        return True

    # date or date:<format> — parse with dateutil, apply format
    if type_lower.startswith('date'):
        fmt = type_spec.split(':', 1)[1].strip() if ':' in type_spec else None
        result = _try_parse_date(clean)
        if result:
            dt, default_fmt = result
            cell.value = dt
            cell.number_format = fmt or default_fmt
        else:
            cell.value = clean
        return True

    # percent or percent:<format> — parse as percent. A declared format
    # applies to every cell in the column, literal and formula alike; that is
    # what declaring it is for. Without one, each literal keeps the precision
    # of its own source text.
    #
    # European decimals read as they do in a `number` column (`4,3` → 4.3%,
    # `0,125` → 0.125%), except the truly ambiguous `4,375`, which stays
    # text — see _parse_percent_body (#143, #211).
    if type_lower.startswith('percent'):
        explicit = type_spec.split(':', 1)[1].strip() if ':' in type_spec else ''
        parsed = _parse_percent_body(_numeric_body(clean, type_spec)[0])
        if parsed is None:
            cell.value = clean
        else:
            cell.value, normalised = parsed
            cell.number_format = explicit or _percent_format_for(normalised)
        return True

    return False


def _kept_as_text(cell, type_spec: str) -> bool:
    """Whether :func:`_apply_column_type` left *cell* as text against its type.

    Every type but ``text`` turns a value it can read into a non-string —
    a float, a bool or a datetime — so a non-empty string after coercion is
    exactly a value the column could not read. An empty cell is not one: it
    has no value to lose.
    """
    return (type_spec.strip().lower() != 'text'
            and isinstance(cell.value, str) and cell.value != '')


# How much of a value a typed_value_kept_as_text message quotes.
_KEPT_VALUE_EXCERPT = 60


def _report_kept_as_text(
    kept: dict[int, list],
    table_data: list[list[str]],
    col_types: list[str | None],
    worksheet,
    warnings,
) -> None:
    """Report each typed column that kept values as text — once per column.

    *kept* maps a 0-based column index to ``[count, first_coordinate,
    first_value, first_ambiguous_or_None]``; the last one — the compacted
    ``4,375`` of the first kept value that is the truly ambiguous form — is
    what the hint quotes, so a column whose first kept value is ``-`` still
    gets it. In a percent column that form is always refused; in a number or
    currency column only when the table writes numbers both ways (#211). One
    warning per table column, with the count, rather than one per cell: a
    column of ``-`` placeholders would otherwise fill the channel up to its
    cap and push out the formula errors after it (#143).
    """
    for col_idx, (count, first_cell, first_value, ambiguous) in sorted(kept.items()):
        type_spec = col_types[col_idx]
        column = _column_label(table_data, col_idx)
        quoted = _excerpt(first_value)
        if count == 1:
            what = (f"'{quoted}' in {column} does not read as {type_spec}, so "
                    f"it was kept as text")
        else:
            what = (f"{count} values in {column} do not read as {type_spec}, "
                    f"so they were kept as text — the first is '{quoted}' at "
                    f"{first_cell}")
        hint = ''
        if ambiguous is not None:
            decimal, whole = ambiguous.replace(',', '.'), ambiguous.replace(',', '')
            if type_spec.strip().lower().startswith('percent'):
                why = f"ambiguous in a percent ({ambiguous} is {decimal} or {whole})"
            else:
                why = (f"ambiguous ({ambiguous} is {decimal} or {whole}), and "
                       f"this table writes its other numbers both ways — with a "
                       f"decimal comma and with a decimal point or comma "
                       f"grouping — so neither reading was chosen")
            hint = (f" A comma followed by exactly three digits is {why}: write "
                    f"the decimal with a dot ({decimal}), or drop the comma if "
                    f"it groups thousands.")
        logger.warning(
            "Sheet '%s' %s: %d value(s) kept as text against type '%s' "
            "(first %s: %r)", worksheet.title, column, count, type_spec,
            first_cell, first_value,
        )
        if warnings is not None:
            warnings.add(
                W.TYPED_VALUE_KEPT_AS_TEXT,
                f"{what}; text will not sum, sort or format with the rest of "
                f"the column.{hint} Correct the value, change the column's "
                f"type, or leave the cell empty if it has no value.",
                sheet=worksheet.title,
                cell=first_cell,
            )


def _report_ambiguous_read_as_thousands(
    read: dict[int, list],
    table_data: list[list[str]],
    worksheet,
    warnings,
) -> None:
    """Report each column that read ``4,375`` as 4375 on no evidence — once.

    *read* maps a 0-based column index to ``[count, first_coordinate,
    first_value, first_compact, first_number]``. A ``number`` or ``currency``
    column whose table shows neither a decimal comma nor a decimal point
    (:func:`_comma3_policies`) reads the truly ambiguous form as thousands,
    as it always did; that is right for ``1,234`` in English and a 1000x
    error for a Czech ``4,375``, so the caller is told, with what to write
    instead (#211). Severity ``info``: the value is in the file, as the
    commoner reading. Aggregated per table column like
    ``typed_value_kept_as_text``, so a column of them is one entry. A value
    the evidence of another column settled is reported by
    :func:`_report_ambiguous_read_from_table` instead, never by both.
    """
    for col_idx, (count, first_cell, first_value, compact, number) in sorted(read.items()):
        column = _column_label(table_data, col_idx)
        quoted = _excerpt(first_value)
        decimal, whole = compact.replace(',', '.'), compact.replace(',', '')
        shown = f"{number:.15g}"
        if count == 1:
            what = f"'{quoted}' in {column} was read as {shown}"
            comma = "a comma before exactly three digits"
            which = "this one uses"
            fix = (f"If it is a decimal, write {decimal}; if it is the whole "
                   f"number, write {whole}")
        else:
            what = (f"{count} values in {column} with a comma before exactly "
                    f"three digits were read as thousands — the first is "
                    f"'{quoted}' at {first_cell}, read as {shown}")
            comma = "such a comma"
            which = "these use"
            fix = (f"If they are decimals, write them with a dot ({decimal}); "
                   f"if they are whole numbers, without the comma ({whole})")
        logger.info(
            "Sheet '%s' %s: %d value(s) with a comma before three digits read "
            "as thousands, with no evidence in the table (first %s: %r)",
            worksheet.title, column, count, first_cell, first_value,
        )
        if warnings is not None:
            warnings.add(
                W.AMBIGUOUS_COMMA_READ_AS_THOUSANDS,
                f"{what}: {comma} groups thousands in English but is a "
                f"decimal comma in most of Europe, and no "
                f"other number in the table shows which {which}. {fix}, and "
                f"the reading no longer depends on a guess.",
                sheet=worksheet.title,
                cell=first_cell,
            )


def _report_ambiguous_read_from_table(
    read: dict[int, list],
    table_data: list[list[str]],
    start_row: int,
    worksheet,
    warnings,
) -> None:
    """Report each column whose ambiguous values another column decided — once.

    *read* maps a 0-based column index to ``[count, first_coordinate,
    first_value, first_compact, first_kind, first_number, source]``, where
    *source* is the :class:`_Evidence` that decided the column's policy
    (:func:`_comma3_policies`, steps 2 and 3). The column itself showed
    nothing, so ``4,375`` — or, in a table read as European, ``1.234`` — was
    read by how *other* columns write numbers; that is right for a table in
    one locale and a 1000x error for a column that writes the other one, so
    the caller is told which value decided, and what to write to settle it
    whatever the rest of the table says (#211 review). Severity ``info``,
    aggregated per table column like ``ambiguous_comma_read_as_thousands``,
    which a cell never also gets.

    The explicit forms: the whole number without its separator (``4375``);
    the decimal written with a dot when the table reads English (``4.375``),
    and otherwise with a fourth decimal place in its own separator
    (``4,3750``, ``1.2340``) — in a table read as European, ``4.375`` is
    itself read as 4375.
    """
    for col_idx, (count, first_cell, first_value, compact, kind, number,
                  source) in sorted(read.items()):
        column = _column_label(table_data, col_idx)
        quoted = _excerpt(first_value)
        shown = f"{number:.15g}"
        is_comma = kind == _KIND_AMBIGUOUS_COMMA
        separator = 'comma' if is_comma else 'dot'
        whole = compact.replace(',' if is_comma else '.', '')
        # Read as English, only 4,375 is reported (1.234 reads as it would
        # with no evidence), and 4.375 is its decimal there.
        english = _KIND_EVIDENCE[source.kind] == EVIDENCE_ENGLISH
        decimal = compact.replace(',', '.') if english else compact + '0'
        if is_comma:
            why = ("a comma before exactly three digits groups thousands in "
                   "English but is a decimal comma in most of Europe")
        else:
            why = ("a dot before exactly three digits is a decimal point in "
                   "English but groups thousands in most of Europe")
        evidence_cell = f"{get_column_letter(source.col_idx + 1)}{start_row + source.row_idx}"
        evidence = (f"'{_excerpt(source.text)}' at {evidence_cell} "
                    f"{_KIND_SHOWS[source.kind]}")
        if count == 1:
            what = f"'{quoted}' in {column} was read as {shown}"
            whose = "its own column"
            if english:
                fix = (f"If it is a decimal, write {decimal}; if it is the "
                       f"whole number, write {whole}")
            else:
                fix = (f"If it is the whole number, write {whole}; if it is a "
                       f"decimal, write {decimal}")
        else:
            what = (f"{count} values in {column} were read by how other "
                    f"columns write numbers — the first is '{quoted}' at "
                    f"{first_cell}, read as {shown}")
            whose = "their own column"
            if english:
                fix = (f"If they are decimals, write them with a dot "
                       f"({decimal}); if they are whole numbers, without the "
                       f"{separator} ({whole})")
            else:
                fix = (f"If they are whole numbers, write them without the "
                       f"{separator} ({whole}); if they are decimals, with a "
                       f"fourth decimal place ({decimal})")
        logger.info(
            "Sheet '%s' %s: %d ambiguous value(s) read by the evidence of %s "
            "(first %s: %r)", worksheet.title, column, count, evidence_cell,
            first_cell, first_value,
        )
        if warnings is not None:
            warnings.add(
                W.AMBIGUOUS_NUMBER_READ_FROM_TABLE,
                f"{what}: {why}; nothing in {whose} shows which, and "
                f"{evidence}. {fix}, and the reading no longer depends on the "
                f"other columns.",
                sheet=worksheet.title,
                cell=first_cell,
            )


def _column_label(table_data: list[list[str]], col_idx: int) -> str:
    """``column B ('Amount')`` — how a column warning names its column."""
    letter = get_column_letter(col_idx + 1)
    header_row = table_data[0] if table_data else []
    header = (_strip_markdown_formatting(header_row[col_idx])[0]
              if col_idx < len(header_row) else '')
    return f"column {letter} ('{header}')" if header else f"column {letter}"


def _excerpt(value: str) -> str:
    """*value* cut to what a column warning quotes."""
    return (value if len(value) <= _KEPT_VALUE_EXCERPT
            else value[:_KEPT_VALUE_EXCERPT] + '…')


@dataclass(frozen=True)
class _Evidence:
    """One value that shows how a table writes numbers, and where it is."""
    row_idx: int          # in table_data; the header is row 0
    col_idx: int
    text: str             # the cell as written, markdown stripped
    kind: str             # the _KIND_* it was read as — what _KIND_SHOWS names


@dataclass(frozen=True)
class _Comma3Policy:
    """How one ``number`` or ``currency`` column reads ``4,375`` (and ``1.234``).

    *comma3* goes to :func:`_apply_column_type`. *guessed*: no evidence
    anywhere, reported as ``ambiguous_comma_read_as_thousands``. *source*:
    the first value that decided it when it is not in the column itself,
    reported as ``ambiguous_number_read_from_table``; None when the column
    decided for itself, or nothing did.
    """
    comma3: str = COMMA3_THOUSANDS
    guessed: bool = False
    source: _Evidence | None = None


def _untyped_number_text(raw_text: str) -> str | None:
    """The number an untyped cell holds, if the cell loop reads it as one.

    An untyped cell is evidence only when :func:`resolve_cell` itself makes
    it a number or a percent — the one test the cell loop applies — and then
    as the text that number was read from (its ``%`` taken off). A ``1,2`` or
    ``12,15`` the cell loop keeps as text shows nothing about how the table
    writes numbers: before the #211 review, a footnote marker beside an
    English dollar column turned ``$12,500`` into 12.5. None for text,
    dates, formulas and empty cells.
    """
    resolved = resolve_cell(raw_text)
    if resolved.is_formula or resolved.is_date or not isinstance(resolved.value, float):
        return None
    text = _strip_markdown_formatting(raw_text)[0]
    return text[:-1] if resolved.is_percent else text


def _evidence_values(table_data: list[list[str]], col_idx: int, spec: str | None):
    """Yield ``(row_idx, text, body)`` for each value of a column that may be evidence.

    A typed column's values are walked by :func:`_data_cell_texts` and
    extracted by :func:`_numeric_body`, exactly as its coercion reads them;
    an untyped column's (*spec* None) are what :func:`_untyped_number_text`
    finds the cell loop reading as a number. Data rows only, ragged rows
    passed over.
    """
    if spec is not None:
        for row_idx, text in _data_cell_texts(table_data, col_idx):
            yield row_idx, text, _numeric_body(text, spec)[0]
        return
    for row_idx, row in enumerate(table_data[1:], start=1):
        if col_idx < len(row):
            body = _untyped_number_text(row[col_idx])
            if body is not None:
                yield row_idx, _strip_markdown_formatting(row[col_idx])[0].strip(), body


def _evidence_in(table_data: list[list[str]],
                 columns: list[tuple[int, str | None]]) -> dict[str, _Evidence]:
    """The evidence *columns* give: ``{EVIDENCE_*: the first value showing it}``.

    Columns in the order given, rows top to bottom; values as
    :func:`_evidence_values` yields them. Stops once both kinds are found.
    """
    found: dict[str, _Evidence] = {}
    for col_idx, spec in columns:
        for row_idx, text, body in _evidence_values(table_data, col_idx, spec):
            read = _decimal_evidence(body)
            if read is not None and read[0] not in found:
                found[read[0]] = _Evidence(row_idx, col_idx, text, read[1])
                if len(found) == 2:
                    return found
    return found


def _comma3_policies(table_data: list[list[str]],
                     col_types: list[str | None]) -> dict[int, _Comma3Policy]:
    """How each ``number`` and ``currency`` column reads ``4,375``: the evidence pass.

    Returns ``{col_idx: _Comma3Policy}``; its ``comma3`` goes to
    :func:`_apply_column_type`. Nobody declares a locale; the table's own
    numbers do (#211). Evidence is what a value proves
    (:func:`_decimal_evidence`): a decimal comma or dot grouping is European,
    a decimal point or comma grouping English; an integer, a trailing ``1.``
    or ``5,``, ``1234,567`` and the ambiguous form itself prove nothing. It
    is looked for in order, and the first place that has any decides:

    1. the column's own literal values;
    2. every literal in the table's typed ``number``, ``currency`` and
       ``percent`` columns;
    3. the cells of its untyped columns that the cell loop reads as a number
       or a percent (:func:`_untyped_number_text`) — never one it keeps as
       text.

    European only → ``'decimal'`` (``4,375`` is 4.375, and ``1.234`` is
    1234). English only → ``'thousands'``. Both → ``'refuse'``: the table
    writes numbers both ways, so ``4,375`` stays text and is reported by
    ``typed_value_kept_as_text``. None anywhere → ``'thousands'`` with
    *guessed* set, so the loop reports each such value as
    ``ambiguous_comma_read_as_thousands``. A decision taken in step 2 or 3
    keeps its first deciding value as *source*, so the loop reports each
    value it settled as ``ambiguous_number_read_from_table`` (#211 review):
    another column is a guess about this one, and a table that mixes locales
    across columns guesses wrong. Header, formula and empty cells never
    count. A ``percent`` column is not given a policy — it always refuses the
    form — but its values are evidence for the others.
    """
    width = max((len(row) for row in table_data), default=0)
    typed = [(idx, spec) for idx, spec in enumerate(col_types) if _is_numeric_type(spec)]
    untyped = [(idx, None) for idx in range(width)
               if idx >= len(col_types) or not col_types[idx]]
    table_wide: dict[str, _Evidence] | None = None    # steps 2 and 3, on first need

    policies: dict[int, _Comma3Policy] = {}
    for idx, spec in typed:
        if spec.strip().lower().startswith('percent'):
            continue
        found = _evidence_in(table_data, [(idx, spec)])
        own = bool(found)
        if not own:
            if table_wide is None:
                table_wide = (_evidence_in(table_data, typed)
                              or _evidence_in(table_data, untyped))
            found = table_wide
        if not found:
            policies[idx] = _Comma3Policy(COMMA3_THOUSANDS, guessed=True)
            continue
        if len(found) == 2:
            policies[idx] = _Comma3Policy(COMMA3_REFUSE)
            continue
        [(evidence, first)] = found.items()
        comma3 = COMMA3_DECIMAL if evidence == EVIDENCE_EUROPEAN else COMMA3_THOUSANDS
        policies[idx] = _Comma3Policy(comma3, source=None if own else first)
    return policies


def _number_format_for_type(type_spec: str | None) -> str | None:
    """Return the Excel number format a column ``types`` spec implies, or None.

    Used to format a *formula* cell sitting in a typed column. Formula cells
    bypass :func:`_apply_column_type` — their value is a formula to resolve,
    not a literal to coerce — so without this they lose the column's intended
    format and a ``=SUM(...)`` in a ``currency:$`` column renders as a bare
    number.

    Mirrors the format selection in :func:`_apply_column_type` without
    touching the cell value. Returns None for types with no numeric format
    (``text``/``bool``), for unknown specs, and for the bare ``number``/``date``
    forms whose format is chosen from the parsed value (which a formula
    doesn't have until Excel evaluates it).
    """
    if not type_spec:
        return None
    type_lower = type_spec.lower()

    if type_lower.startswith('currency'):
        symbol = _currency_symbol(type_spec)
        return _CURRENCY_FORMATS.get(symbol, f'#,##0.00 "{symbol}"')

    if type_lower.startswith('percent'):
        if ':' in type_spec:
            return type_spec.split(':', 1)[1].strip() or '0%'
        # The bare form has no precision of its own. add_table_to_sheet()
        # prefers what the column's literals imply; this is the fallback for
        # a column that has none.
        return '0%'

    if type_lower.startswith(('number', 'date')):
        # Only an explicit format applies; the bare form is value-derived.
        if ':' not in type_spec:
            return None
        return type_spec.split(':', 1)[1].strip() or None

    return None


# ── Workbook limits ───────────────────────────────────────────────────────────

# Excel rejects a formula longer than 8192 characters, and rejects the whole
# FILE rather than the one cell: the workbook opens to a repair prompt.
MAX_FORMULA_LENGTH = 8192


def _write_formula(cell, formula: str, coordinate: str,
                   warnings=None, location: dict | None = None) -> None:
    """Write a resolved formula, degrading to text if Excel would reject it.

    A single over-length formula makes Excel refuse the entire workbook, so
    the cell is stored as an inline string instead: one visibly wrong cell in
    a file that opens beats a file that doesn't — and the caller is told, since
    the cell shows the formula's text rather than its result.
    """
    cell.value = formula
    if len(formula) > MAX_FORMULA_LENGTH:
        cell.data_type = 's'  # inline string, not <f>
        logger.warning(
            "Formula in %s is %d characters, over Excel's %d limit; stored as "
            "text so the workbook still opens. Split it across helper columns.",
            coordinate, len(formula), MAX_FORMULA_LENGTH,
        )
        if warnings is not None:
            warnings.add(
                W.FORMULA_TOO_LONG,
                f"the formula is {len(formula)} characters, over Excel's "
                f"{MAX_FORMULA_LENGTH} limit; it was stored as text so the "
                f"workbook still opens. Split it across helper columns.",
                **(location or {}),
            )


def _ensure_unique_table_headers(worksheet, header_row: int, num_cols: int,
                                 warnings=None) -> None:
    """Make a header row usable as an Excel Table header.

    Excel requires every column name in a Table to be non-empty and unique
    (compared case-insensitively). openpyxl copies the names straight from the
    header cells without validating, so a duplicate or blank heading — two
    ``Q1`` columns, say — produces a file Excel opens with "we found a problem
    with some content". Blank headings get a positional name and duplicates a
    numeric suffix; both are logged, since the rename is visible to the user.

    Only called when ``auto_filter`` builds a Table; a plain table's headers
    are left exactly as written.
    """
    seen: set[str] = set()
    for col_idx in range(1, num_cols + 1):
        cell = worksheet.cell(row=header_row, column=col_idx)
        original = str(cell.value).strip() if cell.value is not None else ""

        # The positional fallback for a blank heading is just a candidate name
        # like any other: it has to clear the uniqueness check too. Exempting
        # it — as an earlier version did — reintroduced the exact duplicate
        # this function exists to prevent, whenever a real column happened to
        # be named "ColumnN" and column N was blank.
        name = original or f"Column{col_idx}"
        if name.casefold() in seen:
            base, suffix = name, 2
            while f"{base}_{suffix}".casefold() in seen:
                suffix += 1
            name = f"{base}_{suffix}"

        seen.add(name.casefold())
        if name != original:
            logger.warning(
                "Table header %s (%s) can't be used as an Excel Table column "
                "name — they must be non-empty and unique — using '%s'.",
                cell.coordinate,
                f"'{original}'" if original else "empty",
                name,
            )
            if warnings is not None:
                problem = ("is empty" if not original
                           else f"repeats the column name '{original}'")
                warnings.add(
                    W.HEADER_RENAMED,
                    f"the header {problem}, which an auto-filtered table "
                    f"cannot use — Excel Table column names must be non-empty "
                    f"and unique — so it was renamed to '{name}'.",
                    sheet=worksheet.title,
                    cell=cell.coordinate,
                )
            cell.value = name


# ── Table Rendering ───────────────────────────────────────────────────────────

def add_table_to_sheet(
    table_data: list[list[str]],
    worksheet,
    start_row: int,
    table_positions: dict[str, int] | None = None,
    all_sheet_table_positions: dict[str, dict[str, int]] | None = None,
    auto_filter: bool = False,
    table_index: int = 0,
    directives: dict[str, str] | None = None,
    available_styles: set[str] | None = None,
    warnings=None,
) -> int:
    """Add table data to Excel worksheet with proper formatting and formula support.

    Args:
        available_styles: Named styles defined in the workbook, used to
            validate a ``styles: ... style:Name`` reference. None disables the
            check (the reference is attempted regardless).
        warnings: The build's :class:`~warning_channel.WarningChannel`. Cells
            that could not be written, formulas that would not resolve and
            styling that was skipped are reported there, located by sheet and
            cell coordinate (#114).
    """
    if not table_data:
        return start_row

    directives = directives or {}

    # Parse column type hints from <!-- types: text, currency:$, date, bool --> directive
    col_types: list[str | None] = _parse_types_directive(directives.get('types', ''))

    # What a formula cell in a bare `percent` column should be formatted as.
    # It has no source text of its own, so the column's literals answer for it
    # (#126); a column with no literal percent keeps the flat fallback from
    # _number_format_for_type().
    learned_percent_formats: dict[int, str] = {}
    for idx, spec in enumerate(col_types):
        if spec and spec.strip().lower() == 'percent':
            learned = _percent_format_for_column(table_data, idx)
            if learned:
                learned_percent_formats[idx] = learned

    # Extract column alignments if available (from TableData subclass)
    col_alignments: list[str | None] = []
    if hasattr(table_data, 'col_alignments'):
        col_alignments = table_data.col_alignments

    # How each number/currency column reads the ambiguous `4,375`, from the
    # table's own evidence (#211): col_idx -> _Comma3Policy.
    comma3_policies = _comma3_policies(table_data, col_types)
    no_policy = _Comma3Policy()

    # Typed values the column could not read and kept as text, per column:
    # col_idx -> [count, first coordinate, first value, first ambiguous].
    # Reported once per column after the loop by _report_kept_as_text (#143).
    kept_as_text: dict[int, list] = {}
    # `4,375` read as 4375 with no evidence either way, per column:
    # col_idx -> [count, first coordinate, first value, compact, number].
    # Reported once per column by _report_ambiguous_read_as_thousands (#211).
    read_as_thousands: dict[int, list] = {}
    # `4,375` (or a European `1.234`) settled by another column's evidence,
    # per column: col_idx -> [count, first coordinate, first value, compact,
    # kind, number, source]. Reported once per column by
    # _report_ambiguous_read_from_table (#211 review).
    read_from_table: dict[int, list] = {}

    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    formula_fill = PatternFill(start_color="E7F3FF", end_color="E7F3FF", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Fill cells
    for row_idx, row_data in enumerate(table_data):
        current_excel_row = start_row + row_idx
        for col_idx, cell_text in enumerate(row_data):
            try:
                cell = worksheet.cell(row=current_excel_row, column=col_idx + 1)

                # If column type directive applies (data rows only), use it
                col_type = col_types[col_idx] if col_idx < len(col_types) else None
                if row_idx > 0 and col_type:
                    # Strip markdown formatting before type coercion
                    clean_text, fmt_info = _strip_markdown_formatting(cell_text)
                    # A formula cell must NOT go through type coercion: its
                    # references still need resolving by
                    # adjust_formula_references, and float('=SUM(B[0]:B[2])')
                    # raises — leaving the unresolved literal in the cell,
                    # which Excel shows as #NAME?. Fall through to the formula
                    # path below, which re-applies the column's number format.
                    is_formula = clean_text.startswith('=')
                    policy = comma3_policies.get(col_idx, no_policy)
                    if not is_formula and _apply_column_type(cell, clean_text, col_type,
                                                             comma3=policy.comma3):
                        form = (_ambiguous_form(_numeric_body(clean_text, col_type)[0])
                                if _is_numeric_type(col_type) else None)
                        compact, kind = form or (None, None)
                        ambiguous = compact if kind == _KIND_AMBIGUOUS_COMMA else None
                        # One report per cell at most: kept as text, read
                        # from another column, or read on no evidence. The
                        # mirror `1.234` only counts where the evidence moved
                        # it — read as European; read as English it is the
                        # decimal it would be with no evidence at all.
                        if _kept_as_text(cell, col_type):
                            entry = kept_as_text.setdefault(
                                col_idx, [0, cell.coordinate, cell.value, None])
                            entry[0] += 1
                            if entry[3] is None:
                                entry[3] = ambiguous
                        elif policy.source is not None and (
                                ambiguous is not None
                                or (kind == _KIND_AMBIGUOUS_DOT
                                    and policy.comma3 == COMMA3_DECIMAL)):
                            entry = read_from_table.setdefault(
                                col_idx, [0, cell.coordinate, clean_text, compact, kind,
                                          cell.value, policy.source])
                            entry[0] += 1
                        elif policy.guessed and ambiguous is not None:
                            entry = read_as_thousands.setdefault(
                                col_idx, [0, cell.coordinate, clean_text, ambiguous, cell.value])
                            entry[0] += 1
                        # Type directive handled the cell value — apply formatting, border, alignment
                        apply_cell_formatting(cell, fmt_info)
                        cell.border = border
                        explicit_align = col_alignments[col_idx] if col_idx < len(col_alignments) else None
                        if explicit_align:
                            cell.alignment = Alignment(horizontal=explicit_align)
                        elif isinstance(cell.value, bool):
                            cell.alignment = Alignment(horizontal='center')
                        elif isinstance(cell.value, (int, float, datetime)):
                            cell.alignment = Alignment(horizontal='right')
                        else:
                            cell.alignment = Alignment(horizontal='left')
                        continue

                resolved = resolve_cell(cell_text)

                if resolved.is_formula:
                    where = {"sheet": worksheet.title, "cell": cell.coordinate}
                    adjusted_formula = adjust_formula_references(
                        resolved.value, current_excel_row, table_positions,
                        all_sheet_table_positions,
                        warnings=warnings, location=where,
                    )
                    _write_formula(cell, adjusted_formula, cell.coordinate,
                                   warnings=warnings, location=where)
                    cell.fill = formula_fill
                    # A formula in a typed column takes the column's intended
                    # number format for its (numeric) result — e.g. a =SUM(...)
                    # in a `currency:$` column displays as currency. Harmless
                    # on non-numeric results: Excel ignores number formats on
                    # strings and errors.
                    if row_idx > 0 and col_type:
                        type_fmt = (learned_percent_formats.get(col_idx)
                                    or _number_format_for_type(col_type))
                        if type_fmt:
                            cell.number_format = type_fmt
                else:
                    # Header row must remain as strings — Excel Tables require
                    # string headers; numeric-looking headers (e.g. "2024") must
                    # not be converted to numbers.
                    if row_idx == 0:
                        # Use the original stripped text for headers to avoid
                        # artifacts like "2024.0" from float conversion
                        clean_header, _ = _strip_markdown_formatting(cell_text)
                        cell.value = clean_header
                    else:
                        cell.value = resolved.value

                # Apply inline formatting (bold/italic/monospace) — skip for header row
                # since header styling will override it immediately below
                if row_idx > 0:
                    apply_cell_formatting(cell, resolved.formatting_info)
                cell.border = border

                # Alignment — use explicit column alignment from separator if available,
                # otherwise fall back to heuristic
                explicit_align = col_alignments[col_idx] if col_idx < len(col_alignments) else None
                if row_idx == 0:
                    cell.alignment = Alignment(horizontal='center')
                elif explicit_align:
                    cell.alignment = Alignment(horizontal=explicit_align)
                elif isinstance(cell.value, (int, float, datetime)) or (isinstance(cell.value, str) and cell.value.startswith('=')):
                    cell.alignment = Alignment(horizontal='right')
                else:
                    cell.alignment = Alignment(horizontal='left')

                # Header row styling (overrides inline formatting)
                if row_idx == 0:
                    cell.font = header_font
                    cell.fill = header_fill
                elif isinstance(cell.value, (int, float)) and abs(cell.value) >= _THOUSANDS_THRESHOLD:
                    # abs() so that -5000 is grouped like 5000; the old
                    # `>= 1000` test left every negative value unformatted.
                    cell.number_format = _thousands_format_for(cell.value)

                # Apply percentage number format, at the source text's precision
                if resolved.is_percent and isinstance(cell.value, (int, float)):
                    cell.number_format = resolved.percent_format or '0%'

                # Apply date number format
                if resolved.is_date and resolved.date_format:
                    cell.number_format = resolved.date_format
            except Exception as e:
                logger.warning("Error processing cell [row=%d, col=%d]: %s", current_excel_row, col_idx + 1, e)
                if warnings is not None:
                    warnings.add(
                        W.CELL_FAILED,
                        f"the cell could not be written ({e}); it is empty in "
                        f"the workbook.",
                        sheet=worksheet.title,
                        cell=f"{get_column_letter(col_idx + 1)}{current_excel_row}",
                    )

    _report_kept_as_text(kept_as_text, table_data, col_types, worksheet, warnings)
    _report_ambiguous_read_as_thousands(read_as_thousands, table_data, worksheet, warnings)
    _report_ambiguous_read_from_table(read_from_table, table_data, start_row,
                                      worksheet, warnings)

    # Column widths — based on clean text length (not raw markdown with formatting markers)
    # When type directives are active, estimate display width from the type spec.
    FORMULA_WIDTH_CAP = 12  # Formulas display as numbers, cap their width contribution
    for col_idx in range(len(table_data[0]) if table_data else 0):
        column_letter = get_column_letter(col_idx + 1)
        col_type = col_types[col_idx] if col_idx < len(col_types) else None
        max_length = 0
        for row_idx, row in enumerate(table_data):
            if col_idx < len(row):
                # For data rows with a type directive, estimate from the directive.
                # Formula cells are the exception — Excel renders them as the
                # result, so their source length says nothing about display
                # width (and would blow the column out to MAX_COLUMN_WIDTH).
                # Uses the same formula test as the rendering loop above rather
                # than a second, cruder one.
                cell_is_formula = _strip_markdown_formatting(row[col_idx])[0].startswith('=')
                if row_idx > 0 and col_type and not cell_is_formula:
                    type_lower = col_type.lower()
                    if type_lower == 'bool':
                        length = 5  # "FALSE" is longest
                    elif type_lower.startswith('currency'):
                        # Symbol + number — use raw text length as decent estimate
                        length = len(row[col_idx].strip())
                    elif type_lower.startswith('date'):
                        fmt = col_type.split(':', 1)[1].strip() if ':' in col_type else "YYYY-MM-DD"
                        length = len(fmt)
                    elif type_lower.startswith('percent'):
                        fmt = (col_type.split(':', 1)[1].strip()
                               if ':' in col_type else '')
                        length = max(len(fmt), 6)  # e.g. "85.0%"
                    else:
                        length = len(row[col_idx].strip())
                else:
                    resolved = resolve_cell(row[col_idx])
                    if resolved.is_formula:
                        length = FORMULA_WIDTH_CAP
                    elif resolved.is_date:
                        length = len(resolved.date_format)
                    else:
                        length = len(str(resolved.value))
                max_length = max(max_length, length)
        adjusted_width = min(max(max_length + COLUMN_WIDTH_PADDING, MIN_COLUMN_WIDTH), MAX_COLUMN_WIDTH)
        worksheet.column_dimensions[column_letter].width = adjusted_width

    # Explicit cell styling from <!-- styles: B2=bg:yellow, C[0]=style:Total -->.
    # Applied as a final pass over the finished table, deliberately: it is the
    # caller's explicit instruction, so it must win over the header fill, the
    # formula fill and inline **bold**, and a single pass here beats repeating
    # the lookup in each branch of the cell loop above.
    styles_directive = directives.get('styles', '')
    if styles_directive:
        from .styles import apply_style_spec, parse_styles_directive

        parsed = parse_styles_directive(styles_directive, start_row,
                                        warnings=warnings, sheet=worksheet.title)
        for coordinate, spec in parsed.items():
            try:
                apply_style_spec(worksheet[coordinate], spec, available_styles)
            except Exception as e:
                logger.warning("Could not style %s: %s", coordinate, e)
                if warnings is not None:
                    warnings.add(
                        W.STYLE_FAILED,
                        f"the styles directive could not be applied ({e}); the "
                        f"cell keeps its default formatting.",
                        sheet=worksheet.title,
                        cell=coordinate,
                    )

    # Auto-filter: create a proper Excel Table object (supports multiple per sheet)
    if auto_filter:
        num_cols = len(table_data[0]) if table_data else 0
        if num_cols > 0:
            # Must run before the Table is built: openpyxl reads the column
            # names off these cells, and a blank or duplicate one yields a
            # workbook Excel refuses to open.
            _ensure_unique_table_headers(worksheet, start_row, num_cols,
                                         warnings=warnings)
            last_col_letter = get_column_letter(num_cols)
            last_data_row = start_row + len(table_data) - 1
            table_ref = f"A{start_row}:{last_col_letter}{last_data_row}"
            # Excel table names must be unique across the workbook
            table_name = f"Table_{worksheet.title.replace(' ', '_')}_{table_index + 1}"
            # Sanitize: Excel table names allow only letters, digits, underscores
            table_name = re.sub(r'[^A-Za-z0-9_]', '', table_name)
            excel_table = Table(displayName=table_name, ref=table_ref)
            excel_table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2", showFirstColumn=False,
                showLastColumn=False, showRowStripes=True, showColumnStripes=False,
            )
            worksheet.add_table(excel_table)

    return start_row + len(table_data) + TABLE_BOTTOM_SPACING
