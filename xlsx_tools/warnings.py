"""The Excel builder's warning codes, and the channel it writes them to.

Excel used to lose content and mis-resolve formulas in silence (#114): a line
that is not a table row is dropped, a formula naming a table or sheet that
does not exist still resolves — to the wrong cell, or to ``#REF!`` once Excel
opens it — and an over-length formula is stored as text. Each produced a
success response and a log line the caller never sees. A value a ``types``
column cannot read is kept as text the same way (#143); that one is reported
once per column, with a count, rather than once per cell. A ``number`` or
``currency`` column that reads ``4,375`` as 4375 with nothing in its table to
say whether that comma groups thousands or marks a decimal says so too, the
same way (#211), and so does one whose ambiguous values were settled by how
*other* columns of its table write numbers, since that is a guess about this
one (#211 review).

Each code below names one of those situations, and :data:`WARNING_SEVERITY`
fixes what it costs. Locations are the ``sheet`` the problem is on plus either
the ``cell`` coordinate or the source ``line`` of the markdown, counting
from 1.

See :mod:`warning_channel` for the record shape and the collector, and
docs/development/tools/excel.md for where each one is raised.
"""

from __future__ import annotations

from warning_channel import (
    SEVERITY_ERROR,
    SEVERITY_INFO,
    SEVERITY_WARNING,
    WarningChannel,
)

# Content the caller wrote that is not in the workbook.
LINE_DROPPED = "line_dropped"
TABLE_INCOMPLETE = "table_incomplete"
CELL_FAILED = "cell_failed"

# Formulas that were written, but will not compute what was asked.
TABLE_REFERENCE_MISSING = "table_reference_missing"
SHEET_REFERENCE_MISSING = "sheet_reference_missing"
FORMULA_UNRESOLVED = "formula_unresolved"
FORMULA_TOO_LONG = "formula_too_long"
CIRCULAR_REFERENCE = "circular_reference"

# Names and formatting the workbook did not take as asked.
TABLE_SEPARATOR_MISSING = "table_separator_missing"
SHEET_NAME_COLLISION = "sheet_name_collision"
SHEET_NAME_INVALID = "sheet_name_invalid"
HEADER_RENAMED = "header_renamed"
STYLE_ENTRY_INVALID = "style_entry_invalid"
STYLE_RANGE_TOO_LARGE = "style_range_too_large"
STYLE_FAILED = "style_failed"
TYPED_VALUE_KEPT_AS_TEXT = "typed_value_kept_as_text"

# A value read one of two ways, with nothing in the table to say which — or
# with only the table's other columns to say which.
AMBIGUOUS_COMMA_READ_AS_THOUSANDS = "ambiguous_comma_read_as_thousands"
AMBIGUOUS_NUMBER_READ_FROM_TABLE = "ambiguous_number_read_from_table"

#: One severity per code, in one place.
WARNING_SEVERITY: dict[str, str] = {
    LINE_DROPPED: SEVERITY_ERROR,
    TABLE_INCOMPLETE: SEVERITY_ERROR,
    CELL_FAILED: SEVERITY_ERROR,

    TABLE_REFERENCE_MISSING: SEVERITY_ERROR,
    SHEET_REFERENCE_MISSING: SEVERITY_ERROR,
    FORMULA_UNRESOLVED: SEVERITY_ERROR,
    FORMULA_TOO_LONG: SEVERITY_ERROR,
    CIRCULAR_REFERENCE: SEVERITY_ERROR,

    TABLE_SEPARATOR_MISSING: SEVERITY_WARNING,
    SHEET_NAME_COLLISION: SEVERITY_WARNING,
    SHEET_NAME_INVALID: SEVERITY_WARNING,
    HEADER_RENAMED: SEVERITY_WARNING,
    STYLE_ENTRY_INVALID: SEVERITY_WARNING,
    STYLE_RANGE_TOO_LARGE: SEVERITY_WARNING,
    STYLE_FAILED: SEVERITY_WARNING,
    TYPED_VALUE_KEPT_AS_TEXT: SEVERITY_WARNING,

    AMBIGUOUS_COMMA_READ_AS_THOUSANDS: SEVERITY_INFO,
    AMBIGUOUS_NUMBER_READ_FROM_TABLE: SEVERITY_INFO,
}


def channel() -> WarningChannel:
    """A fresh channel for one Excel build, bound to the table above."""
    return WarningChannel(WARNING_SEVERITY)
