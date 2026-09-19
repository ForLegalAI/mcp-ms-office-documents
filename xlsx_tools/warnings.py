"""The Excel builder's warning codes, and the channel it writes them to.

Excel used to lose content and mis-resolve formulas in silence (#114): a line
that is not a table row is dropped, a formula naming a table or sheet that
does not exist still resolves — to the wrong cell, or to ``#REF!`` once Excel
opens it — and an over-length formula is stored as text. Each produced a
success response and a log line the caller never sees.

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
SHEET_NAME_COLLISION = "sheet_name_collision"
SHEET_NAME_INVALID = "sheet_name_invalid"
HEADER_RENAMED = "header_renamed"
STYLE_ENTRY_INVALID = "style_entry_invalid"
STYLE_RANGE_TOO_LARGE = "style_range_too_large"
STYLE_FAILED = "style_failed"

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

    SHEET_NAME_COLLISION: SEVERITY_WARNING,
    SHEET_NAME_INVALID: SEVERITY_WARNING,
    HEADER_RENAMED: SEVERITY_WARNING,
    STYLE_ENTRY_INVALID: SEVERITY_WARNING,
    STYLE_RANGE_TOO_LARGE: SEVERITY_WARNING,
    STYLE_FAILED: SEVERITY_WARNING,
}


def channel() -> WarningChannel:
    """A fresh channel for one Excel build, bound to the table above."""
    return WarningChannel(WARNING_SEVERITY)
