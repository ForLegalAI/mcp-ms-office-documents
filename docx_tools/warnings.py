"""The Word builder's warning codes, and the channel it writes them to.

Word used to lose content in silence (#114): a block whose rendering raised
was logged and skipped, an image that would not download became a bracketed
placeholder, a style the template does not define fell back to ``Normal``.
Every one of those produced a success response, so the caller — a model that
could have fixed its markdown — never learned.

Each code below names one of those situations, and
:data:`WARNING_SEVERITY` fixes what it costs. Locations are the source
``line`` of the markdown the caller sent, counting from 1.

See :mod:`warning_channel` for the record shape and the collector, and
docs/development/tools/word.md for where each one is raised.
"""

from __future__ import annotations

from warning_channel import (
    SEVERITY_ERROR,
    SEVERITY_INFO,
    SEVERITY_WARNING,
    WarningChannel,
)

# Content the caller wrote that is not in the document.
BLOCK_FAILED = "block_failed"
TABLE_FAILED = "table_failed"
TABLE_CELL_FAILED = "table_cell_failed"
IMAGE_FAILED = "image_failed"

# Instructions the document did not follow.
TABLE_NOT_RECOGNISED = "table_not_recognised"
TABLE_SEPARATOR_MISSING = "table_separator_missing"
STYLE_MISSING = "style_missing"
STYLE_FALLBACK_MISSING = "style_fallback_missing"
WIDTHS_INVALID = "widths_invalid"
LINK_REFUSED = "link_refused"

# Dynamic templates: a Word content control the caller sent a value for, that
# the document could not take (see content_controls). Located by ``tag``.
CONTROL_ITEM_MISSING = "control_item_missing"
CONTROL_NOT_FILLED = "control_not_filled"
CONTROL_VALUE_INVALID = "control_value_invalid"
CONTROL_DATE_FORMAT_SIMPLIFIED = "control_date_format_simplified"

#: One severity per code, in one place.
WARNING_SEVERITY: dict[str, str] = {
    BLOCK_FAILED: SEVERITY_ERROR,
    TABLE_FAILED: SEVERITY_ERROR,
    TABLE_CELL_FAILED: SEVERITY_ERROR,
    IMAGE_FAILED: SEVERITY_ERROR,
    CONTROL_ITEM_MISSING: SEVERITY_ERROR,
    CONTROL_NOT_FILLED: SEVERITY_ERROR,
    CONTROL_VALUE_INVALID: SEVERITY_ERROR,

    TABLE_NOT_RECOGNISED: SEVERITY_WARNING,
    TABLE_SEPARATOR_MISSING: SEVERITY_WARNING,
    STYLE_MISSING: SEVERITY_WARNING,
    STYLE_FALLBACK_MISSING: SEVERITY_WARNING,
    WIDTHS_INVALID: SEVERITY_WARNING,
    LINK_REFUSED: SEVERITY_WARNING,

    CONTROL_DATE_FORMAT_SIMPLIFIED: SEVERITY_INFO,
}


def channel() -> WarningChannel:
    """A fresh channel for one Word build, bound to the table above."""
    return WarningChannel(WARNING_SEVERITY)
