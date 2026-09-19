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
    SEVERITY_WARNING,
    WarningChannel,
)

# Content the caller wrote that is not in the document.
BLOCK_FAILED = "block_failed"
TABLE_FAILED = "table_failed"
TABLE_CELL_FAILED = "table_cell_failed"
IMAGE_FAILED = "image_failed"

# Instructions the document did not follow.
STYLE_MISSING = "style_missing"
STYLE_FALLBACK_MISSING = "style_fallback_missing"
WIDTHS_INVALID = "widths_invalid"

#: One severity per code, in one place.
WARNING_SEVERITY: dict[str, str] = {
    BLOCK_FAILED: SEVERITY_ERROR,
    TABLE_FAILED: SEVERITY_ERROR,
    TABLE_CELL_FAILED: SEVERITY_ERROR,
    IMAGE_FAILED: SEVERITY_ERROR,

    STYLE_MISSING: SEVERITY_WARNING,
    STYLE_FALLBACK_MISSING: SEVERITY_WARNING,
    WIDTHS_INVALID: SEVERITY_WARNING,
}


def channel() -> WarningChannel:
    """A fresh channel for one Word build, bound to the table above."""
    return WarningChannel(WARNING_SEVERITY)
