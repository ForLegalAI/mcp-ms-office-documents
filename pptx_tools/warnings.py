"""The warnings channel: what the builder had to work around, as data.

``PowerpointPresentation.warnings`` used to be a list of English sentences.
A caller could show them to a human and nothing else — not branch on them,
not count them by kind, not decide whether a deck was worth keeping (#122).

Each entry is now a :class:`SlideWarning` carrying:

* ``code`` — a stable identifier, the thing to branch on;
* ``slide`` — the slide's index in the ``slides`` list the caller sent, so a
  model can correct the element it got wrong, or None for the deck as a whole.
  It is the number the message has always carried, counting from zero;
* ``severity`` — how much it cost, derived from the code so one code always
  means one severity;
* ``message`` — the same sentence as before, unchanged.

Severity is about the deck, not about how unusual the situation is:

``error``
    Something the caller asked for is not in the file. A dropped subtitle, an
    image that would not load, an element skipped for being off-slide.
``warning``
    It is in the file, but not as asked: shrunk, clamped, or laid out on a
    layout that was not the obvious one.
``info``
    A substitution the caller probably does not mind — a format or template
    picked for them, a table font reduced a point or two.

``str(warning)`` renders the line the channel used to hold, which is what the
logs and the admin preview header still show.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional

SEVERITY_ERROR = "error"
SEVERITY_WARNING = "warning"
SEVERITY_INFO = "info"

# -- codes ------------------------------------------------------------------
# Deck-level: no single slide is responsible.
FORMAT_SUBSTITUTED = "format_substituted"
TEMPLATE_SUBSTITUTED = "template_substituted"
TEMPLATE_UNREADABLE = "template_unreadable"
FOOTER_UNSUPPORTED = "footer_unsupported"

# Template configuration, reported against the slide that read it.
TEMPLATE_OPTION_INVALID = "template_option_invalid"
TEMPLATE_FONT_SIZE_INVALID = "template_font_size_invalid"
TEMPLATE_FONT_SIZE_CLAMPED = "template_font_size_clamped"

# Layout and placeholders.
LAYOUT_SUBSTITUTED = "layout_substituted"
TITLE_POSITION_GUESSED = "title_position_guessed"
SUBTITLE_DROPPED = "subtitle_dropped"
BULLETS_DROPPED = "bullets_dropped"
CLOSING_LINES_DROPPED = "closing_lines_dropped"
AGENDA_DROPPED = "agenda_dropped"
AGENDA_EMPTY = "agenda_empty"

# Content that did not fit, or did not arrive.
TEXT_OVERFLOW = "text_overflow"
TABLE_EMPTY = "table_empty"
TABLE_OVERFLOW = "table_overflow"
TABLE_FONT_REDUCED = "table_font_reduced"
IMAGE_FAILED = "image_failed"
CHART_FAILED = "chart_failed"
CHART_SERIES_LENGTH = "chart_series_length"
KPI_CROWDED = "kpi_crowded"
TIMELINE_DETAIL_DROPPED = "timeline_detail_dropped"

# Positioned elements on a blank slide.
ELEMENT_OFF_SLIDE = "element_off_slide"
ELEMENT_CLAMPED = "element_clamped"
ELEMENT_ZERO_SIZE = "element_zero_size"

# One severity per code, in one place: a caller that learns
# ``image_failed`` is an error must never meet one that is not.
WARNING_SEVERITY: Dict[str, str] = {
    FORMAT_SUBSTITUTED: SEVERITY_INFO,
    TEMPLATE_SUBSTITUTED: SEVERITY_INFO,
    TEMPLATE_UNREADABLE: SEVERITY_WARNING,
    FOOTER_UNSUPPORTED: SEVERITY_WARNING,

    TEMPLATE_OPTION_INVALID: SEVERITY_WARNING,
    TEMPLATE_FONT_SIZE_INVALID: SEVERITY_WARNING,
    TEMPLATE_FONT_SIZE_CLAMPED: SEVERITY_WARNING,

    LAYOUT_SUBSTITUTED: SEVERITY_WARNING,
    TITLE_POSITION_GUESSED: SEVERITY_WARNING,
    SUBTITLE_DROPPED: SEVERITY_ERROR,
    BULLETS_DROPPED: SEVERITY_ERROR,
    CLOSING_LINES_DROPPED: SEVERITY_ERROR,
    AGENDA_DROPPED: SEVERITY_ERROR,
    AGENDA_EMPTY: SEVERITY_WARNING,

    TEXT_OVERFLOW: SEVERITY_WARNING,
    TABLE_EMPTY: SEVERITY_WARNING,
    TABLE_OVERFLOW: SEVERITY_WARNING,
    TABLE_FONT_REDUCED: SEVERITY_INFO,
    IMAGE_FAILED: SEVERITY_ERROR,
    CHART_FAILED: SEVERITY_ERROR,
    CHART_SERIES_LENGTH: SEVERITY_WARNING,
    KPI_CROWDED: SEVERITY_INFO,
    TIMELINE_DETAIL_DROPPED: SEVERITY_ERROR,

    ELEMENT_OFF_SLIDE: SEVERITY_ERROR,
    ELEMENT_CLAMPED: SEVERITY_WARNING,
    ELEMENT_ZERO_SIZE: SEVERITY_ERROR,
}


@dataclass(frozen=True)
class SlideWarning:
    """One thing the builder worked around, and where."""

    code: str
    message: str
    slide: Optional[int] = None    # index into the caller's slides list
    severity: str = SEVERITY_WARNING

    def __str__(self) -> str:
        return f"slide {self.slide}: {self.message}" if self.slide is not None else self.message

    def as_dict(self) -> Dict[str, Any]:
        """The shape the tool returns; ``slide`` is omitted for a deck-wide one."""
        payload: Dict[str, Any] = {
            "code": self.code,
            "severity": self.severity,
            "message": self.message,
        }
        if self.slide is not None:
            payload["slide"] = self.slide
        return payload


def make_warning(code: str, message: str, slide: Optional[int] = None) -> SlideWarning:
    """A warning whose severity comes from its code.

    An unregistered code is a programming error, but not one worth failing a
    finished deck over: it is reported at the default severity, and
    ``test_every_code_has_a_severity`` keeps the table honest.
    """
    return SlideWarning(
        code=code,
        message=message,
        slide=slide,
        severity=WARNING_SEVERITY.get(code, SEVERITY_WARNING),
    )
