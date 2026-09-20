"""PowerPoint slide builder class.

Builds a deck from the typed slide models in :mod:`pptx_tools.schema`, using
the SlideHelpers mixin for text, tables, images and charts.

Anything the builder had to work around — an image that would not download, a
slide whose text will not fit, a layout the template does not provide — is
recorded on ``self.warnings`` and returned to the caller alongside the file.
Those situations used to be logged server-side only, so the model was told the
deck was fine and had no way to correct itself.
"""

import io
import copy
import re
import logging
from typing import Any, List, Optional, Sequence, Tuple

from pptx import Presentation
from pptx.dml.color import MSO_THEME_COLOR, RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

from .constants import (
    MARGIN_LEFT, TABLE_ALT_ROW_FILL,
    DEFAULT_BODY_FONT_SIZE, DEFAULT_SUBTITLE_FONT_SIZE,
    DEFAULT_CAPTION_FONT_SIZE, DEFAULT_QUOTE_FONT_SIZE,
    KPI_VALUE_FONT_SIZE, TIMELINE_DETAIL_FONT_SIZE,
    TIMELINE_STEP_HEIGHT, TIMELINE_STEP_MIN_HEIGHT,
    TIMELINE_DETAIL_GAP, TIMELINE_DETAIL_HEIGHT,
    TABLE_HEADER_FILL, TABLE_FONT_SIZE_RANGE,
    DEFAULT_SLIDE_FORMAT, VALID_SLIDE_FORMATS,
)
from .helpers import (
    SlideHelpers,
    apply_autofit, body_to_bullets, estimate_text_fill, parse_table_data,
    resolve_fill, set_runs_language, table_overflows,
)
from .inline_formatting import write_text
from .chart_utils import (
    add_chart_to_slide, add_scatter_to_slide, configure_data_labels,
    set_axis_titles, ChartDataError,
)
from .layouts import LayoutResolver, role_for_slide
from .placeholder_style import (
    TitleStyle, apply_list_style, apply_text_color, content_columns,
    draw_title_box, read_body_color, read_body_font_size,
    read_master_body_font_size,
)
from .text_metrics import theme_body_typeface
from .schema import Bullet, coerce_slides
from . import warnings as W
from .warnings import SlideWarning, make_warning
from .templates import TemplateSpec, open_template, select_template

logger = logging.getLogger(__name__)


class PowerpointPresentation(SlideHelpers):
    """Builder class for creating PowerPoint presentations from structured data."""

    def __init__(self, slides: Sequence[Any], format: str,
                 author: Optional[str] = None,
                 footer_text: Optional[str] = None,
                 show_slide_numbers: bool = False,
                 language: Optional[str] = None,
                 template: Optional[str] = None,
                 template_spec: Optional[TemplateSpec] = None):
        """Initialize and build presentation.

        Args:
            slides: Slide models, or plain dicts in either the current or the
                previous key spelling — both are validated through
                :func:`~pptx_tools.schema.coerce_slides`.
            format: Presentation format ("4:3" or "16:9"). Ignored when
                *template* names a registered template.
            author: Author name stored in document metadata/properties.
            footer_text: Optional footer text displayed on all slides.
            show_slide_numbers: Whether to show slide numbers on all slides.
            language: BCP-47 tag (e.g. "cs-CZ") stamped on every text run of
                the slides, their tables and their notes, so the deck is
                proof-read in the right language. Text that lives inside a
                chart part is not reached, and keeps the viewer's own language.
            template: Name of a registered template. Overrides *format*.
            template_spec: An already-resolved template to build on, bypassing
                the registry lookup. Overrides *template*. This exists for the
                admin UI's preview, which must render a template the admin has
                uploaded but not yet saved — going through the registry would
                mean either finding nothing or writing the entry before the
                admin has agreed to it.
        """
        if not slides:
            raise ValueError("At least one slide is required")

        self.warnings: List[SlideWarning] = []
        self.slides = coerce_slides(slides)

        logger.info(
            "Initializing PowerPoint: slides=%d, format=%s, template=%s",
            len(self.slides), format, template,
        )

        self.spec = None
        self.presentation = self._create_presentation(format, template, template_spec)
        self._layouts = LayoutResolver(
            self.presentation, self.spec.layouts if self.spec else None
        )
        # The face body text is actually set in, so the fit estimate measures
        # this deck's text rather than a generic one (#125).
        self._typeface = theme_body_typeface(self.presentation)
        self._body_size = None
        # Text the builder draws itself is a plain text box, so it inherits
        # the presentation's default text style (tx1, black) rather than the
        # body style a placeholder gets. On a dark template that is black on
        # near-black (#195), so the template's own body colour is applied.
        self._body_color = read_body_color(
            self.presentation.slide_masters[0]
            if self.presentation.slide_masters else None
        )

        defaults = self.spec.defaults if self.spec else {}
        self._footer_text = footer_text if footer_text is not None else defaults.get("footer_text")
        self._show_slide_numbers = (
            show_slide_numbers if show_slide_numbers else bool(defaults.get("show_slide_numbers"))
        )
        self._language = language if language is not None else defaults.get("language")
        self._table_defaults = defaults.get("table") or {}
        self._chart_defaults = defaults.get("chart") or {}

        self._remove_template_slides()
        self._build_slides(self.slides)
        self._drop_unused_placeholders()
        self._apply_sections()
        if self._footer_text or self._show_slide_numbers:
            self._apply_footer_and_slide_numbers()
        if self._language:
            self._apply_language(self._language)
        if author:
            self.presentation.core_properties.author = author

    # -------------------------------------------------------------------------
    # Setup
    # -------------------------------------------------------------------------

    def _warn(self, slide_index: int, code: str, message: str) -> None:
        """Record a caller-visible warning about one slide."""
        self._record(make_warning(code, message, slide=slide_index))

    def _warn_deck(self, code: str, message: str) -> None:
        """Record a warning about the deck rather than any one slide."""
        self._record(make_warning(code, message))

    def _record(self, warning: SlideWarning) -> None:
        self.warnings.append(warning)
        logger.warning("[pptx] %s", warning)

    @property
    def warning_messages(self) -> List[str]:
        """The warnings as the lines they read as, for logs and for people."""
        return [str(warning) for warning in self.warnings]

    @staticmethod
    def _setting(slide_data, field: str, defaults: dict, key: str, fallback):
        """An option's effective value: the slide's, else the template's, else *fallback*.

        The registry documents ``defaults`` as "applied when the tool call does
        not set the same option", so precedence is the explicit slide value,
        then the template default, then the built-in. "Set" is read from
        ``model_fields_set`` rather than the value: ``zebra`` and
        ``data_labels`` default to real booleans, so their value alone cannot
        say whether the caller chose it. An explicit ``null`` counts as unset.

        Until this existed the table and chart defaults were parsed into the
        spec and stored on the builder but never consulted — four documented
        options that changed nothing.
        """
        if field in slide_data.model_fields_set:
            value = getattr(slide_data, field)
            if value is not None:
                return value
        if key in defaults and defaults[key] is not None:
            return defaults[key]
        return fallback

    def _coerce_bool(self, value, option: str, index: int, fallback: bool) -> bool:
        """A boolean from a template default, which YAML may have left as text.

        A quoted ``"false"`` in the registry parses as the string ``"false"``,
        and ``bool("false")`` is True — the setting would invert with no
        warning. Text is read the way a person meant it; anything unrecognised
        is reported and the built-in used. A slide's own value is always a real
        bool by the time it gets here, so this only ever acts on template input.
        """
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return bool(value)
        text = str(value).strip().lower()
        if text in ("true", "yes", "on", "1"):
            return True
        if text in ("false", "no", "off", "0"):
            return False
        self._warn(index, W.TEMPLATE_OPTION_INVALID,
                   f"template {option} {value!r} is not true/false; used {fallback}.")
        return fallback

    def _coerce_font_size(self, value, index: int) -> Optional[int]:
        """An integer point size within the range the slide schema enforces.

        A template default bypasses that schema, so a typo like 200 was applied
        verbatim — a broken-looking table where the non-numeric case already
        produced a warning. Out-of-range values are clamped and reported.
        """
        try:
            points = int(value)
        except (TypeError, ValueError):
            self._warn(index, W.TEMPLATE_FONT_SIZE_INVALID,
                       f"template table font_size {value!r} is not a number; ignored.")
            return None
        low, high = TABLE_FONT_SIZE_RANGE
        if not low <= points <= high:
            clamped = max(low, min(points, high))
            self._warn(
                index, W.TEMPLATE_FONT_SIZE_CLAMPED,
                f"template table font_size {points} is outside {low}–{high}; used {clamped}.",
            )
            return clamped
        return points

    def _create_presentation(self, format: str, template: Optional[str] = None,
                             template_spec: Optional[TemplateSpec] = None) -> Presentation:
        """Create the presentation from the selected registered template."""
        if format not in VALID_SLIDE_FORMATS:
            logger.warning(
                "Unknown presentation format %r; using %s. Valid formats: %s",
                format, DEFAULT_SLIDE_FORMAT, ", ".join(VALID_SLIDE_FORMATS),
            )
            self._warn_deck(
                W.FORMAT_SUBSTITUTED,
                f"Unknown format {format!r}; used {DEFAULT_SLIDE_FORMAT}.",
            )
            format = DEFAULT_SLIDE_FORMAT

        if template_spec is not None:
            spec, note = template_spec, None
        else:
            spec, note = select_template(template, None if template else format)
        if note:
            self._warn_deck(W.TEMPLATE_SUBSTITUTED, note)

        if spec is not None:
            try:
                presentation = open_template(spec.path)
                self.spec = spec
                logger.info(
                    "Using template %r (%s, %s)", spec.name, spec.path.name, spec.aspect
                )
                return presentation
            except Exception as e:
                logger.error("Failed to load template %s: %s", spec.path.name, e)
                self._warn_deck(
                    W.TEMPLATE_UNREADABLE,
                    f"Template {spec.name!r} could not be opened ({e}); "
                    "used the built-in PowerPoint theme.",
                )

        logger.warning("Using the built-in PowerPoint theme for %s", format)
        return Presentation()

    def _apply_title(self, slide, text, index: int) -> None:
        """Set the slide title, drawing it by hand when the layout has none.

        A Blank layout never has a title placeholder — that is what makes it
        blank — so a `blank` slide's title was accepted, validated and then
        dropped with a warning on every single run (#118). It is now drawn as
        a text box where this template puts its titles, in the style this
        template gives them, which is what the placeholder would have done.

        Only a template that defines no title anywhere still warns: the box
        then goes in a band across the top of the slide, which is a guess.
        """
        if self._set_title(slide, text):
            return
        if not text:
            return

        style = self._layouts.title_style()
        if style is None:
            style = self._default_title_style()
            self._warn(
                index, W.TITLE_POSITION_GUESSED,
                f"layout {slide.slide_layout.name!r} has no title placeholder and "
                "no layout in this template has one; the title was drawn as a "
                "text box across the top of the slide.",
            )
        box = draw_title_box(slide, text, style)
        # add_textbox() defaults to grow-the-shape; a title stays the size the
        # template gave it and shrinks its text instead, as a placeholder does.
        apply_autofit(box.text_frame)

    def _default_title_style(self) -> TitleStyle:
        """A title band across the top, for a template that positions none."""
        return TitleStyle(
            left=MARGIN_LEFT,
            top=Inches(0.3),
            width=self.presentation.slide_width - 2 * MARGIN_LEFT,
            height=Inches(1.0),
        )

    def _content_slide(self, slide_data, index: int):
        """Resolve, create and title a slide, and return its content rectangle.

        Shared by every slide type that draws into the body area itself
        (table, chart, scatter, quote). An image slide resolves and titles its
        slide the same way but inspects it for a picture placeholder first, so
        it calls the two steps itself.
        """
        slide = self._new_slide(slide_data, index)
        self._apply_title(slide, slide_data.title, index)
        return self._add_title_content_slide("", slide=slide)

    def _new_slide(self, slide_data, index: int):
        """Add a slide on the layout resolved for its type.

        This is what replaced indexing ``slide_layouts`` by position. A
        template that reorders or omits layouts is now matched by name and by
        placeholder signature, and anything approximate about the match is
        reported to the caller rather than silently producing a deck laid out
        on the wrong layouts.
        """
        role = role_for_slide(slide_data, self._layouts)
        layout, note = self._layouts.resolve(role, slide_data.layout)
        if note:
            self._warn(index, W.LAYOUT_SUBSTITUTED, note)
        return self.presentation.slides.add_slide(layout)

    # Chrome is cloned and filled by _apply_footer_and_slide_numbers, which
    # runs after this pass; an empty one there is not an unused placeholder.
    _CHROME_PLACEHOLDERS = frozenset((
        PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER,
    ))

    def _drop_unused_placeholders(self) -> None:
        """Remove placeholders no slide content reached.

        ``add_slide()`` copies every placeholder its layout defines, so a
        layout offering more than the slide filled — the third card of a
        three-card layout, the heading strip of a Comparison column a caller
        gave no heading, the body of a Section Header — left a box reading
        "Click to add text" in the deck. It does not print and does not show
        in a slideshow, but it is the first thing anyone opening the file to
        edit it sees, and on the template in #195 there were three per slide.

        A placeholder holding a picture, table or chart is not an ``<p:sp>``
        with a text frame any more, so filling one keeps it. Dropping a
        placeholder does not change what a reader sees; PowerPoint's Reset
        Slide puts it back from the layout.
        """
        for slide in self.presentation.slides:
            for placeholder in list(slide.placeholders):
                if placeholder.placeholder_format.type in self._CHROME_PLACEHOLDERS:
                    continue
                if not placeholder.has_text_frame:
                    continue
                if placeholder.text_frame.text.strip():
                    continue
                element = placeholder._element
                element.getparent().remove(element)

    def _remove_template_slides(self) -> None:
        """Remove every slide the template ships with, parts included.

        Dropping only the ``<p:sldId>`` entry leaves the slide's relationship in
        place, and python-pptx serialises parts by walking relationships — so a
        template carrying a sample slide shipped that slide's XML, text and
        images inside every generated file even though PowerPoint showed the
        right slide count. Dropping the relationship removes the part too.

        A template can opt out with ``strip_slides: false`` when its own slides
        are meant to survive — a fixed cover or back page the generated slides
        should follow. Until this check existed the option was parsed, stored
        and offered in the admin UI but never read, so unticking the box
        changed nothing and said nothing.
        """
        if self.spec is not None and not self.spec.strip_slides:
            logger.debug(
                "Template %r sets strip_slides: false; keeping its %d slide(s)",
                self.spec.name, len(self.presentation.slides._sldIdLst),
            )
            return

        sldIdLst = self.presentation.slides._sldIdLst
        prs_part = self.presentation.part

        removed = 0
        for sldId in list(sldIdLst):
            rId = sldId.rId
            try:
                sldIdLst.remove(sldId)
                prs_part.drop_rel(rId)
                removed += 1
            except Exception as e:
                logger.warning("Could not fully remove template slide %s: %s", rId, e)

        if removed:
            logger.debug("Removed %d slide(s) carried by the template", removed)

    # PowerPoint's outline-pane sections live in a p14 extension on
    # <p:presentation>. python-pptx has no API for them, so the XML is written
    # directly, mirroring how the shipped template declares its own p15
    # extension: the namespace on the extension element itself.
    _SECTION_LIST_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
    _P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
    _DEFAULT_SECTION_NAME = "Default Section"

    def _apply_sections(self) -> None:
        """Group slides under each ``section`` slide in the outline pane.

        A ``section`` slide already divides the deck for the audience; this
        makes the same division visible to the presenter, so the slide sorter
        and outline pane show the deck's structure instead of a flat list.
        Slides before the first section slide — the title slide, an agenda —
        go in a section named the way PowerPoint names its own.

        Every slide must belong to a section once a section list exists, or
        PowerPoint reports the file as needing repair, so the grouping is
        computed over the actual slide-id list rather than the models: with
        ``strip_slides: false`` the template's own slides come first, and they
        are folded into the leading section.
        """
        from uuid import uuid4
        from xml.sax.saxutils import escape

        section_at = [i for i, slide in enumerate(self.slides) if slide.type == "section"]
        if not section_at:
            return

        slide_ids = [int(entry.id) for entry in self.presentation.slides._sldIdLst]
        offset = len(slide_ids) - len(self.slides)
        if offset < 0:
            # Fewer slides than models can only mean a build error upstream;
            # a wrong grouping is worse than none.
            logger.warning("[pptx] slide count is below the model count; sections skipped")
            return

        groups: List[Tuple[str, List[int]]] = []
        leading = slide_ids[: offset + section_at[0]]
        if leading:
            groups.append((self._DEFAULT_SECTION_NAME, leading))
        bounds = section_at + [len(self.slides)]
        for number, (start, end) in enumerate(zip(bounds, bounds[1:]), start=1):
            title = (self.slides[start].title or "").strip() or f"Section {number}"
            groups.append((title, slide_ids[offset + start: offset + end]))

        root = self.presentation._element
        ext_lst = root.find(qn("p:extLst"))
        if ext_lst is None:
            ext_lst = parse_xml(f'<p:extLst xmlns:p="{root.nsmap["p"]}"/>')
            root.append(ext_lst)
        # A template may carry a section list of its own; ours replaces it,
        # since the slides it referred to are gone.
        for ext in ext_lst.findall(qn("p:ext")):
            if ext.get("uri") == self._SECTION_LIST_EXT_URI:
                ext_lst.remove(ext)

        sections_xml = "".join(
            f'<p14:section name="{escape(name, {chr(34): "&quot;"})}" '
            f'id="{{{str(uuid4()).upper()}}}"><p14:sldIdLst>'
            + "".join(f'<p14:sldId id="{sid}"/>' for sid in ids)
            + "</p14:sldIdLst></p14:section>"
            for name, ids in groups
        )
        ext_lst.append(parse_xml(
            f'<p:ext xmlns:p="{root.nsmap["p"]}" uri="{self._SECTION_LIST_EXT_URI}">'
            f'<p14:sectionLst xmlns:p14="{self._P14_NS}">{sections_xml}</p14:sectionLst>'
            "</p:ext>"
        ))
        logger.debug("Wrote %d outline section(s)", len(groups))

    def _build_slides(self, slides: Sequence[Any]) -> None:
        """Build all slides from validated models."""
        builders = {
            "title": self._build_title_slide,
            "section": self._build_section_slide,
            "content": self._build_content_slide,
            "table": self._build_table_slide,
            "image": self._build_image_slide,
            "two_column": self._build_two_column_slide,
            "chart": self._build_chart_slide,
            "scatter": self._build_scatter_slide,
            "quote": self._build_quote_slide,
            "kpi": self._build_kpi_slide,
            "agenda": self._build_agenda_slide,
            "closing": self._build_closing_slide,
            "timeline": self._build_timeline_slide,
            "blank": self._build_blank_slide,
        }

        logger.info("Building %d slides", len(slides))

        for i, slide in enumerate(slides):
            builder = builders.get(slide.type)
            if builder is None:  # pragma: no cover - schema forbids it
                raise ValueError(f"No builder for slide type {slide.type!r} at slide {i}")

            try:
                logger.debug("Building slide %d: type=%s", i, slide.type)
                builder(slide, i)
            except Exception as e:
                logger.error("Failed to create slide %d: %s", i, e)
                raise ValueError(f"Error creating slide {i} ({slide.type}): {e}")

    # -------------------------------------------------------------------------
    # Slide Builders
    # -------------------------------------------------------------------------

    def _build_title_slide(self, slide_data, index: int) -> None:
        """Build a title slide with title and optional subtitle."""
        slide = self._new_slide(slide_data, index)

        self._apply_title(slide, slide_data.title, index)

        subtitle = self._placeholder_of_type(
            slide, (PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)
        )
        if subtitle is not None:
            write_text(subtitle.text_frame, slide_data.subtitle or "")
        elif slide_data.subtitle:
            self._warn(index, W.SUBTITLE_DROPPED,
                       "this layout has no subtitle placeholder; the subtitle was dropped.")

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_section_slide(self, slide_data, index: int) -> None:
        """Build a section divider slide."""
        slide = self._new_slide(slide_data, index)
        self._apply_title(slide, slide_data.title, index)
        self._add_speaker_notes(slide, slide_data.notes)

    def _build_content_slide(self, slide_data, index: int) -> None:
        """Build a content slide with bullet points."""
        slide = self._new_slide(slide_data, index)
        self._apply_title(slide, slide_data.title, index)

        bullets = body_to_bullets(slide_data.body)
        if bullets:
            placeholders = self._content_placeholders(slide)
            if placeholders:
                placeholder = placeholders[0]
                placeholder.text = ""
                self._fill_bullets(placeholder.text_frame, bullets)
                self._fit_text(placeholder, bullets, index)
            else:
                self._warn(index, W.BULLETS_DROPPED,
                           "this layout has no body placeholder; the bullets were dropped.")

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_table_slide(self, slide_data, index: int) -> None:
        """Build a table slide with a styled table."""
        slide, left, top, width, height = self._content_slide(slide_data, index)

        rows, col_alignments = parse_table_data(slide_data.rows)
        if not rows:
            self._warn(index, W.TABLE_EMPTY, "table has no rows; the slide is empty.")
            return

        # An explicit `align` beats an inline markdown separator row.
        if slide_data.align:
            col_alignments = [
                {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[a]
                for a in slide_data.align
            ]

        # Registry key is header_fill (a theme name or hex); the slide field is
        # header_color. Accept either spelling from the template.
        table_defaults = dict(self._table_defaults)
        if "header_color" in table_defaults and "header_fill" not in table_defaults:
            table_defaults["header_fill"] = table_defaults["header_color"]
        header_fill = resolve_fill(
            self._setting(slide_data, "header_color", table_defaults, "header_fill", None),
            TABLE_HEADER_FILL,
        )
        zebra = self._coerce_bool(
            self._setting(slide_data, "zebra", table_defaults, "zebra", True),
            "table zebra", index, fallback=True,
        )
        font_size = self._setting(slide_data, "font_size", table_defaults, "font_size", None)
        if font_size is not None:
            font_size = self._coerce_font_size(font_size, index)

        num_cols = max((len(row) for row in rows), default=0)
        _, points = self._create_styled_table(
            slide,
            rows,
            left=left,
            top=top,
            width=width,
            height=height,
            header_color=header_fill,
            alternate_rows=zebra,
            column_alignments=col_alignments,
            font_size=font_size,
            column_widths=self._table_widths(slide_data.widths, num_cols, index),
            cell_fills=self._table_fills(slide_data.fills, len(rows), num_cols, index),
            merges=self._table_merges(slide_data.merges, len(rows), num_cols, index),
        )

        if points and table_overflows(len(rows), height, points):
            self._warn(
                index, W.TABLE_OVERFLOW,
                f"table of {len(rows)} rows will not fit the content area even at "
                f"{points}pt; split it across slides.",
            )
        elif font_size is None and points and points < int(DEFAULT_BODY_FONT_SIZE.pt):
            self._warn(
                index, W.TABLE_FONT_REDUCED,
                f"table font reduced to {points}pt to fit {len(rows)} rows.",
            )

        self._add_speaker_notes(slide, slide_data.notes)

    def _table_widths(self, widths, num_cols: int, index: int):
        """Column weights, or None when they do not describe this table.

        One weight per column: a list of a different length cannot be matched
        to columns without guessing which ones the caller meant, so it is
        reported and the table stays evenly divided.
        """
        if not widths:
            return None
        if len(widths) != num_cols:
            self._warn(
                index, W.TABLE_WIDTHS_IGNORED,
                f"widths has {len(widths)} value(s) for {num_cols} column(s); "
                "the columns were left equal.",
            )
            return None
        return list(widths)

    def _table_fills(self, fills, num_rows: int, num_cols: int, index: int):
        """``(row, col, colour)`` triples for the cells that exist."""
        if not fills:
            return None

        applied, skipped = [], []
        for fill in fills:
            if fill.row >= num_rows or (fill.col is not None and fill.col >= num_cols):
                skipped.append(fill)
                continue
            applied.append((fill.row, fill.col, resolve_fill(fill.color, TABLE_ALT_ROW_FILL)))

        if skipped:
            self._warn(
                index, W.TABLE_FILL_IGNORED,
                f"{len(skipped)} fill(s) name a cell outside a "
                f"{num_rows}x{num_cols} table and were ignored.",
            )
        return applied or None

    def _table_merges(self, merges, num_rows: int, num_cols: int, index: int):
        """``(row, col, row_span, col_span)`` blocks that fit and do not overlap.

        Overlaps are rejected here rather than by python-pptx, which raises
        halfway through and would take the whole deck with it. The first block
        to claim a cell keeps it, so a caller sees which one was dropped.
        """
        if not merges:
            return None

        claimed = set()
        applied, skipped = [], []
        for merge in merges:
            cells = {
                (row, col)
                for row in range(merge.row, merge.row + merge.row_span)
                for col in range(merge.col, merge.col + merge.col_span)
            }
            outside = merge.row + merge.row_span > num_rows or merge.col + merge.col_span > num_cols
            if outside or cells & claimed:
                skipped.append(merge)
                continue
            claimed |= cells
            applied.append((merge.row, merge.col, merge.row_span, merge.col_span))

        if skipped:
            self._warn(
                index, W.TABLE_MERGE_IGNORED,
                f"{len(skipped)} merge(s) fell outside a {num_rows}x{num_cols} table "
                "or overlapped another merge, and were ignored.",
            )
        return applied or None

    def _build_image_slide(self, slide_data, index: int) -> None:
        """Build a slide with an image from a URL or inline data URI.

        On a template with a picture layout the image goes into the layout's
        own PICTURE placeholder, which frames, crops and positions it the way
        the designer meant; the text goes into that layout's body placeholder.
        Only when the template has no such layout — or when the one it has
        cannot hold this slide's text — does the builder compute a rectangle
        and scale the picture into it (#120).
        """
        slide = self._new_slide(slide_data, index)
        self._apply_title(slide, slide_data.title, index)

        caption = slide_data.caption
        bullets = body_to_bullets(slide_data.body)

        picture_placeholder = self._placeholder_of_type(slide, {PP_PLACEHOLDER.PICTURE})
        if picture_placeholder is not None:
            body_placeholder = max(
                self._content_placeholders(slide),
                key=lambda ph: (ph.width or 0) * (ph.height or 0),
                default=None,
            )
            if body_placeholder is not None or not (bullets or caption):
                self._fill_picture_layout(
                    slide, slide_data, index,
                    picture_placeholder, body_placeholder, bullets, caption,
                )
                return
            # Text with nowhere to go: fall through to the computed rectangle
            # rather than dropping it, and take the empty picture placeholder
            # with us so it does not show its prompt.
            self._remove_placeholder(picture_placeholder)

        slide, left, top, width, height = self._add_title_content_slide("", slide=slide)

        # With text beside it the picture takes the left 55%, the text the rest.
        image_width = width
        if bullets:
            image_width = int(width * 0.55)
            gutter = Inches(0.3)
            text_left = left + image_width + gutter
            text_width = width - image_width - gutter
            self._add_bulleted_textbox(slide, bullets, text_left, top, text_width, height)

        max_height = height - (Inches(0.6) if caption else 0)

        picture, error = self._add_image(
            slide, slide_data.source,
            left=left, top=top, max_width=image_width, max_height=max_height,
            center_horizontal=not bullets,
        )

        if picture and caption:
            self._add_text_box(
                slide, caption,
                left=picture.left,
                top=picture.top + picture.height + Inches(0.1),
                width=picture.width,
                height=Inches(0.5),
                font_size=DEFAULT_CAPTION_FONT_SIZE,
                italic=True,
                alignment=PP_ALIGN.CENTER,
            )
        elif not picture:
            self._add_image_placeholder(
                slide, "Image could not be loaded", left, top + Inches(1), width
            )
            self._warn(index, W.IMAGE_FAILED,
                       f"image could not be loaded ({error}); a placeholder was drawn.")

        self._add_speaker_notes(slide, slide_data.notes)

    def _fill_picture_layout(self, slide, slide_data, index: int,
                             picture_placeholder, body_placeholder,
                             bullets, caption) -> None:
        """Fill a picture layout's own placeholders."""
        picture, error = self._fill_picture_placeholder(
            picture_placeholder, slide_data.source
        )
        if not picture:
            # The placeholder is gone either way — insert_picture() replaces it
            # on success, and an empty one would show its prompt on failure.
            left, top, width = (picture_placeholder.left, picture_placeholder.top,
                                picture_placeholder.width)
            self._remove_placeholder(picture_placeholder)
            self._add_image_placeholder(slide, "Image could not be loaded", left, top, width)
            self._warn(index, W.IMAGE_FAILED,
                       f"image could not be loaded ({error}); a placeholder was drawn.")

        if body_placeholder is None:
            self._add_speaker_notes(slide, slide_data.notes)
            return

        if bullets or caption:
            frame = body_placeholder.text_frame
            if bullets:
                self._fill_bullets(frame, bullets)
            if caption:
                # The caption follows the bullets in the same frame: on a
                # picture layout the body placeholder *is* the caption area,
                # and a text box under the picture would sit on the layout.
                self._add_caption_paragraph(frame, caption, after_bullets=bool(bullets))
            # Through _fit_text, like every other placeholder-filled body: a
            # picture layout's text area is the smallest one a template
            # offers, so it is the likeliest to overflow and the one a caller
            # most needs warned about. The caption counts towards the fill;
            # it is set smaller than body text, so the estimate is generous.
            measured = bullets + ([Bullet(text=caption)] if caption else [])
            self._fit_text(body_placeholder, measured, index)
        else:
            self._remove_placeholder(body_placeholder)

        self._add_speaker_notes(slide, slide_data.notes)

    @staticmethod
    def _add_caption_paragraph(frame, caption: str, after_bullets: bool) -> None:
        """Write *caption* into *frame*, italic, after any bullets already there."""
        paragraph = frame.add_paragraph() if after_bullets else frame.paragraphs[0]
        write_text(paragraph, caption, font_size=DEFAULT_CAPTION_FONT_SIZE, italic=True)

    def _build_two_column_slide(self, slide_data, index: int) -> None:
        """Build a slide with two text columns.

        Columns are matched to placeholders by geometry, never by placeholder
        ``idx``: the indices PowerPoint's own Two Content and Comparison
        layouts use are a convention a corporate template need not follow, and
        addressing by number silently wrote one column into the other's box and
        dropped the rest (#195). :func:`~pptx_tools.placeholder_style.content_columns`
        reads left-to-right columns and each column's heading strip instead.

        A template that reserves fewer columns than the slide has still keeps
        every word: the columns merge into the one body, with each heading as a
        bold lead line. Whatever it costs is reported.
        """
        sources = (slide_data.left, slide_data.right)

        slide = self._new_slide(slide_data, index)
        self._apply_title(slide, slide_data.title, index)

        columns = content_columns(slide)

        if not columns:
            if any(column.heading or column.body for column in sources):
                self._warn(index, W.COLUMN_DROPPED,
                           "this layout has no body placeholder; both columns "
                           "were dropped.")
            self._add_speaker_notes(slide, slide_data.notes)
            return

        if len(columns) == 1:
            merged: List[Bullet] = []
            for column in sources:
                merged.extend(self._column_bullets(column))
            if merged:
                _, body = columns[0]
                self._fill_bullets(body.text_frame, merged)
                self._fit_text(body, merged, index)
            self._warn(index, W.COLUMNS_MERGED,
                       "this layout reserves one content area, not two; the "
                       "columns were merged into it in order.")
            self._add_speaker_notes(slide, slide_data.notes)
            return

        for column, (heading_placeholder, body) in zip(sources, columns):
            bullets = body_to_bullets(column.body)
            if column.heading and heading_placeholder is None:
                bullets = self._column_bullets(column)
                self._warn(index, W.HEADING_INLINED,
                           f"this layout has no heading placeholder; the heading "
                           f"{column.heading!r} was written as a bold first line.")
            elif column.heading:
                write_text(heading_placeholder.text_frame, column.heading)
            if bullets:
                self._fill_bullets(body.text_frame, bullets)
                self._fit_text(body, bullets, index)

        self._add_speaker_notes(slide, slide_data.notes)

    @staticmethod
    def _column_bullets(column) -> List[Bullet]:
        """A column's bullets with its heading folded in as a bold lead line.

        What a column becomes when there is no heading strip to write it into.
        """
        bullets = list(body_to_bullets(column.body))
        if column.heading:
            bullets.insert(0, Bullet(text=f"**{column.heading}**"))
        return bullets

    def _build_chart_slide(self, slide_data, index: int) -> None:
        """Build a slide with a category chart."""
        slide, left, top, width, height = self._content_slide(slide_data, index)

        bullets = body_to_bullets(slide_data.body)
        if bullets:
            chart_width = int(width * 0.6)
            gutter = Inches(0.3)
            text_left = left + chart_width + gutter
            text_width = width - chart_width - gutter
            self._add_bulleted_textbox(slide, bullets, text_left, top, text_width, height)
            width = chart_width

        chart_data = {
            "categories": slide_data.categories,
            "series": [
                {"name": s.name, "values": s.values} for s in slide_data.series
            ],
        }

        try:
            chart = add_chart_to_slide(
                slide,
                chart_type=slide_data.chart_type,
                chart_data=chart_data,
                left=left, top=top, width=width, height=height,
                has_legend=slide_data.legend != "none",
                legend_position=slide_data.legend,
                title=slide_data.chart_title,
            )
            data_labels = self._coerce_bool(
                self._setting(slide_data, "data_labels", self._chart_defaults, "data_labels", False),
                "chart data_labels", index, fallback=False,
            )
            configure_data_labels(chart, data_labels, slide_data.number_format)
            set_axis_titles(chart, slide_data.x_title, slide_data.y_title)
            self._paint_chart(chart)
        except ChartDataError as e:
            logger.error(f"Chart error: {e}")
            self._add_text_box(
                slide, f"[Chart error: {e}]",
                left, top, width, Inches(1), alignment=PP_ALIGN.CENTER,
            )
            self._warn(index, W.CHART_FAILED, f"chart could not be built ({e}).")
            self._add_speaker_notes(slide, slide_data.notes)
            return

        # A series whose length disagrees with the categories is accepted by
        # python-pptx but renders with gaps, so say so rather than let it pass.
        for series in slide_data.series:
            if len(series.values) != len(slide_data.categories):
                self._warn(
                    index, W.CHART_SERIES_LENGTH,
                    f"series {series.name!r} has {len(series.values)} values for "
                    f"{len(slide_data.categories)} categories.",
                )

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_scatter_slide(self, slide_data, index: int) -> None:
        """Build a slide with an XY (scatter) chart."""
        slide, left, top, width, height = self._content_slide(slide_data, index)

        try:
            self._paint_chart(add_scatter_to_slide(
                slide,
                series=slide_data.series,
                left=left, top=top, width=width, height=height,
                legend=slide_data.legend,
                title=slide_data.chart_title,
                x_title=slide_data.x_title,
                y_title=slide_data.y_title,
            ))
        except ChartDataError as e:
            logger.error(f"Scatter chart error: {e}")
            self._add_text_box(
                slide, f"[Chart error: {e}]",
                left, top, width, Inches(1), alignment=PP_ALIGN.CENTER,
            )
            self._warn(index, W.CHART_FAILED, f"scatter chart could not be built ({e}).")

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_quote_slide(self, slide_data, index: int) -> None:
        """Build a quote/citation slide."""
        slide, left, top, width, height = self._content_slide(slide_data, index)

        quote_box = slide.shapes.add_textbox(left, top, width, height)
        tf = quote_box.text_frame
        tf.word_wrap = True

        para = tf.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        formatted_quote = f'"{slide_data.text}"'
        write_text(para, formatted_quote, font_size=DEFAULT_QUOTE_FONT_SIZE, italic=True)

        if slide_data.attribution:
            author_para = tf.add_paragraph()
            write_text(
                author_para, f"— {slide_data.attribution}",
                font_size=DEFAULT_SUBTITLE_FONT_SIZE, bold=True, alignment=PP_ALIGN.CENTER,
            )
            author_para.space_before = Pt(24)

        self._paint(tf)
        self._add_speaker_notes(slide, slide_data.notes)

    # -------------------------------------------------------------------------
    # Slide types that draw their own shapes
    # -------------------------------------------------------------------------

    def _build_kpi_slide(self, slide_data, index: int) -> None:
        """Build a row of headline figures.

        Drawn rather than placed in a body placeholder: the point of a KPI row
        is the typographic contrast between a large figure and a small label,
        which a bullet list cannot express.
        """
        slide, left, top, width, height = self._content_slide(slide_data, index)

        items = slide_data.items
        gutter = Inches(0.25)
        cell_width = int((width - gutter * (len(items) - 1)) / len(items))
        # Sit the row a little above centre; a KPI slide reads as a headline.
        block_height = min(height, Inches(2.4))
        block_top = top + int(max(0, (height - block_height)) / 3)

        for position, item in enumerate(items):
            cell_left = left + position * (cell_width + gutter)
            box = slide.shapes.add_textbox(cell_left, block_top, cell_width, block_height)
            frame = box.text_frame
            frame.word_wrap = True

            value = frame.paragraphs[0]
            write_text(value, item.value, font_size=KPI_VALUE_FONT_SIZE,
                       bold=True, alignment=PP_ALIGN.CENTER)

            label = frame.add_paragraph()
            write_text(label, item.label, font_size=DEFAULT_CAPTION_FONT_SIZE,
                       alignment=PP_ALIGN.CENTER)
            label.space_before = Pt(4)

            if item.delta:
                delta = frame.add_paragraph()
                write_text(delta, item.delta, font_size=DEFAULT_CAPTION_FONT_SIZE,
                           italic=True, alignment=PP_ALIGN.CENTER)
                delta.space_before = Pt(2)
                # Theme accent, so the figure follows the template's palette.
                delta.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1

            self._paint(frame)

        if len(items) > 4:
            self._warn(
                index, W.KPI_CROWDED,
                f"{len(items)} figures on one KPI slide will be cramped; "
                "two to four read best.",
            )

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_agenda_slide(self, slide_data, index: int) -> None:
        """Build a numbered agenda.

        With no explicit items the entries are taken from the deck's own
        ``section`` slides, in order — which is the usual case and cannot drift
        out of step with the deck the way a hand-written list does.
        """
        items = slide_data.items
        derived = False
        if items is None:
            items = [s.title for s in self.slides if s.type == "section" and s.title]
            derived = True

        if not items:
            self._warn(
                index, W.AGENDA_EMPTY,
                "agenda has no items and the deck has no section slides to derive "
                "them from; the slide is empty.",
            )

        slide = self._new_slide(slide_data, index)
        self._apply_title(slide, slide_data.title or "Agenda", index)

        if items:
            bullets = [Bullet(text=f"{n}.  {text}") for n, text in enumerate(items, 1)]
            placeholders = self._content_placeholders(slide)
            if placeholders:
                placeholders[0].text = ""
                self._fill_bullets(placeholders[0].text_frame, bullets)
                self._fit_text(placeholders[0], bullets, index)
            else:
                self._warn(index, W.AGENDA_DROPPED,
                           "this layout has no body placeholder; the agenda was dropped.")

        if derived and items:
            logger.debug("Agenda derived from %d section slides", len(items))

        self._add_speaker_notes(slide, slide_data.notes)

    def _build_closing_slide(self, slide_data, index: int) -> None:
        """Build a closing / thank-you slide."""
        slide = self._new_slide(slide_data, index)
        self._apply_title(slide, slide_data.title or "Thank you", index)

        lines = []
        if slide_data.subtitle:
            lines.append(slide_data.subtitle)
        lines.extend(slide_data.contact or [])

        if lines:
            target = self._placeholder_of_type(
                slide, (PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)
            )
            if target is not None:
                target.text = ""
                self._fill_bullets(target.text_frame, [Bullet(text=line) for line in lines])
            else:
                self._warn(
                    index, W.CLOSING_LINES_DROPPED,
                    "this layout has no subtitle or body placeholder; the closing "
                    "lines were dropped.",
                )

        self._add_speaker_notes(slide, slide_data.notes)

    # -------------------------------------------------------------------------
    # Blank slide with positioned elements
    # -------------------------------------------------------------------------

    _ELEMENT_SHAPES = {
        "rectangle": MSO_SHAPE.RECTANGLE,
        "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
        "ellipse": MSO_SHAPE.OVAL,
        "chevron": MSO_SHAPE.CHEVRON,
        "arrow": MSO_SHAPE.RIGHT_ARROW,
    }
    _ELEMENT_ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}

    @staticmethod
    def _resolve_length(value, extent: int) -> int:
        """A schema position — inches, "1.5in" or "40%" — as EMU along *extent*."""
        if isinstance(value, (int, float)):
            return int(Inches(value))
        text = str(value).strip()
        if text.endswith("%"):
            return int(extent * float(text[:-1]) / 100.0)
        if text.endswith("in"):
            text = text[:-2]
        return int(Inches(float(text)))

    @staticmethod
    def _fill_shape(shape, fill) -> None:
        """Apply a resolved fill: an ``RGBColor``, or a theme colour name."""
        shape.fill.solid()
        if isinstance(fill, str):
            # "accent1" -> ACCENT_1, "dark2" -> DARK_2: the enum spells the
            # digit with an underscore.
            member = re.sub(r"(\d)$", r"_\1", fill.upper())
            shape.fill.fore_color.theme_color = getattr(MSO_THEME_COLOR, member)
        else:
            shape.fill.fore_color.rgb = fill

    def _build_blank_slide(self, slide_data, index: int) -> None:
        """Draw positioned elements on an empty layout.

        The escape hatch: every other slide type decides where things go, this
        one is told. Positions are resolved against the real slide size, so a
        percentage means the same thing on a 4:3 and a 16:9 template. An
        element that would run past the slide edge is shrunk to fit and
        reported; one that starts off the slide is skipped and reported.
        Silently drawing off-slide content is the failure the warnings channel
        exists to prevent.
        """
        slide = self._new_slide(slide_data, index)
        self._apply_title(slide, slide_data.title, index)
        slide_w = self.presentation.slide_width
        slide_h = self.presentation.slide_height

        for number, element in enumerate(slide_data.elements, start=1):
            left = self._resolve_length(element.x, slide_w)
            top = self._resolve_length(element.y, slide_h)
            width = self._resolve_length(element.w, slide_w)
            if element.h is not None:
                height = self._resolve_length(element.h, slide_h)
            elif element.kind == "text":
                height = int(Inches(1))
            elif element.kind == "shape":
                height = width
            else:
                height = slide_h - top

            if left >= slide_w or top >= slide_h:
                self._warn(index, W.ELEMENT_OFF_SLIDE,
                           f"element {number} ({element.kind}) starts off the slide; skipped.")
                continue
            reduced = []
            if left + width > slide_w:
                width, reduced = slide_w - left, reduced + ["width"]
            if top + height > slide_h:
                height, reduced = slide_h - top, reduced + ["height"]
            if reduced:
                self._warn(
                    index, W.ELEMENT_CLAMPED,
                    f"element {number} ({element.kind}) ran past the slide edge; "
                    f"{' and '.join(reduced)} reduced to fit.",
                )
            # Clamping cannot reach zero — the off-slide skip above fires first —
            # so a zero here is what the caller asked for: a 0-wide text box or
            # a 1-EMU picture that draws nothing and says nothing.
            if width <= 0 or height <= 0:
                self._warn(index, W.ELEMENT_ZERO_SIZE,
                           f"element {number} ({element.kind}) has no size; skipped.")
                continue

            if element.kind == "text":
                box = slide.shapes.add_textbox(left, top, width, height)
                box.text_frame.word_wrap = True
                write_text(
                    box.text_frame, element.text,
                    font_size=Pt(element.font_size) if element.font_size else None,
                    bold=element.bold,
                    alignment=self._ELEMENT_ALIGN[element.align],
                )
                self._paint(box.text_frame)
            elif element.kind == "image":
                picture, error = self._add_image(
                    slide, element.source,
                    left=left, top=top, max_width=width, max_height=height,
                    center_horizontal=False,
                )
                if not picture:
                    self._add_image_placeholder(slide, "Image could not be loaded", left, top, width)
                    self._warn(
                        index, W.IMAGE_FAILED,
                        f"element {number} image could not be loaded ({error}); "
                        "a placeholder was drawn.",
                    )
            else:
                shape = slide.shapes.add_shape(
                    self._ELEMENT_SHAPES[element.shape], left, top, width, height
                )
                fill = resolve_fill(element.fill, None)
                if fill is not None:
                    self._fill_shape(shape, fill)
                if element.text:
                    shape.text_frame.word_wrap = True
                    write_text(shape.text_frame, element.text, alignment=PP_ALIGN.CENTER)

        self._add_speaker_notes(slide, slide_data.notes)

    def _timeline_detail_band(self, steps, height, index: int):
        """Height to reserve under the shapes for step detail, as (gap, height).

        Returns ``(0, 0)`` when no step has detail, or when the content
        rectangle is too short to carry a legible caption — in which case the
        detail is dropped and reported, rather than drawn off the slide.
        """
        if not any(step.detail for step in steps):
            return 0, 0

        wanted = TIMELINE_DETAIL_GAP + TIMELINE_DETAIL_HEIGHT
        if height - wanted >= TIMELINE_STEP_MIN_HEIGHT:
            return TIMELINE_DETAIL_GAP, TIMELINE_DETAIL_HEIGHT

        self._warn(
            index, W.TIMELINE_DETAIL_DROPPED,
            "the content area is too short to fit step detail under the timeline "
            "shapes; the detail lines were dropped. Use a layout with a taller "
            "body area, or move the detail into speaker notes.",
        )
        return 0, 0

    def _build_timeline_slide(self, slide_data, index: int) -> None:
        """Build a row of steps as chevrons or boxes.

        Autoshapes rather than SmartArt: python-pptx cannot create SmartArt, and
        a row of chevrons carries the same "these follow one another" meaning
        without the dependency.
        """
        slide, left, top, width, height = self._content_slide(slide_data, index)

        steps = slide_data.steps
        shape_type = MSO_SHAPE.CHEVRON if slide_data.style == "chevron" else MSO_SHAPE.ROUNDED_RECTANGLE
        # Chevrons interlock, so they overlap slightly; boxes get a gutter.
        gutter = Inches(-0.12) if slide_data.style == "chevron" else Inches(0.15)

        step_width = int((width - gutter * (len(steps) - 1)) / len(steps))

        # Reserve the detail band up front. Sizing the shapes first and then
        # hanging the captions underneath overflowed the content rectangle on
        # any layout with a short body placeholder — an explicit ``layout``
        # override onto a section-header layout put the detail lines 0.54in
        # below the box, and a custom template with a low, short placeholder
        # would put them off the slide entirely.
        gap, detail_height = self._timeline_detail_band(steps, height, index)
        band = gap + detail_height
        step_height = min(TIMELINE_STEP_HEIGHT, max(TIMELINE_STEP_MIN_HEIGHT, height - band))
        step_top = top + int(max(0, (height - (step_height + band))) / 3)

        for position, step in enumerate(steps):
            shape = slide.shapes.add_shape(
                shape_type,
                left + position * (step_width + gutter),
                step_top, step_width, step_height,
            )
            shape.fill.solid()
            shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
            shape.line.fill.background()

            frame = shape.text_frame
            frame.word_wrap = True
            label = frame.paragraphs[0]
            write_text(label, step.label, font_size=DEFAULT_CAPTION_FONT_SIZE,
                       bold=True, alignment=PP_ALIGN.CENTER)

            if step.detail and detail_height:
                # Detail below the shape, so a long line cannot burst the chevron.
                caption = slide.shapes.add_textbox(
                    left + position * (step_width + gutter),
                    step_top + step_height + gap,
                    step_width, detail_height,
                )
                caption.text_frame.word_wrap = True
                detail = caption.text_frame.paragraphs[0]
                write_text(detail, step.detail, font_size=TIMELINE_DETAIL_FONT_SIZE,
                           alignment=PP_ALIGN.CENTER)
                self._paint(caption.text_frame)

        self._add_speaker_notes(slide, slide_data.notes)

    # -------------------------------------------------------------------------
    # Fit, language, footer
    # -------------------------------------------------------------------------

    def _add_bulleted_textbox(self, slide, bullets, left, top, width, height):
        """Bullets in a text box, looking like bullets.

        A text box is not a placeholder, so it inherits its paragraph
        formatting from the presentation's default text style, which has no
        bullet glyphs: the same markdown that bulleted correctly in a content
        placeholder came out as plain lines beside a picture or a chart
        (#123). The master's body style is applied explicitly instead.
        """
        box = slide.shapes.add_textbox(left, top, width, height)
        self._fill_bullets(box.text_frame, bullets)
        apply_list_style(box.text_frame, slide.slide_layout.slide_master)
        self._paint(box.text_frame)
        apply_autofit(box.text_frame, scale=self._fit_scale(bullets, width, height))
        return box

    def _fit_scale(self, bullets, width, height):
        """Shrink factor for a text box, or None when the text already fits.

        Measured at the master's body size, because that is what
        :func:`apply_list_style` gives these boxes — not the built-in default.
        """
        fill = estimate_text_fill(
            bullets, width, height, font_size_pt=self._master_body_font_size(),
            typeface=self._typeface,
        )
        return (1.0 / fill) if fill > 1.0 else None

    def _paint(self, target) -> None:
        """Give text the builder drew the colour this template uses for body text."""
        apply_text_color(target, self._body_color)

    def _paint_chart(self, chart) -> None:
        """Give a chart's axis labels, legend and title the template's text colour.

        Chart text lives in its own part and inherits nothing from the slide,
        so it came out `tx1` — black — whatever the deck looked like. Setting
        it on ``chart.font`` reaches every label that does not override it.
        """
        if chart is None or self._body_color is None:
            return
        try:
            if isinstance(self._body_color, RGBColor):
                chart.font.color.rgb = self._body_color
            else:
                chart.font.color.theme_color = self._body_color
        except (AttributeError, ValueError) as error:  # pragma: no cover
            logger.debug("Could not colour chart text: %s", error)

    def _master_body_font_size(self) -> float:
        """The template's own body size, read once per deck."""
        if self._body_size is None:
            master = self.presentation.slide_masters[0] if self.presentation.slide_masters else None
            size = read_master_body_font_size(master)
            self._body_size = size if size else float(DEFAULT_BODY_FONT_SIZE.pt)
        return self._body_size

    def _body_font_size(self, placeholder) -> float:
        """The size body text in *placeholder* really renders at, in points.

        Measuring against ``DEFAULT_BODY_FONT_SIZE`` regardless of the template
        made the fit estimate wrong by the square of the ratio. Both templates
        this server ships set 28pt, not 18: a body needing 1.9x its box
        measured as 0.86x, so no shrink factor was written and the overflow
        warning — computed from the same number — never fired either (#195).
        """
        size = read_body_font_size(placeholder)
        return size if size else float(DEFAULT_BODY_FONT_SIZE.pt)

    def _fit_text(self, placeholder, bullets, index: int) -> None:
        """Ask PowerPoint to shrink overfull body text, and warn when it is far gone."""
        fill = estimate_text_fill(
            bullets, placeholder.width, placeholder.height,
            font_size_pt=self._body_font_size(placeholder),
            typeface=self._typeface,
        )
        apply_autofit(placeholder.text_frame, scale=(1.0 / fill) if fill > 1.0 else None)

        # Below the shrink floor the slide is genuinely overfull; shrinking
        # further would produce text nobody can read from a room.
        if fill > 1.0 / 0.6:
            self._warn(
                index, W.TEXT_OVERFLOW,
                f"body text is about {fill:.1f}x the space available and will be "
                "shrunk to fit; consider splitting it across slides.",
            )

    def _apply_language(self, language: str) -> None:
        """Stamp the proofing language on every run in the deck."""
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    set_runs_language(shape.text_frame, language)
                elif getattr(shape, "has_table", False):
                    for row in shape.table.rows:
                        for cell in row.cells:
                            set_runs_language(cell.text_frame, language)
            if slide.has_notes_slide:
                set_runs_language(slide.notes_slide.notes_text_frame, language)

        logger.debug("Applied language %s to all runs", language)

    def _apply_footer_and_slide_numbers(self) -> None:
        """Apply footer text and/or slide numbers to all slides.

        Clones the footer/slide-number placeholder shapes from each slide's
        layout into the slide itself so they become visible. Assigns unique
        shape IDs to avoid PPTX corruption.
        """
        from xml.sax.saxutils import escape as xml_escape

        missing_footer = 0

        for slide in self.presentation.slides:
            layout = slide.slide_layout
            spTree = slide.shapes._spTree

            # Determine next available shape ID on this slide by collecting the
            # id attributes of the tree's direct children.
            existing_ids = set()
            for sp in spTree:
                cNvPr = sp.find('.//' + qn('p:cNvPr'))
                if cNvPr is None:
                    cNvPr = sp.find('.//' + qn('p:nvSpPr') + '/' + qn('p:cNvPr'))
                if cNvPr is not None and cNvPr.get('id'):
                    existing_ids.add(int(cNvPr.get('id')))
            next_id = max(existing_ids, default=0) + 1

            layout_indices = {ph.placeholder_format.idx for ph in layout.placeholders}
            if self._footer_text and 11 not in layout_indices:
                missing_footer += 1

            for ph in layout.placeholders:
                idx = ph.placeholder_format.idx

                if idx == 11 and self._footer_text:  # FOOTER placeholder
                    sp = copy.deepcopy(ph._element)
                    cNvPr = sp.find(qn('p:nvSpPr') + '/' + qn('p:cNvPr'))
                    if cNvPr is not None:
                        cNvPr.set('id', str(next_id))
                        next_id += 1
                    txBody = sp.find(qn('p:txBody'))
                    if txBody is not None:
                        for p in txBody.findall(qn('a:p')):
                            txBody.remove(p)
                        ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                        safe_text = xml_escape(self._footer_text)
                        p_xml = f'<a:p {ns}><a:r><a:t>{safe_text}</a:t></a:r></a:p>'
                        txBody.append(parse_xml(p_xml))
                    spTree.append(sp)

                elif idx == 12 and self._show_slide_numbers:  # SLIDE_NUMBER placeholder
                    sp = copy.deepcopy(ph._element)
                    cNvPr = sp.find(qn('p:nvSpPr') + '/' + qn('p:cNvPr'))
                    if cNvPr is not None:
                        cNvPr.set('id', str(next_id))
                        next_id += 1
                    spTree.append(sp)

        if missing_footer:
            # Previously this just produced a deck with no footer and no
            # explanation of why the argument appeared to do nothing.
            self._warn_deck(
                W.FOOTER_UNSUPPORTED,
                f"footer_text was dropped on {missing_footer} slide(s): their layout "
                "has no footer placeholder.",
            )

        logger.debug("Applied footer/slide numbers to all slides")

    # -------------------------------------------------------------------------
    # Output
    # -------------------------------------------------------------------------

    def save(self) -> io.BytesIO:
        """Save presentation to a BytesIO object.

        Returns:
            BytesIO containing the presentation.

        Raises:
            RuntimeError: If saving fails.
        """
        logger.info("Saving PowerPoint to memory buffer")
        try:
            buffer = io.BytesIO()
            self.presentation.save(buffer)
            buffer.seek(0)
            return buffer
        except Exception as e:
            logger.error("Failed to save PowerPoint presentation: %s", e, exc_info=True)
            raise RuntimeError(f"Failed to save presentation: {e}") from e
