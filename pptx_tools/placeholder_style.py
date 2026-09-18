"""Read what a template's placeholders say, and replay a title as a text box.

Some layouts carry no title placeholder at all — every template's Blank
layout, and any full-bleed layout a designer trimmed. Slides built on them
used to drop their ``title`` and report it (#118), which made ``title`` a
field the schema accepts, validates and then discards.

This module supplies the alternative: read where the template puts a title
and how it styles one, then draw the text there as a plain text box. What it
reads, in this order:

* **geometry** from a title placeholder elsewhere in the template —
  python-pptx already resolves a layout placeholder's position through the
  master, so a layout that defines no ``<a:xfrm>`` still answers;
* **character style** from the first ``<a:defRPr>`` found on the layout's
  title placeholder, then the master's, then the master's
  ``<p:titleStyle>``. That element is copied wholesale onto the run rather
  than re-derived attribute by attribute, so theme references (``+mj-lt``,
  ``schemeClr``) survive and keep following the theme;
* **alignment and anchoring** from the same chain.

:class:`~pptx_tools.layouts.LayoutResolver` picks which layout to read from;
the builder calls :func:`draw_title_box` when ``_set_title`` finds no
placeholder to fill.

:func:`apply_list_style` answers a third: a plain text box inherits its
paragraph formatting from the presentation's default text style, which has no
bullets, so bullets drawn beside a picture or a chart came out as plain lines
while the same markdown on a `content` slide showed glyphs (#123). The
master's ``<p:bodyStyle>`` is where those glyphs are defined, and it is
applied per level.

:func:`read_content_rect` answers the neighbouring question for the body:
where this template expects content to sit. A layout with no body placeholder
— Blank, Title Only — used to fall back to a hardcoded band starting 1.5
inches down, which on a template with a header rule or a deep title band
drew straight over the decoration (#119).
"""

from __future__ import annotations

import copy
import logging
from dataclasses import dataclass
from typing import Any, Dict, Iterator, Optional

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn

logger = logging.getLogger(__name__)

TITLE_PLACEHOLDER_TYPES = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
# BODY and OBJECT are both "content goes here"; templates use them
# interchangeably. Same pair as layouts.py's signature reading.
CONTENT_PLACEHOLDER_TYPES = (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)

# Character properties that point at a relationship (an image fill, a
# hyperlink) would reference a rel that only exists in the layout's part.
# Nothing in a title style uses them; dropping them keeps a hand-built
# template from producing a slide PowerPoint has to repair.
_RELATIONSHIP_BEARING = ('a:blipFill', 'a:hlinkClick', 'a:hlinkMouseOver')

# The bullet half of CT_TextParagraphProperties, in schema order. A paragraph
# may carry one of each group; copying them as a block keeps that order.
_BULLET_TAGS = (
    'a:buClrTx', 'a:buClr',
    'a:buSzTx', 'a:buSzPct', 'a:buSzPts',
    'a:buFontTx', 'a:buFont',
    'a:buNone', 'a:buAutoNum', 'a:buChar',
)
# What may precede them, and what must follow, inside <a:pPr>.
_BEFORE_BULLETS = ('a:lnSpc', 'a:spcBef', 'a:spcAft')
_AFTER_BULLETS = ('a:tabLst', 'a:defRPr', 'a:extLst')


@dataclass(frozen=True)
class Rect:
    """A rectangle read from a template, and the layout it came from."""

    left: int
    top: int
    width: int
    height: int
    source: Optional[str] = None

    def as_tuple(self):
        return self.left, self.top, self.width, self.height


@dataclass(frozen=True)
class TitleStyle:
    """Where a template puts a title and how it dresses one."""

    left: int
    top: int
    width: int
    height: int
    defRPr: Optional[object] = None
    algn: Optional[str] = None
    anchor: Optional[str] = None
    source: Optional[str] = None      # layout name, for diagnostics


def title_placeholder(container):
    """The title placeholder of a layout, master or slide, or None."""
    for placeholder in container.placeholders:
        if placeholder.placeholder_format.type in TITLE_PLACEHOLDER_TYPES:
            return placeholder
    return None


def content_placeholders(container):
    """Body/object placeholders of a layout or slide, in document order."""
    return [placeholder for placeholder in container.placeholders
            if placeholder.placeholder_format.type in CONTENT_PLACEHOLDER_TYPES]


def read_content_rect(layout) -> Optional[Rect]:
    """The rectangle *layout* reserves for content, or None if it reserves none.

    The largest body placeholder, not the first: on a Comparison layout the
    first is the heading strip, which is the same trap ``_add_title_content_
    slide()`` avoids when the slide itself has placeholders.
    """
    placeholder = max(
        content_placeholders(layout),
        key=lambda ph: (ph.width or 0) * (ph.height or 0),
        default=None,
    )
    if placeholder is None:
        return None
    box = (placeholder.left, placeholder.top, placeholder.width, placeholder.height)
    if any(value is None for value in box):
        return None
    return Rect(*box, source=getattr(layout, "name", None))


def _master_of(layout):
    master = getattr(layout, "slide_master", None)
    return master if master is not None else layout


def _paragraph_properties(placeholder) -> Iterator[object]:
    """``<a:lvl1pPr>``/``<a:pPr>`` elements of *placeholder*, most specific first."""
    if placeholder is None:
        return
    txBody = placeholder._element.find(qn('p:txBody'))
    if txBody is None:
        return
    lstStyle = txBody.find(qn('a:lstStyle'))
    if lstStyle is not None:
        lvl1 = lstStyle.find(qn('a:lvl1pPr'))
        if lvl1 is not None:
            yield lvl1
    paragraph = txBody.find(qn('a:p'))
    if paragraph is not None:
        pPr = paragraph.find(qn('a:pPr'))
        if pPr is not None:
            yield pPr


def _style_chain(placeholder, master) -> Iterator[object]:
    """Every paragraph-property element a title inherits from, nearest first."""
    yield from _paragraph_properties(placeholder)
    master_title = title_placeholder(master) if master is not None else None
    if master_title is not None and master_title is not placeholder:
        yield from _paragraph_properties(master_title)
    if master is not None:
        txStyles = master._element.find(qn('p:txStyles'))
        if txStyles is not None:
            titleStyle = txStyles.find(qn('p:titleStyle'))
            if titleStyle is not None:
                lvl1 = titleStyle.find(qn('a:lvl1pPr'))
                if lvl1 is not None:
                    yield lvl1


def _body_properties(placeholder, master) -> Iterator[object]:
    for shape in (placeholder, title_placeholder(master) if master is not None else None):
        if shape is None:
            continue
        txBody = shape._element.find(qn('p:txBody'))
        if txBody is None:
            continue
        bodyPr = txBody.find(qn('a:bodyPr'))
        if bodyPr is not None:
            yield bodyPr


def read_title_style(layout) -> Optional[TitleStyle]:
    """The geometry and character style of *layout*'s title placeholder.

    None when the layout has no title placeholder, or when neither it nor the
    master says where the title goes — there is nothing to copy then, and the
    caller falls back to a band across the top of the slide.
    """
    placeholder = title_placeholder(layout)
    if placeholder is None:
        return None

    box = (placeholder.left, placeholder.top, placeholder.width, placeholder.height)
    if any(value is None for value in box):
        return None

    master = _master_of(layout)

    defRPr = None
    algn = None
    for properties in _style_chain(placeholder, master):
        if defRPr is None:
            candidate = properties.find(qn('a:defRPr'))
            if candidate is not None:
                defRPr = candidate
        if algn is None:
            algn = properties.get('algn')
        if defRPr is not None and algn is not None:
            break

    anchor = None
    for bodyPr in _body_properties(placeholder, master):
        anchor = bodyPr.get('anchor')
        if anchor is not None:
            break

    return TitleStyle(
        left=box[0], top=box[1], width=box[2], height=box[3],
        defRPr=defRPr, algn=algn, anchor=anchor,
        source=getattr(layout, "name", None),
    )


def body_list_levels(master) -> Dict[int, Any]:
    """The master's ``<p:bodyStyle>`` levels, keyed by 0-based paragraph level."""
    if master is None:
        return {}
    txStyles = master._element.find(qn('p:txStyles'))
    if txStyles is None:
        return {}
    bodyStyle = txStyles.find(qn('p:bodyStyle'))
    if bodyStyle is None:
        return {}

    levels: Dict[int, Any] = {}
    for depth in range(9):
        level = bodyStyle.find(qn(f'a:lvl{depth + 1}pPr'))
        if level is not None:
            levels[depth] = level
    return levels


def apply_list_style(text_frame, master) -> bool:
    """Give a plain text box the bullet glyphs and indents of a body placeholder.

    A text box inherits from the presentation's default text style, which has
    no bullets — so a body drawn beside a picture or a chart rendered as plain
    lines while the identical markdown bulleted correctly in a placeholder
    (#123). Each paragraph takes the ``marL``/``indent`` and the bullet
    definition of its own level from the master's body style.

    Only the list formatting is copied. Size, colour and line spacing stay as
    the builder set them: the body style's 28pt first level is meant for a
    full-width placeholder, not for a column beside a picture.

    Returns False when the master defines no body style, in which case the
    text box is left as it was.
    """
    levels = body_list_levels(master)
    if not levels:
        return False

    for paragraph in text_frame.paragraphs:
        pPr = paragraph._p.get_or_add_pPr()
        source = levels.get(int(pPr.get('lvl') or 0))
        if source is None:
            continue

        for name in ('marL', 'indent'):
            value = source.get(name)
            if value is not None:
                pPr.set(name, value)

        _replace_bullet_properties(pPr, source)

    return True


def _replace_bullet_properties(pPr, source) -> None:
    """Swap *pPr*'s bullet elements for *source*'s, keeping schema order."""
    before, after = [], []
    for child in list(pPr):
        tag = child.tag
        if any(tag == qn(name) for name in _BEFORE_BULLETS):
            before.append(child)
        elif any(tag == qn(name) for name in _AFTER_BULLETS):
            after.append(child)
        # Anything else is a bullet property being replaced, and is dropped.
        pPr.remove(child)

    bullets = [copy.deepcopy(child) for child in source
               if any(child.tag == qn(name) for name in _BULLET_TAGS)]

    for child in before + bullets + after:
        pPr.append(child)


def draw_title_box(slide, text: str, style: TitleStyle):
    """Add a text box holding *text*, dressed as this template's titles are."""
    box = slide.shapes.add_textbox(style.left, style.top, style.width, style.height)
    box.name = "Title"
    frame = box.text_frame
    frame.word_wrap = True
    frame.text = text

    bodyPr = frame._txBody.find(qn('a:bodyPr'))
    if bodyPr is not None and style.anchor:
        bodyPr.set('anchor', style.anchor)

    # Every paragraph, not just the first: ``TextFrame.text`` splits a title
    # containing a newline into one paragraph per line, and a second line left
    # at python-pptx's defaults would sit under the first in the wrong face
    # and half the size.
    for paragraph in frame.paragraphs:
        if style.algn:
            paragraph._p.get_or_add_pPr().set('algn', style.algn)
        if style.defRPr is None:
            continue
        for run in paragraph.runs:
            element = run._r
            existing = element.find(qn('a:rPr'))
            if existing is not None:
                element.remove(existing)
            element.insert(0, _run_properties(style.defRPr))

    return box


def _run_properties(defRPr):
    """*defRPr* as an ``<a:rPr>``, with anything part-specific removed.

    The two share a content model, so the copy is schema-valid as it stands;
    what cannot travel is a reference to a relationship, which lives in the
    layout's part and not the slide's.
    """
    rPr = copy.deepcopy(defRPr)
    rPr.tag = qn('a:rPr')
    for tag in _RELATIONSHIP_BEARING:
        for child in rPr.findall(qn(tag)):
            rPr.remove(child)
    return rPr
