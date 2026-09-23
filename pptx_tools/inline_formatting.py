"""Inline markdown formatting for PowerPoint text runs.

Renders **bold**, *italic*, ***bold italic***, ~~strikethrough~~,
__underline__, `code`, ^superscript^, ~subscript~ and [links](url) into
python-pptx runs, with nesting and backslash escapes. The grammar itself is
:mod:`inline_markdown`, shared with the Word renderer; this module is the
PowerPoint-specific rendering of its tokens.

Note on dunders: "__init__" is emphasis under standard markdown too (both
markers are flanked), so it is matched here as well. Write it inside
backticks — `__init__` — to render it verbatim; the code branch consumes the
span before any emphasis branch sees it.
"""

import logging
from typing import Optional

from inline_markdown import ESCAPE_RE, LINK_RE, build_inline_pattern, is_safe_link_target

logger = logging.getLogger(__name__)

# The grammar is shared with the Word renderer — see inline_markdown.py for
# the rules and why there is one copy. PowerPoint cannot highlight, so that
# span is left out and "==x==" stays literal; superscript and subscript are
# drawn through the run's baseline offset.
_INLINE_FORMAT_RE = build_inline_pattern(superscript=True, subscript=True)
_LINK_RE = LINK_RE
_ESCAPE_RE = ESCAPE_RE


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def has_inline_formatting(text: str) -> bool:
    """Quick check whether text contains any inline markdown markers."""
    return bool(_INLINE_FORMAT_RE.search(text))


def has_escapes(text: str) -> bool:
    """Quick check whether text contains a backslash escape."""
    return bool(_ESCAPE_RE.search(text))


def needs_inline_processing(text: str) -> bool:
    """True when *text* must go through :func:`apply_inline_formatting`.

    Callers used to test :func:`has_inline_formatting` alone and assign
    ``paragraph.text`` directly otherwise. That skipped the escape pass, so
    text carrying only an escape and no live marker (``price \\* qty``) kept its
    backslash. Escapes are now honoured whether or not the text also formats.
    """
    return has_inline_formatting(text) or has_escapes(text)


def write_text(
    target,
    text: str,
    font_size=None,
    bold: bool = False,
    italic: bool = False,
    alignment=None,
    hyperlinks: bool = True,
) -> None:
    """Write *text* into *target*, rendering inline markdown when it carries any.

    The one way to put a caller's string on a slide. The tool tells the model
    that inline markdown works in any text field, and it did not: a title, a
    caption, a KPI label or a two-column heading was assigned with
    ``paragraph.text``, which shows ``**bold**`` as four asterisks and a link
    as its markdown source. Every such assignment goes through here instead,
    so the promise holds wherever a caller can put text.

    *target* is a text frame or a paragraph. The style arguments are the
    paragraph's own defaults, applied to the plain path and inherited by the
    formatted one; leaving one out leaves that property inherited from the
    placeholder or layout, which is what an untouched title relies on.
    """
    if needs_inline_processing(text):
        apply_inline_formatting(
            target, text, font_size=font_size, bold=bold, italic=italic,
            alignment=alignment, hyperlinks=hyperlinks,
        )
        return

    paragraph = target.paragraphs[0] if hasattr(target, 'paragraphs') else target
    paragraph.text = text
    if font_size is not None:
        paragraph.font.size = font_size
    if bold:
        paragraph.font.bold = True
    if italic:
        paragraph.font.italic = True
    if alignment is not None:
        paragraph.alignment = alignment


def apply_inline_formatting(
    text_frame_or_paragraph,
    text: str,
    font_size: Optional[int] = None,
    bold: bool = False,
    italic: bool = False,
    alignment=None,
    hyperlinks: bool = True,
) -> None:
    """Parse inline markdown and render into a pptx text frame paragraph.

    If *text_frame_or_paragraph* is a TextFrame, uses its first paragraph.
    Otherwise treats it as a paragraph directly.

    Args:
        text_frame_or_paragraph: pptx TextFrame or paragraph object.
        text: Text potentially containing inline markdown.
        font_size: Optional font size to apply to all runs.
        bold: Inherited bold context.
        italic: Inherited italic context.
        alignment: Optional PP_ALIGN value for the paragraph.
        hyperlinks: False renders a link as its label without making it one.
            Chart text is the case: a hyperlink needs a relationship in the
            part that owns the run, and python-pptx cannot resolve a chart
            title's part — PowerPoint would not follow the link there anyway.
    """
    # Determine if we got a text frame or a paragraph
    if hasattr(text_frame_or_paragraph, 'paragraphs'):
        paragraph = text_frame_or_paragraph.paragraphs[0]
    else:
        paragraph = text_frame_or_paragraph

    if alignment is not None:
        paragraph.alignment = alignment

    # Handle escape sequences
    escape_ctx = {"map": {}, "counter": 0}
    text = _handle_escapes(text, escape_ctx)

    # Parse and render
    _parse_segment(text, paragraph, font_size=font_size, bold=bold, italic=italic,
                   escape_ctx=escape_ctx, hyperlinks=hyperlinks)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _handle_escapes(text: str, escape_ctx: dict) -> str:
    """Replace backslash-escaped characters with PUA placeholders."""
    def _replace(match):
        placeholder = chr(0xE000 + escape_ctx["counter"])
        escape_ctx["map"][placeholder] = match.group(1)
        escape_ctx["counter"] += 1
        return placeholder
    return _ESCAPE_RE.sub(_replace, text)


def _restore_escapes(text: str, escape_ctx: dict) -> str:
    """Replace PUA placeholders back with their original literal characters."""
    if not escape_ctx or not escape_ctx.get("map"):
        return text
    for placeholder, char in escape_ctx["map"].items():
        text = text.replace(placeholder, char)
    return text


# DrawingML expresses super/subscript as a baseline offset in thousandths of
# a percent; these are the values PowerPoint itself writes.
_SUPERSCRIPT_BASELINE = "30000"
_SUBSCRIPT_BASELINE = "-25000"


def _add_run(paragraph, text: str, font_size=None, bold=False, italic=False,
             strike=False, underline=False, font_name=None, hyperlink=None,
             baseline=None):
    """Add a formatted run to a paragraph."""
    run = paragraph.add_run()
    run.text = text
    if hyperlink:
        # Setting the address makes it a real, clickable link and lets the
        # template's hyperlink theme colour apply, rather than faking one with
        # blue underlined text.
        run.hyperlink.address = hyperlink
    if font_size:
        run.font.size = font_size
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True
    if strike:
        # python-pptx Font doesn't expose strikethrough — set via XML attribute
        rPr = run._r.get_or_add_rPr()
        rPr.set('strike', 'sngStrike')
    if underline:
        run.font.underline = True
    if font_name:
        run.font.name = font_name
    if baseline:
        # Nor super/subscript — same route.
        run._r.get_or_add_rPr().set('baseline', baseline)
    return run


def _parse_segment(text: str, paragraph, font_size=None, bold=False, italic=False,
                   escape_ctx=None, hyperlinks=True):
    """Parse a text segment for inline markdown and create runs."""
    for part in _INLINE_FORMAT_RE.split(text):
        if not part:
            continue

        if part.startswith('***') and part.endswith('***') and len(part) > 6:
            # Bold italic
            _parse_segment(part[3:-3], paragraph, font_size=font_size, hyperlinks=hyperlinks,
                           bold=True, italic=True, escape_ctx=escape_ctx)

        elif part.startswith('**') and part.endswith('**') and len(part) > 4:
            # Bold
            _parse_segment(part[2:-2], paragraph, font_size=font_size, hyperlinks=hyperlinks,
                           bold=True, italic=italic, escape_ctx=escape_ctx)

        elif part.startswith('~~') and part.endswith('~~') and len(part) > 4:
            # Strikethrough
            _add_run(paragraph, _restore_escapes(part[2:-2], escape_ctx),
                     font_size=font_size, bold=bold, italic=italic, strike=True)

        elif part.startswith('__') and part.endswith('__') and len(part) > 4 and not part.startswith('___'):
            # Underline
            _add_run(paragraph, _restore_escapes(part[2:-2], escape_ctx),
                     font_size=font_size, bold=bold, italic=italic, underline=True)

        elif part.startswith('*') and part.endswith('*') and not part.startswith('**') and len(part) > 2:
            # Italic
            _parse_segment(part[1:-1], paragraph, font_size=font_size, hyperlinks=hyperlinks,
                           bold=bold, italic=True, escape_ctx=escape_ctx)

        elif part.startswith('`') and part.endswith('`') and len(part) > 2:
            # Code (monospace)
            _add_run(paragraph, _restore_escapes(part[1:-1], escape_ctx),
                     font_size=font_size, bold=bold, italic=italic, font_name='Courier New')

        elif part.startswith('^') and part.endswith('^') and len(part) > 2:
            # Superscript
            _add_run(paragraph, _restore_escapes(part[1:-1], escape_ctx),
                     font_size=font_size, bold=bold, italic=italic,
                     baseline=_SUPERSCRIPT_BASELINE)

        elif part.startswith('~') and part.endswith('~') and not part.startswith('~~') and len(part) > 2:
            # Subscript (single tilde; the strikethrough branch owns "~~")
            _add_run(paragraph, _restore_escapes(part[1:-1], escape_ctx),
                     font_size=font_size, bold=bold, italic=italic,
                     baseline=_SUBSCRIPT_BASELINE)

        elif part.startswith('[') and _LINK_RE.match(part):
            # Hyperlink: render the label with its own formatting, then point
            # every run it produced at the target.
            label, target = _LINK_RE.match(part).groups()
            first_new = len(paragraph.runs)
            _parse_segment(label, paragraph, font_size=font_size, hyperlinks=hyperlinks,
                           bold=bold, italic=italic, escape_ctx=escape_ctx)
            address = _restore_escapes(target, escape_ctx)
            # A refused scheme keeps its label as plain text; the builder
            # reports it once from a pre-scan (refused_link_targets).
            if hyperlinks and is_safe_link_target(address):
                for run in paragraph.runs[first_new:]:
                    run.hyperlink.address = address

        else:
            # Plain text
            _add_run(paragraph, _restore_escapes(part, escape_ctx),
                     font_size=font_size, bold=bold, italic=italic)




