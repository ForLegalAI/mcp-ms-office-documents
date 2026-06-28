"""Conditional block resolution for DOCX templates (Phase 1: block-level).

A template author can wrap a range of body content between marker paragraphs so
that the range is kept or dropped depending on a boolean tool argument:

    {{#if flag}}        keep the inner content when ``flag`` is truthy
    ...
    {{/if}}

    {{^if flag}}        keep the inner content when ``flag`` is falsy
    ...
    {{/if}}

Markers are recognised only when a paragraph's entire (trimmed) text is exactly
one marker token. Whitespace immediately inside the braces is tolerated, e.g.
``{{ #if flag }}``. Because the marker text is read from the paragraph's combined
runs, a marker that Word has split across several runs is still detected.

Granularity is *block-level*: everything between the markers — paragraphs,
lists, page breaks and whole tables — is a sibling element in the document body,
so all of it is pruned together. Markers are evaluated against the document body
only; conditionals inside table cells, headers and footers are not handled yet.

Error handling is intentionally forgiving ("warn and keep content"): if the
markers in the body are not well balanced, a warning is logged, **all content is
preserved**, and only the recognisable marker paragraphs are stripped so the
literal ``{{#if}}`` text does not leak into the generated document. Likewise, a
condition naming an unknown argument keeps its content (both for ``{{#if}}`` and
``{{^if}}``) so a typo never silently deletes a block.
"""
from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING, Any

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:  # imported only for type hints; avoids a runtime import
    from docx import Document as DocxDocument

logger = logging.getLogger(__name__)

# Marker paragraphs: the whole trimmed paragraph text must be exactly one token.
# ``\s*`` around the inner keyword/name tolerates stray spacing inside the braces
# (e.g. autocorrect or HTML paste producing "{{#if flag }}").
_OPEN_IF = re.compile(r'^\{\{\s*#if\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}$')
_OPEN_UNLESS = re.compile(r'^\{\{\s*\^if\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}$')
_CLOSE_IF = re.compile(r'^\{\{\s*/if\s*\}\}$')


class _Marker:
    """A recognised conditional marker on a paragraph."""

    __slots__ = ("kind", "name", "negate")

    def __init__(self, kind: str, name: str | None = None, negate: bool = False):
        self.kind = kind  # "open" or "close"
        self.name = name  # condition name (open markers only)
        self.negate = negate  # True for {{^if ...}}


def _parse_marker(text: str) -> _Marker | None:
    """Return the marker a paragraph represents, or None if it is not a marker."""
    s = text.strip()
    if not s:
        return None
    m = _OPEN_IF.match(s)
    if m:
        return _Marker("open", m.group(1), negate=False)
    m = _OPEN_UNLESS.match(s)
    if m:
        return _Marker("open", m.group(1), negate=True)
    if _CLOSE_IF.match(s):
        return _Marker("close")
    return None


# Public alias: the admin analyzer reuses marker parsing across the package
# boundary, so expose it under a non-underscore name.
parse_marker = _parse_marker


def _paragraph_text(p_elem) -> str:
    """Combined run text of a ``<w:p>`` element.

    Wrapping the raw element in a ``Paragraph`` with a ``None`` parent is safe
    here because only ``.text`` is read, which simply concatenates the runs and
    never touches the parent. This also transparently joins markers that Word
    has fragmented across multiple runs.
    """
    return Paragraph(p_elem, None).text


def _evaluate(marker: _Marker, conditions: dict[str, Any]) -> bool:
    """Whether the inner content of an open marker should be kept."""
    if marker.name not in conditions:
        # Unknown name -> always keep content, for both {{#if}} and {{^if}}, so a
        # typo'd condition never silently deletes a block.
        logger.warning(
            "[dynamic-docx] Condition '%s' is not a known argument; "
            "keeping its content.",
            marker.name,
        )
        return True
    value = bool(conditions[marker.name])
    return (not value) if marker.negate else value


def resolve_conditionals(doc: "DocxDocument", conditions: dict[str, Any]) -> None:
    """Prune conditional blocks in the document body in place.

    Must run *before* placeholder substitution, because it deletes whole block
    elements (paragraphs and tables) and operates on the raw body element list.

    Args:
        doc: The Word document to process (mutated in place).
        conditions: Mapping of argument name -> value; truthiness decides keeps.
    """
    # ``doc.element`` is the public CT_Document; ``.body`` is its CT_Body, whose
    # direct children are the body block elements (<w:p>, <w:tbl>, <w:sectPr>).
    _resolve_block_container(doc.element.body, conditions)


def _resolve_block_container(container, conditions: dict[str, Any]) -> None:
    """Resolve markers among the direct block children of ``container``."""
    p_tag = qn('w:p')
    children = list(container)

    # Pass 1: locate marker paragraphs and check that they are balanced.
    markers: dict[int, _Marker] = {}
    depth = 0
    balanced = True
    for elem in children:
        if elem.tag != p_tag:
            continue
        marker = _parse_marker(_paragraph_text(elem))
        if marker is None:
            continue
        markers[id(elem)] = marker
        if marker.kind == "open":
            depth += 1
        else:  # close
            depth -= 1
            if depth < 0:
                balanced = False
                depth = 0  # recover so later counting stays sane
    if depth != 0:
        balanced = False

    if not markers:
        return

    if not balanced:
        logger.warning(
            "[dynamic-docx] Unbalanced {{#if}}/{{/if}} markers in template body; "
            "keeping all content and stripping %d marker paragraph(s).",
            len(markers),
        )
        for elem in children:
            if id(elem) in markers:
                container.remove(elem)
        return

    # Pass 2: prune. Markers are guaranteed balanced here, so every close has a
    # matching open on the stack. A frame holds whether its inner content is
    # kept; an element survives only when every enclosing frame keeps its content.
    stack: list[bool] = []
    for elem in children:
        marker = markers.get(id(elem))
        if marker is not None:
            if marker.kind == "open":
                stack.append(_evaluate(marker, conditions))
            else:  # close — balance guarantees a frame to pop
                stack.pop()
            container.remove(elem)  # marker paragraphs are always removed
            continue
        if stack and not all(stack):
            container.remove(elem)
