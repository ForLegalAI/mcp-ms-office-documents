"""Numbered-list restart support.

python-docx exposes no high-level API for restarting list numbering, but it *does*
ship the oxml helper methods needed to do it without hand-authoring XML strings
(``CT_Numbering.add_num``, ``CT_Num.add_lvlOverride``, ``CT_NumLvl.add_startOverride``).

The OOXML-correct way to restart a numbered list is to create a fresh numbering
*instance* (``<w:num>``) that references the same ``<w:abstractNum>`` as the list style
but overrides its start value, then attach that instance to the paragraph via an explicit
``<w:numPr>``. That overrides the style's shared numbering for just that run of items, so
each logical list counts independently.

See ``docs/plan-issues-66-67.md`` (Issue #67) for the design rationale.
"""
import logging

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)

# Style names whose abstract numbering definition we reuse for ordered lists.
_ORDERED_LIST_STYLES = ("List Number", "List Number 2", "List Number 3")

# Indentation (twips) per level for the synthesized fallback definition.
_FALLBACK_INDENT_STEP = 360  # 0.25"


def _numbering_root(doc):
    """Return the ``<w:numbering>`` root element, creating the part if absent."""
    return doc.part.numbering_part.element


def _abstract_id_from_style(doc, numbering_root, style_name):
    """Resolve the abstractNumId backing *style_name*, or ``None`` if it has none."""
    try:
        style_el = doc.styles[style_name]._element
    except KeyError:
        return None
    num_ids = style_el.xpath('.//w:numPr/w:numId/@w:val')
    if not num_ids:
        return None
    try:
        num = numbering_root.num_having_numId(int(num_ids[0]))
    except (KeyError, ValueError):
        return None
    return num.abstractNumId.val


def _find_decimal_abstract(numbering_root):
    """Return the id of any existing decimal abstractNum, or ``None``."""
    for abstract in numbering_root.xpath('./w:abstractNum'):
        fmt = abstract.xpath('./w:lvl[@w:ilvl="0"]/w:numFmt/@w:val')
        if fmt and fmt[0] == 'decimal':
            return abstract.get(qn('w:abstractNumId'))
    return None


def _create_decimal_abstract(numbering_root):
    """Synthesize a minimal 3-level decimal ``<w:abstractNum>`` and return its id.

    Degraded fallback used only when a template carries neither a numbered list style
    nor any reusable decimal definition. Inserted before the first ``<w:num>`` so the
    numbering part keeps schema order (all abstractNum before all num).
    """
    existing = [int(v) for v in numbering_root.xpath('./w:abstractNum/@w:abstractNumId')]
    abstract_id = max(existing) + 1 if existing else 0

    abstract = OxmlElement('w:abstractNum')
    abstract.set(qn('w:abstractNumId'), str(abstract_id))
    multi = OxmlElement('w:multiLevelType')
    multi.set(qn('w:val'), 'multilevel')
    abstract.append(multi)
    for ilvl in range(3):
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(ilvl))
        start = OxmlElement('w:start'); start.set(qn('w:val'), '1'); lvl.append(start)
        fmt = OxmlElement('w:numFmt'); fmt.set(qn('w:val'), 'decimal'); lvl.append(fmt)
        text = OxmlElement('w:lvlText'); text.set(qn('w:val'), '%%%d.' % (ilvl + 1)); lvl.append(text)
        jc = OxmlElement('w:lvlJc'); jc.set(qn('w:val'), 'left'); lvl.append(jc)
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), str(_FALLBACK_INDENT_STEP * (ilvl + 1)))
        ind.set(qn('w:hanging'), str(_FALLBACK_INDENT_STEP))
        pPr.append(ind)
        lvl.append(pPr)
        abstract.append(lvl)

    first_num = numbering_root.find(qn('w:num'))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering_root.append(abstract)
    return str(abstract_id)


def resolve_ordered_abstract_num_id(doc):
    """Return ``(abstract_num_id, numbering_root)`` for ordered-list restart instances.

    Prefers the abstract definition backing the ``List Number`` style so restarted lists
    keep the template's numbering format; falls back to any decimal definition, and finally
    synthesizes one (degraded mode). Returns ``(None, root)`` only if synthesis itself fails.
    """
    numbering_root = _numbering_root(doc)
    for style_name in _ORDERED_LIST_STYLES:
        abstract_id = _abstract_id_from_style(doc, numbering_root, style_name)
        if abstract_id is not None:
            return abstract_id, numbering_root
    abstract_id = _find_decimal_abstract(numbering_root)
    if abstract_id is not None:
        return abstract_id, numbering_root
    try:
        return _create_decimal_abstract(numbering_root), numbering_root
    except Exception:
        logger.warning("Could not synthesize a numbering definition; "
                       "ordered lists will not restart.", exc_info=True)
        return None, numbering_root


def new_restarted_num(numbering_root, abstract_num_id, level, start=1):
    """Create a fresh ``<w:num>`` restarting *level* at *start*; return its numId."""
    num = numbering_root.add_num(abstract_num_id)
    num.add_lvlOverride(level).add_startOverride(start)
    return num.numId


def apply_numbering(paragraph, num_id, level):
    """Attach an explicit ``<w:numPr>`` (numId + ilvl) to *paragraph*."""
    numPr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    numPr.get_or_add_numId().val = num_id
    numPr.get_or_add_ilvl().val = level
