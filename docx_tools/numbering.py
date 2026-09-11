"""Numbered-list restart support.

python-docx exposes no high-level API for restarting list numbering, but it *does*
ship the oxml helper methods needed to do it without hand-authoring XML strings
(``CT_Numbering.add_num``, ``CT_Num.add_lvlOverride``, ``CT_NumLvl.add_startOverride``).

The OOXML-correct way to restart a numbered list is to create a fresh numbering
*instance* (``<w:num>``) that references the same ``<w:abstractNum>`` as the list style
but overrides its start value, then attach that instance to the paragraph via an explicit
``<w:numPr>``. That overrides the style's shared numbering for just that run of items, so
each logical list counts independently.

See docs/development/tools/word.md ("Ordered lists") for the design rationale;
the original discussion is issue #67.
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
    """Return the ``<w:numbering>`` root element.

    python-docx's ``numbering_part`` normally materialises an (empty) part on
    access, but guard against a missing part so the caller can degrade instead
    of raising deep inside list rendering.
    """
    numbering_part = doc.part.numbering_part
    if numbering_part is None:
        raise RuntimeError(
            "Document has no numbering part; ordered-list restart needs a template "
            "that defines at least one list style (e.g. 'List Number')."
        )
    return numbering_part.element


def _style_numbering(doc, numbering_root, style_name):
    """Return ``(abstract_num_id, ilvl)`` declared by *style_name*'s numbering.

    Follows the style's ``basedOn`` chain so a style that inherits its list
    definition still resolves. The declared ``w:ilvl`` (default 0) identifies
    which level of the abstract definition the style renders — e.g. Word's
    built-in ``List Number 2`` points at ``ilvl 1`` of a shared multi-level
    definition, while a template with three single-level definitions declares
    no ``ilvl`` at all. Returns ``(None, None)`` when the style is missing or
    carries no usable numbering reference.
    """
    try:
        style = doc.styles[style_name]
    except KeyError:
        return None, None
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        style_el = style._element
        num_ids = style_el.xpath('./w:pPr/w:numPr/w:numId/@w:val')
        if num_ids:
            try:
                num = numbering_root.num_having_numId(int(num_ids[0]))
            except (KeyError, ValueError):
                return None, None
            ilvls = style_el.xpath('./w:pPr/w:numPr/w:ilvl/@w:val')
            try:
                ilvl = int(ilvls[0]) if ilvls else 0
            except ValueError:
                ilvl = 0
            # Normalise to str so all resolver paths return the same type.
            return str(num.abstractNumId.val), ilvl
        style = style.base_style
    return None, None


def _abstract_defined_ilvls(numbering_root, abstract_num_id):
    """Return the sorted ``ilvl`` values the abstract definition declares.

    Uses qualified-name ``findall`` rather than ``.xpath('./w:…')`` because
    ``<w:abstractNum>`` has no registered oxml class — on plain lxml elements
    the ``w:`` prefix is undefined and such an xpath raises.
    """
    for abstract in numbering_root.findall(qn('w:abstractNum')):
        if abstract.get(qn('w:abstractNumId')) == str(abstract_num_id):
            ilvls = []
            for lvl in abstract.findall(qn('w:lvl')):
                try:
                    ilvls.append(int(lvl.get(qn('w:ilvl'))))
                except (TypeError, ValueError):
                    pass
            return sorted(ilvls)
    return []


def _clamp_ilvl(numbering_root, abstract_num_id, ilvl):
    """Clamp *ilvl* to a level the abstract actually defines.

    Referencing an undefined level renders without number or indent in Word,
    so fall back to the nearest defined level below (or the lowest defined).
    """
    defined = _abstract_defined_ilvls(numbering_root, abstract_num_id)
    if not defined or ilvl in defined:
        return ilvl
    lower = [v for v in defined if v <= ilvl]
    return max(lower) if lower else defined[0]


def _find_decimal_abstract(numbering_root):
    """Return the id of any existing decimal abstractNum, or ``None``.

    Uses qualified-name ``findall`` rather than ``.xpath('./w:…')`` because
    ``<w:abstractNum>`` has no registered oxml class — on plain lxml elements
    the ``w:`` prefix is undefined and such an xpath raises.
    """
    for abstract in numbering_root.findall(qn('w:abstractNum')):
        for lvl in abstract.findall(qn('w:lvl')):
            if lvl.get(qn('w:ilvl')) != '0':
                continue
            fmt = lvl.find(qn('w:numFmt'))
            if fmt is not None and fmt.get(qn('w:val')) == 'decimal':
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
    def _el(tag, **attrs):
        """Build a ``w:``-namespaced OxmlElement with ``w:``-namespaced attributes."""
        el = OxmlElement(tag)
        for key, value in attrs.items():
            el.set(qn('w:' + key), value)
        return el

    for ilvl in range(3):
        lvl = _el('w:lvl', ilvl=str(ilvl))
        lvl.append(_el('w:start', val='1'))
        lvl.append(_el('w:numFmt', val='decimal'))
        lvl.append(_el('w:lvlText', val='%%%d.' % (ilvl + 1)))
        lvl.append(_el('w:lvlJc', val='left'))
        pPr = OxmlElement('w:pPr')
        pPr.append(_el('w:ind',
                       left=str(_FALLBACK_INDENT_STEP * (ilvl + 1)),
                       hanging=str(_FALLBACK_INDENT_STEP)))
        lvl.append(pPr)
        abstract.append(lvl)

    first_num = numbering_root.find(qn('w:num'))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering_root.append(abstract)
    return str(abstract_id)


def resolve_ordered_numbering(doc, level=0, style_name=None):
    """Return ``(abstract_num_id, ilvl, numbering_root)`` for restart instances.

    *level* is the 0-based markdown nesting level; *style_name* is the ordered-list
    paragraph style actually applied at that level (a ``style_mapping`` override or a
    ``<!-- style: … -->`` directive style). Resolution order:

    1. *style_name*'s own numbering (declared ``numId`` + ``ilvl``, following
       ``basedOn``) — restarted lists keep that style's numeral format and level
       indents instead of degrading to the built-in ``List Number`` definition.
    2. The built-in style for this level (``List Number`` / ``… 2`` / ``… 3``),
       then the other built-ins — each with its *own* declared ``ilvl``, so a
       template whose level-2 style carries a separate single-level definition
       restarts that definition's level 0 rather than pointing at an ``ilvl``
       the abstract does not define (which Word renders unnumbered).
    3. Any existing decimal ``abstractNum`` (``ilvl`` clamped to its levels).
    4. A synthesized minimal decimal definition (degraded mode).

    Returns ``(None, None, None)`` if no numbering part is available or synthesis
    fails, so the caller renders lists without restart rather than raising.
    """
    try:
        numbering_root = _numbering_root(doc)
    except RuntimeError:
        logger.warning("No numbering part available; ordered lists will not restart.",
                       exc_info=True)
        return None, None, None
    builtin = _ORDERED_LIST_STYLES[min(level, len(_ORDERED_LIST_STYLES) - 1)]
    candidates = [style_name] if style_name else []
    candidates.append(builtin)
    candidates.extend(name for name in _ORDERED_LIST_STYLES if name != builtin)
    seen = set()
    for name in candidates:
        if name in seen:
            continue
        seen.add(name)
        abstract_id, ilvl = _style_numbering(doc, numbering_root, name)
        if abstract_id is not None:
            return abstract_id, _clamp_ilvl(numbering_root, abstract_id, ilvl), numbering_root
    abstract_id = _find_decimal_abstract(numbering_root)
    if abstract_id is not None:
        return abstract_id, _clamp_ilvl(numbering_root, abstract_id, level), numbering_root
    try:
        abstract_id = _create_decimal_abstract(numbering_root)
        return abstract_id, min(level, 2), numbering_root
    except Exception:
        logger.warning("Could not synthesize a numbering definition; "
                       "ordered lists will not restart.", exc_info=True)
        return None, None, None


def new_restarted_num(numbering_root, abstract_num_id, ilvl, start=1):
    """Create a fresh ``<w:num>`` restarting *ilvl* at *start*; return its numId."""
    num = numbering_root.add_num(abstract_num_id)
    num.add_lvlOverride(ilvl).add_startOverride(start)
    return num.numId


def apply_numbering(paragraph, num_id, ilvl):
    """Attach an explicit ``<w:numPr>`` (numId + ilvl) to *paragraph*."""
    numPr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    numPr.get_or_add_numId().val = num_id
    numPr.get_or_add_ilvl().val = ilvl


def style_indent_attrs(style):
    """Return the ``w:ind`` attributes *style* defines (following ``basedOn``), or None.

    *style* is a python-docx paragraph style object (e.g. ``paragraph.style``).
    """
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        ppr = style._element.find(qn('w:pPr'))
        if ppr is not None:
            ind = ppr.find(qn('w:ind'))
            if ind is not None:
                return dict(ind.attrib)
        style = style.base_style
    return None


def apply_style_indent(paragraph, ind_attrs):
    """Re-assert a style's indents as direct formatting on *paragraph*.

    Word treats the ``w:ind`` inside a numbering level referenced by a *direct*
    ``w:numPr`` as direct formatting, silently overriding the paragraph style's
    own indents — so a template's styled left indent would be lost on every
    ordered-list item. Re-stating the style's ``w:ind`` directly on the
    paragraph restores the intended precedence; attributes the style does not
    define (typically the number's ``hanging``) still come from the numbering
    level, because ``w:ind`` attributes merge individually across the cascade.
    """
    if not ind_attrs:
        return
    ind = paragraph._p.get_or_add_pPr().get_or_add_ind()
    for key, value in ind_attrs.items():
        ind.set(key, value)
