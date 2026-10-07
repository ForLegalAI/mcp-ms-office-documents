"""Fill Word content controls in dynamic DOCX templates from template arguments.

A content control (``<w:sdt>``, Word's Developer ▸ Controls) is bound to an
argument by its **Tag** (Developer ▸ Properties ▸ Tag)::

    full_name     the control takes the value of the argument ``full_name``
    size=small    check boxes only: ticked when ``size`` equals ``small``
                  (a list argument: when it contains ``small``)

Supported kinds (``FILLERS``):

- check box (``w14:checkbox``): bare tag -> true/false; ``name=option`` -> the
  value equals the option, or a list value contains it. Only the state and the
  glyph change, so the box stays clickable.
- drop-down list (``w:dropDownList``): the item whose value or display text
  equals the argument is selected (content text + ``w:lastValue``).
- plain text (``w:text``): the content becomes the value, in the control's own
  run formatting; a multi-line control gets line breaks.
- combo box (``w:comboBox``): an item whose value or display text matches is
  shown as that item; any other value is written as free text — a combo box
  allows it.
- date picker (``w:date``): the value is an ISO date (``2026-10-06``). It is
  stored in ``w:fullDate`` and shown in the control's own ``w:dateFormat``,
  numerically: a format that spells out month or day names, or has a time,
  is replaced by the numeric default for the control's language (``w:lid``).

Rules shared by every kind:

- Only a tag naming a declared argument binds; untagged controls and foreign
  tags are left exactly as the template has them, as are the kinds without a
  filler (rich text, combo box, date, picture, repeating section, …).
- A value that is not there (``None``, an empty or blank string) leaves the
  control as the template has it — a form keeps its "click here" prompt so the
  client can finish it by hand. ``False`` and an empty list are values: they
  untick.
- A filled control stays a content control; its ``w:dataBinding`` is dropped,
  or Word would overwrite the new content from the document's XML data store.
- A value the document could not take — a drop-down with no such item, a tag
  on a kind that is not filled, ``=option`` on a non-check box, a date that is
  not an ISO date — is reported on the build's warning channel
  (``control_item_missing`` / ``control_not_filled`` / ``control_value_invalid``,
  located by ``tag``), so the caller learns it from the tool result, not from
  the server log. A date shown in a numeric format other than the control's
  own is reported as ``control_date_format_simplified`` (info).

``dynamic_docx_tools`` calls :func:`resolve_content_controls` after the
conditionals and before placeholder substitution; the admin preview calls the
same function, and the admin analyser reads :func:`describe_content_controls`.
"""
from __future__ import annotations

import copy
import datetime as dt
import logging
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Callable, Dict, Iterator, List, Optional, Tuple

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from . import warnings as W

if TYPE_CHECKING:  # imported only for type hints
    from docx import Document as DocxDocument

logger = logging.getLogger(__name__)

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"

# Kinds, from the one element in <w:sdtPr> that types the control.
CHECKBOX = "checkbox"
DROPDOWN = "dropdown"
TEXT = "text"
RICH_TEXT = "richtext"
COMBO_BOX = "combobox"
DATE = "date"
PICTURE = "picture"
REPEATING = "repeating"
OTHER = "other"

_TYPE_ELEMENTS: List[Tuple[str, str]] = [
    (f"{{{W14}}}checkbox", CHECKBOX),
    (qn("w:dropDownList"), DROPDOWN),
    (qn("w:comboBox"), COMBO_BOX),
    (qn("w:text"), TEXT),
    (qn("w:date"), DATE),
    (qn("w:picture"), PICTURE),
    (f"{{{W15}}}repeatingSection", REPEATING),
    (f"{{{W15}}}repeatingSectionItem", OTHER),
    (qn("w:docPartObj"), OTHER),
    (qn("w:docPartList"), OTHER),
    (qn("w:group"), OTHER),
    (qn("w:citation"), OTHER),
    (qn("w:bibliography"), OTHER),
    (qn("w:equation"), OTHER),
]

# A bare-tag check box bound to a text argument is ticked by these words.
_TRUE_WORDS = {"1", "true", "yes", "y", "ano", "x"}


@dataclass
class ContentControl:
    """One ``<w:sdt>`` and what its properties say about it."""

    element: Any
    kind: str
    tag: str = ""
    alias: str = ""
    name: str = ""                 # argument the tag names ("" when untagged)
    option: Optional[str] = None   # the part after "=" (check boxes)
    items: List[Tuple[str, str]] = field(default_factory=list)  # (value, display)

    @property
    def properties(self):
        return self.element.find(qn("w:sdtPr"))


def parse_tag(tag: str) -> Tuple[str, Optional[str]]:
    """``"size=small"`` -> ``("size", "small")``; ``"full_name"`` -> ``("full_name", None)``."""
    name, sep, option = (tag or "").partition("=")
    return name.strip(), (option.strip() if sep else None)


def classify(sdt_pr) -> str:
    """The kind of control an ``<w:sdtPr>`` describes (no type element = rich text)."""
    if sdt_pr is None:
        return RICH_TEXT
    for tag, kind in _TYPE_ELEMENTS:
        if sdt_pr.find(tag) is not None:
            return kind
    return RICH_TEXT


def _attr(el, name: str) -> str:
    return (el.get(qn(name)) or "") if el is not None else ""


def _list_items(sdt_pr, kind: str) -> List[Tuple[str, str]]:
    container = sdt_pr.find(qn("w:dropDownList" if kind == DROPDOWN else "w:comboBox"))
    items = []
    for item in container.findall(qn("w:listItem")) if container is not None else []:
        value = _attr(item, "w:value")
        if not value:
            continue  # Word's own "Choose an item." entry: a prompt, not a choice
        items.append((value, _attr(item, "w:displayText") or value))
    return items


def _roots(doc: "DocxDocument") -> Iterator[Any]:
    """The body plus every header/footer part the document really owns."""
    yield doc.element.body
    # Elements, not their id(): an id is only unique while its object lives.
    seen: List[Any] = []
    for section in doc.sections:
        for part in (section.header, section.footer,
                     section.first_page_header, section.first_page_footer,
                     section.even_page_header, section.even_page_footer):
            if part is None or part.is_linked_to_previous:
                continue  # a linked part belongs to an earlier section; don't create one
            element = part._element
            if not any(element is s for s in seen):
                seen.append(element)
                yield element


def iter_content_controls(doc: "DocxDocument") -> Iterator[ContentControl]:
    """Every content control in the body, tables, headers and footers, in order."""
    for root in _roots(doc):
        for sdt in list(root.iter(qn("w:sdt"))):
            pr = sdt.find(qn("w:sdtPr"))
            kind = classify(pr)
            tag = _attr(pr.find(qn("w:tag")) if pr is not None else None, "w:val").strip()
            alias = _attr(pr.find(qn("w:alias")) if pr is not None else None, "w:val").strip()
            name, option = parse_tag(tag)
            items = _list_items(pr, kind) if kind in (DROPDOWN, COMBO_BOX) else []
            yield ContentControl(sdt, kind, tag, alias, name, option, items)


def describe_content_controls(doc: "DocxDocument") -> List[Dict[str, Any]]:
    """What the admin analyser shows: one dict per control."""
    return [
        {"tag": cc.tag, "name": cc.name, "option": cc.option, "kind": cc.kind,
         "alias": cc.alias, "items": [v for v, _ in cc.items],
         "labels": [d for _, d in cc.items],
         "supported": cc.kind in FILLERS}
        for cc in iter_content_controls(doc)
    ]


# ---------------------------------------------------------------------------
# Writing content
# ---------------------------------------------------------------------------

def _is_unset(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _matches(value: Any, wanted: str) -> bool:
    """Case-insensitive equality; a list matches when any item does."""
    items = value if isinstance(value, (list, tuple)) else [value]
    wanted = wanted.strip().casefold()
    return any(("" if i is None else str(i)).strip().casefold() == wanted for i in items)


def _drop_binding(cc: ContentControl) -> None:
    for binding in cc.properties.findall(qn("w:dataBinding")):
        cc.properties.remove(binding)


def _content_rpr(cc: ContentControl):
    """Run properties for new content: the control's first run, else its own rPr."""
    content = cc.element.find(qn("w:sdtContent"))
    run = content.find(".//" + qn("w:r")) if content is not None else None
    rpr = run.find(qn("w:rPr")) if run is not None else None
    if rpr is None:
        rpr = cc.properties.find(qn("w:rPr"))
    return copy.deepcopy(rpr) if rpr is not None else None


def _set_text(cc: ContentControl, text: str, multiline: bool = False) -> None:
    """Replace the control's content with *text*, keeping its run formatting.

    While a control shows its prompt (``w:showingPlcHdr``) the run carries the
    "Placeholder Text" character style; both go, or the value would print grey.
    """
    content = cc.element.find(qn("w:sdtContent"))
    if content is None:
        return
    rpr = _content_rpr(cc)
    showing_prompt = cc.properties.find(qn("w:showingPlcHdr"))
    if showing_prompt is not None:
        cc.properties.remove(showing_prompt)
        if rpr is not None:
            for style in rpr.findall(qn("w:rStyle")):
                rpr.remove(style)

    # Run-level controls hold runs directly; block-, row- and cell-level ones
    # hold paragraphs. Write into the first paragraph, empty the others.
    paragraphs = list(content.iter(qn("w:p")))
    target = paragraphs[0] if paragraphs else content
    for run in list(content.iter(qn("w:r"))):
        run.getparent().remove(run)
    for extra in paragraphs[1:]:
        for child in list(extra):
            if child.tag != qn("w:pPr"):
                extra.remove(child)

    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(rpr)
    if multiline:
        lines = text.splitlines() or [""]
    else:
        lines = [" ".join(text.splitlines())]
    for i, line in enumerate(lines):
        if i:
            run.append(OxmlElement("w:br"))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = line
        run.append(t)
    target.append(run)


def _fill_checkbox(cc: ContentControl, value: Any, warnings=None) -> bool:
    if cc.option is None:
        on = (value.strip().casefold() in _TRUE_WORDS) if isinstance(value, str) else bool(value)
    else:
        on = _matches(value, cc.option)
    box = cc.properties.find(f"{{{W14}}}checkbox")
    checked = box.find(f"{{{W14}}}checked")
    if checked is None:
        checked = box.makeelement(f"{{{W14}}}checked", {})
        box.insert(0, checked)
    checked.set(f"{{{W14}}}val", "1" if on else "0")

    state = box.find(f"{{{W14}}}{'checkedState' if on else 'uncheckedState'}")
    code = state.get(f"{{{W14}}}val") if state is not None else None
    glyph = chr(int(code, 16)) if code else ("☒" if on else "☐")
    font = state.get(f"{{{W14}}}font") if state is not None else None
    content = cc.element.find(qn("w:sdtContent"))
    run = content.find(".//" + qn("w:r")) if content is not None else None
    text = run.find(qn("w:t")) if run is not None else None
    if text is None:
        return True  # state set; a box with no glyph run has nothing to redraw
    text.text = glyph
    if font:
        rpr = run.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            run.insert(0, rpr)
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rpr.insert(0, fonts)
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            fonts.set(qn(attr), font)
    return True


def _fill_dropdown(cc: ContentControl, value: Any, warnings=None) -> bool:
    for item_value, display in cc.items:
        if _matches(value, item_value) or _matches(value, display):
            _set_text(cc, display)
            cc.properties.find(qn("w:dropDownList")).set(qn("w:lastValue"), item_value)
            return True
    logger.warning("[dynamic-docx] Drop-down '%s' has no item %r; leaving it as the "
                   "template has it.", cc.tag, value)
    if warnings is not None:
        choices = ", ".join(v for v, _ in cc.items) or "none"
        warnings.add(W.CONTROL_ITEM_MISSING,
                     f"The drop-down has no item {value!r}, so it was left unselected. "
                     f"Its items are: {choices}.", tag=cc.tag)
    return False


def _fill_text(cc: ContentControl, value: Any, warnings=None) -> bool:
    text = ", ".join(map(str, value)) if isinstance(value, (list, tuple)) else str(value)
    multiline = _attr(cc.properties.find(qn("w:text")), "w:multiLine") in ("1", "true", "on")
    _set_text(cc, text, multiline=multiline)
    return True


def _fill_combobox(cc: ContentControl, value: Any, warnings=None) -> bool:
    text = ", ".join(map(str, value)) if isinstance(value, (list, tuple)) else str(value)
    for item_value, display in cc.items:
        if _matches(text, item_value) or _matches(text, display):
            text, last = display, item_value
            break
    else:
        last = text  # free text is a valid combo-box value, not a miss
    _set_text(cc, text)
    cc.properties.find(qn("w:comboBox")).set(qn("w:lastValue"), last)
    return True


# --- date picker ----------------------------------------------------------------

#: The numeric format a date is shown in when the control's own cannot be used
#: (it names months or days, has a time, or is missing), by language of w:lid.
_DEFAULT_DATE_FORMATS = {
    "cs": "d. M. yyyy", "sk": "d. M. yyyy",
    "de": "dd.MM.yyyy", "pl": "dd.MM.yyyy",
    "en-us": "M/d/yyyy", "en": "dd/MM/yyyy",
}
_ISO_FORMAT = "yyyy-MM-dd"
_DATE_TOKEN = re.compile(r"'[^']*'|\"[^\"]*\"|y+|M+|d+|[HhmsAaPpt]+|.", re.S)


def _default_date_format(lid: str) -> str:
    lid = (lid or "").casefold()
    return (_DEFAULT_DATE_FORMATS.get(lid) or _DEFAULT_DATE_FORMATS.get(lid.split("-")[0])
            or _ISO_FORMAT)


def _format_date(day: dt.date, fmt: str) -> Optional[str]:
    """*day* in Word date-format codes, or None when *fmt* is not numeric-only.

    ``d dd M MM yy yyyy`` and quoted literals are rendered; names (``MMM``,
    ``ddd`` and longer) and any time code make the format unusable here.
    """
    out = []
    for token in _DATE_TOKEN.findall(fmt):
        if token[0] in "'\"" and len(token) >= 2:
            out.append(token[1:-1])
        elif token in ("d", "dd"):
            out.append(f"{day.day:0{len(token)}d}")
        elif token in ("M", "MM"):
            out.append(f"{day.month:0{len(token)}d}")
        elif token == "yy":
            out.append(f"{day.year % 100:02d}")
        elif token == "yyyy":
            out.append(f"{day.year:04d}")
        elif token[0] in "dMyHhmsAaPpt":
            return None  # a name, a time, or a year code Word does not define

        else:
            out.append(token)
    return "".join(out)


def _parse_iso_date(value: Any) -> Optional[dt.date]:
    """A ``date``, or an ISO date / date-time string; anything else is None."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value).strip()
    try:
        return dt.date.fromisoformat(text)
    except ValueError:
        pass
    try:
        return dt.datetime.fromisoformat(text).date()  # "2026-10-06T09:30:00"
    except ValueError:
        return None


def _fill_date(cc: ContentControl, value: Any, warnings=None) -> bool:
    day = _parse_iso_date(value)
    if day is None:
        logger.warning("[dynamic-docx] Date control '%s': %r is not an ISO date; leaving "
                       "it.", cc.tag, value)
        if warnings is not None:
            warnings.add(W.CONTROL_VALUE_INVALID,
                         f"{value!r} is not a date in ISO form (YYYY-MM-DD), so the date "
                         "control was left as the template has it.", tag=cc.tag)
        return False
    date_el = cc.properties.find(qn("w:date"))
    own = _attr(date_el.find(qn("w:dateFormat")), "w:val")
    lid = _attr(date_el.find(qn("w:lid")), "w:val")
    text = _format_date(day, own) if own else None
    if text is None:
        fallback = _default_date_format(lid)
        text = _format_date(day, fallback)
        if own:
            logger.info("[dynamic-docx] Date control '%s': format %r is not numeric; "
                        "shown as %r.", cc.tag, own, fallback)
            if warnings is not None:
                warnings.add(W.CONTROL_DATE_FORMAT_SIMPLIFIED,
                             f"The date was shown as {text!r} (numeric) instead of the "
                             f"control's format {own!r}.", tag=cc.tag)
    date_el.set(qn("w:fullDate"), f"{day.isoformat()}T00:00:00Z")
    _set_text(cc, text)
    return True


#: kind -> filler ``(control, value, warnings) -> filled``. A kind missing here
#: is left as the template has it.
FILLERS: Dict[str, Callable[..., bool]] = {
    CHECKBOX: _fill_checkbox,
    DROPDOWN: _fill_dropdown,
    TEXT: _fill_text,
    COMBO_BOX: _fill_combobox,
    DATE: _fill_date,
}


def resolve_content_controls(doc: "DocxDocument", values: Dict[str, Any],
                             warnings=None) -> int:
    """Fill every tagged, supported control whose tag names a key of *values*.

    Args:
        doc: The Word document (mutated in place).
        values: Argument name -> validated value, as the tool received it.
        warnings: The build's :class:`~warning_channel.WarningChannel`, or None
            to discard them (the admin preview). A value the caller sent that a
            control could not take is reported here.

    Returns:
        How many controls were filled.
    """
    filled = 0
    for cc in list(iter_content_controls(doc)):
        if not cc.name or cc.name not in values:
            continue  # untagged, or a tag that names no argument of this template
        value = values[cc.name]
        if _is_unset(value):
            continue  # nothing sent: keep the template's prompt / state
        filler = FILLERS.get(cc.kind)
        reason = None
        if filler is None:
            reason = f"it is a {cc.kind} control, which is not filled yet"
        elif cc.option is not None and cc.kind != CHECKBOX:
            reason = "'=option' in a tag only works on check boxes"
        if reason:
            logger.warning("[dynamic-docx] Content control '%s' left as is: %s.", cc.tag, reason)
            if warnings is not None:
                warnings.add(W.CONTROL_NOT_FILLED,
                             f"The value for '{cc.name}' was not written into this "
                             f"control: {reason}. Fix the template.", tag=cc.tag)
            continue
        if filler(cc, value, warnings):
            _drop_binding(cc)
            filled += 1
    return filled
