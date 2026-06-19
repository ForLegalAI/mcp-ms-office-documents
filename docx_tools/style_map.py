"""Style mapping for markdown → DOCX rendering (issue #66, part A).

The renderer applies Word paragraph styles by name (``List Number``, ``Quote``,
``Table Grid``, ``Heading N`` …). Custom templates may name their styles
differently, so :class:`StyleMap` lets those built-in names be remapped without
touching the call sites — the defaults reproduce today's behaviour exactly.

A map is threaded explicitly through the processors (not held in global state) so
concurrent conversions on worker threads never share mutable mapping state. Config
overrides come from the ``style_mapping`` section of ``config/docx_templates.yaml``
(global) and each template's own ``style_mapping`` (per-template, wins over global).
See ``docs/plan-issues-66-67.md`` (Issue #66) for the design rationale.
"""
import logging
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Optional

from docx.table import Table
from docx.text.paragraph import Paragraph

logger = logging.getLogger(__name__)

_DEFAULT_HEADING = ("Heading 1", "Heading 2", "Heading 3",
                    "Heading 4", "Heading 5", "Heading 6")
_DEFAULT_LIST_NUMBER = ("List Number", "List Number 2", "List Number 3")
_DEFAULT_LIST_BULLET = ("List Bullet", "List Bullet 2", "List Bullet 3")


@dataclass(frozen=True)
class StyleMap:
    """Resolved style names used by the markdown renderer.

    Tuple fields are indexed by nesting level (0-based); ``heading`` is indexed by
    heading level 1-6 via :meth:`heading_style`.
    """
    heading: tuple = _DEFAULT_HEADING
    list_number: tuple = _DEFAULT_LIST_NUMBER
    list_bullet: tuple = _DEFAULT_LIST_BULLET
    quote: str = "Quote"
    table: str = "Table Grid"
    normal: str = "Normal"
    # Paragraph style for fenced code blocks. Default None = no paragraph style
    # (runs are still set to a monospace font); map it to a template style to
    # add shading/spacing.
    code: str = None

    def heading_style(self, level: int) -> str:
        """Style name for a 1-based heading *level* (clamped to 1..len)."""
        idx = min(max(level, 1), len(self.heading)) - 1
        return self.heading[idx]


DEFAULT_STYLE_MAP = StyleMap()

# Recognised config keys → where they land in StyleMap.
_HEADING_KEYS = {f"heading_{i}": i - 1 for i in range(1, 7)}
_LIST_NUMBER_KEYS = {"list_number": 0, "list_number_2": 1, "list_number_3": 2}
_LIST_BULLET_KEYS = {"list_bullet": 0, "list_bullet_2": 1, "list_bullet_3": 2}
_SCALAR_KEYS = ("quote", "table", "normal", "code")


def _normalize(mapping: dict) -> dict:
    """Translate a flat config dict into ``StyleMap`` field overrides."""
    heading = list(_DEFAULT_HEADING)
    list_number = list(_DEFAULT_LIST_NUMBER)
    list_bullet = list(_DEFAULT_LIST_BULLET)
    touched_heading = touched_ln = touched_lb = False
    out = {}
    for raw_key, raw_val in mapping.items():
        key = str(raw_key).strip().lower()
        if raw_val is None or str(raw_val).strip() == "":
            continue
        val = str(raw_val)
        if key in _HEADING_KEYS:
            heading[_HEADING_KEYS[key]] = val
            touched_heading = True
        elif key in _LIST_NUMBER_KEYS:
            list_number[_LIST_NUMBER_KEYS[key]] = val
            touched_ln = True
        elif key in _LIST_BULLET_KEYS:
            list_bullet[_LIST_BULLET_KEYS[key]] = val
            touched_lb = True
        elif key in _SCALAR_KEYS:
            out[key] = val
        else:
            logger.warning("Unknown style_mapping key %r ignored.", raw_key)
    if touched_heading:
        out["heading"] = tuple(heading)
    if touched_ln:
        out["list_number"] = tuple(list_number)
    if touched_lb:
        out["list_bullet"] = tuple(list_bullet)
    return out


def build_style_map(*mappings) -> StyleMap:
    """Merge zero or more config dicts onto the defaults; later mappings win.

    Returns the shared :data:`DEFAULT_STYLE_MAP` when nothing is overridden.
    """
    merged = {}
    for mapping in mappings:
        if mapping:
            merged.update({str(k).strip().lower(): v for k, v in mapping.items()})
    overrides = _normalize(merged)
    return replace(DEFAULT_STYLE_MAP, **overrides) if overrides else DEFAULT_STYLE_MAP


def apply_style(obj, style_name, fallback="Normal") -> None:
    """Set ``obj.style`` to *style_name*, falling back if the style is missing.

    *obj* is a paragraph or table. A missing style raises ``KeyError`` in
    python-docx; we log and fall back instead of letting it abort the render.
    """
    if not style_name:
        return
    try:
        obj.style = style_name
        return
    except KeyError:
        logger.warning("Style %r not found in document; falling back to %r.",
                       style_name, fallback)
    if fallback and fallback != style_name:
        try:
            obj.style = fallback
        except KeyError:
            logger.warning("Fallback style %r also missing; leaving default style.",
                           fallback)


def apply_style_to_block_element(doc, element, style_name, fallback="Normal") -> None:
    """Apply *style_name* to a raw body element (``<w:p>`` or ``<w:tbl>``).

    Used by the ``<!-- style: … -->`` directive to style block content after it has
    been rendered. The element is re-wrapped in its python-docx proxy so the style
    *name* is resolved to a style id correctly. Tables take no paragraph fallback.
    """
    tag = element.tag
    if tag.endswith('}p'):
        apply_style(Paragraph(element, doc._body), style_name, fallback)
    elif tag.endswith('}tbl'):
        apply_style(Table(element, doc._body), style_name, fallback=None)


def add_mapped_heading(doc, level, style_map=DEFAULT_STYLE_MAP):
    """Add a heading paragraph using *style_map*'s style for *level* (1-based).

    With default styles this is equivalent to ``doc.add_heading('', level)``.
    """
    para = doc.add_paragraph()
    apply_style(para, style_map.heading_style(level), fallback="Normal")
    return para


# Candidate locations for the global config (mirrors main.py resolution).
_CONFIG_PATHS = (
    Path("/app/config") / "docx_templates.yaml",
    Path(__file__).resolve().parent.parent / "config" / "docx_templates.yaml",
)
_cached_global_style_map: Optional[StyleMap] = None


def load_global_style_map() -> StyleMap:
    """Build the global :class:`StyleMap` from ``docx_templates.yaml`` (cached).

    Reads the top-level ``style_mapping`` section. Returns :data:`DEFAULT_STYLE_MAP`
    if no config file or section is present. Result is cached for the process.
    """
    global _cached_global_style_map
    if _cached_global_style_map is not None:
        return _cached_global_style_map
    mapping = {}
    for path in _CONFIG_PATHS:
        try:
            if not path.is_file():
                continue
            import yaml
            cfg = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
            mapping = cfg.get("style_mapping") or {}
            break
        except Exception:
            logger.warning("Failed to read style_mapping from %s", path, exc_info=True)
    _cached_global_style_map = build_style_map(mapping)
    return _cached_global_style_map
