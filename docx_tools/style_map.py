"""Style mapping for markdown → DOCX rendering (issue #66, part A).

The renderer applies Word paragraph styles by name (``List Number``, ``Quote``,
``Table Grid``, ``Heading N`` …). Custom templates may name their styles
differently, so :class:`StyleMap` lets those built-in names be remapped without
touching the call sites — the defaults reproduce today's behaviour exactly.

A map is threaded explicitly through the processors (not held in global state) so
concurrent conversions on worker threads never share mutable mapping state. Config
overrides come from the ``style_mapping`` section of ``config/docx_templates.yaml``
(global, overridable by the admin UI through
``config/docx_templates.d/_global.yaml`` — see
:func:`template_registry.global_config`) and each template's own
``style_mapping`` (per-template, wins over global).
See docs/development/tools/word.md ("Style mapping") for the design rationale;
the original discussion is issue #66.
"""
import logging
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Optional, Tuple

from docx.table import Table
from docx.text.paragraph import Paragraph

from template_registry import (
    GLOBAL_SETTINGS_FILENAME, global_config, spec_dir_for,
)

from . import warnings as W

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

#: Every ``style_mapping`` key :func:`_normalize` acts on. Public because the
#: admin UI's style editor must offer exactly these — offering fewer hides part
#: of the feature, offering more promises a setting that silently does nothing.
#: ``tests/test_admin_style_keys.py`` compares the two.
RECOGNISED_KEYS = frozenset(_HEADING_KEYS) | frozenset(_LIST_NUMBER_KEYS) \
    | frozenset(_LIST_BULLET_KEYS) | frozenset(_SCALAR_KEYS)


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


def apply_style(obj, style_name, fallback="Normal", warnings=None, line=None) -> None:
    """Set ``obj.style`` to *style_name*, falling back if the style is missing.

    *obj* is a paragraph or table. A missing style raises ``KeyError`` in
    python-docx; we log and fall back instead of letting it abort the render.

    *warnings* is the build's :class:`~warning_channel.WarningChannel`. A
    fallback means the document does not look as the caller asked, which is
    only visible to them if it is reported — so the substitution goes on the
    channel as well as into the log. The channel de-duplicates, so a template
    missing ``List Number`` reports once, not once per list item.

    *line* is the 1-based source line, known only where the caller named the
    style itself — a ``<!-- style: … -->`` directive.
    """
    if not style_name:
        return
    try:
        obj.style = style_name
        return
    except KeyError:
        fallback_desc = repr(fallback) if fallback else "the document default"
        logger.warning("Style %r not found in document; falling back to %s.",
                       style_name, fallback_desc)
        if warnings is not None:
            warnings.add(
                W.STYLE_MISSING,
                f"style '{style_name}' is not defined in the Word template; "
                f"{fallback or 'the document default'} was used instead.",
                line=line,
            )
    if fallback and fallback != style_name:
        try:
            obj.style = fallback
        except KeyError:
            logger.warning("Fallback style %r also missing; leaving default style.",
                           fallback)
            if warnings is not None:
                warnings.add(
                    W.STYLE_FALLBACK_MISSING,
                    f"fallback style '{fallback}' is missing from the Word "
                    f"template too; the document default was used.",
                    line=line,
                )


def apply_style_to_block_element(doc, element, style_name, fallback="Normal",
                                 warnings=None, line=None) -> None:
    """Apply *style_name* to a raw body element (``<w:p>`` or ``<w:tbl>``).

    Used by the ``<!-- style: … -->`` directive to style block content after it has
    been rendered. The element is re-wrapped in its python-docx proxy so the style
    *name* is resolved to a style id correctly. Tables take no paragraph fallback.
    """
    tag = element.tag
    if tag.endswith('}p'):
        apply_style(Paragraph(element, doc._body), style_name, fallback,
                    warnings=warnings, line=line)
    elif tag.endswith('}tbl'):
        apply_style(Table(element, doc._body), style_name, fallback=None,
                    warnings=warnings, line=line)


def add_mapped_heading(doc, level, style_map=DEFAULT_STYLE_MAP, warnings=None):
    """Add a heading paragraph using *style_map*'s style for *level* (1-based).

    With default styles this is equivalent to ``doc.add_heading('', level)``.
    """
    para = doc.add_paragraph()
    apply_style(para, style_map.heading_style(level), fallback="Normal",
                warnings=warnings)
    return para


# Candidate locations for the global config (mirrors main.py resolution).
_CONFIG_PATHS = (
    Path("/app/config") / "docx_templates.yaml",
    Path(__file__).resolve().parent.parent / "config" / "docx_templates.yaml",
)


def _resolve_config() -> Tuple[Optional[Path], Optional[Path]]:
    """``(master YAML, spec dir)`` for the first candidate location in use.

    A location counts as in use when *either* half is there: the admin UI can
    write ``docx_templates.d/_global.yaml`` on a deployment that never had a
    master YAML, and that mapping still has to apply.
    """
    for path in _CONFIG_PATHS:
        spec_dir = spec_dir_for(path)
        try:
            if path.is_file() or (spec_dir and spec_dir.is_dir()):
                return path, spec_dir
        except OSError:  # pragma: no cover - unreadable candidate directory
            logger.warning("Could not stat %s", path, exc_info=True)
    return None, None


def _stamp(path: Optional[Path]) -> Tuple:
    """A cheap fingerprint of one config file: absent, or its mtime and size."""
    if path is None:
        return ()
    try:
        st = path.stat()
        return (str(path), st.st_mtime_ns, st.st_size)
    except OSError:
        return (str(path), None, None)


#: ``(fingerprint, map)``. A plain tuple assignment, so two threads racing here
#: both compute the same answer and one wins — never a half-built cache.
_cached_global_style_map: Optional[Tuple[Tuple, StyleMap]] = None


def invalidate_global_style_map() -> None:
    """Drop the cached global map, so the next read re-resolves from disk.

    Called by whoever *writes* the mapping — the admin UI — because the
    fingerprint below cannot be trusted to notice two writes inside one
    filesystem timestamp tick when the file happens to keep its size.
    Hand edits on the volume are what the fingerprint is for; a write we made
    ourselves we simply know about.
    """
    global _cached_global_style_map
    _cached_global_style_map = None


def load_global_style_map() -> StyleMap:
    """Build the global :class:`StyleMap` from the docx template config.

    Reads the top-level ``style_mapping`` in force — the master
    ``docx_templates.yaml`` under whatever ``docx_templates.d/_global.yaml``
    overrides — through the same :func:`template_registry.global_config` the
    dynamic loader and the admin UI use, so the static Word tool cannot end up
    on a different mapping from a template tool.

    Cached against the two files' mtime and size, *not* for the life of the
    process as it used to be: that meant editing the mapping in the admin UI
    left every subsequent document on the old one until a restart — the
    setting with the widest blast radius being the one that could not be
    changed live (#161). Re-parsing unconditionally would be correct too, but
    the master file is a few hundred lines of worked examples and parsing it
    costs more than the markdown render it would be paying for.

    Two ``stat`` calls, once per document build
    (:mod:`docx_tools.base_docx_tool` calls this only when the caller passed
    no map). A caller generating in a loop should still build the map once and
    pass it down, which is what every dynamic template tool already does.
    """
    global _cached_global_style_map

    master, spec_dir = _resolve_config()
    settings = (spec_dir / GLOBAL_SETTINGS_FILENAME) if spec_dir else None
    fingerprint = (_stamp(master), _stamp(settings))

    cached = _cached_global_style_map
    if cached is not None and cached[0] == fingerprint:
        return cached[1]

    try:
        mapping = global_config(master, spec_dir).get("style_mapping") or {}
    except Exception:
        logger.warning("Failed to read the global style_mapping; using defaults.",
                       exc_info=True)
        mapping = {}
    style_map = build_style_map(mapping)
    _cached_global_style_map = (fingerprint, style_map)
    return style_map
