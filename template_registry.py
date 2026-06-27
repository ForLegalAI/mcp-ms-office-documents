"""Shared helpers for loading and live-(un)registering dynamic template tools.

Both the docx and email dynamic-tool modules use these to:

* merge the heavily-documented master YAML (``config/<kind>_templates.yaml``)
  with the UI-managed per-template files in ``config/<kind>_templates.d/`` —
  the per-template file wins when names collide; and
* remove a live MCP tool by name (so a template can be re-registered after an
  edit, or unregistered on delete) tolerantly across FastMCP versions.

Keeping this here (a sibling of ``template_utils.py``) avoids a dependency from
the core dynamic-tool modules onto the optional ``admin`` package.
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import yaml

logger = logging.getLogger(__name__)


def safe_remove_tool(mcp, name: str) -> bool:
    """Remove a live MCP tool *name* if present.

    Tolerant of FastMCP version differences (prefers ``local_provider`` to
    avoid the deprecation warning on ``mcp.remove_tool``) and of the tool not
    existing. Returns True only when a tool was actually removed.
    """
    provider = getattr(mcp, "local_provider", None)
    remover = getattr(provider, "remove_tool", None) if provider else None
    if remover is None:
        remover = getattr(mcp, "remove_tool", None)
    if remover is None:  # pragma: no cover - unexpected FastMCP build
        logger.warning("[template-registry] No remove_tool available on MCP instance.")
        return False
    try:
        remover(name)
        return True
    except Exception as e:
        logger.debug("[template-registry] remove_tool(%r) no-op: %s", name, e)
        return False


def _read_one_spec(path: Path) -> Optional[Dict[str, Any]]:
    """Load a single per-template ``*.yaml`` file into a spec dict.

    Accepts either a bare spec mapping or a ``{templates: [spec]}`` wrapper.
    Returns ``None`` (with a log) for unreadable / malformed files so one bad
    file never aborts the whole load.
    """
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except Exception as e:
        logger.error("[template-registry] Failed to parse %s: %s", path, e)
        return None
    if isinstance(data, dict) and isinstance(data.get("templates"), list):
        templates = data["templates"]
        data = templates[0] if templates else {}
    if not isinstance(data, dict) or not data.get("name"):
        logger.warning("[template-registry] Ignoring malformed spec file %s", path)
        return None
    return data


def read_spec_dir(spec_dir: Path) -> List[Dict[str, Any]]:
    """Return the specs in ``spec_dir`` (one per ``*.yaml``), sorted by filename."""
    if not spec_dir or not spec_dir.is_dir():
        return []
    specs: List[Dict[str, Any]] = []
    for path in sorted(spec_dir.glob("*.yaml")):
        spec = _read_one_spec(path)
        if spec is not None:
            specs.append(spec)
    return specs


def gather_specs(
    master_yaml: Optional[Path], spec_dir: Optional[Path]
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Merge the master YAML's ``templates`` with the per-template ``spec_dir``.

    Returns ``(templates, master_cfg)`` where:

    * ``templates`` is the merged list — entries from ``spec_dir`` override
      master entries with the same ``name`` (and new names are appended), with
      original master order preserved and dir-only templates appended after; and
    * ``master_cfg`` is the parsed master mapping (``{}`` if absent), so callers
      can still read top-level keys such as ``style_mapping``.
    """
    master_cfg: Dict[str, Any] = {}
    master_templates: List[Dict[str, Any]] = []
    if master_yaml and master_yaml.is_file():
        try:
            master_cfg = yaml.safe_load(master_yaml.read_text(encoding="utf-8")) or {}
        except Exception as e:
            logger.error("[template-registry] Failed to load master YAML %s: %s", master_yaml, e)
            master_cfg = {}
        raw = master_cfg.get("templates")
        if isinstance(raw, list):
            master_templates = [t for t in raw if isinstance(t, dict)]
        elif raw is not None:
            logger.error("[template-registry] 'templates' in %s must be a list.", master_yaml)

    overrides = {
        s["name"]: s for s in read_spec_dir(spec_dir) if isinstance(s, dict) and s.get("name")
    } if spec_dir else {}

    merged: List[Dict[str, Any]] = []
    seen = set()
    for spec in master_templates:
        name = spec.get("name")
        if name in overrides:
            merged.append(overrides[name])  # dir wins
        else:
            merged.append(spec)
        seen.add(name)
    # Append dir-only templates (not present in the master) in name order.
    for name in sorted(overrides):
        if name not in seen:
            merged.append(overrides[name])

    return merged, master_cfg
