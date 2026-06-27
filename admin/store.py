"""Storage abstraction for UI-managed dynamic templates (docx + email).

The admin UI never edits the heavily-documented master YAML files
(``config/docx_templates.yaml`` / ``config/email_templates.yaml``). Instead each
UI-managed template is persisted as a **single file per template** in a sibling
directory:

    config/docx_templates.d/<name>.yaml      + custom_templates/<docx_path>
    config/email_templates.d/<name>.yaml     + custom_templates/<html_path>

The loader (see :mod:`docx_tools.dynamic_docx_tools` /
:mod:`email_tools.dynamic_email_tools`) merges these per-template files on top of
the master YAML, so the master files stay pristine and UI edits produce clean,
reviewable diffs.

Everything goes through the :class:`TemplateStore` interface. The default
:class:`FileTemplateStore` reads/writes the mounted volume; a future
``DbTemplateStore`` (for multi-replica deployments) can replace it without
touching the admin views — they depend only on the interface.
"""
from __future__ import annotations

import logging
import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml

logger = logging.getLogger(__name__)

# Supported template kinds.
KIND_DOCX = "docx"
KIND_EMAIL = "email"

# Per-kind metadata: where managed spec files live, the asset extension, and the
# spec key that names the asset file.
_KIND_META: Dict[str, Dict[str, str]] = {
    KIND_DOCX: {"subdir": "docx_templates.d", "asset_ext": ".docx", "path_key": "docx_path"},
    KIND_EMAIL: {"subdir": "email_templates.d", "asset_ext": ".html", "path_key": "html_path"},
}

# Tool names must be safe identifiers — they become MCP tool names and file
# stems. Keep this conservative (letters, digits, underscore; not starting with
# a digit) so a name can never escape its directory or collide with YAML syntax.
_NAME_RE = re.compile(r"^[a-zA-Z_][a-zA-Z0-9_]*$")

# Default production / local directories (mirrors template_utils resolution).
_APP_CUSTOM_DIR = Path("/app/custom_templates")
_APP_CONFIG_DIR = Path("/app/config")
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_LOCAL_CUSTOM_DIR = _PROJECT_ROOT / "custom_templates"
_LOCAL_CONFIG_DIR = _PROJECT_ROOT / "config"


class TemplateStoreError(Exception):
    """Raised for invalid template operations (bad name, missing asset, …)."""


def valid_kind(kind: str) -> bool:
    """Return True if *kind* is a supported template kind."""
    return kind in _KIND_META


def _require_kind(kind: str) -> Dict[str, str]:
    meta = _KIND_META.get(kind)
    if meta is None:
        raise TemplateStoreError(f"Unknown template kind: {kind!r}")
    return meta


def validate_name(name: str) -> str:
    """Validate and return a template/tool *name*, or raise ``TemplateStoreError``."""
    if not name or not _NAME_RE.match(name):
        raise TemplateStoreError(
            f"Invalid template name {name!r}: use letters, digits and underscores "
            "and do not start with a digit."
        )
    return name


def validate_asset_filename(filename: str, kind: str) -> str:
    """Validate that *filename* is a bare filename with the kind's extension.

    Mirrors the filename-only guard enforced by the dynamic-tool loaders so the
    asset can never reference a directory or absolute path.
    """
    meta = _require_kind(kind)
    if not filename:
        raise TemplateStoreError("Asset filename must not be empty.")
    p = Path(filename)
    if p.is_absolute() or len(p.parts) != 1:
        raise TemplateStoreError(
            f"Asset filename must be a bare filename (no directories); got {filename!r}."
        )
    if p.suffix.lower() != meta["asset_ext"]:
        raise TemplateStoreError(
            f"{kind} asset must end with {meta['asset_ext']}; got {filename!r}."
        )
    return filename


class TemplateStore(ABC):
    """Interface the admin UI uses to persist and load managed templates.

    A "spec" is the plain ``dict`` that also appears as one entry under
    ``templates:`` in the master YAML (``name``, ``description``, ``args``, the
    asset path key, optional ``style_mapping`` / ``annotations`` …).
    """

    @abstractmethod
    def list_specs(self, kind: str) -> List[Dict[str, Any]]:
        """Return all managed specs for *kind*, sorted by name."""

    @abstractmethod
    def get_spec(self, kind: str, name: str) -> Optional[Dict[str, Any]]:
        """Return the managed spec named *name*, or ``None`` if absent."""

    @abstractmethod
    def save_spec(
        self,
        kind: str,
        spec: Dict[str, Any],
        asset_bytes: Optional[bytes] = None,
        asset_filename: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Persist *spec* (and optionally its asset). Returns the stored spec."""

    @abstractmethod
    def delete_spec(self, kind: str, name: str, delete_asset: bool = False) -> bool:
        """Delete the managed spec (and optionally its asset). Returns True if removed."""

    @abstractmethod
    def read_asset(self, kind: str, filename: str) -> bytes:
        """Return the bytes of an asset file (raises if not found)."""

    @abstractmethod
    def asset_exists(self, kind: str, filename: str) -> bool:
        """Return True if the asset file exists in the writable custom dir."""


class FileTemplateStore(TemplateStore):
    """Filesystem-backed store writing to ``custom_templates/`` and ``config/*.d/``.

    Args:
        custom_dir: Directory that holds template assets (the writable
            ``custom_templates``). Created on first write if missing.
        config_dir: Directory that holds the master YAML and the per-template
            ``*.d`` directories.
    """

    def __init__(self, custom_dir: Path, config_dir: Path):
        self.custom_dir = Path(custom_dir)
        self.config_dir = Path(config_dir)

    # -- construction -------------------------------------------------------

    @classmethod
    def from_config(cls) -> "FileTemplateStore":
        """Resolve writable directories the way the rest of the app does.

        Prefers the container mount points (``/app/...``) when they exist,
        otherwise the in-project directories for local development.
        """
        custom = _APP_CUSTOM_DIR if _APP_CUSTOM_DIR.exists() else _LOCAL_CUSTOM_DIR
        config = _APP_CONFIG_DIR if _APP_CONFIG_DIR.exists() else _LOCAL_CONFIG_DIR
        return cls(custom_dir=custom, config_dir=config)

    # -- internal paths -----------------------------------------------------

    def _spec_dir(self, kind: str) -> Path:
        return self.config_dir / _require_kind(kind)["subdir"]

    def _spec_path(self, kind: str, name: str) -> Path:
        return self._spec_dir(kind) / f"{validate_name(name)}.yaml"

    def asset_path(self, kind: str, filename: str) -> Path:
        validate_asset_filename(filename, kind)
        return self.custom_dir / filename

    # -- reads --------------------------------------------------------------

    def list_specs(self, kind: str) -> List[Dict[str, Any]]:
        spec_dir = self._spec_dir(kind)
        if not spec_dir.is_dir():
            return []
        specs: List[Dict[str, Any]] = []
        for path in sorted(spec_dir.glob("*.yaml")):
            spec = self._read_spec_file(path)
            if spec is not None:
                specs.append(spec)
        specs.sort(key=lambda s: str(s.get("name", "")))
        return specs

    def get_spec(self, kind: str, name: str) -> Optional[Dict[str, Any]]:
        path = self._spec_path(kind, name)
        if not path.is_file():
            return None
        return self._read_spec_file(path)

    @staticmethod
    def _read_spec_file(path: Path) -> Optional[Dict[str, Any]]:
        """Load one managed spec file.

        A managed file holds a single spec mapping. For resilience a file that
        instead wraps it as ``{templates: [spec]}`` is also accepted (first
        entry wins).
        """
        try:
            data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        except Exception as e:
            logger.error("[template-store] Failed to parse %s: %s", path, e)
            return None
        if isinstance(data, dict) and isinstance(data.get("templates"), list):
            templates = data["templates"]
            data = templates[0] if templates else {}
        if not isinstance(data, dict) or not data.get("name"):
            logger.warning("[template-store] Ignoring malformed spec file %s", path)
            return None
        return data

    def read_asset(self, kind: str, filename: str) -> bytes:
        path = self.asset_path(kind, filename)
        if not path.is_file():
            raise TemplateStoreError(f"Asset not found: {filename}")
        return path.read_bytes()

    def asset_exists(self, kind: str, filename: str) -> bool:
        try:
            return self.asset_path(kind, filename).is_file()
        except TemplateStoreError:
            return False

    # -- writes -------------------------------------------------------------

    def save_spec(
        self,
        kind: str,
        spec: Dict[str, Any],
        asset_bytes: Optional[bytes] = None,
        asset_filename: Optional[str] = None,
    ) -> Dict[str, Any]:
        meta = _require_kind(kind)
        if not isinstance(spec, dict):
            raise TemplateStoreError("spec must be a mapping.")

        spec = dict(spec)  # shallow copy; we may set the asset path key
        name = validate_name(str(spec.get("name", "")))

        path_key = meta["path_key"]
        # Determine the asset filename: explicit arg wins, else the spec's path
        # key, else derive from the template name.
        filename = asset_filename or spec.get(path_key) or f"{name}{meta['asset_ext']}"
        validate_asset_filename(filename, kind)
        spec[path_key] = filename

        # Write the asset first (so the spec never points at a missing file).
        if asset_bytes is not None:
            self.custom_dir.mkdir(parents=True, exist_ok=True)
            self.asset_path(kind, filename).write_bytes(asset_bytes)
        elif not self.asset_exists(kind, filename):
            raise TemplateStoreError(
                f"No asset bytes provided and asset {filename!r} does not exist yet."
            )

        spec_path = self._spec_path(kind, name)
        spec_path.parent.mkdir(parents=True, exist_ok=True)
        spec_path.write_text(self._dump_spec(spec), encoding="utf-8")
        logger.info("[template-store] Saved %s template %r -> %s", kind, name, spec_path)
        return spec

    def delete_spec(self, kind: str, name: str, delete_asset: bool = False) -> bool:
        spec_path = self._spec_path(kind, name)
        existed = spec_path.is_file()
        if existed:
            spec = self._read_spec_file(spec_path)
            spec_path.unlink()
            logger.info("[template-store] Deleted %s spec %r", kind, name)
            if delete_asset and spec:
                filename = spec.get(_require_kind(kind)["path_key"])
                if filename:
                    try:
                        asset = self.asset_path(kind, filename)
                        if asset.is_file():
                            asset.unlink()
                            logger.info("[template-store] Deleted asset %s", filename)
                    except TemplateStoreError:
                        pass
        return existed

    @staticmethod
    def _dump_spec(spec: Dict[str, Any]) -> str:
        """Serialise a spec to YAML with a header marking it UI-managed."""
        header = (
            "# Managed by the template-admin UI. Edits here are merged on top of\n"
            "# the master YAML at startup. Prefer editing via the admin UI.\n"
        )
        body = yaml.safe_dump(
            spec,
            sort_keys=False,
            allow_unicode=True,
            default_flow_style=False,
            width=4096,
        )
        return header + body
