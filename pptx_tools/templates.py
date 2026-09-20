"""Named PowerPoint template registry.

Before this, a deployment got exactly two slots — ``custom_pptx_template_4_3
.pptx`` and ``custom_pptx_template_16_9.pptx`` — discovered by filename and
cached for the life of the process. There was no way to offer a model a choice
of brand decks, no per-template configuration, no validation, and dropping a
new template into the mounted volume needed a restart.

This mirrors the pattern the docx and email tools already use: a documented
master ``config/pptx_templates.yaml`` merged with UI-managed per-template files
in ``config/pptx_templates.d/``, loaded through
:func:`template_registry.gather_specs`.

When no registry file exists the previous behaviour is preserved exactly: the
two filename slots are synthesised as specs named ``16_9`` and ``4_3``, so an
existing deployment sees no change.
"""

from __future__ import annotations

import io
import logging
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.util import Emu

import template_utils
from template_registry import gather_specs, is_enabled
from template_utils import find_file_in_template_dirs

from .constants import SLIDE_FORMAT_4_3, SLIDE_FORMAT_16_9

logger = logging.getLogger(__name__)


def candidate_template_dirs() -> List[Path]:
    """The directories template resolution searches, in priority order."""
    return list(template_utils._candidate_dirs())

MASTER_YAML_NAME = "pptx_templates.yaml"
SPEC_SUBDIR = "pptx_templates.d"

_APP_CONFIG_DIR = Path("/app/config")
_LOCAL_CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"

# Content types of a .potx and the .pptx equivalent. python-pptx refuses the
# template flavour outright ("is not a PowerPoint file, content type is ...")
# even though the package is otherwise identical, so a .potx is rewritten in
# memory rather than rejected — handing a designer's .potx back with that error
# is a poor answer when the fix is one string.
_POTX_CONTENT_TYPE = "presentationml.template.main+xml"
_PPTX_CONTENT_TYPE = "presentationml.presentation.main+xml"


@dataclass
class TemplateSpec:
    """One registered template."""

    name: str
    path: Path
    description: str = ""
    is_default: bool = False
    layouts: Dict[str, str] = field(default_factory=dict)
    defaults: Dict[str, Any] = field(default_factory=dict)
    strip_slides: bool = True
    aspect: Optional[str] = None

    def summary(self) -> Dict[str, Any]:
        """Serialisable description, for the template-listing tool."""
        return {
            "name": self.name,
            "description": self.description,
            "aspect": self.aspect,
            "default": self.is_default,
            "file": self.path.name,
        }


def config_dir() -> Path:
    """Production mount if present, else the in-project config directory."""
    return _APP_CONFIG_DIR if _APP_CONFIG_DIR.exists() else _LOCAL_CONFIG_DIR


# ---------------------------------------------------------------------------
# Loading a template file
# ---------------------------------------------------------------------------

def open_template(path) -> Presentation:
    """Open a .pptx or .potx as a Presentation.

    Raises:
        ValueError: If the file cannot be opened as a presentation at all.
    """
    path = Path(path)
    try:
        return Presentation(str(path))
    except ValueError as exc:
        if _POTX_CONTENT_TYPE not in str(exc):
            raise
        logger.debug("Rewriting .potx content type for %s", path.name)
        return Presentation(_potx_as_pptx(path))


def _potx_as_pptx(path: Path) -> io.BytesIO:
    """Copy a .potx into memory with the presentation content type."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(path) as source:
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = data.decode("utf-8").replace(
                        _POTX_CONTENT_TYPE, _PPTX_CONTENT_TYPE
                    ).encode("utf-8")
                target.writestr(item, data)
    buffer.seek(0)
    return buffer


def aspect_of(presentation: Presentation) -> str:
    """Report a presentation's aspect ratio from its actual slide size.

    Read rather than declared: a registry entry that claims 16:9 for a 4:3 file
    would otherwise mis-place every element on the slide.
    """
    width = Emu(presentation.slide_width).inches
    height = Emu(presentation.slide_height).inches
    if height <= 0:
        return SLIDE_FORMAT_16_9
    ratio = width / height
    return SLIDE_FORMAT_16_9 if ratio > 1.45 else SLIDE_FORMAT_4_3


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

def _fingerprint() -> Tuple:
    """Cheap signature of everything that can change the registry.

    Comparing this on each call is what replaces the process-lifetime cache,
    so dropping a new template into the mounted volume takes effect without a
    restart.

    Directory mtimes alone are not enough. A directory's mtime changes when an
    entry is added, removed or renamed — but *not* when an existing file's
    contents are replaced, which is exactly how a template gets updated in
    practice. Without the per-file mtimes below, overwriting brand.pptx with a
    deck of a different slide size left the cached aspect ratio stale, and
    selection by aspect then picked the wrong template.
    """
    parts: List[Tuple[str, Optional[int]]] = []
    cfg = config_dir()
    # Watch exactly the directories template resolution searches, rather than a
    # hardcoded list: otherwise the watched set and the searched set can drift
    # apart and an edit in a searched-but-unwatched directory goes unnoticed.
    template_dirs = list(candidate_template_dirs())

    watched: List[Path] = [cfg, cfg / MASTER_YAML_NAME, cfg / SPEC_SUBDIR]
    watched.extend(template_dirs)

    spec_dir = cfg / SPEC_SUBDIR
    if spec_dir.is_dir():
        watched.extend(sorted(spec_dir.glob("*.yaml")))

    for directory in template_dirs:
        if directory.is_dir():
            for pattern in ("*.pptx", "*.potx"):
                watched.extend(sorted(directory.glob(pattern)))

    for item in watched:
        try:
            parts.append((str(item), item.stat().st_mtime_ns))
        except OSError:
            parts.append((str(item), None))
    return tuple(parts)


_cache: Dict[str, Any] = {"fingerprint": None, "specs": []}


def _legacy_specs() -> List[TemplateSpec]:
    """Synthesise specs from the two historical filename slots."""
    specs: List[TemplateSpec] = []
    for name, candidates, aspect in (
        ("16_9", ("custom_pptx_template_16_9.pptx", "default_pptx_template_16_9.pptx"), SLIDE_FORMAT_16_9),
        ("4_3", ("custom_pptx_template_4_3.pptx", "default_pptx_template_4_3.pptx"), SLIDE_FORMAT_4_3),
    ):
        for filename in candidates:
            found = find_file_in_template_dirs(filename)
            if found:
                specs.append(TemplateSpec(
                    name=name,
                    path=found,
                    description=f"Built-in {aspect} template ({found.name}).",
                    is_default=(aspect == SLIDE_FORMAT_16_9),
                    aspect=aspect,
                ))
                break
    return specs


def _spec_from_yaml(entry: Dict[str, Any]) -> Optional[TemplateSpec]:
    name = entry.get("name")
    filename = entry.get("pptx_path") or entry.get("path")
    if not name or not filename:
        logger.error("[pptx-templates] Ignoring entry without a name and pptx_path: %r", entry)
        return None

    if Path(filename).name != filename:
        logger.error(
            "[pptx-templates] %s: pptx_path must be a bare filename, got %r", name, filename
        )
        return None

    found = find_file_in_template_dirs(filename)
    if not found:
        logger.error(
            "[pptx-templates] %s: %s not found in the template directories; skipping.",
            name, filename,
        )
        return None

    layouts = entry.get("layouts") or {}
    if not isinstance(layouts, dict):
        logger.error("[pptx-templates] %s: 'layouts' must be a mapping; ignoring it.", name)
        layouts = {}

    defaults = entry.get("defaults") or {}
    if not isinstance(defaults, dict):
        logger.error("[pptx-templates] %s: 'defaults' must be a mapping; ignoring it.", name)
        defaults = {}

    return TemplateSpec(
        name=str(name),
        path=found,
        description=str(entry.get("description") or ""),
        is_default=bool(entry.get("default")),
        layouts={str(k): str(v) for k, v in layouts.items()},
        defaults=defaults,
        strip_slides=bool(entry.get("strip_slides", True)),
    )


def load_specs(force: bool = False) -> List[TemplateSpec]:
    """Return the registered templates, reloading when anything changed."""
    fingerprint = _fingerprint()
    if not force and _cache["fingerprint"] == fingerprint and _cache["specs"]:
        return _cache["specs"]

    cfg = config_dir()
    master = cfg / MASTER_YAML_NAME
    spec_dir = cfg / SPEC_SUBDIR

    specs: List[TemplateSpec] = []
    all_disabled = False
    if master.is_file() or spec_dir.is_dir():
        all_entries, _ = gather_specs(master, spec_dir, include_disabled=True)
        entries = [e for e in all_entries if is_enabled(e)]
        all_disabled = bool(all_entries) and not entries
        for entry in entries:
            spec = _spec_from_yaml(entry)
            if spec is not None:
                specs.append(spec)

    # The historical filename slots are the safety net for "nothing readable is
    # configured" — a missing config, or entries too malformed to build a spec
    # from. Turning every template off is neither: it is a deliberate choice,
    # and answering it by resurrecting the built-ins would hand back the very
    # templates the admin just took away (#165).
    if not specs and not all_disabled:
        specs = _legacy_specs()
    elif specs:
        # Fill in the aspect from each file, and make sure exactly one default.
        for spec in specs:
            try:
                spec.aspect = aspect_of(open_template(spec.path))
            except Exception as e:
                logger.error("[pptx-templates] %s: could not open %s: %s", spec.name, spec.path.name, e)
        if not any(spec.is_default for spec in specs):
            specs[0].is_default = True

    _cache["fingerprint"] = fingerprint
    _cache["specs"] = specs
    return specs


def clear_cache() -> None:
    """Drop the cached registry (tests, and the admin UI after a save)."""
    _cache["fingerprint"] = None
    _cache["specs"] = []


def template_names() -> List[str]:
    return [spec.name for spec in load_specs()]


def select_template(name: Optional[str] = None, format: Optional[str] = None):
    """Pick a template by name, else by aspect ratio, else the default.

    Returns ``(spec, warning)``; *spec* is None only when no template file
    could be found at all, in which case the caller falls back to python-pptx's
    built-in theme.
    """
    specs = load_specs()
    if not specs:
        return None, "no PowerPoint template is configured; used the built-in theme."

    if name:
        for spec in specs:
            if spec.name == name:
                return spec, None
        available = ", ".join(spec.name for spec in specs)
        fallback, _ = select_template(None, format)
        return fallback, (
            f"template {name!r} is not registered (available: {available}); "
            f"used {fallback.name!r}."
        )

    if format:
        for spec in specs:
            if spec.aspect == format:
                return spec, None
        # Nothing in that aspect: better to say so than silently reshape the deck.
        fallback = next((s for s in specs if s.is_default), specs[0])
        return fallback, (
            f"no registered template is {format}; used {fallback.name!r} "
            f"({fallback.aspect})."
        )

    return next((spec for spec in specs if spec.is_default), specs[0]), None


# ---------------------------------------------------------------------------
# Startup validation
# ---------------------------------------------------------------------------

def content_area_summary(resolver, presentation) -> Optional[Dict[str, str]]:
    """The template's content rectangle as percentages of the slide.

    Percentages rather than inches because that is what a caller positioning
    a ``blank`` slide's elements writes, and they carry across aspect ratios.
    A blank layout usually keeps the template's logo, rules and footer, so
    this is the band inside them (#119). Public because the admin UI reports
    the same rectangle in the same units (#171); one formatter means the page
    and the tool's own diagnostics cannot disagree about it.
    """
    rect = resolver.content_area()
    if rect is None:
        return None

    width, height = presentation.slide_width, presentation.slide_height
    if not width or not height:
        return None

    def share(value: int, total: int) -> str:
        return f"{round(value / total * 100, 1)}%"

    return {
        "layout": rect.source,
        "x": share(rect.left, width),
        "y": share(rect.top, height),
        "w": share(rect.width, width),
        "h": share(rect.height, height),
    }


def validate_templates() -> List[Dict[str, Any]]:
    """Open every registered template once and report what it provides.

    Called at startup so a template missing the layouts the tool needs is
    visible in the log then, rather than as a strangely laid-out deck later.
    """
    from .layouts import LayoutResolver  # imported here to avoid a cycle

    reports: List[Dict[str, Any]] = []
    for spec in load_specs():
        report: Dict[str, Any] = {"name": spec.name, "file": spec.path.name}
        try:
            presentation = open_template(spec.path)
        except Exception as e:
            report["error"] = str(e)
            reports.append(report)
            logger.error("[pptx-templates] %s: cannot open %s: %s", spec.name, spec.path.name, e)
            continue

        resolver = LayoutResolver(presentation, spec.layouts)
        report["aspect"] = aspect_of(presentation)
        # Each layout with the role it classified as, not just its name: a
        # template whose section layout carries only a title is detected as
        # title_only, which then also serves kpi and timeline slides. That is
        # invisible from a list of names, and diagnosing it meant reading
        # layouts.py and running classify_layout() by hand (#121).
        report["layouts"] = resolver.describe()
        report["coverage"] = resolver.coverage()
        report["missing_roles"] = resolver.missing_roles()
        report["layouts_without_footer"] = resolver.layouts_without_footer()
        report["content_area"] = content_area_summary(resolver, presentation)

        # A configured name that is not in the file is a silent mis-render
        # waiting to happen, so name it explicitly.
        unknown = [
            f"{role}->{layout_name}"
            for role, layout_name in spec.layouts.items()
            if resolver.by_name(layout_name) is None
        ]
        if unknown:
            report["unknown_configured_layouts"] = unknown
            logger.warning(
                "[pptx-templates] %s: configured layouts not found in the file: %s",
                spec.name, ", ".join(unknown),
            )

        if report["missing_roles"]:
            logger.warning(
                "[pptx-templates] %s (%s): no layout matches %s; those slide types "
                "will fall back by position.",
                spec.name, spec.path.name, ", ".join(report["missing_roles"]),
            )
        else:
            logger.info(
                "[pptx-templates] %s (%s, %s): all slide-type layouts resolved.",
                spec.name, spec.path.name, report["aspect"],
            )

        reports.append(report)

    return reports
