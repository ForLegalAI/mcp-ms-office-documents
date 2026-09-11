"""Template file resolution shared by every document tool.

Templates are searched in four directories, most specific first: the
container's ``/app/custom_templates``, the checkout's ``custom_templates/``,
then the same pair for ``default_templates/``. A custom file therefore always
wins over the shipped default of the same name, and the same code runs in the
container and in a local checkout. Never hard-code a template path; call
``find_file_in_template_dirs`` or one of the ``find_*_template`` helpers.
See docs/development/shared-modules.md.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional, Tuple
import logging

logger = logging.getLogger(__name__)

# Base directory of the project (this file lives at project root)
BASE_DIR = Path(__file__).resolve().parent

# Production mount points inside container
APP_CUSTOM_DIR = Path("/app/custom_templates")
APP_DEFAULT_DIR = Path("/app/default_templates")

# Local development directories
LOCAL_CUSTOM_DIR = BASE_DIR / "custom_templates"
LOCAL_DEFAULT_DIR = BASE_DIR / "default_templates"


def _candidate_dirs() -> list[Path]:
    """Return template search directories in priority order.

    Order:
    1) /app/custom_templates (production)
    2) <project>/custom_templates (local dev)
    3) /app/default_templates (production)
    4) <project>/default_templates (local dev)
    """
    return [
        APP_CUSTOM_DIR,
        LOCAL_CUSTOM_DIR,
        APP_DEFAULT_DIR,
        LOCAL_DEFAULT_DIR,
    ]


def _classify_template_source(p: Path) -> str:
    """Classify the resolved template path as 'custom' or 'default'.

    Falls back to 'unknown' if not under expected template directories.
    """
    parts = {part.lower() for part in p.parts}
    if "custom_templates" in parts:
        return "custom"
    if "default_templates" in parts:
        return "default"
    return "unknown"


def find_file_in_template_dirs(filename: str) -> Optional[Path]:
    """Find a file by name in custom/default template directories.

    Returns the first existing Path or None if not found.
    Emits an INFO log when a template is selected indicating whether it is
    a custom or default template, and from which directory it was loaded.
    """
    for d in _candidate_dirs():
        p = d / filename
        if p.exists():
            source = _classify_template_source(p)
            # Visible by default (INFO): announce which template was chosen
            logger.info("Using %s template: %s (from %s)", source, p.name, p.parent)
            # Verbose trace (DEBUG): full resolved path detail
            logger.debug("Template resolved: %s", p)
            return p
    logger.debug("Template not found in search paths: %s", filename)
    return None


def _resolve_from_candidates(filenames: list[str]) -> Optional[str]:
    """Try a list of candidate filenames (in order) across template dirs.

    Returns first match as a string path, or None if none exist.
    """
    for name in filenames:
        p = find_file_in_template_dirs(name)
        if p is not None:
            return str(p)
    return None


def find_pptx_templates() -> Tuple[Optional[str], Optional[str]]:
    """Resolve PPTX templates for 4:3 and 16:9 using strict new naming.

    For each aspect, priority is:
    - custom_pptx_template_<aspect>.pptx
    - default_pptx_template_<aspect>.pptx

    Returns tuple[str|None, str|None] for (4:3, 16:9).
    """
    t43 = _resolve_from_candidates([
        "custom_pptx_template_4_3.pptx",
        "default_pptx_template_4_3.pptx",
    ])

    t169 = _resolve_from_candidates([
        "custom_pptx_template_16_9.pptx",
        "default_pptx_template_16_9.pptx",
    ])

    return (t43, t169)


def find_docx_template() -> Optional[str]:
    """Resolve DOCX template using strict new naming.

    Order:
    1) custom_docx_template.docx
    2) default_docx_template.docx
    """
    return _resolve_from_candidates([
        "custom_docx_template.docx",
        "default_docx_template.docx",
    ])


def find_email_template(filename: Optional[str] = None) -> Optional[str]:
    """Resolve HTML email template.

    Behavior:
    - If filename is provided, search exactly that filename across template dirs.
    - If filename is None, use new naming convention:
      custom_email_template.html -> default_email_template.html
    """
    if filename:
        found = find_file_in_template_dirs(filename)
        return str(found) if found else None

    return _resolve_from_candidates([
        "custom_email_template.html",
        "default_email_template.html",
    ])
