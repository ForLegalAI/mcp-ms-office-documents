"""Render a managed template with sample values for in-UI preview.

Preview never touches the configured upload backend — it renders the document
to an in-memory buffer and returns the bytes, so an admin can sanity-check a
template even on a server configured for S3/GCS/etc.

The docx path reuses the exact production substitution pipeline
(``resolve_conditionals`` + ``_replace_placeholders_in_document``) so what the
preview shows matches what the live tool produces. The email path mirrors the
dynamic email tool's pystache rendering.
"""
from __future__ import annotations

import io
from pathlib import Path
from typing import Any, Dict, List

import pystache
from docx import Document as DocxDocument

from docx_tools.conditionals import resolve_conditionals
from docx_tools.dynamic_docx_tools import replace_placeholders_in_document
from docx_tools.style_map import build_style_map


def sample_values(args: List[Dict[str, Any]], conditionals: List[str] = None) -> Dict[str, Any]:
    """Build a plausible sample value for each declared arg.

    Strings use their declared default when non-empty, else a bracketed name
    like ``[recipient_name]`` so the placeholder is obvious in the output.
    Booleans (including conditional flags) default to True so conditional blocks
    are visible in the preview.
    """
    conditionals = set(conditionals or [])
    values: Dict[str, Any] = {}
    for arg in args or []:
        if not isinstance(arg, dict):
            continue
        name = arg.get("name")
        if not name:
            continue
        atype = str(arg.get("type", "string")).lower()
        default = arg.get("default")
        if atype in ("bool", "boolean"):
            values[name] = True if default in (None, "") else bool(default)
        elif atype in ("int", "integer"):
            values[name] = default if isinstance(default, int) else 1
        elif atype == "float":
            values[name] = default if isinstance(default, (int, float)) else 1.0
        else:
            values[name] = default if (default not in (None, "")) else f"[{name}]"
    # Any conditional without a declared arg still defaults to shown.
    for cond in conditionals:
        values.setdefault(cond, True)
    return values


def render_docx_preview(
    template_bytes: bytes,
    spec: Dict[str, Any],
    values: Dict[str, Any],
    global_style_mapping: Dict[str, Any] = None,
) -> bytes:
    """Render a docx template with *values*; return the generated ``.docx`` bytes."""
    doc = DocxDocument(io.BytesIO(template_bytes))
    style_map = build_style_map(global_style_mapping, spec.get("style_mapping"))

    resolve_conditionals(doc, values)
    context = {k: ("" if v is None else str(v)) for k, v in values.items()}
    replace_placeholders_in_document(doc, context, style_map)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# A fixed sample deck, rendered through whatever template is being previewed.
# Chosen to exercise the layouts an admin actually needs to check: the title
# layout, a section divider, bulleted content, a two-column split, a drawn slide
# that uses no placeholder at all, and the closing slide. Six slides keeps the
# preview quick to open and quick to scan; a longer deck buries the problem.
SAMPLE_DECK = [
    {
        "type": "title",
        "title": "Quarterly Business Review",
        "subtitle": "Sample deck — rendered through this template",
    },
    {"type": "section", "title": "Where we stand"},
    {
        "type": "content",
        "title": "Highlights",
        "body": (
            "- Revenue **ahead of plan** for the third quarter running\n"
            "- Churn down to 18%, the lowest since launch\n"
            "  - Enterprise renewals carried most of the improvement\n"
            "- One risk: onboarding time is still climbing"
        ),
    },
    {
        "type": "two_column",
        "title": "What worked, what did not",
        "left": {"heading": "Worked", "body": "- Partner channel\n- Pricing change"},
        "right": {"heading": "Did not", "body": "- Self-serve funnel\n- Support backlog"},
    },
    {
        "type": "kpi",
        "title": "At a glance",
        "items": [
            {"value": "€4.2M", "label": "ARR", "delta": "+12% vs Q2"},
            {"value": "18%", "label": "Churn", "delta": "−3pp"},
            {"value": "94", "label": "NPS"},
        ],
    },
    {
        "type": "closing",
        "title": "Thank you",
        "subtitle": "Questions welcome",
        "contact": ["hello@example.com"],
    },
]


def render_pptx_preview(
    template_path,
    spec: Dict[str, Any],
    slides: List[Dict[str, Any]] = None,
) -> tuple:
    """Build the sample deck on a PowerPoint template; return ``(bytes, lines)``.

    Goes through :class:`~pptx_tools.slide_builder.PowerpointPresentation` with a
    spec built from the submitted form, so the preview exercises the same layout
    resolution, role mapping and defaults the live tool will use — including any
    layout overrides the admin has just typed but not yet saved. The warnings
    the builder produces are handed back, because "which slides could not be
    laid out on this template" is the single most useful thing a preview can
    tell an admin — as their rendered lines, since they go straight into a
    response header for a person to read.
    """
    from pptx_tools.slide_builder import PowerpointPresentation
    from pptx_tools.templates import TemplateSpec, aspect_of, open_template

    path = Path(template_path)
    aspect = aspect_of(open_template(path))

    template_spec = TemplateSpec(
        name=str(spec.get("name") or "preview"),
        path=path,
        description=str(spec.get("description") or ""),
        layouts=dict(spec.get("layouts") or {}),
        defaults=dict(spec.get("defaults") or {}),
        strip_slides=bool(spec.get("strip_slides", True)),
        aspect=aspect,
    )

    presentation = PowerpointPresentation(
        slides if slides is not None else SAMPLE_DECK,
        format=aspect,
        template_spec=template_spec,
    )
    return presentation.save().getvalue(), presentation.warning_messages


def render_email_preview(
    template_bytes: bytes,
    spec: Dict[str, Any],
    values: Dict[str, Any],
) -> str:
    """Render an email HTML template with *values*; return rendered HTML."""
    html_source = template_bytes.decode("utf-8", errors="replace")
    safe = {k: ("" if v is None else v) for k, v in values.items()}
    return pystache.Renderer(file_encoding="utf-8").render(html_source, safe)
