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


def render_email_preview(
    template_bytes: bytes,
    spec: Dict[str, Any],
    values: Dict[str, Any],
) -> str:
    """Render an email HTML template with *values*; return rendered HTML."""
    html_source = template_bytes.decode("utf-8", errors="replace")
    safe = {k: ("" if v is None else v) for k, v in values.items()}
    # Mirror the dynamic email tool's convenience promo block, if present.
    if "promo_code" in safe and "promo_code_block" not in safe:
        promo = safe.get("promo_code")
        safe["promo_code_block"] = (
            f'<div class="promo">Use promo code <strong>{promo}</strong>.</div>' if promo else ""
        )
    return pystache.Renderer(file_encoding="utf-8").render(html_source, safe)
