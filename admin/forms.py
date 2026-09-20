"""Turn a submitted admin form into a template spec.

The inverse of :mod:`admin.views`: where the views render a spec as a form,
this reads a form back into the plain ``dict`` that :mod:`admin.store` persists
and the dynamic-tool loaders consume.

Kept free of FastHTML so the parsing rules — how a default is coerced to its
declared type, what an unticked checkbox means, which keys a spec omits when
they match the default — can be unit tested without rendering anything.
"""
from __future__ import annotations

import json
from typing import Any, Dict, List, Optional

from admin.kinds import STYLE_KEYS, descriptor
from admin.store import KIND_DOCX, KIND_PPTX, validate_name


def coerce_default(atype: str, raw: str) -> Any:
    """Coerce a string default from the form into the arg's declared type.

    A value that will not convert is kept as the raw string rather than
    rejected: the admin sees what they typed on the next render, which is more
    useful than an error page that discards the rest of the form.
    """
    raw = (raw or "").strip()
    atype = (atype or "string").lower()
    if atype in ("bool", "boolean"):
        return raw.lower() in ("1", "true", "yes", "on")
    if atype in ("int", "integer") and raw:
        try:
            return int(raw)
        except ValueError:
            return raw
    if atype == "float" and raw:
        try:
            return float(raw)
        except ValueError:
            return raw
    return raw


def parse_args_from_form(form) -> List[Dict[str, Any]]:
    """Build the ``args`` list from the parallel-indexed form fields."""
    names = form.getlist("arg_name")
    types = form.getlist("arg_type")
    reqs = form.getlist("arg_required")
    defs = form.getlist("arg_default")
    descs = form.getlist("arg_desc")
    args: List[Dict[str, Any]] = []
    for i, raw_name in enumerate(names):
        name = (raw_name or "").strip()
        if not name:
            continue
        atype = types[i] if i < len(types) else "string"
        required = (reqs[i] if i < len(reqs) else "true") == "true"
        default_raw = defs[i] if i < len(defs) else ""
        desc = descs[i] if i < len(descs) else ""
        arg: Dict[str, Any] = {
            "name": name, "type": atype, "required": required, "description": desc,
        }
        if not required or (default_raw or "").strip():
            arg["default"] = coerce_default(atype, default_raw)
        args.append(arg)
    return args


def parse_style_mapping(form) -> Dict[str, str]:
    """Collect non-default style-mapping selections."""
    mapping: Dict[str, str] = {}
    for key in STYLE_KEYS:
        val = (form.get(f"style_{key}") or "").strip()
        if val and val != "__default__":
            mapping[key] = val
    return mapping


def checked(form, field: str) -> bool:
    """True when a checkbox was submitted ticked.

    An unticked HTML checkbox sends nothing at all, so absence has to mean
    False rather than "unset".
    """
    return (form.get(field) or "").lower() in ("1", "true", "yes", "on")


def build_pptx_spec(form) -> Dict[str, Any]:
    """Assemble a PowerPoint template spec from the submitted edit form.

    Nothing like the docx/email spec: no ``args``, because a presentation
    template takes no arguments. What it carries instead is how the deck should
    be laid out — which layout plays each role, and the deck-wide defaults.
    """
    from pptx_tools.layouts import ROLES

    name = validate_name((form.get("name") or "").strip())
    spec: Dict[str, Any] = {
        "name": name,
        "description": (form.get("description") or "").strip(),
        "pptx_path": (form.get("asset_filename") or "").strip(),
    }
    if checked(form, "is_default"):
        spec["default"] = True
    # Only write strip_slides when it departs from the default, so a spec file
    # stays as small as what the admin actually chose.
    if not checked(form, "strip_slides"):
        spec["strip_slides"] = False

    layouts = {
        role: (form.get(f"layout_{role}") or "").strip()
        for role in ROLES
        if (form.get(f"layout_{role}") or "").strip()
    }
    if layouts:
        spec["layouts"] = layouts

    defaults: Dict[str, Any] = {}
    footer = (form.get("default_footer_text") or "").strip()
    if footer:
        defaults["footer_text"] = footer
    language = (form.get("default_language") or "").strip()
    if language:
        defaults["language"] = language
    if checked(form, "default_slide_numbers"):
        defaults["show_slide_numbers"] = True
    if defaults:
        spec["defaults"] = defaults
    return spec


#: Hidden field carrying an already-built spec between two POSTs. The preview
#: values form needs the spec the admin is editing — unsaved changes included
#: — without re-emitting every field `build_spec` reads and drifting from it.
CARRIED_SPEC_FIELD = "spec_json"


def carried_spec(form) -> Optional[Dict[str, Any]]:
    """The spec carried in `CARRIED_SPEC_FIELD`, or ``None``.

    Treated as untrusted like any other form field: anything that is not a
    named mapping is ignored and the caller rebuilds from the form instead.
    It only ever drives an in-memory preview render.
    """
    raw = form.get(CARRIED_SPEC_FIELD)
    if not raw:
        return None
    try:
        spec = json.loads(str(raw))
    except (TypeError, ValueError):
        return None
    if not isinstance(spec, dict) or not spec.get("name"):
        return None
    return spec


def build_spec(kind: str, form) -> Dict[str, Any]:
    """Assemble a template spec dict from the submitted edit form."""
    if kind == KIND_PPTX:
        return build_pptx_spec(form)
    desc = descriptor(kind)
    name = validate_name((form.get("name") or "").strip())
    description = (form.get("description") or "").strip()
    title = (form.get("title") or "").strip()

    spec: Dict[str, Any] = {"name": name, "description": description or f"Generate {name}"}
    if title:
        spec["annotations"] = {"title": title}
    spec[desc.path_key] = (form.get("asset_filename") or "").strip()
    spec["args"] = parse_args_from_form(form)
    if kind == KIND_DOCX:
        style_mapping = parse_style_mapping(form)
        if style_mapping:
            spec["style_mapping"] = style_mapping
    return spec
