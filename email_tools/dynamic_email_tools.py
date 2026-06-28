"""Dynamic registration of email draft MCP tools (Mustache-only, simplified).

Updated assumptions:
  - YAML lives in `config/` (e.g. config/email_templates.yaml)
  - Each template's `html_path` is ONLY a filename (no path separators)
  - The HTML file is resolved via custom_templates first, then default_templates
    (and corresponding /app/* paths in production)
"""
from __future__ import annotations

import io
import pystache
import logging
import threading
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
from typing import Any, Dict, Optional, Literal
from pydantic import Field, create_model
from fastmcp import FastMCP
from upload_tools import upload_file
from template_utils import find_email_template
from template_registry import gather_specs, safe_remove_tool
from async_runner import run_blocking
import metrics
from fastmcp.exceptions import ToolError


__all__ = [
    "register_email_template_tools_from_yaml",
    "register_email_template",
    "unregister_email_template",
    "registered_email_template_names",
]

logger = logging.getLogger(__name__)

# Live registry of email template tools registered on the running server,
# mapping tool name -> the spec it was built from.
_REGISTERED_EMAIL: Dict[str, Dict[str, Any]] = {}
_REG_LOCK = threading.Lock()

TYPE_MAP = {
    "string": str, "str": str,
    "int": int, "integer": int,
    "float": float,
    "bool": bool, "boolean": bool,
    "list": list[str], "list[str]": list[str], "list[string]": list[str],
    "dict": dict, "object": dict,
}

BASE_FIELDS: Dict[str, Any] = {
    "subject": (str, Field(..., description="Email subject line (also sets Subject header)")),
    "to": (Optional[list[str]], Field(None, description="List of recipient email addresses")),
    "cc": (Optional[list[str]], Field(None, description="List of CC recipient email addresses")),
    "bcc": (Optional[list[str]], Field(None, description="List of BCC recipient email addresses")),
}


def email_spec_dir(yaml_path: Path) -> Path:
    """Directory holding UI-managed per-template email specs, beside *yaml_path*."""
    return yaml_path.parent / "email_templates.d"


def registered_email_template_names() -> list[str]:
    """Return the names of email template tools currently registered."""
    with _REG_LOCK:
        return sorted(_REGISTERED_EMAIL)


def register_email_template_tools_from_yaml(mcp: FastMCP, yaml_path: Path) -> None:
    """Register dynamic email template tools from YAML.

    Merges the master YAML at *yaml_path* with any UI-managed per-template files
    in the sibling ``email_templates.d/`` directory (the per-template file wins
    on a name clash). Either source may be absent.
    """
    templates, _cfg = gather_specs(yaml_path, email_spec_dir(yaml_path))
    for spec in templates:
        try:
            register_email_template(mcp, spec)
        except Exception as e:  # pragma: no cover
            name = spec.get("name", "<unknown>") if isinstance(spec, dict) else "<unknown>"
            logger.error(f"[dynamic-email] Failed to register template '{name}': {e}")


def register_email_template(mcp: FastMCP, spec: Dict[str, Any]) -> bool:
    """Register (or replace) a single email template tool on the live server.

    Removes any existing tool of the same name first, so it is safe to call
    after startup to apply an edited template. Returns True on success.
    """
    name = spec.get("name") if isinstance(spec, dict) else None
    if name:
        safe_remove_tool(mcp, name)
        with _REG_LOCK:
            _REGISTERED_EMAIL.pop(name, None)
    return _register_single_email_template(mcp, spec)


def unregister_email_template(mcp: FastMCP, name: str) -> bool:
    """Remove an email template tool from the live server. Returns True if removed."""
    removed = safe_remove_tool(mcp, name)
    with _REG_LOCK:
        _REGISTERED_EMAIL.pop(name, None)
    return removed


def _register_single_email_template(mcp: FastMCP, spec: Dict[str, Any]) -> bool:
    """Build and register one email template tool. Returns True on success."""
    if not isinstance(spec, dict) or not spec.get("name"):
        logger.warning("[dynamic-email] Template missing 'name', skipping.")
        return False
    name = spec["name"]
    description = spec.get("description")
    annotations = spec.get("annotations")
    meta = spec.get("meta")
    html_path = spec.get("html_path")

    if not html_path:
        logger.warning(f"[dynamic-email] Missing html_path for {name}, skipping.")
        return False
    html_path_obj = Path(html_path)
    if html_path_obj.is_absolute() or len(html_path_obj.parts) != 1:
        logger.error(f"[dynamic-email] html_path must be filename only (no directories, no absolute paths) for {name}; got '{html_path}'")
        return False

    resolved = find_email_template(html_path)
    if not resolved:
        logger.error(f"[dynamic-email] Template file not found for {name}: {html_path}")
        return False
    logger.info(f"[dynamic-email] Using template for {name}: {resolved}")
    html_source = Path(resolved).read_text(encoding="utf-8")

    fields: Dict[str, Any] = dict(BASE_FIELDS)

    for arg in spec.get("args", []):
        arg_name = arg.get("name")
        if not arg_name or arg_name in fields:
            continue

        enum_values = arg.get("enum")
        if enum_values and isinstance(enum_values, list) and enum_values:
            if all(isinstance(v, int) for v in enum_values):
                lit_values = tuple(int(v) for v in enum_values)
            elif all(isinstance(v, (int, float)) for v in enum_values):
                lit_values = tuple(float(v) for v in enum_values)
            else:
                lit_values = tuple(str(v) for v in enum_values)
            py_type = Literal[lit_values]  # type: ignore[index]
            required = bool(arg.get("required", True))
            default = arg.get("default", (Ellipsis if required else None))
            if default is not Ellipsis and default is not None and default not in lit_values:
                logger.warning(f"[dynamic-email] Default '{default}' not in enum for {arg_name}; ignoring default.")
                default = Ellipsis if required else None
            desc = arg.get("description") or f"One of: {', '.join(map(str, lit_values))}"
            fields[arg_name] = (py_type, Field(default, description=desc))
            continue

        py_type = TYPE_MAP.get(str(arg.get("type", "string")).lower(), str)
        required = bool(arg.get("required", True))
        field_type = py_type if required else Optional[py_type]  # type: ignore[index]
        default = arg["default"] if "default" in arg else (Ellipsis if required else None)
        desc = arg.get("description")
        fields[arg_name] = (field_type, Field(default, description=desc) if desc is not None else default)

    model = create_model(f"{name}_Args", **fields)  # type: ignore
    # See dynamic_docx_tools: the tool annotation is resolved by name against this
    # module's globals when FastMCP builds the schema, so the model must live here.
    globals()[model.__name__] = model

    renderer = pystache.Renderer(file_encoding="utf-8")

    # Tool body is synchronous and blocking (mustache rendering,
    # MIME construction, synchronous upload). Wrap it in an
    # `async def` that dispatches through `run_blocking()` so
    # behaviour follows the RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED
    # flag uniformly with the rest of the tools.
    def make_tool_fn(_model=model, _html=html_source, _renderer=renderer, _name=name):
        def _sync_impl(data):
            try:
                payload = data.model_dump()
                safe_payload = {k: ("" if v is None else v) for k, v in payload.items()}

                if "promo_code" in safe_payload and "promo_code_block" not in safe_payload:
                    promo_val = safe_payload.get("promo_code")
                    safe_payload["promo_code_block"] = (
                        f"<div class=\"promo\">Use promo code <strong>{promo_val}</strong>.</div>" if promo_val else ""
                    )
                try:
                    html_rendered = _renderer.render(_html, safe_payload)
                except Exception as e:  # pragma: no cover
                    logger.error(f"[dynamic-email] Error rendering template {_name}: {e}")
                    raise ToolError(f"Error rendering template {_name}: {e}")

                # Mirror static create_eml: single HTML body base64 encoded.
                msg = MIMEText(html_rendered, 'html', 'utf-8')
                encoders.encode_base64(msg)  # sets proper Content-Transfer-Encoding and encodes payload

                subject = str(safe_payload.get("subject", ""))
                if subject:
                    msg['Subject'] = subject
                for hdr in ("To", "Cc", "Bcc"):
                    key = hdr.lower()
                    val = safe_payload.get(key)
                    if isinstance(val, list) and val:
                        msg[hdr] = ", ".join(val)
                    elif isinstance(val, str) and val:
                        msg[hdr] = val
                msg['X-Unsent'] = '1'

                buffer = io.BytesIO()
                try:
                    buffer.write(msg.as_bytes())
                    buffer.seek(0)
                    result = upload_file(buffer, "eml", filename=safe_payload.get("file_name") or safe_payload.get("subject") or _name)
                    metrics.record_call("email", _name)
                    return result
                except Exception as e:  # pragma: no cover
                    logger.error(f"[dynamic-email] Error creating email draft for template '{_name}': {e}")
                    raise ToolError(f"Error creating email draft for template '{_name}': {e}")
                finally:
                    buffer.close()

            except ToolError as e:
                metrics.record_error("email", _name, str(e))
                raise
            except Exception as e:
                metrics.record_error("email", _name, str(e))
                logger.error(f"[dynamic-email] Unexpected error in tool '{_name}': {e}", exc_info=True)
                raise ToolError(f"Error generating email from template '{_name}': {e}")

        async def tool_impl(data):
            return await run_blocking(_sync_impl, data)

        tool_impl.__annotations__['data'] = _model  # type: ignore[index]
        tool_impl.__annotations__['return'] = str  # type: ignore[index]
        return tool_impl

    mcp.tool(name=name, description=description, annotations=annotations, meta=meta)(make_tool_fn())
    with _REG_LOCK:
        _REGISTERED_EMAIL[name] = spec
    logger.info(f"[dynamic-email] Registered tool: {name}")
    return True
