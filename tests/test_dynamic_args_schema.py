"""Dynamic tool args must expose a FLAT JSON schema with descriptions intact.

Optional args (required: false) used to be typed Optional[...], which pydantic
renders as ``anyOf: [{type}, {"type": "null"}]`` with the description as a
SIBLING of anyOf. Several MCP clients normalize anyOf schemas and silently drop
that sibling description, so every optional argument reached the model with no
description at all. Optionality must instead come from the default alone: the
field keeps its plain type (description survives everywhere) and is simply
absent from the schema's "required" list.
"""
import asyncio
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from fastmcp import FastMCP  # noqa: E402

from docx_tools.dynamic_docx_tools import register_docx_template  # noqa: E402
from email_tools.dynamic_email_tools import register_email_template  # noqa: E402

DOCX_SPEC = {
    "name": "schema_probe_docx",
    "description": "schema probe",
    "docx_path": "default_docx_template.docx",
    "args": [
        {"name": "req_arg", "type": "string", "required": True,
         "description": "Required arg"},
        {"name": "opt_default", "type": "string", "required": False,
         "default": " ", "description": "Optional with default"},
        {"name": "opt_nodefault", "type": "string", "required": False,
         "description": "Optional without default"},
        {"name": "opt_bool", "type": "bool", "required": False,
         "default": True, "description": "Optional bool"},
    ],
}

EMAIL_SPEC = {
    "name": "schema_probe_email",
    "description": "schema probe",
    "html_path": "default_email_template.html",
    "args": [
        {"name": "req_arg", "type": "string", "required": True,
         "description": "Required arg"},
        {"name": "opt_default", "type": "string", "required": False,
         "default": " ", "description": "Optional with default"},
        {"name": "opt_nodefault", "type": "string", "required": False,
         "description": "Optional without default"},
    ],
}


def _model_schema(register, spec, probe_field):
    mcp = FastMCP("schema-test")
    assert register(mcp, spec), "registration failed"
    tool = asyncio.run(mcp.get_tool(spec["name"]))
    schema = tool.parameters
    for definition in schema.get("$defs", {}).values():
        if probe_field in definition.get("properties", {}):
            return definition
    assert probe_field in schema.get("properties", {}), "arg model not found in schema"
    return schema


def _assert_flat_with_description(props, names):
    for name in names:
        prop = props[name]
        assert "anyOf" not in prop, f"{name}: anyOf leaks the description in some clients"
        assert "type" in prop, f"{name}: flat type expected"
        assert prop.get("description"), f"{name}: description missing from schema"


def test_docx_optional_args_keep_flat_schema_and_description():
    schema = _model_schema(register_docx_template, DOCX_SPEC, "req_arg")
    _assert_flat_with_description(
        schema["properties"], ["req_arg", "opt_default", "opt_nodefault", "opt_bool"])
    # Optionality is preserved: only the truly required arg is required.
    assert schema.get("required") == ["req_arg"]
    assert schema["properties"]["opt_default"]["default"] == " "
    assert schema["properties"]["opt_bool"]["type"] == "boolean"


def test_email_optional_args_keep_flat_schema_and_description():
    schema = _model_schema(register_email_template, EMAIL_SPEC, "req_arg")
    _assert_flat_with_description(
        schema["properties"], ["req_arg", "opt_default", "opt_nodefault"])
    # Base fields stay flat too (to/cc/bcc used to be Optional[list[str]]).
    _assert_flat_with_description(schema["properties"], ["subject", "to", "cc", "bcc"])
    required = schema.get("required", [])
    assert "req_arg" in required and "subject" in required
    assert "opt_default" not in required and "to" not in required


def test_explicit_null_for_optional_arg_is_rejected():
    """Documented tradeoff of the flat schema: optional args are OMITTED, not nulled.

    With the plain (non-Optional) type, explicitly passing null no longer
    validates — this test pins the behavior change down so an accidental
    revert to Optional[...] (and thus to anyOf schemas) fails loudly.
    """
    import pydantic
    import pytest

    import docx_tools.dynamic_docx_tools as ddt

    _model_schema(register_docx_template, DOCX_SPEC, "req_arg")
    model = getattr(ddt, f"{DOCX_SPEC['name']}_DocxArgs")
    # Omitting the optional arg uses its default...
    assert model(req_arg="x").opt_default == " "
    # ...but explicit null is rejected instead of silently accepted.
    with pytest.raises(pydantic.ValidationError):
        model(req_arg="x", opt_default=None)


def test_an_optional_arg_without_default_publishes_no_null_default():
    """``"default": null`` told the model null was valid while validation refuses
    it, and cost a key per optional argument; it is not published. A real
    default still is, and omitting the argument still yields None."""
    import asyncio
    import json

    from fastmcp import Client

    docx = _model_schema(register_docx_template, DOCX_SPEC, "req_arg")["properties"]
    assert "default" not in docx["opt_nodefault"]
    assert docx["opt_default"]["default"] == " " and docx["opt_bool"]["default"] is True

    email = _model_schema(register_email_template, EMAIL_SPEC, "req_arg")["properties"]
    for name in ("opt_nodefault", "to", "cc", "bcc"):
        assert "default" not in email[name], name

    import docx_tools.dynamic_docx_tools as ddt
    model = getattr(ddt, f"{DOCX_SPEC['name']}_DocxArgs")
    assert model(req_arg="x").opt_nodefault is None

    mcp = FastMCP("wire")
    assert register_docx_template(mcp, DOCX_SPEC)

    async def listed():
        async with Client(mcp) as c:
            return (await c.list_tools())[0].inputSchema
    assert '"default": null' not in json.dumps(asyncio.run(listed()))
