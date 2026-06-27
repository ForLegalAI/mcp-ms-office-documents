# AGENTS.md

## Project Overview

MCP (Model Context Protocol) server built with **FastMCP 3.0** that exposes Office document generation as MCP tools. Runs as a Docker container (Python 3.12, Alpine) on port **8958** at `/mcp` using streamable-HTTP transport. Entry point: `main.py`.

## Architecture

```
main.py                  ← Registers all MCP tools on a single FastMCP instance
├── {docx,xlsx,pptx,email,xml}_tools/
│   ├── __init__.py      ← Re-exports the public function (e.g. markdown_to_word)
│   ├── base_*_tool.py   ← Core conversion logic (markdown → document bytes)
│   ├── helpers.py        ← Parsing, formatting, shared utilities
│   └── dynamic_*_tools.py  ← YAML-driven tool registration (docx, email only)
├── upload_tools/
│   ├── main.py          ← upload_file() dispatches to strategy backend
│   └── backends/{local,s3,gcs,azure,minio}.py
├── config.py            ← Singleton Config from env vars (Pydantic v2), logging setup
├── template_utils.py    ← Template resolution: custom_templates/ → default_templates/
└── middleware.py         ← Optional API key auth (Bearer / x-api-key header)
```

**Data flow:** Every tool converts input → in-memory bytes → calls `upload_file(file_obj, suffix)` → backend saves/uploads → returns URL or path string to the MCP client.

## Key Conventions

- **Config is centralized in `config.py`** — no module reads `os.environ` directly. Access via `get_config()` singleton.
- **Template resolution** (`template_utils.py`): searches `custom_templates/` before `default_templates/`, with `/app/*` container paths tried first, then local paths. Never hardcode template paths.
- **Dynamic tool registration**: YAML files in `config/` define parameterized email/docx templates. Each entry becomes a separate MCP tool at startup via `register_*_template_tools_from_yaml(mcp, path)`. Placeholders use Mustache syntax `{{name}}`. See `config/docx_templates.yaml` for the canonical example.
  - **Merged loading**: the master `config/<kind>_templates.yaml` is merged with UI-managed per-template files in `config/<kind>_templates.d/<name>.yaml` (the `.d` file wins on a name clash). The documented master files are never rewritten by tooling. Merge logic lives in `template_registry.py`.
  - **Live (un)registration**: `register_docx_template` / `register_email_template` (and their `unregister_*` counterparts) can add/replace/remove tools on a running `FastMCP` instance via `local_provider`, so the admin UI applies template edits without a restart. A module-level registry tracks live dynamic tool names.
- **Admin UI** (`admin/`, opt-in via `ADMIN_ENABLED`): a FastHTML template-admin mounted in the same ASGI process. `admin/store.py` defines `TemplateStore` (interface) + `FileTemplateStore`, which the UI uses to persist managed templates; the storage layer is abstracted so a shared backend can replace it for multi-replica deployments.
- **Pydantic models** for tool arguments are created dynamically with `create_model()` in `dynamic_*_tools.py`. The `TYPE_MAP` dict maps YAML type strings to Python types.
- **Error handling in tools**: raise `fastmcp.exceptions.ToolError` for user-facing errors; use `RuntimeError` in upload/backend layers.
- **Logging**: use `logging.getLogger(__name__)` everywhere. Level controlled by `DEBUG` env var only.

## Adding a New Document Tool

1. Create `<type>_tools/` package with `__init__.py`, `base_<type>_tool.py`, and optional `helpers.py`.
2. The base tool function should: accept content → produce an `io.BytesIO` → call `upload_file(buffer, "<ext>")` → return the result string.
3. Register the async wrapper in `main.py` using `@mcp.tool(name=..., description=..., tags=..., annotations=...)`.
4. Use `Annotated[<type>, Field(description=...)]` for all tool parameters — the descriptions are critical because MCP clients (AI models) rely on them.

## Tests

```bash
pytest                        # Run all tests (asyncio_mode=auto in pytest.ini)
pytest tests/test_docx_base.py  # Single module
```

- Tests live in `tests/` and output generated files to `tests/output/{docx,pptx,xlsx}/` for manual inspection.
- Upload is mocked in tests — patch `upload_file` or the specific `*_tool.upload_file` to capture bytes without needing a real backend. See `test_xlsx_creation.py::_create_workbook_from_markdown` for the pattern.
- PPTX tests instantiate `PowerpointPresentation` directly and call `.save()` to get a buffer, bypassing upload entirely.
- No `.env` required for tests — `config.py` defaults to `LOCAL` strategy and INFO logging.
