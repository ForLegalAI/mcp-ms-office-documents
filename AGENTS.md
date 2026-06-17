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
│   ├── formula_engine.py ← Pure-Python Excel formula evaluation (xlsx only)
│   ├── xml_cache.py      ← Injects cached <v> values into xlsx XML (xlsx only)
│   └── dynamic_*_tools.py  ← YAML-driven tool registration (docx, email only)
├── upload_tools/
│   ├── main.py          ← upload_file() dispatches to strategy backend
│   └── backends/{local,s3,gcs,azure,minio}.py
├── config.py            ← Singleton Config from env vars (Pydantic v2), logging setup
├── template_utils.py    ← Template resolution: custom_templates/ → default_templates/
└── middleware.py         ← Optional API key auth (Bearer / x-api-key header)
```

**Data flow:** Every tool converts input → in-memory bytes → calls `upload_file(file_obj, suffix)` → backend saves/uploads → returns URL or path string to the MCP client.

The **XLSX tool** has an extra recalculation pass between workbook save and upload: after `openpyxl` writes the file (with empty `<v>` tags for formula cells), `xlsx_tools/formula_engine.py` evaluates every formula in-process via the pure-Python `formulas` library, and `xlsx_tools/xml_cache.py` walks the worksheet XML to replace the empty `<v>` tags with the computed values. This makes the file preview correctly in tools that don't recalculate on open (Google Sheets preview, mail clients, openpyxl loaded with `data_only=True`). Controlled by `XLSX_RECALC_ENABLED` / the `recalc` tool parameter.

## Key Conventions

- **Config is centralized in `config.py`** — no module reads `os.environ` directly. Access via `get_config()` singleton.
- **Template resolution** (`template_utils.py`): searches `custom_templates/` before `default_templates/`, with `/app/*` container paths tried first, then local paths. Never hardcode template paths.
- **Dynamic tool registration**: YAML files in `config/` define parameterized email/docx templates. Each YAML entry becomes a separate MCP tool at startup via `register_*_template_tools_from_yaml(mcp, path)`. Placeholders use Mustache syntax `{{name}}`. See `config/docx_templates.yaml` for the canonical example.
- **Pydantic models** for tool arguments are created dynamically with `create_model()` in `dynamic_*_tools.py`. The `TYPE_MAP` dict maps YAML type strings to Python types.
- **Error handling in tools**: raise `fastmcp.exceptions.ToolError` for user-facing errors; use `RuntimeError` in upload/backend layers.
- **Logging**: use `logging.getLogger(__name__)` everywhere. Level controlled by `DEBUG` env var only.

## XLSX-Specific Conventions

- **Formula recalculation** (`xlsx_tools/formula_engine.py` + `xlsx_tools/xml_cache.py`): on by default (`XLSX_RECALC_ENABLED=true`). The `formulas` library is pure-Python — no LibreOffice or external binary is required. Cached values are injected directly into the worksheet XML by replacing empty `<v>` tags inside formula `<c>` cells. Recalculation runs in a worker thread bounded by `XLSX_RECALC_TIMEOUT_SECONDS` (default 30s); on timeout the file is delivered without cached values.
- **Formula error policy**: errors (`#REF!`, `#DIV/0!`, etc.) and circular references (`#CIRC!`) are always detected. When `recalc=True` is **explicitly** passed, the call fails with a descriptive error so the model can fix the formulas and retry (zero-errors standard). When recalc runs as a **default** (parameter not passed), errors are logged but the file is delivered — misconfiguration must never block document generation. The distinction is tracked via `recalc_explicitly_requested` in `base_xlsx_tool.py`. Errors are grouped by type with counts and locations in the message (`_format_grouped_errors()` in `base_xlsx_tool.py`).
- **Circular-reference detection** (`detect_circular_references()` in `formula_engine.py`): the `formulas` library silently resolves cycles to nothing (no cached values, no errors), so an independent DFS-based graph analysis runs on the formula strings themselves. Every cell on a cycle is reported as a `#CIRC!` error (the `CIRCULAR_ERROR_TYPE` constant — deliberately distinct from the seven OOXML sentinels). Runs even when the recalc engine is unavailable. Reference parsing in `extract_formula_references()` handles cross-sheet refs, ranges (expanded via `_expand_range()`), and absolute (`$A$1`) notation.
- **String-result formula limitation**: formulas returning a string value are not injected as cached values (requires fragile shared-strings-table edits post-write). Numeric/boolean/datetime results are cached. The cell still computes correctly when opened in a recalc-on-open client; only static previewers miss it. Documented in `Readme.md`.
- **RecalcResult** (`formula_engine.py`) carries `values_map`, `errors`, `total_formulas`, `recalc_performed`, `skip_reason`. `total_formulas` is populated from `count_formulas()` in `xml_cache.py` (counts `<f>` elements).
- **Financial modeling** (the `financial_modeling` tool parameter): when true, applies CFA-standard color coding via `apply_financial_styling()` in `helpers.py` — blue for hardcoded inputs, black for local formulas, green for cross-sheet references, **red for external-workbook links** (formulas with `[Workbook.xlsx]` syntax), yellow background for source-cited cells. Also treats 4-digit-year strings in data rows as text labels. Combine with the `sources:` directive and number-format variants (`number:dash`, `currency:$:parens`, `multiple`).
- **Number-format variants** (`helpers.py`): `NUMBER_FORMAT_VARIANTS`, `PERCENT_FORMAT_VARIANTS`, and `MULTIPLES_FORMAT_VARIANTS` dicts map variant keywords (`dash`, `parens`) to Excel format strings. The `multiple` type renders valuation multiples as `12.5x` (value stored raw, `x` is display only).
- **Markdown directives** are parsed by `xlsx_tools/parser.py` as `<!-- key: value -->` lines above a table. Currently supported: `freeze`, `types`, `sources`. The directive parser is generic — any new directive key is automatically captured and forwarded to `add_table_to_sheet`'s `directives` dict.
- **Default font** (`XLSX_DEFAULT_FONT` env var / `default_font` parameter): optional font family applied to every cell. Inline `code` formatting (Courier New) always overrides it.
- **Docker base image** is `python:3.12-slim` (not alpine) because `formulas` pulls in numpy/scipy, which only have pre-built wheels for glibc. On slim the wheels install directly with no build tools; on alpine they'd need ~400MB of build deps.

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
