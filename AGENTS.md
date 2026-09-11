# AGENTS.md

Operational brief for coding agents and humans alike. Rules here, reasons in
[`docs/development/`](docs/development/architecture.md).

## What this is

An MCP server (FastMCP 3, Python 3.12) that turns Markdown or structured
input into Office files. Runs in Docker on port 8958, endpoint `/mcp`,
streamable-HTTP. Entry point `main.py`. Six static tools (Word, Excel,
PowerPoint, template listing, email, XML) plus dynamic tools registered from
YAML templates.

## Commands

```bash
pip install -r requirements-dev.txt
ruff check .                     # lint (CI)
pytest -m "not network"          # tests (CI); plain `pytest` includes network tests
pytest tests/test_docx_base.py   # one module
python main.py                   # run locally, no .env needed
```

## Layout

```
main.py             tool handlers, startup order, health routes, server launch
config.py           the only module that reads os.environ; get_config() singleton
async_runner.py     run_blocking(): bounded thread pool for every blocking call
librechat_integration.py   request-header user context; upload_and_format_response()
upload_tools/       upload_file() / upload_file_async() dispatch; backends/<strategy>.py
template_utils.py   template file resolution (custom → default, container → local)
template_registry.py  YAML spec merging (master + *.d/) and live tool removal
inline_markdown.py  the inline-emphasis grammar shared by Word and PowerPoint
image_utils.py      image download/decode with the SSRF guard
metrics.py          in-process counters for the admin Status page
docx_tools/ xlsx_tools/ pptx_tools/ email_tools/ xml_tools/   one package per type
admin/              optional FastHTML admin UI (ADMIN_ENABLED)
docs/               user reference (docs/*.md) and development docs (docs/development/)
```

Request path: handler in `main.py` → `await run_blocking(_<type>_buffer, …)`
→ `extract_user_context_from_request()` → `upload_and_format_response()` →
backend → URL string or LibreChat artifact dict. Details:
[architecture.md](docs/development/architecture.md).

## Rules

**Structure**
- Every tool package exposes a private `_…_buffer()` that returns `BytesIO`
  and a public wrapper that also uploads. `main.py` calls the buffer function
  only. New tool: follow [adding-a-tool.md](docs/development/adding-a-tool.md).
- Blocking work goes through `await run_blocking(...)`, always. Never call
  a buffer function or `upload_file()` directly from an `async def` handler.
- Read request headers on the event loop, before dispatch; never inside a
  buffer function.
- Dynamic template tools call `upload_file()` inside their own offloaded
  body and do **not** go through `upload_and_format_response()`.
- Config: read via `get_config()`; never `os.environ` outside `config.py`.
  A new variable goes into `config.py`, `.env.example` and
  `docs/configuration.md`; `tests/test_config_docs.py` enforces it.
- Templates: resolve through `template_utils`; never hard-code a path.
- Images: `image_utils.load_image()`; never `requests` directly.
- Logging: `logging.getLogger(__name__)`. Level comes from `DEBUG` only.

**Errors**
- Raise `fastmcp.exceptions.ToolError` only at the handler boundary in
  `main.py` or a dynamic tool body. `RuntimeError` in the upload layer.
  `ValueError` for input the caller can fix.
- A backend returns a string on success, `None` or raises on failure. Never
  return an error message as a string.
- PowerPoint reports anything it worked around on
  `PowerpointPresentation.warnings`; anything else that grows a warnings
  channel returns `(BytesIO, warnings)` the same way.

**Word**
- One line-break model: every newline reaching the inline layer is a soft
  break; two-space runs are assembled by `_soft_break_run()` in the block
  dispatcher and stop before any block. Never give `expand_br_to_block_breaks()`
  the renderer's numbering state (#110). Details in
  [word.md](docs/development/tools/word.md#one-line-break-model).

**PowerPoint**
- Never index `slide_layouts[N]` in a builder; go through `_new_slide()`.
- Keep the published slide schema flat (no `oneOf`/`$ref`/`discriminator`);
  validation is `coerce_slides()` in the build step.

**Dynamic tools and schemas**
- Never `Optional[...]` on a dynamic-tool argument; optionality is the
  default alone. Descriptions must be siblings of a flat type.
- Do not give `add_unique_prefix` a default in a dynamic tool body; it must
  reach `upload_file()` as `None`.

**Docs**
- A user-visible change touches the tool description in `main.py`, the
  reference page in `docs/`, and the tests together.
- Each package's module docstring names what it owns and which entry point
  `main.py` uses. Keep it true.

## Tests

- One file per behaviour under `tests/`. Regression tests name the issue.
- Build without uploading: patch `upload_file` where it is looked up
  (`<pkg>.base_<type>_tool.upload_file`), or call the buffer function.
  PowerPoint tests instantiate `PowerpointPresentation` and call `.save()`.
- Output for manual inspection goes to `tests/output/{docx,pptx,xlsx}/`.
- Config is a singleton read at import; see
  [testing.md](docs/development/testing.md) for the reload pattern.

## Where to read more

[architecture](docs/development/architecture.md) ·
[word](docs/development/tools/word.md) · [excel](docs/development/tools/excel.md) ·
[powerpoint](docs/development/tools/powerpoint.md) · [email](docs/development/tools/email.md) ·
[xml](docs/development/tools/xml.md) · [dynamic templates](docs/development/dynamic-templates.md) ·
[upload backends](docs/development/upload-backends.md) · [shared modules](docs/development/shared-modules.md) ·
[testing](docs/development/testing.md) · [CONTRIBUTING](CONTRIBUTING.md)
