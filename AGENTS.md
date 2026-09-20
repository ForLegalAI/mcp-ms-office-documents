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
warning_channel.py  DocumentWarning + WarningChannel: what a build worked around
image_utils.py      image download/decode with the SSRF guard
metrics.py          in-process counters for the admin Status page
docx_tools/ xlsx_tools/ pptx_tools/ email_tools/ xml_tools/   one package per type
admin/              optional FastHTML admin UI (ADMIN_ENABLED); app.py holds
                    routes only — components.py (markup + theme), kinds.py
                    (dynamic kinds), base_templates.py (static base
                    templates), forms.py, views/
docs/               user reference (docs/*.md) and development docs (docs/development/)
```

Request path: handler in `main.py` → `await run_blocking(_<type>_buffer, …)`
→ `extract_user_context_from_request()` → `upload_and_format_response()` →
backend → URL string or LibreChat artifact dict. Details:
[architecture.md](docs/development/architecture.md).

## Rules

**Structure**
- Every tool package exposes a private `_…_buffer()` that returns `BytesIO`
  — or `(BytesIO, warnings)` where the tool has a warnings channel (Word,
  Excel, PowerPoint) — and a public wrapper that also uploads. `main.py` calls
  the buffer function only. New tool: follow
  [adding-a-tool.md](docs/development/adding-a-tool.md).
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

**Warnings**
- A branch that skips, substitutes or degrades something the caller asked for
  reports it: `warnings.add(code, message, **location)` beside the `logger`
  call, never the log alone. The caller gets a success response and never
  reads the server log.
- The channel is a `WarningChannel` from `warning_channel.py`, created in the
  `_…_buffer()` function and threaded down as an **argument**. Never module
  state: builds run concurrently on worker threads.
- Every code lives in the package's `warnings.py` with a severity in its
  `WARNING_SEVERITY` (`error` = not in the file, `warning` = there but not as
  asked, `info` = a substitution the caller will not mind). Each package's
  `test_every_code_has_a_severity` enforces it.
- A tool with a channel returns `(BytesIO, warnings)` from its buffer
  function; `main.py` attaches them with `_with_warnings()`, passing the
  tool's `kind` and `name` — that call also records them for the admin Status
  page, and the arguments are required so a new tool cannot go uncounted.
  PowerPoint keeps its own `SlideWarning` record and shares the severities.

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
- Write a caller's text with `inline_formatting.write_text()`, never
  `paragraph.text` — the tool promises inline markdown and links in every
  text field, and a direct assignment prints the markers instead.
- Every warning takes a code from `pptx_tools/warnings.py`, and every code
  takes a severity in `WARNING_SEVERITY`.

**Admin UI**
- Build pages from `admin/components.py`, never hand-written FastHTML trees.
  A labelled control goes through `field()` — it is what mints the `id` and
  points the `<label>` at it; a table goes through `data_table()`, which is
  what keeps a wide table from scrolling the whole page.
- Per-kind wording and flags live in `admin/kinds.py`, storage metadata in
  `admin/store.py`. Add a kind by editing those two tables, not by adding a
  branch to a view.
- `admin/kinds.py`'s `STYLE_KEYS` must match the keys `docx_tools/style_map.py`
  recognises; `tests/test_admin_style_keys.py` enforces it. Style-name labels
  come from `DEFAULT_STYLE_MAP`, never a second copy.
- No external assets, ever: no CDN stylesheet, no `<script src>`, no web font.
  The theme and scripts are inlined by `components.head_tags()` and
  `tests/test_admin_assets.py` enforces it on every rendered page.
- Colours are CSS custom properties on `:root`, redefined under
  `prefers-color-scheme: dark`. A new rule takes a token, never a literal, or
  it will be wrong in one of the two themes.

**Dynamic tools and schemas**
- Never `Optional[...]` on a dynamic-tool argument; optionality is the
  default alone. Descriptions must be siblings of a flat type.
- Do not give `add_unique_prefix` a default in a dynamic tool body; it must
  reach `upload_file()` as `None`.
- A disabled template (`enabled: false`) is filtered in
  `template_registry.gather_specs()` — the one merge point every registration
  path uses — never at a call site. A new consumer of `gather_specs()` gets
  the filtering for free; only the admin UI passes `include_disabled=True`.
  Absent means enabled, so specs written before #165 stay live.

**Docs — part of every change, never a follow-up**
- Before you finish any change, find every page under `docs/` that describes
  what you touched and update it in the same commit. Concretely:
  - **Internal change** (new, renamed, moved or removed module or function; a
    pipeline step added or reordered; a new invariant, gotcha or extension
    point): the package's page under `docs/development/tools/` or the
    relevant `docs/development/*.md` — its pipeline diagram, module map,
    "how it works", extension points and invariants. Change this file too if
    a rule here changes.
  - **User-visible change** (input format, parameter, output shape,
    behaviour): the tool description in `main.py`, the reference page under
    `docs/`, and the tests, together.
  - **New environment variable**: `config.py`, `.env.example`,
    `docs/configuration.md`; `tests/test_config_docs.py` enforces this.
  - **A limitation you fixed**: remove it from the page's "Known limitations"
    and close or update the issue it links to.
  - **A new package or backend**: a new page in the same shape as the others,
    linked from `docs/README.md`.
- A change with no doc update is only correct if you checked and nothing
  described what you touched. Say so in the commit message.
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
