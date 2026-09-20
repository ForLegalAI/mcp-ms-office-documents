# Architecture

How a tool call travels through this server, which module owns each step, and
the rules that keep the process responsive. Read this before the per-tool pages
under [`tools/`](tools/) — they assume it.

For **what** the tools accept (Markdown grammar, slide schema, YAML options), see
the user reference in [`docs/`](../). This page and its siblings explain **how**
the code does it.

## Process layout

One Python process runs everything: a FastMCP 3 server on uvicorn/Starlette,
listening on port **8958**. The MCP endpoint is `/mcp` over the streamable-HTTP
transport. Entry point: `main.py`.

`main.py` has two ways of starting, chosen by `ADMIN_ENABLED`:

| Mode | What runs | Where |
|------|-----------|-------|
| Admin UI off (default) | `mcp.run(transport="streamable-http", path="/mcp")` | `main.py`, `__main__` block |
| Admin UI on | One Starlette app: the admin UI mounted under `ADMIN_PATH` (default `/admin`) and the MCP app mounted at `/`, sharing the MCP app's lifespan | `admin.app.build_combined_app()` |

The admin mount is listed first so its prefix wins over the catch-all MCP
mount. Running both in one process is what lets the admin UI register a
template's MCP tool live, without a restart.

Three plain HTTP routes — `/healthz`, `/readyz`, `/livez` — are registered with
FastMCP's `@custom_route` at the Starlette layer. They sit **outside** the MCP
middleware stack, so they answer without an API key. Kubernetes probes rely on
this; do not move them behind the middleware.

## Startup sequence

Module import order in `main.py` is the startup sequence. Each step logs what it
decided, so a deployment's log shows the full picture.

1. **Configuration.** `get_config()` builds the `Config` singleton from
   environment variables (`config.py`) and configures logging once. A missing
   or contradictory setting raises `ValueError` here and the process does not
   start. Every other module reads settings through `get_config()`; nothing
   reads `os.environ` directly.
2. **Authentication.** `ApiKeyAuthMiddleware` is added only when `API_KEY` is
   set (`middleware.py`).
3. **Health routes** are registered.
4. **Config directory.** `/app/config` if it exists (the container mount),
   otherwise `./config` next to `main.py` (local development).
5. **Dynamic tools.** Email templates, then Word templates, are registered from
   `config/<kind>_templates.yaml` merged with `config/<kind>_templates.d/`.
   Registration runs whenever either source exists. Each is wrapped in its own
   `try`; a bad YAML file is logged and skipped, never fatal.
6. **PowerPoint template validation.** Every registered `.pptx`/`.potx` is
   opened once and its layout coverage logged. A bad template degrades to the
   built-in theme at build time; it never stops startup.
7. **Static tools.** The six tools declared with `@mcp.tool` in `main.py` are
   registered as the module body executes.

Two modules read configuration at import time rather than per call:
`upload_tools/main.py` caches the storage strategy, and `config.py` itself
configures logging. `async_runner.py` deliberately reads its flag on every
call, so tests can flip it without reloading the module.

## Request lifecycle

Every static document tool follows the same path. The Word, Excel and
PowerPoint tools carry a warnings channel back out alongside the file; the
template-listing tool skips the upload.

```
MCP client
  │  POST /mcp  (streamable-HTTP)
  ▼
Starlette / FastMCP
  │  ApiKeyAuthMiddleware.on_request          middleware.py     (only if API_KEY set)
  ▼
Tool handler                                  main.py
  │  Pydantic validates the Annotated params
  │  (slides arrive as a loose list; see below)
  ▼
await run_blocking(_<tool>_buffer, ...)       async_runner.py
  │  runs on a bounded worker thread; returns io.BytesIO,
  │  or (BytesIO, warnings) for Word, Excel and PowerPoint
  │  (PowerPoint: coerce_slides() validates here, pptx_tools/schema.py)
  ▼
extract_user_context_from_request()           librechat_integration.py
  │  reads X-User-Id / X-User-Email / X-Conversation-Id
  ▼
upload_and_format_response(buffer, suffix, file_name, ctx, message, add_unique_prefix)
  │                                           librechat_integration.py
  ├─ LIBRECHAT ──► upload_file_async()        upload_tools/main.py
  │                  └► upload_to_librechat()  upload_tools/backends/librechat.py
  │                  └► format_file_artifact() → dict
  └─ otherwise ─► run_blocking(upload_file)   upload_tools/main.py
                     └► upload_to_<backend>()  upload_tools/backends/*.py → message with URL or path
  ▼
return str | dict                             main.py
  (_with_warnings(): wraps in {"file", …, "warnings"} when there are warnings —
   PowerPoint also carries "slide_count". Each warning is a
   {code, severity, message} record with its location (slide, line, or sheet
   and cell), not a sentence. Nothing to report → the shape is unchanged)
```

Stage by stage:

| Stage | Module | Function | Notes |
|-------|--------|----------|-------|
| Authenticate | `middleware.py` | `ApiKeyAuthMiddleware.on_request` | Bearer, plain token, or `x-api-key`; constant-time compare; throttled warning log |
| Validate input | `main.py` | Pydantic `Annotated[..., Field(...)]` | Field descriptions are what the calling model reads. Slides are deliberately typed loosely here (`SlidesInput` publishes a flat schema and accepts any list) so that clients which mangle `oneOf` schemas still reach the server |
| Build document | `<type>_tools/base_<type>_tool.py` | `_markdown_to_word_buffer`, `_markdown_to_excel_buffer`, `_create_presentation_buffer`, `_create_eml_buffer`, `_create_xml_buffer` | Input to bytes. No upload, no request context. May fetch images over the network. Word, Excel and PowerPoint return `(BytesIO, warnings)`; see [`shared-modules.md`](shared-modules.md#warning_channelpy) |
| Report workarounds | `main.py` | `_with_warnings` | Widens a bare URL string into `{"file", …, "warnings"}`, or adds `warnings` to the LibreChat artifact dict. A clean build returns exactly what it always did. Also records the call and its warnings with `metrics`, since this is the one point every channel-carrying tool passes through on success — hence the required `kind` and `name` |
| Validate slides | `pptx_tools/schema.py`, called from `pptx_tools/slide_builder.py` | `coerce_slides` | Runs inside the build step, on the worker thread. Raises `ValueError` with messages like `slide 2 -> rows.0: …`, which the handler passes through as a `ToolError` |
| Offload | `async_runner.py` | `run_blocking` | See [Threading model](#threading-model) |
| User context | `librechat_integration.py` | `extract_user_context_from_request` | Headers are trusted verbatim; see [Security boundaries](#security-boundaries) |
| Upload | `librechat_integration.py` → `upload_tools/main.py` | `upload_and_format_response` → `upload_file` / `upload_file_async` | Chooses the backend from `UPLOAD_STRATEGY`; resolves the unique-prefix default |
| Name the object | `upload_tools/utils.py` | `generate_named_object_name`, `generate_unique_object_name` | Sanitises the caller's `file_name`; adds an 8-character UUID prefix unless told otherwise |

### Two entry points per tool package

Each document package exposes two functions:

- A **private buffer function** (`_markdown_to_word_buffer` and friends) that
  builds the document and returns `io.BytesIO`. This is what `main.py` calls.
  A tool with a warnings channel returns `(io.BytesIO, warnings)` instead —
  Word, Excel and PowerPoint do.
- A **public wrapper** (`markdown_to_word` and friends) that builds *and*
  uploads synchronously through `upload_file()` and returns the backend's
  string. It exists for direct library use and for tests. It cannot return a
  LibreChat artifact, and it drops the warnings (it has nowhere to put them).

`main.py` uses the buffer functions only, so the upload step is dispatched the
same way for every tool and the LibreChat branch lives in one place. A new tool
must follow the same split; see [`adding-a-tool.md`](adding-a-tool.md).

### Dynamic template tools

Tools registered from YAML (`docx_tools/dynamic_docx_tools.py`,
`email_tools/dynamic_email_tools.py`) take a shorter path. Each registration
builds a Pydantic model with `create_model()` from the template's declared
arguments, then registers an `async def` handler whose body is
`await run_blocking(_sync_impl, data)`. That single synchronous body loads the
template, substitutes placeholders, saves to a buffer, **and calls
`upload_file()` itself**. It does not go through `upload_and_format_response`,
so it is never dispatched twice.

Consequence worth knowing: `upload_file()` is the synchronous entry point and
refuses the `LIBRECHAT` strategy with a `RuntimeError`. Dynamic template tools
therefore do not work under `UPLOAD_STRATEGY=LIBRECHAT` today. The static tools
do. Tracked in [#113](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/113); see [`dynamic-templates.md`](dynamic-templates.md).

## Threading model

FastMCP runs on an asyncio event loop. Every document builder is synchronous
and blocking: it opens zip archives, parses Markdown, downloads images with
`requests`, and uploads with boto3 or its equivalents. Calling one directly from
an `async def` handler freezes the loop for the whole call — no other request
is served, health probes included, and Kubernetes restarts the pod.

`run_blocking()` in `async_runner.py` is the single answer to this:

- **Enabled (the default).** The callable is submitted to one process-wide
  `ThreadPoolExecutor` sized by `RUN_BLOCKING_MAX_WORKERS` (default 4). The
  pool is bounded on purpose: Python's default executor sizes itself from the
  *host's* CPU count, which on a 1-vCPU pod means up to 32 threads all
  contending for the GIL. Extra requests queue in the executor instead.
- **Disabled** (`RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED=false`). The callable
  runs inline on the event loop. Only for local debugging or to rule threading
  out of a regression.

Rules that follow from this:

1. Every tool handler is `async def` and calls blocking work only through
   `await run_blocking(...)`. Call sites never branch on the flag.
2. Static tools dispatch **twice**: once for the build, once for the upload
   (inside `upload_and_format_response`). Dynamic tools dispatch **once** with
   a body that does both. Do not route a dynamic tool through
   `upload_and_format_response`.
3. Nothing inside a buffer function may touch the request context. Request
   headers are read on the event loop by `extract_user_context_from_request()`
   before dispatch.

## Error mapping

| Where | Raised | Reaches the client as |
|-------|--------|-----------------------|
| Tool handler in `main.py` | any `Exception` from build or upload | `fastmcp.exceptions.ToolError` with a message prefixed by the tool ("Error creating Word document: …") |
| PowerPoint slides | `ValueError` from `coerce_slides()`, raised inside the build step | `ToolError` carrying the message verbatim, e.g. `slide 2 -> rows.0: …` |
| Upload dispatcher and backends | `RuntimeError` | wrapped by the handler as above |
| LibreChat without `X-User-Id` | `ValueError` | wrapped by the handler as above |
| Dynamic tool body | any `Exception` | `ToolError("Error generating document from template <name>: …")`, and `metrics.record_error()` is called |
| Auth middleware | `fastmcp.exceptions.AuthorizationError` | MCP error response; the tool is never invoked |
| Configuration at startup | `ValueError("Invalid configuration: …")` | process exits |

Convention: raise `ToolError` only at the handler boundary, `RuntimeError` in
the upload layer, and plain `ValueError` for input that failed validation.
Log with `logging.getLogger(__name__)`; the level is controlled by `DEBUG`.

## Upload strategies and response shape

`UPLOAD_STRATEGY` selects one backend for the whole process.

| Strategy | Backend module | Returns to the client |
|----------|----------------|-----------------------|
| `LOCAL` | `upload_tools/backends/local.py` | a sentence naming the saved path under `./output`, relative to the process working directory |
| `S3` | `upload_tools/backends/s3.py` | a sentence containing a pre-signed URL, valid for `SIGNED_URL_EXPIRES_IN` seconds |
| `GCS` | `upload_tools/backends/gcs.py` | a sentence containing a signed URL |
| `AZURE` | `upload_tools/backends/azure.py` | a sentence containing a SAS URL |
| `MINIO` | `upload_tools/backends/minio.py` | a sentence containing a pre-signed URL |
| `LIBRECHAT` | `upload_tools/backends/librechat.py` | a dict with a `result` object that LibreChat renders as a file attachment |

The traditional backends return prose for the model to relay ("Link to created
document … valid for N seconds"), not a bare URL. A backend signals failure by
raising or by returning `None`, which the dispatcher turns into a
`RuntimeError`. Returning an error *string* is not a failure signal; the
dispatcher would pass it to the client as a success.

Backends import their SDK lazily so an optional backend's dependency is never
loaded for a deployment that does not use it. Keep that rule when adding one;
see [`adding-a-backend.md`](adding-a-backend.md).

The **unique-prefix default** is resolved in the upload layer, not by callers:
when `add_unique_prefix` arrives as `None`, traditional backends get a prefix
(collision safety in shared storage) and LibreChat does not (it adds its own).
Callers pass the parameter through untouched so this rule lives in one place.

## Configuration

`config.py` is the single source of settings. `Config.from_env()` reads the
environment once into nested Pydantic models:

| Section | Holds |
|---------|-------|
| `LoggingSettings` | level derived from `DEBUG` |
| `StorageSettings` | `strategy` plus one optional settings object per backend; a validator checks the selected backend's required fields |
| `AdminSettings` | `enabled`, `path`, `password`; `Config.admin_password_effective` falls back to `API_KEY` |
| flat fields | `api_key`, `run_blocking_by_asyncio_thread_enabled`, `run_blocking_max_workers`, `allow_private_image_addresses`, `stateless_http` |

The full variable list with defaults is in [`../configuration.md`](../configuration.md).

## Templates and directories

Document templates are searched in four directories, custom before default
and, within each, the container mount before the local checkout:

| Order | Directory | Resolver |
|------:|-----------|----------|
| 1 | `/app/custom_templates` | `template_utils.find_file_in_template_dirs()` |
| 2 | `./custom_templates` | |
| 3 | `/app/default_templates` | |
| 4 | `./default_templates` | |

So a custom template in the local checkout beats a shipped default in the
container. The config directory is simpler: `/app/config` if it exists,
otherwise `./config` (`_CONFIG_DIR` in `main.py`).

Never hard-code a template path; go through `template_utils`.

Dynamic template specs are merged by `template_registry.gather_specs()`: the
hand-written master `config/<kind>_templates.yaml` plus one file per template
in `config/<kind>_templates.d/`. A `.d` entry wins on a name clash. Tooling
never rewrites the master file. A spec marked `enabled: false` is dropped
there, so it is never registered by any path; absent means enabled.

PowerPoint templates have their own registry (`pptx_tools/templates.py`) that
maps names to files and layout roles. It is cached against a modification-time
fingerprint of the config and template directories, so a dropped-in template is
picked up without a restart.

## Admin UI and live registration

`admin/` is an optional FastHTML app (`ADMIN_ENABLED=true`). It writes one YAML
file per template into the `.d` directory plus the asset into
`custom_templates/`, then calls `register_docx_template()` /
`register_email_template()` on the running `FastMCP` instance so the tool
appears immediately. `template_registry.safe_remove_tool()` handles
re-registration and deletion across FastMCP versions.

`metrics.py` keeps in-process counters per tool — calls, errors and the
warnings a finished build reported, counted per severity — and a bounded ring
buffer of recent log records, all shown on the admin Status page. It has no
external dependencies and is safe to import from core tool modules.

Live registration assumes a single instance owns the template files. With
several replicas, put the files on shared storage and restart the pods.

## Security boundaries

- **API key.** `ApiKeyAuthMiddleware` gates every MCP request when `API_KEY`
  is set. Comparison is constant-time. Health routes bypass it by design.
- **LibreChat headers.** `X-User-Id` and friends are trusted without
  validation. The assumption is that only LibreChat can reach the endpoint,
  which the API key enforces. Do not expose a LibreChat-strategy server
  without one.
- **Image fetching.** `image_utils.assert_url_is_public()` resolves every
  hostname and refuses private, loopback, link-local and metadata addresses,
  including after redirects. `SSRF_ALLOW_PRIVATE_ADDRESSES=true` disables the
  check for trusted networks.
- **Admin UI.** One shared password (`ADMIN_PASSWORD`, falling back to
  `API_KEY`), CSRF tokens on forms, and a 10 MB cap on uploads.
- **Multi-replica.** `STATELESS_HTTP=true` makes the transport stateless so
  requests can land on any replica; the default keeps sessions in-process.

## Root module map

| Module | Responsibility |
|--------|----------------|
| `main.py` | Tool declarations, startup sequence, health routes, server launch |
| `config.py` | `Config` model, `from_env()`, logging setup, `get_config()` singleton |
| `async_runner.py` | `run_blocking()` and the bounded executor |
| `middleware.py` | API-key middleware |
| `librechat_integration.py` | request-header user context; `upload_and_format_response()` |
| `template_utils.py` | template file resolution across custom/default and container/local dirs |
| `template_registry.py` | YAML spec merging and live tool removal, shared by both dynamic-tool modules |
| `inline_markdown.py` | the inline-markdown grammar shared by the Word and PowerPoint renderers |
| `image_utils.py` | image download, data-URI decoding, validation, SSRF guard |
| `warning_channel.py` | the severity vocabulary, the `DocumentWarning` record and the per-build `WarningChannel` the builders write to |
| `metrics.py` | in-process counters and recent-log buffer for the admin Status page |
| `upload_tools/` | strategy dispatch and one module per backend |
| `docx_tools/`, `xlsx_tools/`, `pptx_tools/`, `email_tools/`, `xml_tools/` | one package per document type; see [`tools/`](tools/) |
| `admin/` | optional template-admin UI |

Shared root modules are described in more detail in
[`shared-modules.md`](shared-modules.md).
