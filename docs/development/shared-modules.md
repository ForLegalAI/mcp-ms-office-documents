# Shared root modules

The modules at the repository root that more than one tool package uses.
Each is small and independent; this page says what each owns and the rule
it enforces. The request path that ties them together is in
[`architecture.md`](architecture.md).

## `config.py`

The only module that reads the environment. `Config.from_env()` builds
nested Pydantic models: `LoggingSettings` (level from `DEBUG`),
`StorageSettings` (strategy plus exactly one populated backend settings
object), `AdminSettings`, and the flat flags for thread offload, worker
count, SSRF override and stateless HTTP. Each backend settings model has a
validator that requires its fields, so a misconfigured deployment fails at
startup with `ValueError("Invalid configuration: …")` rather than on the
first upload.

`get_config()` returns the process singleton and configures logging once.
Invalid numeric values fall back to defaults silently: a bad
`SIGNED_URL_EXPIRES_IN` becomes 3600, a bad `RUN_BLOCKING_MAX_WORKERS`
becomes 4, an unknown `UPLOAD_STRATEGY` becomes `LOCAL`.

The full variable list is in [`../configuration.md`](../configuration.md);
a test guards that the list matches what `from_env()` reads.

## `async_runner.py`

`run_blocking(func, *args, **kwargs)`: awaitable dispatch of a blocking
callable to a bounded process-wide `ThreadPoolExecutor`, or inline on the
event loop when `RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED` is false. The flag
is read on every call so tests can flip the config singleton. The module
docstring records the production incident behind the bounded pool. Every
tool handler, static and dynamic, goes through this and nothing else.

## `middleware.py`

`ApiKeyAuthMiddleware`, a FastMCP middleware registered only when `API_KEY`
is set. Reads `Authorization: Bearer <key>`, a plain `Authorization: <key>`,
or `x-api-key`, compares in constant time, and raises `AuthorizationError`.
Failed attempts are counted and logged at WARNING at most once a minute, so
a misconfigured client cannot flood the log. It reads headers through
`get_http_headers(include={"authorization"})` because FastMCP strips that
header by default.

## `librechat_integration.py`

Two functions on the request path. `extract_user_context_from_request()`
reads `X-User-Id`, `X-User-Email` and `X-Conversation-Id` from the current
HTTP request on the event loop; the headers are trusted as-is, on the
assumption that only LibreChat can reach the endpoint. `upload_and_format_response()`
is the branch between a LibreChat artifact and a traditional URL, and the
place that offloads the synchronous upload. See
[`upload-backends.md`](upload-backends.md).

## `template_utils.py`

Resolves template files across four directories in priority order:
`/app/custom_templates`, `./custom_templates`, `/app/default_templates`,
`./default_templates`. `find_file_in_template_dirs(name)` is the primitive;
`find_docx_template()`, `find_email_template()` and `find_pptx_templates()`
apply the `custom_<kind>_template.<ext>` then `default_<kind>_template.<ext>`
convention. The chosen file and whether it was custom or default are logged
at INFO. The Excel styles module and the PowerPoint registry both go through
the primitive.

## `template_registry.py`

Shared by the Word, email and PowerPoint dynamic registries.
`gather_specs(master_yaml, spec_dir)` merges the hand-written master file
with the admin-written per-template files, the latter winning by name.
`read_spec_file()` is the canonical loader for one `.d` file and tolerates a
`{templates: [spec]}` wrapper. `safe_remove_tool(mcp, name)` removes a live
tool across FastMCP versions. It lives at the root so the core tool modules
never import the optional `admin` package.

## `inline_markdown.py`

The single inline-emphasis grammar. `build_inline_pattern(highlight=,
superscript=, subscript=)` compiles one alternation with a capturing group
so `pattern.split(text)` yields tokens interleaved with plain text, the
convention both renderers use. The flags let a renderer omit spans it cannot
draw: Word asks for all three, PowerPoint omits highlight. Flanking rules
follow CommonMark so arithmetic is not read as emphasis; `ESCAPE_RE` escapes
only ASCII punctuation so `C:\new` survives. The module docstring records
the drift that motivated unifying the two copies. Change the grammar here
and nowhere else; `tests/test_inline_markdown.py` covers it.

## `image_utils.py`

`load_image(source)` returns `(BytesIO, extension)` for an `https` URL or a
`data:` URI. Shared by the Word and PowerPoint renderers.

- `download_image()` validates the scheme, then `_get_following_redirects()`
  fetches with redirects disabled and follows up to five hops by hand,
  calling `assert_url_is_public()` on every hop, under one shared timeout.
  Content type must be an allowed image type and the body is capped at 10 MB
  while streaming.
- `assert_url_is_public()` resolves the hostname through `getaddrinfo` (so
  alternative IP encodings normalise) and refuses any address that is not
  globally routable: private, loopback, link-local including the cloud
  metadata address, reserved, multicast. `SSRF_ALLOW_PRIVATE_ADDRESSES=true`
  skips the check. DNS rebinding between the check and the connect is a
  documented residual risk.
- `decode_data_uri()` accepts base64 data URIs of allowed types, with the
  size guard applied before decoding.

`tests/test_image_ssrf.py` covers the guard.

## `metrics.py`

In-process counters per dynamic template (`record_call`, `record_error`,
`tool_stats`) under a lock, plus a `RecentLogHandler` ring buffer of 300
records that the admin app attaches to the root logger when enabled. No
external dependencies, so the core tool modules can import it. `reset()` is
for tests.

## Where the other root files fit

`main.py` is the entry point and is described in
[`architecture.md`](architecture.md). `admin/` is the optional UI, covered in
[`../admin-ui.md`](../admin-ui.md) for users and touched on in
[`dynamic-templates.md`](dynamic-templates.md) for its registration calls.
