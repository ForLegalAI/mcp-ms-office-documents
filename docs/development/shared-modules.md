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

Two orders decide which file wins, and they are independent: the directory
order settles a filename that exists in more than one directory, the
candidate-name order settles `custom_` against `default_`. A test that gives
the two files different names exercises only the second — and one that names
a single pair of directories exercises only that boundary, which is how the
`APP_DEFAULT_DIR` / `LOCAL_DEFAULT_DIR` pair went unguarded through the first
two drafts of its own test. `tests/test_template_resolution.py` therefore
walks the whole list: the same filename goes into every directory and the
winner is removed each round, so each boundary is exercised in turn and a
fifth directory is covered the day it is added.

## `template_registry.py`

Shared by the Word, email and PowerPoint dynamic registries.
`gather_specs(master_yaml, spec_dir)` merges the hand-written master file
with the admin-written per-template files, the latter winning by name, and
drops anything `is_enabled()` rejects unless `include_disabled=True`.
`read_spec_file()` is the canonical loader for one `.d` file and tolerates a
`{templates: [spec]}` wrapper. `is_enabled(spec)` reads the `enabled` key —
absent means enabled, so every spec written before #165 stays live. It also
accepts the string forms PyYAML leaves uncoerced (`n`, `disable`, `disabled`)
and logs anything it recognises in neither direction. `safe_remove_tool(mcp, name)` removes a live
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

## `warning_channel.py`

The channel a builder grows when it has to work around something: the
severity vocabulary (`error`, `warning`, `info`), the `DocumentWarning`
record and the `WarningChannel` collector. Introduced by
[#114](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/114),
which gave Word and Excel what PowerPoint had had since
[#122](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/122).

A warning is not an error: the file was produced, and something in it is not
what the caller asked for. Left in the log that fact never reaches the model
that could fix it, so it rides back in the response instead.

- `DocumentWarning(code, message, severity, location)` — `code` is what a
  caller branches on, `severity` follows from the code, and `location` is
  whatever that tool counts positions by. `as_dict()` spreads the location in,
  so a caller reads `warning["line"]` or `warning["cell"]`, and `str()`
  renders the log line (`line 42: …`).
- `WarningChannel(severities, limit=DEFAULT_LIMIT)` — one per build, passed
  down the call chain as an argument. Builds run concurrently on
  `async_runner` worker threads, so a module-level list would mix two
  callers' documents together. It de-duplicates identical
  `(code, message, location)` entries (a template missing one style warns
  once, not once per paragraph) and caps the number it carries, appending a
  single `warnings_truncated` record instead.

Each tool owns its codes and their severities: `docx_tools/warnings.py`,
`xlsx_tools/warnings.py`. `pptx_tools/warnings.py` shares the severity
vocabulary but keeps its own `SlideWarning` record — a slide index is not a
line or a cell, and its published shape predates this module.
`main._with_warnings()` is the one place a result is widened from a bare URL
string into `{"file", …, "warnings"}` — and, because it is the one place every
channel-carrying tool passes through on success, the one place those warnings
are counted for the Status page. It takes the tool's `kind` and `name` as
required arguments for that reason: a new tool that forgets them fails there
rather than going silently uncounted.

`tests/test_warning_channel.py` covers the record and the collector; each
tool's own codes are covered by `tests/test_<tool>_warnings.py`.

## `metrics.py`

In-process counters per tool (`record_call`, `record_error`,
`record_warnings`, `tool_stats`) under a lock, plus a `RecentLogHandler` ring
buffer of `LOG_BUFFER_CAPACITY` records that the admin app attaches to the
root logger when enabled. No external dependencies, so the core tool modules
can import it.
`reset()` is for tests.

`recent_logs()` filters the buffer by level, top-level logger package and a
case-insensitive substring (matched against the message *and* the logger name)
before applying its limit, so the limit bounds matches rather than records
scanned. `log_sources()` lists the packages present, for the Status page's
filter; `capture_level()` reports what the buffer is recording at, so the page
can withhold a level filter that could never match.

`record_warnings()` is what makes a degrading build visible. A build that
substitutes a style or drops a row still *succeeds*, so `calls` counts it and
`errors` does not — without this it looked identical to a clean one on the
Status page (#173). Warnings are counted per severity rather than totalled,
because `info` is a substitution the caller will not mind and would otherwise
mask an `error`, which means something they asked for is not in the file;
`ToolStat.degraded` is the total excluding `info`. A severity the table does
not recognise counts as degraded rather than being dropped.

Counters cover static tools as well as template-backed ones — a Word
conversion quietly degrading is the same operator question as a template tool
doing it.

## Where the other root files fit

`main.py` is the entry point and is described in
[`architecture.md`](architecture.md). `admin/` is the optional UI, covered in
[`../admin-ui.md`](../admin-ui.md) for users and touched on in
[`dynamic-templates.md`](dynamic-templates.md) for its registration calls.
