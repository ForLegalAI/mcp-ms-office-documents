# Dynamic template tools

Word and email templates declared in YAML, or saved through the admin UI,
become MCP tools of their own at startup and can be added, replaced or
removed while the server runs. This page explains the mechanism. For **what**
a template file may contain, see [`../templates.md`](../templates.md) and the
commented master files under `config/`.

PowerPoint templates are a different thing: a design, not a fill-in
document. They do not become tools; they become values for the `template`
argument of the presentation tool. See
[`tools/powerpoint.md`](tools/powerpoint.md).

## Where specs come from

| Source | Path | Written by |
|--------|------|------------|
| Master file | `config/docx_templates.yaml`, `config/email_templates.yaml` | a person; heavily commented; never rewritten by tooling |
| Per-template files | `config/docx_templates.d/<name>.yaml`, `config/email_templates.d/<name>.yaml` | the admin UI (`admin/store.py`) |

`template_registry.gather_specs()` merges the two: master entries in their
original order, a `.d` entry replacing a master entry with the same `name`,
and `.d`-only entries appended in name order. It also returns the master
file's top-level mapping so callers can read `style_mapping`. A malformed
`.d` file is logged and skipped; it never aborts the load. The same loader
serves the PowerPoint registry.

`main.py` registers email templates first, then Word templates, whenever
either the master file or the `.d` directory exists.

## What a spec becomes

`_register_single_template()` in `docx_tools/dynamic_docx_tools.py` and
`_register_single_email_template()` in `email_tools/dynamic_email_tools.py`
follow the same steps:

1. **Validate the spec.** `name` is required. `docx_path` or `html_path`
   must be a bare filename; it is resolved through `template_utils`, custom
   directory first. A missing file skips the template with an error log.
2. **Build the argument model.** Each entry in `args` becomes a field on a
   Pydantic model created with `create_model()`. `TYPE_MAP` translates the
   YAML `type` string to a Python type; `enum` produces a `Literal`.
   `required` and `default` decide optionality. The email model starts from
   `BASE_FIELDS` (`subject` required, `to`, `cc`, `bcc` optional); the Word
   model starts empty.
3. **Expose the model in module globals.** FastMCP resolves the handler's
   annotation by name against the module namespace when it builds the
   schema, so the class must be importable from there.
4. **Register an `async def` handler** whose body is
   `await run_blocking(_sync_impl, data)`. The synchronous body renders,
   assembles the file and uploads it in one go. See
   [`architecture.md`](architecture.md#dynamic-template-tools) for why it does
   not go through the shared upload helper.
5. **Record the registration** in a lock-guarded module dict so the admin
   Status page can list live tools.

### Rendering, Word

The `_sync_impl` body opens the template with python-docx, runs
`conditionals.resolve_conditionals()` with the argument payload, then
`_replace_placeholders_in_document()` with every value stringified.

Placeholder replacement is the intricate part. Word splits text across runs
as it is edited, so `{{name}}` is often three runs. `_replace_placeholder_in_paragraph()`
joins the run texts, finds the placeholder, and captures the formatting of
the run it starts in. Then:

- a value with no block-level Markdown is rendered inline through
  `parse_inline_formatting()` into the paragraph, with the placeholder run's
  font, size, colour and emphasis re-applied to the new runs;
- a value with block-level Markdown (`contains_block_markdown()`) is rendered
  through `process_markdown_content(..., return_elements=True)` and the
  produced elements are inserted after the paragraph, with the placeholder
  paragraph's properties propagated.

Body, tables (including nested tables), headers and footers are all scanned.
Table cells, headers and footers get inline rendering only.

Conditionals are block markers on their own paragraphs (`{{#if flag}}`,
`{{^if flag}}`, `{{/if}}`) evaluated against argument truthiness. Unbalanced
markers keep all content and strip the markers; an unknown name keeps its
content. Body only.

### Rendering, email

The body renders the HTML with a default `pystache.Renderer` (so `{{x}}`
escapes and `{{{x}}}` is raw), wraps it in a base64 `MIMEText`, sets
`Subject`, recipients and `X-Unsent`, and uploads. One special case exists:
a `promo_code` argument automatically produces a `promo_code_block` context
value when the template does not declare one. That is a leftover from the
first example template and is worth removing once no template relies on it.

## Live registration

`register_docx_template()` / `register_email_template()` remove any existing
tool of the same name through `template_registry.safe_remove_tool()`, which
prefers FastMCP's `local_provider.remove_tool` and tolerates version
differences, then register afresh. `unregister_*` removes and forgets. The
admin UI calls these after writing the `.d` file and the asset, so a saved
template is callable immediately.

`metrics.record_call()` and `record_error()` are called from the tool body,
keyed by kind and template name, for the Status page.

## The argument-schema rules

Three rules exist because of how MCP clients read schemas. They are tested
by `tests/test_dynamic_args_schema.py`.

- **Never `Optional[...]`.** An optional argument keeps its plain type and is
  simply absent from `required`. `anyOf: [type, null]` puts the description
  beside the `anyOf`, and several clients drop it.
- **A description is a sibling of a flat type.** Same reason.
- **`add_unique_prefix` has no default in the tool body.** The payload is
  built only from declared args, so `payload.get("add_unique_prefix")` yields
  `None` for a template that does not declare it, and `upload_file()` then
  applies the strategy default. A `False` default here would suppress the
  prefix on every backend.

`file_name` and `add_unique_prefix` are ordinary args: a template that wants
them must declare them. The email tool falls back to the subject, then the
template name, for the filename.

## Differences between the two kinds

| | Word | Email |
|-|------|-------|
| `TYPE_MAP` | string, int, float, bool, list | the same plus `dict`/`object` |
| Base fields | none | `subject`, `to`, `cc`, `bcc` |
| Global config | `style_mapping` from the master file, merged with a per-template one | none |
| Escaping | Markdown rendering, no HTML | Mustache HTML escaping on `{{x}}` |
| Conditionals | `{{#if}}` block markers | Mustache sections `{{#x}}…{{/x}}` |
| Headers set | n/a | `Subject`, `To`/`Cc`/`Bcc`, `X-Unsent` only |

## Known limitations

- **Unavailable under `UPLOAD_STRATEGY=LIBRECHAT`.** The tool body calls the
  synchronous `upload_file()`, which refuses that strategy. Fixing this means
  reading the user context on the event loop before dispatch and calling
  `upload_file_async()` from an async body, mirroring the static tools.
- **Single instance.** Live registration assumes one process owns the
  template files. With replicas, use shared storage and restart.
- **Block content is body-only** in Word templates; see rendering above.
- **The `promo_code` special case** in the email body, as above.

## Tests

| File | Covers |
|------|--------|
| `tests/test_docx_templates.py` | Registration from YAML, placeholder replacement across runs, block content insertion |
| `tests/test_docx_placeholder_formatting.py` | Run formatting preserved through replacement |
| `tests/test_docx_conditionals.py` | Marker parsing, balance, nesting, unknown names |
| `tests/test_dynamic_args_schema.py` | The flat-schema rules for both kinds |
| `tests/test_template_registry.py` | `gather_specs()` merging and live (un)registration |
| `tests/test_template_store.py`, `tests/test_admin_app.py` | The admin store and UI paths that write `.d` files |
