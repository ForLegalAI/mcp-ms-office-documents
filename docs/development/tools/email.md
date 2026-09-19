# Email tool (`email_tools`)

Builds an unsent email draft (`.eml`) whose HTML body is wrapped in a
Mustache template. This page explains how the package does it. For **what**
the tool accepts, see the tool description in `main.py` and
[`../../templates.md`](../../templates.md) for the email wrapper template. For
the request path around the build step, see
[`../architecture.md`](../architecture.md).

## Entry points

| Name | Where | Used by |
|------|-------|---------|
| `create_email_draft` | MCP tool declared in `main.py` | MCP clients |
| `_create_eml_buffer()` | `email_tools/base_email_tool.py` | `main.py`, via `run_blocking` |
| `create_eml()` | `email_tools/base_email_tool.py` | direct library use; builds and uploads synchronously |
| YAML-defined email tools | `email_tools/dynamic_email_tools.py` | MCP clients; see [`../dynamic-templates.md`](../dynamic-templates.md) |

The static tool takes `content` (an HTML fragment), `subject`, `to`, `cc`,
`bcc`, `priority` and `language`. `file_name` and `add_unique_prefix` go to
the upload step.

## Pipeline

```
content (HTML fragment), subject, recipients, priority, language
  │
  ▼  base_email_tool._create_eml_buffer()
  ├─ validate priority ∈ {low, normal, high}; require content and subject
  ├─ _load_template()            template_utils.find_email_template()
  │                              custom_email_template.html → default_email_template.html
  ├─ pystache render with escaping DISABLED
  │    {{language}}  sanitised (quotes stripped)
  │    {{subject}}   html.escape()d by hand
  │    {{{content}}} inserted raw
  ├─ MIMEText(html, 'html', 'utf-8'), base64 transfer encoding
  ├─ headers: To/Cc/Bcc, Subject (RFC 2047), Date, Content-Language,
  │           Accept-Language, X-Priority/X-MSMail-Priority/Importance, X-Unsent: 1
  ▼
BytesIO of msg.as_bytes()
```

`X-Unsent: 1` is what makes Outlook open the file as an editable draft
rather than a received message.

## Module map

| Module | Owns |
|--------|------|
| `base_email_tool.py` | Template load, Mustache render, MIME assembly, the two entry points |
| `dynamic_email_tools.py` | YAML-driven email tools: argument model, render, MIME assembly, live registration. Documented in [`../dynamic-templates.md`](../dynamic-templates.md) |

Templates live in `default_templates/default_email_template.html` (the
wrapper for the static tool) and `default_templates/broadcast_email_style_1.html`
(the example dynamic template). Both are resolved through `template_utils`.

## How the interesting parts work

### Escaping is the caller's job

The static tool renders with `pystache.Renderer(escape=lambda u: u)`, so
Mustache does no escaping at all. The subject is escaped by hand before it
goes into the context and the language has quotes stripped; the body is
inserted verbatim because it is meant to be HTML. The tool description tells
the model which tags to use and to send a fragment, not a document. Nothing
sanitises that fragment server-side.

The dynamic email tools use a default `pystache.Renderer`, so there
`{{name}}` escapes and `{{{name}}}` is raw. The two tools therefore have
opposite defaults; keep that in mind when moving a template between them.

### Static and dynamic drafts differ

The static tool sets `Date`, `Content-Language`, `Accept-Language` and the
priority headers. The dynamic tools set only `Subject`, recipients and
`X-Unsent`. A dynamic template that needs the proofing language must carry
it in the HTML `lang` attribute itself.

## Extension points

| To add… | Edit | Notes |
|---------|------|-------|
| A new header on static drafts | `base_email_tool._create_eml_buffer()` | Also add a tool parameter in `main.py` |
| A new template variable for the static wrapper | the `context` dict in `_create_eml_buffer()` and the HTML template | Decide whether it is escaped; the renderer will not do it |
| A new dynamic email template | `config/email_templates.yaml` or the admin UI | No code change; see [`../dynamic-templates.md`](../dynamic-templates.md) |

## Invariants and gotchas

- **Escaping is disabled in the static renderer.** Anything added to the
  context must be escaped explicitly unless it is meant to be HTML.
- **The default language comes from `EMAIL_DEFAULT_LANGUAGE`**, not from the
  signatures. It was a hard-coded `cs-CZ` in both the tool parameter and
  `_create_eml_buffer()` until #116, which made it a setting and changed the
  default to `en-US`; a Czech deployment now sets the variable. The buffer
  function takes `language=None` and resolves it from config, so there is one
  place the default lives.
- **`priority` is validated case-insensitively** but the check reads
  `priority.lower()` twice; a `None` priority would fail on the second read.
  The tool parameter defaults to `"normal"`, so this cannot happen through
  MCP.

## Tests

| File | Covers |
|------|--------|
| `tests/test_email_creation.py` | The draft itself: recipient joining, the RFC 2047 subject, the three priority headers moving together, `X-Unsent`, the base64 UTF-8 HTML body, what is escaped (`subject`) and what is not (`content`), template resolution including the custom override and the missing-template error, the uploading wrapper, and the dynamic tool's deliberately smaller header set |
| `tests/test_email_language.py` | The `EMAIL_DEFAULT_LANGUAGE` setting and where the tag lands |
| `tests/test_dynamic_args_schema.py`, `tests/test_template_registry.py` | The dynamic email tools' argument schema and registration |

The draft is built with `_create_eml_buffer()` and parsed back with
`email.message_from_bytes()`, so the assertions are on what a mail client
reads rather than on the code's intermediates. The body is base64, so a test
that looks at it decodes the payload first.

Both modules were added by [#112](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/112) and [#116](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/116); before them the static tool had no test of its
own and was exercised only incidentally by `tests/test_run_blocking.py`,
`tests/test_upload_unique_prefix.py` and `tests/test_librechat_integration.py`,
none of which assert on the draft they produce.

## Known limitations

- **No HTML sanitisation.** The body fragment is trusted as sent.
- **Dynamic drafts carry fewer headers** than static ones, as above.
- **One template variable set.** The static wrapper knows only `language`,
  `subject` and `content`.
