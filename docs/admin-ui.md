# Template Admin UI

Prefer clicking over editing YAML? Enable the built-in **template-admin UI** to
create and manage dynamic **Word**, **email** and **PowerPoint** templates from
your browser — no YAML, no restart.

## How to enable and use it

**1.** Set these in your `.env`:

```
ADMIN_ENABLED=true
ADMIN_PASSWORD=choose-a-strong-password   # falls back to API_KEY if omitted
# ADMIN_PATH=/admin                        # optional, this is the default
```

**2.** Start the server as usual. The admin UI is served from the **same port**
as the MCP endpoint:

```
http://localhost:8958/admin
```

**3.** Log in with your `ADMIN_PASSWORD`, then:

- **Upload** a Word `.docx` (or email `.html`) that contains `{{placeholders}}`
  (and optionally `{{#if flag}} … {{/if}}` conditionals). Author it in real Word
  — full fidelity is preserved.
- The UI **auto-detects** every placeholder and conditional and pre-builds the
  argument form. Fill in each argument's type, whether it's required, a default,
  and the description the AI sees.
- **Preview** with sample values (generates a real file; never uploaded anywhere).
- **Save** — the template becomes a live MCP tool **immediately**, no restart.
- **Edit** later — adjust arguments, or **Replace document** to upload a new
  version and re-scan it for placeholders (existing arguments are kept).
- **Status** page — see live tool counts, per-template usage (calls/errors/last
  used this session), and a recent activity & error log (filterable to
  warnings/errors).

**PowerPoint templates work differently.** A Word or email template is a
fill-in-the-blanks document: it declares arguments and becomes an MCP tool of
its own. A PowerPoint template is a *design* — slide master, layouts, theme
fonts and palette — with no placeholders and no arguments. It does not become a
tool; it becomes one more value you can pass as the `template` argument of
`create_powerpoint_presentation`. So its page looks different:

- **Upload** a `.pptx` or a `.potx` (a designer's template file is taken as
  readily as a deck).
- The UI reports the **slide size and aspect**, every **layout** with its
  placeholders and the slide role auto-detected for it, and the **theme** fonts
  and colours. It warns about the things that actually spoil a generated deck:
  a role no layout covers (those slides fall back by position), layouts with no
  footer placeholder, and sample slides left in the file.
- **Override a layout per role** only where the automatic choice is wrong — the
  dropdowns are populated with the template's real layout names.
- Set **deck defaults** — footer text, language, slide numbers — applied
  whenever a tool call does not set the same option.
- **Preview** builds a real six-slide sample deck through the template, using
  the layout choices on screen *including ones you have not saved yet*.
- **Save** — the presentation tool can build on it immediately, no restart.

**How it's stored:** the UI writes one file per template into
`config/docx_templates.d/`, `config/email_templates.d/` or
`config/pptx_templates.d/` (plus the asset into `custom_templates/`). Your
hand-written master `config/*.yaml` files are never modified — UI templates are
simply merged on top of them at load time.

> **Single-instance note:** live, no-restart registration assumes one server
> instance owns the template files (the standard docker-compose setup). For a
> multi-replica deployment, put the template files on shared storage and roll
> the pods to pick up changes.

