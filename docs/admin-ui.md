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
- **Edit** later — adjust arguments, **download** the source file the template
  is actually using, or upload a new version over it and re-scan for
  placeholders (existing arguments are kept). The edit page also shows the
  **YAML** the template is stored as, so what you built by clicking is readable
  in the same format `docs/templates.md` teaches.
- **Style names** — if your Word template renames the built-in styles, map them
  under *Advanced*. Every style the renderer understands is there (all six
  heading levels, three numbered and three bulleted list levels, quote, table,
  normal and code), and each dropdown's default option names what that key
  resolves to today — including when a global `style_mapping` sets it.
- **Delete** — asks first, on its own page, listing what goes: the
  configuration and the live tool. The uploaded source file is kept unless you
  tick the box, and it is never removed while another template still uses it.
- **Status** page — filter the activity log by level, by which part of the
  server logged it, or by a search over the message and logger name, and
  optionally have the page refresh itself. Every filter is in the URL, so a
  filtered view can be bookmarked or pasted to someone else. Also see live
  tool counts, per-template usage (calls/errors/last
  used this session), and a recent activity & error log (filterable to
  warnings/errors).

## Base templates

The **Base templates** page manages the five files that style *every* document
the server generates — the Word template, the email wrapper, both PowerPoint
designs and the Excel named-styles workbook. Previously these could only be
changed by copying files onto the volume by hand.

Each slot shows which file is in use (yours, or the bundled default), lets you
**download** it, **replace** it, and **revert** to the bundled default. A
replacement takes effect on the next document — no restart. Uploads are checked
before they are installed: the wrong extension, or a file that cannot be opened,
is refused and nothing changes.

Each slot also reports what its file offers: the Word template's styles (and a
warning when one the renderer needs is missing), the wrapper's Mustache
variables, a deck's layouts, roles and theme, and — for Excel — every named cell
style, which is the set a `styles:` directive can reference as `style:<Name>`.

> **Excel is the exception:** no default workbook ships. Removing a custom one
> leaves no named styles at all, rather than restoring a fallback.

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

The UI follows your system's **light or dark appearance** automatically, and
loads nothing from the internet — no CDN, no web font — so it looks right on an
air-gapped host.

**How it's stored:** the UI writes one file per template into
`config/docx_templates.d/`, `config/email_templates.d/` or
`config/pptx_templates.d/` (plus the asset into `custom_templates/`). Your
hand-written master `config/*.yaml` files are never modified — UI templates are
simply merged on top of them at load time.

> **Single-instance note:** live, no-restart registration assumes one server
> instance owns the template files (the standard docker-compose setup). For a
> multi-replica deployment, put the template files on shared storage and roll
> the pods to pick up changes.

