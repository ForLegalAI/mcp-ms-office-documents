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
- **Preview with my values…** — a form with one control per argument,
  pre-filled with those same samples, so it stays one click if you do not care.
  Text values are textareas because **Markdown works inside a placeholder** —
  a heading, a list or a bold run in a body placeholder comes out as real Word
  formatting, and a generated `[body]` sample can never show you that.
  Conditionals get checkboxes, so you can see the document with a block
  switched off as well as on. It previews what is on screen, including edits
  you have not saved.
- **Save** — the template becomes a live MCP tool **immediately**, no restart.
- **Edit** later — adjust arguments, **download** the source file the template
  is actually using, or upload a new version over it and re-scan for
  placeholders (existing arguments are kept). The edit page also shows the
  **YAML** the template is stored as, so what you built by clicking is readable
  in the same format `docs/templates.md` teaches.
- **PowerPoint diagnostics** — the analysis card reports the **content area**:
  the rectangle the template reserves for content, as the same percentages
  `list_presentation_templates` returns, and which layout it was read from.
  Slides the tool positions itself stay inside it, so this is what keeps
  generated content off your logo, rules and footer. You are warned when no
  layout declares one (a fixed band is used instead), or when the rectangle
  is implausible.
- **Style names** — if your Word template renames the built-in styles, map them
  under *Advanced*. Every style the renderer understands is there (all six
  heading levels, three numbered and three bulleted list levels, quote, table,
  normal and code), and each dropdown's default option names what that key
  resolves to today — including when a global `style_mapping` sets it.
- **Inspect a hand-written template** — a template defined in your master
  `config/<kind>_templates.yaml` now links to a read-only page showing its
  description, arguments, style mapping, which file it uses and where that
  file lives, plus the same analysis the edit page runs. Nothing on it can be
  edited: the master file is yours.
- **Adopt** — from that page, copy the entry into the templates this UI
  manages, after which it edits like any other. Your master YAML is **not**
  modified; the copy simply takes precedence, **matched by name**. From then
  on the copy is the template — later edits to that entry in your YAML,
  including turning it off, no longer have any effect, and renaming it there
  leaves you with two templates. If the document ships in
  `default_templates/`, it is copied into the uploads directory so you can
  replace it later; a document named like one of the base templates is copied
  under the template's own name instead, so adopting can never change the
  look of every document the server generates.
- **Clone** — start a new template from an existing one, from its row or its
  edit page. Name the copy and everything else comes across: the description
  and every argument with its type, default and description, plus a **copy**
  of the source file under the new name. It is a copy, not a shared file, so
  replacing one template's document never changes the other's. You land on
  the copy's edit page, since the description usually needs changing at once.
  A PowerPoint clone does not inherit the default-template flag — only one
  template can be the default.
- **Disable** — takes the tool off the server without touching anything else.
  The configuration, the arguments and the uploaded file all stay put; the AI
  simply stops being able to call it, and **Enable** puts it back. Reach for
  this rather than Delete for a seasonal template, one being revised, or one
  whose wording is under review. It survives a restart, and saving an edit to
  a disabled template leaves it disabled.
- **Rename** — change the name on the edit page and save. The tool the AI
  calls is renamed; the uploaded source file keeps its own filename, so
  anything else pointing at that file still works.
- **Delete** — asks first, on its own page, listing what goes: the
  configuration and the live tool. The uploaded source file is kept unless you
  tick the box, and it is never removed while another template still uses it.
  If you only want the AI to stop calling it, Disable instead — Delete
  destroys the arguments and descriptions with it.
- **Source files** page — every file in the uploads directory, with what
  still points at it: a template, a hand-written master-YAML entry, or a base
  template. Anything nothing references is marked **Unreferenced** and listed
  first, and only those can be deleted — one at a time, with the filename on
  the confirmation. This is where a source file kept by a delete, or stranded
  by a rename or a replacement, turns up.
- **Status** page — filter the activity log by level, by which part of the
  server logged it, or by a search over the message and logger name, and
  optionally have the page refresh itself. Every filter is in the URL, so a
  filtered view can be bookmarked or pasted to someone else. Also see live
  tool counts and per-template usage (calls/errors/last used this session).

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

