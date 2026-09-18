# Templates

Customise the look of generated documents by providing your own templates,
register several PowerPoint designs by name, and turn a Word or email file
with placeholders into a tool of its own.

## Static Templates

Place files in the `custom_templates/` folder:

| Document | Filename | Notes |
|----------|----------|-------|
| PowerPoint 4:3 | `custom_pptx_template_4_3.pptx` | Or register several by name — see below |
| PowerPoint 16:9 | `custom_pptx_template_16_9.pptx` | Or register several by name — see below |
| Word | `custom_docx_template.docx` | |
| Email wrapper | `custom_email_template.html` | Base it on `default_templates/default_email_template.html` |
| Excel named styles | `custom_xlsx_template.xlsx` | Any named cell style in it becomes usable as `style:<Name>` in a `styles:` directive; see [Markdown reference](markdown-reference.md#excel) |

## Named PowerPoint Templates

### Register several decks, and how layouts are matched

Dropping the two files above is still all you need. To offer your AI a *choice*
of decks — or to configure one — create `config/pptx_templates.yaml`:

```yaml
templates:
  - name: corporate_16_9
    description: Brand deck, widescreen. Use for client-facing presentations.
    pptx_path: corporate_16_9.pptx    # bare filename, in custom_templates/
    default: true
    defaults:
      footer_text: "ACME s.r.o. · Confidential"
      language: cs-CZ
  - name: legacy_4_3
    pptx_path: legacy_4_3.potx        # .potx works too
```

Each entry becomes a value for the `template` argument, and
`list_presentation_templates` reports them with their aspect ratio and layout
names. Anything under `defaults:` applies when the tool call does not set it.

With `include_layouts`, each entry also carries the diagnostics that say how
the template will actually be used:

| Field | What it answers |
|-------|-----------------|
| `layouts` | every layout with its index, placeholder types and the **role it was detected as** — `null` when it matched none |
| `roles` | the role → layout mapping the builder will use, `null` for a role nothing serves |
| `missing_roles` | roles no layout provides, where slides fall back to a layout by position and warn |
| `content_area` | the rectangle the template reserves for content, as percentages of the slide |

The per-layout role is what makes a mis-detected template diagnosable. A
section layout that carries only a title is `title_only` by signature, so
`section` goes unserved *and* that layout is what `kpi` and `timeline` slides
get when it comes before the real title-only layout — visible at a glance in
`layouts` and `roles`, where a list of layout names said nothing.

The `content_area` is read from the template's own content layout. Slides the
tool positions itself stay inside it — a `blank` or title-only layout keeps
the template's logo, rules and footer, so drawing from the top-left corner
would land on them — and a `blank` slide's own `elements`, whose coordinates
are absolute, can be positioned against it:

```json
{"content_area": {"layout": "Nadpis a obsah",
                  "x": "6.9%", "y": "26.6%", "w": "86.2%", "h": "63.4%"}}
```

**Layouts are matched by name and by shape, never by position.** A template
that reorders, renames or deletes layouts still works — previously a reordered
template silently built the title slide on the section layout. Resolution order:

1. the `layout` field on the slide itself,
2. a `layouts:` mapping in the template's entry,
3. automatic detection from each layout's placeholders,
4. the positional index, with a warning returned to the caller.

Detection reads placeholder *types*, so it is language-independent — a Czech or
German template classifies from the same rules as an English one. It recognises
`title`, `section`, `content`, `two_column`, `comparison`, `image_text`,
`title_only` and `blank`. Vertical-text layouts are never picked automatically,
since they share a signature with their horizontal counterparts; name one
explicitly if you want it.

Only add a `layouts:` mapping when detection picks the wrong one:

```yaml
    layouts:
      content: "Brand Body"
      section: "Brand Divider"
```

Every registered template is opened once at startup and its coverage logged, so
a template missing a layout the tool needs shows up then rather than as a
strangely laid-out deck later. Templates are re-read when the config or
template directories change, so a new or edited template takes effect **without
restarting the server**.

The file `config/pptx_templates.yaml` in this repository documents every option
in full.

You can also register and configure templates from the browser instead of
writing YAML — see [Template Admin UI](admin-ui.md). It reports
each layout's detected role and previews a sample deck through the template
before you save.

## Dynamic Email Templates

Create reusable, parameterized email layouts that your AI can fill in automatically.

### How to set up dynamic email templates

**1.** Create `config/email_templates.yaml`:

```yaml
templates:
  - name: welcome_email
    description: Welcome email with optional promo code
    html_path: welcome_email.html  # must be in custom_templates/ or default_templates/
    annotations:
      title: Welcome Email
    args:
      - name: first_name
        type: string
        description: Recipient's first name
        required: true
      - name: promo_code
        type: string
        description: Optional promotional code
        required: false
```

**2.** Create the HTML file in `custom_templates/welcome_email.html`:

```html
<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8" /></head>
<body>
  <h2>Welcome {{first_name}}!</h2>
  <p>We're excited to have you on board.</p>
  {{#promo_code}}<div class="promo">Use promo code <strong>{{promo_code}}</strong>.</div>{{/promo_code}}
  <p>Regards,<br/>Support Team</p>
</body>
</html>
```

**How it works:**
- Each template becomes a separate AI tool at startup
- Standard email fields (`subject`, `to`, `cc`, `bcc`) are added automatically; declare `file_name` or `add_unique_prefix` as args if a template needs them
- Use `{{variable}}` for escaped text, `{{{variable}}}` for raw HTML

## Dynamic Word (DOCX) Templates

Create reusable Word documents with `{{placeholders}}` that support full Markdown formatting.

### How to set up dynamic DOCX templates

**1.** Create `config/docx_templates.yaml`:

```yaml
templates:
  - name: formal_letter
    description: Generate a formal business letter
    docx_path: letter_template.docx  # must be in custom_templates/ or default_templates/
    annotations:
      title: Formal Letter Generator
    args:
      - name: recipient_name
        type: string
        description: Full name of the recipient
        required: true
      - name: recipient_address
        type: string
        description: Recipient's address
        required: true
      - name: subject
        type: string
        description: Letter subject
        required: true
      - name: body
        type: string
        description: Main body of the letter (supports markdown)
        required: true
      - name: sender_name
        type: string
        description: Sender's name
        required: true
      - name: date
        type: string
        description: Letter date
        required: false
        default: ""
```

**2.** Create a Word document with placeholders and save as `custom_templates/letter_template.docx`:

```
{{date}}

{{recipient_name}}
{{recipient_address}}

Subject: {{subject}}

{{body}}

{{sender_name}}
```

**How it works:**
- Each template becomes a separate AI tool at startup
- Placeholders can be in the document body, tables, headers, and footers
- Placeholder values support full Markdown (bold, italic, lists, headings…). Block-level Markdown (lists, headings, tables) renders only for placeholders in the body; placeholders in table cells, headers and footers take inline formatting
- The placeholder's own formatting — font, size, colour, **bold, italic, underline, highlight** — is captured and applied to the replacement text (markdown in the value, e.g. `**bold**`, still wins where it sets formatting)
- Formatting of the surrounding text in the same paragraph (before/after the placeholder) is preserved

### Word style requirements for custom templates

For proper formatting, make sure these styles exist in your `.docx` template:

| Category | Styles |
|----------|--------|
| Headings | Heading 1 – Heading 6 |
| Bullet lists | List Bullet, List Bullet 2, List Bullet 3 |
| Numbered lists | List Number, List Number 2, List Number 3 |
| Other | Normal, Quote, Table Grid |

> **Tip:** Customize these styles (font, size, color, spacing) in your template — the server will use your styling.

### Custom style mapping (use your own style names)

If your template defines styles under **different names** than the built-ins above, map them in `config/docx_templates.yaml` so rendered Markdown uses them — no need to rename styles in Word.

A top-level `style_mapping:` applies to every document; each template may add its own `style_mapping:` which overrides the global one for that template.

```yaml
# config/docx_templates.yaml

# Global — applies to all conversions:
style_mapping:
  heading_1: "Brand Title"
  list_number: "Brand Numbers"
  quote: "Brand Quote"
  table: "Brand Table"

templates:
  - name: formal_letter
    docx_path: letter_template.docx
    # Per-template override (wins over the global mapping):
    style_mapping:
      quote: "Letter Quote"
    args: [ ... ]
```

**Recognized keys:** `heading_1`…`heading_6`, `list_number` / `_2` / `_3`, `list_bullet` / `_2` / `_3`, `quote`, `table`, `normal`, `code` (style for fenced code blocks).

**Ad-hoc style tag:** to apply any style to a single block without a mapping, put a directive directly above it:

```markdown
<!-- style: Callout -->
This paragraph uses the "Callout" style.
```

Unknown style names fall back to the document default (with a logged warning) rather than failing the document.

