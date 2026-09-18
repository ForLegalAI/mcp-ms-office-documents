# PowerPoint slide reference

The `create_powerpoint_presentation` tool takes a list of typed slide
objects. This page lists every slide type and field. For templates and
layout matching see [Templates](templates.md); for how the tool is built see
the [developer page](development/tools/powerpoint.md).

**Tool parameters** (`create_powerpoint_presentation`):

| Parameter | Description |
|-----------|-------------|
| `slides` | Ordered list of slide objects (below). Required. |
| `format` | `16:9` (default) or `4:3`. |
| `author` | Stored in document properties. |
| `footer_text` | Shown on every slide whose layout has a footer placeholder. |
| `show_slide_numbers` | Slide numbers on every slide. |
| `language` | BCP-47 proofing tag, e.g. `cs-CZ`. Set it when the deck is not in the template's language, or Word/PowerPoint flags every word as misspelled. |
| `template` | Name of a registered template to build on; overrides `format`. Call `list_presentation_templates` to see the names. See [Templates](templates.md). |
| `file_name` | Output filename without extension. |

Every slide takes `type` plus optional `title`, `notes` (speaker notes) and `layout`.

| `type` | Fields |
|--------|--------|
| `title` | `subtitle?` |
| `section` | — |
| `content` | `body` |
| `two_column` | `left`, `right` — each `{heading?, body}` |
| `table` | `rows`, `align?`, `header_color?`, `zebra?`, `font_size?` |
| `chart` | `chart_type`, `categories`, `series`, `legend?`, `data_labels?`, `number_format?`, `chart_title?`, `x_title?`, `y_title?` |
| `scatter` | `series` (`{name, points: [[x, y], …]}`), `legend?`, `chart_title?`, `x_title?`, `y_title?` |
| `image` | `source`, `caption?`, `body?` |
| `quote` | `text`, `attribution?` |
| `kpi` | `items` (`{value, label, delta?}`), 2–4 read best |
| `timeline` | `steps` (`{label, detail?}`), `style?` (`chevron` or `box`) |
| `agenda` | `items?` — omit to build it from the deck's own `section` slides |
| `closing` | `subtitle?`, `contact?` |
| `blank` | `elements` — positioned items, each `{kind: text\|image\|shape, x, y, w, h?}` with lengths in inches (`1.5`, `"1.5in"`) or as a share of the slide (`"40%"`) |

**Sections.** Every `section` slide also starts a section in PowerPoint's outline pane and slide sorter, so the presenter sees the deck's structure rather than a flat list. Slides before the first section slide go in a "Default Section", as PowerPoint would name it.

**Blank slides** are the escape hatch for the one layout no typed slide fits. Elements draw in order, so a later one sits on top; a `text` element takes the same inline markdown as a bullet, a `shape` is one of `rectangle`, `rounded_rectangle`, `ellipse`, `chevron`, `arrow` with an optional `fill` and centred `text`, and an `image` keeps its aspect ratio within its box. Anything that would run past the slide edge is shrunk to fit and reported in `warnings`; anything starting off the slide is skipped and reported. A `title` on a blank slide is kept: the blank layout has no title placeholder, so it is drawn as a text box where the template puts its titles, and elements you position draw over it. Element coordinates are absolute on the slide, and a blank layout usually still carries the template's logo, header rule and footer — `list_presentation_templates` with `include_layouts` reports each template's `content_area`, the band that stays clear of them, as percentages you can use directly.

**Body text.** `body` takes either a Markdown bullet string or explicit bullet objects. Prefer the string:

```json
{"type": "content", "title": "Q3 results",
 "body": "- Revenue **up 12%**\n  - EMEA +18%\n  - APAC +4%\n- Churn flat"}
```

Indent child items with any consistent unit — two spaces, four spaces or a tab. A line without a `-` marker becomes a top-level bullet. The explicit form is `[{"text": "…", "level": 2}]`, where `level` is 1 (outermost) to 5.

**Inline formatting** works in every text field, table cells included: `**bold**`, `*italic*`, `***bold italic***`, `~~strikethrough~~`, `__underline__`, `` `code` ``, `^superscript^`, `~subscript~`, and `[links](https://example.com)`. The same grammar drives the Word tool, so text formats identically in both. A marker only formats when it hugs its text (`**bold**`, not `** bold **`), so prose like `5 * 3 * 2 = 30` is left alone. Escape a literal marker with `\*`, or wrap it in backticks.

**Tables** take raw values — numbers and `null` are fine, not just strings:

```json
{"type": "kpi", "title": "At a glance",
 "items": [{"value": "€4.2M", "label": "ARR", "delta": "+12% vs Q2"},
           {"value": "18%", "label": "Churn"}]}
```

```json
{"type": "table", "title": "Pricing",
 "rows": [["Plan", "Users", "Price"], ["Basic", 5, 9.0], ["Pro", 25, 29.0]],
 "align": ["left", "right", "right"], "header_color": "accent1"}
```

**Colours** accept 6-digit hex with or without `#`, or a theme name (`accent1`…`accent6`, `dark1`, `dark2`, `light1`, `light2`). Prefer a theme name so the deck follows your template's palette.

**Images** take an https URL or an inline data URI, so an image you already hold can be placed without publishing it first:

```json
{"type": "image", "source": "data:image/png;base64,iVBORw0KGgo…", "caption": "Fig 1"}
```

An image slide uses the template's own picture layout when it has one, so the picture is framed, cropped and positioned the way the template's designer intended, with `body` and `caption` in that layout's text area. Without such a layout — or when the one the template has has no room for the text — the picture is scaled into a rectangle on the content layout, with `body` beside it and `caption` underneath.

**Warnings.** When the deck is produced but not exactly as asked — an image that would not load, body text shrunk to fit, a footer dropped because the layout has no placeholder — the result carries a `warnings` list alongside the file instead of leaving it in the server log:

```json
{"file": "https://…/deck.pptx", "slide_count": 12,
 "warnings": ["slide 4: body text is about 1.9x the space available and will be shrunk to fit; consider splitting it across slides."]}
```

**Compatibility.** The previous key names (`slide_type`, `slide_title`, `slide_text`, `indentation_level`, `speaker_notes`, `table_data`, `alternate_rows`, `image_url`, `image_caption`, `quote_text`, `quote_author`, `left_column`, `right_column`, `chart_data`, `has_legend`, `legend_position`) are still accepted and mapped onto the current ones, with a note in the log. They will be removed in a future release.

**Clients that cannot send objects.** `slides` is published as one flat object schema — every field of every slide type in a single `properties` map, each saying which types accept it — rather than as a `oneOf` union of the fourteen types. A union is the better model but not a portable declaration: clients that bridge MCP to a provider without `oneOf`/`$ref` drop what they cannot express, show the model `slides: array of string`, and then reject the call against the union they kept, so *every* call fails with a schema error no matter what the model sends. Validation is unchanged — it runs on this server, against the union, and names the slide and field it rejected.

As a last resort a slide sent as a **string** is read rather than refused: a JSON object is the slide it encodes, and anything else is markdown for one slide (`# Heading` plus text is a `title` slide, otherwise a `content` slide with the rest as its body). The whole deck arriving as one JSON string is unpacked the same way. A string that starts like JSON but does not parse is reported instead — a deck truncated in transit must not quietly arrive as one slide titled with the raw JSON. This fallback covers titles and bullets only — send slide objects for everything else.
