# PowerPoint tool (`pptx_tools`)

Builds a `.pptx` deck from a list of typed slide objects with python-pptx,
on top of a registered template whose layouts are matched by name and shape.
This page explains how the package does it. For **what** the tool accepts,
see the user reference: [`../../powerpoint-slides.md`](../../powerpoint-slides.md)
and [`../../templates.md`](../../templates.md). For the request path around the
build step, see [`../architecture.md`](../architecture.md).

## Entry points

| Name | Where | Used by |
|------|-------|---------|
| `create_powerpoint_presentation` | MCP tool declared in `main.py` | MCP clients |
| `list_presentation_templates` | MCP tool declared in `main.py` | MCP clients, to discover template and layout names |
| `_create_presentation_buffer()` | `pptx_tools/base_pptx_tool.py` | `main.py`, via `run_blocking`; returns `(BytesIO, warnings)` |
| `PowerpointPresentation` | `pptx_tools/slide_builder.py` | the buffer function; tests; the admin UI preview (with `template_spec=`) |
| `create_presentation()` | `pptx_tools/base_pptx_tool.py` | direct library use; builds and uploads synchronously, drops the warnings |

This tool had the first **warnings channel**. Everything the builder had
to work around, such as an image that would not load, text shrunk to fit, a
layout the template does not provide or a footer dropped, is collected on
`PowerpointPresentation.warnings` and returned alongside the file so the
calling model can correct its next call. `main._with_warnings()` wraps the
result in `{"file", "slide_count", "warnings"}` when there is anything to
report — the same helper the Word and Excel tools use since
[#114](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/114) gave
them channels of their own.

Each entry is a `SlideWarning` from `warnings.py`, not a sentence: `code`,
`slide` (the index in the caller's list, None for the deck), `severity` and
`message`. The severity vocabulary comes from the shared
[`warning_channel.py`](../shared-modules.md#warning_channelpy), so `error`
means the same thing in every tool; the record itself stays here, because a
slide index is not a line or a cell and this shape is already published. `_warn(index, code, message)` and `_warn_deck(code, message)`
record them, `make_warning()` reads the severity out of `WARNING_SEVERITY` so
one code always means one severity, and `str(warning)` renders the line the
channel used to hold — which is what the log, the admin preview header and
`warning_messages` still use. Adding a warning means adding its code to that
table; `test_every_code_has_a_severity` fails otherwise
([#122](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/122)).

## Pipeline

```
slides (list of dicts, models, or strings)  +  format / template / author / footer_text / …
  │
  ▼  PowerpointPresentation.__init__()                        slide_builder.py
  ├─ schema.coerce_slides()                                    validate → typed models
  │    ├─ _migrate_list()        a JSON-string deck or slide is parsed; legacy keys renamed
  │    ├─ slide_from_text()      a markdown string becomes a title or content slide
  │    └─ TypeAdapter(Slides)    discriminated union on "type", extra="forbid"
  ├─ templates.select_template(name, format)                  templates.py
  │    └─ open_template()        .potx rewritten in memory; else Presentation()
  ├─ LayoutResolver(presentation, spec.layouts)               layouts.py
  │    └─ classify_layout()      role per layout from placeholder types
  ├─ template defaults           footer_text, language, slide numbers, table/chart defaults
  ├─ _remove_template_slides()   unless strip_slides: false
  ▼
  _build_slides(): one builder per slide type
  ├─ _new_slide()                role_for_slide() → resolver.resolve(role, slide.layout)
  ├─ _apply_title()              placeholder, else a text box in the template's title style
  ├─ helpers.body_to_bullets()   markdown string or Bullet objects → levels
  ├─ SlideHelpers._fill_bullets / _create_styled_table / _add_image / _add_text_box
  ├─ inline_formatting.apply_inline_formatting()               runs, links, super/subscript
  ├─ chart_utils.add_chart_to_slide / add_scatter_to_slide
  └─ helpers.estimate_text_fill() → apply_autofit(); warn when far over
  ▼
  _apply_sections()              p14 sectionLst XML from section slides
  _apply_footer_and_slide_numbers()   clone layout placeholders idx 11 / 12 onto each slide
  _apply_language()              lang on every run, table cell and notes
  core_properties.author
  ▼
save() → BytesIO; warnings list
```

Unlike the Word and Excel tools, a failure inside one slide builder fails
the whole deck: `_build_slides()` re-raises as `ValueError("Error creating
slide N (type): …")`, which the handler turns into a `ToolError`. Recoverable
problems go to the warnings list instead.

## Module map

| Module | Owns |
|--------|------|
| `base_pptx_tool.py` | The two entry points; nothing else |
| `schema.py` | The fourteen slide models and their sub-models, colour and position validators, legacy-key migration, the text-slide shim, the flat published schema, `coerce_slides()` |
| `slide_builder.py` | `PowerpointPresentation`: template selection, one `_build_*` method per slide type, sections, footer and slide numbers, language, the warnings list |
| `helpers.py` | `SlideHelpers` mixin (titles, placeholders, bullets, tables, images, notes) and free functions: `body_to_bullets()`, `parse_table_data()`, `estimate_text_fill()`, `apply_autofit()`, `fit_table_font_size()`, `set_runs_language()`, `resolve_fill()` |
| `layouts.py` | Layout roles, `classify_layout()`, `role_for_slide()`, `LayoutResolver` |
| `placeholder_style.py` | Reading placeholder geometry, character style, size and list style from a template, and replaying them on plain text boxes (`read_title_style()`, `read_content_rect()`, `content_columns()`, `read_body_font_size()`, `draw_title_box()`, `apply_list_style()`) |
| `templates.py` | `TemplateSpec`, the registry loaded from YAML with an mtime-fingerprint cache, `.potx` handling, `select_template()`, `validate_templates()` |
| `chart_utils.py` | Category charts from `CategoryChartData`, scatter from `XyChartData`, legend, title, data labels, axis titles |
| `inline_formatting.py` | Renders the shared inline grammar into python-pptx runs |
| `constants.py` | Aspect ratios, positional layout fallbacks, typography, autofit ratios, table colours (theme names, not literals) |
| `warnings.py` | The warnings channel as data: `SlideWarning`, the codes, and the one severity per code (severities from the shared `warning_channel.py`) |
| `text_metrics.py` | Font resolution (metric-compatible substitutes, generic fallback) and wrapped line counting through Pillow |

Two root modules are part of this pipeline: `inline_markdown.py` holds the
emphasis grammar (PowerPoint asks for superscript and subscript but not
highlight), and `image_utils.py` downloads and validates images including
the SSRF guard.

## How the interesting parts work

### Validation: a strict union, published flat

Each slide type is a Pydantic model with `extra="forbid"`, so a misspelled
key is an error rather than a silently empty element. Values are loose where
models are loose: table cells accept numbers and `null`, colours accept
`#RRGGBB`, `RRGGBB` or a theme name, bullet levels are clamped rather than
rejected.

The union is the right model and the wrong thing to publish. Several
providers' function-calling dialects have no `oneOf`/`$ref`/`discriminator`,
and clients that bridge to them show the model `slides: array of string`
while still validating against the union, so every call fails client-side.
`flat_slide_schema()` therefore builds one object schema with every field of
every type, each field's description naming the types that accept it, and
`_simplify()` strips everything outside the keyword subset those dialects
read. The tool parameter declares that flat schema through `WithJsonSchema`;
validation still runs against the union in `coerce_slides()`, on the worker
thread, where the error can say `slide 2 -> rows.0: …`.

Two shims sit in front of validation. `migrate_legacy_slide()` renames the
previous key spellings (`slide_type`, `slide_text`, `chart_data`, …) so old
client prompts still work, logging once per call. `slide_from_text()` reads
a slide that arrived as a string: JSON is parsed, anything else is one slide
of markdown. This is a fallback for the same clients, deliberately not
documented as a second spelling.

### Templates and the registry

`templates.load_specs()` merges `config/pptx_templates.yaml` with
`config/pptx_templates.d/` through `template_registry.gather_specs()`. With
no registry at all it synthesises two specs named `16_9` and `4_3` from the
historical filename slots, so an existing deployment is unchanged. That
fallback covers "nothing readable is configured" — a missing config, or
entries too malformed to build a spec from — but deliberately *not* every
template being disabled (`enabled: false`), which is a choice someone made:
falling back there would hand back the very templates the admin just took
away, under the built-in names, and the deck would still build off the wrong
design. The cache
key is a fingerprint of every watched file's modification time, including
each template file, so overwriting a template or dropping in a new one takes
effect without a restart. The aspect ratio is read from the file, never from
the config, and exactly one spec is marked default.

`templates.content_area_summary()` formats that rectangle as percentages of
the slide. It is public because the admin UI reports the same rectangle in the
same units (#171): one formatter means the page and the tool's own diagnostics
cannot drift apart on the number an admin is checking.

`select_template()` picks by name, else by aspect, else the default, and
returns a warning string whenever it had to substitute. A `.potx` is opened
by rewriting its content type in memory, since python-pptx refuses the
template flavour outright.

Template `defaults` apply when the tool call does not set the same option.
`_setting()` reads `model_fields_set` rather than the value, because `zebra`
and `data_labels` default to real booleans and their value alone cannot say
whether the caller chose it. Booleans from YAML may arrive as strings and are
coerced by `_coerce_bool()`; a table `font_size` default is clamped to the
schema's range by `_coerce_font_size()` because it bypasses the schema.

### Layout resolution

`LayoutResolver` never indexes `slide_layouts` by position. Each slide type
asks for a **role** (`title`, `section`, `content`, `two_column`,
`comparison`, `image_text`, `title_only`, `blank`, `closing`); `role_for_slide()` picks
`comparison` over `two_column` when either column has a heading, `blank`
for an untitled quote so no empty title placeholder is left behind, and
`image_text` for an image slide when `resolver.provides()` says the template
has such a layout — the one role that depends on the template rather than on
the slide alone.
Resolution order is the slide's own `layout` name, then the registry's
`layouts:` mapping, then detection, then a **near-neighbour role** from
`ROLE_ALTERNATIVES` with a warning, then the positional index with a warning,
and finally the last layout rather than an `IndexError` on a trimmed template.

The neighbour step exists because the positional index is a guess about a
template that has already proved unusual. On the template in
[#194](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/194) the
`comparison` role was unprovided and position 4 was Title Only, so a
two-column slide landed on a layout with no body placeholder and lost both
columns. Every alternative listed can still hold the content: `comparison`
falls to `two_column` (headings inline) and then `content` (columns merge), so
the deck degrades instead of dropping text.

`LayoutResolver.describe()` is what `validate_templates()` reports and
`list_presentation_templates` returns: each layout with its index, placeholder
types, whether it is vertical, and the role it classified as. Roles are held
in a list parallel to the layouts rather than keyed by name, because two
layouts in one template may share a name. Together with `coverage()` and
`missing_roles()` this is the answer to "why did my slide come out on that
layout"
([#121](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/121)).

`closing` is in `CONFIGURED_ROLE_DEFAULT`: no signature detects it, because a
contact or thank-you slide is a designer's layout rather than a shape — the
template in #194 has a "kontakt" layout carrying a QR code, a photo and the
firm's details. Naming it a role is what lets a template map it in the
registry's `layouts:` block; unmapped, it resolves to `title` **without** a
warning, since that is its documented default and not a substitution.
`missing_roles()` and the admin's analysis both exclude it, or every template
would be reported as lacking something no placeholder arrangement can supply.

`classify_layout()` reads placeholder *types*, ignoring date, footer and
slide-number chrome, so it is language-independent. It is deliberately
conservative: vertical-text layouts and "Content with Caption" return `None`
and are left for an explicit name. The distinguishing tests are a picture
placeholder (image_text), a subtitle (title), no content (title_only), two
content placeholders (two_column), and for one, `BODY` means section and
`OBJECT` means content.

`comparison` is the one role that is **not** decided by counting. Four content
placeholders were enough to claim it, and the template in #194 spent them on
three cards side by side plus a caption bar — so every two-column slide in the
deck was laid out on a three-card layout. A layout now has to resolve, through
`content_columns()`, to exactly two columns *each with a heading strip*.
Anything else is a shape this vocabulary has no name for and returns `None`,
which is what leaves a three-card layout selectable by name and nothing else.

### Titles on a layout that has none

A Blank layout has no title placeholder — that is what makes it blank — and a
trimmed corporate layout may have none either. `_apply_title()` therefore
falls back rather than dropping the text: `LayoutResolver.title_style()`
reads the first title placeholder it finds, preferring the body layouts
(`content`, `title_only`, two-column, `image_text`, `section`) over the cover
layout, whose title is a centred block halfway down the slide, and
`placeholder_style.draw_title_box()` draws a text box with that geometry.

The character style is the `<a:defRPr>` of the first paragraph-property
element in the chain *layout placeholder → master placeholder → the master's
`<p:titleStyle>`*, copied onto the run wholesale rather than re-derived
attribute by attribute, so theme references (`+mj-lt`, `schemeClr`) survive
and keep following the theme. Alignment and vertical anchoring come from the
same chain; `apply_autofit()` then replaces python-pptx's default
`<a:spAutoFit>` so a long title shrinks instead of resizing the band.

Only a template with no title placeholder on any layout still warns: the box
goes in a band across the top of the slide, and that is a guess worth saying
out loud.

### Content placement

Builders that draw their own shapes (table, image, chart, scatter, quote,
KPI, timeline) call `_content_slide()`, which resolves the layout, applies
the title, then takes the rectangle of the **largest** body placeholder and
removes it. Largest, not first: on a Comparison layout the first content
placeholder is the small heading strip.

With no body placeholder on the slide's layout — Blank, Title Only —
`_content_area()` falls back to `LayoutResolver.content_area()`: the largest
body placeholder of a layout that does have one, preferring the
single-content layouts over a two-column one, whose body is half the slide.
A blank layout is rarely an empty canvas; templates keep a logo, a header
rule and a footer band on it, and the old fixed rectangle 1.5 inches down
drew over them ([#119](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/119)).
That fixed rectangle is now the last resort, for a template with no body
placeholder on any layout.

`validate_templates()` reports the same rectangle as `content_area` in
percentages, which `list_presentation_templates` passes through. That is the
answer for a caller positioning a `blank` slide's elements: those coordinates
stay absolute on the slide — changing them would silently move every existing
deck — so the safe band is published instead.

### Colour follows the template

Anything the builder *draws* rather than places in a placeholder is a plain
text box, so it inherits the presentation's `<p:defaultTextStyle>` — `tx1`,
black — not the body style a placeholder gets. On the dark template in #194
the KPI figures and the timeline detail lines were black on near-black.

`read_body_color()` reads the master's `<p:bodyStyle>` colour and returns a
**`MSO_THEME_COLOR`** for a scheme colour, so it keeps tracking the theme
instead of being flattened to whatever it resolves to today; an `srgbClr`
comes back as an `RGBColor`. `_paint()` applies it to the KPI cells, the
timeline detail captions, the quote, blank-slide text elements and the
bulleted box drawn beside a picture or chart. `_paint_chart()` does the same
through `chart.font`, which chart text needs because it lives in its own part
and inherits nothing from the slide.

Two things are deliberately left alone: a timeline chevron's label, which
takes its colour from the autoshape's `<p:style>` against the accent fill it
sits on, and a drawn title, which `draw_title_box()` styles from the
template's *title* style. A template stating no body colour paints nothing.

Table fills were literals — Office's old default blue and a grey beside it —
so a table came out that blue on every template. `TABLE_HEADER_FILL` and
`TABLE_ALT_ROW_FILL` are now the theme names `accent1` and `bg2`;
`_set_cell_fill()` already wrote a theme name as `schemeClr`, only the default
was not one. `TABLE_HEADER_TEXT` stays an explicit white: it is paired with
`accent1`, and choosing it from the theme would need a luminance decision this
tool has no safe way to make. Both remain overridable per slide through
`header_color` and `fills`, and per template through the registry's `table`
defaults.

### Unused placeholders are removed

`add_slide()` copies every placeholder its layout defines, so a layout that
offers more than the slide filled left empty prompt boxes in the deck: the
third card of a three-card layout, the heading strip of a Comparison column
given no heading, the body of a Section Header (a `section` slide carries only
a title), the subtitle of a title slide without one. They neither print nor
appear in a slideshow, but they are the first thing anyone opening the file to
edit it sees, and on the template in #194 there were three on a single slide.

`_drop_unused_placeholders()` runs once after every slide is built and removes
any non-chrome placeholder still holding an empty text frame. A placeholder
that took a picture, table or chart is no longer an `<p:sp>` with a text
frame, so filling one keeps it; date, footer and slide-number placeholders are
skipped because `_apply_footer_and_slide_numbers()` runs after this pass.
Removing a placeholder changes nothing a reader sees, and PowerPoint's Reset
Slide restores it from the layout.

### Two columns, matched by geometry

`two_column` slides do **not** address placeholders by `idx`. The indices
PowerPoint's own Two Content (1, 2) and Comparison (1-4) layouts use are a
convention, not a rule, and a corporate template routinely breaks it: the one
in [#194](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/194)
numbered its two cards 4 and 2 — left and right *in that order* — and its
comparison layout 4, 13, 14, 15. Addressing by number wrote the right column
into the left card, dropped the left column and both headings, left three
placeholders empty, and reported none of it.

`content_columns()` in `placeholder_style.py` reads the shape of the layout
instead:

1. content placeholders that overlap horizontally by more than half the
   narrower one are **one column**;
2. the **tallest** placeholder of a column is its body;
3. a **shorter** placeholder above that body is the column's heading strip —
   at most 60% of the body's height, so a second body is never mistaken for a
   heading;
4. columns are returned left to right, as `(heading, body)` pairs, with
   `heading` None for a layout that reserves no strip (Two Content).

The builder then fills column by column, and every shortfall is a warning
rather than a silent loss:

| Template offers | What happens | Code | Severity |
|---|---|---|---|
| Two columns with heading strips | Each heading and body written as asked | — | — |
| Two columns, no heading strips | Heading becomes a bold first line of its column | `heading_inlined` | info |
| One content area | Both columns merged into it, in order, headings bolded | `columns_merged` | warning |
| No content placeholder | Nothing to write into | `column_dropped` | error |

A layout with *more* columns than the slide has leaves the extra ones as the
template drew them.

### Pictures in a picture placeholder

An `image` slide prefers the template's picture layout and fills its
`PICTURE` placeholder through `_fill_picture_placeholder()`, so the
template's own frame, crop and position apply; `body` and `caption` go into
that layout's body placeholder, the caption as a trailing italic paragraph.
Until this existed `ROLE_IMAGE_TEXT` was classified but never requested and
no builder ever filled a picture placeholder, so a template's photo layouts
were dead weight
([#120](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/120)).

The computed-rectangle path below is the fallback, taken when the template
has no picture layout, or when the layout it resolved to has a picture
placeholder but no body placeholder to hold this slide's text — text is
never dropped to keep the nicer frame. An unfilled picture placeholder is
removed either way, since an empty one shows its prompt as soon as anyone
opens the deck.

Chart and image slides with a `body` split the area, chart or picture on the
left and bullets on the right. Both go through `_add_bulleted_textbox()`,
which fills the box and then calls `placeholder_style.apply_list_style()`: a
text box is not a placeholder, so it inherits from the presentation's default
text style, which has no bullet glyphs — the same markdown that bulleted
correctly in a content placeholder came out as plain lines beside a picture
([#123](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/123)).
Each paragraph takes the `marL`, `indent` and bullet definition of its own
level from the master's `<p:bodyStyle>`; size and spacing stay as the builder
set them, because that style's 28pt first level is meant for a full-width
placeholder. The copy keeps `<a:pPr>`'s children in schema order — bullet
properties after `lnSpc`/`spcBef` and before `defRPr` — since PowerPoint
reports a file whose paragraph properties are out of order as damaged. KPI and timeline slides are drawn from
autoshapes and text boxes because python-pptx cannot create SmartArt. A
timeline reserves its detail band before sizing the chevrons so captions can
never run off a short content box.

### One way to write text

Every string a caller supplies goes through `inline_formatting.write_text()`,
which renders the inline grammar when the text carries any and assigns it
plainly when it does not. The tool has always told the model that inline
markdown and links work in any text field; ten of them — titles, subtitles,
quote attributions, KPI figures, labels and deltas, timeline labels and
details, image captions, two-column headings — assigned `paragraph.text`
directly and printed `**bold**` as four asterisks. A new text field is
written with `write_text()`, not with `paragraph.text`, and
`tests/test_pptx_inline_everywhere.py` parametrises over every field there
is so a new one that forgets is a failing test.

Chart text goes through it too, with `hyperlinks=False`: a hyperlink needs a
relationship in the part that owns the run, and python-pptx cannot resolve a
chart title's part — attempting one raises `'ChartTitle' object has no
attribute 'part'`. PowerPoint does not follow a link inside chart text
anyway, so the label is rendered and the target dropped, which the tool
description says out loud.

The style arguments are the paragraph's defaults: what is left out stays
inherited, which is what a title filling a placeholder relies on. For the
same reason `draw_title_box()` puts the template's title style in the
paragraph's `<a:defRPr>` rather than on each run — stamping the runs would
overwrite the bold a caller asked for and drop a link's `<a:hlinkClick>`.

### Table formatting

`_create_styled_table()` takes three optional extras beyond alignment, the
header colour, zebra shading and a font size
([#124](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/124)):

| Extra | How it is applied |
|-------|-------------------|
| `column_widths` | relative weights normalised to the table's own width, so `[3, 1, 1]` means 60/20/20 whatever rectangle the layout gives it; the last column takes the rounding remainder so the columns still add up |
| `cell_fills` | applied after the header and zebra passes, so an explicit fill wins; a fill with no column covers the row |
| `merges` | applied **before any text is written** — `Cell.merge()` concatenates the text of the cells it joins, so merging afterwards repeats every cell of the block inside the one that survives. Cells reported as `is_spanned` are then skipped |

The builder checks all three against the real table first
(`_table_widths()`, `_table_fills()`, `_table_merges()`) and warns rather than
raising: widths that do not match the column count are dropped whole, since
guessing which column a short list meant is worse than leaving them equal;
fills and merges outside the table are skipped individually. Overlapping
merges are caught there too, because `Cell.merge()` raises part-way through a
block and would take the whole deck with it — the first block to claim a cell
keeps it.

### Fit estimation

python-pptx cannot lay text out, so `estimate_text_fill()` counts the lines
itself and multiplies by a line height. Lines are counted by **measuring**:
`text_metrics.measure_lines()` wraps each bullet against a real font file
through Pillow, which python-pptx already depends on, so "WWW WWW" and
"iii iii" no longer count the same
([#125](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/125)).
The result drives two things: `apply_autofit()` writes
`<a:normAutofit fontScale=…>` so viewers that do not recompute autofit still
shrink the text, and a warning is returned when the estimate passes the
shrink floor.

**Which size gets measured** matters as much as which face. The estimate used
to assume `DEFAULT_BODY_FONT_SIZE` (18pt) for every template, and both
templates this server ships set **28pt** in the master's `<p:bodyStyle>` — so
every estimate was low by the square of the ratio, about 2.4x. Text needing
1.9x its placeholder measured as 0.86x: `apply_autofit()` was handed no scale,
wrote a bare `<a:normAutofit/>`, and PowerPoint rendered the text at full size
straight off the bottom of the slide, because it only recomputes autofit when
someone clicks into the box. The overflow warning is derived from the same
number, so the caller was not told either
([#194](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/194)).

`read_body_font_size()` resolves the real size the way PowerPoint inherits it,
nearest first: the shape's own `<a:lstStyle>`, the layout placeholder it came
from, the master's body placeholder, the master's `<p:bodyStyle>`, then the
presentation's `<p:defaultTextStyle>`. `_fit_text()` uses it per placeholder;
the boxes the builder draws itself use `read_master_body_font_size()` through
`_fit_scale()`, because `apply_list_style()` gives them the master's body
style. `DEFAULT_BODY_FONT_SIZE` remains only as the last resort for a template
that states no size anywhere.

Which face gets measured, best first: the deck's own typeface
(`theme_body_typeface()` reads the theme's minor latin font), then a
**metric-compatible** substitute — Carlito for Calibri, Liberation Sans or
Arimo for Arial, Liberation Serif or Tinos for Times New Roman, Caladea for
Cambria, Gelasio for Georgia — whose advances are identical by design, so
measuring one measures the real thing; then any installed sans face, which is
a real measurement of the wrong font; and if no font file can be loaded at
all, `measure_lines()` returns None and the old `AVG_CHAR_WIDTH_RATIO`
arithmetic runs instead. The runtime image installs a package per family that
Alpine packages (`font-carlito`, `font-liberation`) — the shipped templates
set Aptos, which has no free metric-compatible clone, so the generic tier is
the usual one anyway. **Alpine packages neither Caladea nor Gelasio**, so in
the container Cambria and Georgia fall through to that generic sans; the two
stay in the table because a Debian host installing `fonts-crosextra-caladea`
and `fonts-gelasio` does measure them exactly.

Two tests keep this honest. `test_every_metric_compatible_face_is_installed`
reads the Dockerfile and fails if the table promises a face the image could
carry but does not, with the unpackaged families pinned by name so losing one
more is a deliberate edit. `test_every_font_package_exists_in_alpine`
(`network`) checks every package the Dockerfile names against Alpine's index:
a name Alpine does not have fails the image *build*, not a test, and that is
how v4.0-beta.4 came to publish no image at all.

Tables are sized by row count instead, through `fit_table_font_size()`, and
warn when they still would not fit at the minimum size.

### Sections, footers, language

Outline-pane sections have no python-pptx API. `_apply_sections()` writes
the `p14:sectionLst` extension directly, grouping the real slide-id list
under each `section` slide's title and putting anything before the first
section in "Default Section". Every slide must belong to a section once a
list exists or PowerPoint reports the file as damaged, which is why the
grouping is computed over slide ids rather than models.

Footer text and slide numbers are applied by cloning the layout's footer and
slide-number placeholders onto each slide with fresh shape ids. Layouts
without a footer placeholder are counted and reported in one deck-level
warning. `_apply_language()` stamps `lang` on every run, table cell and notes
frame so the deck is proofed in the right language. It walks shapes, so text
that lives inside a chart part — its title, axis titles and category labels —
is not reached and keeps the viewer's own proofing language; the tool
description says so rather than promising more than it does.

### The blank slide

`blank` is the escape hatch: positioned `text`, `image` and `shape` elements
with lengths in inches or as a percentage of the slide, resolved against the
real slide size so a percentage means the same on 4:3 and 16:9. An element
that starts off the slide is skipped; one that runs past the edge is shrunk.
Both are reported.

## Extension points

| To add… | Edit | Notes |
|---------|------|-------|
| A new slide type | `schema.py` (model, `_SLIDE_MODELS`, `AnySlide`), `layouts.SLIDE_TYPE_ROLE`, `slide_builder._build_slides()` (builder), `_build_<type>()` | Use `_content_slide()` if it draws its own shapes, `_new_slide()` + `_apply_title()` otherwise. Report anything dropped through `_warn(index, code, message)` |
| A new warning | a code in `warnings.py` and its severity in `WARNING_SEVERITY` | `test_every_code_has_a_severity` fails on a code without one. Severity is about the deck: `error` if the content is not in the file, `warning` if it is there but altered, `info` for a substitution |
| A new field on an existing type | the model in `schema.py`, then the builder | The flat schema regenerates itself; describe the field, since the description is what the model reads |
| A new chart type | `chart_utils.CHART_TYPE_MAP` and the `chart_type` literal in `ChartSlide` | Only category charts; an XY variant needs its own data path like scatter |
| A new layout role | `layouts.py` (`ROLE_*`, `ROLES`, `ROLE_FALLBACK_INDEX`, `classify_layout()`) | Keep classification conservative; returning `None` is better than a confident wrong pick |
| A new template default | `templates._spec_from_yaml()` passes `defaults` through; read it in the builder with `_setting()` | Coerce booleans and numbers, since YAML bypasses the slide schema |
| A new inline span | `inline_markdown.py` (grammar) and `inline_formatting._parse_segment()` | Shared with Word; gate it with a `build_inline_pattern()` flag if only one renderer can draw it |

## Invariants and gotchas

- **Never index `slide_layouts[N]` in a builder.** Go through `_new_slide()`.
  Positional lookup is what silently mis-laid decks on reordered templates.
- **Keep the published schema inside the keyword subset `_simplify()`
  allows.** No `oneOf`, `$ref` or `discriminator` may reach the client.
- **Report, do not swallow.** Anything the builder drops or substitutes goes
  through `_warn()` (per slide) or `self.warnings.append()` (per deck). An
  image that fails becomes a placeholder box **and** a warning.
- **A builder exception fails the deck.** This is intentional: a half-built
  deck with a success response is worse than an error naming the slide.
- **Never measure text against a hardcoded point size.** Ask
  `read_body_font_size()` what the template actually renders at; the estimate
  is wrong by the *square* of any error, and it is the same number the
  overflow warning is derived from.
- **Never match a content placeholder by `idx`.** Use `content_columns()`
  (columns) or `_content_placeholders()` (single body). `idx` numbering is a
  PowerPoint convention a customer template need not follow, and assuming it
  silently dropped content ([#194](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/194)).
  Footers and slide numbers are the remaining exception: they are still read
  from `idx` 11 and 12.
- **Template defaults bypass the schema.** Coerce and clamp them in the
  builder, as `_coerce_bool()` and `_coerce_font_size()` do.
- **The registry cache is keyed by file mtimes.** A test that writes a
  template and reads it back within the same clock tick can see a stale
  registry; call `clear_cache()`.

## Tests

| File | Covers |
|------|--------|
| `tests/test_pptx_creation.py` | End-to-end deck building through `PowerpointPresentation` and `.save()` |
| `tests/test_pptx_schema.py` | The typed schema, legacy migration, error messages |
| `tests/test_pptx_client_compat.py` | The flat published schema and the string-slide shim |
| `tests/test_pptx_robustness.py` | Input hardening: numeric cells, bad colours, odd levels |
| `tests/test_pptx_slide_types.py` | KPI, timeline, agenda, closing, hyperlinks, table markdown |
| `tests/test_pptx_blank.py` | Positioned elements, clamping and skipping |
| `tests/test_pptx_title_fallback.py` | A title on a layout with no title placeholder: geometry, style, the untitled-template warning |
| `tests/test_pptx_content_area.py` | Drawing on a layout with no body placeholder: the template's content rectangle, and how it is reported |
| `tests/test_pptx_picture_layout.py` | Image slides on a picture layout: role choice, the filled placeholder, and the fallbacks |
| `tests/test_pptx_warnings.py` | Warning records: codes, severities, the deck-wide case, and the tool boundary |
| `tests/test_pptx_bullet_glyphs.py` | Bullets in a text box: the master's glyphs and indents, and the order of `<a:pPr>` |
| `tests/test_pptx_table_formatting.py` | Column widths, cell and row fills, merged blocks, and what happens when they do not fit the table |
| `tests/test_pptx_text_metrics.py` | Measured line counts, wrapping, face selection, the arithmetic fallback, and that the image's font packages exist and match the table |
| `tests/test_pptx_theme_colors.py` | Drawn text, chart text and table fills taking the template's colours rather than literals |
| `tests/test_pptx_unused_placeholders.py` | That no generated slide keeps an empty "Click to add text" box, and that filled ones survive |
| `tests/test_pptx_body_font_size.py` | Reading the template's real body size, and the shrink factor and overflow warning that follow from it |
| `tests/test_pptx_two_columns.py` | Columns matched by geometry on renumbered templates, and the warning each degraded path owes |
| `tests/test_pptx_sections.py` | Outline-pane sections |
| `tests/test_pptx_templates.py` | Registry loading, `.potx`, layout classification and resolution, defaults |
| `tests/test_admin_pptx.py` | Admin UI support for PowerPoint templates |
| `tests/test_image_ssrf.py` | The image SSRF guard |

Tests instantiate `PowerpointPresentation` directly and call `.save()`,
bypassing upload entirely, then reopen the bytes with `Presentation()` to
assert on shapes and XML. See [`../testing.md`](../testing.md).

## Known limitations

- **Fit estimation measures the shapes, not the layout.** Text is measured
  against a real font file at the size the template states, but PowerPoint's
  own line breaking, kerning and autofit are not reproduced, and the deck's
  face may only be approximated by a substitute. The shrink factor remains a
  hint PowerPoint recomputes on first edit. Every bullet is measured at the
  level-1 size, so a slide of deeply indented bullets is estimated high.
- **Placeholder `idx` conventions are assumed for footers and slide numbers**
  (`idx` 11 and 12). Bodies and columns are matched by geometry instead.
- **No SmartArt.** KPI and timeline slides are built from autoshapes.
- **Category charts only in `chart`.** Scatter is its own slide type because
  it needs a different data object; bubble and combo charts are not offered.
