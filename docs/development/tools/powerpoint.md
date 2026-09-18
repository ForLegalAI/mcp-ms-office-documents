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

This is the one tool with a **warnings channel**. Everything the builder had
to work around, such as an image that would not load, text shrunk to fit, a
layout the template does not provide or a footer dropped, is collected on
`PowerpointPresentation.warnings` and returned alongside the file so the
calling model can correct its next call. `main.py` wraps the result in
`{"file", "slide_count", "warnings"}` when there is anything to report.

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
| `placeholder_style.py` | Reading placeholder geometry and inherited character style from a layout, and replaying a title on a plain text box (`read_title_style()`, `read_content_rect()`, `draw_title_box()`) |
| `templates.py` | `TemplateSpec`, the registry loaded from YAML with an mtime-fingerprint cache, `.potx` handling, `select_template()`, `validate_templates()` |
| `chart_utils.py` | Category charts from `CategoryChartData`, scatter from `XyChartData`, legend, title, data labels, axis titles |
| `inline_formatting.py` | Renders the shared inline grammar into python-pptx runs |
| `constants.py` | Aspect ratios, positional layout fallbacks, typography, autofit ratios, table colours |

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
historical filename slots, so an existing deployment is unchanged. The cache
key is a fingerprint of every watched file's modification time, including
each template file, so overwriting a template or dropping in a new one takes
effect without a restart. The aspect ratio is read from the file, never from
the config, and exactly one spec is marked default.

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
`comparison`, `image_text`, `title_only`, `blank`); `role_for_slide()` picks
`comparison` over `two_column` when either column has a heading, `blank`
for an untitled quote so no empty title placeholder is left behind, and
`image_text` for an image slide when `resolver.provides()` says the template
has such a layout — the one role that depends on the template rather than on
the slide alone.
Resolution order is the slide's own `layout` name, then the registry's
`layouts:` mapping, then detection, then the positional index with a warning,
and finally the last layout rather than an `IndexError` on a trimmed
template.

`LayoutResolver.describe()` is what `validate_templates()` reports and
`list_presentation_templates` returns: each layout with its index, placeholder
types, whether it is vertical, and the role it classified as. Roles are held
in a list parallel to the layouts rather than keyed by name, because two
layouts in one template may share a name. Together with `coverage()` and
`missing_roles()` this is the answer to "why did my slide come out on that
layout"
([#121](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/121)).

`classify_layout()` reads placeholder *types*, ignoring date, footer and
slide-number chrome, so it is language-independent. It is deliberately
conservative: vertical-text layouts and "Content with Caption" return `None`
and are left for an explicit name. The distinguishing tests are a picture
placeholder (image_text), a subtitle (title), no content (title_only), four
content placeholders (comparison), two (two_column), and for one, `BODY`
means section and `OBJECT` means content.

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
left and bullets on the right. KPI and timeline slides are drawn from
autoshapes and text boxes because python-pptx cannot create SmartArt. A
timeline reserves its detail band before sizing the chevrons so captions can
never run off a short content box.

### Fit estimation

python-pptx cannot lay text out, so `estimate_text_fill()` approximates lines
from a mean glyph width and line height, both constants. The result drives
two things: `apply_autofit()` writes `<a:normAutofit fontScale=…>` so viewers
that do not recompute autofit still shrink the text, and a warning is
returned when the estimate passes the shrink floor. Tables are sized by row
count instead, through `fit_table_font_size()`, and warn when they still
would not fit at the minimum size.

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
frame so the deck is proofed in the right language.

### The blank slide

`blank` is the escape hatch: positioned `text`, `image` and `shape` elements
with lengths in inches or as a percentage of the slide, resolved against the
real slide size so a percentage means the same on 4:3 and 16:9. An element
that starts off the slide is skipped; one that runs past the edge is shrunk.
Both are reported.

## Extension points

| To add… | Edit | Notes |
|---------|------|-------|
| A new slide type | `schema.py` (model, `_SLIDE_MODELS`, `AnySlide`), `layouts.SLIDE_TYPE_ROLE`, `slide_builder._build_slides()` (builder), `_build_<type>()` | Use `_content_slide()` if it draws its own shapes, `_new_slide()` + `_apply_title()` otherwise. Report anything dropped through `_warn()` |
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
- **Two-column slides depend on placeholder `idx` values 1–4** matching
  PowerPoint's Two Content and Comparison conventions, and footers on `idx`
  11 and 12. A custom template that renumbers them silently drops content.
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
| `tests/test_pptx_sections.py` | Outline-pane sections |
| `tests/test_pptx_templates.py` | Registry loading, `.potx`, layout classification and resolution, defaults |
| `tests/test_admin_pptx.py` | Admin UI support for PowerPoint templates |
| `tests/test_image_ssrf.py` | The image SSRF guard |

Tests instantiate `PowerpointPresentation` directly and call `.save()`,
bypassing upload entirely, then reopen the bytes with `Presentation()` to
assert on shapes and XML. See [`../testing.md`](../testing.md).

## Known limitations

- **Fit estimation is approximate.** No font metrics are consulted; the
  shrink factor is a hint PowerPoint recomputes on first edit, and the
  warning threshold is a rough estimate.
- **Placeholder `idx` conventions are assumed** for two-column bodies and for
  footers and slide numbers. See the invariants above.
- **No SmartArt.** KPI and timeline slides are built from autoshapes.
- **Category charts only in `chart`.** Scatter is its own slide type because
  it needs a different data object; bubble and combo charts are not offered.
- **`coerce_indent_level()` and `_add_image_from_url()` in `helpers.py` have
  no callers** outside tests. They are backwards-compatibility leftovers and
  candidates for removal ([#116](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/116)).
