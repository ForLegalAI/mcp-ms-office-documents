# Word tool (`docx_tools`)

Converts Markdown to a `.docx` file with python-docx. This page explains how
the package does it. For **what** the tool accepts, see the user reference:
[`../../markdown-reference.md`](../../markdown-reference.md). For the request
path around the build step, see [`../architecture.md`](../architecture.md).

## Entry points

| Name | Where | Used by |
|------|-------|---------|
| `create_word_from_markdown` | MCP tool declared in `main.py` | MCP clients |
| `_markdown_to_word_buffer()` | `docx_tools/base_docx_tool.py` | `main.py`, via `run_blocking` |
| `_markdown_to_doc()` | `docx_tools/base_docx_tool.py` | the buffer function; tests that want the `Document` object |
| `markdown_to_word()` | `docx_tools/base_docx_tool.py` | direct library use; builds and uploads synchronously |
| `process_markdown_content()` | `docx_tools/markdown_processor.py` | the base tool **and** the dynamic template tools |

The tool wrapper in `main.py` passes `markdown_content`, the three metadata
fields (`title`, `author`, `subject`), `header_text`, `footer_text` and
`include_toc` to the buffer function. `file_name` and `add_unique_prefix` go
to the upload step, not the builder.

The dynamic Word template tools (`docx_tools/dynamic_docx_tools.py`) reuse the
same Markdown pipeline to render placeholder values. They are documented in
[`../dynamic-templates.md`](../dynamic-templates.md); this page covers the
pieces they share.

## Pipeline

```
markdown_content
  │
  ▼  base_docx_tool._markdown_to_doc()
  ├─ document_features.load_templates()      resolve custom/default .docx template
  ├─ Document(path)                           open it (blank document if none)
  ├─ core_properties                          title, author, subject
  ├─ document_features.add_toc()              optional TOC field, before the body
  ├─ document_features.set_header_footer()    {page}/{pages} become PAGE/NUMPAGES fields
  ├─ style_map.load_global_style_map()        style names from config/docx_templates.yaml
  ▼
  markdown_processor.process_markdown_content(doc, content, style_map=…)
  ├─ patterns.normalize_escaped_newlines()    literal "\n" typed as text → real newline
  ├─ patterns.expand_br_to_block_breaks()     <br> next to a list/heading → real newline
  ├─ split into lines, then loop:
  │    blank runs        → n-1 spacer paragraphs for n ≥ 2 blanks; one blank is a separator
  │    trailing "  "     → soft-break paragraph (lines joined with line breaks;
  │                        a leading # or > still makes it a heading or quote)
  │    everything else   → process_markdown_block()
  ▼
  markdown_processor.process_markdown_block(doc, lines, i, …)   one block per call
  ├─ heading            → style_map.add_mapped_heading() + inline parse
  ├─ fenced code        → _render_code_block(), verbatim, monospace
  ├─ table              → block_elements.parse_table() + add_table_to_doc()
  ├─ ---                → page break
  ├─ ***                → block_elements.add_horizontal_line()
  ├─ ![alt](url)        → block_elements.add_image_to_doc() (image_utils.download_image)
  ├─ <center>/<div align> → block_elements.detect_alignment(), inline or multi-line block
  ├─ ordered list       → block_elements.process_list_items() (+ numbering.py)
  ├─ unordered list     → block_elements.process_list_items()
  ├─ > quote            → _add_quote()
  ├─ <!-- directive --> → collected, then applied to the next block
  ├─ other <!-- -->     → skipped
  └─ anything else      → paragraph + inline parse
  ▼
  inline_formatting.parse_inline_formatting(text, paragraph)   every text span
  ├─ decode a fixed list of safe HTML entities
  ├─ <br> → soft-break marker
  ├─ backslash escapes → private-use placeholders, restored after tokenising
  └─ inline_markdown.build_inline_pattern().split() → runs, hyperlinks
  ▼
_markdown_to_word_buffer(): doc.save(BytesIO)
```

Every stage after the template is loaded works on one `docx.Document` in
place. Nothing returns intermediate representations; the Markdown is walked
line by line and paragraphs are appended as they are recognised.

## Module map

| Module | Owns |
|--------|------|
| `base_docx_tool.py` | Template load, metadata, TOC, header/footer, then hands off to the processor. The three entry points |
| `markdown_processor.py` | Line walker (`process_markdown_content`), block dispatcher (`process_markdown_block`), soft breaks, code blocks, alignment blocks, comment directives, the running ordered-list count |
| `patterns.py` | Every compiled regex for block detection, plus `ordered_list_is_genuine()`, `normalize_escaped_newlines()`, `expand_br_to_block_breaks()`, `contains_block_markdown()` |
| `inline_formatting.py` | `parse_inline_formatting()`: entities, escapes, soft breaks, run creation, hyperlinks |
| `block_elements.py` | Tables, lists, images, horizontal lines, alignment detection |
| `numbering.py` | Ordered-list restart through fresh `<w:num>` instances; style-aware numbering resolution; indent re-assertion |
| `style_map.py` | `StyleMap` dataclass, config merging, `apply_style()` with fallback, the global map loaded from `config/docx_templates.yaml` |
| `document_features.py` | Template resolution, header/footer with PAGE/NUMPAGES fields, TOC field |
| `conditionals.py` | `{{#if}}`/`{{^if}}`/`{{/if}}` marker paragraphs for dynamic templates |
| `dynamic_docx_tools.py` | YAML-driven template tools, placeholder replacement across split runs, live registration. See [`../dynamic-templates.md`](../dynamic-templates.md) |

Two root modules are part of this pipeline:

- `inline_markdown.py` holds the emphasis grammar. Word asks for the full set
  of spans (`highlight=True, superscript=True, subscript=True`), so
  `patterns._INLINE_FORMAT_RE` is built from it. Change the grammar there, not
  in `patterns.py`; the PowerPoint renderer uses the same source.
- `image_utils.py` downloads and validates images, including the SSRF guard.

## How the interesting parts work

### Block dispatch and the directive look-ahead

`process_markdown_block()` is a chain of `if` tests in a fixed order; the first
match wins. The block patterns are written so that most lines match exactly
one of them (`---` is not a list item because the list marker needs trailing
whitespace). Order matters where they do overlap: a recognised directive is
tested before the generic HTML-comment skip, an inline `<center>` line before
the plain-paragraph fallback, and a numbered line reaches the ordered-list
branch only when it is genuine, otherwise it falls through to a paragraph.
Fenced code consumes every line up to its closing fence, so nothing inside a
code block is ever dispatched to another branch.

Comment directives (`<!-- borderless -->`, `<!-- widths: 30 70 -->`,
`<!-- style: Name -->`) are the one look-ahead. The dispatcher collects every
consecutive directive line, skips blank lines, then recurses into itself for
the next block with the collected `directives` dict. Table directives are read
by the table branch; the `style` directive is applied after the block has been
rendered, to every element it produced. For a list, the style goes in through
a modified `StyleMap` instead, so nested levels keep their own styles and the
numbering definition is resolved from the directive style.

### Placeholder-safe rendering: `return_elements`

The dynamic template tools need to render Markdown *into the middle* of an
existing document. `process_markdown_content(..., return_elements=True)` makes
every branch detach what it appended from the body and return the raw XML
elements instead, so the caller can insert them after a placeholder paragraph.
The `_collect()` closure in the block dispatcher is this mechanism. When you
add a branch, route its output through `_collect()` or the template path will
lose it.

### Ordered lists: genuineness, restarts, and the running count

Three rules combine here, and each exists because of a real document.

1. **A numbered line is a list item only if it is genuine.**
   `patterns.ordered_list_is_genuine()` accepts a line that starts at `1.` or
   is followed by a sibling or nested item. A standalone `23. června 2026` is
   prose. A day-1 date has to be escaped (`1\. června`) because it is
   indistinguishable from a one-item list.
2. **Numbering continues across interposed content.** Legal filings put a
   heading, a quote, a table or an exhibit list between numbered paragraphs.
   `process_markdown_content()` carries a `{'next': N}` cell; every top-level
   ordered list updates it, and `_continues_ordered_run()` lets a later `3.`
   start a list even when it is not locally genuine, provided it equals the
   running count exactly. Only a list starting at `1.` re-bases it.
3. **A restart is a fresh `<w:num>` instance.** python-docx cannot restart
   numbering, so `numbering.py` creates a new numbering instance with a
   `startOverride` whenever `1.` reappears at a level, and attaches it through
   a direct `<w:numPr>`. The abstract definition is resolved from the paragraph
   style actually applied (a mapped or directive style first, then the
   built-in `List Number` styles, then any decimal definition, then a
   synthesised one). Because a direct `numPr` lets the numbering level's
   indents override the style's, `apply_style_indent()` re-asserts the
   style's `w:ind` as direct formatting.

### Nesting by relative indentation

`process_list_items()` treats the first item's indent as the level's base.
Any deeper line that is a list marker starts a child list one level down;
any consistent unit works. The style for a level comes from the
`list_number`/`list_bullet` tuples in the `StyleMap`, clamped to the last
entry. A forward-progress guard renders the line as a plain paragraph and
logs a warning if the loop would not advance; that guard is the fix for a
past infinite-loop hang and must stay.

### Style mapping

`StyleMap` is a frozen dataclass of Word style names. It is threaded through
every function as a parameter, never held globally, so concurrent conversions
on worker threads cannot share mutable state. The defaults reproduce Word's
built-in names. `build_style_map()` merges the global `style_mapping` from
`config/docx_templates.yaml` with a template's own section, the template
winning. `apply_style()` falls back to `Normal` (or the document default for
tables) and logs a warning when the named style is missing from the template,
so a wrong name never aborts a render. The global map is cached for the
process lifetime.

### Inline formatting

`parse_inline_formatting()` runs on every text span, including table cells,
headings and list items. Backslash escapes are swapped for private-use
characters before tokenising and restored on each run, so an escaped `*`
cannot open emphasis. Bold and italic recurse so nesting works; the other
spans create a run directly. Links become real `w:hyperlink` elements with
the relationship registered on the part; on failure the label is written as
plain text and a warning logged.

### Headers, footers, TOC

`set_header_footer()` rewrites the first paragraph of every section's default
header or footer, plus the first-page and even-page variants where the
template uses them, keeping the paragraph's alignment and replacing only the
runs. `{page}` and `{pages}` become PAGE and NUMPAGES fields. `add_toc()`
inserts a TOC field with `w:updateFields` set so Word refreshes it on open.

## Extension points

| To add… | Edit | Notes |
|---------|------|-------|
| A new block element | `patterns.py` (regex), `markdown_processor.process_markdown_block()` (branch) | Place the branch in the right order; route output through `_collect()`; add it to `_BLOCK_PATTERNS` if the template path must detect it |
| A new inline span | `inline_markdown.py` (grammar), `inline_formatting._parse_formatting_segment()` (run creation) | Word and PowerPoint share the grammar; add the span to both renderers or gate it with a `build_inline_pattern()` flag |
| A new comment directive | `markdown_processor.process_markdown_block()` directive branch, and the branch that consumes it | Directives are `key` or `key: value`; keys are lower-cased |
| A new style-map key | `style_map.py` (`StyleMap` field, `_normalize()`), then the call site that applies it | Keep the field name in sync with the YAML key the admin UI writes |
| A new document-level feature | `document_features.py`, wired in `base_docx_tool._markdown_to_doc()` | Also add a tool parameter in `main.py` and pass it through the buffer function |

## Invariants and gotchas

- **A failing block is skipped, not fatal.** `process_markdown_block()` catches
  every exception, logs it at ERROR with the line number, and advances one
  line. The document is still produced, minus that block. Keep this when
  editing: a rendering bug must not cost the user the whole document.
- **A failing image becomes text.** `add_image_to_doc()` writes a paragraph
  reading `[Image could not be loaded: <url>]` and logs a warning. The caller
  never sees an error, and the Word tool has no warnings channel to report it.
- **Never hold a `StyleMap` in module state.** See style mapping above.
- **`normalize_escaped_newlines()` and `expand_br_to_block_breaks()` are
  idempotent.** The template path calls them for routing and the processor
  calls them again. Keep both idempotent if you change them.
- **Do not add `Optional[...]` types to dynamic-tool arguments.** Optionality
  is expressed by the default alone. Several MCP clients drop the sibling
  description when they see `anyOf`. This rule lives in
  `dynamic_docx_tools._register_single_template()`.
- **Do not reintroduce a default for `add_unique_prefix` in the dynamic tool
  body.** It must reach `upload_file()` as `None` so the strategy default
  applies. See [`../architecture.md`](../architecture.md#upload-strategies-and-response-shape).
- **The escape placeholder range** is U+E000 upward. Documents that contain
  genuine private-use characters would collide; none have.

## Tests

| File | Covers |
|------|--------|
| `tests/test_docx_base.py` | End-to-end rendering through `_markdown_to_doc()`; writes inspection files to `tests/output/docx/` |
| `tests/test_docx_code_blocks.py` | Fenced code blocks |
| `tests/test_docx_list_indentation.py`, `test_docx_list_restart.py`, `test_docx_list_continuation_and_br.py`, `test_docx_ordered_list_date.py`, `test_docx_list_infinite_loop_regression.py` | Lists: nesting, restarts, running count, date disambiguation, forward-progress guard |
| `tests/test_docx_style_map.py`, `test_docx_style_numbering.py`, `test_docx_style_tag.py` | Style mapping, style-aware numbering, the `style` directive |
| `tests/test_docx_escaped_newlines.py` | Literal `\n` and backslash escapes |
| `tests/test_docx_templates.py`, `test_docx_placeholder_formatting.py`, `test_docx_conditionals.py` | Dynamic templates: placeholder replacement across runs, formatting preservation, conditionals |
| `tests/test_inline_markdown.py` | The shared inline grammar |

The usual pattern is to call `_markdown_to_doc()` directly and inspect the
returned `Document`, which avoids the upload step entirely. See
[`../testing.md`](../testing.md).

## Known limitations

- **Table widths assume US Letter with 1-inch margins.** `add_table_to_doc()`
  distributes `<!-- widths -->` over a fixed 6.5 inches, while
  `add_image_to_doc()` reads the real section width. Templates with other
  page sizes get proportionally wrong absolute widths.
- **The TOC heading is English and bypasses the style map.** `add_toc()` writes
  a literal "Table of Contents" with `doc.add_heading()`, so a template that
  remaps `heading_1` does not apply to it.
- **Conditionals work in the body only.** Markers inside table cells, headers
  and footers are not evaluated.
- **Block-level placeholder content is not supported inside table cells.**
  `_replace_placeholders_in_table()` passes `doc=None`, so a value containing
  a list or heading is rendered with inline formatting only.
- **Code block font is hard-coded** to Courier New unless `code` is mapped to
  a template style.
