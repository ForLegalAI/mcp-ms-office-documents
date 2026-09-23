# Word tool (`docx_tools`)

Converts Markdown to a `.docx` file with python-docx. This page explains how
the package does it. For **what** the tool accepts, see the user reference:
[`../../markdown-reference.md`](../../markdown-reference.md). For the request
path around the build step, see [`../architecture.md`](../architecture.md).

## Entry points

| Name | Where | Used by |
|------|-------|---------|
| `create_word_from_markdown` | MCP tool declared in `main.py` | MCP clients |
| `_markdown_to_word_buffer()` | `docx_tools/base_docx_tool.py` | `main.py`, via `run_blocking`. Returns `(BytesIO, warnings)` |
| `_markdown_to_doc()` | `docx_tools/base_docx_tool.py` | the buffer function; tests that want the `Document` object |
| `markdown_to_word()` | `docx_tools/base_docx_tool.py` | direct library use; builds and uploads synchronously |
| `process_markdown_content()` | `docx_tools/markdown_processor.py` | the base tool **and** the dynamic template tools |

The tool wrapper in `main.py` passes `markdown_content`, the three metadata
fields (`title`, `author`, `subject`), `header_text`, `footer_text` and
`include_toc` to the buffer function. `file_name` and `add_unique_prefix` go
to the upload step, not the builder. It gets back a buffer **and** a list of
warnings, and wraps them with `main._with_warnings()` — see
"[The warnings channel](#the-warnings-channel)".

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
  markdown_processor.process_markdown_content(doc, content, style_map=…, warnings=…)
  ├─ patterns.normalize_newlines()            literal "\n", CR and CRLF → real newline
  ├─ patterns.expand_br_to_block_breaks()     <br> before a list/heading/quote → real newline
  ├─ split into lines, then loop:
  │    blank runs        → n-1 spacer paragraphs for n ≥ 2 blanks; one blank is a separator
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
  ├─ > quote            → _soft_break_run(strip_quote=True) → _add_quote()
  ├─ <!-- directive --> → collected, then applied to the next block
  ├─ other <!-- -->     → skipped
  └─ anything else      → _soft_break_run() → paragraph + inline parse
                          (a line ending in two spaces continues the paragraph,
                           stopping before any line that starts a block)
  ▼
  inline_formatting.parse_inline_formatting(text, paragraph)   every text span
  ├─ decode a fixed list of safe HTML entities
  ├─ <br>, CR, literal "\n" → "\n"; split, rstrip, one <w:br/> per break
  ├─ backslash escapes → private-use placeholders, restored after tokenising
  └─ inline_markdown.build_inline_pattern().split() → runs, hyperlinks
  ▼
_markdown_to_word_buffer(): doc.save(BytesIO) → (BytesIO, warnings)
```

Every stage after the template is loaded works on one `docx.Document` in
place. Nothing returns intermediate representations; the Markdown is walked
line by line and paragraphs are appended as they are recognised.

## Module map

| Module | Owns |
|--------|------|
| `base_docx_tool.py` | Template load, metadata, TOC, header/footer, then hands off to the processor. The three entry points |
| `markdown_processor.py` | Line walker (`process_markdown_content`), block dispatcher (`process_markdown_block`), soft breaks, code blocks, alignment blocks, comment directives, the running ordered-list count |
| `patterns.py` | Every compiled regex for block detection, plus `ordered_list_is_genuine()`, `normalize_newlines()`, `expand_br_to_block_breaks()`, `line_starts_block()`, `contains_block_markdown()` |
| `inline_formatting.py` | `parse_inline_formatting()`: entities, escapes, soft breaks, run creation, hyperlinks |
| `block_elements.py` | Tables, lists, images, horizontal lines, alignment detection |
| `numbering.py` | Ordered-list restart through fresh `<w:num>` instances; style-aware numbering resolution; indent re-assertion |
| `style_map.py` | `StyleMap` dataclass, config merging, `apply_style()` with fallback, the global map loaded from `config/docx_templates.yaml` |
| `warnings.py` | The Word warning codes and their severities; `channel()` builds the per-build collector |
| `document_features.py` | Template resolution, header/footer with PAGE/NUMPAGES fields, TOC field |
| `conditionals.py` | `{{#if}}`/`{{^if}}`/`{{/if}}` marker paragraphs for dynamic templates |
| `dynamic_docx_tools.py` | YAML-driven template tools, placeholder replacement across split runs, live registration. See [`../dynamic-templates.md`](../dynamic-templates.md) |

Three root modules are part of this pipeline:

- `inline_markdown.py` holds the emphasis grammar. Word asks for the full set
  of spans (`highlight=True, superscript=True, subscript=True`), so
  `patterns._INLINE_FORMAT_RE` is built from it. Change the grammar there, not
  in `patterns.py`; the PowerPoint renderer uses the same source.
- `image_utils.py` downloads and validates images, including the SSRF guard.
- `warning_channel.py` holds the severity vocabulary, the `DocumentWarning`
  record and the `WarningChannel` collector. `docx_tools/warnings.py` supplies
  the codes; see "[The warnings channel](#the-warnings-channel)".

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
   `process_markdown_content()` carries a cell holding the running count
   (`'next'`) and the numbering instance that produced it (`'num_id'`,
   `'ilvl'`, `'style'`); every top-level ordered list updates it, and
   `_continues_ordered_run()` lets a later `3.` start a list even when it is
   not locally genuine, provided it equals the running count exactly. Only a
   list starting at `1.` re-bases it.
3. **A restart is a fresh `<w:num>` instance.** python-docx cannot restart
   numbering, so `numbering.py` creates a new numbering instance with a
   `startOverride` whenever `1.` reappears at a level, and attaches it through
   a direct `<w:numPr>`. The abstract definition is resolved from the paragraph
   style actually applied (a mapped or directive style first, then the
   built-in `List Number` styles, then any decimal definition, then a
   synthesised one). Because a direct `numPr` lets the numbering level's
   indents override the style's, `apply_style_indent()` re-asserts the
   style's `w:ind` as direct formatting.
4. **A continuation is the *same* instance, never a new one** (#136).
   `process_list_items()` reads the instance out of the running cell and
   reuses its `numId`/`ilvl` when the first item equals the running count, so
   the parts of an interrupted list are one Word list: Word computes the later
   numbers and recomputes them when an earlier item is added or removed.
   Minting a second instance with `startOverride=3` — what this used to do —
   only *looks* consecutive and drifts on the first edit. Two cases still get
   their own instance: a restart at `1.` (rule 3), and a continuation whose
   level-0 style differs from the recorded one, which would otherwise inherit
   the earlier list's numeral format and indents. The reuse is tracked at
   level 0 only; nested levels always mint their own instance.

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
built-in names. `build_style_map()` merges the global `style_mapping` with a template's own
section, the template winning. `apply_style()` falls back to `Normal` (or the
document default for tables) and logs a warning when the named style is
missing from the template, so a wrong name never aborts a render.

The global mapping is whatever `template_registry.global_config()` resolves —
`config/docx_templates.yaml`'s top-level section under any
`config/docx_templates.d/_global.yaml` the admin UI has written (see
[dynamic-templates.md](../dynamic-templates.md#kind-wide-settings)). Its three
consumers reach it by different routes and must not diverge:

| Consumer | When it resolves the map | Picks up an edit |
|---|---|---|
| `markdown_to_word` (static) | `load_global_style_map()`, per document | immediately |
| a dynamic template tool | once, at registration | on re-registration |
| the admin preview | per render, via `AdminContext` | immediately |

`load_global_style_map()` used to cache for the life of the process, which
meant an edit in the admin UI left every subsequent document on the old
mapping until a restart
([#161](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/161)). It
now caches against the two config files' **mtime and size**, so a hand edit on
the volume is picked up on the next document.

Re-parsing unconditionally would also be correct, and was the first version of
this. It is not free: the master file is a few hundred lines of worked
examples, and a full resolution measures ~8 ms against ~0.02 ms for a cache
hit — more than the markdown render it would be paying for. Two `stat` calls
per document build is the trade.

The fingerprint alone is not enough, which is why
`invalidate_global_style_map()` exists and
`AdminContext.resync_docx_style_map()` calls it: two writes inside one
filesystem timestamp tick that leave the file the same length look identical,
and swapping one six-letter style name for another does exactly that. The
fingerprint is for edits made *outside* the server; a write the server made
itself it simply knows about.

A caller generating many documents should still build the map once and pass it
down, which is what every dynamic template tool already does.

Because the middle row bakes the map in, the admin UI re-registers every Word
template tool after saving a global mapping (`AdminContext.resync_docx_style_map`).
Without that, a save would move the static tool and leave the template tools
behind — a split with no symptom an admin could act on.

### Inline formatting

`parse_inline_formatting()` runs on every text span, including table cells,
headings and list items. Backslash escapes are swapped for private-use
characters before tokenising and restored on each run, so an escaped `*`
cannot open emphasis. Bold and italic recurse so nesting works; the other
spans create a run directly. Links become real `w:hyperlink` elements with
the relationship registered on the part; on failure the label is written as
plain text and a warning logged.

### One line-break model

Since #110 there is one rule for every entry point: the block layer decides
which newlines separate blocks, and **every newline that reaches the inline
layer is a soft break** (`<w:br/>`). `parse_inline_formatting()` folds
`<br>`, a CR and a literal `\n` into `\n`, splits on it and strips the
trailing-space marker from each segment. A line ending in two spaces is the
source spelling of the same break: `_soft_break_run()` in the block
dispatcher joins such lines with real newlines, for prose and for quotes,
and stops before any line for which `patterns.line_starts_block()` is true
so a list or table on the next line is never swallowed into the paragraph.
`line_starts_block()` is also the predicate behind `contains_block_markdown()`,
which routes template placeholder values.

`expand_br_to_block_breaks()` promotes a `<br>` to a real newline only when
a segment **after the first** begins a block, so `- item<br>second line`
keeps its break inside the item. It reads nothing but the line being split.
Do not give it the renderer's numbering state so that `Note<br>3. Third`
can resume an earlier list; that was tried and removed in #110 because the
copy of the running count drifted from `process_list_items()` in several
ways. A continuation written on its own line still works, since the
dispatcher has the live count there.

Inside a table cell, whose row is one physical line, `<br><br>` stands in
for the blank line and splits the cell into paragraphs
(`_BR_PARAGRAPH_RE` in `add_table_to_doc()`); a single `<br>` is a soft
break like everywhere else.

### The warnings channel

The renderer is deliberately forgiving: a block that raises is skipped, an
image that will not load becomes a bracketed placeholder, a style the template
does not define falls back to `Normal`. Each of those keeps a document that
would otherwise be lost — and each was, until
[#114](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/114),
invisible to the caller, who got a URL and a success message while the log
kept the reason.

A `WarningChannel` (`warning_channel.py`, codes in `docx_tools/warnings.py`)
is created in `_markdown_to_word_buffer()` and threaded down through
`process_markdown_content()` → `process_markdown_block()` → `add_table_to_doc()`,
`add_image_to_doc()`, `process_list_items()`, `apply_style()`. It is an
**argument, never module state**: builds run concurrently on `run_blocking`
worker threads, and a shared list would mix two callers' documents together.

Every site that drops or substitutes something calls
`warnings.add(code, message, line=…)` next to its existing log call. The line
is the 1-based source line of the caller's markdown, so the message points at
what to fix. `style_map.apply_style()` carries a line only where the caller
named the style — a `<!-- style: … -->` directive; a mapped style comes from
configuration and has no line to give.

| Code | Severity | Raised when |
|------|----------|-------------|
| `block_failed` | error | `process_markdown_block()` caught an exception; the block is missing |
| `table_failed` | error | `doc.add_table()` raised; the whole table is missing |
| `table_cell_failed` | error | One cell could not be written; it is empty |
| `image_failed` | error | The image would not load; the placeholder line stands in for it |
| `table_not_recognised` | warning | A line of pipe markup that is not a table at all; it fell through to the paragraph branch and was written as text |
| `table_separator_missing` | warning | A table parsed with no `\|---\|---\|` row directly under its first row; that row was taken as the header |
| `style_missing` | warning | The template has no such style; the fallback was used |
| `style_fallback_missing` | warning | The fallback style is missing too |
| `widths_invalid` | warning | A `<!-- widths -->` directive is not a list of numbers; it was ignored |
| `link_refused` | warning | A link target's scheme is not `http`, `https`, `mailto` or `tel`; `add_hyperlink()` kept the label as text. Reported from a pre-scan in `_markdown_to_word_buffer()`, since the inline renderer has no channel; `_inline_text()` leaves out whole-line images and fenced code, which make no link |

Two properties of the channel matter here. It **de-duplicates** identical
`(code, message, location)` entries, so a template without `List Number` warns
once rather than once per list item; and it **caps** the number of distinct
warnings it carries (`warning_channel.DEFAULT_LIMIT`), appending one
`warnings_truncated` entry rather than returning thousands.

Severity is about the document, which is why the same input can rate
differently here and in Excel: a lone line of pipe markup is a `warning` in
Word, where it falls through to the paragraph branch and is still in the file,
and an `error` (`table_incomplete`) in Excel, where a worksheet has nowhere to
put it. Where the outcome *is* the same in both tools the code is the same
too: a table whose first row was made the header because no separator row sat
under it reports `table_separator_missing` in either, at `warning` severity.
`parse_table()` returns that as its fourth value, and it means the separator
was in the one position markdown gives it meaning. Only row 1 of the run is
tested: a row of dashes anywhere else is data — `| - | - |` is how a caller
writes "not applicable in either column" — and matching by shape alone dropped
it from the table with nothing said. A run of nothing but separator rows has
no header for one of them to sit under, so it is not a table at all: it takes
the `table_not_recognised` path and the lines are written as prose. Excel
applies both rules too, reporting `table_incomplete` for the second because a
worksheet has nowhere to put the lines.

`_is_separator_line()` is the shape test both the guard and the loop use, so
they cannot disagree about what a separator row looks like; position is
decided by the caller of it. A separator row must also carry dashes in
*every* cell: the empty-cell exemption this once had made `all()` vacuously
true for a row of blank cells, so `|  |  |` passed as a separator and the
caller's blank row was swallowed with the table counted as well formed. Excel's
`_is_separator_row()` checks every cell; so does this.

`process_markdown_content()` takes `warnings=None`, which discards them. The
dynamic Word template tools render through that path and report nothing, as
before; only the static tool has a response shape to put them in.

### Headers, footers, TOC

`set_header_footer()` rewrites the first paragraph of every section's default
header or footer, plus the first-page and even-page variants where the
template uses them, keeping the paragraph's alignment and replacing only the
runs. `{page}` and `{pages}` become PAGE and NUMPAGES fields, and `<br>` or a
newline breaks the line; the text is otherwise plain, not Markdown.
`add_toc()` inserts a TOC field with `w:updateFields` set so Word refreshes
it on open.

## Extension points

| To add… | Edit | Notes |
|---------|------|-------|
| A new block element | `patterns.py` (regex), `markdown_processor.process_markdown_block()` (branch) | Place the branch in the right order; route output through `_collect()`; add it to `_BLOCK_PATTERNS` if the template path must detect it |
| A new inline span | `inline_markdown.py` (grammar), `inline_formatting._parse_formatting_segment()` (run creation) | Word and PowerPoint share the grammar; add the span to both renderers or gate it with a `build_inline_pattern()` flag |
| A new comment directive | `markdown_processor.process_markdown_block()` directive branch, and the branch that consumes it | Directives are `key` or `key: value`; keys are lower-cased |
| A new style-map key | `style_map.py` (`StyleMap` field, `_normalize()`), then the call site that applies it | Keep the field name in sync with the YAML key the admin UI writes |
| A new document-level feature | `document_features.py`, wired in `base_docx_tool._markdown_to_doc()` | Also add a tool parameter in `main.py` and pass it through the buffer function |
| A new warning | `docx_tools/warnings.py` (code + severity), then `warnings.add(...)` at the site | `tests/test_docx_warnings.py::test_every_code_has_a_severity` fails if the table is missed; the message names what the caller should change |

## Invariants and gotchas

- **A failing block is skipped, not fatal.** `process_markdown_block()` catches
  every exception, logs it at ERROR with the line number, and advances one
  line. The document is still produced, minus that block. Keep this when
  editing: a rendering bug must not cost the user the whole document.
- **Anything skipped or substituted goes on the warnings channel**
  ([#114](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/114)).
  A log line alone is not enough: the caller gets a success response and never
  reads the server log. A new forgiving branch needs a `warnings.add()` beside
  its `logger` call — see "The warnings channel".
- **A failing image becomes text.** `add_image_to_doc()` writes a paragraph
  reading `[Image could not be loaded: <url>]`, logs a warning and reports
  `image_failed`.
- **Never hold a `StyleMap` in module state.** See style mapping above.
- **`normalize_newlines()` and `expand_br_to_block_breaks()` are
  idempotent.** The template path calls them for routing and the processor
  calls them again. Keep both idempotent if you change them.
- **Do not reintroduce the old `'  \n'` soft-break marker** in the inline
  layer, and do not move `_soft_break_run()` out of `process_markdown_block()`:
  it sits there so prose, quotes, alignment blocks, the base tool and template
  placeholders all get the same behaviour. See "One line-break model".
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
| `tests/test_docx_list_indentation.py`, `test_docx_list_restart.py`, `test_docx_list_continuation_and_br.py`, `test_docx_ordered_list_date.py`, `test_docx_list_infinite_loop_regression.py` | Lists: nesting, restarts, running count, continuation on one instance, date disambiguation, forward-progress guard |
| `tests/test_docx_style_map.py`, `test_docx_style_numbering.py`, `test_docx_style_tag.py` | Style mapping, style-aware numbering, the `style` directive |
| `tests/test_docx_escaped_newlines.py` | Literal `\n` and backslash escapes |
| `tests/test_docx_soft_breaks.py` | The line-break model: `<br>`, trailing spaces, CR, runs stopping before blocks, quotes, cells, headers |
| `tests/test_docx_templates.py`, `test_docx_placeholder_formatting.py`, `test_docx_conditionals.py` | Dynamic templates: placeholder replacement across runs, formatting preservation, conditionals |
| `tests/test_docx_warnings.py` | The warnings channel: every code, the source line, de-duplication, the tool-boundary response shape |
| `tests/test_inline_markdown.py` | The shared inline grammar |
| `tests/test_warning_channel.py` | The shared record and collector the channel is built on |

The usual pattern is to call `_markdown_to_doc()` directly and inspect the
returned `Document`, which avoids the upload step entirely. See
[`../testing.md`](../testing.md).

## Known limitations

The first three are tracked in [#115](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/115).

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
- **Continuation reuse is top-level only.** The running cell tracks one
  numbering instance, for level 0. A nested list resumed after interposed
  content still starts a fresh instance, so its numbers are frozen at
  generation time the way top-level ones were before #136.
