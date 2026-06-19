# DOCX Tooling — Overview, Issues & Improvement Opportunities

> Comprehensive review of the Word (`.docx`) generation path: both the **basic**
> `create_word_from_markdown` tool and the **dynamic** YAML-driven template tools.
> Findings are evidence-based (file:line references) and prioritised. Suggested
> improvements deliberately avoid adding significant new syntax.
>
> **Date:** 2026-06-19

---

## 1. Architecture

Two entry points share one Markdown renderer.

```
                       ┌─────────────────────────────────────────────┐
                       │  Markdown renderer (the shared core)         │
 create_word_from_     │  markdown_processor.process_markdown_content │
   markdown  ─────────▶│    └─ process_markdown_block                 │
 (main.py)             │         ├─ inline_formatting.parse_inline_…  │
   │                   │         ├─ block_elements (tables/lists/img…) │
   │ base_docx_tool    │         ├─ numbering (list restart)          │
   │  ._markdown_to_doc│         └─ style_map (style resolution)      │
   │                   └─────────────────────────────────────────────┘
   │                                     ▲
 dynamic_docx_tools  ───────────────────┘
 (per-template MCP tools, {{placeholder}} substitution)
```

### Module map (`docx_tools/`)

| Module | Responsibility |
|--------|----------------|
| `base_docx_tool.py` | Basic tool: load template, set metadata/header/footer/TOC, drive rendering, upload |
| `markdown_processor.py` | Block dispatcher: empty-line spacing, soft breaks, `process_markdown_block` |
| `block_elements.py` | Tables, lists, images, alignment blocks, horizontal lines |
| `inline_formatting.py` | Run-level markdown (bold/italic/links/super/sub/…), HTML-entity decode, hyperlinks |
| `patterns.py` | All compiled regexes + block detection |
| `document_features.py` | Template loading, header/footer (`{page}`/`{pages}`), TOC field |
| `numbering.py` | Numbered-list restart via python-docx oxml helpers |
| `style_map.py` | `StyleMap`, config merge, `apply_style`, style-tag helper |
| `dynamic_docx_tools.py` | YAML → MCP tools, `{{placeholder}}` substitution |
| `template_utils.py` (root) | Template file resolution (custom → default dirs) |

### Key design facts

- **Hand-rolled parser.** No markdown library — every construct is a regex in
  `patterns.py`. Parsing is line-oriented; blocks are separated by blank lines.
- **Styles drive appearance.** Numbering and look come from the template's named
  styles (`List Number`, `Quote`, `Table Grid`, `Heading N`); `style_map.py` lets
  those names be remapped.
- **One renderer, two callers.** The dynamic path renders placeholder values through
  the same `process_markdown_content`, but block content (lists/headings) is only
  expanded for **body** placeholders — inside table cells, headers and footers,
  placeholders get inline formatting only (`dynamic_docx_tools.py:315,367`).

---

## 2. Feature inventory

**Block:** headings 1–6 · ordered/unordered lists (3-space nesting) with auto-restart ·
tables (column alignment, `<!-- borderless -->`, `<!-- widths: … -->`, `<br>` multi-paragraph
cells) · block quotes · images (downloaded from URL) · page breaks (`---`) · horizontal
lines (`***`) · text alignment (`<center>`, `<div align>`) · `<!-- style: Name -->` tag.

**Inline:** bold/italic/bold-italic · strikethrough · underline · highlight · `code` ·
super/subscript · links · backslash escapes · safe HTML-entity decoding.

**Document:** title/author/subject metadata · headers/footers with `{page}`/`{pages}` ·
auto-updating Table of Contents · custom + default template resolution · global &
per-template style mapping.

---

## 3. Issues & limitations

Ordered by impact. Severity = correctness/security impact × likelihood of being hit.

### 🔴 High

**3.1 Fenced code blocks are unsupported and actively corrupted.** _(✅ resolved — see 4.1)_
There is no ` ``` ` handling, so a fenced block is parsed line-by-line as ordinary
markdown. Verified output for a Python block:

| Input line | Rendered as |
|------------|-------------|
| ` ```python ` | literal paragraph `` ```python `` |
| `    # comment` | **Heading 1** "comment" |
| `    return 1` | paragraph |
| ` ``` ` | literal `` ` `` |

Any technical document with code (or markdown examples) is mangled — comments become
headings, `-`/`1.` lines become lists. *Root cause:* no fenced-block branch in
`process_markdown_block` (`markdown_processor.py`). **High value to fix.**

### 🟠 Medium

**3.2 List nesting only works at 3 (or 4) spaces; 2-space indent silently fails.** _(✅ resolved — see 4.2)_
`current_level = indent // 3` (`block_elements.py:198`, and the look-ahead at `:206`)
hardcodes a 3-space unit. Verified:

| Indent of child | Result |
|-----------------|--------|
| 2 spaces | **not nested** (stays top level) |
| 3 spaces | nested → `List … 2` ✓ |
| 4 spaces | nested → `List … 2` (works by luck: 4//3=1) |

2-space and 4-space are the *common* markdown conventions; 2-space nesting is lost
with no warning. Tabs are also unhandled (counted as 1 character).

**3.3 Image fetch has no SSRF protection.**
`validate_url` (`pptx_tools/image_utils.py`) only checks the scheme is http/https and a
host is present — it does not block loopback/link-local/private addresses. An
`![x](http://169.254.169.254/…)` or `http://localhost:…` triggers a server-side
request to internal infrastructure (blind SSRF; content-type/size limits constrain the
*response* but not the request). Shared with the PowerPoint tool, so one fix covers both.

**3.4 Duplicated rendering logic (maintainability).** _(✅ resolved — see 4.4)_
Heading and quote rendering exist in two places: the soft-break path in
`process_markdown_content` (`markdown_processor.py:72–87`) and `process_markdown_block`.
The recent style-mapping change had to be applied in both; future block changes risk
drifting. Same construct, two code paths.

### 🟡 Low

**3.5 Table header row is not visually distinguished.** Tables get `Table Grid` only —
no bold/shaded header. (Contrast the Excel tool, which bolds + fills the header.)

**3.6 Inconsistent directive mechanisms.** _(✅ resolved — see 4.6)_ `borderless`/`widths` use a *look-back*
(`markdown_processor.py:126–138`); the new `style` tag uses a *look-ahead*. Two patterns
for the same "comment directive above a block" idea.

**3.7 Errors are silently swallowed.** `process_markdown_block`'s broad `except`
(`markdown_processor.py`) logs and advances, dropping malformed content with no signal
in the output.

**3.8 Hardcoded table page width.** `col_widths` assumes `page_width = 6.5"`
(`block_elements.py`), wrong for A4 or non-1" margins — the image path already computes
this correctly from section margins and could be reused.

**3.9 Minor gaps:** no task lists (`- [ ]`), no definition lists, no footnotes; global
style map is cached for the process lifetime (config edits need a restart); `---` as a
page break diverges from standard markdown's thematic-rule meaning (documented, but LLMs
may emit `---` expecting a rule).

**3.10 Nested emphasis at a `***` boundary.** _(✅ resolved)_ `**bold with *italic***`
(and the mirror `*italic with **bold***`) left a literal `*` and mis-attributed the rest of
the line, because the `**bold**` regex couldn't absorb a nested italic whose closing `*`
abuts the bold close. Fixed in `patterns.py` by allowing a paired `\*[^*]+\*` inside the
bold alternative (lone `*` still kept literal via regex backtracking).

---

## 4. Improvement opportunities (no significant new syntax)

Most of these are **behaviour/quality fixes**, not syntax additions. The two feature
additions (4.1, 4.7) use *standard* Markdown, so prompts and existing docs still apply.

| # | Improvement | Addresses | Effort | Syntax impact |
|---|-------------|-----------|:------:|---------------|
| **4.1** ✅ | **Render fenced code blocks** — *Implemented.* ` ``` `/`~~~` blocks are consumed verbatim into monospace (Courier New) paragraphs, skipping markdown parsing inside; the paragraph style is mappable via `StyleMap.code`. | 3.1 | M | None (standard MD) |
| **4.2** ✅ | **Flexible list indentation** — *Implemented.* Nesting is now determined by relative indentation, so 2-, 3- and 4-space units and tabs all nest correctly. | 3.2 | M | None |
| **4.3** | **SSRF hardening** — in `validate_url`, resolve the host and reject loopback/link-local/RFC-1918/metadata ranges (allow-list scheme already present). Benefits Word + PPTX. | 3.3 | S | None |
| **4.4** ✅ | **De-duplicate soft-break rendering** — *Implemented.* Heading/quote creation is shared via `_add_heading`/`_add_quote`, used by both the block dispatcher and the soft-break path. | 3.4 | S | None |
| **4.5** | **Bold the table header row** — emphasise the first row (bold, optional light shading) for readability; opt-out via a `<!-- plain-header -->` directive if needed. | 3.5 | S | None (opt-out only) |
| **4.6** ✅ | **Unify directive handling** — *Implemented.* All comment directives (`borderless`/`widths`/`style`) are collected by one look-ahead mechanism and attached to the next block, so they now compose (e.g. style + borderless + widths on one table). | 3.6 | M | None |
| **4.7** | **Task-list rendering** — `- [ ]`/`- [x]` → `☐`/`☑` prefixes on a bullet. Tiny, common, standard syntax. | 3.9 | S | None (standard MD) |
| **4.8** | **Dynamic table width** — compute usable width from section margins (reuse the image-path calculation) instead of the 6.5" constant. | 3.8 | S | None |
| **4.9** | **Surface dropped content** — count/log swallowed block errors and optionally emit a visible marker, so silent data loss is detectable. | 3.7 | S | None |

### Suggested sequencing

1. **4.3 (SSRF)** — security, small, shared benefit. Do first.
2. **4.1 (code blocks)** — highest correctness value; common real-world content.
3. **4.2 (indentation)** — removes a silent-failure foot-gun.
4. **4.4 / 4.6 (dedup + unify directives)** — pay down maintainability before more features land.
5. **4.5 / 4.7 / 4.8 / 4.9** — polish, any order.

> None of the above changes existing markup. 4.1 and 4.7 add support for *standard*
> Markdown the parser currently ignores, so documents that already use them simply start
> rendering correctly.
