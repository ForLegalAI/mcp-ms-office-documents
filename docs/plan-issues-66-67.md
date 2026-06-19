# Implementation Plan — Issues #66 & #67

> **Scope:** Numbered‑list restart (#67) and custom style mapping + style tag (#66)
> for the markdown → DOCX conversion path.
> **Date:** 2026-06-19

---

## 0. Background — how the DOCX pipeline works today

Markdown is parsed by **hand‑rolled regexes** (no markdown library) and rendered with
**python‑docx 1.2.0**. The relevant call chain:

```
base_docx_tool.markdown_to_word()
  └─ _markdown_to_doc()                       # base_docx_tool.py:24  (loads template, Document(path))
       └─ markdown_processor.process_markdown_content()
            └─ process_markdown_block()        # markdown_processor.py — block dispatcher
                 ├─ headings   → add_heading(level)          (markdown_processor.py:75/115)
                 ├─ blockquote → para.style = 'Quote'        (markdown_processor.py:197)
                 ├─ tables     → word_table.style='Table Grid'(block_elements.py:120)
                 └─ lists      → process_list_items()         (block_elements.py:161)
```

Two facts drive this plan:

1. **There is no numbering machinery.** A repo‑wide grep for `numId`, `abstractNumId`,
   `ilvl`, `numPr`, `numbering.xml`, `restart` returns **zero** source matches. Ordered
   lists rely entirely on the `List Number` / `List Number 2/3` **paragraph styles**, whose
   numbering definitions live inside the loaded template's `numbering.xml`. Because every
   ordered paragraph shares the same style → same `numId`, **separate lists continue one
   running count** instead of restarting. (block_elements.py:170‑191)

2. **Styles are hard‑coded string literals at each call site.** There is no name→style
   indirection layer (only `ALIGNMENT_MAP` exists, block_elements.py:23). The list capture
   pattern (`ORDERED_LIST_CAPTURE_PATTERN`, patterns.py:16) **discards the explicit number**,
   capturing only the item text.

---

## Issue #67 — Numbered list restart

> *Restart numbered list if `1` is used again on the same list level.*

### Problem statement

Two consecutive ordered lists in one document keep counting (…3, 4, 5…) because they share
the `List Number` style's single `numId`. We need each new list — signalled by an item
numbered `1` reappearing at a given level — to **restart at 1**.

### Design decision: per‑list numbering instance with `startOverride`, built via python‑docx's own oxml helpers (no raw XML)

The OOXML‑correct way to restart numbering is **not** to keep relying on the style's
`numId`, but to create a fresh numbering **instance** (`<w:num>`) that points at the same
`<w:abstractNum>` and overrides its start value. The paragraph then carries an explicit
`<w:numPr>` (ilvl + the new numId) in its `pPr`, overriding the style's numbering for that
run of items — exactly what Word emits for "Restart numbering at 1".

**There is no high‑level python‑docx API for this** — `Paragraph` exposes no numbering
control (verified against the installed 1.2.0: its public surface is `style`, `alignment`,
`runs`, `text`, `paragraph_format`, …). **However, we do *not* need to hand‑author XML
strings.** python‑docx 1.2.0 ships the exact oxml helper *methods* for this in
`docx/oxml/numbering.py`, so the implementation is a handful of method calls:

```python
numbering = doc.part.numbering_part.element        # CT_Numbering
num = numbering.add_num(abstract_id)               # new <w:num> -> existing abstract def
num.add_lvlOverride(level).add_startOverride(1)    # restart this level at 1
num_id = num.numId

numPr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
numPr.get_or_add_numId().val = num_id
numPr.get_or_add_ilvl().val  = level
```

`add_num`, `add_lvlOverride`, `add_startOverride` are real methods on python‑docx's
`CT_Numbering` / `CT_Num` / `CT_NumLvl` classes. The only manual read is resolving the
`abstractNumId` off the `List Number` style once. No `OxmlElement('w:…')` string‑building,
no XML templates.

**Options considered**

| Option | Verdict |
|--------|---------|
| **B. New `<w:num>` + `startOverride` via python‑docx oxml helpers (chosen)** | Real, editable Word lists; keeps the named style's appearance; ~15 lines of method calls, no raw XML. |
| A. Render literal "1." as paragraph text, reset a Python counter | Truly zero numbering code, trivial restart — but numbers are **static text**: Word won't auto‑renumber on edit, and `List Number` indentation must be replicated by hand. Kept as documented fallback only. |
| C. Keep style numbering, set `w:start` on the shared abstractNum | Mutates the shared definition → affects *every* list, can't restart selectively. ✗ |

> **Fallback (Option A):** if a target template genuinely lacks any usable numbering
> definition and we can't synthesize one, the system can drop to literal‑text numbering
> (counter reset per logical list) so output is still correct, at the cost of live Word
> list semantics. This path is the documented degraded mode, not the default.

### New module: `docx_tools/numbering.py`

Thin wrapper over python‑docx's oxml helpers so call sites stay clean (and so the fallback
is centralised):

- `get_abstract_num_id(doc, style_name) -> int | None`
  Read the style's `pPr/numPr/numId`, resolve that `<w:num>` to its `<w:abstractNumId>`.
- `ensure_decimal_abstract_num(doc) -> int`
  Fallback: if a style has no numbering (possible in a bare custom template), create a
  minimal decimal `<w:abstractNum>` (levels 0‑2, `%1.` / `%2.` / `%3.`) via the same helpers.
- `new_restarted_num(doc, abstract_num_id, level, start=1) -> int`
  `numbering.add_num(...)` + `add_lvlOverride(level).add_startOverride(start)`; return `numId`.
- `apply_numbering(paragraph, num_id, level)`
  `get_or_add_numPr()` then set `numId.val` / `ilvl.val` — no manual element construction.

Numbering part access: `doc.part.numbering_part.element` (the `<w:numbering>` root; create
the part if absent).

### Changes to `process_list_items` (block_elements.py:161)

1. **Capture the explicit number.** Add `ORDERED_LIST_CAPTURE_PATTERN = r'^(\d+)\.\s+(.+)'`
   (two groups) in patterns.py:16; use group(1)=number, group(2)=text. Keep the bullet
   pattern unchanged.
2. **Allocate a numId per logical list.** On entry to an ordered‑list call, and again
   whenever an item's explicit number is `1` *after* at least one item has already been
   emitted at this level, allocate a fresh restarted `numId` (start = the item's number,
   normally 1) and call `apply_numbering()` on that and subsequent items.
3. **Nested lists** each get their own `numId` (recursion already passes a distinct level),
   so a child list never bleeds into the parent count and restarts independently.
4. Keep applying the named style (`List Number*`) for **appearance**; `<w:numPr>` only
   overrides the *count*, not the look.

### Behaviour after fix

```markdown
1. Apple
2. Banana

1. Carrot      ← restarts at 1 (was 3 before)
2. Potato
```

### Risks / edge cases

- **Bare `Document()` fallback** (no template) — python‑docx's default part includes
  numbering; `ensure_decimal_abstract_num` covers templates whose `List Number` style lacks
  a `numId`.
- **`return_elements=True` path** (placeholder machinery) removes paragraphs from the body
  but the `<w:num>` definitions live in the numbering part and survive re‑insertion — verify
  in tests.
- **Lists that intentionally start at N** (e.g. `5.`) — honour it via `startOverride=5`
  rather than forcing 1; matches author intent and the issue's wording.

---

## Issue #66 — Custom style list (mapping + style tag)

> *For custom templates, allow mapping of custom styles to replace built‑in styles. In
> addition, add a style tag to allow use of a custom style outside of regular mapping.*

Two independent sub‑features.

### Part A — Style **mapping** (replace built‑in style names)

Today each call site hard‑codes `'List Number'`, `'Quote'`, `'Table Grid'`, `Heading N`,
etc. We introduce a single indirection layer.

**Design decision: a `StyleMap` dataclass threaded through the processors, with defaults
equal to today's literals.**

```python
# docx_tools/style_map.py
@dataclass(frozen=True)
class StyleMap:
    heading: tuple[str, ...] = ("Heading 1", ..., "Heading 6")
    list_number: tuple[str, ...] = ("List Number", "List Number 2", "List Number 3")
    list_bullet: tuple[str, ...] = ("List Bullet", "List Bullet 2", "List Bullet 3")
    quote: str = "Quote"
    table: str = "Table Grid"
    normal: str = "Normal"
DEFAULT_STYLE_MAP = StyleMap()
```

- Replace the literals at:
  - block_elements.py:170‑171 (list styles), :120 (table)
  - markdown_processor.py:75/115 (headings), :197 (quote)
- Thread `style_map: StyleMap = DEFAULT_STYLE_MAP` through
  `process_markdown_content → process_markdown_block → process_list_items /
  process_table / …`. This mirrors how `doc` and `return_elements` are already threaded —
  consistent and unit‑testable. (No global/module state.)

**Where the overrides come from** — `config/docx_templates.yaml`:

```yaml
# global default for all base-tool conversions
style_mapping:
  list_number: "Brand Numbers"
  heading_1:   "Brand Title"

templates:
  - name: formal_letter
    docx_path: letter_template.docx
    style_mapping:           # per-template override (wins over global)
      quote: "Letter Quote"
```

`dynamic_docx_tools.register_*` builds a `StyleMap` from `DEFAULT_STYLE_MAP` + global +
per‑template overrides and passes it into `process_markdown_content`. Unknown style names
fall back to the default with a logged warning (reuse the table‑style try/except pattern at
block_elements.py:121‑127, generalised into a `apply_style(obj, name, fallback)` helper).

**Options considered**

| Option | Verdict |
|--------|---------|
| **A. `StyleMap` param threaded through (chosen)** | Explicit, testable, no hidden state; matches existing signature style. |
| B. Module‑level/global current‑map | Simpler call sites but not thread‑safe across concurrent MCP tool calls. ✗ |
| C. Store map on the `Document` object | Hacky attribute; breaks the `doc=None` placeholder paths. ✗ |

### Part B — Style **tag** (ad‑hoc style outside the map)

**Design decision: reuse the existing HTML‑comment directive convention.** Tables already
support `<!-- borderless -->` and `<!-- widths: … -->` directives (markdown_processor.py:130‑134),
and lone `<!-- … -->` lines are already skipped (markdown_processor.py:202). We add:

```markdown
<!-- style: Callout -->
This paragraph is rendered with the "Callout" style.
```

The directive applies the named style to the **next block** (paragraph, list, heading, or
table) — same "directive precedes block" model the table directives already use.

Implementation:
- Add `STYLE_DIRECTIVE_PATTERN = re.compile(r'^<!--\s*style:\s*(.+?)\s*-->$')` (patterns.py).
- In `process_markdown_block`, when the directive matches, capture the style name and apply
  it to the paragraph(s) produced by the following block via the shared
  `apply_style(obj, name, fallback)` helper (KeyError → keep default + warn). For a list the
  style applies to every item paragraph; for an explicit style this **overrides** the
  `StyleMap` for that block only.

### Risks / edge cases

- **Unknown style name** (typo or missing in template) → python‑docx raises `KeyError`;
  the `apply_style` helper catches it, logs a warning, leaves the default style. Document
  generation never fails because of a bad style name.
- **Style tag inside placeholder/table‑cell content** — block styling is already disabled in
  cells/headers/footers (`doc=None`, dynamic_docx_tools.py:315/367); the directive there is
  simply skipped, consistent with current behaviour.
- **Interaction with #67** — a style tag on an ordered list changes *appearance only*; the
  injected `<w:numPr>` from #67 still controls the count.

---

## Testing strategy

New tests under `tests/`, following the existing
`test_docx_list_infinite_loop_regression.py` style (open generated docx, assert on
`paragraph.style.name` and XML):

**#67**
- Two consecutive ordered lists → second list's first item renders as "1" (assert a fresh
  `numId` with `startOverride=1` in the numbering part).
- A list starting at `5.` → `startOverride=5`.
- Nested ordered list restarts independently of parent.
- `return_elements=True` path preserves numbering after re‑insertion.
- Regression: existing infinite‑loop / forward‑progress tests still pass.

**#66**
- Global + per‑template `style_mapping` applied (assert `style.name`).
- Unknown mapped style → falls back + warning, no exception.
- `<!-- style: X -->` applies X to the next paragraph / list / table.
- Default behaviour unchanged when no mapping/tag present (all current tests green).

---

## Suggested sequencing (independent, shippable separately)

1. **#67** — self‑contained (new `numbering.py` + `process_list_items` + pattern). No public
   API change. Ship first.
2. **#66 Part A** (StyleMap threading + YAML wiring) — touches more signatures; defaults keep
   behaviour identical, so low risk.
3. **#66 Part B** (style tag) — small, builds on the `apply_style` helper introduced in A.

Each step is gated by its own tests and leaves `DEFAULT_STYLE_MAP` / no‑directive behaviour
byte‑for‑byte unchanged.

---

## Files touched (summary)

| File | #67 | #66 |
|------|-----|-----|
| `docx_tools/numbering.py` *(new)* | ✓ | |
| `docx_tools/style_map.py` *(new)* | | ✓ |
| `docx_tools/patterns.py` | number capture group | `STYLE_DIRECTIVE_PATTERN` |
| `docx_tools/block_elements.py` | restart logic in `process_list_items` | literals→`StyleMap`, `apply_style` helper |
| `docx_tools/markdown_processor.py` | thread `style_map` | thread `style_map`, style directive, literals→`StyleMap` |
| `docx_tools/base_docx_tool.py` | — | accept/forward `style_map` |
| `docx_tools/dynamic_docx_tools.py` | — | build `StyleMap` from YAML |
| `config/docx_templates.yaml` | — | document `style_mapping` + `<!-- style: -->` |
| `tests/…` | new restart tests | new mapping/tag tests |
