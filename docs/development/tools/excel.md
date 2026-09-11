# Excel tool (`xlsx_tools`)

Converts Markdown tables to an `.xlsx` workbook with openpyxl, resolving a
table-relative formula syntax to real cell references on the way. This page
explains how the package does it. For **what** the tool accepts, see the user
reference: [`../../markdown-reference.md`](../../markdown-reference.md). For the
request path around the build step, see [`../architecture.md`](../architecture.md).

## Entry points

| Name | Where | Used by |
|------|-------|---------|
| `create_excel_from_markdown` | MCP tool declared in `main.py` | MCP clients |
| `_markdown_to_excel_buffer()` | `xlsx_tools/base_xlsx_tool.py` | `main.py`, via `run_blocking` |
| `_build_workbook()` | `xlsx_tools/base_xlsx_tool.py` | the buffer function; tests that want the `Workbook` object |
| `markdown_to_excel()` | `xlsx_tools/base_xlsx_tool.py` | direct library use; builds and uploads synchronously |

The tool takes `markdown_content` and `auto_filter`. `file_name` and
`add_unique_prefix` go to the upload step, not the builder. There are no
dynamic Excel template tools; the only template is an optional
`custom_xlsx_template.xlsx` that contributes named cell styles.

The tool description string in `main.py` is unusually long (about 3,000
characters) because it is the only place the calling model learns the formula
grammar. Any change to that grammar must touch the description, the user
reference and the tests together.

## Pipeline

```
markdown_content
  │
  ▼  base_xlsx_tool._build_workbook()
  ├─ reject empty input
  ▼
  parser.walk_markdown_lines(lines)                 ONE pass → list of events
  ├─ <!-- key: value -->   → pending directive (attaches to the next table)
  ├─ ## Sheet: Name        → SheetEvent (rename of the default sheet, or a new one)
  ├─ # … ######            → HeaderEvent at the current row, then +2 rows
  ├─ | … |                 → helpers.parse_table() → TableEvent at the current row,
  │                           then +len(rows)+2; directives attached; T-number assigned
  └─ anything else         → dropped (and clears pending directives)
  ▼
  parser.collect_table_positions(events)            {sheet: {"T1": header_row, …}}
  ▼
  Workbook(); styles.load_template_styles().register_into(wb)
  ▼
  for each event:
    SheetEvent   → rename the active sheet, or wb.create_sheet(); reset per-sheet positions
    HeaderEvent  → cell A<row> with a level font (HEADER_FONTS)
    TableEvent   → helpers.add_table_to_sheet(); then ws.freeze_panes if "freeze" directive
  ▼
  helpers.add_table_to_sheet(table_data, ws, start_row, positions, …)
  ├─ per cell:
  │    header row          → text as written, white-on-blue, centred
  │    typed column        → _apply_column_type() (text/bool/currency/number/date/percent)
  │    otherwise           → resolve_cell(): formula? percent? number? date? text
  │    formula             → adjust_formula_references() → _write_formula()
  │    inline **bold** etc → apply_cell_formatting(); border; alignment; number format
  ├─ column widths from content, clamped to 12–25
  ├─ <!-- styles --> pass  → styles.parse_styles_directive() + apply_style_spec()
  └─ auto_filter           → _ensure_unique_table_headers(); openpyxl Table object
  ▼
  reject a workbook with zero tables
  ▼
_markdown_to_excel_buffer(): wb.save(BytesIO); circular_refs.detect_circular_references() (log only)
```

The parser runs once and produces events that carry their Excel row numbers.
Both the position map used for cross-sheet formulas and the workbook builder
consume the same events, so the two can never disagree about where a table
sits. That single pass is the reason `parser.py` exists as a separate module;
an earlier version computed positions twice and the copies drifted.

## Module map

| Module | Owns |
|--------|------|
| `base_xlsx_tool.py` | The three entry points; the event loop that turns events into sheets, headers and tables; the zero-tables check; the post-save circular-reference warning |
| `parser.py` | `walk_markdown_lines()`, the event dataclasses, `collect_table_positions()`, sheet-name sanitising, the spacing constants |
| `helpers.py` | Everything cell-level: `parse_table()`, `resolve_cell()`, date and number detection, `adjust_formula_references()`, the `types` directive, formula-length guard, header uniqueness, `add_table_to_sheet()` |
| `styles.py` | The `styles` directive: colour parsing, target and range expansion, named styles loaded from `custom_xlsx_template.xlsx`, and the non-destructive named-style application |
| `circular_refs.py` | Static dependency graph over formula strings and DFS cycle detection; diagnostic only |

## How the interesting parts work

### Row bookkeeping

Rows are assigned by the parser, not by the writer. A header occupies one row
and advances the cursor by two. A table occupies one row per line including
the header and advances by its length plus two. `add_table_to_sheet()` returns
the next free row for API symmetry, but `_build_workbook()` ignores it and
trusts the event's `start_row`. The `TABLE_BOTTOM_SPACING` constant therefore
exists in both `parser.py` and `helpers.py` and the two must stay equal; the
parser cannot import it from helpers without a cycle.

### Sheets and table numbering

`## Sheet: Name` renames the default sheet only when nothing has been placed
yet. If a header or table came first, the default sheet keeps that content
and a new sheet is created. Table numbers (`T1`, `T2`, …) and the row cursor
restart on every new sheet, so `T1` always means the first table of the sheet
being referenced. Sheet names are stripped of the characters Excel forbids
and cut to 31 characters; a name that collides after sanitising is logged,
because openpyxl will silently suffix it and cross-sheet formulas keyed by the
requested name will then point at the wrong sheet.

### Cell resolution

`resolve_cell()` decides what a data cell is, in this order:

1. A whole-cell `**bold**`, `*italic*` or `` `code` `` wrapper is stripped and
   remembered. Partial inline markup is left as literal text.
2. `=` prefix → formula. The text is kept for the reference rewriter.
3. Trailing `%` with a numeric body → a fraction, with a percent format
   carrying as many decimals as the source text, capped at four.
4. A number, after stripping digit-group separators only where the grouping
   is unambiguous (`1,234` and `1.234,56` yes; `1,5` no, since it is 1.5 in
   most of Europe). `nan`, `inf` and PEP 515 underscores are refused because
   the first two serialise as blank cells and the third would turn a product
   code into a number.
5. A date, tried after numbers so `2024` stays numeric. Candidates must be
   at least six characters and pass a cheap shape check. A fixed list of
   `strptime` formats is tried first, each paired with the Excel display
   format to apply, then `dateutil` in day-first mode for strings of eight
   characters or more.
6. Otherwise text.

Numbers from 1000 up get a thousands-grouping format; smaller values are left
in General so they display as written. A `types` directive bypasses this
whole ladder for its column and applies the declared coercion instead, with
formula cells as the one exception: they still go through the formula path
and then take the column's number format.

### Formula reference rewriting

`adjust_formula_references()` is a sequence of regex substitutions over the
formula string. Order is the whole design: cross-sheet forms first, then
same-sheet table forms, then row-relative forms, and within each group ranges
and function shorthands before single cells so a single-cell pattern can never
eat half of a range.

| Form | Example | Resolves against |
|------|---------|------------------|
| Cross-sheet function | `Sales!T1.SUM(B[0]:B[4])` | that sheet's position map; only `SUM`, `AVERAGE`, `MAX`, `MIN` |
| Cross-sheet range | `Sales!T1.B[0]:T1.B[4]` | that sheet's position map |
| Cross-sheet cell | `Sales!T1.B[0]` | that sheet's position map |
| Table range and function | `T1.B[0]:T1.B[4]`, `T1.SUM(B[0]:B[4])` | this sheet's position map |
| Table cell | `T1.B[0]` | this sheet's position map: header row + 1 + offset |
| Row-relative range and cell | `B[-3]:B[-1]`, `B[0]` | the row the formula is on |

Sheet names may be bare or Excel-quoted (`'P&L'!T1.B[0]`); the output always
uses the canonical quoted form and puts the sheet prefix on the first endpoint
of a range only, which is what Excel itself writes. A reference to a table or
sheet that does not exist is logged at WARNING and resolved against the
current row so the file still ships, but the caller gets no signal beyond the
log. A formula over Excel's 8,192-character limit is stored as text, since one
over-length formula makes Excel refuse the whole file.

### The three directives

A directive is an HTML comment on its own line above a table. Blank lines
between are fine; any other content clears it.

- `types:` one spec per column, comma-separated. `_parse_types_directive()`
  re-joins fragments that do not start with a type keyword, because Excel
  formats themselves contain commas (`number:#,##0.00`).
- `styles:` parsed in `styles.py`. Targets are absolute (`B2`), table-relative
  (`B[0]` is the first data row, `B[-1]` the header) or ranges of either. A
  range over 10,000 cells is dropped with a warning. Named styles apply
  first, inline attributes layer on top. Applied as the last pass over the
  table so an explicit instruction wins over the header fill, the formula
  fill and inline bold.
- `freeze` handled in `_build_workbook()`, freezing panes below the header.

### Named styles without collateral damage

`cell.style = name` in openpyxl replaces font, fill, border, alignment and
number format at once. `_apply_named_style()` restores number format, border
and alignment afterwards wherever the style left them at their defaults, on
the reasoning that a default is what a style with no opinion produces. The
one case this cannot handle is a style that deliberately sets `General` to
strip an inherited format: openpyxl serialises "unset" and "explicitly
General" identically, so the information is not in the file. The docstring on
that function records why it is unsolvable rather than unsolved.

### Excel Table objects

With `auto_filter`, each Markdown table becomes an openpyxl `Table` with a
banded style, named from the sheet title and table index. Excel requires
Table column names to be non-empty and case-insensitively unique and rejects
the whole file otherwise, so `_ensure_unique_table_headers()` renames blanks
to `ColumnN` and suffixes duplicates, logging each rename. Without
`auto_filter` headers are left exactly as written.

### Circular references

After saving, `detect_circular_references()` re-reads the bytes, builds a
dependency graph from formula strings (string literals stripped, ranges
expanded, external and structured references ignored) and runs an iterative
DFS. Excel opens a cyclic workbook with a warning and shows 0 in the cells,
which is invisible from the server side, so this exists to put the fact in
the log. It never raises.

## Extension points

| To add… | Edit | Notes |
|---------|------|-------|
| A new column type | `helpers._apply_column_type()`, `_number_format_for_type()`, `_KNOWN_TYPE_KEYWORDS`, the width estimate in `add_table_to_sheet()` | The keyword set is what lets the directive parser tell a column boundary from a comma inside a format |
| A new formula form | `helpers.adjust_formula_references()` | Insert it in the right position of the substitution order; add a case to `circular_refs.extract_formula_references()` if the resolved output could be a new reference shape |
| A new directive | `parser.DIRECTIVE_PATTERN` already accepts any key; read it in `add_table_to_sheet()` or `_build_workbook()` | Directives are lower-cased keys with optional values |
| A new style attribute | `styles.parse_style_spec()`, `StyleSpec`, `apply_style_spec()` | Flags without a value go in `_FLAG_ATTRS` |
| A new date format | `helpers.DATE_FORMATS` | Pair the `strptime` pattern with the Excel display format; order is most specific first |

## Invariants and gotchas

- **Non-table content is dropped.** Paragraphs, lists and anything else that
  is not a heading, a sheet marker, a directive or a table row never reach
  the workbook. There is no warning. The tool description tells the model
  this; the code relies on it.
- **A failing cell is logged and skipped.** The per-cell loop in
  `add_table_to_sheet()` catches every exception at WARNING and moves on,
  leaving that cell empty. The workbook is still produced.
- **Warnings go to the log only.** Missing tables, unknown sheets, oversized
  style ranges and circular references are all logged and never returned to
  the caller. The Excel tool has no warnings channel like PowerPoint's.
- **Header cells are never converted.** A header that looks like `2024` stays
  the string `2024`; Excel Tables require string headers and a float would
  render as `2024.0`.
- **Keep the substitution order in `adjust_formula_references()`.** Every
  pattern after the first assumes the earlier ones have already consumed
  their matches.
- **`resolve_cell()` is called twice per cell** once for the value and again
  for column-width estimation. Keep it cheap and side-effect free.
- **`TABLE_BOTTOM_SPACING` is defined in two modules** and must stay equal ([#116](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/116)).
  See row bookkeeping above.

## Tests

| File | Covers |
|------|--------|
| `tests/test_xlsx_creation.py` | Multi-sheet layout and cross-sheet references end to end; defines the `_create_workbook_from_markdown()` helper that patches `upload_file` and reloads the saved bytes |
| `tests/test_xlsx_tier1_fixes.py` | Cross-sheet range prefix form, comma handling in the `types` directive, thousands separators, accounting negatives, formulas in typed columns, unresolved-reference warnings, formula reference extraction, circular-reference detection |
| `tests/test_xlsx_tier2_semantics.py` | Row-relative versus table-relative reference semantics, percent precision, thousands format, unambiguous digit grouping |
| `tests/test_xlsx_tier4_robustness.py` | Sheet-name quoting, `nan`/`inf`/underscore refusal, font family preserved by inline formatting, Excel Table header uniqueness, formula-length guard, buffer and upload paths agreeing |
| `tests/test_xlsx_styling.py` | The `styles` directive and template-defined named styles |

The usual pattern is to build the workbook, save it to a buffer, and reload
it with `openpyxl.load_workbook` so assertions see what Excel would see. See
[`../testing.md`](../testing.md).

## Known limitations

- **Cross-sheet function shorthand covers four functions.** `SUM`, `AVERAGE`,
  `MAX` and `MIN` only; anything else must be written as a function over a
  cross-sheet range.
- **Inline formatting is whole-cell.** `**Total**` bolds a cell; `**Total**
  revenue` is written literally with the asterisks.
- **No warnings channel.** See the invariants above; tracked in [#114](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/114).
- **Column widths are estimated from text length**, clamped to 12–25
  characters, and do not account for proportional fonts.
- **The Excel template is not listed in the user docs.** `custom_xlsx_template.xlsx`
  is resolved through the same directories as the other templates but the
  README's static-templates table omits it. To be fixed in
  [`../../templates.md`](../../templates.md).
