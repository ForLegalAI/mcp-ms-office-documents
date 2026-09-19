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
| `_markdown_to_excel_buffer()` | `xlsx_tools/base_xlsx_tool.py` | `main.py`, via `run_blocking`. Returns `(BytesIO, warnings)` |
| `_build_workbook()` | `xlsx_tools/base_xlsx_tool.py` | the buffer function; tests that want the `Workbook` object |
| `markdown_to_excel()` | `xlsx_tools/base_xlsx_tool.py` | direct library use; builds and uploads synchronously |

The tool takes `markdown_content` and `auto_filter`. `file_name` and
`add_unique_prefix` go to the upload step, not the builder. The buffer
function hands `main.py` a buffer **and** a list of warnings, wrapped into the
response by `main._with_warnings()` — see
"[The warnings channel](#the-warnings-channel)". There are no
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
  parser.walk_markdown_lines(lines, warnings=…)      ONE pass → list of events
  ├─ <!-- key: value -->   → pending directive (attaches to the next table)
  ├─ ## Sheet: Name        → SheetEvent (rename of the default sheet, or a new one)
  ├─ # … ######            → HeaderEvent at the current row, then +2 rows
  ├─ | … |                 → helpers.parse_table() → TableEvent at the current row,
  │                           then +len(rows)+2; directives attached; T-number assigned
  │                           no table in them → `table_incomplete`
  └─ anything else         → dropped, reported as `line_dropped` (clears pending directives)
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
_markdown_to_excel_buffer(): wb.save(BytesIO); circular_refs.detect_circular_references()
                             → (BytesIO, warnings)
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
| `warnings.py` | The Excel warning codes and their severities; `channel()` builds the per-build collector |

## How the interesting parts work

### Row bookkeeping

Rows are assigned by the parser, not by the writer. A header occupies one row
and advances the cursor by two. A table occupies one row per line including
the header and advances by its length plus two. `add_table_to_sheet()` returns
the next free row for API symmetry, but `_build_workbook()` ignores it and
trusts the event's `start_row`. Both use the same `TABLE_BOTTOM_SPACING`,
defined in `helpers.py` and imported by `parser.py`, so the writer and the
bookkeeping cannot disagree.

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
   carrying as many decimals as the source text, capped at four. A *formula*
   has no source text to take that precision from, so in a `types: percent`
   column it takes the widest precision the column's own literals use —
   otherwise a computed 4.3% rendered as `4%` beside literals that kept their
   decimals ([#126](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/126)).
   `percent:<format>` declares the precision outright and then applies to
   every cell in the column, literal and computed alike.
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
  formats themselves contain commas (`number:#,##0.00`) — `percent:#,##0.0%`
  included. Three specs take a suffix that is a number format
  (`number:`, `date:`, `percent:`) and one takes a symbol (`currency:`);
  `_number_format_for_type()` turns a spec into the format a formula cell in
  that column gets, and is the one place that mapping lives.
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
which is invisible from the server side, so every cell it finds goes into the
log and onto the warnings channel as `circular_reference`. It never raises: a
detector failure must not block delivery of an otherwise valid workbook.

### The warnings channel

Almost everything this tool cannot take literally, it works around: a line
that is not a table row is dropped, a formula naming a table that does not
exist resolves against the current row instead, an over-length formula is
stored as text, a duplicate Table header is renamed. Each keeps a workbook the
caller can open — and each was, until
[#114](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/114),
invisible to them, since the response was a URL and the reason was in the log.

A `WarningChannel` (`warning_channel.py`, codes in `xlsx_tools/warnings.py`)
is created in `_markdown_to_excel_buffer()` and threaded through
`_build_workbook()` → `walk_markdown_lines()`, `add_table_to_sheet()` →
`adjust_formula_references()`, `_write_formula()`,
`_ensure_unique_table_headers()`, `parse_styles_directive()`. It is an
**argument, never module state**: builds run concurrently on `run_blocking`
worker threads.

Warnings are located the way a spreadsheet is: `sheet` plus `cell` where the
problem has a coordinate, `sheet` plus the 1-based source `line` where it is
still markdown.

| Code | Severity | Raised when |
|------|----------|-------------|
| `line_dropped` | error | A line is not a heading, sheet marker, directive or table row, so it is not in the workbook |
| `table_incomplete` | error | Lines starting with `\|` that `parse_table()` could not make a table of (no separator row, or no header); they are consumed there and never reach `line_dropped` |
| `cell_failed` | error | The per-cell loop caught an exception; the cell is empty |
| `table_reference_missing` | error | `T<n>` names a table the target sheet does not have; the reference fell back to the current row |
| `sheet_reference_missing` | error | A cross-sheet reference names a sheet that does not exist; Excel will show `#REF!` |
| `formula_unresolved` | error | `adjust_formula_references()` raised; the formula was written unchanged |
| `formula_too_long` | error | Over Excel's 8,192-character limit; stored as text so the file still opens |
| `circular_reference` | error | The post-save detector found this cell on a cycle |
| `table_separator_missing` | warning | A table parsed with no `\|---\|---\|` row directly under its first row; that row was taken as the header |
| `sheet_name_collision` | warning | A second sheet with the same name; Excel renames it and cross-sheet references break |
| `sheet_name_invalid` | warning | openpyxl rejected the name; the sheet has a different one |
| `header_renamed` | warning | A blank or duplicate header, renamed so the Excel Table is valid |
| `style_entry_invalid` | warning | A `styles:` entry is malformed, unresolvable or sets nothing |
| `style_range_too_large` | warning | A `styles:` range is over `MAX_STYLED_CELLS`; none of it was applied |
| `style_failed` | warning | `apply_style_spec()` raised for one cell |

The channel **de-duplicates** identical `(code, message, location)` entries
and **caps** the number it carries (`warning_channel.DEFAULT_LIMIT`), which
matters most here: a page of prose fed to this tool drops one line per line,
and the cap turns that into fifty warnings plus a `warnings_truncated` note
rather than hundreds.

`parse_table()` never *requires* the separator row — it only skips rows that
look like one, wherever in the run they are — so a table written without one,
or with one somewhere other than under the header, still parses, with its
first row taken as the header. That is usually what the caller meant, but it
is decided for them and it moves every table-relative reference, since
`T1.B[0]` counts from the first row after the header.

A separator cell needs one dash, not three, as in CommonMark and in the Word
tool's own check: demanding three wrote a caller's `|--|--|` into the sheet as
a row of literal dashes and then reported `table_separator_missing` against a
table that had one.

**Position decides, not shape.** Only row 1 of the run can be the separator.
A row of dashes anywhere else is data — `| - | - |` is how a caller writes
"not applicable in either column" — and matching by shape alone dropped it
from the sheet with nothing said. A run of nothing but separator rows has no
header for them to sit under, so it returns an empty `TableData` and is
reported as `table_incomplete`. Word applies the same two rules.

`TableData.has_separator` carries the fact out of the parse so the table
branch can report `table_separator_missing`. It means *the separator was
directly under the first row* — the only position markdown gives it meaning —
not merely that one was seen: a trailing or leading separator leaves which row
is the header exactly as unclear as writing none at all. It defaults to True,
because only markdown that was really parsed can say otherwise. Word's
`parse_table()` returns the same flag as its fourth value and reports the same
code.

Two rules for the sites. A warning explains itself once: `_expand_target()`
returns `None` rather than `[]` when it has already reported why a target was
rejected, so `parse_styles_directive()` does not add a vaguer second entry
about the same range. And a branch that consumes lines owns reporting them:
`parse_table()` eats a run of `|` lines whether or not it finds a table in
them, so the table branch of `walk_markdown_lines()` raises
`table_incomplete` itself — those lines can never reach the `line_dropped`
report in the `else` branch below it.

## Extension points

| To add… | Edit | Notes |
|---------|------|-------|
| A new column type | `helpers._apply_column_type()`, `_number_format_for_type()`, `_KNOWN_TYPE_KEYWORDS`, the width estimate in `add_table_to_sheet()` | The keyword set is what lets the directive parser tell a column boundary from a comma inside a format |
| A new formula form | `helpers.adjust_formula_references()` | Insert it in the right position of the substitution order; add a case to `circular_refs.extract_formula_references()` if the resolved output could be a new reference shape |
| A new directive | `parser.DIRECTIVE_PATTERN` already accepts any key; read it in `add_table_to_sheet()` or `_build_workbook()` | Directives are lower-cased keys with optional values |
| A new style attribute | `styles.parse_style_spec()`, `StyleSpec`, `apply_style_spec()` | Flags without a value go in `_FLAG_ATTRS` |
| A new date format | `helpers.DATE_FORMATS` | Pair the `strptime` pattern with the Excel display format; order is most specific first |
| A new warning | `xlsx_tools/warnings.py` (code + severity), then `warnings.add(...)` at the site | `tests/test_xlsx_warnings.py::test_every_code_has_a_severity` fails if the table is missed; locate it with `sheet=` plus `cell=` or `line=` |

## Invariants and gotchas

- **Non-table content is dropped.** Paragraphs, lists and anything else that
  is not a heading, a sheet marker, a directive or a table row never reach
  the workbook. The tool description tells the model this; the code relies on
  it — and reports each dropped line as `line_dropped`.
- **A failing cell is logged and skipped.** The per-cell loop in
  `add_table_to_sheet()` catches every exception at WARNING and moves on,
  leaving that cell empty. The workbook is still produced.
- **Anything dropped or worked around goes on the warnings channel**
  ([#114](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/114)).
  A log line alone is not enough: the caller gets a success response and never
  reads the server log. A new forgiving branch needs a `warnings.add()` beside
  its `logger` call — see "The warnings channel".
- **Header cells are never converted.** A header that looks like `2024` stays
  the string `2024`; Excel Tables require string headers and a float would
  render as `2024.0`.
- **Keep the substitution order in `adjust_formula_references()`.** Every
  pattern after the first assumes the earlier ones have already consumed
  their matches.
- **`resolve_cell()` is called twice per cell** once for the value and again
  for column-width estimation. Keep it cheap and side-effect free.

## Tests

| File | Covers |
|------|--------|
| `tests/test_xlsx_creation.py` | Multi-sheet layout and cross-sheet references end to end; defines the `_create_workbook_from_markdown()` helper that patches `upload_file` and reloads the saved bytes |
| `tests/test_xlsx_tier1_fixes.py` | Cross-sheet range prefix form, comma handling in the `types` directive, thousands separators, accounting negatives, formulas in typed columns, unresolved-reference warnings, formula reference extraction, circular-reference detection |
| `tests/test_xlsx_tier2_semantics.py` | Row-relative versus table-relative reference semantics, percent precision (literals, formula cells and a declared `percent:<format>`), thousands format, unambiguous digit grouping |
| `tests/test_xlsx_tier4_robustness.py` | Sheet-name quoting, `nan`/`inf`/underscore refusal, font family preserved by inline formatting, Excel Table header uniqueness, formula-length guard, buffer and upload paths agreeing |
| `tests/test_xlsx_styling.py` | The `styles` directive and template-defined named styles |
| `tests/test_xlsx_warnings.py` | The warnings channel: every code, sheet/cell/line locations, the cap, the tool-boundary response shape |
| `tests/test_warning_channel.py` | The shared record and collector the channel is built on |

The usual pattern is to build the workbook, save it to a buffer, and reload
it with `openpyxl.load_workbook` so assertions see what Excel would see. See
[`../testing.md`](../testing.md).

## Known limitations

- **Cross-sheet function shorthand covers four functions.** `SUM`, `AVERAGE`,
  `MAX` and `MIN` only; anything else must be written as a function over a
  cross-sheet range.
- **Inline formatting is whole-cell.** `**Total**` bolds a cell; `**Total**
  revenue` is written literally with the asterisks.
- **Column widths are estimated from text length**, clamped to 12–25
  characters, and do not account for proportional fonts.
