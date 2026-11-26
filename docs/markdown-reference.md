# Markdown reference

The Word and Excel tools accept Markdown documents; the PowerPoint tool takes structured slides whose text fields accept Markdown. These references cover **everything** the parsers understand — including features that are easy to miss.

> **Golden rule:** separate every block element (heading, list, table, quote…) with a **blank line**.

## Word

**Tool parameters** (`create_word_from_markdown`):

| Parameter | Description |
|-----------|-------------|
| `markdown_content` | The document body (see syntax below) |
| `title` / `author` / `subject` | Document properties (file metadata) |
| `header_text` / `footer_text` | Text for the top/bottom of every page. Use `{page}` for the current page number and `{pages}` for the total |
| `include_toc` | Insert an auto-updating Table of Contents at the start |
| `file_name` | Output filename without extension |

**Block elements** (each on its own line, separated by blank lines):

| Syntax | Result |
|--------|--------|
| `# H1` … `###### H6` | Headings 1–6 |
| `- item` / `* item` / `+ item` | Bullet list (nest by indenting children — 2-4 spaces or a tab → `List Bullet 2/3`) |
| `1. item` / `2. item` | Numbered list (nest by indenting children). The count **continues across anything written between the items** (a section title, a paragraph, a bullet list, a table) as long as the numbers run consecutively, and the parts stay **one list in Word** — insert an item in an earlier part and everything after it renumbers; **numbering restarts** only where a list begins again with `1.` |
| `> quote` | Block quote (`Quote` style) |
| `\| A \| B \|` + `\|---\|---\|` | Table (see table features below) |
| ` ``` ` … ` ``` ` (or `~~~`) | Fenced code block — content is rendered verbatim in a monospace font and **not** parsed as markdown |
| `![alt](url)` | Image |
| `---` (3+ dashes) | **Page break** (starts a new page) |
| `***` (3+ asterisks) | Horizontal line (visual separator) |

> ⚠️ Don't confuse `---` (page break) with `***` (horizontal line).

> 💡 Numbered paragraphs of a filing keep counting across the section titles, evidence notes and exhibit lists between them — write `3.`, `4.`, … and they render as one numbered list, not as paragraphs with typed-in numbers.

> 💡 A single numbered line is only treated as a list when it starts at `1.`, continues the count of an earlier list, **or** is followed by another item. This means a standalone date like `23. června 2026` renders as plain text, not a list. Two cases stay ambiguous and need the dot escaped to render as text: a day-1 date (`1. června 2026`), which is indistinguishable from a one-item list, and a date whose day happens to be the next number in a running count (`3. září 2026` right after item `2.`) — write them as `1\. června 2026` / `3\. září 2026`.

**Inline formatting** (works in paragraphs, headings, list items, table cells, quotes):

| Syntax | Result |
|--------|--------|
| `**bold**` · `*italic*` · `***bold italic***` | Bold / italic / both |
| `~~strikethrough~~` | Strikethrough |
| `__underline__` | Underline (double underscore — **not** bold) |
| `==highlight==` | Yellow highlight |
| `` `code` `` | Monospace (Courier New) |
| `^super^` · `~sub~` | Superscript (`x^2^`) / subscript (`H~2~O`) |
| `[text](url)` | Hyperlink — `http`, `https`, `mailto` and `tel` only; any other target keeps its text and is reported as `link_refused` |
| `\*` `\**` `` \` `` `\.` | Escaped literals (render the marker as text — e.g. `1\.` keeps a day-1 date from becoming a list) |

Nesting and combinations work, e.g. `**bold with *italic* inside**`, `**~~bold strikethrough~~**`.

**Table features** — place the directive on the line **directly above** the table:

| Directive / syntax | Effect |
|--------------------|--------|
| `\|:---\|:---:\|---:\|` separator | Column alignment: left / center / right |
| `<!-- borderless -->` | Remove all borders (great for bilingual/parallel layouts) |
| `<!-- widths: 30 70 -->` | Proportional column widths (any number of columns) |
| `<br>` inside a cell | Line break within the cell (same paragraph) |
| `<br><br>` inside a cell | New paragraph within the cell |

Every table row starts with `|`; the trailing `|` is optional, so `| 3 | 4` is the same row as `| 3 | 4 |` (header and separator rows included). A row with fewer cells than the header leaves the rest of its cells empty. Two or more consecutive lines starting with `|` are a table; without a `|---|---|` row under the first, that row becomes the header and the table is reported as `table_separator_missing`. So two lines of prose that both start with `|` (`|x| < 5` then `|y| > 3`) become a table — escape the first pipe (`\|x| < 5`) to keep a line as text. A single line starting with `|`, or lines that are a single `|` each, are written as ordinary text and reported as `table_not_recognised`.

**Text alignment** (HTML tags, single- or multi-line):

```markdown
<center>centered text</center>
<div align="right">right-aligned</div>
<div align="justify">justified paragraph…</div>
```

**Line breaks** — one rule everywhere (prose, headings, list items, quotes, table cells, and the same in [template placeholders](templates.md)):

| Syntax | Result |
|--------|--------|
| blank line | New paragraph / next block |
| `<br>` | **Soft line break** — a new line inside the *same* paragraph (`Street<br>City`) |
| two trailing spaces | The same soft break. Invisible, and many editors strip them — prefer `<br>` |
| `<br>` at the end of a line | Still **one** soft break: the next line continues the same paragraph or list item (`Street<br>` ⏎ `City`). Follow the `<br>` with a blank line and the next line is a new paragraph |
| a literal `\n` typed as text | Read as a real newline, but never write it deliberately |

A heading is always one line: a `<br>` or two spaces at its end do not pull the next line into it.

A soft break never swallows what follows it: a list, table, heading, quote or page break on the next line is still that block. A `<br>` written *before* a list (`Intro:<br>1. A<br>2. B`) is promoted to a real line break so the list renders as one — which is why a line-leading number that is really text (`1. máje 5`) needs the documented `1\.` escape.

> ⚠️ **One thing `<br>` cannot do:** resume a numbered list that started earlier in the document. `Poznámka<br>3. Třetí`, after a list that reached `2.`, stays one paragraph with "3." as text. Write that continuation on its own line (after a blank line, a line ending in two spaces, or a `<br>` followed by Enter) and it becomes item 3 — [list continuation](#word) across intervening content is unaffected, it just needs a real line break rather than a `<br>`.

`header_text` / `footer_text` take `<br>` too; they are plain text, not markdown.

**Custom styles** (issue #66) — remap built-in styles or apply an ad-hoc one:

```markdown
<!-- style: Callout -->
This paragraph uses the "Callout" style from your template.
```

The `<!-- style: Name -->` directive applies a style to the next block only — one paragraph, one heading, the whole table, or the top-level items of a list (nested items keep the `List Bullet 2/3` / `List Number 2/3` styles). On a numbered list the style's **own numbering** is used: the list restarts at `1.` with the style's numeral format and indents. The directive must be alone on its line, and the name must match the Word style name exactly. Unknown styles fall back to the default and are reported in `warnings` (below). To remap styles globally or per template, see [Templates](templates.md).

**Warnings.** The renderer would rather hand you a document with one block missing than no document at all — so a block it cannot render is skipped, an image that will not load becomes a `[Image could not be loaded: …]` line, and a style your template does not define falls back to `Normal`. When any of that happens the result carries a `warnings` list alongside the file, instead of leaving it in the server log:

```json
{"file": "https://…/document.docx",
 "warnings": [
   {"code": "image_failed", "severity": "error", "line": 12,
    "message": "the image at https://… could not be loaded (404 Not Found); a placeholder line stands in its place."},
   {"code": "style_missing", "severity": "warning",
    "message": "style 'Callout' is not defined in the Word template; Normal was used instead."}
 ]}
```

Each warning is an object rather than a sentence, so you can act on it without reading English:

| Field | Meaning |
|-------|---------|
| `code` | a stable identifier for the kind of problem — the thing to branch on |
| `severity` | `error` (something you wrote is not in the document), `warning` (it is there, but not as asked), `info` (a substitution you probably do not mind) |
| `line` | the line of `markdown_content` to look at, counting from 1; absent when the problem has no single line (a style name that came from configuration) |
| `message` | the sentence, naming what to change |

The codes are `block_failed`, `table_failed`, `table_cell_failed`, `image_failed`, `table_not_recognised` (pipe markup that is not a table at all, written as ordinary text instead), `table_separator_missing` (a table with no `|---|---|` row directly under its first row — that row was used as the header), `style_missing`, `style_fallback_missing`, `widths_invalid` and `link_refused` (a link whose scheme is not `http`, `https`, `mailto` or `tel`, kept as plain text). So `any(w["severity"] == "error" for w in warnings)` answers "did the document lose anything". A document with nothing to report comes back exactly as before — the bare link, no `warnings` key.

## Excel

**Tool parameters** (`create_excel_from_markdown`):

| Parameter | Description |
|-----------|-------------|
| `markdown_content` | Markdown containing one or more tables |
| `auto_filter` | Apply Excel auto-filter (dropdown filters) to each table |
| `file_name` | Output filename without extension |

**Sheets & tables:**

| Syntax | Effect |
|--------|--------|
| `\| A \| B \|` + `\|---\|---\|` | A table becomes a block of cells |
| `## Sheet: Name` | Start a new worksheet named `Name` |
| `# Heading` above a table | Used as a title row above the table |

Rows follow the same rule as in Word: each starts with `|`, the trailing `|` is optional.

**Formulas & references** (put a formula in any cell, starting with `=`):

| Reference form | Meaning |
|----------------|---------|
| `=A1`, `=SUM(A1:A5)` | Standard Excel references and functions |
| `B[0]` | **Current row**, column B — moves with the formula. `B[-1]` is the row above, `B[1]` the row below |
| `SUM(B[-3]:B[-1])` | Range over the current row's neighbours |
| `T1.B[0]` | **Table-relative**: table 1, column B, first data row — a fixed cell, does not move |
| `T1.SUM(B[0]:E[0])` | Function over a table range |
| `SheetName!T1.B[0]` | Cross-sheet table reference |

The two `[n]` forms are different on purpose. Write `=B[0]*C[0]` in every data
row to get a per-row calculation; use `T1.B[0]` to point every row at one fixed
cell, such as a total or a shared assumption. The first data row is index `[0]`
— the header is not counted.

```markdown
| Month | Sales | Cumulative  | Growth        |
|-------|-------|-------------|---------------|
| Jan   | 100   | =B[0]       |               |
| Feb   | 200   | =C[-1]+B[0] | =B[0]/B[-1]-1 |
| Mar   | 300   | =C[-1]+B[0] | =B[0]/B[-1]-1 |
| Total | =SUM(T1.B[0]:T1.B[2]) | | |
```

Circular references (a formula that depends on itself, directly or through a
chain) are detected during generation and reported in the server log — Excel
would otherwise show a warning dialog and silently resolve those cells to 0.

**Column directives** — place on the line directly above a table:

| Directive | Effect |
|-----------|--------|
| `<!-- freeze -->` | Freeze panes below the header row (header stays visible when scrolling) |
| `<!-- types: text, currency:$, date, bool, number, percent -->` | Force per-column data types (one entry per column; blank = auto). Options: `text` (preserves leading zeros), `currency:<symbol>` (`$ € £ ¥ Kč zł kr CHF R$ ₹`), `date` / `date:<format>`, `bool`, `number` / `number:<format>` (`number` and `currency` read a decimal comma: `4,3`, `0,125`, `,125`, `1234,567`, `1 234,56` and `1.234,56` are decimals, and Indian grouping (`₹1,23,456.70`) reads too; spaces, non-breaking ones too, only group digits — `10 20` stays text — and `−5` is −5. The one ambiguous form is one to three digits, a comma and exactly three digits — `4,375`, `1,234` — which the table's numbers decide, looked for in this order: the column's own values, then the table's other `number`, `currency` and `percent` columns, then untyped cells that read as numbers (an untyped `2,5` is text, so it decides nothing); the first place that shows anything wins. Integers, `1.`, `5,` and `1234,567` show nothing. If the deciding numbers write decimal commas or dot grouping (`2,5`, `1.234,5`) it is a decimal, and `1.234` is then 1234; if they write decimal points or comma grouping (`1.5`, `1,234.5`) it is thousands; if they write both it stays text; if nothing shows either it is read as thousands and reported as `ambiguous_comma_read_as_thousands`. A decision taken from another column is reported too, as `ambiguous_number_read_from_table`, naming the value that decided it), `percent` / `percent:<format>` (`50%` → `0.5`; each value keeps the precision it was written with, and a formula in the column takes the widest precision the column uses — declare `percent:0.0%` to fix it; the `%` is optional and a decimal comma works as in `number`: `4,3`, `4,3 %`, `0,125`, `1 234,56`, `1.234,56` and Indian-grouped `1,23,456` are percents, but the ambiguous `4,375` and `1,234` could be a decimal or a thousands separator, and a percent column does not look at the rest of the table to decide, so they stay text — write `4.375`). A value the type cannot read — `n/a`, `1,5e3`, `nan` — is kept as text and reported as `typed_value_kept_as_text`; leave a cell empty rather than writing `-` or `n/a` if it has no value |
| `<!-- styles: B2=bg:yellow, C[0]:C[3]=color:red;bold -->` | Set cell background and font colour (see below) |

Without a `types` directive a cell ending in `%` is a percent when the number before it reads by the same rule — `4,3%`, `0,125%`, `1.234,56%` and `1,23,456%` do, `1,234%` stays text.

Without a `types` directive a number with digit grouping is read only where the grouping is certain: `1,234`, `1,234.56` and `1.234,56` are numbers, but a lone decimal comma (`1,5`) stays text, and so does grouping behind a zero (`0,125`, `0,125,000`, `00,125`, `01,234`). An untyped `4,375` is read as 4375 — an untyped cell has no column to learn from and nothing to report with — so declare `types: number` for a column written with decimal commas.

**Cell styling.** Comma-separated `<target>=<attributes>` entries; attributes separated by `;`.

| Target | Meaning |
|--------|---------|
| `B2` | Absolute worksheet cell |
| `B[0]` | Table-relative — first data row of this table; `B[-1]` is the header |
| `B2:D5`, `C[0]:C[3]` | A rectangular block in either form |

| Attribute | Effect |
|-----------|--------|
| `bg:<colour>` | Background fill |
| `color:<colour>` | Font colour |
| `bold`, `italic`, `underline` | Font weight and decoration |
| `style:<Name>` | A named style from your Excel template (see below) |

Colours are 6-digit hex (`FFFF00` or `#FFFF00`) or a name: `red`, `darkred`,
`orange`, `yellow`, `green`, `darkgreen`, `blue`, `darkblue`, `purple`, `pink`,
`brown`, `grey`, `lightgrey`, `darkgrey`, `cyan`, `magenta`, `black`, `white`
(`gray` spellings work too).

```markdown
<!-- styles: A[-1]:D[-1]=bg:darkblue;color:white;bold, D[3]=bg:yellow;bold -->
| Month | Sales | Costs | Profit        |
|-------|-------|-------|---------------|
| Jan   | 100   | 60    | =B[0]-C[0]    |
| Feb   | 200   | 90    | =B[0]-C[0]    |
| Mar   | 300   | 120   | =B[0]-C[0]    |
| Total | =SUM(T1.B[0]:T1.B[2]) | =SUM(T1.C[0]:T1.C[2]) | =SUM(T1.D[0]:T1.D[2]) |
```

Styling overrides the built-in header and formula colouring and never changes
a cell's value. The `bg:` / `color:` / `bold` / `italic` / `underline`
attributes touch nothing else — number format, borders and alignment survive.
A `style:` reference can also carry a number format, border or alignment if
the template's style declares one; where it doesn't, the cell keeps what the
table gave it. An unrecognised colour or malformed entry is logged and skipped
— it never costs you the document.

**Named styles from your own template.** Drop an `.xlsx` containing named cell
styles at `custom_templates/custom_xlsx_template.xlsx`; every style in it
becomes referenceable as `style:<Name>`. This keeps a house look in one file
rather than repeating hex codes in every document:

```markdown
<!-- styles: B[3]=style:Total, C[0]=style:Warning -->
```

Column alignment via the `:---:` separator syntax is honored, and inline `**bold**` / `*italic*` in cells is applied as cell formatting.

**Warnings.** Lines that are not a heading, a `## Sheet:` marker, a directive or a table row have nowhere to go in a spreadsheet and are dropped; a formula naming a table or a sheet that does not exist still resolves, to the wrong cell or to `#REF!`; a formula over Excel's 8,192-character limit is stored as text so the file still opens; a value a `types` column cannot read is kept as text. When any of that happens the result carries a `warnings` list alongside the file, instead of leaving it in the server log:

```json
{"file": "https://…/workbook.xlsx",
 "warnings": [
   {"code": "line_dropped", "severity": "error", "sheet": "Data Report", "line": 3,
    "message": "'This sentence is not a table.' is not a heading, a '## Sheet:' marker, a directive or a table row, so it is not in the workbook. Put prose in a heading or a table cell."},
   {"code": "table_reference_missing", "severity": "error", "sheet": "Summary", "cell": "C7",
    "message": "the formula references T9, but the target sheet has no such table (it has: T1, T2); the reference fell back to the current row and almost certainly points at the wrong cell."}
 ]}
```

Each warning is an object rather than a sentence, so you can act on it without reading English:

| Field | Meaning |
|-------|---------|
| `code` | a stable identifier for the kind of problem — the thing to branch on |
| `severity` | `error` (something you wrote is not in the workbook, or will not compute), `warning` (it is there, but not as asked), `info` (a substitution you probably do not mind) |
| `sheet` | the worksheet it happened on |
| `cell` | the cell coordinate, where the problem has one |
| `line` | the line of `markdown_content` to look at, counting from 1, where it is still markdown |
| `message` | the sentence, naming what to change |

The codes are `line_dropped`, `table_incomplete` (lines that start with `|` but do not form a table — a table needs a header row, a separator row and at least one data row), `cell_failed`, `table_reference_missing`, `sheet_reference_missing`, `formula_unresolved`, `formula_too_long`, `circular_reference`, `table_separator_missing` (a table with no `|---|---|` row directly under its first row — that row was used as the header), `sheet_name_collision`, `sheet_name_invalid`, `header_renamed`, `style_entry_invalid`, `style_range_too_large`, `style_failed`, `typed_value_kept_as_text` (values in a `types` column that do not read as its type and were kept as text — one warning per table column, with the count, located at the first such cell), `ambiguous_comma_read_as_thousands` (severity `info`: a `number` or `currency` column read `4,375` as 4375 because nothing else in the table shows whether it writes decimal commas — one per table column, with the count; write `4.375` or `4375` to settle it) and `ambiguous_number_read_from_table` (severity `info`: the column's own values showed nothing, so its `4,375` — or, in a table read as European, its `1.234` — was read the way another column writes numbers; one per table column, with the count and the cell that decided it; write the whole number without the separator, or the decimal the note gives, to settle it). A workbook with nothing to report comes back exactly as before — the bare link, no `warnings` key; note that a plain English column of thousands (`1,234`, `12,500`) in a table with no decimal anywhere now does have something to report, an `info` note, and so does one settled by another column — write `1234` and `12500` for the bare link. A very long run of warnings is cut off after fifty, with a final `warnings_truncated` note saying how many were left out.
