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
| `1. item` / `2. item` | Numbered list (nest by indenting children). The count **continues across anything written between the items** (a section title, a paragraph, a bullet list, a table) as long as the numbers run consecutively; **numbering restarts** only where a list begins again with `1.` |
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
| `[text](url)` | Hyperlink |
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
| a literal `\n` typed as text | Read as a real newline, but never write it deliberately |

A soft break never swallows what follows it: a list, table, heading, quote or page break on the next line is still that block. A `<br>` written *before* a list (`Intro:<br>1. A<br>2. B`) is promoted to a real line break so the list renders as one — which is why a line-leading number that is really text (`1. máje 5`) needs the documented `1\.` escape.

> ⚠️ **One thing `<br>` cannot do:** resume a numbered list that started earlier in the document. `Poznámka<br>3. Třetí`, after a list that reached `2.`, stays one paragraph with "3." as text. Write that continuation on its own line (a blank line, or a line ending in two spaces) and it becomes item 3 — [list continuation](#word) across intervening content is unaffected, it just needs a real line break rather than a `<br>`.

`header_text` / `footer_text` take `<br>` too; they are plain text, not markdown.

**Custom styles** (issue #66) — remap built-in styles or apply an ad-hoc one:

```markdown
<!-- style: Callout -->
This paragraph uses the "Callout" style from your template.
```

The `<!-- style: Name -->` directive applies a style to the next block only — one paragraph, one heading, the whole table, or the top-level items of a list (nested items keep the `List Bullet 2/3` / `List Number 2/3` styles). On a numbered list the style's **own numbering** is used: the list restarts at `1.` with the style's numeral format and indents. The directive must be alone on its line, and the name must match the Word style name exactly. Unknown styles fall back to the default with a warning. To remap styles globally or per template, see [Templates](templates.md).

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
| `<!-- types: text, currency:$, date, bool, number, percent -->` | Force per-column data types (one entry per column; blank = auto). Options: `text` (preserves leading zeros), `currency:<symbol>` (`$ € £ ¥ Kč zł kr CHF R$ ₹`), `date` / `date:<format>`, `bool`, `number` / `number:<format>`, `percent` (`50%` → `0.5`) |
| `<!-- styles: B2=bg:yellow, C[0]:C[3]=color:red;bold -->` | Set cell background and font colour (see below) |

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
