<div align="center">

# 📄 MCP Office Documents Server

[![MCP Toplist](https://mcptoplist.com/badge/glama%2FForLegalAI%2Fmcp-ms-office-documents.svg)](https://mcptoplist.com/server/glama%2FForLegalAI%2Fmcp-ms-office-documents)

**Let your AI assistant create professional Office documents — PowerPoint, Word, Excel, emails & XML — with a single prompt.**

[![Docker](https://img.shields.io/badge/Docker-Ready-blue?logo=docker)](https://hub.docker.com/)
[![MCP](https://img.shields.io/badge/Protocol-MCP-green)](https://modelcontextprotocol.io/)
[![License](https://img.shields.io/badge/License-MIT-yellow)]()

</div>

---

## 📋 Table of Contents

- [What is this?](#-what-is-this)
- [Features at a Glance](#-features-at-a-glance)
- [Quick Start](#-quick-start)
- [Configuration](#-configuration)
- [Markdown Reference](#-markdown-reference)
- [Custom Templates](#-custom-templates)
- [Connecting Your AI Client](#-connecting-your-ai-client)
- [Contributing](#-contributing)

---

## 💡 What is this?

This is an **MCP (Model Context Protocol) server** that runs in Docker and gives AI assistants (like Claude, Cursor, or any MCP-compatible client) the ability to generate real Office files on demand.

Just ask your AI to _"create a sales presentation"_ or _"draft a welcome email"_ — and it will produce a ready-to-use file for you.

**No coding required.** Install, connect, and start creating.

---

## ✨ Features at a Glance

| Document Type | Tool | Highlights |
|:---:|---|---|
| 📊 **PowerPoint** | `create_powerpoint_presentation` | Typed slide schema (14 slide types incl. KPI, timeline, agenda, free-form blank) · Outline-pane sections · Markdown bullet bodies · Tables, category charts & XY scatter · Inline & data-URI images · Theme colours · Autofit with overflow warnings · Proofing language · 4:3 or 16:9 · Custom templates |
| 📝 **Word** | `create_word_from_markdown` | Write in Markdown, get a `.docx` · Headings, lists (with auto-restart), tables, links, images, block quotes, page breaks & text alignment · Superscript, subscript, underline & highlighted text · Table column alignment, borderless tables, proportional widths & multi-paragraph cells · Headers/footers with page numbers · Table of Contents · Custom style mapping & per-block style tags |
| 📈 **Excel** | `create_excel_from_markdown` | Markdown tables → `.xlsx` · Multiple sheets · Formulas with table-relative & cross-sheet references · Column data types · Freeze panes & auto-filter · Column alignment |
| 📧 **Email** | `create_email_draft` | HTML email drafts (`.eml`) · Subject, recipients, priority, language |
| 🗂️ **XML** | `create_xml_file` | Well-formed XML files · Auto-validates & adds XML declaration if missing |

All tools accept an optional **`file_name`** parameter. When provided, the output file will use that name (without extension) instead of a randomly generated identifier.

All tools also accept an optional **`add_unique_prefix`** parameter. Left unset, it follows the storage backend: `true` for `LOCAL`/`S3`/`GCS`/`AZURE`/`MINIO`, where an 8-character UUID prefix prevents collisions in shared storage (e.g., `ff8ae81d_My_Report.docx`), and `false` for LibreChat, which adds its own UUID prefix during file storage. Set it explicitly to override — `false` gives clean filenames (e.g., `My_Report.docx`).

**Bonus — Dynamic Templates:**

- 📧 **Reusable Email Templates** — Define parameterized email layouts in YAML. Each becomes its own tool with typed arguments (e.g., `first_name`, `promo_code`).
- 📝 **Reusable Word Templates** — Create `.docx` files with `{{placeholders}}`. Each template becomes an AI tool. Placeholders support full Markdown.

**Output options:**
- **Local** — Files saved to the `output/` folder
- **Cloud** — Upload to S3, Google Cloud Storage, Azure Blob, or MinIO and get a time-limited download link

---

## 🚀 Quick Start

Get up and running in **3 steps**:

### 1. Download the compose file

```bash
curl -L -o docker-compose.yml https://raw.githubusercontent.com/dvejsada/mcp-ms-office-docs/main/docker-compose.yml
```

> Already cloned the repo? Skip this step — `docker-compose.yml` is already there.

### 2. Set up your environment

```bash
cp .env.example .env
```

The defaults work out of the box — files will be saved locally to `output/`.

### 3. Start the server

```bash
docker-compose up -d
```

✅ **Done!** Your MCP endpoint is ready at: `http://localhost:8958/mcp`

---

## ⚙️ Configuration

The server is configured through environment variables in your `.env` file.

### Basic Settings

| Variable | Description | Default |
|----------|-------------|---------|
| `DEBUG` | Enable debug logging (`1`, `true`, `yes`) | _(off)_ |
| `API_KEY` | Protect the server with an API key (see Authentication below) | _(disabled)_ |
| `UPLOAD_STRATEGY` | Where to save files: `LOCAL`, `S3`, `GCS`, `AZURE`, `MINIO` | `LOCAL` |
| `SIGNED_URL_EXPIRES_IN` | How long cloud download links stay valid (seconds) | `3600` |
| `RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED` | Offload blocking tool work to a thread pool, keeping the event loop free for health probes & concurrent requests | `true` |
| `RUN_BLOCKING_MAX_WORKERS` | Maximum concurrent worker threads for blocking tool calls | `4` |

<details>
<summary><strong>🔐 Authentication</strong></summary>

Set `API_KEY` in your `.env` to require an API key for all requests:

```
API_KEY=your-secret-key
```

Clients can send the key in any of these headers:

| Header | Format |
|--------|--------|
| `Authorization` | `Bearer your-secret-key` |
| `Authorization` | `your-secret-key` |
| `x-api-key` | `your-secret-key` |

Leave `API_KEY` empty or unset to allow all requests without authentication.

The health-probe routes (`/healthz`, `/readyz`, `/livez`) are the one exception — they stay reachable without a key so orchestrators can poll them. See **🏥 Performance & Health Probes** below.

</details>

<details>
<summary><strong>☁️ AWS S3 Storage</strong></summary>

Set `UPLOAD_STRATEGY=S3` and provide:

| Variable | Description | Required |
|----------|-------------|----------|
| `S3_BUCKET` | S3 bucket name | ✅ Always |
| `AWS_ACCESS_KEY` | AWS access key ID | ⚠️ See below |
| `AWS_SECRET_ACCESS_KEY` | AWS secret access key | ⚠️ See below |
| `AWS_REGION` | AWS region (e.g., `us-east-1`) | ⚠️ See below |

**Credential modes:**

- **Explicit credentials** — Set all three of `AWS_ACCESS_KEY`, `AWS_SECRET_ACCESS_KEY`, and `AWS_REGION`. Recommended for simple setups.

- **AWS default credential chain** — Leave the credential variables unset and boto3 will automatically discover credentials from the standard chain:
  - `AWS_ACCESS_KEY_ID` / `AWS_SECRET_ACCESS_KEY` environment variables
  - Shared credential / config files (`~/.aws/credentials`)
  - AWS SSO sessions (`aws sso login`) — useful for local development
  - **IRSA (IAM Roles for Service Accounts)** — for AWS EKS deployments
  - ECS container credentials / EC2 instance metadata (IMDSv2)

  In this mode only `S3_BUCKET` is required; region is resolved automatically.

</details>

<details>
<summary><strong>☁️ Google Cloud Storage</strong></summary>

Set `UPLOAD_STRATEGY=GCS` and provide:

| Variable | Description |
|----------|-------------|
| `GCS_BUCKET` | GCS bucket name |
| `GCS_CREDENTIALS_PATH` | Path to service account JSON (default: `/app/config/gcs-credentials.json`) |

Mount the credentials file via `docker-compose.yml` volumes.

</details>

<details>
<summary><strong>☁️ Azure Blob Storage</strong></summary>

Set `UPLOAD_STRATEGY=AZURE` and provide:

| Variable | Description |
|----------|-------------|
| `AZURE_STORAGE_ACCOUNT_NAME` | Storage account name |
| `AZURE_STORAGE_ACCOUNT_KEY` | Storage account key |
| `AZURE_CONTAINER` | Blob container name |
| `AZURE_BLOB_ENDPOINT` | _(Optional)_ Custom endpoint for sovereign clouds |

</details>

<details>
<summary><strong>☁️ MinIO / S3-Compatible Storage</strong></summary>

Set `UPLOAD_STRATEGY=MINIO` and provide:

| Variable | Description | Default |
|----------|-------------|---------|
| `MINIO_ENDPOINT` | MinIO server URL (e.g., `https://minio.example.com`) | _(required)_ |
| `MINIO_ACCESS_KEY` | Access key | _(required)_ |
| `MINIO_SECRET_KEY` | Secret key | _(required)_ |
| `MINIO_BUCKET` | Bucket name | _(required)_ |
| `MINIO_REGION` | Region | `us-east-1` |
| `MINIO_VERIFY_SSL` | Verify SSL certificates | `true` |
| `MINIO_PATH_STYLE` | Use path-style URLs (recommended for MinIO) | `true` |

Make sure the bucket exists and your credentials have `PutObject`/`GetObject` permissions.

</details>

<details>
<summary><strong>💬 LibreChat Integration (Experimental)</strong></summary>

> ⚠️ **Note:** This feature is **not yet implemented in the official LibreChat repository**. It requires a custom LibreChat build with MCP file artifact support. A test version is available at:
> 
> 👉 https://github.com/geodanchev/LibreChat/tree/feature/mcp-ms-office-docs-integration

Set `UPLOAD_STRATEGY=LIBRECHAT` to enable seamless document generation within LibreChat conversations. Documents are uploaded to LibreChat's service endpoint and returned as MCP file artifacts that appear as attachments in the chat UI.

| Variable | Description | Required |
|----------|-------------|----------|
| `LIBRECHAT_SERVICE_URL` | Full URL to LibreChat's service files endpoint | ✅ |
| `LIBRECHAT_SERVICE_TOKEN` | Service token for authentication (must match `MCP_SERVICE_TOKEN` in LibreChat) | ✅ |

**Typical URLs:**
- Inside Docker network: `http://api:3080/api/service/files`
- Local development: `http://localhost:3080/api/service/files`

**Generate a service token:**
```bash
openssl rand -hex 32
```

**LibreChat `librechat.yaml` configuration:**
```yaml
mcpServers:
  Office-documents:
    type: streamable-http
    url: http://mcp-office-docs:8958/mcp
    headers:
      X-User-Id: "{{LIBRECHAT_USER_ID}}"
      X-User-Email: "{{LIBRECHAT_USER_EMAIL}}"
```

The `X-User-Id` and `X-User-Email` headers are automatically populated by LibreChat and used to associate uploaded files with the correct user.

</details>

<details>
<summary><strong>🏥 Performance & Health Probes</strong></summary>

The server exposes health-check endpoints that Kubernetes (or any orchestrator) can use for startup/readiness/liveness probes. Each returns `200` with a short plain-text body:

| Endpoint | Probe | Body |
|----------|-------|------|
| `GET /healthz` | `startupProbe` — the pod has started | `ok` |
| `GET /readyz` | `readinessProbe` — the pod can receive traffic | `ready` |
| `GET /livez` | `livenessProbe` — the event loop is responsive; restart on failure | `alive` |

> **These three routes are not protected by `API_KEY`.** They are registered at the Starlette layer (via FastMCP's `@custom_route`), so they sit outside the MCP middleware stack and deliberately bypass the API-key check described under **🔐 Authentication** above — this lets kubelet poll them without credentials. They expose no data beyond the static strings above.

Use HTTP probes rather than TCP ones: a TCP probe only proves the socket still accepts connections, so a wedged Python process with a bound listener would never be restarted. An HTTP probe forces the application itself to answer.

**Thread-pool offloading:** By default (`RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED=true`), all blocking document-generation work is dispatched to a bounded thread pool (`RUN_BLOCKING_MAX_WORKERS` threads, default 4). This keeps the asyncio event loop free to respond to health probes and handle concurrent requests — critical for Kubernetes deployments where blocked probes lead to pod restarts.

Set `RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED=false` only for local debugging or to rule out threading-related issues.

</details>

---

## 📝 Markdown Reference

The Word and Excel tools accept Markdown documents; the PowerPoint tool takes structured slides whose text fields accept Markdown. These references cover **everything** the parsers understand — including features that are easy to miss.

> **Golden rule:** separate every block element (heading, list, table, quote…) with a **blank line**.

<details>
<summary><strong>📊 PowerPoint slides — full schema</strong></summary>

**Tool parameters** (`create_powerpoint_presentation`):

| Parameter | Description |
|-----------|-------------|
| `slides` | Ordered list of slide objects (below). Required. |
| `format` | `16:9` (default) or `4:3`. |
| `author` | Stored in document properties. |
| `footer_text` | Shown on every slide whose layout has a footer placeholder. |
| `show_slide_numbers` | Slide numbers on every slide. |
| `language` | BCP-47 proofing tag, e.g. `cs-CZ`. Set it when the deck is not in the template's language, or Word/PowerPoint flags every word as misspelled. |
| `file_name` | Output filename without extension. |

Every slide takes `type` plus optional `title`, `notes` (speaker notes) and `layout`.

| `type` | Fields |
|--------|--------|
| `title` | `subtitle?` |
| `section` | — |
| `content` | `body` |
| `two_column` | `left`, `right` — each `{heading?, body}` |
| `table` | `rows`, `align?`, `header_color?`, `zebra?`, `font_size?` |
| `chart` | `chart_type`, `categories`, `series`, `legend?`, `data_labels?`, `number_format?`, `chart_title?`, `x_title?`, `y_title?` |
| `scatter` | `series` (`{name, points: [[x, y], …]}`), `legend?`, `chart_title?`, `x_title?`, `y_title?` |
| `image` | `source`, `caption?`, `body?` |
| `quote` | `text`, `attribution?` |
| `kpi` | `items` (`{value, label, delta?}`), 2–4 read best |
| `timeline` | `steps` (`{label, detail?}`), `style?` (`chevron` or `box`) |
| `agenda` | `items?` — omit to build it from the deck's own `section` slides |
| `closing` | `subtitle?`, `contact?` |
| `blank` | `elements` — positioned items, each `{kind: text\|image\|shape, x, y, w, h?}` with lengths in inches (`1.5`, `"1.5in"`) or as a share of the slide (`"40%"`) |

**Sections.** Every `section` slide also starts a section in PowerPoint's outline pane and slide sorter, so the presenter sees the deck's structure rather than a flat list. Slides before the first section slide go in a "Default Section", as PowerPoint would name it.

**Blank slides** are the escape hatch for the one layout no typed slide fits. Elements draw in order, so a later one sits on top; a `text` element takes the same inline markdown as a bullet, a `shape` is one of `rectangle`, `rounded_rectangle`, `ellipse`, `chevron`, `arrow` with an optional `fill` and centred `text`, and an `image` keeps its aspect ratio within its box. Anything that would run past the slide edge is shrunk to fit and reported in `warnings`; anything starting off the slide is skipped and reported.

**Body text.** `body` takes either a Markdown bullet string or explicit bullet objects. Prefer the string:

```json
{"type": "content", "title": "Q3 results",
 "body": "- Revenue **up 12%**\n  - EMEA +18%\n  - APAC +4%\n- Churn flat"}
```

Indent child items with any consistent unit — two spaces, four spaces or a tab. A line without a `-` marker becomes a top-level bullet. The explicit form is `[{"text": "…", "level": 2}]`, where `level` is 1 (outermost) to 5.

**Inline formatting** works in every text field, table cells included: `**bold**`, `*italic*`, `***bold italic***`, `~~strikethrough~~`, `__underline__`, `` `code` ``, `^superscript^`, `~subscript~`, and `[links](https://example.com)`. The same grammar drives the Word tool, so text formats identically in both. A marker only formats when it hugs its text (`**bold**`, not `** bold **`), so prose like `5 * 3 * 2 = 30` is left alone. Escape a literal marker with `\*`, or wrap it in backticks.

**Tables** take raw values — numbers and `null` are fine, not just strings:

```json
{"type": "kpi", "title": "At a glance",
 "items": [{"value": "€4.2M", "label": "ARR", "delta": "+12% vs Q2"},
           {"value": "18%", "label": "Churn"}]}
```

```json
{"type": "table", "title": "Pricing",
 "rows": [["Plan", "Users", "Price"], ["Basic", 5, 9.0], ["Pro", 25, 29.0]],
 "align": ["left", "right", "right"], "header_color": "accent1"}
```

**Colours** accept 6-digit hex with or without `#`, or a theme name (`accent1`…`accent6`, `dark1`, `dark2`, `light1`, `light2`). Prefer a theme name so the deck follows your template's palette.

**Images** take an https URL or an inline data URI, so an image you already hold can be placed without publishing it first:

```json
{"type": "image", "source": "data:image/png;base64,iVBORw0KGgo…", "caption": "Fig 1"}
```

**Warnings.** When the deck is produced but not exactly as asked — an image that would not load, body text shrunk to fit, a footer dropped because the layout has no placeholder — the result carries a `warnings` list alongside the file instead of leaving it in the server log:

```json
{"file": "https://…/deck.pptx", "slide_count": 12,
 "warnings": ["slide 4: body text is about 1.9x the space available and will be shrunk to fit; consider splitting it across slides."]}
```

**Compatibility.** The previous key names (`slide_type`, `slide_title`, `slide_text`, `indentation_level`, `speaker_notes`, `table_data`, `alternate_rows`, `image_url`, `image_caption`, `quote_text`, `quote_author`, `left_column`, `right_column`, `chart_data`, `has_legend`, `legend_position`) are still accepted and mapped onto the current ones, with a note in the log. They will be removed in a future release.

**Clients that cannot send objects.** `slides` is published as one flat object schema — every field of every slide type in a single `properties` map, each saying which types accept it — rather than as a `oneOf` union of the fourteen types. A union is the better model but not a portable declaration: clients that bridge MCP to a provider without `oneOf`/`$ref` drop what they cannot express, show the model `slides: array of string`, and then reject the call against the union they kept, so *every* call fails with a schema error no matter what the model sends. Validation is unchanged — it runs on this server, against the union, and names the slide and field it rejected.

As a last resort a slide sent as a **string** is read rather than refused: a JSON object is the slide it encodes, and anything else is markdown for one slide (`# Heading` plus text is a `title` slide, otherwise a `content` slide with the rest as its body). The whole deck arriving as one JSON string is unpacked the same way. A string that starts like JSON but does not parse is reported instead — a deck truncated in transit must not quietly arrive as one slide titled with the raw JSON. This fallback covers titles and bullets only — send slide objects for everything else.

</details>

<details>
<summary><strong>📝 Word Markdown — full syntax</strong></summary>

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

**Line breaks** — one rule everywhere (prose, headings, list items, quotes, table cells, and the same in [template placeholders](#-custom-templates)):

| Syntax | Result |
|--------|--------|
| blank line | New paragraph / next block |
| `<br>` | **Soft line break** — a new line inside the *same* paragraph (`Street<br>City`) |
| two trailing spaces | The same soft break. Invisible, and many editors strip them — prefer `<br>` |
| a literal `\n` typed as text | Read as a real newline, but never write it deliberately |

A soft break never swallows what follows it: a list, table, heading, quote or page break on the next line is still that block. A `<br>` written *before* a list (`Intro:<br>1. A<br>2. B`) is promoted to a real line break so the list renders as one — which is why a line-leading number that is really text (`1. máje 5`) needs the documented `1\.` escape.

> ⚠️ **One thing `<br>` cannot do:** resume a numbered list that started earlier in the document. `Poznámka<br>3. Třetí`, after a list that reached `2.`, stays one paragraph with "3." as text. Write that continuation on its own line (a blank line, or a line ending in two spaces) and it becomes item 3 — [list continuation](#-word-markdown--full-syntax) across intervening content is unaffected, it just needs a real line break rather than a `<br>`.

`header_text` / `footer_text` take `<br>` too; they are plain text, not markdown.

**Custom styles** (issue #66) — remap built-in styles or apply an ad-hoc one:

```markdown
<!-- style: Callout -->
This paragraph uses the "Callout" style from your template.
```

The `<!-- style: Name -->` directive applies a style to the next block only — one paragraph, one heading, the whole table, or the top-level items of a list (nested items keep the `List Bullet 2/3` / `List Number 2/3` styles). On a numbered list the style's **own numbering** is used: the list restarts at `1.` with the style's numeral format and indents. The directive must be alone on its line, and the name must match the Word style name exactly. Unknown styles fall back to the default with a warning. To remap styles globally or per template, see [Custom Templates](#-custom-templates).

</details>

<details>
<summary><strong>📈 Excel Markdown — full syntax</strong></summary>

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

</details>

---

## 🎨 Custom Templates

You can customize the look of generated documents by providing your own templates.

### Static Templates

Place files in the `custom_templates/` folder:

| Document | Filename | Notes |
|----------|----------|-------|
| PowerPoint 4:3 | `custom_pptx_template_4_3.pptx` | Or register several by name — see below |
| PowerPoint 16:9 | `custom_pptx_template_16_9.pptx` | Or register several by name — see below |
| Word | `custom_docx_template.docx` | |
| Email wrapper | `custom_email_template.html` | Base it on `default_templates/default_email_template.html` |

### Named PowerPoint Templates

<details>
<summary><strong>📊 Register several decks, and how layouts are matched</strong></summary>

Dropping the two files above is still all you need. To offer your AI a *choice*
of decks — or to configure one — create `config/pptx_templates.yaml`:

```yaml
templates:
  - name: corporate_16_9
    description: Brand deck, widescreen. Use for client-facing presentations.
    pptx_path: corporate_16_9.pptx    # bare filename, in custom_templates/
    default: true
    defaults:
      footer_text: "ACME s.r.o. · Confidential"
      language: cs-CZ
  - name: legacy_4_3
    pptx_path: legacy_4_3.potx        # .potx works too
```

Each entry becomes a value for the `template` argument, and
`list_presentation_templates` reports them with their aspect ratio and layout
names. Anything under `defaults:` applies when the tool call does not set it.

**Layouts are matched by name and by shape, never by position.** A template
that reorders, renames or deletes layouts still works — previously a reordered
template silently built the title slide on the section layout. Resolution order:

1. the `layout` field on the slide itself,
2. a `layouts:` mapping in the template's entry,
3. automatic detection from each layout's placeholders,
4. the positional index, with a warning returned to the caller.

Detection reads placeholder *types*, so it is language-independent — a Czech or
German template classifies from the same rules as an English one. It recognises
`title`, `section`, `content`, `two_column`, `comparison`, `image_text`,
`title_only` and `blank`. Vertical-text layouts are never picked automatically,
since they share a signature with their horizontal counterparts; name one
explicitly if you want it.

Only add a `layouts:` mapping when detection picks the wrong one:

```yaml
    layouts:
      content: "Brand Body"
      section: "Brand Divider"
```

Every registered template is opened once at startup and its coverage logged, so
a template missing a layout the tool needs shows up then rather than as a
strangely laid-out deck later. Templates are re-read when the config or
template directories change, so a new or edited template takes effect **without
restarting the server**.

The file `config/pptx_templates.yaml` in this repository documents every option
in full.

You can also register and configure templates from the browser instead of
writing YAML — see [Template Admin UI](#-template-admin-ui-optional). It reports
each layout's detected role and previews a sample deck through the template
before you save.

</details>

### Dynamic Email Templates

Create reusable, parameterized email layouts that your AI can fill in automatically.

<details>
<summary><strong>📧 How to set up dynamic email templates</strong></summary>

**1.** Create `config/email_templates.yaml`:

```yaml
templates:
  - name: welcome_email
    description: Welcome email with optional promo code
    html_path: welcome_email.html  # must be in custom_templates/ or default_templates/
    annotations:
      title: Welcome Email
    args:
      - name: first_name
        type: string
        description: Recipient's first name
        required: true
      - name: promo_code
        type: string
        description: Optional promotional code (HTML formatted)
        required: false
```

**2.** Create the HTML file in `custom_templates/welcome_email.html`:

```html
<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8" /></head>
<body>
  <h2>Welcome {{first_name}}!</h2>
  <p>We're excited to have you on board.</p>
  {{{promo_code_block}}}
  <p>Regards,<br/>Support Team</p>
</body>
</html>
```

**How it works:**
- Each template becomes a separate AI tool at startup
- Standard email fields (subject, to, cc, bcc, priority, language) are added automatically
- Use `{{variable}}` for escaped text, `{{{variable}}}` for raw HTML

</details>

### Dynamic Word (DOCX) Templates

Create reusable Word documents with `{{placeholders}}` that support full Markdown formatting.

<details>
<summary><strong>📝 How to set up dynamic DOCX templates</strong></summary>

**1.** Create `config/docx_templates.yaml`:

```yaml
templates:
  - name: formal_letter
    description: Generate a formal business letter
    docx_path: letter_template.docx  # must be in custom_templates/ or default_templates/
    annotations:
      title: Formal Letter Generator
    args:
      - name: recipient_name
        type: string
        description: Full name of the recipient
        required: true
      - name: recipient_address
        type: string
        description: Recipient's address
        required: true
      - name: subject
        type: string
        description: Letter subject
        required: true
      - name: body
        type: string
        description: Main body of the letter (supports markdown)
        required: true
      - name: sender_name
        type: string
        description: Sender's name
        required: true
      - name: date
        type: string
        description: Letter date
        required: false
        default: ""
```

**2.** Create a Word document with placeholders and save as `custom_templates/letter_template.docx`:

```
{{date}}

{{recipient_name}}
{{recipient_address}}

Subject: {{subject}}

{{body}}

{{sender_name}}
```

**How it works:**
- Each template becomes a separate AI tool at startup
- Placeholders can be in the document body, tables, headers, and footers
- Placeholder values support full Markdown (bold, italic, lists, headings…)
- The placeholder's own formatting — font, size, colour, **bold, italic, underline, highlight** — is captured and applied to the replacement text (markdown in the value, e.g. `**bold**`, still wins where it sets formatting)
- Formatting of the surrounding text in the same paragraph (before/after the placeholder) is preserved

</details>

<details>
<summary><strong>🎯 Word style requirements for custom templates</strong></summary>

For proper formatting, make sure these styles exist in your `.docx` template:

| Category | Styles |
|----------|--------|
| Headings | Heading 1 – Heading 6 |
| Bullet lists | List Bullet, List Bullet 2, List Bullet 3 |
| Numbered lists | List Number, List Number 2, List Number 3 |
| Other | Normal, Quote, Table Grid |

> **Tip:** Customize these styles (font, size, color, spacing) in your template — the server will use your styling.

</details>

<details>
<summary><strong>🎨 Custom style mapping (use your own style names)</strong></summary>

If your template defines styles under **different names** than the built-ins above, map them in `config/docx_templates.yaml` so rendered Markdown uses them — no need to rename styles in Word.

A top-level `style_mapping:` applies to every document; each template may add its own `style_mapping:` which overrides the global one for that template.

```yaml
# config/docx_templates.yaml

# Global — applies to all conversions:
style_mapping:
  heading_1: "Brand Title"
  list_number: "Brand Numbers"
  quote: "Brand Quote"
  table: "Brand Table"

templates:
  - name: formal_letter
    docx_path: letter_template.docx
    # Per-template override (wins over the global mapping):
    style_mapping:
      quote: "Letter Quote"
    args: [ ... ]
```

**Recognized keys:** `heading_1`…`heading_6`, `list_number` / `_2` / `_3`, `list_bullet` / `_2` / `_3`, `quote`, `table`, `normal`, `code` (style for fenced code blocks).

**Ad-hoc style tag:** to apply any style to a single block without a mapping, put a directive directly above it:

```markdown
<!-- style: Callout -->
This paragraph uses the "Callout" style.
```

Unknown style names fall back to the document default (with a logged warning) rather than failing the document.

</details>

---

## 🖥️ Template Admin UI (optional)

Prefer clicking over editing YAML? Enable the built-in **template-admin UI** to
create and manage dynamic **Word**, **email** and **PowerPoint** templates from
your browser — no YAML, no restart.

<details>
<summary><strong>How to enable and use it</strong></summary>

**1.** Set these in your `.env`:

```
ADMIN_ENABLED=true
ADMIN_PASSWORD=choose-a-strong-password   # falls back to API_KEY if omitted
# ADMIN_PATH=/admin                        # optional, this is the default
```

**2.** Start the server as usual. The admin UI is served from the **same port**
as the MCP endpoint:

```
http://localhost:8958/admin
```

**3.** Log in with your `ADMIN_PASSWORD`, then:

- **Upload** a Word `.docx` (or email `.html`) that contains `{{placeholders}}`
  (and optionally `{{#if flag}} … {{/if}}` conditionals). Author it in real Word
  — full fidelity is preserved.
- The UI **auto-detects** every placeholder and conditional and pre-builds the
  argument form. Fill in each argument's type, whether it's required, a default,
  and the description the AI sees.
- **Preview** with sample values (generates a real file; never uploaded anywhere).
- **Save** — the template becomes a live MCP tool **immediately**, no restart.
- **Edit** later — adjust arguments, or **Replace document** to upload a new
  version and re-scan it for placeholders (existing arguments are kept).
- **Status** page — see live tool counts, per-template usage (calls/errors/last
  used this session), and a recent activity & error log (filterable to
  warnings/errors).

**PowerPoint templates work differently.** A Word or email template is a
fill-in-the-blanks document: it declares arguments and becomes an MCP tool of
its own. A PowerPoint template is a *design* — slide master, layouts, theme
fonts and palette — with no placeholders and no arguments. It does not become a
tool; it becomes one more value you can pass as the `template` argument of
`create_powerpoint_presentation`. So its page looks different:

- **Upload** a `.pptx` or a `.potx` (a designer's template file is taken as
  readily as a deck).
- The UI reports the **slide size and aspect**, every **layout** with its
  placeholders and the slide role auto-detected for it, and the **theme** fonts
  and colours. It warns about the things that actually spoil a generated deck:
  a role no layout covers (those slides fall back by position), layouts with no
  footer placeholder, and sample slides left in the file.
- **Override a layout per role** only where the automatic choice is wrong — the
  dropdowns are populated with the template's real layout names.
- Set **deck defaults** — footer text, language, slide numbers — applied
  whenever a tool call does not set the same option.
- **Preview** builds a real six-slide sample deck through the template, using
  the layout choices on screen *including ones you have not saved yet*.
- **Save** — the presentation tool can build on it immediately, no restart.

**How it's stored:** the UI writes one file per template into
`config/docx_templates.d/`, `config/email_templates.d/` or
`config/pptx_templates.d/` (plus the asset into `custom_templates/`). Your
hand-written master `config/*.yaml` files are never modified — UI templates are
simply merged on top of them at load time.

> **Single-instance note:** live, no-restart registration assumes one server
> instance owns the template files (the standard docker-compose setup). For a
> multi-replica deployment, put the template files on shared storage and roll
> the pods to pick up changes.

</details>

---

## 🔌 Connecting Your AI Client

Point your MCP-compatible client to the server endpoint:

```
http://localhost:8958/mcp
```

**Examples for popular clients:**

<details>
<summary><strong>Claude Desktop</strong></summary>

Add to your Claude Desktop MCP config:

```json
{
  "mcpServers": {
    "office-documents": {
      "url": "http://localhost:8958/mcp"
    }
  }
}
```

</details>

<details>
<summary><strong>LibreChat</strong></summary>

Add the server to your `librechat.yaml` configuration under `mcpServers`:

```yaml
mcpServers:
  office-documents:
    type: streamableHttp
    url: http://mcp-office-docs:8958/mcp
```

> **Note:** If LibreChat and this server run in the same Docker network, use the container name (`mcp-office-docs`) as the hostname. If they run separately, use `http://localhost:8958/mcp` instead.

To place both services on the same network, add a shared network in your `docker-compose.yml`:

```yaml
services:
  mcp-office-docs:
    # ...existing config...
    networks:
      - shared

  librechat:
    # ...existing config...
    networks:
      - shared

networks:
  shared:
    driver: bridge
```

</details>

<details>
<summary><strong>Cursor / Other MCP Clients</strong></summary>

Use the SSE/streamable HTTP transport and set the endpoint URL to:

```
http://localhost:8958/mcp
```

If you have authentication enabled, add the API key header as required by your client.

</details>

---

## 🤝 Contributing

Contributions are welcome! If you'd like to help improve this project:

1. **Fork** the repository
2. **Create a branch** for your feature or fix (`git checkout -b my-feature`)
3. **Commit** your changes (`git commit -m "Add my feature"`)
4. **Push** to your branch (`git push origin my-feature`)
5. **Open a Pull Request**

Whether it's a bug report, a new feature idea, documentation improvement, or a code contribution — all input is appreciated. Feel free to open an [issue](https://github.com/dvejsada/mcp-ms-office-docs/issues) to start a discussion.
