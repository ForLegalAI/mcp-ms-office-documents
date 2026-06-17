<div align="center">

# 📄 MCP Office Documents Server

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
| 📊 **PowerPoint** | `create_powerpoint_presentation` | Title, section & content slides · 4:3 or 16:9 format · Custom templates · Author metadata, footer text & slide numbers · Inline markdown (**bold**, *italic*, ~~strikethrough~~, `code`) · Table column alignment |
| 📝 **Word** | `create_word_from_markdown` | Write in Markdown, get a `.docx` · Headers, lists, tables, links, formatting · Superscript, subscript & highlighted text · Table column alignment, borderless tables & proportional column widths · Multi-paragraph table cells |
| 📈 **Excel** | `create_excel_from_markdown` | Markdown tables → `.xlsx` · Formulas & cell references · Pure-Python formula recalculation (no LibreOffice) so cached values render in previewers · **Circular-reference detection** via graph analysis · Zero-errors-when-recalc-explicit policy with type-grouped error output · Cross-sheet & external-workbook links · Financial-modeling mode (CFA color coding incl. red for external links, source citations, dash/parens/multiples number formats) · Custom default fonts · Configurable recalc timeout |
| 📧 **Email** | `create_email_draft` | HTML email drafts (`.eml`) · Subject, recipients, priority, language |
| 🗂️ **XML** | `create_xml_file` | Well-formed XML files · Auto-validates & adds XML declaration if missing |

All tools accept an optional **`file_name`** parameter. When provided, the output file will use that name (without extension) instead of a randomly generated identifier.

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
| `XLSX_RECALC_ENABLED` | When true, evaluate every Excel formula in-process and inject cached values so the file previews correctly without Excel. Uses the pure-Python `formulas` library (no LibreOffice). | `true` |
| `XLSX_RECALC_TIMEOUT_SECONDS` | Maximum wall-clock seconds for formula recalculation. On timeout the file is delivered without cached values (Excel recalcs on open). | `30` |
| `XLSX_DEFAULT_FONT` | Optional default font family for every cell in generated workbooks (e.g. `Arial`). Empty = openpyxl default (Calibri). Inline `code` formatting always overrides. | _(unset)_ |

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
<summary><strong>🏥 Performance & Health Probes</strong></summary>

The server exposes health-check endpoints that Kubernetes (or any orchestrator) can use for liveness/readiness probes:

| Endpoint | Purpose |
|----------|---------|
| `GET /health` | Basic liveness check |
| `GET /readiness` | Readiness check |

**Thread-pool offloading:** By default (`RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED=true`), all blocking document-generation work is dispatched to a bounded thread pool (`RUN_BLOCKING_MAX_WORKERS` threads, default 4). This keeps the asyncio event loop free to respond to health probes and handle concurrent requests — critical for Kubernetes deployments where blocked probes lead to pod restarts.

Set `RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED=false` only for local debugging or to rule out threading-related issues.

</details>

---

## 🎨 Custom Templates

You can customize the look of generated documents by providing your own templates.

### Static Templates

Place files in the `custom_templates/` folder:

| Document | Filename | Notes |
|----------|----------|-------|
| PowerPoint 4:3 | `custom_pptx_template_4_3.pptx` | |
| PowerPoint 16:9 | `custom_pptx_template_16_9.pptx` | |
| Word | `custom_docx_template.docx` | |
| Email wrapper | `custom_email_template.html` | Base it on `default_templates/default_email_template.html` |

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
- The original font from the placeholder location is preserved

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
<summary><strong>📊 Excel advanced features (formulas, financial modeling, source citations)</strong></summary>

**Formula recalculation.** By default, every formula is evaluated in-process using the pure-Python [`formulas`](https://github.com/vinci1it2000/formulas) library and the computed values are written back into the file as cached values. This means the file previews correctly in tools that don't recalculate on open — Google Sheets preview, mail clients, `openpyxl` loaded with `data_only=True`. No external binary (LibreOffice, Excel) is required. Disable with the `recalc=false` parameter or the `XLSX_RECALC_ENABLED=false` env var. Recalculation is bounded by `XLSX_RECALC_TIMEOUT_SECONDS` (default 30s); on timeout the file is delivered without cached values (Excel recalcs on open).

**Formula error policy.** Formula errors (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#NULL!`, `#NUM!`, `#N/A`) **and circular references** (`#CIRC!`) are always detected. **When `recalc=true` is explicitly requested**, the tool call fails with a descriptive error grouping all errors by type with counts and locations, prefixed with `N/total` so the model understands scope (e.g. `2/15` means 2 errors out of 15 formulas), so the model can fix all errors of a given kind at once — enforcing a zero-errors delivery standard. Example error: `2/15 formula error(s): #DIV/0! (2): Sheet1!B2, Sheet1!B5; #CIRC! (1): Sheet1!A1 — circular references detected (a formula depends on itself, directly or indirectly; fix by breaking the cycle)`. **When recalc runs as a default** (the parameter wasn't passed), errors are logged but the file is still delivered, so a misconfigured environment can never block document generation.

**Circular-reference detection.** Because the `formulas` library silently resolves circular references to nothing (no cached values, no errors) rather than flagging them, this server runs an independent graph-based cycle check: it parses every formula's cell references (including cross-sheet refs, cell ranges, and **3D references** like `SUM(Sheet1:Sheet3!A1)`), builds a dependency graph, and uses DFS to find cycles. Every cell on a cycle is reported as a `#CIRC!` error. This runs even when the recalc engine is unavailable, since it only inspects the formula strings.

**String-result formulas.** Formulas whose result is a **string/text** value (e.g. `=A1&" total"`, `=IF(A1>0,"Yes","No")`, `=VLOOKUP(...)` returning text, `=TEXT(...)`, `=PROPER(...)`) have their cached values injected as OOXML inline strings (`t="str"`) — no shared-strings-table entry is required, so the file previews correctly in Google Sheets, mail clients, and `openpyxl` loaded with `data_only=True`. Numeric, boolean, datetime, and string results are all cached.

**External-workbook links.** Formulas that reference another `.xlsx` file (`=[Forecast.xlsx]Sheet1!A1`) can't be evaluated in-process (the other workbook isn't available to the server) and would otherwise make the recalculation engine abort the entire workbook. Such formulas are detected and skipped during in-process recalculation — the rest of the workbook is still recalculated and cached normally, and the external-link formula is preserved in the file so Excel evaluates it natively when the file opens. (External links also get red font in `financial_modeling=true` mode.)

**Financial-modeling conventions** (recommended for budgets, forecasts, valuation models — apply even without `financial_modeling=true`):

- **Assumptions in their own cells.** Put growth rates, margins, multiples, and other inputs in dedicated assumption cells and reference them in formulas. Write `=B5*(1+$B$6)` not `=B5*1.05` — this keeps the model dynamic (change the assumption, everything recalculates) and makes the assumption auditable and color-codeable.
- **Units in headers.** Always state units in the column header: `Revenue ($mm)`, `Growth Rate (%)`, `EV/EBITDA (x)` — never bare `Revenue`.
- **Formulas, not hardcoded values.** Use Excel formulas for every calculation (totals, growth, ratios, differences) rather than computing values externally and hardcoding them. The spreadsheet should recalculate when source data changes.
- **Color coding** (via `financial_modeling=true`): hardcoded inputs in blue, local formulas in black, cross-sheet references in green, **external-workbook links in red** (formulas referencing another `.xlsx` file, e.g. `=[Forecast.xlsx]Sheet1!A1`).
- **Source-cited cells** get a yellow background when listed in a `sources:` directive.
- **4-digit-year strings** (e.g. `2024`) in data rows are kept as text labels rather than converted to numbers. (Use an explicit `types: text` directive if you want this behaviour without the rest of financial mode.)

**Markdown directives.** Directives (`<!-- freeze -->`, `<!-- types: ... -->`, `<!-- sources: ... -->`) are HTML-comment lines that configure the next table. They can be placed directly above the table OR above a `## Sheet:` header that precedes the table — both work. A directive applies to exactly one table (the next one) and is then cleared, so it never leaks to tables on later sheets. Multiple sheets are created with `## Sheet: Name` headings.

**Source citations.** Add cell comments / notes with a `sources:` directive above the table. A useful citation names where the figure came from, when it was pulled, and a pointer back to the original record, e.g. `Source: 10-K annual report, fiscal year 2024, page 45, line item "Total revenue"` or `Source: internal CRM export, 2025-06-12, row 1843`:

```
<!-- sources: B2=Source: 10-K filing, FY2024, Page 45, Revenue Note, B5=Source: Internal forecast, Q2 2025 -->
| Metric  | Value |
|---------|-------|
| Revenue | 1000  |
| Cost    | 400   |
```

Ranges are supported: `B2:B5=Same source applies`. If a single-cell entry and a range both cover the same cell, the single-cell entry wins (it's more specific), regardless of which is listed first — so `B2=Specific, B2:B5=Range` gives B2 the specific source and B3:B5 the range source. Combine with `financial_modeling=true` to also get the yellow background.

**Number-format variants.** The `types:` directive supports optional variants that control how zeros and negatives render:

| Variant | Example directive | Zero renders as | Negative renders as |
|---|---|---|---|
| (none) | `number` | `0` | `-1,234` |
| `dash` | `number:dash` | `-` | `(1,234)` |
| `parens` | `number:parens` | `0` | `(1,234)` |
| `dash` | `currency:$:dash` | `-` | `($1,234)` |
| `dash` | `percent:dash` | `-` | `(12.5%)` |

**Valuation multiples** (EV/EBITDA, P/E, etc.) get their own top-level type `multiple`, which renders as `12.5x` — the value is stored raw and the `x` is a display suffix. `number:multiple` and `number:multiples` are accepted as aliases (the natural way to ask for "a number formatted as a multiple"):

| Directive | Input | Stored value | Display |
|---|---|---|---|
| `multiple` (or `number:multiple`) | `12.5` or `12.5x` | `12.5` | `12.5x` |
| `multiple:dash` | `12.5` | `12.5` | `12.5x` (zeros as `-`, negatives in parens) |

Accounting-style negative notation in input values is recognised: `($50)` parses to `-50` when the column type is `currency:$`.

**Percent format.** The default percent format is `0.0%` (one decimal, CFA convention) so `50.5%` displays as `50.5%` rather than being rounded to `51%`. Use `percent:integer` if you want no decimals (`0%`).

**Default font.** Set a font family for every cell with the `default_font` parameter or the `XLSX_DEFAULT_FONT` env var. Inline `` `code` `` formatting (Courier New) always overrides.

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
