<div align="center">

# 📄 MCP Office Documents Server

[![MCP Toplist](https://mcptoplist.com/badge/glama%2FForLegalAI%2Fmcp-ms-office-documents.svg)](https://mcptoplist.com/server/glama%2FForLegalAI%2Fmcp-ms-office-documents)

**Let your AI assistant create professional Office documents — PowerPoint, Word, Excel, emails & XML — with a single prompt.**

[![Docker](https://img.shields.io/badge/Docker-georgx22%2Fmcp--office--docs-blue?logo=docker)](https://hub.docker.com/r/georgx22/mcp-office-docs)
[![MCP](https://img.shields.io/badge/Protocol-MCP-green)](https://modelcontextprotocol.io/)
[![License](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

</div>

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

**Dynamic templates:**

- 📧 **Reusable Email Templates** — Define parameterized email layouts in YAML. Each becomes its own tool with typed arguments (e.g., `first_name`, `promo_code`).
- 📝 **Reusable Word Templates** — Create `.docx` files with `{{placeholders}}`. Each template becomes an AI tool. Placeholders support full Markdown.

**Output options:**
- **Local** — Files saved to the `output/` folder
- **Cloud** — Upload to S3, Google Cloud Storage, Azure Blob, or MinIO and get a time-limited download link

---

## 🚀 Quick Start

### 1. Get the compose file

```bash
curl -L -o docker-compose.yml https://raw.githubusercontent.com/ForLegalAI/mcp-ms-office-documents/master/docker-compose.yml
curl -L -o .env.example https://raw.githubusercontent.com/ForLegalAI/mcp-ms-office-documents/master/.env.example
```

> Already cloned the repo? Skip this step — both files are already there.

### 2. Set up your environment

```bash
cp .env.example .env
```

The defaults work out of the box — files are saved locally to `output/`.

### 3. Start the server

```bash
docker compose up -d
```

✅ **Done!** Your MCP endpoint is ready at `http://localhost:8958/mcp`. Point your client at it — snippets for Claude Desktop, LibreChat and Cursor are in [Connecting an AI client](docs/clients.md).

---

## ⚙️ Configuration

Everything is set through environment variables in `.env`. The ones most deployments touch:

| Variable | Description | Default |
|----------|-------------|---------|
| `API_KEY` | Require an API key on every request (`Bearer`, plain `Authorization`, or `x-api-key`) | _(disabled)_ |
| `UPLOAD_STRATEGY` | Where files go: `LOCAL`, `S3`, `GCS`, `AZURE`, `MINIO`, `LIBRECHAT` | `LOCAL` |
| `SIGNED_URL_EXPIRES_IN` | Lifetime of cloud download links, in seconds | `3600` |
| `ADMIN_ENABLED` | Turn on the browser UI for managing templates | _(off)_ |
| `DEBUG` | Debug-level logging | _(off)_ |

The full list, including the per-backend credentials, thread-pool and multi-replica settings, is in [Configuration](docs/configuration.md). Deployment notes (volumes, health probes, running without Docker, more than one replica) are in [Deployment](docs/deployment.md).

---

## 📚 Documentation

**Using the tools**

- [Markdown reference](docs/markdown-reference.md) — everything the Word and Excel tools accept: headings, lists with restart rules, tables and directives, inline formatting, formulas with table-relative and cross-sheet references, column types, cell styling.
- [PowerPoint slide reference](docs/powerpoint-slides.md) — the fourteen slide types and their fields, bodies in Markdown, warnings, compatibility notes.
- [Templates](docs/templates.md) — drop in your own Word, PowerPoint, email and Excel templates; register several PowerPoint designs by name; turn a Word or email file with `{{placeholders}}` into a tool of its own.
- [Template Admin UI](docs/admin-ui.md) — do all of that from the browser, no YAML, no restart.
- [Connecting an AI client](docs/clients.md)

**Working on the code**

- [Architecture](docs/development/architecture.md), one page per tool under [docs/development/tools/](docs/development/tools/), and how-tos for [adding a tool](docs/development/adding-a-tool.md) or [a storage backend](docs/development/adding-a-backend.md).
- [CONTRIBUTING.md](CONTRIBUTING.md) for setup, checks and conventions; [AGENTS.md](AGENTS.md) for the rules coding agents follow; [SECURITY.md](SECURITY.md) for what the server protects.

The index is at [docs/README.md](docs/README.md).

---

## 🤝 Contributing

Contributions are welcome. Read [CONTRIBUTING.md](CONTRIBUTING.md), then fork, branch, and open a pull request. Bug reports, feature ideas and documentation fixes are all appreciated — open an [issue](https://github.com/ForLegalAI/mcp-ms-office-documents/issues) to start a discussion.

## 📄 License

MIT — see [LICENSE](LICENSE).
