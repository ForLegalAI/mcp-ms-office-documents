from typing import Annotated, List, Dict, Optional, Literal

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from pydantic import BaseModel, Field
from xlsx_tools import markdown_to_excel
from docx_tools import markdown_to_word
from docx_tools.dynamic_docx_tools import register_docx_template_tools_from_yaml
from pptx_tools import create_presentation
from email_tools import create_eml
from email_tools.dynamic_email_tools import register_email_template_tools_from_yaml
from pathlib import Path
from config import get_config
from xml_tools import create_xml_file
from middleware import ApiKeyAuthMiddleware
from async_runner import run_blocking
from starlette.requests import Request
from starlette.responses import PlainTextResponse

import logging
mcp = FastMCP("MCP Office Documents")


# Initialize config and logging (must come before run_blocking, which reads
# config.run_blocking_by_asyncio_thread_enabled at call time)
config = get_config()
logger = logging.getLogger(__name__)

# Register API key auth middleware (only when API_KEY is set)
if config.api_key:
    mcp.add_middleware(ApiKeyAuthMiddleware(expected_key=config.api_key))
    logger.info("[auth] API key authentication enabled")
else:
    logger.info("[auth] No API_KEY set – authentication disabled")


# ---------------------------------------------------------------------------
# Async-safe wrapper for blocking document-generation calls
# ---------------------------------------------------------------------------
# `run_blocking()` is imported from `async_runner` (above). It lives in its
# own module so dynamically-registered tool modules (dynamic_docx_tools,
# dynamic_email_tools) can import it without creating a circular dependency
# back through this `main` module. See `async_runner.py` for the full
# rationale, the EKS failure mode it addresses, and the meaning of the
# `RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED` environment variable.
#
# All MCP tool handlers — both the static ones declared below and the
# dynamic ones registered from YAML — uniformly write
# `await run_blocking(sync_func, ...)`. The helper decides at call time
# whether to dispatch the work to a worker thread (when enabled) or to
# run it inline on the event loop (when disabled, legacy behaviour).
# ---------------------------------------------------------------------------

logger.info(
    "[run_blocking] thread-offload %s (RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED=%s)",
    "ENABLED" if config.run_blocking_by_asyncio_thread_enabled else "DISABLED",
    str(config.run_blocking_by_asyncio_thread_enabled).lower(),
)


# ---------------------------------------------------------------------------
# Kubernetes health-check endpoints
# ---------------------------------------------------------------------------
# These plain HTTP routes are registered at the Starlette layer (via
# FastMCP's @custom_route) and therefore sit OUTSIDE the MCP protocol
# middleware stack — they intentionally bypass the API-key auth middleware
# so kubelet can poll them without credentials.
#
# Why HTTP (not TCP) probes?
#   A TCP probe only verifies the socket accepts a connection. If the
#   Python process is wedged but the listener is still bound, the kubelet
#   never restarts the pod. An HTTP probe forces an actual response from
#   the application, catching "alive port, dead process" failures.
#
# Probes (see k8s deployment manifest):
#   /healthz — startupProbe   (pod has started)
#   /readyz — readinessProbe (pod is ready to receive traffic)
#   /livez — livenessProbe  (pod is alive; restart on failure)
#
# Each endpoint returns 200 OK with a short plain-text body. If the server
# event loop is wedged, the request will time out and kubelet will mark
# the probe as failed.
# ---------------------------------------------------------------------------

@mcp.custom_route("/healthz", methods=["GET"])
async def health_startup(request: Request) -> PlainTextResponse:
    """Startup probe — returns 200 OK once the HTTP server is accepting requests."""
    return PlainTextResponse("ok", status_code=200)


@mcp.custom_route("/readyz", methods=["GET"])
async def health_ready(request: Request) -> PlainTextResponse:
    """Readiness probe — returns 200 OK when the server can serve traffic."""
    return PlainTextResponse("ready", status_code=200)


@mcp.custom_route("/livez", methods=["GET"])
async def health_live(request: Request) -> PlainTextResponse:
    """Liveness probe — returns 200 OK while the event loop is responsive."""
    return PlainTextResponse("alive", status_code=200)

# Look for dynamic email templates in production and local locations.
# Production (container): /app/config/email_templates.yaml
# Local development: <project_root>/config/email_templates.yaml
APP_CONFIG_PATH = Path("/app/config") / "email_templates.yaml"
LOCAL_CONFIG_PATH = Path(__file__).resolve().parent / "config" / "email_templates.yaml"

# Prefer the production path when present, otherwise fall back to local config.
_primary_yaml = None
for candidate in (APP_CONFIG_PATH, LOCAL_CONFIG_PATH):
    if candidate.exists():
        _primary_yaml = candidate
        logger.info("[dynamic-email] Found email templates file: %s", candidate)
        break

if _primary_yaml:
    try:
        register_email_template_tools_from_yaml(mcp, _primary_yaml)
    except Exception as e:
        logger.exception("[dynamic-email] Failed to register email templates from %s: %s", _primary_yaml, e)
else:
    logger.info(
        "[dynamic-email] No dynamic email templates file found at /app/config/email_templates.yaml or config/email_templates.yaml - skipping"
    )

# Look for dynamic DOCX templates in production and local locations.
# Production (container): /app/config/docx_templates.yaml
# Local development: <project_root>/config/docx_templates.yaml
APP_DOCX_CONFIG_PATH = Path("/app/config") / "docx_templates.yaml"
LOCAL_DOCX_CONFIG_PATH = Path(__file__).resolve().parent / "config" / "docx_templates.yaml"

_docx_yaml = None
for candidate in (APP_DOCX_CONFIG_PATH, LOCAL_DOCX_CONFIG_PATH):
    if candidate.exists():
        _docx_yaml = candidate
        logger.info("[dynamic-docx] Found DOCX templates file: %s", candidate)
        break

if _docx_yaml:
    try:
        register_docx_template_tools_from_yaml(mcp, _docx_yaml)
    except Exception as e:
        logger.exception("[dynamic-docx] Failed to register DOCX templates from %s: %s", _docx_yaml, e)
else:
    logger.info(
        "[dynamic-docx] No dynamic DOCX templates file found at /app/config/docx_templates.yaml or config/docx_templates.yaml - skipping"
    )

class PowerPointSlide(BaseModel):
    """PowerPoint slide - can be title, section, content, or table slide based on slide_type."""
    slide_type: Literal["title", "section", "content", "table"] = Field(description="Type of slide: 'title' for presentation opening, 'section' for dividers, 'content' for slide with bullet points, 'table' for slide with a table")
    slide_title: str = Field(description="Title text for the slide")

    # Optional fields based on slide type
    author: Optional[str] = Field(default="", description="Author name for title slides - appears in subtitle placeholder. Leave empty for section/content/table slides.")
    slide_text: Optional[List[Dict]] = Field(
        default=None,
        description="Array of bullet points for content slides. Each bullet point must have 'text' (string) and 'indentation_level' (integer 1-5). Leave empty/null for title, section, and table slides."
    )
    table_data: Optional[List[List[str]]] = Field(
        default=None,
        description="Table data for table slides. A list of rows where each row is a list of cell values (strings). The first row is treated as the header row. Leave empty/null for title, section, and content slides."
    )

@mcp.tool(
    name="create_excel_from_markdown",
    description="Converts markdown content with tables and formulas to Excel (.xlsx) format. Use '## Sheet: Name' headings to create multiple sheets.",
    tags={"excel", "spreadsheet", "data"},
    annotations={"title": "Markdown to Excel Converter"}
)
async def create_excel_document(
    markdown_content: Annotated[str, Field(description="Markdown content containing tables, headers, and formulas. Use '## Sheet: Sheet Name' to create multiple worksheets. Use T1.B[0] for cross-table references and B[0] for current row references. Use SheetName!T1.B[0] for cross-sheet references (resolves to SheetName!B2 in Excel). ALWAYS use [0], [1], [2] notation, NEVER use absolute row numbers like B2, B3. Do NOT count table header as first row, first row has index [0]. Supports cell formatting: **bold**, *italic*.")],
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """
    Converts markdown to Excel with advanced formula support.
    """

    logger.info("Converting markdown to Excel document")

    try:
        result = await run_blocking(markdown_to_excel, markdown_content, file_name=file_name)
        logger.info("Excel document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating Excel document: {e}", exc_info=True)
        raise ToolError(f"Error creating Excel document: {e}")

@mcp.tool(
    name="create_word_from_markdown",
    description="Converts markdown content to a professionally formatted Word (.docx) document. Supports headings, lists, tables, images, block quotes, page breaks, horizontal lines, text alignment, and rich inline formatting.",
    tags={"word", "document", "text", "legal", "contract"},
    annotations={"title": "Markdown to Word Converter"}
)
async def create_word_document(
    markdown_content: Annotated[str, Field(description=(
        "Markdown content for the document body. Separate all block elements with blank lines.\n"
        "\n"
        "BLOCK ELEMENTS (each on its own line):\n"
        "- Headings: # H1, ## H2, ### H3, #### H4, ##### H5, ###### H6\n"
        "- Unordered lists: - item (or * or +); nest with 3-space indent\n"
        "- Ordered lists: 1. item, 2. item; nest with 3-space indent\n"
        "- Tables: | H1 | H2 |\\n|---|---|\\n| C1 | C2 | (cells support inline formatting)\n"
        "- Block quotes: > text (supports inline formatting)\n"
        "- Page break: --- (three+ dashes alone on a line — starts new page)\n"
        "- Horizontal line: *** (three+ asterisks alone on a line — visual separator)\n"
        "- Images: ![alt text](url)\n"
        "\n"
        "INLINE FORMATTING (usable in paragraphs, headings, lists, tables, quotes):\n"
        "- **bold**, *italic*, ***bold italic***\n"
        "- ~~strikethrough~~, __underline__ (double underscore — NOT bold)\n"
        "- `code` (Courier New font)\n"
        "- [link text](https://url)\n"
        "- Nesting: **bold with *italic* inside**, *italic with **bold** inside*\n"
        "- Combinations: **~~bold strikethrough~~**, **__bold underline__**, *~~italic strikethrough~~*\n"
        "- Escaped literals: \\* \\** \\` to render *, **, ` without formatting\n"
        "\n"
        "TEXT ALIGNMENT (HTML tags):\n"
        "- <center>text</center> or multi-line: <center>\\nline1\\nline2\\n</center>\n"
        "- <div align=\"right|center|justify|left\">text</div> (single or multi-line)\n"
        "\n"
        "LINE BREAKS: End a line with two trailing spaces for a soft break within the same paragraph.\n"
        "\n"
        "CONVENTIONS:\n"
        "- Do NOT confuse --- (page break) with *** (horizontal line).\n"
        "- Do NOT confuse __underline__ with bold; always use **bold** for bold.\n"
        "- LEGAL CONTRACTS: use numbered lists (1., 2., 3.) for clauses, nested lists for provisions; use headings only for the contract title.\n"
        "- Other documents: use headings (# ## ###) to organize sections.\n"
    ))],
    title: Annotated[Optional[str], Field(description="Document title (shown in file properties)", default=None)] = None,
    author: Annotated[Optional[str], Field(description="Document author name (shown in file properties)", default=None)] = None,
    subject: Annotated[Optional[str], Field(description="Document subject/description (shown in file properties)", default=None)] = None,
    header_text: Annotated[Optional[str], Field(description="Text for document header (top of every page). Use {page} for auto page number, {pages} for total pages.", default=None)] = None,
    footer_text: Annotated[Optional[str], Field(description="Text for document footer (bottom of every page). Use {page} for auto page number, {pages} for total pages.", default=None)] = None,
    include_toc: Annotated[Optional[bool], Field(description="If true, inserts a Table of Contents at the beginning of the document. The TOC updates automatically when opened in Word.", default=False)] = False,
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """
    Converts markdown to professionally formatted Word document.

    """

    logger.info("Converting markdown to Word document")

    try:
        result = await run_blocking(
            markdown_to_word,
            markdown_content,
            title=title,
            author=author,
            subject=subject,
            header_text=header_text,
            footer_text=footer_text,
            include_toc=include_toc or False,
            file_name=file_name,
        )
        logger.info("Word document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating Word document: {e}", exc_info=True)
        raise ToolError(f"Error creating Word document: {e}")

@mcp.tool(
    name="create_powerpoint_presentation",
    description="Creates PowerPoint presentations from structured slides.",
    tags={"powerpoint", "presentation", "slides"},
    annotations={"title": "PowerPoint Presentation Creator"}
)
async def create_powerpoint_presentation(
    slides: Annotated[List[dict], Field(
        description="""List of slide objects. Each slide requires 'slide_type' (str) and type-specific fields:

- title: {slide_type: "title", slide_title: str, author?: str}
- section: {slide_type: "section", slide_title: str}
- content: {slide_type: "content", slide_title: str, slide_text: [{text: str, indentation_level: int (1-3)}]}
- table: {slide_type: "table", slide_title: str, table_data: [[str]] (first row = header), header_color?: str (hex), alternate_rows?: bool}
- image: {slide_type: "image", slide_title?: str, image_url: str, image_caption?: str}
- two_column: {slide_type: "two_column", slide_title: str, left_column: [{text: str, indentation_level: int}], right_column: [{text: str, indentation_level: int}], left_heading?: str, right_heading?: str}
- chart: {slide_type: "chart", slide_title: str, chart_type: str (bar|column|line|pie|doughnut|stacked_bar|area), chart_data: {categories: [str], series: [{name: str, values: [number]}]}, has_legend?: bool, legend_position?: str}
- quote: {slide_type: "quote", slide_title?: str, quote_text: str, quote_author?: str}

All slides support optional 'speaker_notes': str field."""
    )],
    format: Annotated[Literal["4:3", "16:9"], Field(
        default="16:9",
        description="Aspect ratio: '16:9' (widescreen) or '4:3' (traditional)"
    )] = "16:9",
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """Creates PowerPoint presentations with structured slide models and professional templates."""

    logger.info(f"Creating PowerPoint presentation with {len(slides)} slides in {format} format")

    try:
        result = await run_blocking(create_presentation, slides, format, file_name=file_name)
        logger.info(f"PowerPoint presentation created: {result}")
        return result
    except Exception as e:
        logger.error(f"Error creating PowerPoint presentation: {e}", exc_info=True)
        raise ToolError(f"Error creating PowerPoint presentation: {e}")

@mcp.tool(
    name="create_email_draft",
    description="Creates an email draft in EML format with HTML content using preset professional styling.",
    tags={"email", "eml", "communication"},
    annotations={"title": "Email Draft Creator"}
)
async def create_email_draft(
    content: Annotated[str, Field(description="BODY CONTENT ONLY - Do NOT include HTML structure tags like <html>, <head>, <body>, or <style>. Do NOT include any CSS styling. Use <p> for greetings and for signatures, never headers. Use <h2> for section headers (will be bold), <h3> for subsection headers (will be underlined). HTML tags allowed: <p>, <h2>, <h3>, <ul>, <li>, <strong>, <em>, <div>.")],
    subject: Annotated[str, Field(description="Email subject line")],
    to: Annotated[Optional[List[str]], Field(description="List of recipient email addresses", default=None)],
    cc: Annotated[Optional[List[str]], Field(description="List of CC recipient email addresses", default=None)],
    bcc: Annotated[Optional[List[str]], Field(description="List of BCC recipient email addresses", default=None)],
    priority: Annotated[str, Field(description="Email priority: 'low', 'normal', or 'high'", default="normal")],
    language: Annotated[str, Field(description="Language code for proofreading in Outlook (e.g., 'cs-CZ' for Czech, 'en-US' for English, 'de-DE' for German, 'sk-SK' for Slovak)", default="cs-CZ")],
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """
    Creates professional email drafts in EML format with preset styling and language settings.
    """

    logger.info(f"Creating email draft with subject: {subject}")

    try:
        result = await run_blocking(
            create_eml,
            to=to,
            cc=cc,
            bcc=bcc,
            re=subject,
            content=content,
            priority=priority,
            language=language,
            file_name=file_name,
        )
        logger.info(f"Email draft created: {result}")
        return result
    except Exception as e:
        logger.error(f"Error creating email draft: {e}", exc_info=True)
        raise ToolError(f"Error creating email draft: {e}")

@mcp.tool(
    name="create_xml_file",
    description="Creates an XML file from provided XML content.",
    tags={"xml", "data", "configuration"},
    annotations={"title": "XML File Creator"}
)
async def create_xml_document(
    xml_content: Annotated[str, Field(description="Complete, well-formed XML content. Must be valid XML with proper opening and closing tags.")],
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """
    Creates an XML file from provided XML content.
    Validates that the XML is well-formed before saving.
    """

    logger.info("Creating XML file")

    try:
        result = await run_blocking(create_xml_file, xml_content, file_name=file_name)
        logger.info(f"XML file created successfully.")
        return result
    except Exception as e:
        logger.error(f"Error creating XML file: {e}", exc_info=True)
        raise ToolError(f"Error creating XML file: {e}")

if __name__ == "__main__":
    mcp.run(
        transport="streamable-http",
        host="0.0.0.0",
        port=8958,
        log_level=config.logging.mcp_level_str,
        path="/mcp"
    )
