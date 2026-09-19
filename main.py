from typing import Annotated, List, Optional, Literal, Union

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from pydantic import Field
from xlsx_tools import _markdown_to_excel_buffer
from docx_tools import _markdown_to_word_buffer
from docx_tools.dynamic_docx_tools import register_docx_template_tools_from_yaml
from pptx_tools import _create_presentation_buffer, SlidesInput
from pptx_tools.templates import load_specs as load_template_specs, validate_templates
from email_tools import _create_eml_buffer
from email_tools.dynamic_email_tools import register_email_template_tools_from_yaml
from pathlib import Path
from config import get_config
from xml_tools import _create_xml_buffer
from middleware import ApiKeyAuthMiddleware
from async_runner import run_blocking
from starlette.requests import Request
from starlette.responses import PlainTextResponse
from librechat_integration import extract_user_context_from_request, upload_and_format_response

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

# Resolve the config directory: production mount (/app/config) is preferred,
# otherwise the in-project config/ directory for local development.
_APP_CONFIG_DIR = Path("/app/config")
_LOCAL_CONFIG_DIR = Path(__file__).resolve().parent / "config"
_CONFIG_DIR = _APP_CONFIG_DIR if _APP_CONFIG_DIR.exists() else _LOCAL_CONFIG_DIR


def _has_templates(master: Path, spec_subdir: str) -> bool:
    """True if a master YAML or a UI-managed ``*.d`` directory is present.

    The loaders merge both sources, so registration should run whenever EITHER
    exists — including the case where the master YAML was never created but the
    admin UI has written per-template files into ``<spec_subdir>``.
    """
    return master.is_file() or (master.parent / spec_subdir).is_dir()


# Dynamic email templates: master config/email_templates.yaml merged with any
# UI-managed per-template files in config/email_templates.d/.
_email_yaml = _CONFIG_DIR / "email_templates.yaml"
if _has_templates(_email_yaml, "email_templates.d"):
    try:
        register_email_template_tools_from_yaml(mcp, _email_yaml)
    except Exception as e:
        logger.exception("[dynamic-email] Failed to register email templates from %s: %s", _email_yaml, e)
else:
    logger.info("[dynamic-email] No email templates found under %s - skipping", _CONFIG_DIR)

# Dynamic DOCX templates: master config/docx_templates.yaml merged with any
# UI-managed per-template files in config/docx_templates.d/.
_docx_yaml = _CONFIG_DIR / "docx_templates.yaml"
if _has_templates(_docx_yaml, "docx_templates.d"):
    try:
        register_docx_template_tools_from_yaml(mcp, _docx_yaml)
    except Exception as e:
        logger.exception("[dynamic-docx] Failed to register DOCX templates from %s: %s", _docx_yaml, e)
else:
    logger.info("[dynamic-docx] No DOCX templates found under %s - skipping", _CONFIG_DIR)

# PowerPoint templates: validated once at startup so a template missing the
# layouts the tool needs shows up in the log now, rather than as a strangely
# laid-out deck later. Never fatal — a bad template degrades to the built-in
# theme at build time.
try:
    validate_templates()
except Exception as e:
    logger.exception("[pptx-templates] Template validation failed: %s", e)


@mcp.tool(
    name="create_excel_from_markdown",
    description="Converts markdown content with tables and formulas to Excel (.xlsx) format. Use '## Sheet: Name' headings to create multiple sheets.",
    tags={"excel", "spreadsheet", "data"},
    annotations={"title": "Markdown to Excel Converter"}
)
async def create_excel_document(
    markdown_content: Annotated[str, Field(description="Markdown content containing tables, headers, and formulas. Use '## Sheet: Sheet Name' to create multiple worksheets. Formulas MUST start with '=' prefix (e.g. =SUM(A1:A5), =A1/B1*100). There are TWO reference forms and they mean different things. (1) CURRENT-ROW: 'B[0]' is the cell in column B of the row the formula itself is on; 'B[-1]' is the row above, 'B[1]' the row below. Write the same '=B[0]*C[0]' in every data row to get a per-row calculation. (2) TABLE-RELATIVE: 'T1.B[0]' is column B of the FIRST data row of table 1, 'T1.B[1]' the second, and so on — it does not move with the formula, so use it to point at one fixed cell (e.g. a total or an assumption). Do NOT count the table header as a row; the first data row is index [0]. Ranges work in both forms: '=SUM(B[-3]:B[-1])' sums the three rows above, '=SUM(T1.B[0]:T1.B[4])' sums the first five data rows of table 1. Cross-sheet uses the table-relative form only: 'SheetName!T1.B[0]' (resolves to SheetName!B2), and 'SheetName!T1.SUM(B[0]:B[4])' for a cross-sheet range. NEVER use absolute row numbers like B2, B3. Supports cell formatting: **bold**, *italic*. Supports column alignment via separator row: |:---|  left, |:---:| center, |---:| right. Dates are auto-detected and formatted (ISO, European, US formats). Supports HTML comment directives placed on a line directly above a table: '<!-- freeze -->' freezes the header row so it stays visible when scrolling; '<!-- types: text, currency:$, date, bool, number, percent -->' forces column data types per column (comma-separated, one per column; omit or leave blank for auto-detection). Type options: 'text' keeps value as-is (preserves leading zeros), 'currency:<symbol>' strips symbol and stores as number with currency format (symbols: $, €, £, ¥, Kč, zł, kr, CHF, R$, ₹), 'date' or 'date:<format>' parses date with optional Excel format, 'bool' maps true/false/yes/no to Excel boolean, 'number' or 'number:<format>' ensures numeric with optional format, 'percent' converts '50%' to 0.5 with percentage format. '<!-- styles: B2=bg:yellow, C[0]:C[3]=color:red;bold -->' sets cell background and font colour: comma-separated entries of <target>=<attributes>, attributes separated by ';'. Targets are an absolute cell (B2), a table-relative cell where [0] is the first data row and [-1] the header (B[0]), or a range of either form (B2:D5, C[0]:C[3]). Attributes: 'bg:<colour>' background, 'color:<colour>' font colour, 'bold', 'italic', 'underline', and 'style:<Name>' to apply a named style from the server's Excel template. Colours are 6-digit hex (FFFF00 or #FFFF00) or a name (red, darkred, orange, yellow, green, darkgreen, blue, darkblue, purple, pink, brown, grey, lightgrey, darkgrey, cyan, magenta, black, white; 'gray' spellings also work). Styling overrides the built-in header and formula colouring and never changes a cell's value; the colour and font attributes also leave number format, borders and alignment untouched, while a 'style:' reference applies whatever the template's style declares. Multiple directives can be stacked above the same table.")],
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
    auto_filter: Annotated[bool, Field(description="If true, adds Excel auto-filter dropdowns to table headers, enabling sorting and filtering in Excel.", default=False)] = False,
    add_unique_prefix: Annotated[Optional[bool], Field(description="If true, adds 8-char UUID prefix to filename for uniqueness. If not set, defaults to True for traditional storage backends (LOCAL/S3/GCS/AZURE/MINIO) and False for LibreChat.", default=None)] = None,
) -> Union[str, dict]:
    """
    Converts markdown to Excel with advanced formula support.

    Returns:
        For traditional upload strategies: URL string
        For LIBRECHAT strategy: MCP file artifact dict
    """

    logger.info("Converting markdown to Excel document")

    try:
        # Generate document buffer
        file_buffer = await run_blocking(
            _markdown_to_excel_buffer,
            markdown_content,
            auto_filter=auto_filter,
        )
        
        # Upload and format response (handles both LIBRECHAT and traditional backends)
        user_context = extract_user_context_from_request()
        result = await upload_and_format_response(
            file_buffer,
            "xlsx",
            file_name,
            user_context,
            f"Excel spreadsheet '{file_name or 'spreadsheet'}' created successfully.",
            add_unique_prefix=add_unique_prefix,
        )
        file_buffer.close()
        
        logger.info("Excel document created successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating Excel document: {e}", exc_info=True)
        raise ToolError(f"Error creating Excel document: {e}")

@mcp.tool(
    name="create_word_from_markdown",
    description="Converts markdown content to a professionally formatted Word (.docx) document. Supports headings, lists, tables, images, block quotes, page breaks, horizontal lines, text alignment, custom per-block styles, and rich inline formatting.",
    tags={"word", "document", "text", "legal", "contract"},
    annotations={"title": "Markdown to Word Converter"}
)
async def create_word_document(
    markdown_content: Annotated[str, Field(description=(
        "Markdown content for the document body. Separate all block elements with blank lines.\n"
        "\n"
        "BLOCK ELEMENTS (each on its own line):\n"
        "- Headings: # H1, ## H2, ### H3, #### H4, ##### H5, ###### H6\n"
        "- Unordered lists: - item (or * or +); nest by indenting child items (2-4 spaces or a tab)\n"
        "- Ordered lists: 1. item, 2. item; nest by indenting child items. The count continues across anything written between the items (a section title, an explanatory paragraph, a bullet list of exhibits, a table) as long as the numbers run consecutively, so keep numbering 3., 4., … instead of restarting; the parts become a single Word list, which renumbers itself when the document is edited; numbering restarts only where a list begins again at 1.\n"
        "  Because of that, escape the dot of a date that starts a line and would land on a list number — a day-1 date (1\\. ledna 2026) or one whose day is the next number in the running count (3\\. září 2026 written after item 2.) — otherwise it is read as the next list item.\n"
        "- Tables: put each row on its own line — a header row, then a separator row (|---|---|), then data rows; cells support inline formatting, <br> for a line break inside the cell and <br><br> for a new paragraph in it; use :---|:---:|---: in the separator for left/center/right alignment\n"
        "- Borderless table: add <!-- borderless --> on the line before the table (useful for bilingual/parallel layouts)\n"
        "- Column widths: add <!-- widths: 30 70 --> before the table (proportional values, any number of columns)\n"
        "- Block quotes: > text (supports inline formatting)\n"
        "- Page break: --- (three+ dashes alone on a line — starts new page)\n"
        "- Horizontal line: *** (three+ asterisks alone on a line — visual separator)\n"
        "- Images: ![alt text](url)\n"
        "- Code blocks: fence with ``` (or ~~~) on their own lines; content is rendered verbatim in a monospace font and NOT parsed as markdown\n"
        "- Custom style: put <!-- style: Style Name --> on the line directly before a block to render it with that Word style (applies to the next block only — every item of a list, or the table). Unknown styles fall back to the default.\n"
        "\n"
        "INLINE FORMATTING (usable in paragraphs, headings, lists, tables, quotes):\n"
        "- **bold**, *italic*, ***bold italic***\n"
        "- ~~strikethrough~~, __underline__ (double underscore — NOT bold)\n"
        "- `code` (Courier New font)\n"
        "- ^superscript^ (e.g. x^2^), ~subscript~ (e.g. H~2~O)\n"
        "- ==highlighted text== (yellow background)\n"
        "- [link text](https://url)\n"
        "- Nesting: **bold with *italic* inside**, *italic with **bold** inside*\n"
        "- Combinations: **~~bold strikethrough~~**, **__bold underline__**, *~~italic strikethrough~~*\n"
        "- Escaped literals: \\* \\** \\` \\. to render *, **, ` and a line-leading number's dot without formatting\n"
        "\n"
        "TEXT ALIGNMENT (HTML tags):\n"
        "- <center>text</center> for a single line; for multiple lines put <center> and </center> each on their own line with the content on the lines between\n"
        "- <div align=\"right|center|justify|left\">text</div> (single or multi-line)\n"
        "\n"
        "LINE BREAKS: Use real newlines in the text — separate paragraphs and blocks with a blank line. For a soft break (a new line inside the SAME paragraph, e.g. an address block) write <br>: Street<br>City. Ending a line with two trailing spaces does the same, but the spaces are invisible and many editors strip them, so prefer <br>. Both work anywhere text is written — prose, headings, list items, quotes, table cells. Do NOT type the two literal characters backslash-n to mean a newline.\n"
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
    header_text: Annotated[Optional[str], Field(description="Text for document header (top of every page). Use {page} for auto page number, {pages} for total pages, and <br> for a second line. Plain text — markdown is not parsed here.", default=None)] = None,
    footer_text: Annotated[Optional[str], Field(description="Text for document footer (bottom of every page). Use {page} for auto page number, {pages} for total pages, and <br> for a second line. Plain text — markdown is not parsed here.", default=None)] = None,
    include_toc: Annotated[Optional[bool], Field(description="If true, inserts a Table of Contents at the beginning of the document. The TOC updates automatically when opened in Word.", default=False)] = False,
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
    add_unique_prefix: Annotated[Optional[bool], Field(description="If true, adds 8-char UUID prefix to filename for uniqueness. If not set, defaults to True for traditional storage backends (LOCAL/S3/GCS/AZURE/MINIO) and False for LibreChat.", default=None)] = None,
) -> Union[str, dict]:
    """
    Converts markdown to professionally formatted Word document.

    Returns:
        For traditional upload strategies: URL string
        For LIBRECHAT strategy: MCP file artifact dict
    """

    logger.info("Converting markdown to Word document")

    try:
        # Generate document buffer
        file_buffer = await run_blocking(
            _markdown_to_word_buffer,
            markdown_content,
            title=title,
            author=author,
            subject=subject,
            header_text=header_text,
            footer_text=footer_text,
            include_toc=include_toc or False,
        )
        
        # Upload and format response (handles both LIBRECHAT and traditional backends)
        user_context = extract_user_context_from_request()
        result = await upload_and_format_response(
            file_buffer,
            "docx",
            file_name,
            user_context,
            f"Word document '{title or file_name or 'document'}' created successfully.",
            add_unique_prefix=add_unique_prefix,
        )
        file_buffer.close()
        
        logger.info("Word document created successfully")
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
    slides: Annotated[SlidesInput, Field(
        description=(
            "Ordered list of slide objects. Each slide's 'type' selects its shape, and the "
            "item schema lists every field with the types that accept it, so only the rules "
            "it cannot express are repeated here.\n"
            "\n"
            "BODY TEXT ('body' on content slides and on each side of two_column): either a markdown "
            "bullet string or explicit bullet objects. Prefer the string — write '- item' per line and "
            "indent child items to nest them (any consistent unit: 2 spaces, 4 spaces or a tab). "
            "A line with no '-' marker becomes a top-level bullet.\n"
            "\n"
            "INLINE MARKDOWN in every text field — titles, subtitles, bullets, table cells, KPI "
            "figures and labels, timeline steps, captions, quotes and attributions, column headings, "
            "chart and axis titles, blank-slide text: **bold**, *italic*, ***bold italic***, ~~strikethrough~~, "
            "__underline__, `code`, ^superscript^, ~subscript~. A marker formats only when it hugs its "
            "text (**bold**, not ** bold **), so prose such as '5 * 3 * 2' is left alone. Escape a "
            "literal marker with a backslash (\\*) or wrap it in backticks.\n"
            "\n"
            "COLOURS: 6-digit hex with or without '#', or a theme name (accent1..accent6, dark1, dark2, "
            "light1, light2). Prefer a theme name so the deck follows the template's palette.\n"
            "\n"
            "IMAGES: 'source' takes an https URL or an inline data URI (data:image/png;base64,...). "
            "On a template with a picture layout the image is placed in that layout's picture "
            "frame, which crops and positions it the way the template's designer intended, and "
            "'body' and 'caption' go in its text area.\n"
            "\n"
            "TABLES: 'rows' is a list of rows, the first being the header. 'widths' are relative "
            "column weights ([3, 1, 1] gives the first column three times the space), 'fills' colour "
            "a cell ({row, col, color}) or a whole row (omit 'col'), and 'merges' join a block into "
            "one cell ({row, col, row_span, col_span}) for a grouped header. Row 0 is the header row "
            "in all three.\n"
            "\n"
            "CHARTS: use the 'chart' type for category charts and the 'scatter' type for XY data "
            "([x, y] point pairs). Keep each series' 'values' the same length as 'categories'. Both "
            "'chart' and 'image' take an optional 'body' to put takeaways beside them.\n"
            "\n"
            "STRUCTURE: open with 'title', then 'agenda' (omit its 'items' and it is built from the "
            "deck's own section slides), divide with 'section', and end with 'closing'. Each 'section' "
            "slide also starts a section in PowerPoint's outline pane. Use 'kpi' for two to four "
            "headline figures and 'timeline' for three to five sequential steps.\n"
            "\n"
            "BLANK SLIDES: 'blank' is the escape hatch for a layout no other type fits — positioned "
            "'elements' (text, image or shape), each with x, y, w and optional h as inches (1.5) or a "
            "share of the slide ('40%'). Elements draw in order; anything that would run past the "
            "slide edge is shrunk and reported. A 'title' is kept — it is drawn where the "
            "template puts its titles, with your elements over it. Coordinates are absolute "
            "on the slide, and a blank layout usually still carries the template's logo, "
            "header rule and footer: call list_presentation_templates with "
            "include_layouts=true for the 'content_area' that stays clear of them. Prefer a "
            "typed slide whenever one fits.\n"
            "\n"
            "LINKS: [label](https://url) works in any text field except inside a chart, where "
            "PowerPoint does not follow one: there the label is shown as written.\n"
            "\n"
            "Text that overflows its slide is shrunk to fit and reported in the result's 'warnings'; "
            "split the content across slides rather than relying on that. Each warning is an object "
            "with a stable 'code', a 'severity' ('error' means something you asked for is not in the "
            "file, 'warning' that it is there but altered, 'info' a substitution), the 'slide' index "
            "it happened on, and a readable 'message'."
        )
    )],
    format: Annotated[Literal["4:3", "16:9"], Field(
        default="16:9",
        description="Aspect ratio: '16:9' (widescreen) or '4:3' (traditional)"
    )] = "16:9",
    author: Annotated[Optional[str], Field(description="Author name stored in document properties/metadata.", default=None)] = None,
    footer_text: Annotated[Optional[str], Field(description="Footer text displayed on every slide (e.g., company name, confidentiality notice).", default=None)] = None,
    show_slide_numbers: Annotated[bool, Field(description="Show slide numbers on every slide.", default=False)] = False,
    language: Annotated[Optional[str], Field(description="Language tag for proofing in PowerPoint (e.g. 'cs-CZ' for Czech, 'en-US' for English, 'de-DE' for German). Set this when the deck is not in the template's own language, otherwise every word is flagged as a misspelling. Applies to slide text, table cells and speaker notes; text inside a chart (its title, axis titles and category labels) keeps the viewer's own proofing language.", default=None)] = None,
    template: Annotated[Optional[str], Field(description="Name of a registered template to build on. Call list_presentation_templates to see what is available, along with each one's aspect ratio and layout names. Overrides 'format'. Omit to use the default template for the requested aspect ratio.", default=None)] = None,
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
    add_unique_prefix: Annotated[Optional[bool], Field(description="If true, adds 8-char UUID prefix to filename for uniqueness. If not set, defaults to True for traditional storage backends (LOCAL/S3/GCS/AZURE/MINIO) and False for LibreChat.", default=None)] = None,
) -> Union[str, dict]:
    """Creates PowerPoint presentations with structured slide models and professional templates.

    Returns:
        For traditional upload strategies: a dict with the file location,
        slide_count and any warnings — each a {code, severity, slide, message}
        object — or a bare URL string when there is nothing to report, which
        preserves the previous response shape.
        For LIBRECHAT strategy: MCP file artifact dict.
    """

    logger.info(f"Creating PowerPoint presentation with {len(slides)} slides in {format} format")

    try:
        # Generate presentation buffer
        file_buffer, warnings = await run_blocking(
            _create_presentation_buffer,
            slides, format,
            author=author,
            footer_text=footer_text,
            show_slide_numbers=show_slide_numbers,
            language=language,
            template=template,
        )

        # Upload and format response (handles both LIBRECHAT and traditional backends)
        user_context = extract_user_context_from_request()
        result = await upload_and_format_response(
            file_buffer,
            "pptx",
            file_name,
            user_context,
            f"PowerPoint presentation '{file_name or 'presentation'}' with {len(slides)} slides created successfully.",
            add_unique_prefix=add_unique_prefix,
        )
        file_buffer.close()

        logger.info("PowerPoint presentation created successfully")

        # Warnings describe a deck that was produced but not exactly as asked
        # (an image that would not load, text shrunk to fit, a dropped footer).
        # They ride alongside the result so the model can correct its next call
        # instead of the problem living only in the server log. Each one is an
        # object with a stable 'code' and 'severity', so a caller can branch on
        # it rather than parse the sentence.
        reported = [warning.as_dict() for warning in warnings]
        if reported and isinstance(result, str):
            return {
                "file": result,
                "slide_count": len(slides),
                "warnings": reported,
            }
        if reported and isinstance(result, dict):
            result = {**result, "warnings": reported}
        return result
    except ValueError as e:
        # Schema/validation problems: the message names the slide and field.
        logger.error(f"Invalid presentation input: {e}")
        raise ToolError(str(e))
    except Exception as e:
        logger.error(f"Error creating PowerPoint presentation: {e}", exc_info=True)
        raise ToolError(f"Error creating PowerPoint presentation: {e}")

@mcp.tool(
    name="list_presentation_templates",
    description="Lists the PowerPoint templates this server can build on, with their aspect ratio and, on request, each layout's detected slide role, the roles nothing covers, and the slide area kept clear of template decoration.",
    tags={"powerpoint", "presentation", "templates"},
    annotations={"title": "PowerPoint Template Lister", "readOnlyHint": True},
)
async def list_presentation_templates(
    include_layouts: Annotated[bool, Field(description="Include each template's diagnostics: 'layouts' (every layout with its index, name, placeholder types, and the role it was detected as — null when it matched none), 'roles' (the role -> layout mapping the builder will actually use), 'missing_roles' (roles no layout provides, where slides fall back to a layout by position), and 'content_area' — the band the template keeps clear of its logo, rules and footer, as percentages you can pass straight to a 'blank' slide's elements. Useful before setting a slide's 'layout' field, positioning blank elements, or working out why a slide came out on the wrong layout.", default=False)] = False,
) -> dict:
    """Report the registered PowerPoint templates.

    This is what makes the per-slide ``layout`` field usable: a model cannot
    name a layout it has no way to discover.

    Returns:
        A dict with 'templates' (name, description, aspect, default) and, when
        include_layouts is set, each template's layouts with the role each one
        was detected as, the resolved role -> layout mapping, the roles no
        layout provides, and its content area — the rectangle inside the
        template's own decoration, which is what a 'blank' slide's elements
        should stay within.

    The per-layout role is what makes a mis-detected template diagnosable: a
    section layout carrying only a title is 'title_only' by signature, which
    then serves kpi and timeline slides as well, and nothing in a list of
    layout names says so (#121).
    """
    logger.info("Listing PowerPoint templates (include_layouts=%s)", include_layouts)

    try:
        specs = await run_blocking(load_template_specs)
        result = {"templates": [spec.summary() for spec in specs]}

        if include_layouts:
            reports = await run_blocking(validate_templates)
            by_name = {report["name"]: report for report in reports}
            for entry in result["templates"]:
                report = by_name.get(entry["name"], {})
                entry["layouts"] = report.get("layouts", [])
                entry["roles"] = report.get("coverage", {})
                if report.get("content_area"):
                    entry["content_area"] = report["content_area"]
                if report.get("unknown_configured_layouts"):
                    entry["unknown_configured_layouts"] = report["unknown_configured_layouts"]
                if report.get("error"):
                    # A template that would not open was never classified, so
                    # an empty missing_roles beside it would read as "nothing
                    # is missing" about a template nothing is known about.
                    entry["error"] = report["error"]
                else:
                    # Otherwise always present, empty list included: a model
                    # cannot tell an absent key from "nothing is missing", and
                    # this is the field that says a slide type will land on a
                    # fallback layout.
                    entry["missing_roles"] = report.get("missing_roles", [])

        if not result["templates"]:
            result["note"] = (
                "No templates are registered; presentations use the built-in "
                "PowerPoint theme."
            )
        return result
    except Exception as e:
        logger.error(f"Error listing presentation templates: {e}", exc_info=True)
        raise ToolError(f"Error listing presentation templates: {e}")


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
    add_unique_prefix: Annotated[Optional[bool], Field(description="If true, adds 8-char UUID prefix to filename for uniqueness. If not set, defaults to True for traditional storage backends (LOCAL/S3/GCS/AZURE/MINIO) and False for LibreChat.", default=None)] = None,
) -> Union[str, dict]:
    """
    Creates professional email drafts in EML format with preset styling and language settings.

    Returns:
        For traditional upload strategies: URL string
        For LIBRECHAT strategy: MCP file artifact dict
    """

    logger.info(f"Creating email draft with subject: {subject}")

    try:
        # Generate email buffer
        file_buffer = await run_blocking(
            _create_eml_buffer,
            to=to,
            cc=cc,
            bcc=bcc,
            re=subject,
            content=content,
            priority=priority,
            language=language,
        )
        
        # Upload and format response (handles both LIBRECHAT and traditional backends)
        user_context = extract_user_context_from_request()
        result = await upload_and_format_response(
            file_buffer,
            "eml",
            file_name,
            user_context,
            f"Email draft '{subject}' created successfully.",
            add_unique_prefix=add_unique_prefix,
        )
        file_buffer.close()
        
        logger.info("Email draft created successfully")
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
    add_unique_prefix: Annotated[Optional[bool], Field(description="If true, adds 8-char UUID prefix to filename for uniqueness. If not set, defaults to True for traditional storage backends (LOCAL/S3/GCS/AZURE/MINIO) and False for LibreChat.", default=None)] = None,
) -> Union[str, dict]:
    """
    Creates an XML file from provided XML content.
    Validates that the XML is well-formed before saving.

    Returns:
        For traditional upload strategies: URL string
        For LIBRECHAT strategy: MCP file artifact dict
    """

    logger.info("Creating XML file")

    try:
        # Generate XML buffer
        file_buffer = await run_blocking(
            _create_xml_buffer,
            xml_content,
        )
        
        # Upload and format response (handles both LIBRECHAT and traditional backends)
        user_context = extract_user_context_from_request()
        result = await upload_and_format_response(
            file_buffer,
            "xml",
            file_name,
            user_context,
            f"XML file '{file_name or 'document'}' created successfully.",
            add_unique_prefix=add_unique_prefix,
        )
        file_buffer.close()
        
        logger.info("XML file created successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating XML file: {e}", exc_info=True)
        raise ToolError(f"Error creating XML file: {e}")

if __name__ == "__main__":
    if config.admin.enabled:
        # Admin UI enabled: serve the MCP endpoint and the FastHTML template-admin
        # from a SINGLE ASGI process (uvicorn) so saving a template can register
        # its MCP tool live, with no restart.
        import uvicorn
        from admin.app import build_combined_app

        if not config.admin_password_effective:
            logger.warning(
                "[admin] ADMIN_ENABLED is set but neither ADMIN_PASSWORD nor API_KEY "
                "is configured — the admin UI will reject all logins."
            )
        logger.info("[admin] Template-admin UI enabled at %s (port 8958)", config.admin.path)
        uvicorn.run(
            build_combined_app(mcp, config),
            host="0.0.0.0",
            port=8958,
            log_level=config.logging.mcp_level_str,
        )
    else:
        mcp.run(
            transport="streamable-http",
            host="0.0.0.0",
            port=8958,
            log_level=config.logging.mcp_level_str,
            path="/mcp",
            stateless_http=config.stateless_http,
        )
