import re
import uuid

# The canonical extension → MIME map for everything this server uploads.
# ``backends/librechat.py`` imports it rather than keeping a second copy that
# drifts: the two disagreed about ``.eml`` until #116, one calling it
# ``application/octet-stream`` and the other ``message/rfc822``.
MIME_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "pdf": "application/pdf",
    "eml": "message/rfc822",
    "xml": "application/xml",
    "txt": "text/plain",
    "md": "text/markdown",
    "json": "application/json",
    "csv": "text/csv",
}


def file_extension(file_name: str) -> str:
    """The lower-cased extension after the last dot, or ``''`` if there is none."""
    return file_name.rsplit(".", 1)[-1].lower() if "." in file_name else ""


def generate_unique_object_name(suffix: str) -> str:
    """Generate a unique object name using UUID and preserve the file extension."""
    unique_id = str(uuid.uuid4())
    return f"{unique_id}.{suffix}"


def sanitize_filename(name: str) -> str:
    """Sanitize a human-readable name into a safe filename component.

    Strips unsafe characters, replaces whitespace with underscores, and truncates
    to a reasonable length.

    :param name: Raw filename or title string
    :return: Sanitized string safe for use in object/blob names
    """
    # Remove characters that are unsafe for filenames/URLs
    name = re.sub(r'[^\w\s\-.]', '', name)
    # Collapse whitespace to single underscores
    name = re.sub(r'\s+', '_', name.strip())
    # Truncate to 100 chars to avoid overly long names
    name = name[:100]
    return name or "document"


def generate_named_object_name(
    filename: str, 
    suffix: str, 
    add_unique_prefix: bool = False
) -> str:
    """Generate an object name using a human-readable filename.
    
    :param filename: Human-readable filename (will be sanitized)
    :param suffix: File extension (e.g., 'pptx', 'docx', 'xlsx', 'eml')
    :param add_unique_prefix: If True, adds 8-char UUID prefix for uniqueness.
        Default False - LibreChat handles uniqueness with its own UUID prefix.
    :return: Object name like 'My_Report.docx' or 'a1b2c3d4_My_Report.docx'
    """
    safe_name = sanitize_filename(filename)
    if add_unique_prefix:
        short_id = uuid.uuid4().hex[:8]
        return f"{short_id}_{safe_name}.{suffix}"
    return f"{safe_name}.{suffix}"


def get_content_type(file_name: str) -> str:
    """Determine content type from the file's extension.

    The extension is the part after the LAST dot. This used to be a substring
    test — ``"pptx" in file_name`` — which made ``notes.pptx_v2.docx`` a
    PowerPoint, and any name merely containing ``xml`` an XML document (#116).

    Unlike :func:`upload_tools.backends.librechat.get_mime_type`, an unknown
    extension raises rather than falling back: the traditional backends upload
    only what this server generates, so anything else is a bug worth surfacing
    at the call site instead of shipping an object typed as a guess.

    :param file_name: Name of the file
    :return: MIME type string
    :raises ValueError: If file type is unknown
    """
    extension = file_extension(file_name)
    if extension not in MIME_TYPES:
        raise ValueError(f"Unknown file type: {file_name!r}")
    return MIME_TYPES[extension]

