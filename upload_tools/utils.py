import re
import uuid


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


def generate_named_object_name(filename: str, suffix: str) -> str:
    """Generate an object name using a human-readable filename with a short UUID prefix.

    :param filename: Human-readable filename (will be sanitized)
    :param suffix: File extension (e.g., 'pptx', 'docx', 'xlsx', 'eml')
    :return: Object name like 'a1b2c3d4_My_Report.docx'
    """
    safe_name = sanitize_filename(filename)
    short_id = uuid.uuid4().hex[:8]
    return f"{short_id}_{safe_name}.{suffix}"


def get_content_type(file_name: str) -> str:
    """Determine content type based on file extension.

    :param file_name: Name of the file
    :return: MIME type string
    :raises ValueError: If file type is unknown
    """
    if "pptx" in file_name:
        return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    elif "docx" in file_name:
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif "xlsx" in file_name:
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif "eml" in file_name:
        return "application/octet-stream"
    elif "xml" in file_name:
        return "application/xml"
    else:
        raise ValueError("Unknown file type")

