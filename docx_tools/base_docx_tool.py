import io
import logging
from docx import Document

from upload_tools import upload_file
from .document_features import load_templates, set_header_footer, add_toc
from .markdown_processor import process_markdown_content
from .style_map import load_global_style_map

logger = logging.getLogger(__name__)


def _markdown_to_doc(markdown_content, title=None, author=None, subject=None,
                     header_text=None, footer_text=None, include_toc=False,
                     style_map=None):
    """Convert Markdown content to a python-docx Document object.

    This is the core conversion logic, separated from upload concerns so it
    can be used directly in tests or other contexts that need the Document.

    Returns:
        A ``docx.Document`` instance with the rendered content.
    """
    logger.info("Starting markdown_to_doc conversion")
    path = load_templates()

    # Create document with or without template
    try:
        if path:
            logger.debug(f"Using Word template at: {path}")
            doc = Document(path)
        else:
            doc = Document()  # Create blank document if no template
            logger.warning("No template found, creating blank document")
    except Exception as e:
        logger.error("Failed to load Word template '%s': %s", path, e, exc_info=True)
        raise RuntimeError(f"Error loading Word template: {e}") from e

    # Set document metadata
    if title:
        doc.core_properties.title = title
    if author:
        doc.core_properties.author = author
    if subject:
        doc.core_properties.subject = subject

    # Insert Table of Contents if requested (before main content)
    if include_toc:
        add_toc(doc)

    # Set header and footer
    if header_text:
        set_header_footer(doc, header_text, 'header')
    if footer_text:
        set_header_footer(doc, footer_text, 'footer')

    # Parse markdown content into document
    if style_map is None:
        style_map = load_global_style_map()
    try:
        process_markdown_content(doc, markdown_content, return_elements=False,
                                 style_map=style_map)
    except Exception as e:
        logger.error(f"Error in parsing markdown: {e}", exc_info=True)
        raise RuntimeError(f"Error in parsing markdown: {e}") from e

    logger.info("Markdown parsing completed")
    return doc


def markdown_to_word(markdown_content, title=None, author=None, subject=None,
                     header_text=None, footer_text=None, include_toc=False, file_name=None,
                     style_map=None):
    """Convert Markdown to Word document, save to memory and upload."""
    doc = _markdown_to_doc(
        markdown_content,
        title=title,
        author=author,
        subject=subject,
        header_text=header_text,
        footer_text=footer_text,
        include_toc=include_toc,
        style_map=style_map,
    )

    # Save the document to BytesIO and upload
    try:
        logger.info("Saving Word document to memory buffer")
        file_object = io.BytesIO()
        doc.save(file_object)
        file_object.seek(0)

        result = upload_file(file_object, "docx", filename=file_name)
        file_object.close()

        logger.info("Word document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error saving/uploading Word document: {e}", exc_info=True)
        raise RuntimeError(f"Error saving/uploading Word document: {e}") from e
