"""Markdown → Word (.docx) conversion: the three entry points of the Word tool.

``_markdown_to_doc`` builds a python-docx ``Document``; ``_markdown_to_word_buffer``
saves it to ``BytesIO`` and returns ``(BytesIO, warnings)`` (what ``main.py``
calls, so upload is dispatched uniformly and the warnings ride back to the
caller alongside the file, as :class:`~warning_channel.DocumentWarning`
records); ``markdown_to_word`` builds and uploads synchronously for direct
library use and drops the warnings. The Markdown pipeline itself lives in
``markdown_processor``. See docs/development/tools/word.md.
"""
import io
import logging
from typing import List, Tuple

from docx import Document

from inline_markdown import refused_link_targets, refused_links_message
from upload_tools import upload_file
from warning_channel import DocumentWarning
from .document_features import load_templates, set_header_footer, add_toc
from .markdown_processor import process_markdown_content
from .style_map import load_global_style_map
from . import warnings as W
from .warnings import channel as warning_channel

logger = logging.getLogger(__name__)


def _markdown_to_doc(markdown_content, title=None, author=None, subject=None,
                     header_text=None, footer_text=None, include_toc=False,
                     style_map=None, warnings=None):
    """Convert Markdown content to a python-docx Document object.

    This is the core conversion logic, separated from upload concerns so it
    can be used directly in tests or other contexts that need the Document.

    *warnings* is the build's :class:`~warning_channel.WarningChannel`; pass
    one to learn what the render had to work around (a dropped block, an image
    that would not load, a style the template does not define). Omit it and
    those stay in the log, as they were before #114.

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
                                 style_map=style_map, warnings=warnings)
    except Exception as e:
        logger.error(f"Error in parsing markdown: {e}", exc_info=True)
        raise RuntimeError(f"Error in parsing markdown: {e}") from e

    logger.info("Markdown parsing completed")
    return doc


def _markdown_to_word_buffer(markdown_content, title=None, author=None, subject=None,
                             header_text=None, footer_text=None, include_toc=False,
                             style_map=None) -> Tuple[io.BytesIO, List[DocumentWarning]]:
    """Convert Markdown to a Word document and return its bytes and any warnings.

    This function is useful when the caller needs to handle upload separately,
    such as for LibreChat file artifact uploads.

    Returns:
        ``(buffer, warnings)`` — the buffer holds the Word document, positioned
        at the start, and *warnings* are
        :class:`~warning_channel.DocumentWarning` records of anything the
        renderer had to work around, each with a stable ``code``, a
        ``severity`` and the source ``line``, so the caller can fix its
        markdown rather than find the loss in the server log (#114).
    """
    warnings = warning_channel()
    doc = _markdown_to_doc(
        markdown_content,
        title=title,
        author=author,
        subject=subject,
        header_text=header_text,
        footer_text=footer_text,
        include_toc=include_toc,
        style_map=style_map,
        warnings=warnings,
    )
    # The renderer refuses an unsafe link scheme without a channel of its own
    # (add_hyperlink keeps the label as text); the caller hears it here.
    refused = refused_link_targets(markdown_content or "")
    if refused:
        warnings.add(W.LINK_REFUSED, refused_links_message(refused))

    try:
        logger.info("Saving Word document to memory buffer")
        file_object = io.BytesIO()
        doc.save(file_object)
        file_object.seek(0)
    except Exception as e:
        logger.error(f"Error saving Word document to buffer: {e}", exc_info=True)
        raise RuntimeError(f"Error saving Word document: {e}") from e

    if warnings:
        logger.info("Word document rendered with %d warning(s): %s",
                    len(warnings), "; ".join(warnings.messages))
    return file_object, warnings.records()


def markdown_to_word(markdown_content, title=None, author=None, subject=None,
                     header_text=None, footer_text=None, include_toc=False, file_name=None,
                     style_map=None):
    """Convert Markdown to Word document, save to memory and upload."""
    file_object, _warnings = _markdown_to_word_buffer(
        markdown_content,
        title=title,
        author=author,
        subject=subject,
        header_text=header_text,
        footer_text=footer_text,
        include_toc=include_toc,
        style_map=style_map,
    )

    # Upload the document
    try:
        result = upload_file(file_object, "docx", filename=file_name)
        file_object.close()

        logger.info("Word document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error uploading Word document: {e}", exc_info=True)
        raise RuntimeError(f"Error uploading Word document: {e}") from e
