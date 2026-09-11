"""Slides → PowerPoint (.pptx): the entry points of the PowerPoint tool.

``_create_presentation_buffer`` builds a deck through
:class:`pptx_tools.slide_builder.PowerpointPresentation` and returns
``(BytesIO, warnings)`` — what ``main.py`` calls, so the warnings ride back to
the caller alongside the file. ``create_presentation`` builds and uploads
synchronously for direct library use and drops the warnings. See
docs/development/tools/powerpoint.md.
"""
import io
import logging
from typing import Any, List, Optional, Sequence, Tuple

from upload_tools import upload_file
from .constants import DEFAULT_SLIDE_FORMAT
from .slide_builder import PowerpointPresentation

logger = logging.getLogger(__name__)


def _create_presentation_buffer(
    slides: Sequence[Any],
    format: str = DEFAULT_SLIDE_FORMAT,
    author: Optional[str] = None,
    footer_text: Optional[str] = None,
    show_slide_numbers: bool = False,
    language: Optional[str] = None,
    template: Optional[str] = None,
) -> Tuple[io.BytesIO, List[str]]:
    """Create a PowerPoint presentation and return its bytes and any warnings.

    This function is useful when the caller needs to handle upload separately,
    such as for LibreChat file artifact uploads.

    :param slides: Slide models, or dicts in the current or previous spelling
    :param format: "4:3" or "16:9" (defaults to the shared DEFAULT_SLIDE_FORMAT)
    :param author: Author name for document properties
    :param footer_text: Optional footer text displayed on all slides
    :param show_slide_numbers: Whether to show slide numbers
    :param language: BCP-47 tag stamped on every run for proofing
    :param template: Name of a registered template; overrides *format*
    :return: ``(buffer, warnings)`` — the buffer is positioned at the start,
        and warnings describe anything the builder had to work around so the
        caller can act on it instead of only seeing it in the server log.
    """
    if not slides:
        raise ValueError("No slides provided")

    logger.info(f"Starting _create_presentation_buffer: slides={len(slides)}, format={format}")

    presentation = PowerpointPresentation(
        slides, format,
        author=author,
        footer_text=footer_text,
        show_slide_numbers=show_slide_numbers,
        language=language,
        template=template,
    )
    return presentation.save(), presentation.warnings


def create_presentation(
    slides: Sequence[Any],
    format: str = DEFAULT_SLIDE_FORMAT,
    file_name: Optional[str] = None,
    author: Optional[str] = None,
    footer_text: Optional[str] = None,
    show_slide_numbers: bool = False,
    language: Optional[str] = None,
    template: Optional[str] = None,
) -> str:
    """Create a PowerPoint presentation from structured slides and upload it.

    :param slides: Slide models, or dicts in the current or previous spelling
    :param format: "4:3" or "16:9"
    :param file_name: Optional custom filename (without extension)
    :param author: Author name for document properties
    :param footer_text: Optional footer text displayed on all slides
    :param show_slide_numbers: Whether to show slide numbers
    :param language: BCP-47 tag stamped on every run for proofing
    :return: Upload status or URL text
    """
    file_object, warnings = _create_presentation_buffer(
        slides, format,
        author=author,
        footer_text=footer_text,
        show_slide_numbers=show_slide_numbers,
        language=language,
        template=template,
    )

    try:
        text = upload_file(file_object, "pptx", filename=file_name)
    finally:
        file_object.close()

    if warnings:
        logger.info("Presentation created with %d warning(s)", len(warnings))

    logger.info("PowerPoint upload completed")
    return text
