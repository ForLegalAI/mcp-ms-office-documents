import logging
from typing import List, Dict, Any

from upload_tools import upload_file
from .slide_builder import PowerpointPresentation

logger = logging.getLogger(__name__)


def create_presentation(slides: List[Dict[str, Any]], format: str = "4:3", file_name: str | None = None) -> str:
    """Create a PowerPoint presentation from structured slides and upload it.

    :param slides: List of slide dicts with keys based on slide_type
    :param format: "4:3" or "16:9"
    :return: Upload status or URL text
    :raises: Exception on failure (propagated to caller)
    """
    try:
        # Validate input
        if not slides:
            raise ValueError("No slides provided")

        logger.info(f"Starting create_presentation: slides={len(slides)}, format={format}")

        # Create presentation
        presentation = PowerpointPresentation(slides, format)

        # Save presentation
        file_object = presentation.save()

        # Upload presentation
        text = upload_file(file_object, "pptx", filename=file_name)
        file_object.close()

        logger.info("PowerPoint upload completed")
        # Return presentation link
        return text

    except Exception as e:
        logger.error(f"Failed to create presentation: {e}")
        # Re-raise so the MCP tool wrapper can return a proper error
        raise

