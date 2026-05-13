"""XML file creation tool.

This module provides functionality to validate XML content and save it as a file.

Uses defusedxml for safe XML parsing to prevent XXE and entity expansion attacks.
"""

import io
import logging
import re
from typing import Tuple

import defusedxml.ElementTree as DefusedET
from defusedxml import DefusedXmlException

from upload_tools import upload_file

logger = logging.getLogger(__name__)


class XMLValidationError(Exception):
    """Raised when XML content is invalid or incomplete."""
    pass


class XMLFileCreationError(Exception):
    """Raised when XML file creation or upload fails."""
    pass


def validate_xml(xml_content: str) -> Tuple[bool, str]:
    """Validate that the provided string is well-formed XML.

    Uses defusedxml to safely parse XML content, protecting against:
    - XML External Entity (XXE) attacks
    - Entity expansion attacks (e.g., "billion laughs")
    - External DTD fetching

    Args:
        xml_content: The XML content to validate.

    Returns:
        A tuple of (is_valid, error_message). If valid, error_message is empty.
    """
    try:
        # Try to parse the XML content using defusedxml for safety
        DefusedET.fromstring(xml_content)
        return True, ""
    except DefusedET.ParseError as e:
        return False, f"XML parsing error: {str(e)}"
    except DefusedXmlException as e:
        return False, f"XML security error (potentially malicious content): {str(e)}"
    except Exception as e:
        return False, f"Unexpected error during XML validation: {str(e)}"


def create_xml_file(xml_content: str, file_name: str | None = None) -> str:
    """Create an XML file from the provided XML content.

    Validates that the content is well-formed XML before saving.

    Args:
        xml_content: Complete, valid XML content string.

    Returns:
        A status message with the download URL or file path.

    Raises:
        XMLValidationError: If the XML content is invalid or contains security threats.
        XMLFileCreationError: If file creation or upload fails.
    """
    logger.info("Starting XML file creation")

    # Strip leading/trailing whitespace
    xml_content = xml_content.strip()

    # Validate the XML content
    is_valid, error_message = validate_xml(xml_content)
    if not is_valid:
        logger.error(f"XML validation failed: {error_message}")
        raise XMLValidationError(error_message)

    logger.debug("XML content validated successfully")

    # Extract encoding from XML declaration if present, default to UTF-8
    encoding = "utf-8"
    if xml_content.startswith('<?xml'):
        match = re.search(r'encoding=["\']([^"\']+)["\']', xml_content)
        if match:
            encoding = match.group(1)
            logger.debug(f"Using encoding from XML declaration: {encoding}")
    else:
        # Add UTF-8 declaration if no declaration present
        xml_content = '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_content
        logger.debug("Added XML declaration with UTF-8 encoding")

    try:
        # Create a file-like object from the XML content using declared encoding
        xml_bytes = xml_content.encode(encoding)
        file_object = io.BytesIO(xml_bytes)

        try:
            # Upload the file
            result = upload_file(file_object, "xml", filename=file_name)
            logger.info("XML file uploaded successfully")
            return result
        finally:
            file_object.close()

    except (XMLValidationError, XMLFileCreationError):
        # Re-raise our custom exceptions as-is
        raise
    except Exception as e:
        logger.error(f"Error creating XML file: {str(e)}", exc_info=True)
        raise XMLFileCreationError(f"Error creating XML file: {str(e)}") from e

