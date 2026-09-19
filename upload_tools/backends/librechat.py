"""LibreChat file service upload backend.

This module handles uploading generated documents to LibreChat's service
endpoint, which stores files and associates them with conversations.
The uploaded files appear as attachments in the chat UI.
"""

import io
import logging
import mimetypes
from typing import Optional

import httpx

from config import LibreChatSettings
# One MIME table for the whole upload layer; re-exported here because this
# module's name for it is part of its published surface.
from ..utils import MIME_TYPES, file_extension

logger = logging.getLogger(__name__)


def get_mime_type(filename: str) -> str:
    """Get MIME type for a filename.

    Falls back to the ``mimetypes`` module, and then to a generic binary type,
    rather than raising: LibreChat receives whatever a caller attaches, and an
    attachment typed as a guess still reaches the conversation.

    Args:
        filename: Filename with extension

    Returns:
        MIME type string
    """
    ext = file_extension(filename)
    if ext in MIME_TYPES:
        return MIME_TYPES[ext]
    # Fallback to mimetypes module
    mime_type, _ = mimetypes.guess_type(filename)
    return mime_type or "application/octet-stream"


async def upload_to_librechat(
    file_object: io.BytesIO,
    filename: str,
    user_context: dict,
    config: LibreChatSettings,
    timeout: float = 60.0,
) -> dict:
    """Upload a file to LibreChat's service endpoint.

    This function uploads a file to LibreChat and returns metadata that can
    be used to construct an MCP file artifact response.

    Args:
        file_object: BytesIO object containing the file data
        filename: Name of the file (with extension)
        user_context: Dict containing user information from request headers:
            - user_id: Required - LibreChat user ID (from X-User-Id header)
            - user_email: Optional - User email (from X-User-Email header)
            - conversation_id: Optional - Conversation ID (from X-Conversation-Id header)
        config: LibreChatSettings with service_url and service_token
        timeout: Request timeout in seconds (default 60)

    Returns:
        Dict with file metadata:
        {
            "file_id": "uuid-string",
            "filename": "document.docx",
            "filepath": "/path/to/file",
            "type": "application/vnd.openxmlformats-...",
            "bytes": 12345,
            "source": "local",
            "download_url": "/api/files/download/{user_id}/{file_id}"
        }

    Raises:
        ValueError: If user_context is missing required user_id
        RuntimeError: If upload fails
    """
    # Validate user context
    user_id = user_context.get("user_id")
    if not user_id:
        raise ValueError(
            "LibreChat upload requires user_id in user_context. "
            "Ensure X-User-Id header is passed from LibreChat."
        )

    # Prepare headers
    headers = {
        "X-Service-Token": config.service_token,
        "X-User-Id": user_id,
    }

    # Add optional headers
    if user_context.get("user_email"):
        headers["X-User-Email"] = user_context["user_email"]
    if user_context.get("conversation_id"):
        headers["X-Conversation-Id"] = user_context["conversation_id"]

    # Determine MIME type
    mime_type = get_mime_type(filename)

    logger.info(
        "Uploading file to LibreChat: %s (%s) for user %s",
        filename,
        mime_type,
        user_id,
    )

    # Ensure file object is at the beginning
    file_object.seek(0)
    file_size = len(file_object.getvalue())

    try:
        async with httpx.AsyncClient(timeout=timeout) as client:
            # Prepare multipart form data
            files = {
                "file": (filename, file_object, mime_type),
            }
            data = {
                "filename": filename,
            }

            response = await client.post(
                config.service_url,
                headers=headers,
                files=files,
                data=data,
            )

            # Check for errors
            if response.status_code == 401:
                raise RuntimeError(
                    "LibreChat upload failed: Invalid service token. "
                    "Verify LIBRECHAT_SERVICE_TOKEN matches MCP_SERVICE_TOKEN in LibreChat."
                )
            elif response.status_code == 400:
                try:
                    error_detail = response.json().get("message", response.text)
                except (ValueError, KeyError):
                    error_detail = response.text
                raise RuntimeError(f"LibreChat upload failed: {error_detail}")

            response.raise_for_status()

            result = response.json()

    except httpx.TimeoutException as e:
        logger.error("LibreChat upload timeout: %s", e)
        raise RuntimeError(
            f"LibreChat upload timed out after {timeout}s. "
            "Try increasing timeout or check network connectivity."
        ) from e
    except httpx.HTTPStatusError as e:
        logger.error("LibreChat upload HTTP error: %s", e)
        raise RuntimeError(f"LibreChat upload failed: {e}") from e
    except httpx.RequestError as e:
        logger.error("LibreChat upload request error: %s", e)
        raise RuntimeError(
            f"LibreChat upload failed: {e}. "
            f"Check LIBRECHAT_SERVICE_URL ({config.service_url}) is accessible."
        ) from e

    # Extract file info from response
    if not result.get("success", False):
        error_msg = result.get("error") or result.get("message") or "Unknown error"
        raise RuntimeError(f"LibreChat upload failed: {error_msg}")

    file_info = result.get("file", {})
    if not file_info.get("file_id"):
        raise RuntimeError(
            "LibreChat upload succeeded but response missing file_id. "
            f"Response: {result}"
        )

    logger.info(
        "File uploaded successfully to LibreChat: file_id=%s, filename=%s",
        file_info.get("file_id"),
        file_info.get("filename", filename),
    )

    # Construct download URL for LLM to use in responses
    file_id = file_info["file_id"]
    download_url = f"/api/files/download/{user_id}/{file_id}"

    # Return normalized file metadata
    return {
        "file_id": file_id,
        "filename": file_info.get("filename", filename),
        "filepath": file_info.get("filepath"),
        "type": file_info.get("type", mime_type),
        "bytes": file_info.get("bytes", file_size),
        "source": file_info.get("source", "local"),
        "download_url": download_url,
    }


def format_file_artifact(file_info: dict, text_message: Optional[str] = None) -> dict:
    """Format file info as a simple dict that FastMCP will pass through correctly.

    Returns a dict with 'result' property containing file info. This format
    is compatible with FastMCP's output schema validation and LibreChat's
    MCP client expectations.

    Args:
        file_info: Dict returned by upload_to_librechat()
        text_message: Optional text message to include with the file

    Returns:
        Dict with result containing message and file info
    """
    file_id = file_info.get("file_id", "unknown")
    filename = file_info.get("filename", "document")
    filepath = file_info.get("filepath", f"/uploads/{file_id}")
    mime_type = file_info.get("type", "application/octet-stream")
    file_bytes = file_info.get("bytes", 0)
    source = file_info.get("source", "local")

    text = text_message or f"File '{filename}' created successfully."

    # Return dict with 'result' property for schema validation
    return {
        "result": {
            "message": text,
            "file": {
                "file_id": file_id,
                "filename": filename,
                "filepath": filepath,
                "mimeType": mime_type,
                "bytes": file_bytes,
                "source": source,
            }
        }
    }
