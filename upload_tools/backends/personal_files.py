"""personal-files MCP storage backend.

Pushes generated documents (multipart) to the librechat-personal-files-mcp
``POST /binary`` endpoint, which stores the file in the requesting user's
private storage. Returns the relative path + metadata so the agent can point
back at the file (e.g. via the personal-files ``publish_file`` tool).
"""

import io
import logging
from typing import Optional

import httpx

from config import PersonalFilesSettings

logger = logging.getLogger(__name__)


async def upload_to_personal_files(
    file_object: io.BytesIO,
    filename: str,
    user_context: dict,
    config: PersonalFilesSettings,
    timeout: float = 120.0,
) -> dict:
    """Upload a file to the personal-files server.

    Args:
        file_object: BytesIO object containing the file data.
        filename: Name of the file (with extension).
        user_context: Dict with user info from request headers:
            - user_id: Required (from X-User-Id)
            - user_email: Optional
            - conversation_id: Optional
        config: PersonalFilesSettings with url, service_token, path_prefix.
        timeout: Request timeout in seconds (default 120).

    Returns:
        Dict with file metadata:
        {
            "path": "office/My_Report.docx",
            "filename": "My_Report.docx",
            "mime_type": "application/vnd.openxmlformats-...",
            "size_bytes": 12345,
            "sha256": "...",
        }

    Raises:
        ValueError: If user_context is missing required user_id.
        RuntimeError: If upload fails.
    """
    user_id = user_context.get("user_id")
    if not user_id:
        raise ValueError(
            "Personal-files upload requires user_id in user_context. "
            "Ensure X-User-Id header is passed from LibreChat."
        )

    # The server resolves the path itself; we just ask for prefix/filename.
    relative_path = f"{config.path_prefix}/{filename}"

    headers = {
        "X-Service-Token": config.service_token,
        "X-User-Id": user_id,
    }
    if user_context.get("user_email"):
        headers["X-User-Email"] = user_context["user_email"]
    if user_context.get("conversation_id"):
        headers["X-Conversation-Id"] = user_context["conversation_id"]

    file_object.seek(0)
    file_size = len(file_object.getvalue())
    files = {"file": (filename, file_object)}
    data = {"path": relative_path}

    logger.info(
        "Uploading file to personal-files: %s (%d bytes) for user %s",
        relative_path,
        file_size,
        user_id,
    )

    try:
        async with httpx.AsyncClient(timeout=timeout) as client:
            response = await client.post(
                f"{config.url}/binary",
                headers=headers,
                files=files,
                data=data,
            )
            response.raise_for_status()
            result = response.json()
    except httpx.TimeoutException as e:
        logger.error("personal-files upload timeout: %s", e)
        raise RuntimeError(
            f"personal-files upload timed out after {timeout}s. "
            "Check the personal-files MCP server is reachable."
        ) from e
    except httpx.HTTPStatusError as e:
        detail = _extract_error(e.response)
        logger.error("personal-files upload HTTP %s: %s", e.response.status_code, detail)
        raise RuntimeError(f"personal-files upload failed: {detail}") from e
    except httpx.RequestError as e:
        logger.error("personal-files upload request error: %s", e)
        raise RuntimeError(
            f"personal-files upload failed: {e}. "
            f"Check PERSONAL_FILES_URL ({config.url}) is accessible."
        ) from e

    if not isinstance(result, dict) or "path" not in result:
        raise RuntimeError(
            "personal-files upload succeeded but response is missing 'path'. "
            f"Response: {result}"
        )

    logger.info(
        "File uploaded successfully to personal-files: path=%s, size=%s",
        result["path"],
        result.get("size_bytes"),
    )
    return {
        "path": result["path"],
        "filename": result.get("filename", filename),
        "mime_type": result.get("mime_type"),
        "size_bytes": result.get("size_bytes", file_size),
        "sha256": result.get("sha256"),
    }


def _extract_error(response: httpx.Response) -> str:
    try:
        body = response.json()
        if isinstance(body, dict) and body.get("error"):
            return f"HTTP {response.status_code}: {body['error']}"
    except ValueError:
        pass
    return f"HTTP {response.status_code}: {response.text[:200]}"


def format_personal_files_result(file_info: dict, text_message: Optional[str] = None) -> str:
    """Format personal-files upload as a plain-text path + metadata response.

    Returns a string the LLM can use to reference the stored file.
    """
    path = file_info.get("path", "unknown")
    size = file_info.get("size_bytes", 0)
    message = text_message or f"Document saved to personal files: {path}"

    parts = [
        "saved_to_personal_files: true",
        f"path: {path}",
        f"size_bytes: {size}",
    ]
    if file_info.get("mime_type"):
        parts.append(f"mime_type: {file_info['mime_type']}")
    if file_info.get("sha256"):
        parts.append(f"sha256: {file_info['sha256']}")

    return f"{message}\n" + "\n".join(parts)
