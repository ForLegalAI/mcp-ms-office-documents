"""LibreChat integration helpers for MCP tools.

This module provides utilities for extracting user context from request headers
and formatting responses as MCP file artifacts for LibreChat integration.
"""

import io
import logging
from typing import Optional, Union

from async_runner import run_blocking
from upload_tools import upload_file_async, is_librechat_strategy
from config import get_config, StorageStrategy

# NOTE: upload_tools.backends.librechat and upload_tools.backends.personal_files
# are imported lazily inside the strategy branches below, not here. They pull
# in httpx, and this module is imported by main.py on every startup — an eager
# import would make an optional backend's dependency mandatory for all
# strategies. The other backends follow the same rule (see the boto3 comment in
# upload_tools/backends/s3.py).

logger = logging.getLogger(__name__)


def _is_personal_files_strategy() -> bool:
    cfg = get_config()
    return cfg.storage.strategy == StorageStrategy.PERSONAL_FILES


def extract_user_context_from_request() -> dict:
    """Extract user context from the current HTTP request.

    LibreChat sends user information via HTTP headers:
    - X-User-Id: Required - LibreChat user ID
    - X-User-Email: Optional - User email
    - X-Conversation-Id: Optional - Current conversation ID

    Uses FastMCP's get_http_request() to access the current request.

    TRUST BOUNDARY NOTE:
    These headers are trusted verbatim without validation. This is intentional
    because LibreChat validates/authorizes the user before calling this MCP server.
    The MCP endpoint's API key auth (ApiKeyAuthMiddleware) ensures only LibreChat
    can reach this endpoint. LibreChat itself handles user authentication upstream,
    so by the time requests reach here, the user identity has already been verified.

    Returns:
        Dict with user_id, user_email, and conversation_id (may be None)
    """
    try:
        from fastmcp.server.dependencies import get_http_request
        request = get_http_request()
        
        # Access headers from Starlette Request object
        headers = request.headers if request else {}
        
        return {
            "user_id": headers.get("x-user-id"),
            "user_email": headers.get("x-user-email"),
            "conversation_id": headers.get("x-conversation-id"),
        }
    except RuntimeError as e:
        # No active HTTP request found
        logger.warning("Could not extract user context: %s", e)
        return {
            "user_id": None,
            "user_email": None,
            "conversation_id": None,
        }
    except ImportError:
        logger.warning("FastMCP dependencies not available")
        return {
            "user_id": None,
            "user_email": None,
            "conversation_id": None,
        }


async def upload_and_format_response(
    file_buffer: io.BytesIO,
    suffix: str,
    filename: Optional[str],
    user_context: dict,
    success_message: str,
    add_unique_prefix: bool | None = None,
) -> Union[str, dict]:
    """Upload a file and format the response appropriately.

    For LIBRECHAT strategy, uploads to LibreChat and returns file artifact format.
    For other strategies, returns the URL string.

    Args:
        file_buffer: BytesIO containing the file data
        suffix: File extension without dot (e.g., 'docx', 'xlsx')
        filename: Human-readable filename (without extension) or None
        user_context: Dict from extract_user_context()
        success_message: Message to include in file artifact response
        add_unique_prefix: If True, adds UUID prefix. If None, uses strategy-based default
            (True for traditional backends, False for LIBRECHAT).

    Returns:
        str (URL) for traditional backends, dict (file artifact) for LIBRECHAT,
        or str (path + metadata) for PERSONAL_FILES.
    """
    if is_librechat_strategy():
        from upload_tools.backends.librechat import format_file_artifact

        # Validate user context for LIBRECHAT
        if not user_context.get("user_id"):
            raise ValueError(
                "LibreChat upload requires X-User-Id header. "
                "Ensure LibreChat is configured to send user headers to MCP server."
            )

        # Upload to LibreChat
        file_info = await upload_file_async(
            file_buffer,
            suffix,
            filename=filename,
            user_context=user_context,
            add_unique_prefix=add_unique_prefix,
        )

        # Format as MCP file artifact
        return format_file_artifact(file_info, success_message)
    elif _is_personal_files_strategy():
        from upload_tools.backends.personal_files import format_personal_files_result

        # Validate user context for PERSONAL_FILES
        if not user_context.get("user_id"):
            raise ValueError(
                "Personal-files upload requires X-User-Id header. "
                "Ensure LibreChat is configured to send user headers to MCP server."
            )

        file_info = await upload_file_async(
            file_buffer,
            suffix,
            filename=filename,
            user_context=user_context,
            add_unique_prefix=add_unique_prefix,
        )

        # Format as path + metadata string (the agent can publish_file later).
        return format_personal_files_result(file_info, success_message)
    else:
        # Traditional upload - returns URL string.
        #
        # upload_file() is synchronous and the backends behind it block: boto3
        # /GCS/Azure do a network round-trip plus a signed-URL call, LOCAL hits
        # the disk. Calling it inline here would freeze the event loop for that
        # whole duration — no other request served, health probes included, the
        # exact EKS failure mode async_runner.py exists to prevent. Dispatch it
        # the same way the document build itself is dispatched.
        #
        # The dynamic template tools do NOT come through here: they call
        # upload_file() from inside their own run_blocking'd body, so they are
        # already off the loop and must not be double-dispatched.
        from upload_tools import upload_file
        return await run_blocking(
            upload_file,
            file_buffer,
            suffix,
            filename=filename,
            add_unique_prefix=add_unique_prefix,
        )
