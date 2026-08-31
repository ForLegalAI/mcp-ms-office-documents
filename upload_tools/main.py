"""Unified file upload module supporting multiple storage backends.

This module provides a centralized upload interface that dispatches to the
configured storage backend (LOCAL, S3, GCS, AZURE, MINIO, LIBRECHAT, or
PERSONAL_FILES).

For LIBRECHAT strategy, uploads go to LibreChat's service endpoint and return
file metadata for MCP file artifacts instead of URLs.
For PERSONAL_FILES strategy, uploads go to the personal-files MCP server's
``POST /binary`` endpoint and return path + metadata.
"""

import logging
from typing import Union

from config import get_config, StorageStrategy
from .utils import generate_unique_object_name, generate_named_object_name
from .backends.local import upload_to_local_folder
from .backends.s3 import upload_to_s3
from .backends.gcs import upload_to_gcs
from .backends.azure import upload_to_azure
from .backends.minio import upload_to_minio

logger = logging.getLogger(__name__)

# Load centralized configuration
cfg = get_config()

# Convenience aliases
UPLOAD_STRATEGY = cfg.storage.strategy
SIGNED_URL_EXPIRES_IN = cfg.storage.signed_url_expires_in

# Strategy announcement logs
if UPLOAD_STRATEGY == StorageStrategy.LOCAL:
    logger.info("Local upload strategy set.")
elif UPLOAD_STRATEGY == StorageStrategy.S3:
    logger.info("S3 upload strategy set.")
elif UPLOAD_STRATEGY == StorageStrategy.GCS:
    logger.info("GCS upload strategy set.")
elif UPLOAD_STRATEGY == StorageStrategy.AZURE:
    logger.info("Azure Blob upload strategy set.")
elif UPLOAD_STRATEGY == StorageStrategy.MINIO:
    logger.info("MinIO upload strategy set.")
elif UPLOAD_STRATEGY == StorageStrategy.LIBRECHAT:
    logger.info("LibreChat upload strategy set.")
elif UPLOAD_STRATEGY == StorageStrategy.PERSONAL_FILES:
    logger.info("Personal-files upload strategy set.")


def upload_file(
    file_object,
    suffix: str,
    filename: str | None = None,
    user_context: dict | None = None,
    add_unique_prefix: bool | None = None,
) -> Union[str, dict]:
    """Upload a file to configured backend and return appropriate response.

    For traditional backends (LOCAL, S3, GCS, AZURE, MINIO), returns a URL string.
    For LIBRECHAT and PERSONAL_FILES, this function raises an error - use
    upload_file_async instead.

    :param file_object: File-like object to upload
    :param suffix: File extension (e.g., 'pptx', 'docx', 'xlsx', 'eml')
    :param filename: Optional human-readable filename (without extension). When provided,
        the uploaded object will use this name (sanitized).
    :param user_context: Optional dict with user info:
        - user_id: Required for LIBRECHAT and PERSONAL_FILES
        - user_email: Optional
        - conversation_id: Optional
    :param add_unique_prefix: If True, adds 8-char UUID prefix to filename for uniqueness.
        If None (default), uses True for traditional backends (to prevent collisions) and
        False for LIBRECHAT (which handles uniqueness with its own UUID prefix).
    :return: Status message with download URL or save location (str for traditional backends)
    :raises RuntimeError: If upload fails or LIBRECHAT/PERSONAL_FILES strategy used without async
    """
    # Resolve default based on strategy: traditional backends default to True for collision safety,
    # LIBRECHAT and PERSONAL_FILES default to False.
    if add_unique_prefix is None:
        add_unique_prefix = UPLOAD_STRATEGY not in (
            StorageStrategy.LIBRECHAT,
            StorageStrategy.PERSONAL_FILES,
        )
    # LIBRECHAT / PERSONAL_FILES require async - direct callers to upload_file_async
    if UPLOAD_STRATEGY in (StorageStrategy.LIBRECHAT, StorageStrategy.PERSONAL_FILES):
        raise RuntimeError(
            "LIBRECHAT/PERSONAL_FILES strategies require async upload. "
            "Use upload_file_async() instead."
        )

    try:
        if filename:
            object_name = generate_named_object_name(filename, suffix, add_unique_prefix)
        else:
            object_name = generate_unique_object_name(suffix)
    except Exception as e:
        logger.error("Failed to generate object name for suffix '%s': %s", suffix, e, exc_info=True)
        raise RuntimeError(f"Error preparing upload: {e}") from e

    try:
        if UPLOAD_STRATEGY == StorageStrategy.LOCAL:
            result = upload_to_local_folder(file_object, object_name)
        elif UPLOAD_STRATEGY == StorageStrategy.S3:
            result = upload_to_s3(file_object, object_name, cfg.storage.s3, SIGNED_URL_EXPIRES_IN)
        elif UPLOAD_STRATEGY == StorageStrategy.GCS:
            result = upload_to_gcs(file_object, object_name, cfg.storage.gcs, SIGNED_URL_EXPIRES_IN)
        elif UPLOAD_STRATEGY == StorageStrategy.AZURE:
            result = upload_to_azure(file_object, object_name, cfg.storage.azure, SIGNED_URL_EXPIRES_IN)
        elif UPLOAD_STRATEGY == StorageStrategy.MINIO:
            result = upload_to_minio(file_object, object_name, cfg.storage.minio, SIGNED_URL_EXPIRES_IN)
        else:
            logger.error("No upload strategy configured (UPLOAD_STRATEGY='%s')", UPLOAD_STRATEGY)
            raise RuntimeError("No upload strategy set, document cannot be created.")
    except RuntimeError:
        raise
    except Exception as e:
        logger.error("Upload failed (strategy=%s): %s", UPLOAD_STRATEGY, e, exc_info=True)
        raise RuntimeError(f"Error uploading document: {e}") from e

    if result is None:
        logger.error("Upload backend '%s' returned None for %s – check backend logs for details", UPLOAD_STRATEGY, object_name)
        raise RuntimeError(f"Upload to {UPLOAD_STRATEGY} failed. Check server logs for details.")

    return result


async def upload_file_async(
    file_object,
    suffix: str,
    filename: str | None = None,
    user_context: dict | None = None,
    add_unique_prefix: bool | None = None,
) -> Union[str, dict]:
    """Async upload a file to configured backend.

    This function supports all backends. For LIBRECHAT, it returns a dict with
    file metadata for MCP file artifacts. For other backends, it returns a URL string.

    :param file_object: File-like object to upload
    :param suffix: File extension (e.g., 'pptx', 'docx', 'xlsx', 'eml')
    :param filename: Optional human-readable filename (without extension). When provided,
        the uploaded object will use this name (sanitized).
    :param user_context: Dict with user info (required for LIBRECHAT and PERSONAL_FILES):
        - user_id: Required
        - user_email: Optional
        - conversation_id: Optional
    :param add_unique_prefix: If True, adds 8-char UUID prefix to filename for uniqueness.
        If None (default), uses True for traditional backends (to prevent collisions) and
        False for LIBRECHAT/PERSONAL_FILES.
    :return: URL string (traditional backends), dict (LIBRECHAT artifact), or
        path+metadata dict (PERSONAL_FILES)
    :raises RuntimeError: If upload fails
    :raises ValueError: If LIBRECHAT/PERSONAL_FILES used without user_context
    """
    # Resolve default based on strategy: traditional backends default to True for collision safety,
    # LIBRECHAT and PERSONAL_FILES default to False.
    if add_unique_prefix is None:
        add_unique_prefix = UPLOAD_STRATEGY not in (
            StorageStrategy.LIBRECHAT,
            StorageStrategy.PERSONAL_FILES,
        )

    # Generate object name
    try:
        if filename:
            object_name = generate_named_object_name(filename, suffix, add_unique_prefix)
        else:
            object_name = generate_unique_object_name(suffix)
    except Exception as e:
        logger.error("Failed to generate object name for suffix '%s': %s", suffix, e, exc_info=True)
        raise RuntimeError(f"Error preparing upload: {e}") from e

    # Handle LIBRECHAT strategy (async)
    if UPLOAD_STRATEGY == StorageStrategy.LIBRECHAT:
        from .backends.librechat import upload_to_librechat

        if not user_context:
            raise ValueError(
                "LIBRECHAT strategy requires user_context with user_id. "
                "Ensure X-User-Id header is passed from LibreChat."
            )

        try:
            result = await upload_to_librechat(
                file_object,
                object_name,
                user_context,
                cfg.storage.librechat,
            )
            return result
        except (ValueError, RuntimeError):
            raise
        except Exception as e:
            logger.error("LibreChat upload failed: %s", e, exc_info=True)
            raise RuntimeError(f"Error uploading to LibreChat: {e}") from e

    # Handle PERSONAL_FILES strategy (async)
    if UPLOAD_STRATEGY == StorageStrategy.PERSONAL_FILES:
        from .backends.personal_files import upload_to_personal_files

        if not user_context:
            raise ValueError(
                "PERSONAL_FILES strategy requires user_context with user_id. "
                "Ensure X-User-Id header is passed from LibreChat."
            )

        try:
            result = await upload_to_personal_files(
                file_object,
                object_name,
                user_context,
                cfg.storage.personal_files,
            )
            return result
        except (ValueError, RuntimeError):
            raise
        except Exception as e:
            logger.error("personal-files upload failed: %s", e, exc_info=True)
            raise RuntimeError(f"Error uploading to personal-files: {e}") from e

    # For non-LIBRECHAT strategies, use sync upload
    # (They don't need user_context and are not truly async)
    try:
        if UPLOAD_STRATEGY == StorageStrategy.LOCAL:
            result = upload_to_local_folder(file_object, object_name)
        elif UPLOAD_STRATEGY == StorageStrategy.S3:
            result = upload_to_s3(file_object, object_name, cfg.storage.s3, SIGNED_URL_EXPIRES_IN)
        elif UPLOAD_STRATEGY == StorageStrategy.GCS:
            result = upload_to_gcs(file_object, object_name, cfg.storage.gcs, SIGNED_URL_EXPIRES_IN)
        elif UPLOAD_STRATEGY == StorageStrategy.AZURE:
            result = upload_to_azure(file_object, object_name, cfg.storage.azure, SIGNED_URL_EXPIRES_IN)
        elif UPLOAD_STRATEGY == StorageStrategy.MINIO:
            result = upload_to_minio(file_object, object_name, cfg.storage.minio, SIGNED_URL_EXPIRES_IN)
        else:
            logger.error("No upload strategy configured (UPLOAD_STRATEGY='%s')", UPLOAD_STRATEGY)
            raise RuntimeError("No upload strategy set, document cannot be created.")
    except RuntimeError:
        raise
    except Exception as e:
        logger.error("Upload failed (strategy=%s): %s", UPLOAD_STRATEGY, e, exc_info=True)
        raise RuntimeError(f"Error uploading document: {e}") from e

    if result is None:
        logger.error("Upload backend '%s' returned None for %s – check backend logs for details", UPLOAD_STRATEGY, object_name)
        raise RuntimeError(f"Upload to {UPLOAD_STRATEGY} failed. Check server logs for details.")

    return result


def is_librechat_strategy() -> bool:
    """Check if current upload strategy is LIBRECHAT.

    Useful for callers to determine if they need to handle file artifacts
    differently in their response format.
    """
    return UPLOAD_STRATEGY == StorageStrategy.LIBRECHAT
