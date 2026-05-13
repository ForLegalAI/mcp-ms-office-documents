import logging
from config import get_config
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
if UPLOAD_STRATEGY == "LOCAL":
    logger.info("Local upload strategy set.")
elif UPLOAD_STRATEGY == "S3":
    logger.info("S3 upload strategy set.")
elif UPLOAD_STRATEGY == "GCS":
    logger.info("GCS upload strategy set.")
elif UPLOAD_STRATEGY == "AZURE":
    logger.info("Azure Blob upload strategy set.")
elif UPLOAD_STRATEGY == "MINIO":
    logger.info("MinIO upload strategy set.")


def upload_file(file_object, suffix: str, filename: str | None = None) -> str:
    """Upload a file to configured backend and return appropriate response.

    :param file_object: File-like object to upload
    :param suffix: File extension (e.g., 'pptx', 'docx', 'xlsx', 'eml')
    :param filename: Optional human-readable filename (without extension). When provided,
        the uploaded object will use this name (sanitized) with a short UUID prefix instead
        of a full UUID.
    :return: Status message with download URL or save location
    :raises RuntimeError: If upload fails for any reason
    """

    try:
        if filename:
            object_name = generate_named_object_name(filename, suffix)
        else:
            object_name = generate_unique_object_name(suffix)
    except Exception as e:
        logger.error("Failed to generate object name for suffix '%s': %s", suffix, e, exc_info=True)
        raise RuntimeError(f"Error preparing upload: {e}") from e

    try:
        if UPLOAD_STRATEGY == "LOCAL":
            result = upload_to_local_folder(file_object, object_name)
        elif UPLOAD_STRATEGY == "S3":
            result = upload_to_s3(file_object, object_name, cfg.storage.s3, SIGNED_URL_EXPIRES_IN)
        elif UPLOAD_STRATEGY == "GCS":
            result = upload_to_gcs(file_object, object_name, cfg.storage.gcs, SIGNED_URL_EXPIRES_IN)
        elif UPLOAD_STRATEGY == "AZURE":
            result = upload_to_azure(file_object, object_name, cfg.storage.azure, SIGNED_URL_EXPIRES_IN)
        elif UPLOAD_STRATEGY == "MINIO":
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
