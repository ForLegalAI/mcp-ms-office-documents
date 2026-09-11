"""Local filesystem upload backend.

Writes the file into ``./output`` under the process working directory
(``/app/output`` in the container, mounted by docker-compose). Failure
raises, like the remote backends returning ``None``: the dispatcher in
``upload_tools/main.py`` treats any string as a successful result, so
returning an error message here used to reach the MCP client as success.
"""
import os
import logging

logger = logging.getLogger(__name__)


def upload_to_local_folder(file_object, file_name: str) -> str:
    """Save *file_object* as ``./output/<file_name>`` and return a status message.

    Raises:
        RuntimeError: If the directory cannot be created or the file written.
    """
    save_dir = os.path.join(os.getcwd(), "output")
    save_path = os.path.join(save_dir, file_name)

    try:
        os.makedirs(save_dir, exist_ok=True)
        file_object.seek(0)
        with open(save_path, 'wb') as f:
            f.write(file_object.read())
    except Exception as e:
        logger.exception("Failed to save file locally to %s", save_path)
        raise RuntimeError(f"Error saving document locally: {e}") from e

    logger.info("Saved file to %s", save_path)
    return f"Document saved to {save_path}"
