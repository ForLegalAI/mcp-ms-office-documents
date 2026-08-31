"""Tests for the personal-files MCP storage backend.

Covers PersonalFilesSettings validation, upload_to_personal_files (with a
mocked httpx client) and format_personal_files_result.
"""

import io
import pytest
from unittest.mock import patch, AsyncMock, MagicMock


class TestPersonalFilesSettings:
    """Test PersonalFilesSettings configuration model."""

    def test_settings_valid(self):
        from config import PersonalFilesSettings

        settings = PersonalFilesSettings(
            url="http://librechat-personal-files:8080",
            service_token="s" * 16,
        )
        assert settings.url == "http://librechat-personal-files:8080"
        assert settings.service_token == "s" * 16
        assert settings.path_prefix == "office"

    def test_settings_strips_trailing_slash_and_prefix(self):
        from config import PersonalFilesSettings

        settings = PersonalFilesSettings(
            url="http://host:8080/",
            service_token="s" * 16,
            path_prefix="/docs/",
        )
        assert settings.url == "http://host:8080"
        assert settings.path_prefix == "docs"

    def test_settings_missing_url_raises(self):
        from config import PersonalFilesSettings
        from pydantic import ValidationError

        with pytest.raises(ValidationError):
            PersonalFilesSettings(url="", service_token="s" * 16)

    def test_settings_missing_token_raises(self):
        from config import PersonalFilesSettings
        from pydantic import ValidationError

        with pytest.raises(ValidationError):
            PersonalFilesSettings(url="http://host:8080", service_token="")


class TestUploadToPersonalFiles:
    """Test upload_to_personal_files function."""

    def _settings(self):
        from config import PersonalFilesSettings

        return PersonalFilesSettings(
            url="http://librechat-personal-files:8080",
            service_token="s" * 16,
            path_prefix="office",
        )

    @pytest.mark.asyncio
    async def test_upload_missing_user_id_raises(self):
        from upload_tools.backends.personal_files import upload_to_personal_files

        config = self._settings()
        file_buffer = io.BytesIO(b"test content")
        user_context = {"user_id": None}

        with pytest.raises(ValueError, match="user_id"):
            await upload_to_personal_files(file_buffer, "document.docx", user_context, config)

    @pytest.mark.asyncio
    async def test_upload_success(self):
        from upload_tools.backends.personal_files import upload_to_personal_files

        config = self._settings()
        file_buffer = io.BytesIO(b"test content")
        user_context = {
            "user_id": "user-123",
            "user_email": "test@example.com",
            "conversation_id": "conv-456",
        }

        mock_response = MagicMock()
        mock_response.raise_for_status = MagicMock()
        mock_response.json.return_value = {
            "path": "office/document.docx",
            "filename": "document.docx",
            "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "size_bytes": 12,
            "sha256": "abc123",
        }

        with patch("upload_tools.backends.personal_files.httpx.AsyncClient") as mock_client:
            mock_instance = AsyncMock()
            mock_instance.post = AsyncMock(return_value=mock_response)
            mock_instance.__aenter__ = AsyncMock(return_value=mock_instance)
            mock_instance.__aexit__ = AsyncMock(return_value=None)
            mock_client.return_value = mock_instance

            result = await upload_to_personal_files(
                file_buffer, "document.docx", user_context, config
            )

            assert result["path"] == "office/document.docx"
            assert result["filename"] == "document.docx"
            assert result["size_bytes"] == 12
            assert result["sha256"] == "abc123"

            # Verify request headers included service token + user id
            _, kwargs = mock_instance.post.call_args
            headers = kwargs["headers"]
            assert headers["X-Service-Token"] == "s" * 16
            assert headers["X-User-Id"] == "user-123"
            assert headers["X-User-Email"] == "test@example.com"
            assert kwargs["data"]["path"] == "office/document.docx"
            # URL hits the /binary endpoint
            assert mock_instance.post.call_args.args[0] == "http://librechat-personal-files:8080/binary"

    @pytest.mark.asyncio
    async def test_upload_http_error(self):
        from upload_tools.backends.personal_files import upload_to_personal_files

        config = self._settings()
        file_buffer = io.BytesIO(b"data")
        user_context = {"user_id": "user-123"}

        mock_response = MagicMock()
        mock_response.status_code = 401
        mock_response.json.return_value = {"error": "invalid_service_token"}
        mock_response.text = '{"error": "invalid_service_token"}'
        from httpx import HTTPStatusError

        mock_response.raise_for_status.side_effect = HTTPStatusError(
            "401", request=MagicMock(), response=mock_response
        )

        with patch("upload_tools.backends.personal_files.httpx.AsyncClient") as mock_client:
            mock_instance = AsyncMock()
            mock_instance.post = AsyncMock(return_value=mock_response)
            mock_instance.__aenter__ = AsyncMock(return_value=mock_instance)
            mock_instance.__aexit__ = AsyncMock(return_value=None)
            mock_client.return_value = mock_instance

            with pytest.raises(RuntimeError, match="invalid_service_token"):
                await upload_to_personal_files(file_buffer, "document.docx", user_context, config)


class TestFormatPersonalFilesResult:
    """Test format_personal_files_result."""

    def test_formats_path_and_metadata(self):
        from upload_tools.backends.personal_files import format_personal_files_result

        file_info = {
            "path": "office/document.docx",
            "filename": "document.docx",
            "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "size_bytes": 12345,
            "sha256": "abc",
        }

        result = format_personal_files_result(file_info, "Document saved.")
        assert "Document saved." in result
        assert "path: office/document.docx" in result
        assert "size_bytes: 12345" in result
        assert "mime_type: application/vnd.openxmlformats" in result
        assert "sha256: abc" in result

    def test_formats_minimal_info(self):
        from upload_tools.backends.personal_files import format_personal_files_result

        result = format_personal_files_result({"path": "office/x.docx", "size_bytes": 3})
        assert "saved_to_personal_files: true" in result
        assert "path: office/x.docx" in result
