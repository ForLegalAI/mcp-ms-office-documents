"""Tests for LibreChat file artifacts integration.

These tests verify the LIBRECHAT upload strategy and file artifact formatting.
"""

import io
import pytest
from unittest.mock import patch, MagicMock, AsyncMock


class TestLibreChatSettings:
    """Test LibreChatSettings configuration model."""

    def test_librechat_settings_valid(self):
        """Test valid LibreChat settings."""
        from config import LibreChatSettings

        settings = LibreChatSettings(
            service_url="http://api:3080/api/service/files",
            service_token="test_token_123",
        )
        assert settings.service_url == "http://api:3080/api/service/files"
        assert settings.service_token == "test_token_123"

    def test_librechat_settings_strips_whitespace(self):
        """Test that LibreChat settings strips whitespace."""
        from config import LibreChatSettings

        settings = LibreChatSettings(
            service_url="  http://api:3080/api/service/files/  ",
            service_token="  test_token  ",
        )
        assert settings.service_url == "http://api:3080/api/service/files"
        assert settings.service_token == "test_token"

    def test_librechat_settings_missing_url_raises(self):
        """Test that missing service_url raises ValueError."""
        from config import LibreChatSettings
        from pydantic import ValidationError

        with pytest.raises(ValidationError):
            LibreChatSettings(
                service_url="",
                service_token="test_token",
            )

    def test_librechat_settings_missing_token_raises(self):
        """Test that missing service_token raises ValueError."""
        from config import LibreChatSettings
        from pydantic import ValidationError

        with pytest.raises(ValidationError):
            LibreChatSettings(
                service_url="http://api:3080/api/service/files",
                service_token="",
            )


class TestStorageStrategy:
    """Test StorageStrategy enum includes LIBRECHAT."""

    def test_librechat_in_storage_strategy(self):
        """Test LIBRECHAT is a valid storage strategy."""
        from config import StorageStrategy

        assert hasattr(StorageStrategy, "LIBRECHAT")
        assert StorageStrategy.LIBRECHAT.value == "LIBRECHAT"

    def test_personal_files_in_storage_strategy(self):
        """Test PERSONAL_FILES is a valid storage strategy."""
        from config import StorageStrategy

        assert hasattr(StorageStrategy, "PERSONAL_FILES")
        assert StorageStrategy.PERSONAL_FILES.value == "PERSONAL_FILES"

    def test_all_strategies_present(self):
        """Test all expected strategies are present."""
        from config import StorageStrategy

        expected = {"LOCAL", "S3", "GCS", "AZURE", "MINIO", "LIBRECHAT", "PERSONAL_FILES"}
        actual = {s.value for s in StorageStrategy}
        assert expected == actual


class TestLibreChatBackend:
    """Test LibreChat upload backend functions."""

    def test_get_mime_type_docx(self):
        """Test MIME type for docx."""
        from upload_tools.backends.librechat import get_mime_type

        assert get_mime_type("document.docx") == \
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    def test_get_mime_type_xlsx(self):
        """Test MIME type for xlsx."""
        from upload_tools.backends.librechat import get_mime_type

        assert get_mime_type("spreadsheet.xlsx") == \
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    def test_get_mime_type_pptx(self):
        """Test MIME type for pptx."""
        from upload_tools.backends.librechat import get_mime_type

        assert get_mime_type("presentation.pptx") == \
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    def test_get_mime_type_eml(self):
        """Test MIME type for eml."""
        from upload_tools.backends.librechat import get_mime_type

        assert get_mime_type("email.eml") == "message/rfc822"

    def test_get_mime_type_xml(self):
        """Test MIME type for xml."""
        from upload_tools.backends.librechat import get_mime_type

        assert get_mime_type("data.xml") == "application/xml"

    def test_get_mime_type_unknown_fallback(self):
        """Test MIME type fallback for unknown extension."""
        from upload_tools.backends.librechat import get_mime_type

        result = get_mime_type("file.unknown123")
        assert result == "application/octet-stream"

    def test_format_file_artifact_with_text(self):
        """Test format_file_artifact with text message returns dict with result property."""
        from upload_tools.backends.librechat import format_file_artifact

        file_info = {
            "file_id": "test-file-id-123",
            "filename": "document.docx",
            "filepath": "/uploads/document.docx",
            "type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "bytes": 12345,
            "source": "local",
            "download_url": "/api/files/download/user-123/test-file-id-123",
        }

        result = format_file_artifact(file_info, "Document created successfully.")

        # Result should be dict with result property
        assert isinstance(result, dict)
        assert "result" in result

        # Check result.message
        assert result["result"]["message"] == "Document created successfully."

        # Check result.file
        file_data = result["result"]["file"]
        assert file_data["file_id"] == "test-file-id-123"
        assert file_data["filename"] == "document.docx"
        assert file_data["filepath"] == "/uploads/document.docx"
        assert file_data["mimeType"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        assert file_data["bytes"] == 12345
        assert file_data["source"] == "local"

    def test_format_file_artifact_without_text(self):
        """Test format_file_artifact without text message uses default message."""
        from upload_tools.backends.librechat import format_file_artifact

        file_info = {
            "file_id": "test-file-id-456",
            "filename": "data.xlsx",
            "type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }

        result = format_file_artifact(file_info, None)

        # Result should be dict with result property
        assert isinstance(result, dict)
        assert "result" in result

        # Check result.message has default text including filename
        assert "data.xlsx" in result["result"]["message"]

        # Check result.file
        file_data = result["result"]["file"]
        assert file_data["file_id"] == "test-file-id-456"
        assert file_data["filename"] == "data.xlsx"


class TestUploadToLibreChat:
    """Test upload_to_librechat function."""

    @pytest.mark.asyncio
    async def test_upload_missing_user_id_raises(self):
        """Test that missing user_id raises ValueError."""
        from upload_tools.backends.librechat import upload_to_librechat
        from config import LibreChatSettings

        config = LibreChatSettings(
            service_url="http://api:3080/api/service/files",
            service_token="test_token",
        )

        file_buffer = io.BytesIO(b"test content")
        user_context = {"user_id": None}

        with pytest.raises(ValueError, match="user_id"):
            await upload_to_librechat(
                file_buffer,
                "document.docx",
                user_context,
                config,
            )

    @pytest.mark.asyncio
    async def test_upload_success(self):
        """Test successful upload to LibreChat."""
        from upload_tools.backends.librechat import upload_to_librechat
        from config import LibreChatSettings

        config = LibreChatSettings(
            service_url="http://api:3080/api/service/files",
            service_token="test_token",
        )

        file_buffer = io.BytesIO(b"test content")
        user_context = {
            "user_id": "user-123",
            "user_email": "test@example.com",
            "conversation_id": "conv-456",
        }

        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "success": True,
            "file": {
                "file_id": "uploaded-file-id",
                "filename": "document.docx",
                "filepath": "/uploads/document.docx",
                "type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "bytes": 12,
            },
        }
        mock_response.raise_for_status = MagicMock()

        with patch("upload_tools.backends.librechat.httpx.AsyncClient") as mock_client:
            mock_instance = AsyncMock()
            mock_instance.post = AsyncMock(return_value=mock_response)
            mock_instance.__aenter__ = AsyncMock(return_value=mock_instance)
            mock_instance.__aexit__ = AsyncMock(return_value=None)
            mock_client.return_value = mock_instance

            result = await upload_to_librechat(
                file_buffer,
                "document.docx",
                user_context,
                config,
            )

            assert result["file_id"] == "uploaded-file-id"
            assert result["filename"] == "document.docx"
            # Verify download_url is constructed correctly
            assert result["download_url"] == "/api/files/download/user-123/uploaded-file-id"


class TestBufferFunctions:
    """Test document buffer creation functions."""

    def test_markdown_to_word_buffer(self):
        """Test _markdown_to_word_buffer returns BytesIO."""
        from docx_tools import _markdown_to_word_buffer

        result = _markdown_to_word_buffer("# Hello World\n\nThis is a test.")

        assert isinstance(result, io.BytesIO)
        assert result.tell() == 0
        content = result.read()
        assert len(content) > 0
        result.close()

    def test_markdown_to_excel_buffer(self):
        """Test _markdown_to_excel_buffer returns BytesIO."""
        from xlsx_tools import _markdown_to_excel_buffer

        markdown = """| Name | Value |
|------|-------|
| A    | 1     |
| B    | 2     |
"""
        result = _markdown_to_excel_buffer(markdown)

        assert isinstance(result, io.BytesIO)
        assert result.tell() == 0
        content = result.read()
        assert len(content) > 0
        result.close()

    def test_create_presentation_buffer(self):
        """Test _create_presentation_buffer returns BytesIO."""
        from pptx_tools import _create_presentation_buffer

        slides = [
            {"slide_type": "title", "slide_title": "Test Presentation", "subtitle": "Test"}
        ]
        result = _create_presentation_buffer(slides)

        assert isinstance(result, io.BytesIO)
        assert result.tell() == 0
        content = result.read()
        assert len(content) > 0
        result.close()

    def test_create_eml_buffer(self):
        """Test _create_eml_buffer returns BytesIO."""
        from email_tools import _create_eml_buffer

        result = _create_eml_buffer(
            to=["test@example.com"],
            re="Test Subject",
            content="<p>Test content</p>",
        )

        assert isinstance(result, io.BytesIO)
        assert result.tell() == 0
        content = result.read()
        assert len(content) > 0
        # The Subject header is RFC 2047 encoded (=?utf-8?q?Test_Subject?=)
        # because it is written as email.header.Header(..., 'utf-8'), so the
        # literal string never appears in the raw bytes. Decode before
        # comparing rather than asserting on the encoded form, which would
        # pin the choice of charset and encoding.
        import email
        from email.header import decode_header, make_header

        parsed = email.message_from_bytes(content)
        assert str(make_header(decode_header(parsed["Subject"]))) == "Test Subject"
        result.close()

    def test_create_xml_buffer(self):
        """Test _create_xml_buffer returns BytesIO."""
        from xml_tools import _create_xml_buffer

        xml_content = "<root><item>Test</item></root>"
        result = _create_xml_buffer(xml_content)

        assert isinstance(result, io.BytesIO)
        assert result.tell() == 0
        content = result.read()
        assert b"<root>" in content
        result.close()


class TestIsLibreChatStrategy:
    """Test is_librechat_strategy function."""

    def test_is_librechat_strategy_callable(self):
        """Test is_librechat_strategy is callable."""
        from upload_tools import is_librechat_strategy

        assert callable(is_librechat_strategy)
        # Default should be LOCAL, not LIBRECHAT
        result = is_librechat_strategy()
        assert isinstance(result, bool)
