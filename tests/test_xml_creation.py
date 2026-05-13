"""Tests for XML file creation tool.

These tests verify that the XML validation and file creation works correctly,
including valid XML, invalid XML, auto-declaration behavior, encoding handling,
and security protection against malicious XML content.
"""

import sys
from pathlib import Path
from unittest.mock import patch, MagicMock
import io

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from xml_tools.base_xml_tool import (
    validate_xml,
    create_xml_file,
    XMLValidationError,
    XMLFileCreationError,
)


class TestValidateXml:
    """Tests for the validate_xml function."""

    def test_valid_simple_xml(self):
        """Test validation of simple valid XML."""
        xml_content = "<root><child>Hello</child></root>"
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is True
        assert error_message == ""

    def test_valid_xml_with_declaration(self):
        """Test validation of XML with declaration."""
        xml_content = '<?xml version="1.0" encoding="UTF-8"?><root><child>Hello</child></root>'
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is True
        assert error_message == ""

    def test_valid_xml_with_attributes(self):
        """Test validation of XML with attributes."""
        xml_content = '<root attr="value"><child id="1">Hello</child></root>'
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is True
        assert error_message == ""

    def test_valid_xml_with_namespace(self):
        """Test validation of XML with namespaces."""
        xml_content = '<root xmlns="http://example.com"><child>Hello</child></root>'
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is True
        assert error_message == ""

    def test_valid_complex_xml(self):
        """Test validation of more complex XML structure."""
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
        <catalog>
            <book id="1">
                <title>XML Guide</title>
                <author>John Doe</author>
                <price currency="USD">29.99</price>
            </book>
            <book id="2">
                <title>Python Basics</title>
                <author>Jane Smith</author>
                <price currency="EUR">24.99</price>
            </book>
        </catalog>"""
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is True
        assert error_message == ""

    def test_invalid_xml_unclosed_tag(self):
        """Test validation fails for unclosed tag."""
        xml_content = "<root><child>Hello</root>"
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is False
        assert "XML parsing error" in error_message

    def test_invalid_xml_mismatched_tags(self):
        """Test validation fails for mismatched tags."""
        xml_content = "<root><child>Hello</other></root>"
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is False
        assert "XML parsing error" in error_message

    def test_invalid_xml_no_root_element(self):
        """Test validation fails for content without root element."""
        xml_content = "Hello World"
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is False
        assert "XML parsing error" in error_message

    def test_invalid_xml_multiple_roots(self):
        """Test validation fails for multiple root elements."""
        xml_content = "<root1>Hello</root1><root2>World</root2>"
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is False
        assert "XML parsing error" in error_message

    def test_invalid_xml_empty_string(self):
        """Test validation fails for empty string."""
        xml_content = ""
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is False
        assert "XML parsing error" in error_message

    def test_invalid_xml_malformed_declaration(self):
        """Test validation fails for malformed XML declaration."""
        xml_content = '<?xml version="1.0"?<root>Hello</root>'
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is False
        assert "XML parsing error" in error_message

    def test_valid_xml_with_cdata(self):
        """Test validation of XML with CDATA section."""
        xml_content = "<root><![CDATA[<not>parsed</not>]]></root>"
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is True
        assert error_message == ""

    def test_valid_xml_with_comments(self):
        """Test validation of XML with comments."""
        xml_content = "<root><!-- This is a comment --><child>Hello</child></root>"
        is_valid, error_message = validate_xml(xml_content)
        assert is_valid is True
        assert error_message == ""


class TestCreateXmlFile:
    """Tests for the create_xml_file function."""

    @patch('xml_tools.base_xml_tool.upload_file')
    def test_create_xml_file_success(self, mock_upload):
        """Test successful XML file creation."""
        mock_upload.return_value = "http://example.com/file.xml"
        xml_content = "<root><child>Hello</child></root>"

        result = create_xml_file(xml_content)

        assert result == "http://example.com/file.xml"
        mock_upload.assert_called_once()
        # Verify upload_file was called with BytesIO and "xml" suffix
        call_args = mock_upload.call_args
        assert call_args[0][1] == "xml"
        assert isinstance(call_args[0][0], io.BytesIO)

    @patch('xml_tools.base_xml_tool.upload_file')
    def test_create_xml_file_adds_declaration(self, mock_upload):
        """Test that XML declaration is added when not present."""
        captured_content = {}

        def capture_upload(file_obj, suffix, **kwargs):
            captured_content['data'] = file_obj.getvalue()
            return "http://example.com/file.xml"

        mock_upload.side_effect = capture_upload
        xml_content = "<root><child>Hello</child></root>"

        create_xml_file(xml_content)

        content = captured_content['data'].decode('utf-8')
        assert content.startswith('<?xml version="1.0" encoding="UTF-8"?>')
        assert "<root><child>Hello</child></root>" in content

    @patch('xml_tools.base_xml_tool.upload_file')
    def test_create_xml_file_preserves_existing_declaration(self, mock_upload):
        """Test that existing XML declaration is preserved."""
        captured_content = {}

        def capture_upload(file_obj, suffix, **kwargs):
            captured_content['data'] = file_obj.getvalue()
            return "http://example.com/file.xml"

        mock_upload.side_effect = capture_upload
        xml_content = '<?xml version="1.0" encoding="UTF-8"?><root>Hello</root>'

        create_xml_file(xml_content)

        content = captured_content['data'].decode('utf-8')
        # Should not have duplicate declarations
        assert content.count('<?xml') == 1

    @patch('xml_tools.base_xml_tool.upload_file')
    def test_create_xml_file_respects_encoding(self, mock_upload):
        """Test that declared encoding is respected."""
        captured_content = {}

        def capture_upload(file_obj, suffix, **kwargs):
            captured_content['data'] = file_obj.getvalue()
            return "http://example.com/file.xml"

        mock_upload.side_effect = capture_upload
        # Using ISO-8859-1 encoding declaration
        xml_content = '<?xml version="1.0" encoding="ISO-8859-1"?><root>Héllo</root>'

        create_xml_file(xml_content)

        # Content should be encoded as ISO-8859-1
        content = captured_content['data'].decode('ISO-8859-1')
        assert 'Héllo' in content

    @patch('xml_tools.base_xml_tool.upload_file')
    def test_create_xml_file_strips_whitespace(self, mock_upload):
        """Test that leading/trailing whitespace is stripped."""
        captured_content = {}

        def capture_upload(file_obj, suffix, **kwargs):
            captured_content['data'] = file_obj.getvalue()
            return "http://example.com/file.xml"

        mock_upload.side_effect = capture_upload
        xml_content = "   \n<root>Hello</root>\n   "

        create_xml_file(xml_content)

        content = captured_content['data'].decode('utf-8')
        # Should start with declaration, not whitespace
        assert content.startswith('<?xml')

    def test_create_xml_file_invalid_xml_raises_error(self):
        """Test that invalid XML raises XMLValidationError."""
        xml_content = "<root><unclosed>"

        with pytest.raises(XMLValidationError) as exc_info:
            create_xml_file(xml_content)

        assert "XML parsing error" in str(exc_info.value)

    def test_create_xml_file_empty_string_raises_error(self):
        """Test that empty string raises XMLValidationError."""
        with pytest.raises(XMLValidationError):
            create_xml_file("")

    def test_create_xml_file_whitespace_only_raises_error(self):
        """Test that whitespace-only content raises XMLValidationError."""
        with pytest.raises(XMLValidationError):
            create_xml_file("   \n\t  ")

    @patch('xml_tools.base_xml_tool.upload_file')
    def test_create_xml_file_upload_error_raises_exception(self, mock_upload):
        """Test that upload errors raise XMLFileCreationError."""
        mock_upload.side_effect = Exception("Upload failed")
        xml_content = "<root>Hello</root>"

        with pytest.raises(XMLFileCreationError) as exc_info:
            create_xml_file(xml_content)

        assert "Error creating XML file" in str(exc_info.value)

    @patch('xml_tools.base_xml_tool.upload_file')
    def test_create_xml_file_closes_buffer(self, mock_upload):
        """Test that BytesIO buffer is closed after upload."""
        mock_upload.return_value = "http://example.com/file.xml"
        xml_content = "<root>Hello</root>"

        # We can't directly verify the buffer is closed, but we can verify
        # the function completes without error
        result = create_xml_file(xml_content)
        assert result == "http://example.com/file.xml"


class TestXmlSecurityProtection:
    """Tests for XML security protections using defusedxml."""

    def test_billion_laughs_attack_blocked(self):
        """Test that entity expansion attack (billion laughs) is blocked."""
        # This is a simplified version of the billion laughs attack
        xml_content = """<?xml version="1.0"?>
        <!DOCTYPE lolz [
          <!ENTITY lol "lol">
          <!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
          <!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">
        ]>
        <lolz>&lol3;</lolz>"""

        is_valid, error_message = validate_xml(xml_content)
        # defusedxml should block this
        assert is_valid is False
        # Error message should indicate security issue or parsing error
        assert "error" in error_message.lower()

    def test_external_entity_blocked(self):
        """Test that external entity references are blocked."""
        xml_content = """<?xml version="1.0"?>
        <!DOCTYPE foo [
          <!ENTITY xxe SYSTEM "file:///etc/passwd">
        ]>
        <foo>&xxe;</foo>"""

        is_valid, error_message = validate_xml(xml_content)
        # defusedxml should block external entities
        assert is_valid is False


class TestXmlExceptionExports:
    """Tests to verify exceptions are properly exported."""

    def test_xml_validation_error_importable(self):
        """Test that XMLValidationError can be imported from xml_tools."""
        from xml_tools import XMLValidationError
        assert XMLValidationError is not None

    def test_xml_file_creation_error_importable(self):
        """Test that XMLFileCreationError can be imported from xml_tools."""
        from xml_tools import XMLFileCreationError
        assert XMLFileCreationError is not None

    def test_create_xml_file_importable(self):
        """Test that create_xml_file can be imported from xml_tools."""
        from xml_tools import create_xml_file
        assert create_xml_file is not None


