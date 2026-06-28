"""Tests for API-key authentication middleware and related config handling.

Covers:
- Key extraction from various HTTP header formats
- Case-insensitive header matching
- Constant-time comparison via secrets.compare_digest
- Request rejection when key is missing or wrong
- Request passthrough when key is valid
- Config.api_key population from environment
- Middleware registration logic (enabled / disabled based on env)
"""

import os
import sys
from pathlib import Path
from unittest.mock import patch, AsyncMock, MagicMock

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from middleware import ApiKeyAuthMiddleware
from fastmcp.exceptions import AuthorizationError


# ======================================================================
# Key extraction
# ======================================================================

class TestExtractKey:
    """Unit tests for ApiKeyAuthMiddleware._extract_key."""

    def test_bearer_token(self):
        headers = {"Authorization": "Bearer my-secret-key"}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "my-secret-key"

    def test_bearer_token_extra_whitespace(self):
        headers = {"Authorization": "Bearer   my-secret-key  "}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "my-secret-key"

    def test_bearer_case_insensitive_scheme(self):
        """The 'Bearer' prefix check should be case-insensitive."""
        headers = {"Authorization": "BEARER my-key"}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "my-key"

    def test_bearer_lowercase_scheme(self):
        headers = {"Authorization": "bearer my-key"}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "my-key"

    def test_plain_authorization(self):
        """Authorization header without scheme prefix is treated as plain token."""
        headers = {"Authorization": "plain-api-key"}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "plain-api-key"

    def test_plain_authorization_stripped(self):
        headers = {"Authorization": "  plain-api-key  "}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "plain-api-key"

    def test_x_api_key_header(self):
        headers = {"x-api-key": "key-from-header"}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "key-from-header"

    def test_x_api_key_stripped(self):
        headers = {"x-api-key": "  key-from-header  "}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "key-from-header"

    def test_empty_headers(self):
        assert ApiKeyAuthMiddleware._extract_key({}) is None

    def test_irrelevant_headers(self):
        headers = {"Content-Type": "application/json", "Accept": "*/*"}
        assert ApiKeyAuthMiddleware._extract_key(headers) is None

    # -- Case-insensitive header name matching --

    def test_authorization_mixed_case(self):
        headers = {"AUTHORIZATION": "Bearer case-key"}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "case-key"

    def test_x_api_key_mixed_case(self):
        headers = {"X-Api-Key": "mixed-case-key"}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "mixed-case-key"

    def test_x_api_key_upper_case(self):
        headers = {"X-API-KEY": "upper-key"}
        assert ApiKeyAuthMiddleware._extract_key(headers) == "upper-key"

    # -- Priority: Authorization wins over x-api-key --

    def test_authorization_takes_precedence_over_x_api_key(self):
        """When both headers are present, Authorization should be used."""
        headers = {
            "Authorization": "Bearer auth-key",
            "x-api-key": "api-key-value",
        }
        assert ApiKeyAuthMiddleware._extract_key(headers) == "auth-key"


# ======================================================================
# Constructor validation
# ======================================================================

class TestConstructor:
    """Ensure the middleware rejects invalid construction arguments."""

    def test_empty_key_raises(self):
        with pytest.raises(ValueError, match="non-empty"):
            ApiKeyAuthMiddleware("")

    def test_none_key_raises(self):
        """Falsy values must be rejected."""
        with pytest.raises((ValueError, TypeError)):
            ApiKeyAuthMiddleware(None)  # type: ignore[arg-type]

    def test_valid_key_accepted(self):
        mw = ApiKeyAuthMiddleware("good-key")
        assert mw.expected_key == "good-key"


# ======================================================================
# on_request hook
# ======================================================================

def _make_context(method: str = "tools/call"):
    """Build a minimal MiddlewareContext-like mock."""
    ctx = MagicMock()
    ctx.method = method
    return ctx


class TestAuthorizationHeaderSurfaced:
    """Regression: FastMCP's get_http_headers() strips `authorization` by default.

    These use the REAL get_http_headers (only the underlying request is faked),
    so they would fail if on_request stopped opting the header back in.
    """

    async def test_real_get_http_headers_surfaces_authorization(self):
        import fastmcp.server.dependencies as deps
        from starlette.datastructures import Headers

        fake_request = MagicMock()
        fake_request.headers = Headers({"Authorization": "Bearer secret-123", "host": "x"})

        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock(return_value="ok")
        context = _make_context()

        # Patch only FastMCP's internal request accessor so the real
        # get_http_headers runs (it would strip Authorization unless on_request
        # passes include=). If this breaks after a FastMCP upgrade, check
        # fastmcp.server.dependencies.get_http_request.
        with patch.object(deps, "get_http_request", return_value=fake_request):
            result = await mw.on_request(context, call_next)

        call_next.assert_awaited_once_with(context)
        assert result == "ok"

    async def test_on_request_opts_in_authorization_header(self):
        """on_request must ask get_http_headers to include the auth headers."""
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock(return_value="ok")
        context = _make_context()

        with patch("middleware.get_http_headers",
                   return_value={"authorization": "Bearer secret-123"}) as ghh:
            await mw.on_request(context, call_next)

        _args, kwargs = ghh.call_args
        included = {h.lower() for h in (kwargs.get("include") or set())}
        assert "authorization" in included


class TestOnRequest:
    """Integration-style tests for the on_request middleware hook."""

    async def test_valid_bearer_token_passes(self):
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock(return_value="ok")
        context = _make_context()

        with patch("middleware.get_http_headers", return_value={"Authorization": "Bearer secret-123"}):
            result = await mw.on_request(context, call_next)

        call_next.assert_awaited_once_with(context)
        assert result == "ok"

    async def test_valid_x_api_key_passes(self):
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock(return_value="ok")
        context = _make_context()

        with patch("middleware.get_http_headers", return_value={"x-api-key": "secret-123"}):
            result = await mw.on_request(context, call_next)

        call_next.assert_awaited_once_with(context)
        assert result == "ok"

    async def test_valid_plain_authorization_passes(self):
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock(return_value="ok")
        context = _make_context()

        with patch("middleware.get_http_headers", return_value={"Authorization": "secret-123"}):
            result = await mw.on_request(context, call_next)

        call_next.assert_awaited_once_with(context)
        assert result == "ok"

    async def test_missing_key_rejects(self):
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock()
        context = _make_context()

        with patch("middleware.get_http_headers", return_value={}):
            with pytest.raises(AuthorizationError):
                await mw.on_request(context, call_next)

        call_next.assert_not_awaited()

    async def test_wrong_key_rejects(self):
        mw = ApiKeyAuthMiddleware("correct-key")
        call_next = AsyncMock()
        context = _make_context()

        with patch("middleware.get_http_headers", return_value={"Authorization": "Bearer wrong-key"}):
            with pytest.raises(AuthorizationError):
                await mw.on_request(context, call_next)

        call_next.assert_not_awaited()

    async def test_none_headers_rejects(self):
        """get_http_headers() may return None for non-HTTP transports."""
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock()
        context = _make_context()

        with patch("middleware.get_http_headers", return_value=None):
            with pytest.raises(AuthorizationError):
                await mw.on_request(context, call_next)

        call_next.assert_not_awaited()

    async def test_constant_time_comparison_used(self):
        """Ensure secrets.compare_digest is used (not plain !=)."""
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock(return_value="ok")
        context = _make_context()

        with patch("middleware.get_http_headers", return_value={"x-api-key": "secret-123"}), \
             patch("middleware.secrets.compare_digest", return_value=True) as mock_compare:
            await mw.on_request(context, call_next)

        mock_compare.assert_called_once_with("secret-123", "secret-123")

    async def test_failed_attempt_increments_counter(self):
        """Each auth failure should increment the internal counter."""
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock()
        context = _make_context()

        assert mw._failed_attempts == 0

        with patch("middleware.get_http_headers", return_value={"x-api-key": "wrong"}):
            with pytest.raises(AuthorizationError):
                await mw.on_request(context, call_next)

        # First failure triggers a WARNING which resets counter to 0
        # so after one reject the counter is 0 (just emitted warning)
        assert mw._failed_attempts == 0
        assert mw._last_warn_time > 0

    async def test_throttled_warning_not_emitted_within_interval(self):
        """WARNING should NOT fire again within the throttle window."""
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock()
        context = _make_context()

        # Simulate that a warning was just emitted
        import time as _time
        mw._last_warn_time = _time.monotonic()

        with patch("middleware.get_http_headers", return_value={}), \
             patch("middleware.logger") as mock_logger:
            with pytest.raises(AuthorizationError):
                await mw.on_request(context, call_next)

        # DEBUG is always emitted
        mock_logger.debug.assert_called_once()
        # WARNING should NOT be emitted (within interval)
        mock_logger.warning.assert_not_called()
        # Counter should have incremented without reset
        assert mw._failed_attempts == 1

    async def test_throttled_warning_emitted_after_interval(self):
        """WARNING should fire again once the throttle window has elapsed."""
        mw = ApiKeyAuthMiddleware("secret-123")
        call_next = AsyncMock()
        context = _make_context()

        # Pretend the last warning was long ago
        mw._last_warn_time = 0.0
        mw._failed_attempts = 5  # accumulated silently

        with patch("middleware.get_http_headers", return_value={}), \
             patch("middleware.logger") as mock_logger:
            with pytest.raises(AuthorizationError):
                await mw.on_request(context, call_next)

        # WARNING should have been emitted (interval elapsed)
        mock_logger.warning.assert_called_once()
        # Counter should reset after warning
        assert mw._failed_attempts == 0


# ======================================================================
# Config integration
# ======================================================================

class TestConfigApiKey:
    """Verify that Config.from_env reads API_KEY correctly."""

    def test_api_key_set(self):
        env = {
            "DEBUG": "",
            "UPLOAD_STRATEGY": "LOCAL",
            "API_KEY": "env-secret-key",
        }
        with patch.dict(os.environ, env, clear=True):
            from config import Config
            cfg = Config.from_env()
        assert cfg.api_key == "env-secret-key"

    def test_api_key_empty(self):
        env = {
            "DEBUG": "",
            "UPLOAD_STRATEGY": "LOCAL",
            "API_KEY": "",
        }
        with patch.dict(os.environ, env, clear=True):
            from config import Config
            cfg = Config.from_env()
        assert cfg.api_key is None

    def test_api_key_missing(self):
        env = {
            "DEBUG": "",
            "UPLOAD_STRATEGY": "LOCAL",
        }
        with patch.dict(os.environ, env, clear=True):
            from config import Config
            cfg = Config.from_env()
        assert cfg.api_key is None

    def test_api_key_whitespace_only(self):
        env = {
            "DEBUG": "",
            "UPLOAD_STRATEGY": "LOCAL",
            "API_KEY": "   ",
        }
        with patch.dict(os.environ, env, clear=True):
            from config import Config
            cfg = Config.from_env()
        assert cfg.api_key is None

    def test_api_key_stripped(self):
        env = {
            "DEBUG": "",
            "UPLOAD_STRATEGY": "LOCAL",
            "API_KEY": "  spaced-key  ",
        }
        with patch.dict(os.environ, env, clear=True):
            from config import Config
            cfg = Config.from_env()
        assert cfg.api_key == "spaced-key"

