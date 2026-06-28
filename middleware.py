"""API-key authentication middleware for the MCP Office Documents server.

When an API key is configured (via the API_KEY environment variable), every
incoming request must carry a matching key in one of the supported HTTP
headers.  If no key is configured the middleware is not registered at all.

Supported header formats (checked in order):
  1. ``Authorization: Bearer <key>``
  2. ``Authorization: <key>``   (plain token, no scheme)
  3. ``x-api-key: <key>``
"""

from __future__ import annotations

import logging
import secrets
import time
from typing import Optional

from fastmcp.server.middleware import Middleware, MiddlewareContext
from fastmcp.server.dependencies import get_http_headers
from fastmcp.exceptions import AuthorizationError

logger = logging.getLogger(__name__)


class ApiKeyAuthMiddleware(Middleware):
    """Reject requests that do not carry a valid API key.

    The middleware inspects the incoming HTTP headers for an API key and
    compares it against the *expected_key* provided at construction time.

    Look-up order:
      1. ``Authorization`` header with ``Bearer`` scheme
      2. ``Authorization`` header used as a plain token
      3. ``x-api-key`` header

    If the key is missing or does not match, the request is rejected with
    an ``McpError`` (error code ``-32001``).
    """

    # Throttle window (seconds) for WARNING-level auth failure messages
    # to avoid log flooding from brute-force or misconfigured clients.
    _WARN_INTERVAL_SECONDS: float = 60.0

    def __init__(self, expected_key: str) -> None:
        if not expected_key:
            raise ValueError("ApiKeyAuthMiddleware requires a non-empty expected_key")
        self.expected_key = expected_key
        self._failed_attempts: int = 0
        self._last_warn_time: float = 0.0

    # ------------------------------------------------------------------
    # Key extraction
    # ------------------------------------------------------------------
    @staticmethod
    def _extract_key(headers: dict[str, str]) -> Optional[str]:
        """Return the API key from the first matching header, or *None*.

        Header names are compared case-insensitively (HTTP headers are
        case-insensitive per RFC 7230 §3.2).
        """
        # Normalize header names to lower-case for reliable look-up
        lower = {k.lower(): v for k, v in headers.items()}

        # 1. Authorization header
        auth = lower.get("authorization")
        if auth:
            # "Bearer <token>" – standard OAuth 2.0 scheme
            if auth.lower().startswith("bearer "):
                return auth[7:].strip()
            # Plain token (no scheme prefix)
            return auth.strip()

        # 2. x-api-key header (common API-gateway convention)
        api_key = lower.get("x-api-key")
        if api_key:
            return api_key.strip()

        return None

    # ------------------------------------------------------------------
    # Middleware hook
    # ------------------------------------------------------------------
    async def on_request(self, context: MiddlewareContext, call_next):
        # NOTE: get_http_headers() strips a set of "problematic" headers by
        # default — and that set INCLUDES `authorization` (and
        # `proxy-authorization`). Without opting them back in via `include=`,
        # an `Authorization: Bearer <key>` request arrives here with the header
        # already removed, so only `x-api-key` would ever authenticate. Pass
        # include={...} so the Authorization header is surfaced to us.
        headers = get_http_headers(include={"authorization", "proxy-authorization"}) or {}
        api_key = self._extract_key(headers)

        if api_key is None or not secrets.compare_digest(api_key, self.expected_key):
            self._failed_attempts += 1
            now = time.monotonic()

            # Always log at DEBUG (cheap, only visible when debug is on)
            logger.debug("Auth failure on %s (attempt #%d)", context.method, self._failed_attempts)

            # Throttled WARNING: emit at most once per interval
            if now - self._last_warn_time >= self._WARN_INTERVAL_SECONDS:
                logger.warning(
                    "Rejected %d auth attempt(s) in the last %.0fs",
                    self._failed_attempts,
                    self._WARN_INTERVAL_SECONDS,
                )
                self._failed_attempts = 0
                self._last_warn_time = now

            raise AuthorizationError("Unauthorized – invalid or missing API key")

        return await call_next(context)


