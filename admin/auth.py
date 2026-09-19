"""Single shared-password gate for the admin UI.

Authentication is intentionally minimal (one shared secret, a signed session
cookie) per the chosen access model. The password is
``config.admin_password_effective`` (explicit ``ADMIN_PASSWORD`` or, failing
that, ``API_KEY``). Comparison is constant-time.

The gate is a FastHTML ``before`` callable: it lets the login route through and
redirects everything else to the login page until the session is marked
authenticated. Nothing else is public.
"""
from __future__ import annotations

import hmac
import logging
import secrets
from typing import Optional

from fasthtml.common import RedirectResponse

logger = logging.getLogger(__name__)

SESSION_KEY = "admin_authed"
CSRF_KEY = "csrf_token"


def check_password(supplied: Optional[str], expected: Optional[str]) -> bool:
    """Constant-time comparison of *supplied* against the *expected* secret."""
    if not expected or supplied is None:
        return False
    return hmac.compare_digest(str(supplied), str(expected))


def ensure_csrf(sess) -> str:
    """Return the session's CSRF token, generating one on first use."""
    token = sess.get(CSRF_KEY)
    if not token:
        token = secrets.token_urlsafe(32)
        sess[CSRF_KEY] = token
    return token


def valid_csrf(sess, supplied: Optional[str]) -> bool:
    """Constant-time check that *supplied* matches the session CSRF token."""
    token = sess.get(CSRF_KEY)
    if not token or not supplied:
        return False
    return hmac.compare_digest(str(token), str(supplied))


def make_before(login_path: str):
    """Return a FastHTML ``before`` callable gating everything but login/static.

    *login_path* is the absolute (mount-prefixed) login URL, e.g. ``/admin/login``.
    """
    allowed = {login_path, login_path + "/"}

    def _before(req, sess):
        path = req.url.path
        # Only the login endpoint is public. There is deliberately no exemption
        # for static files: the UI serves none — the theme and scripts are
        # inlined into every page — so a suffix rule like ".css"/".ico" would
        # buy nothing and silently expose the first route that ever matched it
        # (#159). If static files are added, exempt their mount prefix here.
        if path in allowed:
            return None
        if sess.get(SESSION_KEY):
            ensure_csrf(sess)  # make a token available to rendered forms
            return None
        return RedirectResponse(login_path, status_code=303)

    return _before
