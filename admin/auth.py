"""Single shared-password gate for the admin UI.

Authentication is intentionally minimal (one shared secret, a signed session
cookie) per the chosen access model. The password is
``config.admin_password_effective`` (explicit ``ADMIN_PASSWORD`` or, failing
that, ``API_KEY``). Comparison is constant-time.

The gate is a FastHTML ``before`` callable: it lets the login route and static
assets through and redirects everything else to the login page until the
session is marked authenticated.
"""
from __future__ import annotations

import hmac
import logging
from typing import Optional

from fasthtml.common import RedirectResponse

logger = logging.getLogger(__name__)

SESSION_KEY = "admin_authed"


def check_password(supplied: Optional[str], expected: Optional[str]) -> bool:
    """Constant-time comparison of *supplied* against the *expected* secret."""
    if not expected or supplied is None:
        return False
    return hmac.compare_digest(str(supplied), str(expected))


def make_before(login_path: str):
    """Return a FastHTML ``before`` callable gating everything but login/static.

    *login_path* is the absolute (mount-prefixed) login URL, e.g. ``/admin/login``.
    """
    def _before(req, sess):
        path = req.url.path
        # Always allow the login endpoint and obvious static asset requests.
        if path.rstrip("/").endswith("/login") or path.endswith(".ico") or path.endswith(".css"):
            return None
        if sess.get(SESSION_KEY):
            return None
        return RedirectResponse(login_path, status_code=303)

    return _before
