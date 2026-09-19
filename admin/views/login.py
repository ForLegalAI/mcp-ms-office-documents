"""The sign-in and sign-out pages.

Both are rendered without the nav (there is nothing to navigate to when signed
out, and the sign-out confirmation is a dead end by design), so they pass
``authed=False`` / render inside the narrow ``login-wrap`` column.
"""
from __future__ import annotations

from typing import Optional

from fasthtml.common import A, Button, Div, H1, Input, P

from admin import components as c
from admin.views.shell import BRAND, page


def login_page(ctx, login_url: str, error: Optional[str] = None):
    return page(
        ctx, "Sign in",
        Div(
            c.card(
                H1(BRAND),
                P("Sign in to manage document templates.", cls="muted"),
                c.flash(error, "err"),
                c.post_form(
                    login_url,
                    c.field("Password",
                            Input(name="password", type="password", required=True,
                                  autofocus=True)),
                    Button("Sign in", type="submit", cls="btn btn-primary"),
                ),
            ),
            cls="login-wrap",
        ),
        authed=False,
    )


def logout_page(ctx, logout_url: str, csrf: str = ""):
    return page(
        ctx, "Sign out",
        Div(
            c.card(
                H1("Sign out?"),
                P("You'll need your password to sign back in.", cls="muted"),
                c.post_form(
                    logout_url,
                    c.action_bar(
                        Button("Sign out", type="submit", cls="btn btn-primary"),
                        A("Cancel", href=ctx.u("/"), cls="btn"),
                    ),
                    csrf=csrf,
                ),
            ),
            cls="login-wrap",
        ),
    )
