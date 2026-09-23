"""The admin UI with ADMIN_ENABLED set but no password configured.

Regression for the hardening review: with neither ADMIN_PASSWORD nor API_KEY
the session secret fell back to a hash of a constant string in admin/app.py,
and the gate admitted any validly signed session. So a cookie signed with that
public constant opened every admin page, though no login could ever succeed.
The gate is now locked in that state and the secret is random per process.
"""
import base64
import hashlib
import json
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import itsdangerous
import pytest
from starlette.testclient import TestClient

import admin.store as store_mod
import template_utils as tu
from config import Config

_OLD_CONSTANT_SECRET = hashlib.sha256(b"mcp-office-admin:").hexdigest()


def _signed_session(secret: str, session: dict) -> str:
    """A cookie value as Starlette's SessionMiddleware would sign it."""
    data = base64.b64encode(json.dumps(session).encode())
    return itsdangerous.TimestampSigner(secret).sign(data).decode()


class _FakeRequest:
    def __init__(self, path):
        self.url = type("U", (), {"path": path})()


@pytest.fixture
def passwordless_client(tmp_path, monkeypatch):
    for name in ("custom", "config"):
        (tmp_path / name).mkdir()
    monkeypatch.setattr(store_mod, "_APP_CUSTOM_DIR", tmp_path / "noexist_app_custom")
    monkeypatch.setattr(store_mod, "_APP_CONFIG_DIR", tmp_path / "noexist_app_config")
    monkeypatch.setattr(store_mod, "_LOCAL_CUSTOM_DIR", tmp_path / "custom")
    monkeypatch.setattr(store_mod, "_LOCAL_CONFIG_DIR", tmp_path / "config")
    monkeypatch.setattr(tu, "APP_CUSTOM_DIR", tmp_path / "noexist_app_custom")
    monkeypatch.setattr(tu, "LOCAL_CUSTOM_DIR", tmp_path / "custom")

    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.delenv("ADMIN_PASSWORD", raising=False)
    monkeypatch.delenv("API_KEY", raising=False)

    from fastmcp import FastMCP
    from admin.app import build_combined_app

    app = build_combined_app(FastMCP("test-admin-nopw"), Config.from_env())
    with TestClient(app) as client:
        yield client


def test_locked_gate_ignores_an_authenticated_session():
    from admin import auth

    before = auth.make_before("/admin/login", locked=True)
    signed_in = {auth.SESSION_KEY: True}
    assert before(_FakeRequest("/admin/login"), signed_in) is None
    for path in ("/admin/", "/admin/powerpoint", "/admin/server/status"):
        result = before(_FakeRequest(path), signed_in)
        assert getattr(result, "status_code", None) == 303, f"{path} was let through"


def test_unlocked_gate_still_admits_an_authenticated_session():
    from admin import auth

    before = auth.make_before("/admin/login")
    assert before(_FakeRequest("/admin/"), {auth.SESSION_KEY: True}) is None


def test_a_cookie_signed_with_the_old_constant_key_is_refused(passwordless_client):
    from admin import auth

    cookie = _signed_session(_OLD_CONSTANT_SECRET, {auth.SESSION_KEY: True})
    passwordless_client.cookies.set("session_", cookie)
    for path in ("/admin/", "/admin/powerpoint"):
        response = passwordless_client.get(path, follow_redirects=False)
        assert response.status_code == 303, f"{path} opened without a password"
        assert response.headers["location"].endswith("/admin/login")


def test_no_login_succeeds_without_a_password(passwordless_client):
    for supplied in ("", "anything"):
        passwordless_client.post("/admin/login", data={"password": supplied})
        response = passwordless_client.get("/admin/", follow_redirects=False)
        assert response.status_code == 303
