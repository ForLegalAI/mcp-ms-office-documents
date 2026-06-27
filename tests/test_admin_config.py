"""Tests for the optional admin-UI configuration in config.py."""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from config import Config, AdminSettings


def _clear_admin_env(monkeypatch):
    for var in ("ADMIN_ENABLED", "ADMIN_PASSWORD", "ADMIN_PATH", "API_KEY"):
        monkeypatch.delenv(var, raising=False)


def test_admin_disabled_by_default(monkeypatch):
    _clear_admin_env(monkeypatch)
    cfg = Config.from_env()
    assert cfg.admin.enabled is False
    assert cfg.admin.path == "/admin"
    assert cfg.admin_password_effective is None


def test_admin_enabled_with_password(monkeypatch):
    _clear_admin_env(monkeypatch)
    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.setenv("ADMIN_PASSWORD", "s3cret")
    cfg = Config.from_env()
    assert cfg.admin.enabled is True
    assert cfg.admin_password_effective == "s3cret"


def test_admin_password_falls_back_to_api_key(monkeypatch):
    _clear_admin_env(monkeypatch)
    monkeypatch.setenv("ADMIN_ENABLED", "1")
    monkeypatch.setenv("API_KEY", "apikey123")
    cfg = Config.from_env()
    assert cfg.admin_password_effective == "apikey123"


def test_explicit_admin_password_wins_over_api_key(monkeypatch):
    _clear_admin_env(monkeypatch)
    monkeypatch.setenv("ADMIN_PASSWORD", "adminpw")
    monkeypatch.setenv("API_KEY", "apikey123")
    cfg = Config.from_env()
    assert cfg.admin_password_effective == "adminpw"


def test_path_normalization():
    assert AdminSettings(path="admin/").path == "/admin"
    assert AdminSettings(path="/console//").path == "/console"
    assert AdminSettings(path="").path == "/admin"
    assert AdminSettings(path="manage").path == "/manage"
