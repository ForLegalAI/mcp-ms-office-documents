"""`ADMIN_PATH` must not take over a path the server already serves (#193).

The admin routes are registered before the catch-all MCP mount so they win, so
an `ADMIN_PATH` equal to the MCP endpoint or a Kubernetes probe shadows it.
Nothing rejected such a value, and the failure is silent: the pod starts, the
log says the admin UI is enabled, and a probe or an MCP client quietly gets the
wrong app.

The list lives in `config.RESERVED_PATHS` rather than being imported from
`main`, because `config` stays free of application imports. That is a copy, and
a copy drifts — so the first test here compares it against the routes `main`
actually registers, the way `tests/test_admin_style_keys.py` ties the admin
style keys to the renderer's.
"""
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest

from config import RESERVED_PATHS, AdminSettings


def test_the_reserved_list_matches_what_main_actually_serves():
    """The copy against the original.

    A path `main` starts serving and this list does not know about is one
    `ADMIN_PATH` can still shadow; a path here that `main` stopped serving
    rejects a value that would have been fine.
    """
    import main

    served = set()
    for route in main.mcp.http_app(path="/mcp").routes:
        path = getattr(route, "path", None)
        # Mounts and the session-manager plumbing are not what this guards;
        # only the concrete paths a client or a probe asks for by name.
        if path and type(route).__name__ == "Route":
            served.add(path)
    served.add("/mcp")  # mcp.http_app(path=...), not a Route on this router

    assert RESERVED_PATHS == served, (
        "config.RESERVED_PATHS has drifted from the paths main.py serves: "
        f"missing {served - RESERVED_PATHS}, stale {RESERVED_PATHS - served}"
    )


@pytest.mark.parametrize("path", sorted(RESERVED_PATHS))
def test_a_colliding_admin_path_falls_back(path, caplog):
    """Falls back rather than refusing to start.

    The admin UI is optional; the MCP server and the probes are not. Refusing
    to boot would take document generation down over a misconfiguration of a
    secondary feature, so the reserved path keeps doing its job and the UI
    lands on the documented default — loudly.
    """
    with caplog.at_level("ERROR"):
        settings = AdminSettings(enabled=True, path=path)

    assert settings.path == "/admin"
    # getMessage(), not .message: the latter is only populated once a handler
    # has formatted the record, so it is missing here and the assertion would
    # have raised rather than checked anything.
    logged = [r.getMessage() for r in caplog.records]
    assert any("reserved path" in m and path in m for m in logged), (
        "a silently relocated admin UI is the same class of surprise; "
        f"logged: {logged}")


@pytest.mark.parametrize("path", ["/admin", "/manage", "/mcp2", "/healthz-ui",
                                  "/Admin"])
def test_an_ordinary_admin_path_is_left_alone(path):
    """The guard must not cry wolf on a path that merely resembles a reserved
    one: matching is exact, and `/mcp2` is nobody's MCP endpoint."""
    assert AdminSettings(enabled=True, path=path).path == path


def test_the_probes_and_the_mcp_endpoint_survive_a_colliding_config(tmp_path,
                                                                    monkeypatch):
    """End to end, because the point is what a probe gets, not what a
    validator returns.

    Before the guard, `ADMIN_PATH=/healthz` answered the startup probe with a
    307 into the admin UI — and that redirect is one this PR introduced, so
    without the guard this change would have made the collision worse than it
    found it.
    """
    from starlette.testclient import TestClient

    import admin.store as store
    from config import Config

    monkeypatch.setattr(store, "_APP_CUSTOM_DIR", tmp_path / "nx1")
    monkeypatch.setattr(store, "_APP_CONFIG_DIR", tmp_path / "nx2")
    monkeypatch.setattr(store, "_LOCAL_CUSTOM_DIR", tmp_path / "custom")
    monkeypatch.setattr(store, "_LOCAL_CONFIG_DIR", tmp_path / "config")
    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.setenv("ADMIN_PASSWORD", "pw")
    monkeypatch.setenv("ADMIN_PATH", "/healthz")
    monkeypatch.delenv("API_KEY", raising=False)

    import main
    from admin.app import build_combined_app

    with TestClient(build_combined_app(main.mcp, Config.from_env())) as client:
        probe = client.get("/healthz", follow_redirects=False)
        admin = client.get("/admin", follow_redirects=False)

    assert probe.status_code == 200, \
        f"the startup probe got {probe.status_code} {probe.headers.get('location', '')}"
    assert probe.text == "ok", "and from the health route, not the admin app"
    assert admin.status_code == 307, "the UI moved to the default, and is there"
