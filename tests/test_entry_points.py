"""Every URL this server advertises to the outside world answers (#193).

The class of bug this exists for: a URL nothing inside the app ever generates
for itself, so nothing inside the app ever tests it.

`GET /admin` — the address the startup log prints and `docs/admin-ui.md` tells
people to open — returned 404 through ten merged PRs of admin-UI work. Every
test reached the UI from *inside* the mount (16 requests to `/admin/`, none to
`/admin`), because tests were written from the route table and in-app links all
come from `ctx.u("/")`, which already carries the slash. Test and code agreed
because both were built from the same model of the mount; neither checked it
from outside.

So the assertions here are deliberately not written from the route table. Each
one is a URL a *stranger* uses — a Kubernetes probe, an MCP client, a person
typing the documented address — and the only thing asserted is "not a 404",
because what went wrong was never about the body.
"""
import os
import sys
from pathlib import Path
from urllib.parse import urlparse

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from starlette.testclient import TestClient

import admin.store as store_mod
import metrics
from config import Config

#: Every externally-advertised entry point, and who advertises it. A URL here
#: is a promise made somewhere a user can read — a probe in the deployment
#: manifest, the documented admin address, the MCP endpoint in the README — so
#: it is the promise that is under test, not the routing that happens to honour
#: it today.
ENTRY_POINTS = [
    ("/healthz", "startupProbe (main.py, docs/configuration.md)"),
    ("/readyz", "readinessProbe (main.py)"),
    ("/livez", "livenessProbe (main.py)"),
    ("/mcp", "the MCP endpoint (README, AGENTS.md)"),
    ("/admin", "the admin UI (startup log, docs/admin-ui.md)"),
]


@pytest.fixture(scope="module")
def combined(tmp_path_factory):
    """The real combined app, built from `main`'s own MCP instance.

    Not a fresh `FastMCP()`: the health routes are registered in `main.py` via
    `@mcp.custom_route`, so a stand-in instance does not have them and this
    file would pass while saying nothing.
    """
    tmp = tmp_path_factory.mktemp("entry")
    (tmp / "custom").mkdir()
    (tmp / "config").mkdir()
    store_mod._APP_CUSTOM_DIR = tmp / "nx1"
    store_mod._APP_CONFIG_DIR = tmp / "nx2"
    store_mod._LOCAL_CUSTOM_DIR = tmp / "custom"
    store_mod._LOCAL_CONFIG_DIR = tmp / "config"
    os.environ["ADMIN_ENABLED"] = "true"
    os.environ["ADMIN_PASSWORD"] = "pw"
    os.environ.pop("API_KEY", None)

    import main
    from admin.app import build_combined_app

    client = TestClient(build_combined_app(main.mcp, Config.from_env()))
    metrics.reset()
    with client:
        yield client
    metrics.reset()


@pytest.mark.parametrize("url, advertised_by", ENTRY_POINTS)
def test_an_advertised_url_is_not_a_404(combined, url, advertised_by):
    """Exactly as typed — no trailing slash added to make it work.

    A redirect counts: following it is what a browser and a probe both do. A
    404 does not, and neither does the 405/406 that would mean the path
    resolved to the wrong app.
    """
    r = combined.get(url, follow_redirects=False)

    assert r.status_code != 404, (
        f"{url} is advertised by {advertised_by} and does not resolve. "
        "It most likely fell through to the catch-all MCP mount — see the "
        "Mounting section in docs/development/admin-ui.md."
    )
    assert r.status_code != 405, f"{url} resolved, but to something that refuses GET"


def test_the_admin_url_leads_all_the_way_to_the_ui(combined):
    """Not a 404 is the floor; this is the thing a person actually wants.

    Following the redirects, as a browser does, has to end at a real page —
    a redirect into open air would satisfy the test above and still be broken.
    """
    r = combined.get("/admin")

    assert r.status_code == 200
    assert "Template Admin" in r.text
    assert r.url.path.startswith("/admin/"), "it should end up inside the mount"


@pytest.mark.parametrize("root_path", ["", "/office", "/a/b"])
def test_the_root_redirect_keeps_a_proxy_prefix(root_path):
    """Behind `uvicorn --root-path` or a path-rewriting proxy.

    The first version of this redirect built its `Location` from
    `config.admin.path`, which looks identical in development and is wrong the
    moment a prefix is stripped in front of the app: `GET /office/admin`
    answered `/admin/`, sending the browser somewhere that does not exist on
    that deployment. Starlette's own `redirect_slashes` avoids this by building
    the target from the request scope, and so must this.

    Asserted against Starlette's *own* output rather than a literal, so the two
    cannot drift: the point is not "this string" but "whatever the framework
    would have done if the catch-all were not suppressing it".
    """
    from starlette.applications import Starlette
    from starlette.routing import Mount
    from starlette.testclient import TestClient as TC

    from admin.app import build_admin_app

    cfg = Config.from_env()
    import main
    only_admin = Starlette(routes=[Mount(cfg.admin.path,
                                         app=build_admin_app(main.mcp, cfg))])
    with TC(only_admin, root_path=root_path) as c:
        native = c.get(root_path + "/admin", follow_redirects=False)

    # A combined app of its own rather than the module fixture's: entering a
    # TestClient runs the app's lifespan, and re-running it on the shared
    # instance restarts the MCP session manager underneath every later test in
    # this file. (Found by this test failing only in file order.)
    from admin.app import build_combined_app

    with TC(build_combined_app(main.mcp, cfg), root_path=root_path) as c:
        ours = c.get(root_path + "/admin", follow_redirects=False)

    assert native.status_code == 307, "the control itself must redirect"
    assert urlparse(ours.headers["location"]).path == \
        urlparse(native.headers["location"]).path, (
            f"root_path={root_path!r}: we answer "
            f"{ours.headers['location']!r}, Starlette would answer "
            f"{native.headers['location']!r}")


def test_the_root_redirect_keeps_the_query_string(combined):
    """Falls out of deriving the target from the request, and is worth keeping:
    a bookmarked `/admin?...` should not silently lose its parameters."""
    r = combined.get("/admin?next=status", follow_redirects=False)

    assert urlparse(r.headers["location"]).query == "next=status"


def test_the_mcp_endpoint_survives_an_admin_path_collision(tmp_path, monkeypatch):
    """Why the root redirect is GET/HEAD only, and must stay that way.

    `ADMIN_PATH` is free text, so `/mcp` is a legal if unwise value. Nothing
    rejects it, and three things then have to line up for the MCP endpoint to
    keep working: the redirect Route does not take POST, `Mount("/mcp")` does
    not match the bare `/mcp`, and the catch-all still does. A POST therefore
    falls past both admin routes to the MCP app.

    Widen the Route's methods "for uniformity" and this breaks: the POST gets
    a 307 to `/mcp/` and the client never gets a session. Verified — that is
    what this test is here to stop, and a comment would not have held it.
    """
    import admin.store as store
    from starlette.testclient import TestClient as TC

    monkeypatch.setattr(store, "_APP_CUSTOM_DIR", tmp_path / "nx1")
    monkeypatch.setattr(store, "_APP_CONFIG_DIR", tmp_path / "nx2")
    monkeypatch.setattr(store, "_LOCAL_CUSTOM_DIR", tmp_path / "custom")
    monkeypatch.setattr(store, "_LOCAL_CONFIG_DIR", tmp_path / "config")
    monkeypatch.setenv("ADMIN_ENABLED", "true")
    monkeypatch.setenv("ADMIN_PASSWORD", "pw")
    monkeypatch.setenv("ADMIN_PATH", "/mcp")
    monkeypatch.delenv("API_KEY", raising=False)

    import main
    from admin.app import build_combined_app

    with TC(build_combined_app(main.mcp, Config.from_env())) as client:
        r = client.post("/mcp", json={
            "jsonrpc": "2.0", "id": 1, "method": "initialize",
            "params": {"protocolVersion": "2025-06-18", "capabilities": {},
                       "clientInfo": {"name": "t", "version": "1"}},
        }, headers={"Accept": "application/json, text/event-stream",
                    "Content-Type": "application/json"},
           follow_redirects=False)

    assert r.status_code == 200, \
        f"an MCP client got {r.status_code} {r.headers.get('location', '')}"
    assert "mcp-session-id" in r.headers


def test_the_mcp_endpoint_is_not_shadowed_by_the_admin_mount(combined):
    """The other side of the same seam.

    The admin mount is registered first so it wins over the catch-all. If its
    prefix ever widened, the MCP endpoint would start resolving to the admin
    app, and every MCP client would get HTML.
    """
    r = combined.post("/mcp", json={
        "jsonrpc": "2.0", "id": 1, "method": "initialize",
        "params": {"protocolVersion": "2025-06-18", "capabilities": {},
                   "clientInfo": {"name": "t", "version": "1"}},
    }, headers={"Accept": "application/json, text/event-stream",
                "Content-Type": "application/json"},
       follow_redirects=False)

    # follow_redirects=False on purpose: TestClient follows by default, so a
    # bare 200 cannot tell "matched directly" from "307'd somewhere and the
    # client quietly went there". That distinction is the entire subject of
    # this file, and this assertion was blind to it.
    assert r.status_code == 200, \
        f"got {r.status_code} {r.headers.get('location', '')}"
    assert "mcp-session-id" in r.headers
