"""The admin UI's section/tab navigation (the product-level pages).

The UI is organised by what an admin is working on — Word, PowerPoint, Excel,
Email, XML, Server — with the old function-per-page views as tabs inside each
section. Three things here are easy to break and impossible to notice:

* **A section that no link reaches.** The nav is derived from the section
  table, so a new section appears in it automatically — but only if the table
  is the single source. #157 was this exact bug one layer down, when the
  PowerPoint create page became unreachable.
* **A shadowed route.** A section slug occupies the same first path segment as
  a template kind (``email`` is both) and as the admin's literal pages
  (``/new/docx``), and the section routes are registered first. Starlette
  matches in registration order, so a section route that is too greedy silently
  eats a handler — a POST to ``/admin/email/save`` rendering the Email section
  would look like a save that did nothing.
* **A dead bookmark.** ``/status``, ``/files``, ``/styles`` and ``/base`` were
  top-bar links for the whole life of the UI.

The shadowing test walks the real route table rather than asserting the two
rules that currently prevent it (GET-only section routes, POST-only
two-segment kind routes), so a future change that keeps the invariant by other
means still passes, and one that breaks it fails wherever it happens.
"""
import re
import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

import pytest
from starlette.routing import Route

from admin import sections as sec_mod
from admin.app import MOVED_PATHS
from admin.sections import SECTIONS, assert_slugs_free, nav_items, tab_items

from tests.test_admin_app import admin_client  # noqa: F401  (fixture)


def _admin_routes(app):
    """Every route under the admin mount, in registration order."""
    for route in app.routes:
        sub = getattr(route, "app", None)
        if sub is not None and getattr(sub, "routes", None):
            for inner in sub.routes:
                if isinstance(inner, Route):
                    yield inner
        if isinstance(route, Route):
            yield route


# ---------------------------------------------------------------------------
# Every section and tab is reachable
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("slug", [s.slug for s in SECTIONS])
def test_every_section_renders(admin_client, slug):
    client, _mcp = admin_client
    r = client.get(f"/admin/{slug}")
    assert r.status_code == 200, r.text[:400]


@pytest.mark.parametrize(
    "slug,tab",
    [(s.slug, t.slug) for s in SECTIONS for t in s.tabs],
)
def test_every_tab_renders(admin_client, slug, tab):
    client, _mcp = admin_client
    r = client.get(f"/admin/{slug}/{tab}")
    assert r.status_code == 200, r.text[:400]


@pytest.mark.parametrize(
    "slug,tab",
    [(s.slug, t.slug) for s in SECTIONS for t in s.tabs if len(s.tabs) > 1],
)
def test_the_open_tab_is_marked_active(admin_client, slug, tab):
    """Exactly one tab carries the active class, and it is the one requested."""
    client, _mcp = admin_client
    html = client.get(f"/admin/{slug}/{tab}").text
    active = re.findall(r'<a href="([^"]+)" class="tab active"', html)
    assert len(active) == 1, f"{len(active)} active tabs on {slug}/{tab}"
    expected = sec_mod.section(slug).href(lambda p="": f"/admin{p}", tab)
    assert active[0] == expected


def test_the_open_section_is_marked_in_the_nav(admin_client):
    client, _mcp = admin_client
    html = client.get("/admin/excel").text
    assert '<a href="/admin/excel" class="active">Excel</a>' in html
    assert '<a href="/admin/word" class="active">' not in html


def test_a_single_tab_section_renders_no_tab_bar(admin_client):
    """XML has one tab; a tab bar of one says nothing and is not drawn."""
    client, _mcp = admin_client
    assert len(sec_mod.section("xml").tabs) == 1
    assert 'class="tabs"' not in client.get("/admin/xml").text


def test_an_unknown_tab_redirects_to_the_section(admin_client):
    client, _mcp = admin_client
    r = client.get("/admin/word/nosuchtab", follow_redirects=False)
    assert r.status_code == 303
    assert r.headers["location"] == "/admin/word"


def test_an_unknown_section_is_not_found(admin_client):
    """No catch-all: a section route exists only for a slug in the table.

    The alternative — a ``/{slug}`` pattern that redirects home for anything
    unknown — would also match ``/new`` and ``/base``, and being registered
    first it would swallow them. A 404 for a path nothing serves is both
    correct and what keeps those pages reachable.
    """
    client, _mcp = admin_client
    assert client.get("/admin/nosuchsection").status_code == 404


# ---------------------------------------------------------------------------
# The dashboard and the nav link to everything
# ---------------------------------------------------------------------------

def test_the_dashboard_links_to_every_section(admin_client):
    """#157 in its general form: a section no page links to is invisible."""
    client, _mcp = admin_client
    html = client.get("/admin/").text
    for s in SECTIONS:
        assert f'href="/admin/{s.slug}"' in html, f"{s.slug} is not on the dashboard"


def test_the_nav_covers_every_section():
    items = nav_items(lambda p="": f"/admin{p}")
    assert [label for label, _href, _active in items] == [s.label for s in SECTIONS]


def test_the_nav_is_empty_when_not_authenticated():
    assert nav_items(lambda p="": f"/admin{p}", authed=False) == ()


def test_tab_items_default_to_the_first_tab():
    s = sec_mod.section("word")
    actives = [active for _l, _h, active in tab_items(s, lambda p="": p)]
    assert actives == [True] + [False] * (len(s.tabs) - 1)


# ---------------------------------------------------------------------------
# Moved URLs
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("old,target", sorted(MOVED_PATHS.items()))
def test_a_moved_page_redirects_to_its_tab(admin_client, old, target):
    client, _mcp = admin_client
    slug, tab = target
    r = client.get(f"/admin{old}", follow_redirects=False)
    assert r.status_code == 303
    assert r.headers["location"] == sec_mod.section(slug).href(
        lambda p="": f"/admin{p}", tab)
    # And the target actually exists, so the redirect is not to a 404.
    assert client.get(r.headers["location"]).status_code == 200


# ---------------------------------------------------------------------------
# Nothing a section route shadows
# ---------------------------------------------------------------------------

def _sample_path(path: str) -> str:
    """A concrete path for a parameterised route, for match testing."""
    return re.sub(r"\{[^}]+\}", "x", path)


def test_no_route_is_shadowed_by_a_section_route(admin_client):
    """Every handler registered after the sections is still reachable.

    Starlette matches in registration order and the section routes come first,
    so this is the invariant that keeps ``POST /admin/email/save`` reaching the
    save handler rather than rendering the Email section. It is checked against
    the real route table: a change that preserves the property another way
    passes, and one that breaks it fails here rather than in production.
    """
    client, _mcp = admin_client
    routes = list(_admin_routes(client.app))
    slugs = {s.slug for s in SECTIONS}
    section_routes = [
        r for r in routes
        if r.path.lstrip("/").split("/")[0] in slugs
        and not r.path.lstrip("/").split("/")[0] == ""
    ]
    assert section_routes, "no section routes found — the walk is wrong"
    first_section = routes.index(section_routes[0])

    for later in routes[first_section:]:
        if later in section_routes:
            continue
        sample = _sample_path(later.path)
        for method in sorted(later.methods or {"GET"}):
            if method in ("HEAD", "OPTIONS"):
                continue
            for earlier in section_routes:
                if routes.index(earlier) > routes.index(later):
                    continue
                match, _scope = earlier.matches(
                    {"type": "http", "path": sample, "method": method,
                     "path_params": {}})
                assert match.value < 2, (
                    f"{method} {sample} ({later.name}) is shadowed by the "
                    f"section route {earlier.path}")


def test_a_post_to_a_kind_route_is_not_eaten_by_its_section(admin_client):
    """The concrete case: 'email' is both a section slug and a template kind."""
    client, _mcp = admin_client
    assert "email" in {s.slug for s in SECTIONS}
    # No CSRF token, so the handler rejects it — but it is the *save handler*
    # rejecting it, not the section page rendering a 200.
    r = client.post("/admin/email/save", data={"name": "x"},
                    follow_redirects=False)
    assert r.status_code == 403, r.text[:200]


def test_the_new_template_page_is_not_shadowed(admin_client):
    """'/new/docx' must not be read as the tab 'docx' of a section 'new'."""
    client, _mcp = admin_client
    html = client.get("/admin/new/docx").text
    assert "New Word template" in html


# ---------------------------------------------------------------------------
# The guard that keeps a new section from shadowing a page
# ---------------------------------------------------------------------------

def test_slugs_free_passes_for_the_shipped_sections():
    assert_slugs_free() is None


def test_slugs_free_refuses_a_slug_that_shadows_a_page():
    with pytest.raises(RuntimeError, match="word"):
        assert_slugs_free(("word",))
