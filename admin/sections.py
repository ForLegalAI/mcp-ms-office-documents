"""One descriptor per *product section* — the admin UI's top-level navigation.

The UI used to be organised by function: one page listing every template of
every kind, one page holding all five base templates, one page for the Word
style mapping. Everything about Word was therefore spread across three pages
and nothing about Excel had a page at all. This module organises it by the
thing an admin is actually working on — Word, PowerPoint, Excel, Email, XML —
with the functional views as *tabs* inside each section.

It owns nothing of its own. A section is a join over the two tables that
already existed:

=========================  ====================================================
:mod:`admin.kinds`         the dynamic template kinds (docx / email / pptx) —
                           the templates that become MCP tools, or values for
                           the presentation tool's ``template`` argument
:mod:`admin.base_templates`  the fixed-filename base templates that style every
                           generated document
=========================  ====================================================

A section may have either, both or neither: Word has a template kind, a base
slot and the global style mapping; Excel has only a base slot; XML has neither
and exists so that every tool the server exposes is accounted for somewhere.

Adding a section means adding a row to :data:`SECTIONS` — the nav, the routes
and the tab bars are all derived from it, so there is nowhere else to
remember. The same promise :mod:`admin.kinds` makes for kinds.

.. warning::
   A section's :attr:`Section.slug` occupies the same first path segment as a
   template *kind* (``/admin/word`` beside ``/admin/docx/save``) and as the
   admin's other literal pages (``/admin/new/docx``). Two rules keep that from
   becoming a shadowed route, because the section routes are registered first:

   * the section pages are registered **GET-only**, and every two-segment
     ``/{kind}/...`` route is a POST — which is what lets the Email section
     live at ``/email`` beside the ``email`` kind's own handlers; and
   * :func:`assert_slugs_free` refuses a slug that is one of the literal first
     segments the app already serves, failing at startup rather than quietly
     breaking that page.

   ``tests/test_admin_sections.py`` walks the real route table and fails if
   anything registered after the section routes is shadowed by one. That is
   the actual invariant; the two rules above are how it is currently met.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Sequence, Tuple

from admin import base_templates
from admin.store import KIND_DOCX, KIND_EMAIL, KIND_PPTX

#: Tab slugs with a fixed meaning. A section lists the ones it has.
TAB_OVERVIEW = "overview"
TAB_TEMPLATES = "templates"
TAB_BASE = "base"
TAB_STYLES = "styles"
TAB_STATUS = "status"
TAB_FILES = "files"
TAB_LOG = "log"


@dataclass(frozen=True)
class Tab:
    """One tab inside a section: its slug, its label, and a one-line blurb."""

    slug: str
    label: str
    blurb: str = ""


@dataclass(frozen=True)
class Section:
    """One product area of the admin UI."""

    slug: str                       # "word" — the URL segment and nav key
    label: str                      # "Word"
    icon: str                       # "📝"
    blurb: str                      # one line under the page title
    tabs: Tuple[Tab, ...]
    #: The dynamic template kind this section manages, where it has one.
    kind: Optional[str] = None
    #: Base-template slot keys shown on this section's Base tab, in order.
    slot_keys: Tuple[str, ...] = ()
    #: The static MCP tools this section's documents come from, for Overview.
    tools: Tuple[str, ...] = ()
    #: Reference page under docs/, linked from Overview.
    doc: str = ""
    #: True for the server-maintenance section, which is not a document type.
    is_server: bool = False

    def __post_init__(self):
        if not self.tabs:  # pragma: no cover - a programming error, not input
            raise ValueError(f"Section {self.slug!r} declares no tabs")

    @property
    def default_tab(self) -> str:
        return self.tabs[0].slug

    @property
    def slots(self) -> Tuple[base_templates.BaseSlot, ...]:
        """This section's base slots, resolved from the slot table."""
        return tuple(base_templates.slot(k) for k in self.slot_keys)

    def has_tab(self, slug: str) -> bool:
        return any(t.slug == slug for t in self.tabs)

    def tab(self, slug: str) -> Tab:
        for t in self.tabs:
            if t.slug == slug:
                return t
        raise KeyError(f"Section {self.slug!r} has no tab {slug!r}")

    def href(self, url, tab: Optional[str] = None) -> str:
        """URL for this section, or for one of its tabs."""
        if tab is None or tab == self.default_tab:
            return url(f"/{self.slug}")
        return url(f"/{self.slug}/{tab}")

    @property
    def title(self) -> str:
        return f"{self.icon} {self.label}"


_WORD = Section(
    slug="word",
    label="Word",
    icon="📝",
    blurb="Everything that shapes a .docx: the templates the AI can call, the "
          "base document they are rendered onto, and the styles markdown maps to.",
    tabs=(
        Tab(TAB_OVERVIEW, "Overview",
            "What this server can produce as a Word document, and how it is "
            "being used."),
        Tab(TAB_TEMPLATES, "Templates",
            "Each one becomes an MCP tool of its own, with the arguments you "
            "declare here."),
        Tab(TAB_BASE, "Base template",
            "The document every Word file is built on — its styles, fonts, "
            "margins and page setup."),
        Tab(TAB_STYLES, "Style mapping",
            "Which Word style each markdown block is rendered as, for every "
            "Word document this server produces."),
    ),
    kind=KIND_DOCX,
    slot_keys=("docx",),
    tools=("create_word_document",),
    doc="docs/markdown-reference.md",
)

_PPTX = Section(
    slug="powerpoint",
    label="PowerPoint",
    icon="📊",
    blurb="The designs decks are built on. A PowerPoint template is not a tool "
          "of its own — it is one more value for the presentation tool's "
          "'template' argument.",
    tabs=(
        Tab(TAB_OVERVIEW, "Overview",
            "What this server can produce as a deck, and how it is being used."),
        Tab(TAB_TEMPLATES, "Templates",
            "Registered designs. Each is a choice for the 'template' argument."),
        Tab(TAB_BASE, "Base designs",
            "The two designs used when a call names no template at all."),
    ),
    kind=KIND_PPTX,
    slot_keys=("pptx_16_9", "pptx_4_3"),
    tools=("create_powerpoint_presentation", "list_presentation_templates"),
    doc="docs/powerpoint-slides.md",
)

_XLSX = Section(
    slug="excel",
    label="Excel",
    icon="📈",
    blurb="Workbooks are built from markdown tables rather than from a "
          "template, so the one thing to manage here is the set of named cell "
          "styles a 'styles:' directive can reference.",
    tabs=(
        Tab(TAB_OVERVIEW, "Overview",
            "What this server can produce as a workbook, and how it is being "
            "used."),
        Tab(TAB_BASE, "Named styles",
            "The workbook whose named cell styles a 'styles:' directive can "
            "apply as style:<Name>."),
    ),
    slot_keys=("xlsx",),
    tools=("create_excel_document",),
    doc="docs/markdown-reference.md",
)

_EMAIL = Section(
    slug="email",
    label="Email",
    icon="✉️",
    blurb="Email templates become MCP tools of their own; the wrapper is the "
          "HTML every generated email is placed inside.",
    tabs=(
        Tab(TAB_OVERVIEW, "Overview",
            "What this server can produce as an email, and how it is being "
            "used."),
        Tab(TAB_TEMPLATES, "Templates",
            "Each one becomes an MCP tool of its own, with the arguments you "
            "declare here."),
        Tab(TAB_BASE, "Base wrapper",
            "The HTML wrapper around every generated email."),
    ),
    kind=KIND_EMAIL,
    slot_keys=("email",),
    tools=("create_email_draft",),
    doc="docs/development/tools/email.md",
)

_XML = Section(
    slug="xml",
    label="XML",
    icon="🧩",
    blurb="XML documents are written from the structure the caller supplies "
          "and validated on the way out. There is nothing to configure — this "
          "page is here so the tool is not invisible.",
    tabs=(
        Tab(TAB_OVERVIEW, "Overview",
            "What this server can produce as XML, and how it is being used."),
    ),
    tools=("create_xml_document",),
    doc="docs/development/tools/xml.md",
)

_SERVER = Section(
    slug="server",
    label="Server",
    icon="⚙️",
    blurb="How this server is doing: what has been called, what builds worked "
          "around, what is on disk, and what has been logged.",
    tabs=(
        Tab(TAB_STATUS, "Status",
            "Live counts and what every tool has done this session."),
        Tab(TAB_FILES, "Source files",
            "Every file in the uploads directory, and what still points at it."),
        Tab(TAB_LOG, "Activity log",
            "A filterable tail of what this server has logged."),
    ),
    is_server=True,
)

#: Every section, in nav order. Document types first, the server last.
SECTIONS: Tuple[Section, ...] = (_WORD, _PPTX, _XLSX, _EMAIL, _XML, _SERVER)

_BY_SLUG = {s.slug: s for s in SECTIONS}

#: The slug of the section a kind's template pages belong to, for breadcrumbs
#: and for the "back" link out of an editor.
_BY_KIND = {s.kind: s for s in SECTIONS if s.kind}


def section(slug: str) -> Section:
    """The section for *slug*, or ``KeyError``."""
    return _BY_SLUG[slug]


def is_section(slug: str) -> bool:
    return slug in _BY_SLUG


def section_for_kind(kind: str) -> Section:
    """The section whose Templates tab lists *kind*.

    Every editor page belongs to a section — that is what its breadcrumb and
    its Cancel link point at — so a kind with no section is a programming
    error rather than something to fall back from.
    """
    return _BY_KIND[kind]


def slot_section(key: str) -> Section:
    """The section whose Base tab holds base slot *key*."""
    for s in SECTIONS:
        if key in s.slot_keys:
            return s
    raise KeyError(f"No section holds base slot {key!r}")


#: First path segments the admin serves that are not sections. A slug taken
#: from this set would shadow that page, because the section routes are
#: registered first.
RESERVED_SEGMENTS: Tuple[str, ...] = (
    "new", "base", "files", "styles", "login", "logout",
)


def assert_slugs_free(reserved: Sequence[str] = RESERVED_SEGMENTS) -> None:
    """Fail if a section slug is one of the admin's other literal pages.

    Called where the routes are built, so a slug that would shadow ``/new/…``
    or ``/base/…`` stops the server at startup rather than quietly breaking
    one page. Template *kinds* are deliberately not checked: a section page is
    registered GET-only and every two-segment ``/{kind}/…`` route is a POST,
    which is why the Email section can live at ``/email`` beside the ``email``
    kind's handlers. ``tests/test_admin_sections.py`` is what pins that.
    """
    clash = sorted(set(reserved) & set(_BY_SLUG))
    if clash:
        raise RuntimeError(
            "Admin section slug(s) collide with a page the server already "
            "serves: " + ", ".join(clash)
            + ". A section would shadow it; rename the section."
        )


def nav_items(url, active: str = "", authed: bool = True):
    """``(label, href, active)`` triples for the top bar.

    The "New …" links the bar used to carry are gone: creating a template
    happens on the section's Templates tab, beside the list it adds to, which
    is both where it is looked for and how the bar stays at six items.
    """
    if not authed:
        return ()
    return [(s.label, s.href(url), s.slug == active) for s in SECTIONS]


def tab_items(s: Section, url, active: str = ""):
    """``(label, href, active)`` triples for one section's tab bar."""
    active = active or s.default_tab
    return [(t.label, s.href(url, t.slug), t.slug == active) for t in s.tabs]


__all__ = [
    "Section", "Tab", "SECTIONS", "RESERVED_SEGMENTS", "TAB_OVERVIEW",
    "TAB_TEMPLATES", "TAB_BASE", "TAB_STYLES", "TAB_STATUS", "TAB_FILES",
    "TAB_LOG", "section", "is_section", "section_for_kind", "slot_section",
    "assert_slugs_free", "nav_items", "tab_items",
]
