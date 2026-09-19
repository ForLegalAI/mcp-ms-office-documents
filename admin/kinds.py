"""One descriptor per template kind, for the admin views.

:mod:`admin.store` owns where a kind is *stored* — its spec directory, accepted
extensions and the spec key naming its asset. This module owns how a kind is
*presented*: its label and icon, the noun for the thing it produces, the wording
on its forms and confirmations, and whether it declares arguments at all.

The two halves are deliberately separate — ``store`` stays import-light and
FastHTML-free — but a view should only ever need this one lookup, so
:class:`KindDescriptor` re-exposes the storage fields as properties. That is
what ``store.kind_meta``'s docstring promises when it says adding a kind means
editing one table: the table for storage is there, the table for presentation is
here, and neither is a pile of ``if kind == KIND_PPTX`` branches scattered
through the views.

The split that drives most of the wording: a **docx or email** template is a
parameterised document — it declares arguments and becomes an MCP tool of its
own, so "live" means a tool is registered. A **pptx** template is a design —
layouts, theme, aspect, no arguments — and becomes one more value for the
``template`` argument of ``create_powerpoint_presentation``, so "live" means the
registry has re-read it. :attr:`KindDescriptor.has_args` is the flag the views
branch on where behaviour genuinely differs.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Sequence, Tuple

from admin.store import KIND_DOCX, KIND_EMAIL, KIND_PPTX, kind_meta

# Argument types offered in the argument editor.
ARG_TYPES = ("string", "int", "float", "bool", "list")

# Style-mapping keys surfaced in the UI (a subset of the set
# docx_tools.style_map recognises — see #160).
STYLE_KEYS = ("heading_1", "list_number", "list_bullet", "quote", "table")


@dataclass(frozen=True)
class KindDescriptor:
    """Everything the views need to know about one template kind."""

    kind: str
    label: str                  # "Word"
    icon: str                   # "📝"
    has_args: bool              # declares arguments and becomes a tool of its own
    name_label: str             # "Tool name" / "Template name"
    name_placeholder: str
    name_hint: str
    detail_header: str          # the middle column of the template table
    file_help: str              # how the source file is authored
    draft_hint: str             # flash after a successful upload+analysis
    save_ok: str                # confirmation when registration succeeded
    save_warn: str              # confirmation when it did not
    preview_label: str
    preview_hint: str = ""
    section_blurb: str = ""     # extra prose under the index section heading

    # -- storage metadata, delegated so a view needs one lookup --------------

    @property
    def _meta(self) -> Dict[str, Any]:
        return kind_meta(self.kind)

    @property
    def subdir(self) -> str:
        return self._meta["subdir"]

    @property
    def asset_ext(self) -> str:
        """The canonical extension, used when deriving a filename from a name."""
        return self._meta["asset_ext"]

    @property
    def asset_exts(self) -> Tuple[str, ...]:
        """Every extension accepted on upload."""
        return tuple(self._meta["asset_exts"])

    @property
    def master_file(self) -> str:
        """The hand-written master YAML's filename, beside the ``*.d`` dir."""
        subdir = self.subdir
        stem = subdir[:-2] if subdir.endswith(".d") else subdir
        return f"{stem}.yaml"

    @property
    def path_key(self) -> str:
        """The spec key naming the asset file (``docx_path``, ``pptx_path``…)."""
        return self._meta["path_key"]

    @property
    def accept(self) -> str:
        """The accepted extensions as an HTML ``accept`` list."""
        return ",".join(self.asset_exts)

    def detail_cell(self, spec: Dict[str, Any]) -> str:
        """The middle column for *spec* in the template table.

        Argument count where a kind has arguments; for a presentation template,
        which declares none, what distinguishes one is whether it is the default.
        """
        if self.has_args:
            return str(len(spec.get("args") or []))
        return "★ default" if spec.get("default") else "—"


_DOCX = KindDescriptor(
    kind=KIND_DOCX,
    label="Word",
    icon="📝",
    has_args=True,
    name_label="Tool name",
    name_placeholder="e.g. formal_letter",
    name_hint="Letters, digits and underscores — this is what the AI calls.",
    detail_header="Args",
    file_help="Word",
    draft_hint="Analyzed {filename} — review the arguments below, preview, then save.",
    save_ok="Saved — the tool '{name}' is now live and ready for the AI to use.",
    save_warn="Saved, but the tool '{name}' could not be registered (check the logs).",
    preview_label="Preview",
)

_EMAIL = KindDescriptor(
    kind=KIND_EMAIL,
    label="Email",
    icon="✉️",
    has_args=True,
    name_label="Tool name",
    name_placeholder="e.g. welcome_email",
    name_hint="Letters, digits and underscores — this is what the AI calls.",
    detail_header="Args",
    file_help="any HTML editor",
    draft_hint="Analyzed {filename} — review the arguments below, preview, then save.",
    save_ok="Saved — the tool '{name}' is now live and ready for the AI to use.",
    save_warn="Saved, but the tool '{name}' could not be registered (check the logs).",
    preview_label="Preview",
)

_PPTX = KindDescriptor(
    kind=KIND_PPTX,
    label="PowerPoint",
    icon="📊",
    has_args=False,
    name_label="Template name",
    name_placeholder="e.g. corporate_16_9",
    name_hint="Letters, digits and underscores — this is the value passed as the "
              "'template' argument when generating a deck.",
    detail_header="Default",
    file_help="PowerPoint",
    draft_hint="Analyzed {filename} — check the layouts below, preview a sample deck, "
               "then save.",
    # No tool is created for a presentation template, so calling one "live"
    # would be a lie; what changed is the set of templates the presentation
    # tool can build on.
    save_ok="Saved — '{name}' is now one of the templates the presentation tool "
            "can build on.",
    save_warn="Saved, but '{name}' did not come back out of the template registry — "
              "check that its file is in custom_templates/ and see the logs.",
    preview_label="Preview sample deck",
    preview_hint="Preview builds a six-slide sample deck on this template, using the "
                 "layout choices above — including ones you have not saved yet.",
    section_blurb="Designs the presentation tool can build on. These are not tools of "
                  "their own — they are choices for its 'template' argument.",
)

DESCRIPTORS: Dict[str, KindDescriptor] = {d.kind: d for d in (_DOCX, _EMAIL, _PPTX)}

#: Every supported kind, in the order the index page shows them.
KINDS: Tuple[str, ...] = tuple(DESCRIPTORS)

#: Kinds offered as "New …" in the top bar — all of them. PowerPoint used to be
#: missing, which left its create page unreachable once one template existed,
#: because the only other link was the template table's empty state (#157).
NAV_KINDS: Tuple[str, ...] = KINDS


def descriptor(kind: str) -> KindDescriptor:
    """The descriptor for *kind*, or ``KeyError`` if it is not a known kind."""
    return DESCRIPTORS[kind]


def is_kind(kind: str) -> bool:
    return kind in DESCRIPTORS


def nav_links(url, authed: bool = True) -> Sequence[Tuple[str, str]]:
    """``(label, href)`` pairs for the top bar, given a URL builder."""
    if not authed:
        return ()
    links = [("All templates", url("/"))]
    links += [(f"New {DESCRIPTORS[k].label}", url(f"/new/{k}")) for k in NAV_KINDS]
    links += [("Status", url("/status")), ("Log out", url("/logout"))]
    return links
