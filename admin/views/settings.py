"""The Word Style mapping tab: the ``style_mapping`` every Word document uses.

Deliberately not a template editor. There is nothing to name, no arguments and
no source file — one mapping, applying to the static ``markdown_to_word`` tool
and to every Word template that does not override the key itself.

Two things make this page different from the per-template style block, and both
are about not repeating the lie #161 is named after:

* it shows what a key is set to **now**, from the mapping actually in force
  (the master YAML under whatever this UI has stored), not from an empty form;
  and
* where the stored override and the hand-written master YAML disagree, it says
  so, because the override replaces the master's mapping outright rather than
  merging into it — so a master setting this page has switched off would
  otherwise simply vanish from view.
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence

from fasthtml.common import Button, Div, Option, P, Select

from admin import components as c
from admin.kinds import STYLE_GROUPS
from admin.views.templates import builtin_style_names

#: The select value meaning "do not set this key", shared with the per-template
#: editor so ``forms.parse_style_mapping`` reads both the same way.
UNSET = "__default__"


#: Appended to a style name the base Word template does not define. Says only
#: what is true — see :func:`_reach_note` for why it cannot say more.
MISSING_SUFFIX = " — not in the base Word template"


def _option_label(style: str, offered: Sequence[str]) -> str:
    """A style name, marked when the base Word template does not define it.

    A mapping may legitimately name a style only some other document has, and
    a value already in the config must never be dropped just because this page
    cannot offer it — so it is offered *and* flagged, rather than silently
    listed as though the base template provided it.
    """
    return style if style in offered else f"{style}{MISSING_SUFFIX}"


def _reach_note(mapping: Dict[str, str], offered: Sequence[str]):
    """Why a marked style is not necessarily a mistake.

    The dropdowns list the *base* Word template's styles, because that is what
    `markdown_to_word` renders onto. But this mapping also reaches every Word
    template tool, and each of those renders onto its **own** document — which
    may well define the style. So the marker means "missing here", and saying
    only that would leave an admin to read it as "missing everywhere", which is
    the shape of wrongness #161 is about.

    Only shown when something is actually marked; a note that is always there
    is a note nobody reads.
    """
    marked = sorted({v for v in mapping.values() if v and v not in offered})
    if not marked or not offered:
        return None
    return P(
        "Marked styles (" + ", ".join(marked) + ") are not in the base Word "
        "template, so documents rendered onto it fall back to the built-in. "
        "That is not necessarily wrong: a Word template that defines the style "
        "in its own document still uses it. It is only a mistake where neither "
        "document has it.", cls="muted")


def _style_select(key: str, current: str, offered: Sequence[str],
                  builtin: Dict[str, Optional[str]]):
    default = builtin.get(key)
    label = f"(use built-in: {default})" if default else "(use built-in: none)"
    opts = [Option(label, value=UNSET, selected=not current)]
    names = list(offered)
    if current and current not in names:
        names.append(current)
    opts += [Option(_option_label(s, offered), value=s, selected=(s == current))
             for s in names]
    return Select(*opts, name=f"style_{key}")


def _lower(mapping: Optional[Dict[str, Any]]) -> Dict[str, str]:
    """A style mapping keyed as the renderer keys it.

    ``build_style_map`` lower-cases and strips every key before matching, so a
    hand-written ``Heading_1:`` is live. Comparing or looking up the raw keys
    would show it as unset on a page whose whole job is saying what is set.
    """
    out: Dict[str, str] = {}
    for key, value in (mapping or {}).items():
        if value is None or str(value).strip() == "":
            continue
        out[str(key).strip().lower()] = str(value)
    return out


def _provenance(stored: bool, mapping: Dict[str, str], master: Dict[str, str],
                keys: Sequence[str], offered: Sequence[str]):
    """The card saying where the mapping on this page came from."""
    body: List[Any] = []
    if stored:
        body.append(c.flash(
            "This mapping is managed here. It replaces the master YAML's "
            "style_mapping entirely — not key by key — so anything the master "
            "sets and this page leaves alone is off.", "info"))
        # Only where they actually differ. A master key this page reproduces
        # is not something an admin has to know about; one it drops or changes
        # is the whole reason the master file can look wrong from the volume.
        differs = sorted(k for k in master
                         if k in keys and mapping.get(k) != master[k])
        if differs:
            body.append(c.static_row(
                f"Not in force: {len(differs)} key(s) docx_templates.yaml sets",
                c.chips([f"{k}: {master[k]} → {mapping.get(k) or 'built-in'}"
                         for k in differs])))
    else:
        body.append(c.flash(
            "Nothing is stored here yet, so the mapping below is what the "
            "hand-written config/docx_templates.yaml declares. Saving copies "
            "it into this server's managed config; the master file is never "
            "rewritten.", "info"))
        body.append(P(
            "Saving writes every field on this page, so reload it before "
            "saving if the master file may have been edited on the volume "
            "since it loaded — this form would otherwise store what it was "
            "rendered with.", cls="muted"))
    if not offered:
        body.append(P(
            "The base Word template's styles could not be read, so the "
            "dropdowns offer only what is already configured. Install a base "
            "Word template to pick from its styles.", cls="muted"))
    return c.card(*body, title="Where this comes from")


def _unrecognised(mapping: Dict[str, str], master: Dict[str, str],
                  keys: Sequence[str]):
    """A warning for configured keys the renderer does not act on.

    They already do nothing — ``style_map._normalize`` logs each one and drops
    it — but they are in a config file somebody wrote on purpose, so a page
    claiming to show what is set has to account for them rather than leave them
    to a log nobody is reading.

    Read from the master as well as from what is in force: once an override is
    stored the key is no longer in the effective mapping, but it is still
    sitting in ``docx_templates.yaml`` looking like a setting.
    """
    unknown = sorted({k for k in mapping if k not in keys}
                     | {k for k in master if k not in keys})
    if not unknown:
        return None
    return c.flash(
        "These keys are configured but are not ones the renderer acts on, so "
        "they change nothing: " + ", ".join(unknown) + ". This page cannot "
        "set them and the mapping it manages does not carry them.", "warn")


def style_mapping_panel(ctx, *, csrf: str, mapping: Dict[str, Any],
                        master: Dict[str, Any], stored: bool,
                        offered: Sequence[str] = (),
                        message: Optional[str] = None,
                        message_kind: str = "ok"):
    """The editor for the Word ``style_mapping`` that applies to everything.

    *mapping* is what is in force, *master* what the master YAML alone says,
    *stored* whether this UI has an override on file, and *offered* the styles
    the base Word template defines.

    A panel on the Word section's Style mapping tab. It was a top-level page
    called "Global styles", which said nothing about the mapping being Word's
    alone — there is no such thing as a global style for a spreadsheet.
    """
    builtin = builtin_style_names()
    keys = [key for _group, ks in STYLE_GROUPS for key in ks]
    mapping = _lower(mapping)
    master = _lower(master)

    groups = []
    for title, group_keys in STYLE_GROUPS:
        fields = [
            c.field(key, _style_select(key, mapping.get(key, ""),
                                       offered, builtin))
            for key in group_keys
        ]
        groups.append(Div(P(title, cls="group-title"),
                          Div(*fields, cls="role-grid")))
    reach = _reach_note(mapping, offered)
    if reach is not None:
        groups.append(reach)

    controls = [Button("Save & apply", type="submit", cls="btn btn-primary")]
    if stored:
        controls.append(Button(
            "Revert to docx_templates.yaml", type="submit",
            formaction=ctx.u("/styles/revert"), cls="btn btn-secondary",
            title="Discard the mapping managed here and use the master file's"))
    controls.append(Button("Reset", type="reset", cls="btn",
                           title="Undo unsaved changes on this page"))

    warn = _unrecognised(mapping, master, keys)
    return [
        c.flash(message, message_kind),
        warn,
        _provenance(stored, mapping, master, keys, offered),
        c.post_form(
            ctx.u("/styles/save"),
            c.card(*groups, title="Style names", level=2),
            c.card(c.action_bar(*controls)),
            csrf=csrf,
        ),
    ]
