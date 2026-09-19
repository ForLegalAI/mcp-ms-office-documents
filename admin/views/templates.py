"""The template pages: the index, the new-template form, and the editor.

Three kinds share these views. A docx or email template is a parameterised
document, so its editor is an argument table; a pptx template is a design, so
its editor is a layout-role map and deck defaults. Where the two genuinely
differ they get their own function (:func:`edit_form` dispatches); where they
only differ in wording, the wording comes from
:class:`~admin.kinds.KindDescriptor` and the markup is shared.
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional

from fasthtml.common import (
    A, Button, Div, Form, H1, Input, Option, P, Select, Span, Textarea, Th, Tr, Td,
)

from admin import components as c
from admin.analysis import Analysis, PptxAnalysis, reconcile
from admin.kinds import ARG_TYPES, KINDS, STYLE_KEYS, descriptor
from admin.store import KIND_DOCX, KIND_PPTX
from admin.views.shell import form_actions, name_field, page

# ---------------------------------------------------------------------------
# Template table
# ---------------------------------------------------------------------------


def template_table(ctx, kind: str, csrf: str = ""):
    """Every managed template of *kind*, plus read-only ones from the master YAML."""
    d = descriptor(kind)
    managed = ctx.store.list_specs(kind)
    managed_names = {s.get("name") for s in managed}
    live = set(ctx.live_names(kind))

    rows = []
    for spec in managed:
        name = spec.get("name")
        rows.append(Tr(
            Td(A(name, href=ctx.u(f"/{kind}/{name}/edit"))),
            Td(d.detail_cell(spec)),
            Td(c.status_badge(name in live)),
            Td(c.action_bar(
                A("Edit", href=ctx.u(f"/{kind}/{name}/edit"), cls="btn btn-secondary btn-sm"),
                c.post_form(
                    ctx.u(f"/{kind}/{name}/delete"),
                    Button("Delete", type="submit", cls="btn btn-danger btn-sm",
                           onclick="return confirm('Delete this template? This removes the "
                                   "live tool. The source file is kept in custom_templates/.')"),
                    csrf=csrf, cls="inline-form",
                ),
            )),
        ))
    # Read-only templates from the master YAML (live but not UI-managed).
    for name in sorted(live - managed_names):
        rows.append(Tr(
            Td(name), Td("—"),
            Td(c.badge("● Live", "live")),
            Td(c.badge("from master YAML — read-only", "ro")),
        ))

    if not rows:
        return c.empty_state(
            f"No {d.label} templates yet.",
            A(f"{d.icon} Create your first {d.label} template",
              href=ctx.u(f"/new/{kind}"), cls="btn btn-primary"),
        )
    return c.data_table(["Name", d.detail_header, "Status", "Actions"], rows)


def index_page(ctx, csrf: str = ""):
    sections = []
    for kind in KINDS:
        d = descriptor(kind)
        body = [P(d.section_blurb, cls="muted")] if d.section_blurb else []
        body.append(template_table(ctx, kind, csrf))
        sections.append(c.card(*body, title=f"{d.icon} {d.label} templates", level=2))
    return page(
        ctx, "Templates",
        H1("Templates"),
        P("Create reusable Word and email templates — each becomes a tool the AI can "
          "call — and register PowerPoint templates for it to build decks on.",
          cls="muted"),
        *sections,
    )


# ---------------------------------------------------------------------------
# Argument editor
# ---------------------------------------------------------------------------


def arg_row(arg: Optional[Dict[str, Any]] = None, cond: bool = False):
    """One editable argument row (parallel-indexed fields + remove button).

    The fields carry no ids: they repeat once per row under the same names, so
    an id would not be unique. The column headers label them instead.
    """
    arg = arg or {}
    name = arg.get("name", "")
    atype = str(arg.get("type", "string")).lower()
    required = bool(arg.get("required", True))
    default = arg.get("default", "")
    desc = arg.get("description", "")

    type_opts = [Option(t, value=t, selected=(t == atype)) for t in ARG_TYPES]
    req_opts = [
        Option("required", value="true", selected=required),
        Option("optional", value="false", selected=not required),
    ]
    name_cell = [Input(name="arg_name", value=name, placeholder="argument_name")]
    if cond:
        name_cell.append(c.badge("if", "if", title="Conditional flag"))
    return Tr(
        Td(*name_cell, cls="col-name"),
        Td(Select(*type_opts, name="arg_type"), cls="col-type"),
        Td(Select(*req_opts, name="arg_required"), cls="col-req"),
        Td(Input(name="arg_default", value="" if default in (None,) else str(default),
                 placeholder="optional"), cls="col-def"),
        Td(Input(name="arg_desc", value=desc, placeholder="what this is, for the AI")),
        Td(Button("✕", type="button", cls="btn-icon", title="Remove",
                  onclick="adminRemoveRow(this)"), cls="col-x"),
        cls="arg-row",
    )


def _arg_rows(spec: Dict[str, Any], analysis: Optional[Analysis]):
    """Declared args first, then detected-but-undeclared placeholders.

    No blank spares — users add rows with the "Add argument" button.
    """
    cond_set = set(analysis.conditionals) if analysis else set()
    existing = spec.get("args") or []
    existing_names = {a.get("name") for a in existing if isinstance(a, dict)}
    rows = [arg_row(a, cond=a.get("name") in cond_set) for a in existing]
    if analysis:
        for ph in analysis.placeholders:
            if ph not in existing_names:
                is_cond = ph in cond_set
                rows.append(arg_row(
                    {"name": ph, "type": "bool" if is_cond else "string",
                     "required": not is_cond, "description": ""}, cond=is_cond))
                existing_names.add(ph)
        for cond in analysis.conditionals:
            if cond not in existing_names:
                rows.append(arg_row({"name": cond, "type": "bool", "required": False,
                                     "description": ""}, cond=True))
                existing_names.add(cond)
    return rows or [arg_row()]


def style_mapping_block(analysis: Optional[Analysis], spec: Dict[str, Any]):
    styles = (analysis.styles_present if analysis else []) or []
    current = (spec or {}).get("style_mapping") or {}
    selects = []
    for key in STYLE_KEYS:
        cur = current.get(key, "")
        opts = [Option("(use built-in)", value="__default__", selected=not cur)]
        opts += [Option(s, value=s, selected=(s == cur)) for s in styles]
        selects.append(c.field(key, Select(*opts, name=f"style_{key}")))
    return c.details_block(
        "Advanced — map markdown styles to your template's own style names",
        P("Only needed if your Word template renames the built-in styles "
          "(Heading 1, List Number, …). Otherwise leave these alone.", cls="muted"),
        *selects,
    )


# ---------------------------------------------------------------------------
# Analysis reports
# ---------------------------------------------------------------------------


def analysis_report(analysis, spec: Dict[str, Any]):
    if isinstance(analysis, PptxAnalysis):
        return pptx_analysis_report(analysis)
    rec = reconcile(analysis, spec.get("args") or [])
    cond_set = set(analysis.conditionals)
    # c.static_row, not c.field: these are read-only reports, not controls.
    items: List[Any] = [
        c.static_row("Detected placeholders", c.chips(analysis.placeholders)),
        c.static_row("Detected conditionals",
                     c.chips(analysis.conditionals, cond_set)),
    ]
    if analysis.missing_required_styles:
        items.append(c.flash(
            "⚠ The document is missing some Word styles the renderer uses: "
            + ", ".join(analysis.missing_required_styles)
            + ". Lists/headings using them will fall back to the default style.", "warn"))
    for w in analysis.warnings:
        items.append(c.flash("⚠ " + w, "warn"))
    if rec.orphan_args:
        items.append(c.flash("Args with no matching placeholder (they'll be ignored by "
                             "the document): " + ", ".join(rec.orphan_args), "warn"))
    if rec.non_bool_conditions:
        items.append(c.flash("Conditional flags should be type 'bool': "
                             + ", ".join(rec.non_bool_conditions), "warn"))
    return c.card(*[i for i in items if i is not None],
                  title="What we found in the document")


def pptx_analysis_report(analysis: PptxAnalysis):
    """What we found in a PowerPoint template: size, layouts, roles, theme."""
    rows = [
        Tr(
            Td(layout.name),
            Td(c.chip(layout.role) if layout.role
               else Span("name it explicitly", cls="muted")),
            Td(c.chips(layout.placeholders)),
            Td("—" if layout.has_footer else Span("no footer", cls="muted")),
        )
        for layout in analysis.layouts
    ]

    facts = Div(
        c.static_row("Slide size",
                    Span(f"{analysis.slide_size} ({analysis.aspect})"
                         if analysis.slide_size else "unknown")),
        c.static_row("Theme fonts",
                    Span(", ".join(f"{k}: {v}" for k, v in analysis.theme_fonts.items())
                         or "not declared",
                         cls="muted" if not analysis.theme_fonts else "")),
        c.static_row("Theme colours",
                    Span(*[c.swatch(k, v) for k, v in analysis.theme_colors.items()])
                    if analysis.theme_colors else Span("not declared", cls="muted")),
        cls="facts",
    )

    items: List[Any] = [facts]
    for w in analysis.warnings:
        items.append(c.flash("⚠ " + w, "warn"))
    if rows:
        items.append(c.data_table(
            ["Layout", "Detected role", "Placeholders", "Footer"], rows))
    return c.card(*items, title="What we found in the template")


# ---------------------------------------------------------------------------
# Edit forms
# ---------------------------------------------------------------------------


def edit_form(ctx, kind: str, spec: Dict[str, Any],
              analysis, is_new: bool, csrf: str = ""):
    if kind == KIND_PPTX:
        return _pptx_edit_form(ctx, spec, analysis, is_new, csrf)
    return _document_edit_form(ctx, kind, spec, analysis, is_new, csrf)


def _form_shell(ctx, kind: str, spec: Dict[str, Any], is_new: bool,
                csrf: str, asset_filename: str, *cards):
    """The ``<form>`` every editor shares: hidden state, cards, action bar."""
    return Form(
        c.csrf_input(csrf),
        c.hidden("kind", kind),
        c.hidden("asset_filename", asset_filename),
        c.hidden("name", spec.get("name", "")) if not is_new else None,
        *cards,
        form_actions(ctx, kind),
        action=ctx.u(f"/{kind}/save"), method="post",
    )


def _document_edit_form(ctx, kind: str, spec: Dict[str, Any],
                        analysis: Optional[Analysis], is_new: bool, csrf: str = ""):
    d = descriptor(kind)
    annotations = spec.get("annotations") or {}

    details = c.card(
        name_field(kind, spec.get("name", ""), is_new),
        c.field("Title", Input(name="title", value=annotations.get("title", ""),
                               placeholder="Friendly name shown to the user")),
        c.field("Description",
                Textarea(spec.get("description", ""), name="description", rows="3",
                         placeholder="Tell the AI when and how to use this template."),
                hint="This is the instruction the AI sees when choosing the tool."),
        title="Tool details",
    )

    args = c.card(
        P("Every placeholder and conditional in your document needs an argument. "
          "We pre-filled them from the document — adjust types and descriptions, "
          "then save.", cls="muted"),
        c.data_table(
            [Th("Name", cls="col-name"), Th("Type", cls="col-type"),
             Th("Required", cls="col-req"), Th("Default", cls="col-def"),
             Th("Description"), Th("", cls="col-x")],
            _arg_rows(spec, analysis),
            cls="args-table", body_id="argrows",
        ),
        Div(Button("+ Add argument", type="button", cls="btn btn-secondary btn-sm",
                   onclick="adminAddArgRow()"), style="margin-top:.6rem"),
        title="Arguments",
    )

    cards = [details, args]
    if kind == KIND_DOCX:
        cards.append(c.card(style_mapping_block(analysis, spec)))
    return _form_shell(ctx, kind, spec, is_new, csrf, spec.get(d.path_key, ""), *cards)


def _pptx_edit_form(ctx, spec: Dict[str, Any], analysis: Optional[PptxAnalysis],
                    is_new: bool, csrf: str = ""):
    """The PowerPoint template form.

    Deliberately not the argument editor. A presentation template declares no
    arguments; what an admin needs to control is which layout plays each slide
    role, and the deck-wide defaults.
    """
    from pptx_tools.layouts import ROLES

    kind = KIND_PPTX
    configured = spec.get("layouts") or {}
    defaults = spec.get("defaults") or {}
    detected = analysis.role_map if analysis else {}
    layout_names = analysis.layout_names if analysis else []

    details = c.card(
        name_field(kind, spec.get("name", ""), is_new),
        c.field("Description",
                Textarea(spec.get("description", ""), name="description", rows="3",
                         placeholder="When should the AI reach for this deck? "
                                     "e.g. 'Brand deck, widescreen. Use for "
                                     "client-facing work.'"),
                hint="This is what the AI sees when choosing between templates."),
        c.checkbox_field("is_default", "Use this template when none is named",
                         checked=bool(spec.get("default"))),
        c.checkbox_field("strip_slides",
                         "Drop any slides the template file itself contains",
                         checked=bool(spec.get("strip_slides", True)),
                         hint="Sample slides a designer left in the file never belong "
                              "in generated output. Leave this on unless you know "
                              "otherwise."),
        title="Template details",
    )

    # One dropdown per role, populated from the template's real layout names.
    role_fields = []
    for role in ROLES:
        current = configured.get(role, "")
        auto = detected.get(role)
        opts = [Option(f"(auto-detected: {auto})" if auto else "(not detected)",
                       value="", selected=not current)]
        opts += [Option(n, value=n, selected=(n == current)) for n in layout_names]
        role_fields.append(c.field(
            role, Select(*opts, name=f"layout_{role}"),
            hint=None if auto else "No layout matches this role — slides needing it "
                                   "fall back by position unless you pick one.",
            hint_cls="muted warn-text",
        ))

    layouts = c.card(
        P("Layouts are matched automatically by their placeholders, so most templates "
          "need nothing here. Override a role only when the automatic choice is wrong, "
          "or when it says 'not detected'.", cls="muted"),
        Div(*role_fields, cls="role-grid"),
        title="Layout for each slide role",
    )

    deck_defaults = c.card(
        P("Applied when the tool call does not set the same option.", cls="muted"),
        c.field("Footer text",
                Input(name="default_footer_text",
                      value=defaults.get("footer_text", "") or "",
                      placeholder="e.g. ACME s.r.o. · Confidential")),
        c.field("Language",
                Input(name="default_language", value=defaults.get("language", "") or "",
                      placeholder="e.g. cs-CZ"),
                hint="BCP-47 tag stamped on every text run, so the deck is proof-read "
                     "in the right language."),
        c.checkbox_field("default_slide_numbers", "Show slide numbers",
                         checked=bool(defaults.get("show_slide_numbers"))),
        title="Deck defaults",
    )

    return _form_shell(ctx, kind, spec, is_new, csrf, spec.get("pptx_path", ""),
                       details, layouts, deck_defaults)


# ---------------------------------------------------------------------------
# Whole pages
# ---------------------------------------------------------------------------


def _authoring_help(kind: str):
    """How a template of this kind is authored, as a collapsible block."""
    d = descriptor(kind)
    if not d.has_args:
        return _pptx_help()
    blocks = [
        P("Author your document in ", d.file_help,
          " and insert placeholders like ", c.chip("{{recipient_name}}"),
          " wherever a value should go."),
    ]
    if kind == KIND_DOCX:
        blocks.append(
            P("Wrap optional content in ", c.chip("{{#if include_clause}}"),
              " … ", c.chip("{{/if}}"),
              " to show it only when a flag is set.", cls="muted"))
    blocks.append(
        P("When you upload, we scan the file, list every placeholder, and pre-build "
          "the argument form for you. Nothing goes live until you press Save.",
          cls="muted"))
    return c.details_block("How does this work?", *blocks)


def _pptx_help():
    """How a presentation template differs from the other two kinds."""
    return c.details_block(
        "How does this work?",
        P("A PowerPoint template is a ", c.chip("design"),
          ", not a fill-in-the-blanks document. It has no placeholders and takes no "
          "arguments: what it contributes is the slide master, the layouts, the theme "
          "fonts and the colour palette."),
        P("So this does not become a tool of its own, the way a Word template does. It "
          "becomes one more choice for the ", c.chip("template"),
          " argument of the presentation tool, alongside any others you register."),
        P("Design your deck in PowerPoint and save it as .pptx or .potx. Every layout "
          "you want used should carry the placeholders it needs — a title, a content "
          "box, a picture frame — and we work out which slide role each layout plays "
          "from those. Sample slides left in the file are dropped from generated "
          "decks.", cls="muted"),
        P("When you upload, we report the slide size, every layout with the role we "
          "detected for it, and the theme. Nothing goes live until you press Save.",
          cls="muted"),
    )


def new_page(ctx, kind: str, csrf: str = "", error: Optional[str] = None):
    d = descriptor(kind)
    return page(
        ctx, f"New {d.label} template",
        H1(f"{d.icon} New {d.label} template"),
        c.flash(error, "err"),
        c.card(c.post_form(
            ctx.u(f"/{kind}/draft"),
            c.field(d.name_label,
                    Input(name="name", required=True, placeholder=d.name_placeholder),
                    hint=d.name_hint),
            c.field(f"{d.label} file ({d.accept})",
                    Input(name="file", type="file", accept=d.accept, required=True)),
            Button("Upload & analyze", type="submit", cls="btn btn-primary"),
            csrf=csrf, enctype="multipart/form-data",
        )),
        _authoring_help(kind),
    )


def configure_page(ctx, kind: str, name: str, filename: str,
                   analysis, spec: Dict[str, Any], csrf: str = ""):
    """Shown straight after a successful upload, with the form pre-filled."""
    d = descriptor(kind)
    return page(
        ctx, f"Configure {name}",
        H1(f"Configure {name}"),
        c.flash(d.draft_hint.format(filename=filename), "ok"),
        analysis_report(analysis, spec),
        edit_form(ctx, kind, spec, analysis, is_new=False, csrf=csrf),
    )


def edit_page(ctx, kind: str, name: str, spec: Dict[str, Any], analysis,
              csrf: str = "", message: Optional[str] = None,
              message_kind: str = "ok"):
    from admin.views.shell import replace_card

    body: List[Any] = [H1(f"Edit {name}")]
    if message:
        body.append(c.flash(message, message_kind))
    if analysis is not None:
        body.append(analysis_report(analysis, spec))
    else:
        body.append(c.flash("The template's source file is missing — arguments can "
                            "still be edited.", "warn"))
    body.append(edit_form(ctx, kind, spec, analysis, is_new=False, csrf=csrf))
    body.append(replace_card(ctx, kind, name, csrf))
    return page(ctx, f"Edit {name}", *body)
