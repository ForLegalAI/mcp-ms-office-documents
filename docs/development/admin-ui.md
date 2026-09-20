# Admin UI internals

The optional template-admin UI (`ADMIN_ENABLED`) is a FastHTML app mounted in
the same ASGI process as the MCP server, so saving a template takes effect with
no restart. This page is about **how it is built**. For what an admin can do
with it, see [`../admin-ui.md`](../admin-ui.md); for what a template spec may
contain, [`dynamic-templates.md`](dynamic-templates.md).

## Module map

| Module | Owns |
|--------|------|
| `admin/app.py` | `AdminContext` (the services a view needs), the routes, `build_admin_app()` / `build_combined_app()` — and nothing else |
| `admin/components.py` | markup primitives (`card`, `field`, `data_table`, …) and the inlined theme |
| `admin/kinds.py` | one `KindDescriptor` per *dynamic* template kind: label, icon, wording, `has_args` |
| `admin/base_templates.py` | one `BaseSlot` per *static* base template (fixed filename, one of each) |
| `admin/forms.py` | reading a submitted form back into a spec dict |
| `admin/assets.py` | what is in `custom_templates/` and what still references it |
| `admin/views/` | the pages — `shell`, `templates`, `base`, `assets`, `settings`, `status`, `login` |
| `admin/store.py` | persistence: `config/<kind>_templates.d/<name>.yaml` + the asset |
| `admin/analysis.py` | what is inside an uploaded `.docx` / `.html` / `.pptx` |
| `admin/preview.py` | rendering a template without touching the upload backend |
| `admin/auth.py` | the shared-password gate and CSRF tokens |

`components`, `kinds`, `forms`, `store`, `analysis` and `assets` do not import
`admin.app`; views take the `AdminContext` as a parameter. `store`, `analysis`
and `forms` have no FastHTML dependency at all, so their rules can be unit
tested without rendering anything.

A request is: route in `app.py` → load from `store` (and maybe `analysis`) →
a function in `views/` → primitives from `components.py`.

## The three invariants

**1. No external assets.** No CDN stylesheet, no `<script src>`, no web font,
no off-origin image — the UI must render correctly offline, in an air-gapped
deployment and behind a restrictive CSP. This is why `build_admin_app()`
constructs `FastHTML(..., default_hdrs=False, htmx=False, surreal=False)`:
FastHTML's own defaults are CDN-loaded. The theme and the ~8 lines of
JavaScript are inlined by `components.head_tags()`, which also injects a blank
argument row as `window.__ARG_ROW_HTML__` so "Add argument" can clone it
without a client-side templating library.

`tests/test_admin_assets.py` walks every page type and enforces this. It also
asserts the theme and script *are* present, so deleting them is not a way to
pass.

**2. Every control is labelled.** `components.field()` mints an `id` from the
control's `name` and points the `<label>` at it. Building a field by hand
silently loses the association, which is what the whole UI used to do. Use
`static_row()` for a read-only labelled value — there is no control for a label
to point at, and it must not be used for input.

**One gotcha inside that.** `views.edit_page()` takes the analysis twice over:
once to prefill the argument rows, once to render "What we found in the
document". The rejected-re-upload path passes `report=False` to get the first
without the second — the file on disk is still the old one, so it is the right
thing to prefill from and the wrong thing to describe under a heading that
means "the file you just uploaded" everywhere else it appears.

**Every kind's create page stays reachable.** `kinds.NAV_KINDS` puts a "New …"
entry in the top bar for each kind, and `template_table()` renders a create
link in *both* its states. That redundancy is deliberate: when only the empty
state linked to the page, adding a template removed the last route to it, which
was invisible for Word and Email (the top bar covered them) and a dead end for
PowerPoint (it did not) — #157.

**Deleting never removes a file something else uses.** Assets share one flat
`custom_templates/` directory, so several templates can name the same file.
`AdminContext.other_specs_using_asset()` checks both sources — the managed
`*.d` specs *and* the hand-written master YAML — and the delete page withholds
the "also delete the source file" option when anything else points at it; the
POST re-derives the answer server-side rather than trusting the form. A master
entry sharing the name being deleted counts too: the managed spec was
overriding it, so deleting the override revives the master definition, still
pointing at that file.

**The style editor offers exactly what the renderer recognises.**
`kinds.STYLE_KEYS` must match the keys `docx_tools/style_map.py` acts on —
offering fewer hides part of the feature (it offered 5 of 16 for a long time),
offering more promises a setting that silently does nothing.
`tests/test_admin_style_keys.py` compares the two lists directly, so adding a
key to the renderer fails the suite until the UI catches up. The labels come
from `DEFAULT_STYLE_MAP` for the same reason: a second hand-written copy of the
defaults would drift.

**A spec's filename is untrusted input.** `save_spec()` validates what the UI
writes, but a `*.d` spec is plain YAML a person can hand-write, so the filename
in it is not necessarily one this code produced. It reaches the filesystem
*and* a `Content-Disposition` header, so `validate_asset_filename()` rejects
quotes, backslashes and control characters (spaces and non-ASCII stay legal —
`Brand Deck.pptx` is an ordinary name), and `kinds.content_disposition()`
builds the header per RFC 6266 rather than interpolating the name.

**Base-template routes are registered before the generic ones.**
`/base/{slot}/download` also matches `/{kind}/{name}/download`, and Starlette
takes the first route that fits — so the `/base/…` routes are declared ahead of
the `/{kind}/…` block in `build_admin_app()`. Move them and the download
silently redirects home instead of serving a file;
`test_download_route_is_not_swallowed_by_the_generic_pattern` pins it.

**A slot's state is decided by path identity, not directory naming.**
`base_templates.source_of()` compares the resolved file against the exact path
a replacement is written to. `template_utils._classify_template_source()` looks
for a path part called `custom_templates`, which is right in the normal layout
and wrong when the bundled defaults sit under such a directory — it would
report a default as custom and offer to revert a file the UI cannot delete.

**One predicate decides whether a file is usable.** The analysers report two
different things: a file that could not be read at all, and observations about
a file that read fine. Only the first should refuse an upload, and
`analysis.is_unusable()` is what every upload route asks — it matches any
warning opening "Could not …". The routes used to match `"Could not open"`
inline, which missed `analyze_xlsx`'s "Could not read named styles" and
installed a workbook whose whole purpose had failed.

**An oversized upload is refused before it is read.** Starlette spools a
multipart upload to a temp file above 1 MB, so the request costs little memory
— it is `await upload.read()` that materialises it. `_read_upload()` therefore
checks `UploadFile.size` (falling back to a seek) *first*, and only reads a
file that is within `ADMIN_MAX_UPLOAD_MB`. The post-read check stays as a
backstop for an upload whose size could not be known in advance.

**The log view filters the whole buffer, then limits.** `metrics.recent_logs()`
takes the level, source and search and applies them before the `limit`, so the
limit counts *matches* rather than records scanned — limiting first would lose
a match that happens to sit behind the newest page and make a filter look
empty. The level choices come from `available_levels()`, which drops any level
below what the buffer is actually capturing at: offering "debug and above" on
a server running at INFO is a filter that can only ever come back empty.

**The scan snapshots the ring buffer before filtering it.** `recent_logs()`
and `counts_by_level()` both do `list(_LOG_HANDLER.records)` first. Builds run
on worker threads and each one logs, so a record can arrive mid-scan — and a
`deque` raises `RuntimeError: deque mutated during iteration` if it does.
Filter the snapshot, never the live deque.

**`refresh_seconds()` clamps to `REFRESH_CHOICES`.** A value outside the
offered set turns auto-refresh off rather than arming a timer, because no
`<option>` would render as selected for it — the control would read "off"
while the page reloaded under the reader.

**3. Colours are tokens.** Custom properties on `:root`, redefined under
`@media (prefers-color-scheme: dark)`. A rule written with a literal colour
will be wrong in one of the two themes.

## Adding a template kind

Two tables, no new branches:

1. `admin/store._KIND_META` — spec subdirectory, accepted extensions, the spec
   key naming the asset, and `clone_drops`: the keys a clone must not inherit
   (omit it when there are none).
2. `admin/kinds.DESCRIPTORS` — label, icon, `has_args`, and the wording for the
   name field, the table's detail column, the post-upload flash and the save
   confirmation.

`KindDescriptor` re-exposes the storage fields as properties, so a view needs
one lookup (`descriptor(kind)`) rather than reaching into both.

Where behaviour genuinely differs, branch on `descriptor(kind).has_args` — the
real distinction, not the kind name. A docx or email template is a
parameterised document: it declares arguments and becomes an MCP tool, so
"live" means a tool is registered. A pptx template is a design with no
arguments; it becomes one more value for the presentation tool's `template`
argument, so "live" means the registry re-read it.

**Disabling is a spec key, not UI state.** `enabled: false` lives in the spec
file and `template_registry.gather_specs()` drops it, which is why a disabled
template stays disabled across a restart with nothing having to remember it.
`AdminContext.sync()` is the one place that decides whether saving a template
registers it or takes it off — a disabled spec must not come back as a live
tool just because it was saved. The edit form carries no `enabled` control, so
the save route reads the stored flag and carries it forward; without that, an
edit would silently switch a disabled template back on.

**A rename goes through `store.rename_spec()`.** It owns the collision check,
the name validation and the write-before-unlink ordering, so an interrupted
rename leaves two templates rather than none. The save route calls it rather
than re-doing the sequence inline: an inline copy is a second implementation
of the same invariant, and it is the one no test reaches. A failed unlink
raises `OSError`, which the route catches alongside `TemplateStoreError` —
otherwise a rename that cannot remove the old file 500s instead of saying so.
The asset keeps its own filename: renaming it would break any master-YAML
entry pointing at the same file, and the spec names it explicitly anyway.

**`original_name` is what separates a create from an edit.** The edit form
renders it, the create form does not, and the save route keys off its
*presence* — not off whether a spec with that name happens to exist. Keyed
the other way, a fresh upload whose name collided with an existing template
read as an ordinary edit: it overwrote the occupant in silence and inherited
its `enabled` flag, so a brand-new template could arrive disabled. A create
that lands on an occupied name is now refused with the same message a rename
gets.

**Adopting is additive: the master YAML is never rewritten.** `gather_specs()`
lets a `.d` entry win over a master entry of the same name, so `AdminContext.
adopt()` only has to write the copy — the hand-written, commented config the
admin owns stays byte-identical, which a test asserts directly. The asset is
copied into `custom_templates/` when it is not already there, because a master
entry usually points at a file shipped in `default_templates/`, and a template
you can edit but whose document you cannot replace is a confusing half-state.

**An adopted template never owns a base-template filename.** `template_utils`
searches `custom_templates/` before `default_templates/`, so copying a file
named `default_docx_template.docx` into the uploads directory shadows the base
Word template — and replacing that one template's document would then restyle
every Word document the server generates. `base_templates.RESERVED_FILENAMES`
names every such file (both the `custom_` and the `default_` spelling of all
five slots); adopting a master entry that points at one gives the template a
private copy under its own name instead.

**`save_spec()` will not adopt a file it only guessed the name of.** The asset
filename comes from the argument, else the spec's path key, else the template's
name — and that last branch is a guess. The guard under it used to read "the
asset must exist", when what a caller in that branch means is "the asset must
be the one I just wrote", so a file left behind by a deleted template was
adopted by the next template of the same name in silence (#184). Deriving onto
an existing file is refused unless bytes are supplied; naming the file
explicitly still reuses it, which is how every edit of an existing template
saves. The two refusals — nothing there, and something unexpected there — say
different things, because they are different problems.

**The pptx analysis reports the content area, read with no overrides.**
`analysis.content_area` comes from the same `templates.content_area_summary()`
the tool's own diagnostics use, so the page and `list_presentation_templates`
cannot disagree about it — a test asserts they match, and another pins the
example `docs/templates.md` prints. Like the role map beside it, it is read
with no layout overrides, and that is exact rather than a simplification:
`LayoutResolver` consults a configured `layouts:` mapping in `provides()` and
`resolve()`, but `_by_role` — which is what `content_area()` reads — is built
from `classify_layout()` alone. Passing a spec's overrides here would pick the
same rectangle, so the card shows what the builder uses. A test pins that,
because if overrides ever start affecting the reading the card has to follow
(#189 proposes exactly that). The warnings cover what an admin can act on — no layout declaring one (the builder then
uses a fixed band that knows nothing about this template), a rectangle
covering essentially the whole slide, and an implausibly small one.

**Preview values travel in a second POST, carrying the spec with them.** The
values form is reached from the edit form, so it has to preview what is on
screen — unsaved edits included — not what is on disk. Rather than re-emit
every field `build_spec()` reads and drift from it, the rebuilt spec rides
along as JSON in `forms.CARRIED_SPEC_FIELD`, and `carried_spec()` ignores
anything that is not a named mapping. The values themselves are prefixed
(`preview.VALUE_PREFIX`) so they cannot collide with the spec fields, and
their absence is what means "generate samples" — which is what keeps the
plain Preview button one click. That absence is not a reliable signal on its
own: an unticked checkbox sends nothing, so a template whose arguments are
all booleans submits no `value_` key when every box is off. The form carries
`preview.VALUES_MARKER` so the submission is recognised whatever its controls
happen to send; without it the admin's explicit off came back as a sample on.

A sampled boolean is subtler than it looks: `sample_values()` uses the
declared default and only falls back to True when there is none. An
undeclared or required flag therefore always previews on, and an *optional*
one always previews off, because `build_spec()` writes `default: false` for a
blank default. Both are pinned by tests, because either way the sample is a
guess and the point of #168 is not to have to guess.

**Adoption records no provenance, so the copy is keyed by name and nothing
else.** Two consequences follow from `gather_specs()` replacing the whole spec
rather than merging fields, and both are pinned by tests so a change to either
is deliberate: editing the master entry after adoption — including disabling
it — has no effect, because the override replaces it entirely; and renaming
the master entry by hand leaves *two* templates, the renamed master entry
(no longer overridden) and the adopted copy (now standing alone). Following a
rename would mean recording where a copy came from, which is a bigger feature
than adopting; the adopt card says plainly that the copy stops following the
entry instead.

The master rows are listed from `_master_specs()` rather than from the live
tool names: keyed off what is registered, a master template that is disabled
or that failed to load vanished from the page entirely, with no way to inspect
it and no way to adopt it.

**A clone copies the asset; it never shares it.** Two specs pointing at one
file would make "Replace document" on either one silently change the other,
and nothing in the UI would report it — the duplication is the cheaper
mistake. `store.clone_spec()` owns that, along with the collision check and
keeping the source's extension (a clone of a `.potx` stays a `.potx`). Keys a
clone must not inherit are declared per kind as `clone_drops` in
`store._KIND_META`, so a new kind states its own rather than a branch
appearing in the route: PowerPoint drops `default`, because two defaults is a
state the registry resolves silently by picking one.

A clone of a disabled template is itself disabled. `AdminContext.sync()`
decides that, the same call that decides it on save — landing a live tool
from a template someone had deliberately taken out of service, with a
description that has not been edited yet, is the wrong default.

**Saving has three outcomes, not two.** Live, deliberately off, and genuinely
failed to register. `sync()` answers "does the server match the spec now?", so
taking a disabled template off succeeds; it used to return `unregister()`'s own
bool, which is `False` whenever there was no live tool to remove — every
disabled save and every clone of a disabled template — and the page then told
the admin to go and read the logs about a registration that was never meant to
happen. `saved_page()` takes the enabled state and picks between `save_ok`,
`save_disabled` and `save_warn` rather than inferring it from one boolean.

**An orphan is defined by what does *not* reference a file, so the reference
map has to be complete.** `assets.reference_map()` counts three things, and
the page offers to delete anything it misses: managed specs, master-YAML
entries (read with `include_disabled=True`, because a disabled template still
owns its file), and the base-template slots — which live in the same flat
directory under fixed names and are named by no spec at all. Deleting one of
those would silently restyle every document the server produces.

Deletion is keyed off the scan rather than the URL: the filename has to match
one `scan()` produced, which is a basename read from the directory, and it
has to still be an orphan when the request arrives. A crafted path matches
nothing, and a file that gained a reference since the page was rendered is no
longer deletable.

## Global styles (#161)

`/styles` edits the Word `style_mapping` that applies to **every** document —
the setting with the widest blast radius on the server, and the last one that
required hand-editing YAML on the volume.

It writes `config/docx_templates.d/_global.yaml`, the merge layer's home for
kind-wide settings, never the master file. The storage rules and why the
override *replaces* rather than merges are in
[dynamic-templates.md](dynamic-templates.md#kind-wide-settings); what the page
itself has to get right:

- **The form opens on what is in force**, not empty. On a server whose master
  YAML has set a mapping all along, an empty form would read as "nothing is
  set", and the first save would silently clear it. Opening pre-filled means
  saving without touching anything is a no-op in effect.
- **A value the base Word template does not define is still offered**, marked
  *not in the base template*. The dropdowns list the base template's styles
  because that is what the static Word tool renders onto — but dropping a
  configured value for being unrecognised would lose configuration by
  rendering a page.
- **A master key the override drops is named on the page** ("Not in force"),
  because the override replaces the master's mapping. Otherwise an admin
  reading `docx_templates.yaml` on the volume sees a setting that is simply
  not happening.
- **Keys the renderer does not act on are flagged**, read from the master as
  well as from the effective mapping, and kept out of the "Not in force" list:
  they were never in force, and listing them beside a real one sends an admin
  looking for a setting they never had.
- **Saving re-registers every Word template tool.** The mapping is baked into
  each tool at registration time, so a save would otherwise move the static
  tool and leave the template tools on the old mapping.
- **A template lost to that re-registration is named, as a warning.**
  `register_docx_template()` removes the existing tool *before* rebuilding it,
  so a template whose source file has gone missing since startup is not merely
  stale — it is off the server, and the registration loop logs it and moves
  on. `resync_docx_style_map()` therefore returns `(live, lost)` from a
  before/after comparison of the live names. From the admin's side the only
  thing that happened was saving a style, so this must not arrive as a success
  message with a smaller number in it. The remove-before-rebuild ordering
  itself is [#192](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/192).

Route ordering matters here as it does for `/base` and `/files`: `/styles/save`
also fits `/{kind}/save`, so the three `/styles` routes are registered before
the generic ones and `tests/test_admin_global_styles.py` pins it.

## Known limitations

- An **accepted** upload is still read fully into memory. Streaming it
  straight to its destination would need the analysers and
  `store.write_asset()` to take a file object rather than bytes — worth doing
  if `ADMIN_MAX_UPLOAD_MB` is ever raised far, not before.

## Tests

| File | Covers |
|------|--------|
| `tests/test_admin_app.py` | routes, auth, CSRF, live registration, the docx/email flow |
| `tests/test_admin_pptx.py` | pptx analysis, the layout form, preview |
| `tests/test_admin_assets.py` | the no-external-assets and `lang` invariants, on every page |
| `tests/test_admin_style_keys.py` | the style-key lists cannot drift from the renderer |
| `tests/test_admin_base_templates.py` | the five base-template slots: state, upload, download, revert |
| `tests/test_metrics_warnings.py` | warnings reach the counters and the Status page |
| `tests/test_admin_log_view.py` | the log view's level, source and search filters |
| `tests/test_admin_template_lifecycle.py` | disabling, enabling and renaming a template |
| `tests/test_admin_source_files.py` | the reference map, and that only an orphan can be deleted |
| `tests/test_admin_clone.py` | what a clone carries, what it drops, and that the asset is copied |
| `tests/test_admin_master_templates.py` | inspecting a master-YAML entry, and adopting it without touching the file |
| `tests/test_admin_preview_values.py` | previewing with your own values: Markdown through a placeholder, conditionals off |
| `tests/test_admin_content_area.py` | the pptx content area: agreeing with the tool, and the cases worth warning about |
| `tests/test_admin_config.py` | `ADMIN_*` settings |
