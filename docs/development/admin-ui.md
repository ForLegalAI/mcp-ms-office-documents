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
| `admin/views/` | the pages — `shell`, `templates`, `base`, `status`, `login` |
| `admin/store.py` | persistence: `config/<kind>_templates.d/<name>.yaml` + the asset |
| `admin/analysis.py` | what is inside an uploaded `.docx` / `.html` / `.pptx` |
| `admin/preview.py` | rendering a template without touching the upload backend |
| `admin/auth.py` | the shared-password gate and CSRF tokens |

`components`, `kinds`, `forms`, `store` and `analysis` do not import
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

**3. Colours are tokens.** Custom properties on `:root`, redefined under
`@media (prefers-color-scheme: dark)`. A rule written with a literal colour
will be wrong in one of the two themes.

## Adding a template kind

Two tables, no new branches:

1. `admin/store._KIND_META` — spec subdirectory, accepted extensions, the spec
   key naming the asset.
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

## Known limitations

- The **global** `style_mapping` is still read-only: the editor now says what
  each key inherits from it, but there is no way to change it from the UI
  ([#161](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/161)).
  Editing it needs a home for global config in the `*.d` merge layer —
  `gather_specs()` reads top-level keys from the master YAML only — and
  `style_map.load_global_style_map()` caches its result for the process, so
  both would have to change together for an edit to take effect without a
  restart.
- A kept source file is still invisible: deleting now offers to remove it, but
  nothing lists the files in `custom_templates/` that no template points at
  ([#166](https://github.com/ForLegalAI/mcp-ms-office-documents/issues/166)).
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
| `tests/test_admin_config.py` | `ADMIN_*` settings |
