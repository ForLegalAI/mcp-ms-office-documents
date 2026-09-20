"""Shared view primitives and the self-contained theme for the admin UI.

Every admin page is built from the helpers here rather than from hand-written
FastHTML trees, so "what a card looks like" or "how a form field is laid out"
is decided once. Four properties the module exists to keep:

* **No external assets.** :data:`ADMIN_CSS` and :data:`ADMIN_JS` are inlined
  into every page by :func:`head_tags`. Nothing here may emit a ``<link>`` to
  another origin, a ``<script src>`` or a ``url(https://…)`` — the UI has to
  look right offline and in locked-down deployments, which is also why
  :mod:`admin.app` builds its ``FastHTML`` with ``default_hdrs=False``.
  ``tests/test_admin_assets.py`` enforces this on every rendered page.
* **Every control is labelled.** :func:`field` mints an ``id`` and points the
  ``<label>`` at it. That is the only reason the UI has label/control
  associations at all, so build a field with this rather than by hand.
* **Wide tables scroll, the page does not.** :func:`data_table` wraps its table
  in a scroll container; a six-column table must not force the whole document
  sideways on a phone. The same applies to the two link bars: :func:`nav` and
  :func:`tab_bar` scroll sideways rather than wrapping into a second row that
  pushes the page down.
* **Navigation is links, not script.** :func:`tab_bar` renders anchors to real
  URLs, so a tab is bookmarkable, survives a reload and works with JavaScript
  off. The only script on any page is the argument-row cloner.

The theme is token-based: colours, spacing, radii, type sizes and shadows are
custom properties on ``:root``, with the colours redefined under
``prefers-color-scheme: dark``. A new rule takes a token, never a literal, or
it will be wrong in one of the two themes — or out of step with the scale.

:mod:`admin.views` builds the pages from these; :mod:`admin.sections` decides
what goes in the nav and the tab bars; :mod:`admin.app` only wires them to
routes.
"""
from __future__ import annotations

import re
from typing import Any, Dict, Iterable, Optional, Sequence, Tuple

from fasthtml.common import (
    A, B, Details, Div, Form, H1, H2, H3, Header, Hidden, Input, Label,
    Main, Meta, Nav, NotStr, P, Script, Span, Style, Summary, Table, Tbody,
    Th, Thead, Title, Tr,
)

# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------

ADMIN_CSS = """
:root{
  color-scheme:light;
  --bg:#f6f7fb; --card:#ffffff; --raised:#ffffff;
  --ink:#111827; --ink-2:#374151; --muted:#6b7280;
  --line:#e5e7eb; --line-2:#eef0f4; --row-line:#f3f4f6;
  --brand:#4f46e5; --brand-d:#4338ca; --brand-soft:#eef2ff; --on-brand:#ffffff;
  --link:#4338ca;
  --topbar-bg:#ffffff; --topbar-ink:#111827; --topbar-link:#4b5563;
  --topbar-line:#e5e7eb;
  --ctl-bg:#ffffff; --ctl-disabled-bg:#f9fafb; --btn-bg:#ffffff;
  --accent-bg:#eef2ff; --accent-ink:#4338ca;
  --neutral-bg:#f3f4f6; --hover-bg:#f9fafb;
  --chip-if-bg:#fef9c3; --chip-if-ink:#854d0e;
  --ok:#16a34a; --warn:#b45309; --err:#dc2626; --info:#2563eb;
  --ok-bg:#f0fdf4; --warn-bg:#fffbeb; --err-bg:#fef2f2; --info-bg:#eff6ff;
  --ok-line:#bbf7d0; --warn-line:#fde68a; --err-line:#fecaca; --info-line:#bfdbfe;
  --ok-ink:#15803d; --warn-ink:#92400e; --err-ink:#b91c1c; --info-ink:#1d4ed8;

  --sp-1:.25rem; --sp-2:.5rem; --sp-3:.75rem; --sp-4:1rem; --sp-5:1.5rem;
  --sp-6:2rem; --sp-7:3rem;
  --fs-xs:.75rem; --fs-sm:.8125rem; --fs-md:.875rem; --fs-base:.9375rem;
  --fs-lg:1.0625rem; --fs-xl:1.25rem; --fs-2xl:1.6rem;
  --r-sm:6px; --r-md:10px; --r-lg:14px; --r-full:999px;
  --shadow:0 1px 2px rgba(17,24,39,.04),0 1px 3px rgba(17,24,39,.03);
  --shadow-md:0 2px 4px rgba(17,24,39,.05),0 4px 12px rgba(17,24,39,.05);
  --shadow-bar:0 1px 2px rgba(17,24,39,.04);
  --ring:0 0 0 3px rgba(79,70,229,.25);
  --mono:ui-monospace,SFMono-Regular,Menlo,Consolas,"Liberation Mono",monospace;
}
@media (prefers-color-scheme:dark){
  :root{
    color-scheme:dark;
    --bg:#0b1020; --card:#151b2e; --raised:#1b2236;
    --ink:#e8ecf5; --ink-2:#c7cfe0; --muted:#8e9ab4;
    --line:#2a3350; --line-2:#222b45; --row-line:#1f2740;
    --brand:#6366f1; --brand-d:#818cf8; --brand-soft:#1e2445; --on-brand:#ffffff;
    --link:#a5b4fc;
    --topbar-bg:#111729; --topbar-ink:#e8ecf5; --topbar-link:#9aa6c2;
    --topbar-line:#242c47;
    --ctl-bg:#0f1527; --ctl-disabled-bg:#151b2e; --btn-bg:#1b2236;
    --accent-bg:#1e2445; --accent-ink:#c7d2fe;
    --neutral-bg:#222b45; --hover-bg:#1b2236;
    --chip-if-bg:#3f2d06; --chip-if-ink:#fde68a;
    --ok:#4ade80; --warn:#fbbf24; --err:#f87171; --info:#60a5fa;
    --ok-bg:#0f2a1c; --warn-bg:#2e2408; --err-bg:#331414; --info-bg:#131f39;
    --ok-line:#15803d; --warn-line:#92400e; --err-line:#991b1b; --info-line:#1e40af;
    --ok-ink:#86efac; --warn-ink:#fcd34d; --err-ink:#fca5a5; --info-ink:#bfdbfe;
    --shadow:0 1px 2px rgba(0,0,0,.35);
    --shadow-md:0 2px 6px rgba(0,0,0,.4),0 8px 24px rgba(0,0,0,.28);
    --shadow-bar:0 1px 0 rgba(0,0,0,.3);
    --ring:0 0 0 3px rgba(129,140,248,.3);
  }
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font-family:system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  font-size:var(--fs-base);line-height:1.55;
  -webkit-font-smoothing:antialiased}
a{color:var(--link);text-decoration:none}
a:hover{text-decoration:underline}
:focus-visible{outline:2px solid var(--brand);outline-offset:2px}

/* ---- App bar ---------------------------------------------------------- */
.topbar{background:var(--topbar-bg);color:var(--topbar-ink);
  border-bottom:1px solid var(--topbar-line);box-shadow:var(--shadow-bar);
  position:sticky;top:0;z-index:20}
.topbar-inner{max-width:1180px;margin:0 auto;padding:0 var(--sp-5);
  display:flex;align-items:center;gap:var(--sp-5);min-height:56px}
.topbar .brand{font-weight:650;font-size:var(--fs-base);color:var(--topbar-ink);
  display:flex;align-items:center;gap:var(--sp-2);white-space:nowrap;
  letter-spacing:-.01em}
.topbar .brand:hover{text-decoration:none}
.brand-mark{display:inline-flex;align-items:center;justify-content:center;
  width:26px;height:26px;border-radius:var(--r-sm);
  background:var(--brand);color:var(--on-brand);font-size:.8rem;flex:none}
.topbar nav{display:flex;align-items:center;gap:var(--sp-1);
  overflow-x:auto;scrollbar-width:none;margin-left:auto}
.topbar nav::-webkit-scrollbar{display:none}
.topbar nav a{color:var(--topbar-link);font-size:var(--fs-md);font-weight:550;
  padding:var(--sp-2) var(--sp-3);border-radius:var(--r-sm);white-space:nowrap}
.topbar nav a:hover{color:var(--topbar-ink);background:var(--hover-bg);
  text-decoration:none}
.topbar nav a.active{color:var(--brand-d);background:var(--brand-soft)}
@media (prefers-color-scheme:dark){
  .topbar nav a.active{color:var(--brand-d)}
}
.topbar nav a.nav-end{margin-left:var(--sp-3);color:var(--muted);
  font-weight:500}

/* ---- Page shell ------------------------------------------------------- */
.container{max-width:1180px;margin:0 auto;padding:var(--sp-5) var(--sp-5) var(--sp-7)}
.page-head{margin:0 0 var(--sp-5)}
.page-head .crumb{font-size:var(--fs-sm);color:var(--muted);
  margin:0 0 var(--sp-2)}
.page-head .head-row{display:flex;align-items:flex-start;gap:var(--sp-4);
  flex-wrap:wrap}
.page-head h1{margin:0;flex:1;min-width:14rem}
.page-head .head-actions{display:flex;gap:var(--sp-2);flex-wrap:wrap;
  align-items:center}
.page-head .subtitle{margin:var(--sp-2) 0 0;color:var(--muted);
  font-size:var(--fs-md);max-width:72ch}
h1{font-size:var(--fs-2xl);line-height:1.25;letter-spacing:-.02em;
  margin:.1rem 0 var(--sp-4);font-weight:680}
h2{font-size:var(--fs-lg);margin:0 0 var(--sp-3);font-weight:640;
  letter-spacing:-.01em}
h3{font-size:var(--fs-base);margin:var(--sp-4) 0 var(--sp-2);font-weight:640}
.muted{color:var(--muted);font-size:var(--fs-md)}

/* ---- Tabs ------------------------------------------------------------- */
.tabs{display:flex;gap:var(--sp-1);border-bottom:1px solid var(--line);
  margin:0 0 var(--sp-5);overflow-x:auto;scrollbar-width:none}
.tabs::-webkit-scrollbar{display:none}
.tab{padding:var(--sp-3) var(--sp-3);font-size:var(--fs-md);font-weight:560;
  color:var(--muted);border-bottom:2px solid transparent;white-space:nowrap;
  margin-bottom:-1px}
.tab:hover{color:var(--ink);text-decoration:none}
.tab.active{color:var(--brand-d);border-bottom-color:var(--brand)}
.tab-blurb{color:var(--muted);font-size:var(--fs-md);margin:0 0 var(--sp-4);
  max-width:72ch}

/* ---- Cards ------------------------------------------------------------ */
.card{background:var(--card);border:1px solid var(--line);
  border-radius:var(--r-lg);padding:var(--sp-5);margin-bottom:var(--sp-4);
  box-shadow:var(--shadow)}
.card>*:last-child{margin-bottom:0}
.card-head{display:flex;align-items:center;gap:var(--sp-4);
  flex-wrap:wrap;margin:calc(var(--sp-5)*-1) calc(var(--sp-5)*-1) var(--sp-4);
  padding:var(--sp-3) var(--sp-5);border-bottom:1px solid var(--line-2);
  background:transparent}
.card-head .card-title{margin:0;flex:1;min-width:10rem}
.card-head .card-actions{display:flex;gap:var(--sp-2);flex-wrap:wrap;
  align-items:center}
.card-sub{margin:calc(var(--sp-2)*-1) 0 var(--sp-4);color:var(--muted);
  font-size:var(--fs-md)}

/* ---- Forms ------------------------------------------------------------ */
.field{margin-bottom:var(--sp-4)}
.field label{display:block;font-weight:600;font-size:var(--fs-sm);
  margin-bottom:var(--sp-1);color:var(--ink-2)}
input,select,textarea{width:100%;padding:.5rem .65rem;
  border:1px solid var(--line);border-radius:var(--r-sm);font:inherit;
  font-size:var(--fs-md);background:var(--ctl-bg);color:var(--ink);
  transition:border-color .12s,box-shadow .12s}
input:hover:not([disabled]),select:hover,textarea:hover{border-color:var(--muted)}
input:focus,select:focus,textarea:focus{outline:none;border-color:var(--brand);
  box-shadow:var(--ring)}
input[disabled]{background:var(--ctl-disabled-bg);color:var(--muted);
  cursor:not-allowed}
input[type=file]{padding:.35rem .5rem;font-size:var(--fs-sm)}
textarea{resize:vertical}
.field .muted{margin:var(--sp-1) 0 0;font-size:var(--fs-sm)}

/* ---- Buttons ---------------------------------------------------------- */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:var(--sp-2);
  cursor:pointer;border-radius:var(--r-sm);padding:.5rem .9rem;font:inherit;
  font-size:var(--fs-md);font-weight:600;border:1px solid var(--line);
  background:var(--btn-bg);color:var(--ink);text-decoration:none;
  white-space:nowrap;transition:background .12s,border-color .12s,box-shadow .12s}
.btn:hover{text-decoration:none;background:var(--hover-bg);
  border-color:var(--muted)}
.btn:focus-visible{outline:none;box-shadow:var(--ring)}
.btn-primary{background:var(--brand);color:var(--on-brand);
  border-color:var(--brand);box-shadow:var(--shadow)}
.btn-primary:hover{background:var(--brand-d);border-color:var(--brand-d);
  color:var(--on-brand)}
.btn-secondary{background:var(--btn-bg);color:var(--link);
  border-color:var(--line)}
.btn-secondary:hover{background:var(--accent-bg);border-color:var(--brand)}
.btn-danger{background:var(--btn-bg);color:var(--err);border-color:var(--err-line)}
.btn-danger:hover{background:var(--err-bg);border-color:var(--err)}
.btn-sm{padding:.3rem .6rem;font-size:var(--fs-sm)}
.btn-icon{padding:.25rem .5rem;background:var(--btn-bg);
  border:1px solid var(--line);color:var(--err);border-radius:var(--r-sm)}
.btn-icon:hover{background:var(--err-bg);border-color:var(--err-line)}
.actions{display:flex;gap:var(--sp-2);flex-wrap:wrap;align-items:center}

/* ---- Tables ----------------------------------------------------------- */
.table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch;
  margin:0 calc(var(--sp-5)*-1);padding:0 var(--sp-5)}
.table-actions{margin-top:var(--sp-4)}
table{width:100%;border-collapse:collapse}
th,td{text-align:left;padding:.6rem .6rem;border-bottom:1px solid var(--line-2);
  vertical-align:middle}
th{font-size:var(--fs-xs);text-transform:uppercase;letter-spacing:.04em;
  color:var(--muted);font-weight:650;border-bottom-color:var(--line);
  white-space:nowrap}
tbody tr:last-child td{border-bottom:none}
.tpl-table tbody tr:hover td{background:var(--hover-bg)}
.tpl-table td{font-size:var(--fs-md)}
.tpl-table td:first-child,.tpl-table th:first-child{padding-left:0}
.tpl-table td:last-child,.tpl-table th:last-child{padding-right:0}
.row-name{font-weight:600}
.args-table td{padding:.25rem .3rem;border-bottom:none}
.args-table th{padding-bottom:var(--sp-1)}
.args-table input,.args-table select{padding:.35rem .5rem;font-size:var(--fs-sm)}
.args-table tbody tr:hover td{background:transparent}
.col-name{width:22%}.col-type{width:12%}.col-req{width:13%}.col-def{width:16%}
.col-x{width:38px}

/* ---- Badges, chips ---------------------------------------------------- */
.badge{display:inline-flex;align-items:center;gap:.35em;padding:.15rem .55rem;
  border-radius:var(--r-full);font-size:var(--fs-xs);font-weight:650;
  line-height:1.6;border:1px solid transparent}
.badge::before{content:"";width:.45em;height:.45em;border-radius:var(--r-full);
  background:currentColor;flex:none}
.badge-live{background:var(--ok-bg);color:var(--ok-ink);border-color:var(--ok-line)}
.badge-off{background:var(--neutral-bg);color:var(--muted);border-color:var(--line)}
.badge-ro{background:var(--neutral-bg);color:var(--muted);border-color:var(--line)}
.badge-if{background:var(--accent-bg);color:var(--accent-ink);
  border-color:var(--accent-ink);margin-left:var(--sp-1)}
.badge-plain::before{display:none}
.chip{display:inline-block;padding:.15rem .5rem;
  margin:.15rem .25rem .15rem 0;border-radius:var(--r-sm);
  background:var(--accent-bg);color:var(--accent-ink);font-size:var(--fs-sm);
  font-family:var(--mono)}
.chip-if{background:var(--chip-if-bg);color:var(--chip-if-ink)}

/* ---- Notices ---------------------------------------------------------- */
.flash{padding:.7rem .9rem;border-radius:var(--r-md);margin:var(--sp-3) 0;
  border:1px solid transparent;font-size:var(--fs-md)}
.flash-info{background:var(--info-bg);border-color:var(--info-line);color:var(--info-ink)}
.flash-ok{background:var(--ok-bg);border-color:var(--ok-line);color:var(--ok-ink)}
.flash-warn{background:var(--warn-bg);border-color:var(--warn-line);color:var(--warn-ink)}
.flash-err{background:var(--err-bg);border-color:var(--err-line);color:var(--err-ink)}
.empty{text-align:center;padding:var(--sp-6) var(--sp-4);color:var(--muted)}
.empty .empty-icon{font-size:1.6rem;display:block;margin-bottom:var(--sp-2);
  opacity:.65}
.empty p{margin:0 0 var(--sp-4)}

/* ---- Metrics ---------------------------------------------------------- */
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
  gap:var(--sp-3);margin-bottom:var(--sp-4)}
.stat{background:var(--card);border:1px solid var(--line);
  border-radius:var(--r-md);padding:var(--sp-4);box-shadow:var(--shadow)}
.stat .num{font-size:1.5rem;font-weight:700;line-height:1.2;
  letter-spacing:-.02em}
.stat .lbl{font-size:var(--fs-xs);text-transform:uppercase;letter-spacing:.04em;
  color:var(--muted);font-weight:650;margin-top:var(--sp-1)}

/* ---- Section tiles (the dashboard) ------------------------------------ */
.tile-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));
  gap:var(--sp-4)}
.tile{display:block;background:var(--card);border:1px solid var(--line);
  border-radius:var(--r-lg);padding:var(--sp-5);box-shadow:var(--shadow);
  color:var(--ink);transition:box-shadow .15s,border-color .15s,transform .15s}
.tile:hover{text-decoration:none;box-shadow:var(--shadow-md);
  border-color:var(--brand);transform:translateY(-1px)}
.tile .tile-head{display:flex;align-items:center;gap:var(--sp-2);
  font-weight:650;font-size:var(--fs-lg);letter-spacing:-.01em}
.tile .tile-blurb{color:var(--muted);font-size:var(--fs-sm);
  margin:var(--sp-2) 0 var(--sp-4);min-height:2.6em}
.tile .tile-facts{display:flex;gap:var(--sp-4);flex-wrap:wrap;
  border-top:1px solid var(--line-2);padding-top:var(--sp-3)}
.tile .tile-fact{font-size:var(--fs-sm)}
.tile .tile-fact b{display:block;font-size:var(--fs-lg);font-weight:700;
  line-height:1.3}
.tile .tile-fact span{color:var(--muted);font-size:var(--fs-xs);
  text-transform:uppercase;letter-spacing:.04em}

/* ---- Misc ------------------------------------------------------------- */
.group-title{font-weight:700;font-size:var(--fs-xs);text-transform:uppercase;
  letter-spacing:.04em;color:var(--muted);margin:var(--sp-4) 0 var(--sp-2)}
.yaml-view{background:var(--ctl-bg);border:1px solid var(--line);
  border-radius:var(--r-sm);padding:var(--sp-3) var(--sp-4);overflow-x:auto;
  margin:0;font-family:var(--mono);font-size:var(--fs-sm);line-height:1.5;
  color:var(--ink)}
.yaml-view code{font:inherit;color:inherit;background:none}
code{font-family:var(--mono);font-size:.92em}
.role-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));
  gap:0 var(--sp-4)}
.facts .field{margin-bottom:var(--sp-3)}
.warn-text{color:var(--warn)}
.swatch-wrap{display:inline-flex;align-items:center;gap:var(--sp-1);
  margin:.15rem var(--sp-3) .15rem 0}
.swatch{display:inline-block;width:14px;height:14px;border-radius:var(--r-sm);
  border:1px solid var(--line)}
.swatch-label{font-size:var(--fs-xs);color:var(--muted);font-family:var(--mono)}
.field label input[type=checkbox]{width:auto;margin-right:var(--sp-1);
  vertical-align:-2px}
details{margin:var(--sp-4) 0;border:1px solid var(--line);
  border-radius:var(--r-md);padding:var(--sp-2) var(--sp-4);
  background:var(--raised)}
details[open]{padding-bottom:var(--sp-4)}
summary{cursor:pointer;font-weight:600;font-size:var(--fs-md);
  padding:var(--sp-1) 0}
summary:hover{color:var(--brand-d)}
.login-wrap{max-width:400px;margin:10vh auto}
.inline-form{display:inline}
hr{border:none;border-top:1px solid var(--line);margin:var(--sp-5) 0}
.logs{font-family:var(--mono);font-size:var(--fs-sm);max-height:460px;
  overflow:auto;border:1px solid var(--line);border-radius:var(--r-md)}
.logs table{width:100%}
.logs td{padding:.25rem var(--sp-2);border-bottom:1px solid var(--row-line);
  white-space:nowrap}
.logs tbody tr:hover td{background:var(--hover-bg)}
.logs td.msg{white-space:normal;word-break:break-word}
.logs .ts{color:var(--muted)}
.lvl{font-weight:700}
.lvl-ERROR,.lvl-CRITICAL{color:var(--err)}
.lvl-WARNING{color:var(--warn)}
.lvl-INFO{color:var(--info)}
.lvl-DEBUG{color:var(--muted)}
.toggle-row{display:flex;gap:var(--sp-2);align-items:center;
  margin-bottom:var(--sp-3);flex-wrap:wrap}
.filters{display:flex;gap:var(--sp-3);align-items:flex-end;flex-wrap:wrap;
  margin-bottom:var(--sp-3)}
.filters .field{margin-bottom:0}
.filters .field.grow{flex:1;min-width:180px}
.filters select,.filters input{min-width:140px}
.num-err{color:var(--err)}
.sev{display:inline-block;margin-right:var(--sp-2);font-size:var(--fs-sm);
  font-weight:600}
.tool-list{list-style:none;margin:0;padding:0}
.tool-list li{display:flex;align-items:center;gap:var(--sp-2);
  padding:var(--sp-2) 0;border-bottom:1px solid var(--line-2);
  font-size:var(--fs-md)}
.tool-list li:last-child{border-bottom:none}
.tool-list code{font-weight:600}

@media (max-width:640px){
  .container{padding:var(--sp-4) var(--sp-4) var(--sp-6)}
  .topbar-inner{padding:0 var(--sp-4);gap:var(--sp-3)}
  .card{padding:var(--sp-4)}
  .card-head{margin:calc(var(--sp-4)*-1) calc(var(--sp-4)*-1) var(--sp-3);
    padding:var(--sp-3) var(--sp-4)}
  .table-wrap{margin:0 calc(var(--sp-4)*-1);padding:0 var(--sp-4)}
  h1{font-size:var(--fs-xl)}
}
"""

# Vanilla JS for dynamic argument rows (add / remove) — avoids a CDN htmx dep.
ARG_ROWS_JS = """
function adminAddArgRow(){
  var tb=document.getElementById('argrows');
  if(!tb) return;
  tb.insertAdjacentHTML('beforeend', window.__ARG_ROW_HTML__);
}
function adminRemoveRow(btn){
  var tr=btn.closest('tr'); if(tr) tr.remove();
}
"""

#: Kept as an alias: the page's whole script payload is the argument-row
#: helpers and nothing else, and callers should not have to know that.
ADMIN_JS = ARG_ROWS_JS


def auto_refresh(seconds: int):
    """Reload the page every *seconds*, or nothing when it is off.

    Inline rather than a `<meta http-equiv="refresh">` so it can carry the
    current query string, and inline rather than anything external because
    the UI loads no third-party scripts (see the module docstring).
    """
    if not seconds or seconds <= 0:
        return None
    return Script(NotStr(
        f"setTimeout(function(){{location.reload();}}, {int(seconds) * 1000});"
    ))


def head_tags(blank_row_html: str):
    """The full set of ``<head>`` children: charset, viewport, theme and script.

    *blank_row_html* is a JSON-encoded argument row, injected as
    ``window.__ARG_ROW_HTML__`` so "Add argument" can clone it without a
    client-side templating library.

    Everything is inline by design — see the module docstring.
    """
    return (
        Meta(charset="utf-8"),
        Meta(name="viewport", content="width=device-width, initial-scale=1"),
        Style(ADMIN_CSS),
        Script(NotStr(f"window.__ARG_ROW_HTML__ = {blank_row_html};\n{ADMIN_JS}")),
    )


# ---------------------------------------------------------------------------
# Page shell
# ---------------------------------------------------------------------------

def _link_triples(links):
    """``(label, href, active)`` for items that may omit the active flag."""
    out = []
    for item in links or ():
        if len(item) == 3:
            label, href, active = item
        else:
            (label, href), active = item, False
        out.append((label, href, bool(active)))
    return out


def nav(links, end_links=()):
    """The top-bar navigation.

    *links* and *end_links* are ``(label, href)`` or ``(label, href, active)``.
    Exactly one item should be active; nothing enforces it, because a page
    outside every section (an editor, a confirmation) legitimately has none.

    Scrolls sideways rather than wrapping: a second row of links would push
    the page content down by a variable amount depending on the viewport,
    which is worse than a bar you swipe.
    """
    items = [A(label, href=href, cls="active" if active else None)
             for label, href, active in _link_triples(links)]
    items += [A(label, href=href, cls="nav-end")
              for label, href, _active in _link_triples(end_links)]
    return Nav(*items) if items else Span()


def topbar(brand: str, links=(), end_links=(), home: str = ""):
    """The app bar: brand mark, then the section navigation.

    *brand* is split into its leading glyph and the rest so the glyph can be
    the coloured mark; a brand without one still renders.
    """
    glyph, _, text = str(brand).partition(" ")
    mark = Span(Span(glyph, cls="brand-mark"), text, cls="brand")
    brand_el = A(mark, href=home, cls="brand-link") if home else mark
    return Header(Div(brand_el, nav(links, end_links), cls="topbar-inner"),
                  cls="topbar")


def tab_bar(items):
    """One section's tabs, as ``(label, href, active)`` triples.

    Real anchors to real URLs — a tab is bookmarkable, reloads into itself and
    needs no JavaScript. Returns ``None`` for a section with a single tab:
    a tab bar of one is furniture that says nothing.
    """
    triples = _link_triples(items)
    if len(triples) < 2:
        return None
    return Nav(*[A(label, href=href, cls="tab active" if active else "tab")
                 for label, href, active in triples], cls="tabs")


def page(title: str, header, *content):
    """Wrap page *content* in the themed shell (title + topbar + container)."""
    return (
        Title(f"{title} · Template Admin"),
        header,
        Main(Div(*content, cls="container")),
    )


def page_header(title, subtitle: Optional[str] = None, *actions,
                crumb: Optional[Any] = None):
    """The page title block: optional breadcrumb, an H1, actions and a blurb.

    The actions sit on the title row rather than under the content, so the
    primary thing a page offers is visible without scrolling past a table.
    """
    children = []
    if crumb is not None:
        children.append(Div(crumb, cls="crumb"))
    row = [H1(title)]
    if actions:
        row.append(Div(*actions, cls="head-actions"))
    children.append(Div(*row, cls="head-row"))
    if subtitle:
        children.append(P(subtitle, cls="subtitle"))
    return Div(*children, cls="page-head")


def flash(msg: Optional[str], kind: str = "info"):
    """A coloured notice, or ``None`` when there is no message to show.

    Returning ``None`` for an empty message lets callers splat the result into
    a page unconditionally; FastHTML drops ``None`` children.
    """
    if not msg:
        return None
    return Div(msg, cls=f"flash flash-{kind}")


def card(*content, title: Optional[str] = None, level: int = 3,
         cls: str = "card", actions: Sequence[Any] = (),
         subtitle: Optional[str] = None):
    """A titled panel. *level* picks ``H2`` (section) or ``H3`` (sub-section).

    *actions* puts controls on the title row, in a header strip divided from
    the body — the place a card's own buttons go, so they are not mistaken for
    part of the content below them. A card with actions must have a title;
    there would be nothing for them to sit beside otherwise.
    """
    heading = {1: H1, 2: H2, 3: H3}[level]
    children = []
    if title and actions:
        children.append(Div(heading(title, cls="card-title"),
                            Div(*actions, cls="card-actions"), cls="card-head"))
    elif title:
        children.append(heading(title))
    if subtitle:
        children.append(P(subtitle, cls="card-sub"))
    children += list(content)
    return Div(*children, cls=cls)


def action_bar(*controls):
    """The row of buttons that closes a form."""
    return Div(*controls, cls="actions")


# ---------------------------------------------------------------------------
# Form fields
# ---------------------------------------------------------------------------

_SLUG_RE = re.compile(r"[^a-z0-9]+")


def _slug(text: str) -> str:
    return _SLUG_RE.sub("-", str(text).strip().lower()).strip("-") or "field"


def control_id(control, label: str) -> str:
    """The ``id`` a control should carry, derived from its ``name``.

    Falls back to the label text so a control without a ``name`` (a disabled
    display-only input, say) is still associated with its label rather than
    silently losing the association.
    """
    existing = control.attrs.get("id")
    if existing:
        return str(existing)
    name = control.attrs.get("name")
    return f"f-{_slug(name)}" if name else f"f-{_slug(label)}"


def field(label: str, control, hint: Optional[str] = None, cls: str = "field",
          hint_cls: str = "muted"):
    """A labelled form control, with the label pointing at the control.

    This is the only place that associates a ``<label>`` with its input, so
    every field in the UI goes through it. The ``id`` is derived from the
    control's ``name`` (see :func:`control_id`); pass an explicit ``id=`` on the
    control when two fields on one page would otherwise collide.

    The ``id`` is set on *control* in place, so pass a freshly built element.
    Handing the same instance to two ``field()`` calls silently rewrites the
    first one's ``id`` and leaves its label pointing at nothing.

    *hint_cls* lets a hint be a warning ("no layout matches this role") rather
    than the usual muted aside.
    """
    cid = control_id(control, label)
    control.attrs["id"] = cid
    children = [Label(label, **{"for": cid}), control]
    if hint:
        children.append(P(hint, cls=hint_cls))
    return Div(*children, cls=cls)


def static_row(label: str, value, cls: str = "field"):
    """A read-only labelled row.

    Not a form field: there is no control for a label to point at, so this
    renders a bare ``<label>`` on purpose and must not be used for input.
    """
    return Div(Label(label), value, cls=cls)


def checkbox_field(name: str, label: str, checked: bool = False,
                   hint: Optional[str] = None, value: str = "1"):
    """A checkbox with its label wrapped around it, plus an optional hint.

    Wrapping is the association here — the label is the clickable text beside
    the box, so it needs no ``for``.
    """
    box = Input(name=name, type="checkbox", value=value, checked=checked,
                id=f"f-{_slug(name)}")
    children = [Label(box, f" {label}")]
    if hint:
        children.append(P(hint, cls="muted"))
    return Div(*children, cls="field")


def hidden(name: str, value: str):
    return Hidden(name=name, value=value if value is not None else "")


def csrf_input(token: str):
    return Hidden(name="csrf", value=token or "")


# ---------------------------------------------------------------------------
# Data display
# ---------------------------------------------------------------------------

def data_table(headers: Sequence[Any], rows: Sequence[Any],
               cls: str = "tpl-table", body_id: Optional[str] = None):
    """A table wrapped in a horizontal scroll container.

    The wrapper is the point: a wide table (the status page's six columns, the
    layout report's placeholder lists) scrolls inside its card instead of
    forcing the whole page sideways on a narrow screen.

    *headers* may be plain strings or ready-made ``Th`` cells when a column
    needs a width class. *body_id* puts an id on the ``<tbody>``, which the
    argument editor needs so its "Add argument" script can append to it.
    """
    head = [h if getattr(h, "tag", None) == "th" else Th(h) for h in headers]
    body = Tbody(*rows, id=body_id) if body_id else Tbody(*rows)
    return Div(Table(Thead(Tr(*head)), body, cls=cls), cls="table-wrap")


def badge(text: str, variant: str = "off", plain: bool = False, **kwargs):
    """A pill label. The leading dot means *state*; *plain* drops it.

    Keeping the dot for state alone is what makes "Live" and "Disabled" read
    as a condition of the thing, while a classification beside them ("From a
    template", "master YAML") reads as a label. With a dot on everything the
    distinction is lost and the row is just busier.
    """
    cls = f"badge badge-{variant}" + (" badge-plain" if plain else "")
    return Span(text, cls=cls, **kwargs)


def status_badge(live: bool, enabled: bool = True):
    """Live / Disabled / Not live.

    Disabled is its own state rather than another "Not live": one is a choice
    the admin made and can undo from the same row, the other is a template that
    failed to register and wants investigating (#165).
    """
    if not enabled:
        return badge("Disabled", "off", title="Turned off — the AI cannot call it")
    return badge("Live", "live") if live else badge("Not live", "off")


def chip(text: str, conditional: bool = False):
    """One monospace chip — a placeholder, a role, a literal bit of syntax."""
    return Span(text, cls="chip chip-if" if conditional else "chip")


def chips(items: Iterable[str], highlight: Optional[set] = None,
          empty: str = "none"):
    """Monospace chips for a list of names; *highlight* marks conditionals.

    Wrapped in a ``<span>`` so the group is one node; use :func:`chip` where a
    single chip goes straight into a cell and the wrapper would be noise.
    """
    items = list(items or [])
    if not items:
        return Span(empty, cls="muted")
    highlight = highlight or set()
    return Span(*[chip(it, it in highlight) for it in items])


def swatch(name: str, value: str):
    """One theme colour, shown rather than described."""
    return Span(
        Span(cls="swatch", style=f"background:{value}"),
        Span(f"{name} {value}", cls="swatch-label"),
        cls="swatch-wrap", title=f"{name}: {value}",
    )


def stat(label: str, value: Any, num_cls: str = ""):
    return Div(Div(str(value), cls=f"num {num_cls}".strip()),
               Div(label, cls="lbl"), cls="stat")


def stats_row(*tiles):
    return Div(*tiles, cls="stats")


def tile(href: str, title: str, blurb: str = "", facts: Sequence[Tuple[Any, str]] = ()):
    """A linked summary card for the dashboard.

    *facts* are ``(value, label)`` pairs — the two or three numbers that say
    whether this area needs attention, so the dashboard answers "is anything
    off?" rather than only "where do I click?".
    """
    children = [Div(title, cls="tile-head")]
    if blurb:
        children.append(Div(blurb, cls="tile-blurb"))
    if facts:
        children.append(Div(*[
            Div(B(str(value)), Span(label), cls="tile-fact")
            for value, label in facts
        ], cls="tile-facts"))
    return A(*children, href=href, cls="tile")


def tile_grid(*tiles):
    return Div(*tiles, cls="tile-grid")


def details_block(summary: str, *content):
    return Details(Summary(summary), *content)


def empty_state(message: str, *actions, icon: str = "📭"):
    """The "nothing here yet" block, with the action that fills it."""
    children = [Span(icon, cls="empty-icon"), P(message)]
    if actions:
        children.append(Div(*actions, cls="actions",
                            style="justify-content:center"))
    return Div(*children, cls="empty")


def post_form(action: str, *content, csrf: str = "", cls: str = "",
              enctype: Optional[str] = None):
    """A POST form that always carries the session CSRF token."""
    kwargs: Dict[str, Any] = {"action": action, "method": "post"}
    if cls:
        kwargs["cls"] = cls
    if enctype:
        kwargs["enctype"] = enctype
    return Form(csrf_input(csrf), *content, **kwargs)


__all__ = [
    "ADMIN_CSS", "ADMIN_JS", "ARG_ROWS_JS", "auto_refresh", "head_tags",
    "nav", "topbar", "tab_bar", "page", "page_header", "flash", "card",
    "action_bar", "field", "static_row", "checkbox_field", "hidden", "csrf_input",
    "data_table", "badge", "status_badge", "chip", "chips", "swatch",
    "stat", "stats_row", "tile", "tile_grid", "details_block", "empty_state",
    "post_form", "control_id",
]
