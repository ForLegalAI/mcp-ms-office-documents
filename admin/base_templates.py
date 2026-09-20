"""The static base templates: the files that style *every* generated document.

Separate from :mod:`admin.kinds`, which describes the *dynamic* templates — the
ones that become tools, or values for the presentation tool's ``template``
argument. These five are different in kind:

* they are referenced by a **fixed filename**, not by a spec, so there is no
  YAML entry and nothing to name or configure;
* there is **exactly one of each**, so they are replaced rather than created
  and deleted; and
* they apply to every document the server produces, which makes them the
  highest-blast-radius thing an admin can change here.

Resolution follows :mod:`template_utils`: ``custom_templates/<custom_name>``
wins over ``default_templates/<default_name>``. "Reverting" is therefore just
deleting the custom file and letting the bundled default surface again — except
for Excel, which ships no default at all (see :attr:`BaseSlot.default_name`).

Nothing here takes a filename from a caller. Every path is built from this
table, so a slot key is the only untrusted input and it is looked up, never
joined.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional, Tuple

# Kept in step with docs/templates.md's "Static Templates" table.


@dataclass(frozen=True)
class BaseSlot:
    """One replaceable base template."""

    key: str
    label: str
    icon: str
    custom_name: str
    #: The bundled fallback, or ``None`` where nothing ships (Excel).
    default_name: Optional[str]
    exts: Tuple[str, ...]
    #: One line on what changing this file actually affects.
    controls: str
    #: Which :func:`admin.analysis.analyze` branch reads it.
    analysis_kind: str
    #: What its analysis is reporting, for the report heading.
    report_title: str

    @property
    def accept(self) -> str:
        return ",".join(self.exts)

    @property
    def has_default(self) -> bool:
        return self.default_name is not None


SLOTS: Tuple[BaseSlot, ...] = (
    BaseSlot(
        key="docx", label="Word", icon="📝",
        custom_name="custom_docx_template.docx",
        default_name="default_docx_template.docx",
        exts=(".docx",),
        controls="The look of every Word document the server generates — its "
                 "styles, fonts, margins and page setup.",
        analysis_kind="docx",
        report_title="Styles this template defines",
    ),
    BaseSlot(
        key="email", label="Email wrapper", icon="✉️",
        custom_name="custom_email_template.html",
        default_name="default_email_template.html",
        exts=(".html",),
        controls="The HTML wrapper around every generated email.",
        analysis_kind="email",
        report_title="What we found in the wrapper",
    ),
    BaseSlot(
        key="pptx_16_9", label="PowerPoint 16:9", icon="📊",
        custom_name="custom_pptx_template_16_9.pptx",
        default_name="default_pptx_template_16_9.pptx",
        exts=(".pptx",),
        controls="The default widescreen deck design, used when a call names "
                 "no template.",
        analysis_kind="pptx",
        report_title="What we found in the template",
    ),
    BaseSlot(
        key="pptx_4_3", label="PowerPoint 4:3", icon="📊",
        custom_name="custom_pptx_template_4_3.pptx",
        default_name="default_pptx_template_4_3.pptx",
        exts=(".pptx",),
        controls="The default 4:3 deck design.",
        analysis_kind="pptx",
        report_title="What we found in the template",
    ),
    BaseSlot(
        key="xlsx", label="Excel named styles", icon="📈",
        custom_name="custom_xlsx_template.xlsx",
        # Nothing ships: with no custom workbook there are simply no named
        # styles, rather than a fallback set. Reverting here removes the
        # feature instead of restoring a default.
        default_name=None,
        exts=(".xlsx",),
        controls="The named cell styles a 'styles:' directive can reference as "
                 "style:<Name>. Without one, no named styles exist.",
        analysis_kind="xlsx",
        report_title="Named styles this workbook offers",
    ),
)

_BY_KEY: Dict[str, BaseSlot] = {s.key: s for s in SLOTS}

#: Filenames that mean "base template" to :mod:`template_utils`, whichever
#: directory they sit in. A dynamic template must never point at one of these
#: inside ``custom_templates/``: resolution searches the custom directory
#: first, so the file would shadow the base template for its kind, and
#: replacing that one template's document would restyle every document the
#: server generates (#167).
RESERVED_FILENAMES: frozenset = frozenset(
    [s.custom_name for s in SLOTS]
    + [s.default_name for s in SLOTS if s.default_name]
)


def slot(key: str) -> BaseSlot:
    return _BY_KEY[key]


def is_slot(key: str) -> bool:
    return key in _BY_KEY


def custom_path(custom_dir: Path, s: BaseSlot) -> Path:
    """Where a replacement for *s* is written. Never caller-supplied."""
    return Path(custom_dir) / s.custom_name


def has_custom(custom_dir: Path, s: BaseSlot) -> bool:
    return custom_path(custom_dir, s).is_file()


def active_path(s: BaseSlot) -> Optional[Path]:
    """The file the server would actually use, resolved as the tools resolve it.

    Goes through :mod:`template_utils` rather than reimplementing the search
    order, so this page cannot disagree with what a generated document uses.
    """
    from template_utils import find_file_in_template_dirs

    for name in filter(None, (s.custom_name, s.default_name)):
        found = find_file_in_template_dirs(name)
        if found is not None:
            return found
    return None


def source_of(custom_dir: Path, s: BaseSlot, active: Optional[Path]) -> str:
    """``"custom"``, ``"default"`` or ``"none"`` for a slot's resolved file.

    Compares *active* against the exact path a replacement is written to,
    rather than looking for ``custom_templates`` in the path's parts the way
    :func:`template_utils._classify_template_source` does. That heuristic is
    right in production and wrong anywhere the directory is named something
    else — including every test — and this decides whether the UI offers to
    revert, so it should not depend on what a directory is called.
    """
    if active is None:
        return "none"
    try:
        if active.resolve() == custom_path(custom_dir, s).resolve():
            return "custom"
    except OSError:  # pragma: no cover - unreadable path
        pass
    return "default"
