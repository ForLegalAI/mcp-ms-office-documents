"""What is in ``custom_templates/``, and what still points at it (#166).

The directory is write-only from the admin's point of view: files go in
through uploads and are never shown again. Deleting a template keeps its
source file by design, a rename or a replace can strand one, and files copied
onto the volume by hand are equally unlisted — so orphans accumulate where
nothing reports them.

This module answers "what references this file?" for every file in the
directory. Getting that answer wrong in the *unreferenced* direction is what
would make the feature dangerous — the view offers to delete an orphan, so a
file wrongly called an orphan is a file wrongly offered for deletion. Three
things count as a reference, and all three are easy to forget:

* a **managed spec** in ``config/<kind>_templates.d/``;
* a **master-YAML template**, hand-written and never rewritten by tooling —
  read through the same merge the registry uses, including disabled ones,
  because a disabled template still owns its file; and
* a **base-template slot** (``custom_docx_template.docx`` and its four
  siblings), which lives in the same flat directory under a fixed name and is
  referenced by nothing at all in the spec layer.

:mod:`admin.app` uses :func:`scan` for the maintenance view.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Tuple

from admin import base_templates
from admin.kinds import KINDS, descriptor

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class AssetFile:
    """One file in ``custom_templates/`` and what still points at it."""

    name: str
    size: int
    mtime: float
    #: Human-readable descriptions of what references this file, e.g.
    #: ``"Word template 'formal_letter'"``. Empty means an orphan.
    references: Tuple[str, ...] = ()

    @property
    def orphaned(self) -> bool:
        return not self.references


def reference_map(store, master_specs: Callable[[str], Iterable[dict]]
                  ) -> Dict[str, List[str]]:
    """Map each referenced filename to what references it.

    *master_specs* takes a kind and returns that kind's master-YAML entries as
    written — :meth:`admin.app.AdminContext._master_specs`, which passes
    ``include_disabled=True``. A disabled template is still using its file, so
    leaving those out would report their assets as orphans and offer them for
    deletion.
    """
    refs: Dict[str, List[str]] = {}

    def add(filename: Optional[str], description: str) -> None:
        if isinstance(filename, str) and filename:
            refs.setdefault(filename, []).append(description)

    for kind in KINDS:
        d = descriptor(kind)
        for spec in store.list_specs(kind):
            add(spec.get(d.path_key), f"{d.label} template '{spec.get('name')}'")
        for spec in master_specs(kind):
            add(spec.get(d.path_key),
                f"{d.label} template '{spec.get('name')}' (master YAML)")

    # The base templates are not specs and no spec names them, but they are
    # the files that style every document the server produces. Offering one
    # for deletion as an orphan would quietly change the look of everything.
    for slot in base_templates.SLOTS:
        add(slot.custom_name, f"Base template — {slot.label}")

    return refs


def scan(custom_dir: Path, refs: Dict[str, List[str]]) -> List[AssetFile]:
    """Every file in *custom_dir*, orphans first, then by name.

    Orphans lead because they are the only rows with an action on them, and
    the point of the page is to find them. Subdirectories are skipped: the
    store writes a flat directory and resolves assets by bare filename, so
    anything nested is not reachable as a template asset anyway.
    """
    directory = Path(custom_dir)
    if not directory.is_dir():
        return []

    files: List[AssetFile] = []
    for path in sorted(directory.iterdir()):
        if not path.is_file():
            continue
        try:
            stat = path.stat()
        except OSError:  # pragma: no cover - vanished between listing and stat
            logger.warning("[admin-assets] Could not stat %s", path)
            continue
        files.append(AssetFile(
            name=path.name,
            size=stat.st_size,
            mtime=stat.st_mtime,
            references=tuple(refs.get(path.name, ())),
        ))

    files.sort(key=lambda f: (not f.orphaned, f.name.lower()))
    return files


def orphan_names(files: Iterable[AssetFile]) -> List[str]:
    """The filenames safe to offer for deletion."""
    return [f.name for f in files if f.orphaned]
