"""Inject cached formula values into an XLSX produced by openpyxl.

openpyxl writes formula cells as::

    <c r="A4"><f>SUM(A1:A3)</f><v></v></c>

The ``<v>`` element is empty because openpyxl does not evaluate
formulas. Many downstream consumers (Google Sheets preview, mail-client
preview, openpyxl loaded with ``data_only=True``) display empty cells
in this state — the formula only recomputes once a real spreadsheet
application opens the file.

This module post-processes the XLSX bytes to replace the empty ``<v>``
with the value computed by :mod:`xlsx_tools.formula_engine`. It walks
the zip archive, parses each worksheet XML with stdlib ElementTree,
and rewrites only the cells that contain a formula and have a computed
value available. All other content (styles, fonts, number formats,
freeze panes, table definitions, shared strings) is preserved
byte-for-byte because we only touch the worksheet XML files.

No third-party XML library is required — stdlib ``xml.etree.ElementTree``
is sufficient and keeps the Docker image lean.
"""

from __future__ import annotations

import io
import logging
import re
import zipfile
from datetime import date, datetime
from typing import Any
from xml.etree import ElementTree as ET

logger = logging.getLogger(__name__)


# XML namespaces used in OOXML worksheet files.
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Register prefixes so ElementTree writes back the same ``<c>`` / ``<v>``
# tags without inventing ``ns0:`` prefixes.
ET.register_namespace("", _NS_MAIN)
ET.register_namespace("r", _NS_R)

_TAG_C = f"{{{_NS_MAIN}}}c"
_TAG_V = f"{{{_NS_MAIN}}}v"
_TAG_F = f"{{{_NS_MAIN}}}f"
_TAG_ROW = f"{{{_NS_MAIN}}}row"
_TAG_WORKSHEET = f"{{{_NS_MAIN}}}worksheet"
_TAG_SHEET = f"{{{_NS_MAIN}}}sheet"
_TAG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"

_ATTR_R = f"{{{_NS_R}}}id"
_ATTR_TYPE = "t"
_ATTR_COORD = "r"
_ATTR_TARGET = "Target"
_ATTR_ID = "Id"
_ATTR_NAME = "name"

# Excel serial date epoch (1900 system, with the well-known 1900-02-29
# compatibility quirk baked in by adding 1 to all serials after Feb 28 1900;
# openpyxl uses the same convention).
_EXCEL_EPOCH = datetime(1899, 12, 30)


# ── Sheet-name → worksheet-file mapping ──────────────────────────────────────


def _build_sheet_file_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """Map sheet name -> archive path of its worksheet XML.

    Reads ``xl/workbook.xml`` for ``<sheet name=... r:id=...>`` and
    ``xl/_rels/workbook.xml.rels`` for ``<Relationship Id=... Target=...>``
    to translate sheet names to the actual ``xl/worksheets/sheetN.xml``
    path inside the archive.
    """
    rels_by_id: dict[str, str] = {}
    try:
        rels_xml = zf.read("xl/_rels/workbook.xml.rels")
        rels_root = ET.fromstring(rels_xml)
        for rel in rels_root:
            rid = rel.get(_ATTR_ID)
            target = rel.get(_ATTR_TARGET)
            if rid and target and "worksheet" in (rel.get("Type") or ""):
                rels_by_id[rid] = _normalise_worksheet_path(target)
    except KeyError:
        # No rels file (shouldn't happen for a valid xlsx but be defensive).
        pass

    sheet_map: dict[str, str] = {}
    try:
        wb_xml = zf.read("xl/workbook.xml")
        wb_root = ET.fromstring(wb_xml)
        for sheet in wb_root.iter(_TAG_SHEET):
            name = sheet.get(_ATTR_NAME)
            rid = sheet.get(_ATTR_R)
            if name and rid and rid in rels_by_id:
                sheet_map[name] = rels_by_id[rid]
    except KeyError:
        pass

    return sheet_map


def _normalise_worksheet_path(target: str) -> str:
    """Normalise a relationship Target to an absolute archive path.

    openpyxl sometimes writes absolute (``/xl/worksheets/sheet1.xml``)
    and sometimes relative (``worksheets/sheet1.xml``) targets. Resolve
    both to the canonical archive path.
    """
    if target.startswith("/"):
        return target.lstrip("/")
    # Relative: resolve against the workbook.xml location (xl/).
    return f"xl/{target}"


# ── Value serialisation ──────────────────────────────────────────────────────


def _format_value_for_xml(value: Any) -> tuple[str, str | None]:
    """Serialise a Python scalar to OOXML ``<v>`` text and an optional ``t`` attr.

    Returns ``(value_text, cell_type_attr)`` where ``cell_type_attr`` is
    ``None`` when the default (numeric) type applies, or one of ``"b"``
    (boolean), ``"str"`` (string result of a formula), ``"e"`` (error).

    Datetimes are converted to Excel serial numbers (numeric type).
    """
    # Boolean — must come before int check because bool is a subclass of int.
    if isinstance(value, bool):
        return ("1" if value else "0", "b")

    # Integer / float.
    if isinstance(value, int):
        return (str(value), None)
    if isinstance(value, float):
        return (_format_float(value), None)

    # Date / datetime → Excel serial number (numeric type).
    if isinstance(value, datetime):
        serial = _datetime_to_serial(value)
        return (_format_float(serial), None)
    if isinstance(value, date):
        serial = _datetime_to_serial(datetime(value.year, value.month, value.day))
        return (_format_float(serial), None)

    # Fallback: treat as string. (Caller normally filters strings out
    # upstream, but handle defensively.)
    if isinstance(value, str):
        return (value, "str")

    # Unknown type — best effort.
    return (str(value), "str")


def _format_float(value: float) -> str:
    """Format a float for XML without trailing-zero noise.

    ``repr(float)`` gives the shortest round-trippable representation
    in Python 3, which is what we want. Integers stored as floats
    (e.g. ``600.0``) become ``"600"`` for cleaner output.
    """
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    return repr(value)


def _datetime_to_serial(dt: datetime) -> float:
    """Convert a datetime to an Excel serial number (1900 system)."""
    delta = dt - _EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400.0


# ── Public API ───────────────────────────────────────────────────────────────


def count_formulas(xlsx_bytes: bytes) -> int:
    """Count the total number of formula cells in an XLSX byte string.

    A formula cell is any ``<c>`` element containing an ``<f>`` child.
    Used for telemetry and to populate ``RecalcResult.total_formulas``.
    Returns 0 if the archive can't be read (best-effort).
    """
    try:
        zf = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
        try:
            count = 0
            for info in zf.infolist():
                # Worksheet XML files live under xl/worksheets/.
                if not info.filename.startswith("xl/worksheets/sheet"):
                    continue
                data = zf.read(info.filename)
                root = ET.fromstring(data)
                # A formula cell has an <f> child. Count cells where
                # find('f') succeeds.
                for cell in root.iter(_TAG_C):
                    if cell.find(_TAG_F) is not None:
                        count += 1
            return count
        finally:
            zf.close()
    except Exception as e:
        logger.debug("Formula count failed: %s", e)
        return 0


def inject_cached_values(
    xlsx_bytes: bytes,
    values_map: dict[str, Any],
) -> bytes:
    """Inject computed formula values into an XLSX byte string.

    Args:
        xlsx_bytes: Raw xlsx bytes (e.g. from ``openpyxl.Workbook.save``).
        values_map: Mapping of ``"Sheet!Cell"`` (sheet name quoted with
            single quotes if it contains spaces) -> computed Python
            scalar. Cells not present in this map are left untouched.

    Returns:
        New xlsx bytes with cached ``<v>`` values written into matching
        formula cells. The original bytes are not modified.

    The function is tolerant: any internal error is logged and the
    original bytes are returned unchanged. Injection is best-effort —
    failing to inject one cell never prevents the file from shipping.
    """
    if not values_map:
        return xlsx_bytes

    try:
        return _inject_impl(xlsx_bytes, values_map)
    except Exception as e:
        logger.warning(
            "Cached-value injection skipped due to error: %s", e, exc_info=True
        )
        return xlsx_bytes


def _inject_impl(
    xlsx_bytes: bytes,
    values_map: dict[str, Any],
) -> bytes:
    """Implementation of :func:`inject_cached_values` — may raise."""

    # Read the entire archive into memory (xlsx files from this server
    # are small, generated workbooks).
    source_zip = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    try:
        sheet_file_map = _build_sheet_file_map(source_zip)

        # Build a reverse index: archive path -> {cell_coord: value}.
        # The values_map is keyed by "Sheet!Cell"; we need to translate
        # sheet name to worksheet archive path.
        per_file_targets: dict[str, dict[str, Any]] = {}
        for location_key, value in values_map.items():
            sheet_name, coordinate = _parse_location(location_key)
            archive_path = sheet_file_map.get(sheet_name)
            if not archive_path:
                # Sheet name not found in the workbook — skip silently.
                # This can happen for case-casing mismatches or when the
                # engine returned a sheet that no longer exists.
                logger.debug(
                    "No worksheet file for sheet '%s' (location %s) — skipping",
                    sheet_name, location_key,
                )
                continue
            per_file_targets.setdefault(archive_path, {})[coordinate] = value

        if not per_file_targets:
            logger.debug("No injectable targets matched any worksheet — returning original bytes")
            return xlsx_bytes

        # Copy all entries, replacing worksheet XML files where we have
        # injected content.
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for item in source_zip.infolist():
                data = source_zip.read(item.filename)
                if item.filename in per_file_targets:
                    try:
                        data = _inject_into_sheet_xml(
                            data, per_file_targets[item.filename]
                        )
                    except Exception as e:
                        logger.warning(
                            "Failed to inject into %s: %s — leaving unchanged",
                            item.filename, e,
                        )
                out_zip.writestr(item, data)
        return out_buf.getvalue()
    finally:
        source_zip.close()


def _parse_location(location_key: str) -> tuple[str, str]:
    """Parse ``"Sheet!Cell"`` or ``"'Sheet Name'!Cell"`` into (sheet, cell)."""
    # Strip optional surrounding single quotes from sheet name.
    if location_key.startswith("'"):
        # Find the closing quote followed by '!'.
        close = location_key.find("'", 1)
        if close == -1 or location_key[close + 1:close + 2] != "!":
            # Malformed; best effort.
            parts = location_key.split("!", 1)
            return (parts[0].strip("'"), parts[1] if len(parts) > 1 else "")
        sheet = location_key[1:close]
        coordinate = location_key[close + 2:]
        return (sheet, coordinate)
    parts = location_key.split("!", 1)
    if len(parts) == 1:
        return (location_key, "")
    return (parts[0], parts[1])


def _inject_into_sheet_xml(
    sheet_xml: bytes,
    targets: dict[str, Any],
) -> bytes:
    """Inject values into a single worksheet XML document.

    ``targets`` maps cell coordinate (e.g. ``"B5"``) to the computed
    Python scalar value. Only ``<c>`` elements that already contain an
    ``<f>`` child (i.e. formula cells) are modified; this prevents us
    from accidentally overwriting literal values.
    """
    root = ET.fromstring(sheet_xml)
    injected = 0

    for cell in root.iter(_TAG_C):
        coord = cell.get(_ATTR_COORD)
        if not coord or coord not in targets:
            continue

        # Only modify cells that contain a formula.
        formula_elem = cell.find(_TAG_F)
        if formula_elem is None:
            continue

        value = targets[coord]
        value_text, type_attr = _format_value_for_xml(value)

        # Replace existing <v> child if present, else append a new one.
        v_elem = cell.find(_TAG_V)
        if v_elem is None:
            v_elem = ET.SubElement(cell, _TAG_V)
        v_elem.text = value_text

        # Update cell type attribute if needed (e.g. for booleans).
        if type_attr is not None:
            cell.set(_ATTR_TYPE, type_attr)
        else:
            # If the cell had a stale type attr (e.g. openpyxl wrote
            # t="str" by default for formula cells), clear it so Excel
            # treats the cached value as numeric.
            if _ATTR_TYPE in cell.attrib:
                del cell.attrib[_ATTR_TYPE]

        injected += 1

    if injected == 0:
        return sheet_xml

    logger.debug("Injected %d cached values into worksheet", injected)
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)
