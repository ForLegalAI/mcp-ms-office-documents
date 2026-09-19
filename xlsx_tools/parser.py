"""Shared markdown line-walking logic for xlsx tools.

Provides a single-pass generator that yields structured events, used by both
the position-scanning pass and the workbook-building pass to ensure they stay
in sync.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from . import warnings as W
from .helpers import TABLE_BOTTOM_SPACING, parse_table


# Pattern for multi-sheet heading: ## Sheet: Name
SHEET_HEADING_PATTERN = re.compile(r'^##\s+Sheet:\s+(.+)$')
# Pattern for comment directives: <!-- key: value --> or <!-- key -->
DIRECTIVE_PATTERN = re.compile(r'^<!--\s*(\w[\w-]*)(?:\s*:\s*(.*?))?\s*-->$')

# Spacing inserted after a header row (rows)
HEADER_ROW_SPACING = 2
# Spacing inserted after a table (rows) is TABLE_BOTTOM_SPACING, imported from
# helpers so the writer and the row bookkeeping here can never disagree.
# Maximum allowed Excel sheet name length
MAX_SHEET_NAME_LENGTH = 31

DEFAULT_SHEET_NAME = "Data Report"


def _sanitize_sheet_name(name: str) -> str:
    """Sanitize a sheet name for Excel compatibility.

    Excel sheet names must be ≤31 chars and cannot contain []*?:/\\ characters.
    """
    sanitized = re.sub(r'[\[\]*?:/\\]', '', name)
    return sanitized[:MAX_SHEET_NAME_LENGTH].strip() or "Sheet"


@dataclass
class SheetEvent:
    """A new sheet was declared via '## Sheet: Name'."""
    sheet_name: str = ""
    is_rename: bool = False  # True if this renames the default first sheet


@dataclass
class HeaderEvent:
    """A markdown header line (# ... through ######)."""
    level: int = 1
    text: str = ""
    row: int = 1  # The Excel row where this header will be placed


@dataclass
class TableEvent:
    """A parsed markdown table."""
    table_data: list[list[str]] = field(default_factory=list)
    table_key: str = ""  # e.g. "T1", "T2"
    start_row: int = 1  # The Excel row where this table starts
    sheet_name: str = ""  # Which sheet this table belongs to
    directives: dict[str, str] = field(default_factory=dict)  # Comment directives above the table


# Union type for all events
LineEvent = SheetEvent | HeaderEvent | TableEvent


#: How much of a dropped line to quote back at the caller.
_DROPPED_LINE_EXCERPT = 60


def _excerpt(line: str) -> str:
    """The line, quoted and cut short enough to find it by without echoing it."""
    if len(line) > _DROPPED_LINE_EXCERPT:
        return f"'{line[:_DROPPED_LINE_EXCERPT]}…'"
    return f"'{line}'"


def walk_markdown_lines(lines: list[str], warnings=None) -> list[LineEvent]:
    """Parse markdown lines and return a list of structured events.

    This is the single source of truth for how markdown maps to Excel row
    positions. Both the position-scanning pass and the workbook-building pass
    consume these events, ensuring they never diverge.

    A line that is none of the four things a spreadsheet can hold — a heading,
    a ``## Sheet:`` marker, a directive or a table row — has nowhere to go and
    is dropped. That is the right call for a workbook, but it is content the
    caller wrote, so each dropped line is reported on *warnings* (the build's
    :class:`~warning_channel.WarningChannel`) with its 1-based source line.

    Lines that begin with ``|`` but do not form a table get their own report:
    they are consumed by :func:`~xlsx_tools.helpers.parse_table` before the
    dropped-line branch can see them, and the fix for them is a specific one —
    add the separator row — rather than "put this somewhere else".
    """
    events: list[LineEvent] = []

    current_sheet = DEFAULT_SHEET_NAME
    current_row = 1
    table_counter = 1
    first_sheet_named = False
    pending_directives: dict[str, str] = {}

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if not line:
            i += 1
            continue

        # Check for comment directives (<!-- key: value --> or <!-- key -->)
        directive_match = DIRECTIVE_PATTERN.match(line)
        if directive_match:
            key = directive_match.group(1).lower()
            value = (directive_match.group(2) or "").strip()
            pending_directives[key] = value
            i += 1
            continue

        # Check for sheet heading
        sheet_match = SHEET_HEADING_PATTERN.match(line)
        if sheet_match:
            pending_directives = {}  # Directives don't carry across sheets
            sheet_name = _sanitize_sheet_name(sheet_match.group(1).strip())
            is_rename = not first_sheet_named and current_row == 1

            events.append(SheetEvent(sheet_name=sheet_name, is_rename=is_rename))

            if is_rename:
                current_sheet = sheet_name
            else:
                current_sheet = sheet_name
                current_row = 1
                table_counter = 1

            first_sheet_named = True
            i += 1
            continue

        # Headers
        if line.startswith('#'):
            pending_directives = {}  # Directives don't carry across headers
            header_level = len(line) - len(line.lstrip('#'))
            header_text = line.lstrip('#').strip()

            events.append(HeaderEvent(level=header_level, text=header_text, row=current_row))

            current_row += HEADER_ROW_SPACING
            i += 1

        # Tables
        elif line.startswith('|'):
            table_start = i
            table_data, i = parse_table(lines, i)
            if table_data:
                if warnings is not None and not getattr(
                        table_data, "has_separator", True):
                    # parse_table() does not require the separator row, it only
                    # skips rows that look like one — so a table written
                    # without it still parses, and its first row becomes the
                    # header. Usually what the caller meant, but decided for
                    # them, and it moves every table-relative reference: those
                    # count from the first row AFTER the header.
                    warnings.add(
                        W.TABLE_SEPARATOR_MISSING,
                        "the table starting here has no separator row "
                        "(|---|---|), so its first row was used as the header "
                        "and the rest as data. Add the separator to say which "
                        "row is the header — references like T1.B[0] count "
                        "from the first row after it.",
                        sheet=current_sheet,
                        line=table_start + 1,
                    )
                table_key = f"T{table_counter}"
                events.append(TableEvent(
                    table_data=table_data,
                    table_key=table_key,
                    start_row=current_row,
                    sheet_name=current_sheet,
                    directives=pending_directives,
                ))
                current_row += len(table_data) + TABLE_BOTTOM_SPACING
                table_counter += 1
            elif warnings is not None:
                # parse_table consumed the run of pipe lines and found no table
                # in it — a row with no separator, or separators with no header.
                # The lines are gone from the workbook, and this is the only
                # branch that can say so: having been consumed here, they never
                # reach the dropped-line report below.
                warnings.add(
                    W.TABLE_INCOMPLETE,
                    f"{_excerpt(line)} starts something that looks like a "
                    f"table but is not one, so it is not in the workbook. A "
                    f"table needs a header row, a separator row (|---|---|) "
                    f"and at least one data row.",
                    sheet=current_sheet,
                    line=table_start + 1,
                )
            pending_directives = {}

        # Skip other content — directives must be directly above a table
        else:
            pending_directives = {}
            if warnings is not None:
                warnings.add(
                    W.LINE_DROPPED,
                    f"{_excerpt(line)} is not a heading, a '## Sheet:' marker, a "
                    f"directive or a table row, so it is not in the workbook. "
                    f"Put prose in a heading or a table cell.",
                    sheet=current_sheet,
                    line=i + 1,
                )
            i += 1

    return events


def collect_table_positions(events: list[LineEvent]) -> dict[str, dict[str, int]]:
    """Build the all_sheet_table_positions map from parsed events.

    Returns ``{sheet_name: {"T1": start_row, "T2": start_row, ...}}``.
    """
    all_positions: dict[str, dict[str, int]] = {}
    current_sheet = DEFAULT_SHEET_NAME
    all_positions[current_sheet] = {}

    for event in events:
        if isinstance(event, SheetEvent):
            if event.is_rename:
                # Rename default sheet key
                all_positions[event.sheet_name] = all_positions.pop(current_sheet)
                current_sheet = event.sheet_name
            else:
                current_sheet = event.sheet_name
                all_positions.setdefault(current_sheet, {})
        elif isinstance(event, TableEvent):
            sheet = event.sheet_name
            all_positions.setdefault(sheet, {})
            all_positions[sheet][event.table_key] = event.start_row

    return all_positions



