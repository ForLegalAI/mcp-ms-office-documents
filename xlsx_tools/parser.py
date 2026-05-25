"""Shared markdown line-walking logic for xlsx tools.

Provides a single-pass generator that yields structured events, used by both
the position-scanning pass and the workbook-building pass to ensure they stay
in sync.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .helpers import parse_table


# Pattern for multi-sheet heading: ## Sheet: Name
SHEET_HEADING_PATTERN = re.compile(r'^##\s+Sheet:\s+(.+)$')
# Pattern for comment directives: <!-- key: value --> or <!-- key -->
DIRECTIVE_PATTERN = re.compile(r'^<!--\s*(\w[\w-]*)(?:\s*:\s*(.*?))?\s*-->$')

# Spacing inserted after a header row (rows)
HEADER_ROW_SPACING = 2
# Spacing inserted after a table (rows)
TABLE_BOTTOM_SPACING = 2
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


def walk_markdown_lines(lines: list[str]) -> list[LineEvent]:
    """Parse markdown lines and return a list of structured events.

    This is the single source of truth for how markdown maps to Excel row
    positions. Both the position-scanning pass and the workbook-building pass
    consume these events, ensuring they never diverge.
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
            table_data, i = parse_table(lines, i)
            if table_data:
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
            pending_directives = {}

        # Skip other content
        else:
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



