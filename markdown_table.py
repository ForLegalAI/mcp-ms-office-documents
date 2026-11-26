"""The Markdown table-row grammar shared by the Word and Excel renderers.

Each renderer used to decide for itself which lines were table rows, and the
two disagreed: Excel took any line starting with ``|`` and added a missing
trailing pipe, while Word demanded both pipes. A table whose last row lacked
its trailing pipe kept every row in Excel and lost that row from the Word
table, printing it below as prose with nothing reported (#144).

The row grammar lives here once:

* :func:`is_table_line` — a row starts with ``|``. The trailing pipe is
  optional, as in GitHub-flavoured Markdown; the leading pipe is not, in
  either tool.
* :func:`split_row` — adds the missing trailing pipe, then splits the row
  into stripped cells.
* :func:`is_separator_cells` — whether split cells are a ``|---|:---:|`` row.

Only the *shape* of a row is decided here. Where a table starts, where it
ends, and which row is the separator (only the one directly under the header)
stay with each renderer's ``parse_table()``. Word's ``line_starts_block()``
deliberately reads pipe lines more strictly than :func:`is_table_line` — see
``docx_tools.patterns``. PowerPoint takes its rows as JSON arrays, not
Markdown lines, and has its own separator check.

It sits at the root, beside ``inline_markdown.py``, because ``xlsx_tools``
must not import ``docx_tools``.
"""

from __future__ import annotations

import re

_SEPARATOR_CELL_RE = re.compile(r'^:?-+:?$')


def is_table_line(stripped: str) -> bool:
    """True if the already-stripped *stripped* line is a table row.

    A row starts with ``|``; the trailing pipe is optional. A lone ``|`` is a
    row too — an empty one — so it does not end a table it sits in.
    """
    return stripped.startswith('|')


def split_row(line: str) -> list[str]:
    """Split a table row into its stripped cells.

    The line is stripped and a missing trailing ``|`` is added first, so
    ``| a | b`` and ``| a | b |`` give the same ``['a', 'b']``. The text
    before the leading pipe (none, for a line :func:`is_table_line` accepts)
    is dropped.
    """
    line = line.strip()
    if not line.endswith('|'):
        line += '|'
    return [cell.strip() for cell in line.split('|')[1:-1]]


def is_separator_cells(cells: list[str]) -> bool:
    """True if *cells* (from :func:`split_row`) are a separator row's.

    Every cell must be one or more dashes with an optional colon at either
    end (``-``, ``:--``, ``:-:``, ``--:``). Every cell, including empty ones:
    exempting those made a row of blank cells (``|  |  |``) pass as a
    separator and the caller's blank row was swallowed. A row with no cells
    is not a separator.

    Shape only — whether a row is *the* separator also depends on its
    position, which each renderer's ``parse_table()`` decides.
    """
    return bool(cells) and all(_SEPARATOR_CELL_RE.match(cell) for cell in cells)
