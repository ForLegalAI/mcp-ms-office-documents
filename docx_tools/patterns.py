"""Compiled regex patterns and block-level markdown detection.

All block-level patterns are centralised here so that every module in the
docx_tools package can import them without circular dependencies.
"""

import re

from inline_markdown import ESCAPE_RE, LINK_RE, build_inline_pattern

# ---------------------------------------------------------------------------
# Block-level patterns (compiled once, used by many modules)
# ---------------------------------------------------------------------------

ORDERED_LIST_PATTERN = re.compile(r'^\d+\.\s+')
UNORDERED_LIST_PATTERN = re.compile(r'^[-*+]\s+')
# Capture variants used by process_list_items() to extract the item text.
# The ordered pattern captures the explicit number (group 1) so the renderer can
# restart numbering when "1." reappears at a level; group 2 is the item text.
# Item text is (.*) — matching the detection patterns above — so a marker with no
# text (e.g. "1." or "-") still captures (as empty) rather than failing the match.
ORDERED_LIST_CAPTURE_PATTERN = re.compile(r'^(\d+)\.\s+(.*)')
UNORDERED_LIST_CAPTURE_PATTERN = re.compile(r'^[-*+]\s+(.*)')
# Comment directive: <!-- key --> or <!-- key: value --> placed on its own line
# directly above the block it modifies. One mechanism for all block directives
# (borderless, widths, style, …). Group 1 = key, group 2 = optional value.
COMMENT_DIRECTIVE_PATTERN = re.compile(r'^<!--\s*([\w-]+)(?:\s*:\s*(.*?))?\s*-->$',
                                       re.IGNORECASE)
HEADING_PATTERN = re.compile(r'^(#{1,6})\s+(.+)$')
BLOCKQUOTE_PATTERN = re.compile(r'^>')
PAGE_BREAK_PATTERN = re.compile(r'^-{3,}\s*$')
HORIZONTAL_LINE_PATTERN = re.compile(r'^\*{3,}\s*$')
IMAGE_PATTERN = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')
TABLE_LINE_PATTERN = re.compile(r'^\|.+\|$')
# Fenced code block opener: 3+ backticks or tildes, optional info/language string.
# Group 1 is the fence run (its char/length identify the matching close).
CODE_FENCE_PATTERN = re.compile(r'^(`{3,}|~{3,})(.*)$')

# All block-level patterns checked by contains_block_markdown
_BLOCK_PATTERNS = [
    ORDERED_LIST_PATTERN, UNORDERED_LIST_PATTERN, HEADING_PATTERN,
    BLOCKQUOTE_PATTERN, PAGE_BREAK_PATTERN, HORIZONTAL_LINE_PATTERN,
    IMAGE_PATTERN, TABLE_LINE_PATTERN, CODE_FENCE_PATTERN,
]

# ---------------------------------------------------------------------------
# Inline formatting patterns
# ---------------------------------------------------------------------------

# The emphasis grammar is shared with the PowerPoint renderer — see
# inline_markdown.py for the rules and why there is one copy. Word can draw
# every span, so it asks for all of them. These names are kept so the five
# modules that import them need not change.
_INLINE_FORMAT_RE = build_inline_pattern(highlight=True, superscript=True, subscript=True)
_LINK_RE = LINK_RE
_ESCAPE_RE = ESCAPE_RE
# Newline spellings that must all become a real "\n": the literal escape
# sequences ("\n", "\r\n", "\r" written as text) that LLMs often emit instead of
# a real newline, and genuine CR/CRLF line endings. Longest alternative first in
# each pair so "\r\n" collapses to a single break rather than two.
_NEWLINE_RE = re.compile(r'\\r\\n|\\[nr]|\r\n|\r')

# ---------------------------------------------------------------------------
# Alignment patterns
# ---------------------------------------------------------------------------

# Inline (single-line):  <center>text</center>  or  <div align="x">text</div>
_ALIGN_INLINE_RE = re.compile(
    r'^(?:<center>(.*)</center>'
    r'|<div\s+align="(right|center|justify|left)">(.*)</div>)$',
    re.IGNORECASE,
)
# Block open:  <center>  or  <div align="x">  (content on following lines)
_ALIGN_OPEN_RE = re.compile(
    r'^(?:<center>'
    r'|<div\s+align="(right|center|justify|left)">)\s*$',
    re.IGNORECASE,
)
# Block close:  </center>  or  </div>
_ALIGN_CLOSE_RE = re.compile(r'^</(?:center|div)>\s*$', re.IGNORECASE)

# ---------------------------------------------------------------------------
# Word field / header-footer patterns
# ---------------------------------------------------------------------------

_PAGE_TOKEN_RE = re.compile(r'(\{page}|\{pages})')

# HTML <br> tag variants — the soft-break spelling, everywhere.
_BR_RE = re.compile(r'<br\s*/?>', re.IGNORECASE)
# Two or more consecutive <br> tags: a paragraph split where a real blank line
# cannot be written (inside a table cell, whose row is one physical line).
_BR_PARAGRAPH_RE = re.compile(r'(?:\s*<br\s*/?>\s*){2,}', re.IGNORECASE)

# The other soft-break spelling: a line ending in two (or more) spaces continues
# into the next line within the same paragraph.
SOFT_BREAK_SUFFIX = '  '


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------

def normalize_newlines(text: str) -> str:
    """Turn every newline spelling into a real ``\\n``.

    Two kinds are folded here:

    * **Literal escape sequences.** LLMs frequently emit a newline as the two
      characters ``\\n`` (or ``\\r\\n``) inside a tool argument instead of a real
      line break — often because tool/argument descriptions demonstrate ``\\n``
      as if it were syntax. Left untouched, the backslash handler would strip the
      slash and leave a stray ``n`` while the break is lost.
    * **CR and CRLF line endings.** The block parser splits on ``\\n`` alone, so a
      surviving ``\\r`` would sit at the end of every line and defeat the
      trailing-space soft-break check (and land in the document as a stray
      character).

    Normalising both up front means one line-ending convention reaches the rest
    of the renderer.
    """
    if not text:
        return text
    return _NEWLINE_RE.sub('\n', text)


def ordered_list_is_genuine(lines, idx) -> bool:
    """Return True if the ordered-list marker at ``lines[idx]`` should start a list.

    A numbered line begins an ordered list only when its number is ``1`` (a list
    may legitimately have a single item) OR a continuation follows — another
    sibling ordered item at the same indent, or a more-indented nested list item.
    This stops a standalone numbered line such as a date ("23. června 2026") from
    being misread as an ordered list.

    Note: a day-1 date ("1. června 2026") still matches the ``number == 1`` case
    and must be escaped ("1\\. června 2026") to render as prose; dates on days
    2–31 are handled automatically here.
    """
    raw = lines[idx]
    match = ORDERED_LIST_CAPTURE_PATTERN.match(raw.strip())
    if not match:
        return False
    if int(match.group(1)) == 1:
        return True
    base_indent = len(raw) - len(raw.lstrip())
    for nxt in lines[idx + 1:]:
        stripped = nxt.strip()
        if not stripped:
            return False  # blank line ends the run before any continuation
        indent = len(nxt) - len(nxt.lstrip())
        if indent > base_indent:
            # A more-indented list item nested under this one is a continuation.
            return bool(ORDERED_LIST_PATTERN.match(stripped)
                        or UNORDERED_LIST_PATTERN.match(stripped))
        if indent == base_indent:
            return bool(ORDERED_LIST_PATTERN.match(stripped))  # sibling ordered item
        return False  # dedent ends the run
    return False


def _segment_is_block(segments, idx, next_number=None) -> bool:
    """Return True if ``segments[idx]`` begins a block element.

    A heading, unordered-list or quote marker always counts; an ordered-list
    marker counts when :func:`ordered_list_is_genuine` accepts it within the
    *segments* list (so a lone number that is really a date does not), when the
    previous segment is an ordered marker too — ``"1. a<br>2. b"`` is a list even
    though neither ``2.`` on its own nor anything after it says so — or when it
    matches *next_number*, the count a running ordered list would continue with.
    """
    seg = segments[idx]
    if (HEADING_PATTERN.match(seg) or UNORDERED_LIST_PATTERN.match(seg)
            or BLOCKQUOTE_PATTERN.match(seg)):
        return True
    match = ORDERED_LIST_CAPTURE_PATTERN.match(seg)
    if not match:
        return False
    if next_number is not None and int(match.group(1)) == next_number:
        return True
    return bool(ordered_list_is_genuine(segments, idx)
                or (idx and ORDERED_LIST_PATTERN.match(segments[idx - 1])))


def expand_br_to_block_breaks(text: str) -> str:
    """Promote ``<br>`` to a real newline when it borders block content.

    ``<br>`` is normally an *inline* soft line break (see
    :func:`docx_tools.inline_formatting.parse_inline_formatting`). But the
    block-level parser splits only on real newlines, so a list or heading that a
    model separated from surrounding text with ``<br>`` is never recognised — a
    ``"1. ..."`` after a ``<br>`` would render as literal text instead of a
    numbered list. This pre-pass splits a physical line on ``<br>`` *only* when a
    segment **after the first** is itself a block element (a heading, or a
    genuine ordered/unordered list item); plain-prose ``<br>`` is left untouched
    so it still renders as a within-paragraph soft break.

    Only segments after the first count, because what the first one begins is
    already being parsed as a block by the caller: in ``"- item<br>second line"``
    the ``<br>`` is a break *inside* the list item, and promoting it would tear
    the item's second line out into a loose paragraph. A block that a later
    segment begins, by contrast, is invisible to the line-based parser unless it
    is promoted here.

    A numbered segment is also a block when it continues a numbered run already
    under way (``next_number`` below), so ``"Note<br>3. Treti"`` after a list that
    reached ``2.`` renders as the third item — the same thing the equivalent
    newline spelling does through the renderer's own running count. It inherits
    that rule's documented ambiguity: a date whose day happens to be the next
    number needs its dot escaped (``3\\. zari 2026``), exactly as it does when
    written on its own line.

    Table rows (``| ... |``) and fenced code blocks are never touched: a row is
    one physical line, so a ``<br>`` in a cell is an in-cell break (rendered by
    the inline layer, with ``<br><br>`` splitting the cell into paragraphs in
    ``add_table_to_doc``), and code is taken verbatim. Idempotent — running it on already-expanded text is a
    no-op, so it is safe to call on both the base and placeholder paths.
    """
    if not text or not _BR_RE.search(text):
        return text
    out = []
    in_code = False
    # The number that would continue the ordered run seen so far, tracked from
    # top-level numbered lines only (a nested item's number is not the run's).
    next_number = None

    def _remember(emitted):
        nonlocal next_number
        for produced in emitted:
            if produced[:1].isspace():
                continue
            match = ORDERED_LIST_CAPTURE_PATTERN.match(produced.strip())
            if match:
                next_number = int(match.group(1)) + 1

    for line in text.split('\n'):
        stripped = line.strip()
        if CODE_FENCE_PATTERN.match(stripped):
            in_code = not in_code
            out.append(line)
            continue
        if in_code:
            out.append(line)
            continue  # code is verbatim: a "3." in it is not part of any run
        if TABLE_LINE_PATTERN.match(stripped) or not _BR_RE.search(line):
            out.append(line)
            _remember([line])
            continue
        segments = [seg.strip() for seg in _BR_RE.split(line)]
        if len(segments) > 1 and any(_segment_is_block(segments, idx, next_number)
                                     for idx in range(1, len(segments))):
            out.extend(segments)
            _remember(segments)
        else:
            out.append(line)
            _remember([line])
    return '\n'.join(out)


def line_starts_block(lines, idx) -> bool:
    """Return True if ``lines[idx]`` begins a block-level element.

    The single answer to "is this line a block?", used both to route a template
    placeholder value (:func:`contains_block_markdown`) and to stop a soft-break
    run before a block it must not swallow
    (:func:`docx_tools.markdown_processor.process_markdown_block`).
    """
    from .block_elements import detect_alignment  # deferred to avoid circular

    stripped = lines[idx].strip()
    for pattern in _BLOCK_PATTERNS:
        if not pattern.match(stripped):
            continue
        # A lone numbered line (e.g. a date "23. června 2026") is not a list
        # unless it starts at 1 or has a continuation — keep it inline prose.
        if pattern is ORDERED_LIST_PATTERN and not ordered_list_is_genuine(lines, idx):
            continue
        return True
    return detect_alignment(stripped) is not None


def contains_block_markdown(value: str) -> bool:
    """Return True if *value* contains block-level markdown content."""
    lines = value.split('\n')
    return any(line_starts_block(lines, idx) for idx in range(len(lines)))

