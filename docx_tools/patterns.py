"""Compiled regex patterns and block-level markdown detection.

All block-level patterns are centralised here so that every module in the
docx_tools package can import them without circular dependencies.
"""

import re

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
    PAGE_BREAK_PATTERN, HORIZONTAL_LINE_PATTERN, IMAGE_PATTERN,
    TABLE_LINE_PATTERN, CODE_FENCE_PATTERN,
]

# ---------------------------------------------------------------------------
# Inline formatting patterns
# ---------------------------------------------------------------------------

_INLINE_FORMAT_RE = re.compile(
    r'(\*{3}(?:[^*]|\*(?!\*{2}))+\*{3}'  # ***bold italic***
    r'|\*\*(?:[^*]|\*[^*]+\*|\*(?!\*))+\*\*'  # **bold** (allows nested *italic*, incl. at the ***close)
    r'|~~.+?~~'                           # ~~strikethrough~~
    r'|==.+?=='                           # ==highlight==
    r'|__(?!_).+?__'                      # __underline__
    r'|\*(?:[^*]|\*\*[^*]+\*\*)+\*'       # *italic* (allows nested **bold**)
    r'|`[^`]+`'                           # `code`
    r'|\^[^^]+\^'                         # ^superscript^
    r'|~(?!~)[^~]+~'                      # ~subscript~ (single tilde, not ~~)
    r'|\[[^\]]*\]\([^)]*\))'             # [link](url)
)

_LINK_RE = re.compile(r'\[(.*?)]\((.*?)\)')        # [link text](url)
_ESCAPE_RE = re.compile(r'\\(.)')                   # backslash-escaped character

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

# HTML <br> tag variants (used for line breaks inside table cells etc.)
_BR_RE = re.compile(r'<br\s*/?>', re.IGNORECASE)


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------

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


def contains_block_markdown(value: str) -> bool:
    """Return True if *value* contains block-level markdown content."""
    from .block_elements import detect_alignment  # deferred to avoid circular

    lines = value.split('\n')
    for idx, line in enumerate(lines):
        stripped = line.strip()
        for pattern in _BLOCK_PATTERNS:
            if not pattern.match(stripped):
                continue
            # A lone numbered line (e.g. a date "23. června 2026") is not a list
            # unless it starts at 1 or has a continuation — keep it inline prose.
            if pattern is ORDERED_LIST_PATTERN and not ordered_list_is_genuine(lines, idx):
                continue
            return True
        if detect_alignment(stripped) is not None:
            return True
    return False

