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
ORDERED_LIST_CAPTURE_PATTERN = re.compile(r'^(\d+)\.\s+(.+)')
UNORDERED_LIST_CAPTURE_PATTERN = re.compile(r'^[-*+]\s+(.+)')
# Ad-hoc style directive: <!-- style: My Style --> applies the named style to the
# next block (issue #66, part B). Group 1 is the style name.
STYLE_DIRECTIVE_PATTERN = re.compile(r'^<!--\s*style:\s*(.+?)\s*-->$', re.IGNORECASE)
HEADING_PATTERN = re.compile(r'^(#{1,6})\s+(.+)$')
PAGE_BREAK_PATTERN = re.compile(r'^-{3,}\s*$')
HORIZONTAL_LINE_PATTERN = re.compile(r'^\*{3,}\s*$')
IMAGE_PATTERN = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')
TABLE_LINE_PATTERN = re.compile(r'^\|.+\|$')

# All block-level patterns checked by contains_block_markdown
_BLOCK_PATTERNS = [
    ORDERED_LIST_PATTERN, UNORDERED_LIST_PATTERN, HEADING_PATTERN,
    PAGE_BREAK_PATTERN, HORIZONTAL_LINE_PATTERN, IMAGE_PATTERN,
    TABLE_LINE_PATTERN,
]

# ---------------------------------------------------------------------------
# Inline formatting patterns
# ---------------------------------------------------------------------------

_INLINE_FORMAT_RE = re.compile(
    r'(\*{3}(?:[^*]|\*(?!\*{2}))+\*{3}'  # ***bold italic***
    r'|\*\*(?:[^*]|\*(?!\*))+\*\*'       # **bold**
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

def contains_block_markdown(value: str) -> bool:
    """Return True if *value* contains block-level markdown content."""
    from .block_elements import detect_alignment  # deferred to avoid circular

    for line in value.split('\n'):
        stripped = line.strip()
        if any(p.match(stripped) for p in _BLOCK_PATTERNS):
            return True
        if detect_alignment(stripped) is not None:
            return True
    return False

