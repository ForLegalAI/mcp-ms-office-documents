"""The inline-markdown grammar shared by the Word and PowerPoint renderers.

Each renderer used to carry its own copy of the emphasis regex, and the two
had drifted: Word had no flanking rules, so ``5 * 3 * 2 = 30`` italicised the
3 and ``10 ~ 20 ~ 30`` subscripted the 20; PowerPoint had flanking but
mis-parsed a bold span ending in italics (``**a *b***``), and its escape
pattern swallowed the backslash before *any* character, turning ``C:\\new``
into ``C:new``. Fixing one copy left the other wrong.

The grammar lives here once. A renderer asks for the pattern with the spans
it can actually draw (Word can highlight; PowerPoint cannot) and dispatches
on the token shapes, which are identical for both. The token *shapes* are the
contract: ``**…**``, ``*…*``, ``~~…~~``, ``__…__``, `````…`````,
``^…^``, ``~…~``, ``==…==`` and ``[label](target)``. A link target is only
made clickable when :func:`is_safe_link_target` allows its scheme.

Flanking follows CommonMark: an opening marker must be followed, and a
closing marker preceded, by something that is neither whitespace nor the
marker character itself. That is what keeps arithmetic and prose from being
read as emphasis. Nesting is stated structurally rather than with a closing
lookbehind, because a span whose last element is itself a nested span ends in
a marker character and a lookbehind would reject it.
"""

from __future__ import annotations

import re
import string
from typing import List
from urllib.parse import urlsplit

# A backslash escapes only the ASCII punctuation markdown uses as markers. It
# must not swallow the backslash before other characters: a literal "\n" is
# the two characters backslash and n, and r'\\(.)' collapsed it to a stray
# "n" — and corrupted "\t" and Windows paths like C:\new the same way.
ESCAPE_RE = re.compile(r"\\([" + re.escape(string.punctuation) + r"])")

# [label](target): the label is parsed for emphasis like any other segment;
# the target takes no nesting and no whitespace. Anchored, because a renderer
# applies it to a whole token the inline pattern has already isolated.
LINK_RE = re.compile(r"^\[([^\]\n]+)\]\(([^)\s]+)\)$")

# The schemes a link may point at. Anything else — file:, a UNC or drive path,
# a custom protocol handler, javascript: — is written as an external
# relationship that the reader's Office resolves on click, so a model steered
# by injected content could plant a link that leaks credentials or launches a
# local handler. A target with no scheme at all is refused too: Office
# resolves it as a path relative to the document. A refused link keeps its
# label as plain text. Shared, so Word and PowerPoint draw the same line.
SAFE_LINK_SCHEMES = frozenset({"http", "https", "mailto", "tel"})

# A link token anywhere in a string, for counting refusals before rendering.
# "![alt](src)" counts too: the inline renderers have no image branch, so they
# draw "!" and then a link. Only a caller that turns some of that shape into
# an image block (Word, for a line that is nothing but the image) removes
# those lines before scanning.
_LINK_ANYWHERE_RE = re.compile(r"\[[^\]\n]+\]\(([^)\s]+)\)")
_CODE_SPAN_RE = re.compile(r"`[^`]+`")


def is_safe_link_target(target: str) -> bool:
    """True when *target* may become a clickable hyperlink."""
    try:
        scheme = urlsplit(target).scheme.lower()
    except ValueError:
        return False
    return scheme in SAFE_LINK_SCHEMES


def refused_link_targets(text: str) -> List[str]:
    """The link targets in *text* a renderer will draw as plain text.

    A pre-scan for the warnings channel, so the renderers themselves need no
    channel: they refuse through :func:`is_safe_link_target` and the builder
    reports once. Code spans are skipped, since a link inside one is literal.
    """
    if "](" not in text:
        return []
    text = _CODE_SPAN_RE.sub("", text)
    return [target for target in _LINK_ANYWHERE_RE.findall(text)
            if not is_safe_link_target(target)]


def refused_links_message(targets: List[str]) -> str:
    """The warning both tools give for :func:`refused_link_targets`' result."""
    shown = ", ".join(t if len(t) <= 60 else t[:57] + "..." for t in targets[:3])
    more = f" and {len(targets) - 3} more" if len(targets) > 3 else ""
    return (f"{len(targets)} link(s) kept as plain text, not made clickable: "
            f"{shown}{more}. Only http, https, mailto and tel links are allowed.")


# One nested italic unit, usable inside a bold span: flanked on both sides.
_NESTED_ITALIC = r"\*[^\s*][^*]*?(?<=[^\s*])\*"
# One nested bold unit, usable inside an italic span.
_NESTED_BOLD = r"\*\*[^*]+\*\*"
# A single '*' that opens nothing — "**a * b**" — stays literal inside a span.
# Both renderers' old grammars had this; the first unified version dropped it,
# and a span containing one then failed to match at all. It is tried after the
# nested unit, so a real nested span is never read as a stray marker, and it
# refuses a '*' followed by another so it can never eat half of a "**" closer.
_LONE_STAR = r"\*(?!\*)"

_BOLD_ITALIC = r"\*{3}(?=[^\s*])(?:[^*]|\*(?!\*{2}))+?(?<=[^\s*])\*{3}"
# Bold, allowing a nested *italic* — including as the very last thing before
# the closer, which a plain lookbehind would reject ("**a *b***").
#
# The body takes plain characters and lone stars only; the nested-italic unit
# appears in the closer alone. A nested italic mid-span is still inside the
# token — its two stars are lone stars to the body, and the renderer parses the
# inner text again — so nothing is lost by leaving it out of the body. Keeping
# it there made every "*x" readable two ways (lone star, or the opener of a
# nested unit), so an unclosed span had exponentially many parses to reject:
# a 67-character title held the GIL for 1.5 s, doubling every two characters.
# With one reading per character the scan from each opener is linear.
_BOLD = (
    r"\*\*(?=[^\s*])(?:[^*]|" + _LONE_STAR + r")*?"
    r"(?:[^\s*]|" + _NESTED_ITALIC + r")\*\*"
)
_STRIKE = r"~~(?=[^\s~]).+?(?<=[^\s~])~~"
_HIGHLIGHT = r"==(?=[^\s=]).+?(?<=[^\s=])=="
_UNDERLINE = r"__(?=[^\s_]).+?(?<=[^\s_])__"
# Italic, allowing a nested **bold** — same structural closer as bold.
_ITALIC = (
    r"\*(?=[^\s*])(?:[^*]|" + _NESTED_BOLD + r"|" + _LONE_STAR + r")*?"
    r"(?:[^\s*]|" + _NESTED_BOLD + r")\*"
)
_CODE = r"`[^`]+`"
_SUPERSCRIPT = r"\^(?=[^\s^])[^^]+?(?<=[^\s^])\^"
# Single tilde only: the strikethrough branch has already claimed "~~".
_SUBSCRIPT = r"~(?!~)(?=[^\s~])[^~]+?(?<=[^\s~])~"
_LINK = r"\[[^\]\n]+\]\([^)\s]+\)"


def build_inline_pattern(
    *, highlight: bool = False, superscript: bool = False, subscript: bool = False
) -> re.Pattern:
    """Compile the inline pattern with the spans a renderer can draw.

    The result has one capturing group around the whole alternation, so
    ``pattern.split(text)`` yields the tokens interleaved with plain text —
    the calling convention both renderers already use. Branch order matters:
    longer markers first, so ``***`` is not read as ``**`` + ``*``, and
    strikethrough before subscript so ``~~`` is never read as two ``~``.
    """
    branches = [_BOLD_ITALIC, _BOLD, _STRIKE]
    if highlight:
        branches.append(_HIGHLIGHT)
    branches += [_UNDERLINE, _ITALIC, _CODE]
    if superscript:
        branches.append(_SUPERSCRIPT)
    if subscript:
        branches.append(_SUBSCRIPT)
    branches.append(_LINK)
    return re.compile("(" + "|".join(branches) + ")")
