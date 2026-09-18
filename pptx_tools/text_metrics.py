"""Measure text with real glyph widths instead of counting characters.

The fill estimate used to be ``len(text) / chars_per_line`` with one mean
glyph width for every face and every letter (#125), so "WWW WWW WWW" and
"iii iii iii" were the same length and the shrink warnings were indicative at
best. Pillow — already installed, as python-pptx depends on it — can measure a
string against an actual font file, which turns the estimate into a
measurement wherever the face is available.

What "available" means in practice:

* the deck's own typeface, when a file of that name is installed;
* a **metric-compatible** substitute, whose glyph advances are identical by
  design: Carlito for Calibri, Liberation Sans or Arimo for Arial, Liberation
  Serif or Tinos for Times New Roman, Caladea for Cambria. Measuring one of
  these is measuring the real thing;
* failing both, any installed sans face, which is a real measurement of the
  wrong font — still shape-aware, and still far closer than a character count;
* failing everything, nothing: :func:`measure_lines` returns None and the
  caller keeps its arithmetic estimate.

Sizes are in points throughout. A font loaded at size *N* measures in pixels
that are points here, because nothing is being rasterised — only advances are
read, and those scale linearly with the requested size.
"""

from __future__ import annotations

import logging
from functools import lru_cache
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

logger = logging.getLogger(__name__)

# Faces whose advances match the named typeface exactly. Substituting one of
# these is not an approximation; they exist to be dropped in for the other.
METRIC_COMPATIBLE = {
    "calibri": ("Carlito",),
    "cambria": ("Caladea",),
    "arial": ("LiberationSans", "Liberation Sans", "Arimo"),
    "helvetica": ("LiberationSans", "Liberation Sans", "Arimo"),
    "times new roman": ("LiberationSerif", "Liberation Serif", "Tinos"),
    "courier new": ("LiberationMono", "Liberation Mono", "Cousine"),
    "georgia": ("Gelasio",),
}

# Anything else falls back to one of these: a real face, just not the one the
# deck asks for.
GENERIC_SANS = (
    "LiberationSans-Regular", "LiberationSans", "Arimo",
    "DejaVuSans", "NotoSans-Regular", "FreeSans",
)

_SUFFIXES = ("", "-Regular", "-Roman")


def _candidate_names(typeface: Optional[str]) -> List[str]:
    """Font file stems to try for *typeface*, best first."""
    names: List[str] = []
    if typeface:
        cleaned = typeface.strip()
        # "Aptos Display" and "+mj-lt" style names: the theme reference form
        # has no file of its own, so it is left to the generic fallback.
        if cleaned and not cleaned.startswith("+"):
            names.extend((cleaned, cleaned.replace(" ", "")))
            names.extend(METRIC_COMPATIBLE.get(cleaned.lower(), ()))
    names.extend(GENERIC_SANS)
    return names


@lru_cache(maxsize=64)
def load_font(typeface: Optional[str], size_pt: int):
    """A Pillow font for *typeface* at *size_pt*, or None if none can be found.

    Cached: a deck asks for the same face and size on every slide, and each
    miss costs a filesystem walk through the font directories.
    """
    try:
        from PIL import ImageFont
    except ImportError:      # pragma: no cover - Pillow ships with python-pptx
        logger.debug("Pillow is unavailable; text is estimated, not measured")
        return None

    for name in _candidate_names(typeface):
        for suffix in _SUFFIXES:
            for extension in (".ttf", ".otf"):
                try:
                    return ImageFont.truetype(f"{name}{suffix}{extension}", size_pt)
                except OSError:
                    continue
    logger.debug("No font file found for %r; text is estimated, not measured", typeface)
    return None


def font_name(font) -> Optional[str]:
    """The file a font was loaded from, for diagnostics."""
    path = getattr(font, "path", None)
    return Path(path).name if path else None


def measure_lines(
    texts: Sequence[Tuple[str, float]],
    typeface: Optional[str],
    size_pt: float,
) -> Optional[int]:
    """Total wrapped line count for ``(text, usable_width_pt)`` pairs.

    None when no font file could be loaded, which is the caller's signal to
    fall back to its own arithmetic rather than to trust a measurement that
    did not happen.
    """
    font = load_font(typeface, max(int(round(size_pt)), 1))
    if font is None:
        return None

    return sum(_wrapped_lines(text, font, width_pt) for text, width_pt in texts)


def _wrapped_lines(text: str, font, width_pt: float) -> int:
    """How many lines *text* takes in *width_pt*, wrapped at spaces.

    A word longer than the line is broken across lines the way a text box
    breaks it, rather than counted as one line and quietly under-measured.
    """
    if not text.strip():
        return 1
    if width_pt <= 0:
        return 1

    lines, current = 1, ""
    for word in text.split():
        candidate = f"{current} {word}".strip()
        if _advance(font, candidate) <= width_pt:
            current = candidate
            continue
        if current:
            lines += 1
        if _advance(font, word) <= width_pt:
            current = word
            continue
        # A single word wider than the box: count the lines it breaks into.
        whole, remainder = _broken_word_lines(font, word, width_pt)
        lines += whole
        current = remainder
    return lines


def _broken_word_lines(font, word: str, width_pt: float) -> Tuple[int, str]:
    """``(extra lines, remainder)`` for a word that does not fit one line."""
    extra, current = 0, ""
    for char in word:
        if _advance(font, current + char) <= width_pt or not current:
            current += char
            continue
        extra += 1
        current = char
    return extra, current


def _advance(font, text: str) -> float:
    """Width of *text* in points, with the font loaded at its point size."""
    try:
        return font.getlength(text)
    except AttributeError:   # pragma: no cover - very old Pillow
        return font.getsize(text)[0]


def theme_body_typeface(presentation) -> Optional[str]:
    """The minor latin font of the deck's theme — what body text is set in.

    Read from the theme part rather than assumed, because it is the whole
    point: measuring Calibri when the template sets Aptos is measuring the
    wrong face. Returns None when the theme cannot be read, which leaves the
    generic fallback to :func:`load_font`.
    """
    from lxml import etree
    from pptx.oxml.ns import qn

    try:
        master = presentation.slide_masters[0]
        theme = master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        )
        root = etree.fromstring(theme.blob)
    except Exception:
        return None

    scheme = root.find(".//" + qn("a:fontScheme"))
    if scheme is None:
        return None
    minor = scheme.find(qn("a:minorFont"))
    if minor is None:
        return None
    latin = minor.find(qn("a:latin"))
    return latin.get("typeface") if latin is not None else None
