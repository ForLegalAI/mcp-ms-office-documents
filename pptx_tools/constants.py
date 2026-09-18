"""PowerPoint presentation constants.

This module contains layout indices, typography settings, colors, and margins
used throughout the PowerPoint generation.
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


# =============================================================================
# Presentation Formats
# =============================================================================
# Single source of truth for the accepted aspect ratios. Previously the tool
# handler defaulted to "16:9" while _create_presentation_buffer() defaulted to
# "4:3", so a caller that omitted the argument got a different deck depending on
# which entry point it used.

SLIDE_FORMAT_4_3 = "4:3"
SLIDE_FORMAT_16_9 = "16:9"
VALID_SLIDE_FORMATS = (SLIDE_FORMAT_4_3, SLIDE_FORMAT_16_9)
DEFAULT_SLIDE_FORMAT = SLIDE_FORMAT_16_9


# =============================================================================
# Slide Layout Indices (PowerPoint default template)
# =============================================================================

TITLE_LAYOUT = 0           # Title Slide
CONTENT_LAYOUT = 1         # Title and Content
SECTION_LAYOUT = 2         # Section Header
TWO_COLUMN_LAYOUT = 3      # Two Content (no subheaders)
TWO_COLUMN_TEXT_LAYOUT = 4 # Comparison (with subheaders)
TITLE_ONLY_LAYOUT = 5      # Title Only
BLANK_LAYOUT = 6           # Blank


# =============================================================================
# Typography
# =============================================================================

DEFAULT_SUBTITLE_FONT_SIZE = Pt(20)
DEFAULT_BODY_FONT_SIZE = Pt(18)
DEFAULT_CAPTION_FONT_SIZE = Pt(14)
DEFAULT_QUOTE_FONT_SIZE = Pt(28)

# KPI figures are meant to read from the back of a room; the label beside them
# stays at caption size so the contrast does the work.
KPI_VALUE_FONT_SIZE = Pt(40)
# Timeline step detail sits under the shape, smaller than the step label.
TIMELINE_DETAIL_FONT_SIZE = Pt(11)

# Timeline step geometry. The detail band is reserved out of the content
# rectangle *before* the shapes are sized, so a short content box shrinks the
# chevrons rather than pushing the detail lines out past the bottom of the box.
TIMELINE_STEP_HEIGHT = Inches(1.1)
TIMELINE_STEP_MIN_HEIGHT = Inches(0.3)
TIMELINE_DETAIL_GAP = Inches(0.1)
TIMELINE_DETAIL_HEIGHT = Inches(0.8)


# =============================================================================
# Bullet Indentation
# =============================================================================
# PowerPoint supports paragraph levels 0-8; the slide schema documents 1-3.
# Values outside 1..MAX_INDENT_LEVEL are clamped rather than rejected, because a
# model occasionally emits a deeper level than the prompt asked for.

MAX_INDENT_LEVEL = 5


# =============================================================================
# Autofit / overflow estimation
# =============================================================================
# python-pptx cannot lay text out, so the builder counts the lines itself.
# It does that by measuring each bullet against a real font file through
# Pillow (text_metrics.py); the ratios below are the fallback for an
# environment with no font file at all, and the line height and shrink floor
# the measured path uses too. What they drive either way is (a) a shrink
# factor written into <a:normAutofit>, which PowerPoint recomputes exactly
# when the deck is opened, and (b) a warning returned to the caller.

# Mean glyph advance as a fraction of the font size, for a mixed-case latin
# sentence in the template's body face. The fallback path only: it cannot tell
# "WWW" from "iii", which is why measuring replaced it.
AVG_CHAR_WIDTH_RATIO = 0.5
# How much of the text width each bullet level gives up to its indent.
BULLET_INDENT_INCHES = 0.3
# Baseline-to-baseline distance as a multiple of the font size.
LINE_HEIGHT_RATIO = 1.22
# Never shrink text below this fraction of its nominal size; past here the
# slide is genuinely overfull and the caller should split it.
MIN_AUTOFIT_SCALE = 0.6

# Tables get shrunk by row count rather than by text metrics.
TABLE_MIN_FONT_SIZE = 9
# The point sizes a caller may ask for. The slide schema enforces this on the
# tool call; the builder enforces it on a template default, which bypasses the
# schema and would otherwise apply a typo like 200 verbatim.
TABLE_FONT_SIZE_RANGE = (6, 40)
# Rough row height (in points) per point of font size, including cell padding.
TABLE_ROW_HEIGHT_PER_POINT = 2.1


# =============================================================================
# Table Colors
# =============================================================================

TABLE_HEADER_FILL = RGBColor(0x41, 0x72, 0xC4)
TABLE_HEADER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ALT_ROW_FILL = RGBColor(0xE9, 0xEC, 0xEF)


# =============================================================================
# Margins and Dimensions
# =============================================================================

MARGIN_LEFT = Inches(0.5)

