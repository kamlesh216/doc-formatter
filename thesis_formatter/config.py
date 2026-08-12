"""
config.py — All formatting constants for the thesis formatter.

Edit this file to change the visual style. Nothing is hard-coded in the
implementation modules; they all import from here.
"""

from docx.shared import Inches, Pt, RGBColor

# ---------------------------------------------------------------------------
# Page layout
# ---------------------------------------------------------------------------
PAGE_WIDTH  = Inches(8.27)   # A4 width
PAGE_HEIGHT = Inches(11.69)  # A4 height

MARGIN_TOP    = Inches(1.0)
MARGIN_BOTTOM = Inches(1.0)
MARGIN_LEFT   = Inches(1.25)
MARGIN_RIGHT  = Inches(1.0)

PRINTABLE_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT   # ≈ 6.02 in

# ---------------------------------------------------------------------------
# Body text
# ---------------------------------------------------------------------------
BODY_FONT          = "Arial"
BODY_SIZE          = Pt(12)
BODY_LINE_SPACING  = 1.15       # multiple (not Pt)
BODY_SPACE_BEFORE  = Pt(0)
BODY_SPACE_AFTER   = Pt(6)

# ---------------------------------------------------------------------------
# Heading fonts & sizes
# ---------------------------------------------------------------------------
HEADING_FONT = "Arial"

CHAPTER_SIZE        = Pt(16)
CHAPTER_BOLD        = True
CHAPTER_SPACE_BEFORE = Pt(0)
CHAPTER_SPACE_AFTER  = Pt(12)

HEADING1_SIZE        = Pt(14)
HEADING1_BOLD        = True
HEADING1_SPACE_BEFORE = Pt(12)
HEADING1_SPACE_AFTER  = Pt(6)

HEADING2_SIZE        = Pt(12)
HEADING2_BOLD        = True
HEADING2_ITALIC      = False
HEADING2_SPACE_BEFORE = Pt(10)
HEADING2_SPACE_AFTER  = Pt(4)

HEADING3_SIZE        = Pt(12)
HEADING3_BOLD        = True
HEADING3_ITALIC      = True
HEADING3_SPACE_BEFORE = Pt(8)
HEADING3_SPACE_AFTER  = Pt(4)

# Heading colors  (None = inherit / black)
HEADING_COLORS = {
    "Chapter":   "1F4E79",
    "Heading 1": "1F4E79",
    "Heading 2": "2E74B5",
    "Heading 3": "2E74B5",
}

# ---------------------------------------------------------------------------
# Captions
# ---------------------------------------------------------------------------
CAPTION_FONT       = "Arial"
CAPTION_SIZE       = Pt(11)
CAPTION_BOLD       = False
CAPTION_ITALIC     = False
CAPTION_SPACE_BEFORE = Pt(4)
CAPTION_SPACE_AFTER  = Pt(4)

# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------
MAX_TABLE_WIDTH  = PRINTABLE_WIDTH        # fit within printable area
TABLE_FONT       = "Arial"
TABLE_FONT_SIZE  = Pt(10)
TABLE_BORDER_SZ  = 4                      # eighths of a point (0.5 pt)
TABLE_BORDER_CLR = "000000"

# ---------------------------------------------------------------------------
# Images / Figures
# ---------------------------------------------------------------------------
MAX_IMAGE_WIDTH  = PRINTABLE_WIDTH        # cap at printable width
IMAGE_SPACE_BEFORE = Pt(6)
IMAGE_SPACE_AFTER  = Pt(6)

# ---------------------------------------------------------------------------
# Lists
# ---------------------------------------------------------------------------
BULLET_STYLE  = "List Bullet"
NUMBER_STYLE  = "List Number"
LIST_FONT     = BODY_FONT
LIST_FONT_SIZE = BODY_SIZE

# ---------------------------------------------------------------------------
# Bibliography
# ---------------------------------------------------------------------------
BIB_FONT        = BODY_FONT
BIB_SIZE        = BODY_SIZE
BIB_SPACE_AFTER = Pt(4)
BIB_FIRST_LINE_INDENT = Pt(0)
BIB_HANGING_INDENT    = Pt(0)   # set > 0 for hanging indent style

# ---------------------------------------------------------------------------
# Caption regex patterns  (compiled in captions.py)
# ---------------------------------------------------------------------------
TABLE_CAPTION_PATTERNS = [
    r"^Table\s+[\d.]+",
    r"^TABLE\s+[\d.]+",
    r"^TABLE\b",
]
FIGURE_CAPTION_PATTERNS = [
    r"^Figure\s+[\d.]+",
    r"^FIGURE\s+[\d.]+",
    r"^Fig\.\s*[\d.]+",
    r"^Fig\s+[\d.]+",
    r"^Plate\s+[\d.]+",
    r"^PLATE\s+[\d.]+",
    r"^Graph\s+[\d.]+",
    r"^Chart\s+[\d.]+",
]

# ---------------------------------------------------------------------------
# Chapter / Heading detection patterns
# ---------------------------------------------------------------------------
CHAPTER_PATTERNS = [
    r"^Chapter\s*[-–—]\s*[IVXivx\d]+",
    r"^CHAPTER\s+[IVXivx\d]+",
    r"^CHAPTER\b",
]

HEADING1_PATTERNS = [
    r"^\d+\.\d+\s+\S",          # 4.1 Growth …
]
HEADING2_PATTERNS = [
    r"^\d+\.\d+\.\d+\s+\S",     # 4.1.1 Plant height …
]
HEADING3_PATTERNS = [
    r"^\d+\.\d+\.\d+\.\d+\s+\S",
]

# ---------------------------------------------------------------------------
# Front-matter section titles (case-insensitive match)
# ---------------------------------------------------------------------------
FRONT_MATTER_TITLES = {
    "title page", "certificate", "declaration", "acknowledgement",
    "acknowledgements", "contents", "table of contents",
    "list of tables", "list of figures", "list of plates",
    "list of symbols", "abstract", "abbreviations",
}

# ---------------------------------------------------------------------------
# Safe whitespace cleanup
# ---------------------------------------------------------------------------
MAX_CONSECUTIVE_BLANK_PARAS = 1   # allow at most 1 blank para in a row
