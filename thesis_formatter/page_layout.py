"""
page_layout.py — Apply page size, margins, and orientation to every section.
"""
from __future__ import annotations

import logging
from docx import Document
from docx.enum.section import WD_ORIENT

from . import config

log = logging.getLogger(__name__)


def apply_page_layout(doc: Document) -> int:
    """
    Set page dimensions and margins on all sections.
    Returns the number of sections processed.
    """
    count = 0
    for section in doc.sections:
        # orientation
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width  = config.PAGE_WIDTH
        section.page_height = config.PAGE_HEIGHT

        section.top_margin    = config.MARGIN_TOP
        section.bottom_margin = config.MARGIN_BOTTOM
        section.left_margin   = config.MARGIN_LEFT
        section.right_margin  = config.MARGIN_RIGHT
        count += 1

    log.info("Page layout applied to %d section(s).", count)
    return count
