"""
headings.py — Detect and format chapter/heading paragraphs.

Rules:
- Chapter → centered, 16 pt bold Arial, page break before
- Heading 1 → left, 14 pt bold Arial, keep_with_next
- Heading 2 → left, 12 pt bold Arial, keep_with_next
- Heading 3 → left, 12 pt bold italic Arial, keep_with_next
"""
from __future__ import annotations

import logging
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.text.paragraph import Paragraph

from . import config
from .analyzer import DocAnalysis, ParaType
from .utils import (
    apply_font_to_para_runs,
    set_keep_with_next,
    set_para_spacing,
    set_para_alignment,
    set_widow_control,
)

log = logging.getLogger(__name__)


def format_headings(doc: Document, analysis: DocAnalysis) -> int:
    """Apply heading formatting. Returns count of headings processed."""
    count = 0
    for para, ptype in analysis.classified:
        if ptype == ParaType.CHAPTER:
            _format_chapter(para)
            count += 1
        elif ptype == ParaType.HEADING1:
            _format_h1(para)
            count += 1
        elif ptype == ParaType.HEADING2:
            _format_h2(para)
            count += 1
        elif ptype == ParaType.HEADING3:
            _format_h3(para)
            count += 1
    log.info("Formatted %d heading(s).", count)
    return count


# ---------------------------------------------------------------------------
# Per-level formatters
# ---------------------------------------------------------------------------

def _format_chapter(para: Paragraph) -> None:
    # Ensure Heading 1 style (carries TOC + numbering anchors)
    _set_style_safe(para, "Heading 1")
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(
        para,
        space_before=config.CHAPTER_SPACE_BEFORE,
        space_after=config.CHAPTER_SPACE_AFTER,
        line_spacing=1.0,
    )
    set_keep_with_next(para, True)
    set_widow_control(para, True)
    _ensure_page_break_before(para)
    # Font
    apply_font_to_para_runs(para, config.HEADING_FONT, config.CHAPTER_SIZE)
    _set_runs_bold(para, True)
    _set_heading_color(para, "Heading 1")


def _format_h1(para: Paragraph) -> None:
    _set_style_safe(para, "Heading 2")
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_para_spacing(
        para,
        space_before=config.HEADING1_SPACE_BEFORE,
        space_after=config.HEADING1_SPACE_AFTER,
        line_spacing=1.15,
    )
    set_keep_with_next(para, True)
    apply_font_to_para_runs(para, config.HEADING_FONT, config.HEADING1_SIZE)
    _set_runs_bold(para, True)
    _set_heading_color(para, "Heading 1")


def _format_h2(para: Paragraph) -> None:
    _set_style_safe(para, "Heading 3")
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_para_spacing(
        para,
        space_before=config.HEADING2_SPACE_BEFORE,
        space_after=config.HEADING2_SPACE_AFTER,
        line_spacing=1.15,
    )
    set_keep_with_next(para, True)
    apply_font_to_para_runs(para, config.HEADING_FONT, config.HEADING2_SIZE)
    _set_runs_bold(para, True)
    _set_heading_color(para, "Heading 2")


def _format_h3(para: Paragraph) -> None:
    _set_style_safe(para, "Heading 4")
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_para_spacing(
        para,
        space_before=config.HEADING3_SPACE_BEFORE,
        space_after=config.HEADING3_SPACE_AFTER,
        line_spacing=1.15,
    )
    set_keep_with_next(para, True)
    apply_font_to_para_runs(para, config.HEADING_FONT, config.HEADING3_SIZE)
    _set_runs_bold(para, True)
    if config.HEADING3_ITALIC:
        _set_runs_italic(para, True)
    _set_heading_color(para, "Heading 3")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _set_style_safe(para: Paragraph, style_name: str) -> None:
    try:
        doc = para._p.getroottree().getroot()  # not ideal but works
    except Exception:
        pass
    try:
        para.style = style_name
    except Exception:
        log.debug("Could not set style '%s' on paragraph.", style_name)


def _set_runs_bold(para: Paragraph, bold: bool) -> None:
    for run in para.runs:
        if run.bold is None:
            run.bold = bold


def _set_runs_italic(para: Paragraph, italic: bool) -> None:
    for run in para.runs:
        if run.italic is None:
            run.italic = italic


def _set_heading_color(para: Paragraph, key: str) -> None:
    hex_color = config.HEADING_COLORS.get(key)
    if not hex_color:
        return
    color = RGBColor.from_string(hex_color)
    for run in para.runs:
        try:
            if run.font.color.rgb is None:
                run.font.color.rgb = color
        except Exception:
            pass


def _ensure_page_break_before(para: Paragraph) -> None:
    """Add a page break before the paragraph via pPr/pageBreakBefore XML element."""
    pPr = para._p.get_or_add_pPr()
    # Check if already has pageBreakBefore
    existing = pPr.find(qn("w:pageBreakBefore"))
    if existing is None:
        pb = OxmlElement("w:pageBreakBefore")
        pPr.append(pb)
