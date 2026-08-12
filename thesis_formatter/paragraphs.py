"""
paragraphs.py — Format normal body paragraphs.

CRITICAL: Does NOT call paragraph.clear() or paragraph.add_run().
           All processing is at the run level to preserve bold, italic,
           superscript, subscript, hyperlinks, fields, etc.
"""
from __future__ import annotations

import logging
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph

from . import config
from .analyzer import DocAnalysis, ParaType
from .utils import (
    apply_font_to_para_runs,
    clean_run_text,
    set_para_spacing,
    set_widow_control,
)

log = logging.getLogger(__name__)

# Paragraph types to skip (they have their own formatters)
_SKIP = frozenset({
    ParaType.CHAPTER,
    ParaType.HEADING1,
    ParaType.HEADING2,
    ParaType.HEADING3,
    ParaType.TABLE_CAPTION,
    ParaType.FIGURE_CAPTION,
    ParaType.IMAGE,
    ParaType.EQUATION,
    ParaType.FRONT_MATTER,
    ParaType.BLANK,
    ParaType.BULLET,
    ParaType.NUMBERED,
})


def format_paragraphs(doc: Document, analysis: DocAnalysis) -> int:
    """Format body paragraphs. Returns count processed."""
    count = 0
    for para, ptype in analysis.classified:
        if ptype == ParaType.BODY:
            _format_body(para)
            count += 1
        elif ptype == ParaType.BIBLIOGRAPHY:
            _format_bibliography(para)
            count += 1
    log.info("Formatted %d body paragraph(s).", count)
    return count


def _format_body(para: Paragraph) -> None:
    # Spacing & alignment
    set_para_spacing(
        para,
        space_before=config.BODY_SPACE_BEFORE,
        space_after=config.BODY_SPACE_AFTER,
        line_spacing=config.BODY_LINE_SPACING,
    )
    para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_widow_control(para, True)

    # Normalize whitespace per run — never modify content
    for run in para.runs:
        clean_run_text(run)

    # Apply font only where run doesn't already specify it
    apply_font_to_para_runs(para, config.BODY_FONT, config.BODY_SIZE)


def _format_bibliography(para: Paragraph) -> None:
    set_para_spacing(
        para,
        space_before=config.BODY_SPACE_BEFORE,
        space_after=config.BIB_SPACE_AFTER,
        line_spacing=config.BODY_LINE_SPACING,
    )
    para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_widow_control(para, True)

    for run in para.runs:
        clean_run_text(run)
    apply_font_to_para_runs(para, config.BIB_FONT, config.BIB_SIZE)
