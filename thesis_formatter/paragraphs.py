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
    force_apply_font_to_para_runs,
    normalize_para_spaces,
    set_para_spacing,
    set_widow_control,
    strip_body_run_formatting,
    strip_para_shading_and_highlight,
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

    # Reset left/right indents to 0 (clear copy-paste indents)
    para.paragraph_format.left_indent = 0
    para.paragraph_format.right_indent = 0

    # Strip background shading / theme fill / text highlights from copy-paste
    if getattr(config, "REMOVE_BODY_SHADING", True):
        strip_para_shading_and_highlight(para)

    # Clean run text, reset discolored text colors, and strip unwanted bold (w:b, w:bCs, rStyle)
    remove_bold = getattr(config, "REMOVE_BODY_BOLD", True)

    # Smart bold stripping: if ALL runs in a body paragraph are bold,
    # this is likely an inline sub-heading (e.g., "At initial stage") — PRESERVE its bold.
    text_runs = [r for r in para.runs if r.text and r.text.strip()]
    all_bold = bool(text_runs) and all(r.bold for r in text_runs)
    effective_remove_bold = remove_bold and not all_bold

    for run in para.runs:
        clean_run_text(run)
        strip_body_run_formatting(run, remove_bold=effective_remove_bold)

    # Normalize spaces across paragraph runs (non-breaking \xa0, multi-spaces, cross-run double spaces)
    normalize_para_spaces(para)

    # Forcefully apply Arial 12pt font to all body runs
    force_apply_font_to_para_runs(para, config.BODY_FONT, config.BODY_SIZE)


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
