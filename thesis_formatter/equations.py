"""
equations.py — Preserve equation / formula paragraphs.

Equation paragraphs are:
- centered
- font preserved as-is (do NOT change superscript/subscript)
- spacing adjusted but content NEVER touched
"""
from __future__ import annotations

import logging
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph

from . import config
from .analyzer import DocAnalysis, ParaType
from .utils import set_para_spacing

log = logging.getLogger(__name__)


def format_equations(doc: Document, analysis: DocAnalysis) -> int:
    """Preserve and lightly format equation paragraphs. Returns count."""
    count = 0
    for para, ptype in analysis.classified:
        if ptype == ParaType.EQUATION:
            _format_eq(para)
            count += 1
    log.info("Preserved %d equation paragraph(s).", count)
    return count


def _format_eq(para: Paragraph) -> None:
    # Center equations
    para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(
        para,
        space_before=config.BODY_SPACE_BEFORE,
        space_after=config.BODY_SPACE_AFTER,
        line_spacing=config.BODY_LINE_SPACING,
    )
    # DO NOT touch runs — superscript, subscript, symbols must be preserved
