"""
lists.py — Detect and style bullet/numbered list paragraphs.
"""
from __future__ import annotations

import logging
import re
from docx import Document
from docx.text.paragraph import Paragraph

from . import config
from .analyzer import DocAnalysis, ParaType
from .utils import apply_font_to_para_runs, set_para_spacing, style_exists

log = logging.getLogger(__name__)

_BULLET_RE  = re.compile(r"^\s*[-*•◆▪]\s+")
_NUMBER_RE  = re.compile(r"^\s*(\d+)[.)]\s+")


def format_lists(doc: Document, analysis: DocAnalysis) -> int:
    """Apply list styles. Returns count processed."""
    count = 0
    for para, ptype in analysis.classified:
        if ptype == ParaType.BULLET:
            _apply_bullet(doc, para)
            count += 1
        elif ptype == ParaType.NUMBERED:
            _apply_number(doc, para)
            count += 1
    log.info("Formatted %d list paragraph(s).", count)
    return count


def _apply_bullet(doc: Document, para: Paragraph) -> None:
    current = para.style.name if para.style else ""
    # Strip manual bullet char if style is being applied
    if _BULLET_RE.match(para.text):
        _strip_prefix(para, _BULLET_RE)
    if style_exists(doc, config.BULLET_STYLE):
        try:
            para.style = config.BULLET_STYLE
        except Exception:
            pass
    _common_list_fmt(para)


def _apply_number(doc: Document, para: Paragraph) -> None:
    if _NUMBER_RE.match(para.text):
        _strip_prefix(para, _NUMBER_RE)
    if style_exists(doc, config.NUMBER_STYLE):
        try:
            para.style = config.NUMBER_STYLE
        except Exception:
            pass
    _common_list_fmt(para)


def _common_list_fmt(para: Paragraph) -> None:
    set_para_spacing(
        para,
        space_before=None,
        space_after=config.BODY_SPACE_AFTER,
        line_spacing=config.BODY_LINE_SPACING,
    )
    apply_font_to_para_runs(para, config.LIST_FONT, config.LIST_FONT_SIZE)


def _strip_prefix(para: Paragraph, pattern: re.Pattern) -> None:
    """Remove leading bullet/number chars from the first non-empty run."""
    for run in para.runs:
        if run.text:
            run.text = pattern.sub("", run.text, count=1)
            break
