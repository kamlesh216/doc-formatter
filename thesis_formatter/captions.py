"""
captions.py — Detect and format table/figure captions.

Table captions → ABOVE table, keep_with_next=True
Figure captions → BELOW image, keep_with_next=True (for previous image para)
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
    set_keep_with_next,
    set_para_spacing,
)

log = logging.getLogger(__name__)


def format_captions(doc: Document, analysis: DocAnalysis) -> int:
    """Format all caption paragraphs. Returns count processed."""
    count = 0
    classified = analysis.classified

    for i, (para, ptype) in enumerate(classified):
        if ptype == ParaType.TABLE_CAPTION:
            _format_caption(para)
            set_keep_with_next(para, True)   # keep with the table below
            count += 1

        elif ptype == ParaType.FIGURE_CAPTION:
            _format_caption(para)
            # keep with whatever follows (next image or next text)
            set_keep_with_next(para, False)
            # The IMAGE paragraph before should keep_with_next → caption
            if i > 0:
                prev_para, prev_type = classified[i - 1]
                if prev_type == ParaType.IMAGE:
                    set_keep_with_next(prev_para, True)
            count += 1

    log.info("Formatted %d caption(s).", count)
    return count


def _format_caption(para: Paragraph) -> None:
    para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(
        para,
        space_before=config.CAPTION_SPACE_BEFORE,
        space_after=config.CAPTION_SPACE_AFTER,
        line_spacing=1.0,
    )
    apply_font_to_para_runs(para, config.CAPTION_FONT, config.CAPTION_SIZE)
    if config.CAPTION_BOLD:
        for run in para.runs:
            if run.bold is None:
                run.bold = True
    if config.CAPTION_ITALIC:
        for run in para.runs:
            if run.italic is None:
                run.italic = True
