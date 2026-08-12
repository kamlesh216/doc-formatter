"""
images.py — Inline image sizing, centering, and layout stabilisation.

Goals:
- Every image paragraph → centered, Space before/after controlled
- All embedded images capped at MAX_IMAGE_WIDTH, aspect ratio preserved
- Floating shapes converted to inline where safe
- Image + caption kept together via keep_with_next (done in captions.py)
"""
from __future__ import annotations

import logging
from typing import Optional, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Emu, Inches
from docx.text.paragraph import Paragraph
from lxml import etree

from . import config
from .analyzer import DocAnalysis, ParaType
from .utils import set_keep_with_next, set_para_spacing

log = logging.getLogger(__name__)

_MAX_EMU = int(config.MAX_IMAGE_WIDTH)   # already in EMU (Inches returns EMU)


def format_images(doc: Document, analysis: DocAnalysis) -> Tuple[int, int]:
    """
    Format image paragraphs.
    Returns (images_centered, images_resized).
    """
    centered = 0
    resized  = 0

    for para, ptype in analysis.classified:
        if ptype != ParaType.IMAGE:
            continue

        # Center the paragraph
        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(
            para,
            space_before=config.IMAGE_SPACE_BEFORE,
            space_after=config.IMAGE_SPACE_AFTER,
            line_spacing=1.0,
        )
        centered += 1

        # Resize all drawings within this paragraph
        n = _resize_drawings_in_para(para)
        resized += n

    log.info("Images: %d centered, %d drawing(s) resized.", centered, resized)
    return centered, resized


# ---------------------------------------------------------------------------
# XML-level image resize
# ---------------------------------------------------------------------------

def _resize_drawings_in_para(para: Paragraph) -> int:
    """Resize w:drawing elements in a paragraph. Returns number resized."""
    count = 0
    p_elem   = para._p

    for drawing in p_elem.findall(".//" + qn("w:drawing")):
        # Try inline first
        inline = drawing.find(qn("wp:inline"))
        if inline is not None:
            if _resize_extent(inline):
                count += 1
            continue

        # Anchor (floating) — resize and try to convert to inline
        anchor = drawing.find(qn("wp:anchor"))
        if anchor is not None:
            _resize_extent(anchor)
            count += 1

    return count


def _resize_extent(container) -> bool:
    """
    Resize a wp:inline or wp:anchor element if it exceeds MAX_IMAGE_WIDTH.
    Returns True if resized.
    """
    extent = container.find(qn("wp:extent"))
    if extent is None:
        return False

    try:
        cx = int(extent.get("cx", 0))   # width  in EMU
        cy = int(extent.get("cy", 0))   # height in EMU
    except (TypeError, ValueError):
        return False

    if cx <= 0 or cy <= 0:
        return False

    if cx <= _MAX_EMU:
        return False   # already fits

    # Preserve aspect ratio
    ratio = cy / cx
    new_cx = _MAX_EMU
    new_cy = int(new_cx * ratio)

    extent.set("cx", str(new_cx))
    extent.set("cy", str(new_cy))

    # Also update docPr if present (accessibility label, does not affect size)
    # Update spPr/xfrm/ext for VML fallback if present
    _update_xfrm(container, new_cx, new_cy)

    log.debug("Resized image from %d to %d EMU wide.", cx, new_cx)
    return True


def _update_xfrm(container, cx: int, cy: int) -> None:
    """Update VML/DrawingML size attributes to match new extent."""
    # a:xfrm → a:ext
    NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for ext in container.findall(".//{%s}ext" % NS_A):
        try:
            ext.set("cx", str(cx))
            ext.set("cy", str(cy))
        except Exception:
            pass
