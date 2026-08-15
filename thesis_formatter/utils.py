"""
utils.py — Shared helper utilities.
All helpers are purely functional; none modify the document directly.
"""
from __future__ import annotations

import logging
import re
from typing import Optional

from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------
_MULTI_SPACE = re.compile(r"[ \t]{2,}")
_TAB         = re.compile(r"\t")


def normalize_para_spaces(para: Paragraph) -> None:
    """
    Clean all multi-spaces, non-breaking spaces (\xa0), tabs, and cross-run double spaces in a paragraph.
    Preserves text content and inline formatting.
    """
    runs = [r for r in para.runs if r.text]
    if not runs:
        return

    # Step 1: Replace non-breaking spaces (\xa0) and tabs (\t) with regular spaces, and collapse internal multi-spaces
    for r in runs:
        text = r.text.replace("\xa0", " ").replace("\t", " ")
        text = re.sub(r" {2,}", " ", text)
        if text != r.text:
            r.text = text

    # Step 2: Fix cross-run double spaces (where Run i ends with space and Run i+1 starts with space)
    for i in range(len(runs) - 1):
        r1, r2 = runs[i], runs[i + 1]
        if r1.text and r2.text and r1.text.endswith(" ") and r2.text.startswith(" "):
            r2.text = r2.text.lstrip(" ")

    # Step 3: Strip leading space from paragraph start and trailing space from paragraph end
    if runs[0].text:
        runs[0].text = runs[0].text.lstrip(" ")
    if runs[-1].text:
        runs[-1].text = runs[-1].text.rstrip(" ")


def clean_run_text(run: Run) -> None:
    """Normalize whitespace WITHIN a single run without touching other properties."""
    if not run.text:
        return
    text = run.text.replace("\xa0", " ")
    text = _TAB.sub(" ", text)
    text = _MULTI_SPACE.sub(" ", text)
    if text != run.text:
        run.text = text


def strip_leading_trailing_spaces_from_para(para: Paragraph) -> None:
    """Strip leading/trailing spaces from the first/last run of a paragraph."""
    runs = [r for r in para.runs if r.text]
    if not runs:
        return
    runs[0].text = runs[0].text.lstrip()
    runs[-1].text = runs[-1].text.rstrip()


# ---------------------------------------------------------------------------
# Pt / EMU helpers
# ---------------------------------------------------------------------------

def pt_val(value) -> Optional[float]:
    """Return float pt value from a Pt/EMU or None safely."""
    if value is None:
        return None
    try:
        return value.pt
    except AttributeError:
        return None


# ---------------------------------------------------------------------------
# Paragraph / run inspection
# ---------------------------------------------------------------------------

def is_blank_para(para: Paragraph) -> bool:
    """Return True if a paragraph has no visible text and no embedded objects."""
    text = para.text.strip()
    if text:
        return False
    # check for images / drawings
    p = para._p
    if p.findall(".//" + qn("w:drawing")):
        return False
    if p.findall(".//" + qn("w:pict")):
        return False
    if p.findall(".//" + qn("w:object")):
        return False
    return True


def para_has_image(para: Paragraph) -> bool:
    """Return True if the paragraph contains at least one embedded image."""
    p = para._p
    return bool(
        p.findall(".//" + qn("w:drawing"))
        or p.findall(".//" + qn("w:pict"))
        or p.findall(".//" + qn("w:object"))
    )


def para_has_math(para: Paragraph) -> bool:
    """Return True if the paragraph contains OMML math."""
    p = para._p
    return bool(p.findall(".//" + qn("m:oMath")))


def count_inline_images(para: Paragraph) -> int:
    return len(para._p.findall(".//" + qn("w:drawing")))


def is_all_caps_or_upper(text: str) -> bool:
    stripped = text.strip()
    return bool(stripped) and stripped == stripped.upper() and any(c.isalpha() for c in stripped)


def get_style_name(para: Paragraph) -> str:
    """Return the paragraph style name, defaulting to 'Normal'."""
    try:
        return para.style.name
    except Exception:
        return "Normal"


# ---------------------------------------------------------------------------
# Paragraph format helpers
# ---------------------------------------------------------------------------

def set_para_spacing(
    para: Paragraph,
    space_before=None,
    space_after=None,
    line_spacing=None,
) -> None:
    """Apply paragraph spacing. Any arg left None is not changed."""
    pf = para.paragraph_format
    if space_before is not None:
        pf.space_before = space_before
    if space_after is not None:
        pf.space_after = space_after
    if line_spacing is not None:
        pf.line_spacing = line_spacing


def set_para_alignment(para: Paragraph, alignment) -> None:
    para.paragraph_format.alignment = alignment


def set_keep_with_next(para: Paragraph, value: bool = True) -> None:
    para.paragraph_format.keep_with_next = value


def set_keep_together(para: Paragraph, value: bool = True) -> None:
    para.paragraph_format.keep_together = value


def set_widow_control(para: Paragraph, value: bool = True) -> None:
    para.paragraph_format.widow_control = value


# ---------------------------------------------------------------------------
# Run font helpers — only set where the run doesn't already override
# ---------------------------------------------------------------------------

def apply_font_to_run(
    run: Run,
    name: str,
    size: "Pt | None" = None,
) -> None:
    """Set font name/size on a run ONLY if the run doesn't already specify it."""
    if run.font.name is None:
        run.font.name = name
    if size is not None and run.font.size is None:
        run.font.size = size


def apply_font_to_para_runs(
    para: Paragraph,
    name: str,
    size: "Pt | None" = None,
) -> None:
    """Apply font to every run in para only where not already set."""
    for run in para.runs:
        apply_font_to_run(run, name, size)


def force_apply_font_to_para_runs(
    para: Paragraph,
    name: str,
    size: "Pt | None" = None,
) -> None:
    """Forcefully set font name and size on ALL runs in paragraph."""
    for run in para.runs:
        run.font.name = name
        if size is not None:
            run.font.size = size


def strip_body_run_formatting(run: Run, remove_bold: bool = True) -> None:
    """Clean run styling: reset font color (to black), remove character style, and strip w:b / w:bCs bold tags."""
    if hasattr(run, "_r") and run._r is not None:
        rPr = run._r.get_or_add_rPr()
        clr_tag = qn("w:color")
        b_tag = qn("w:b")
        bcs_tag = qn("w:bCs")
        rstyle_tag = qn("w:rStyle")

        # 1. Remove custom text color tags to prevent faded/grey text vs black text
        for child in list(rPr):
            if child.tag == clr_tag or child.tag.endswith("color"):
                rPr.remove(child)

        # 2. Strip bold XML elements and character styles if bold removal requested
        if remove_bold:
            for child in list(rPr):
                if child.tag in (b_tag, bcs_tag, rstyle_tag) or child.tag.endswith("b") or child.tag.endswith("bCs") or child.tag.endswith("rStyle"):
                    rPr.remove(child)
            run.bold = False


# ---------------------------------------------------------------------------
# Shading / Highlighting Clean-up
# ---------------------------------------------------------------------------

def strip_run_shading_and_highlight(run: Run) -> None:
    """Remove background fill (w:shd) and text highlight (w:highlight) from a run."""
    try:
        run.font.highlight_color = None
    except Exception:
        pass
    if hasattr(run, "_r") and run._r is not None:
        rPr = run._r.get_or_add_rPr()
        shd_tag = qn("w:shd")
        hl_tag = qn("w:highlight")
        for child in list(rPr):
            if child.tag in (shd_tag, hl_tag) or child.tag.endswith("shd") or child.tag.endswith("highlight"):
                rPr.remove(child)


def strip_para_shading_and_highlight(para: Paragraph) -> None:
    """Remove background fill and highlight from a paragraph and all its runs."""
    if hasattr(para, "_p") and para._p is not None:
        pPr = para._p.get_or_add_pPr()
        shd_tag = qn("w:shd")
        hl_tag = qn("w:highlight")
        for child in list(pPr):
            if child.tag in (shd_tag, hl_tag) or child.tag.endswith("shd") or child.tag.endswith("highlight"):
                pPr.remove(child)
    for run in para.runs:
        strip_run_shading_and_highlight(run)


# ---------------------------------------------------------------------------
# Style existence check
# ---------------------------------------------------------------------------

def style_exists(doc, style_name: str) -> bool:
    try:
        doc.styles[style_name]
        return True
    except KeyError:
        return False

