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


def clean_run_text(run: Run) -> None:
    """Normalize whitespace WITHIN a single run without touching other properties."""
    if not run.text:
        return
    text = _TAB.sub(" ", run.text)
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


# ---------------------------------------------------------------------------
# Style existence check
# ---------------------------------------------------------------------------

def style_exists(doc, style_name: str) -> bool:
    try:
        doc.styles[style_name]
        return True
    except KeyError:
        return False
