"""
analyzer.py — Single-pass document analysis.

Classifies every body paragraph into a ParaType and returns a DocAnalysis
object with counts and lists of detected elements.
"""
from __future__ import annotations

import dataclasses
import logging
import re
from enum import Enum, auto
from typing import List, Tuple

from docx import Document
from docx.text.paragraph import Paragraph

from . import config
from .utils import is_blank_para, para_has_image, para_has_math, get_style_name

log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Classification enum
# ---------------------------------------------------------------------------

class ParaType(Enum):
    BODY          = auto()
    CHAPTER       = auto()
    HEADING1      = auto()
    HEADING2      = auto()
    HEADING3      = auto()
    BULLET        = auto()
    NUMBERED      = auto()
    TABLE_CAPTION = auto()
    FIGURE_CAPTION= auto()
    EQUATION      = auto()
    IMAGE         = auto()
    FRONT_MATTER  = auto()
    BIBLIOGRAPHY  = auto()
    BLANK         = auto()


# ---------------------------------------------------------------------------
# Compiled regex helpers
# ---------------------------------------------------------------------------

def _compile_all(patterns: list[str]) -> list[re.Pattern]:
    return [re.compile(p, re.IGNORECASE) for p in patterns]


_CHAPTER_RE  = _compile_all(config.CHAPTER_PATTERNS)
_H1_RE       = _compile_all(config.HEADING1_PATTERNS)
_H2_RE       = _compile_all(config.HEADING2_PATTERNS)
_H3_RE       = _compile_all(config.HEADING3_PATTERNS)
_TBL_CAP_RE  = _compile_all(config.TABLE_CAPTION_PATTERNS)
_FIG_CAP_RE  = _compile_all(config.FIGURE_CAPTION_PATTERNS)

_BULLET_RE   = re.compile(r"^\s*[-*•◆▪]\s+")
_NUMBER_RE   = re.compile(r"^\s*\d+[.)]\s+")
_EQ_CHARS    = re.compile(r"[²³⁻¹⁰⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉±×÷≤≥≠∑∏∫√]")


def _matches_any(text: str, patterns: list[re.Pattern]) -> bool:
    for p in patterns:
        if p.match(text):
            return True
    return False


# ---------------------------------------------------------------------------
# Word style → ParaType mapping
# ---------------------------------------------------------------------------

_STYLE_MAP = {
    "Heading 1": ParaType.HEADING1,
    "Heading 2": ParaType.HEADING2,
    "Heading 3": ParaType.HEADING3,
    "Heading 4": ParaType.HEADING3,
    "Caption":   ParaType.FIGURE_CAPTION,
    "List Bullet":   ParaType.BULLET,
    "List Bullet 2": ParaType.BULLET,
    "List Bullet 3": ParaType.BULLET,
    "List Number":   ParaType.NUMBERED,
    "List Number 2": ParaType.NUMBERED,
    "List Number 3": ParaType.NUMBERED,
}

_FRONT_MATTER_STYLES = {"Title", "Subtitle", "TOC Heading"}


# ---------------------------------------------------------------------------
# DataClass for analysis results
# ---------------------------------------------------------------------------

@dataclasses.dataclass
class DocAnalysis:
    para_count:      int = 0
    table_count:     int = 0
    image_count:     int = 0
    heading_count:   int = 0
    chapter_count:   int = 0
    blank_count:     int = 0
    table_cap_count: int = 0
    figure_cap_count:int = 0
    list_count:      int = 0
    equation_count:  int = 0
    front_matter_count: int = 0

    # Detailed lists
    chapters:       List[str] = dataclasses.field(default_factory=list)
    figures:        List[str] = dataclasses.field(default_factory=list)
    table_captions: List[str] = dataclasses.field(default_factory=list)

    # Per-paragraph classification: list of (Paragraph, ParaType)
    classified:     List[Tuple[Paragraph, ParaType]] = dataclasses.field(default_factory=list)


# ---------------------------------------------------------------------------
# Main analysis function
# ---------------------------------------------------------------------------

def analyse(doc: Document) -> DocAnalysis:
    """Analyse the document and return a DocAnalysis with per-paragraph types."""
    result = DocAnalysis()
    result.table_count = len(doc.tables)

    paras = doc.paragraphs
    n = len(paras)

    for i, para in enumerate(paras):
        result.para_count += 1
        ptype = _classify(para, paras, i, n)
        result.classified.append((para, ptype))

        # counts
        if ptype == ParaType.BLANK:
            result.blank_count += 1
        elif ptype == ParaType.IMAGE:
            result.image_count += 1
        elif ptype in (ParaType.HEADING1, ParaType.HEADING2, ParaType.HEADING3):
            result.heading_count += 1
        elif ptype == ParaType.CHAPTER:
            result.chapter_count += 1
            result.chapters.append(para.text.strip())
        elif ptype == ParaType.TABLE_CAPTION:
            result.table_cap_count += 1
            result.table_captions.append(para.text.strip())
        elif ptype == ParaType.FIGURE_CAPTION:
            result.figure_cap_count += 1
            result.figures.append(para.text.strip())
        elif ptype in (ParaType.BULLET, ParaType.NUMBERED):
            result.list_count += 1
        elif ptype == ParaType.EQUATION:
            result.equation_count += 1
        elif ptype == ParaType.FRONT_MATTER:
            result.front_matter_count += 1

    return result


def _classify(para: Paragraph, paras: list, idx: int, n: int) -> ParaType:
    text       = para.text.strip()
    style_name = get_style_name(para)
    text_lower = text.lower()

    # 1. Blank
    if is_blank_para(para):
        return ParaType.BLANK

    # 2. Image
    if para_has_image(para):
        return ParaType.IMAGE

    # 3. Math
    if para_has_math(para):
        return ParaType.EQUATION

    # 4. Front matter by style
    if style_name in _FRONT_MATTER_STYLES:
        return ParaType.FRONT_MATTER

    # 5. Front matter by title text
    if text_lower in config.FRONT_MATTER_TITLES:
        return ParaType.FRONT_MATTER

    # 6. Word style map
    if style_name in _STYLE_MAP:
        mapped = _STYLE_MAP[style_name]
        # Could still be a Chapter — check text
        if mapped in (ParaType.HEADING1, ParaType.HEADING2, ParaType.HEADING3):
            if _matches_any(text, _CHAPTER_RE):
                return ParaType.CHAPTER
        return mapped

    # 7. Caption detection (before heading — short lines starting with Table/Figure)
    if _matches_any(text, _TBL_CAP_RE):
        return ParaType.TABLE_CAPTION
    if _matches_any(text, _FIG_CAP_RE):
        return ParaType.FIGURE_CAPTION

    # 8. Chapter by text pattern
    if _matches_any(text, _CHAPTER_RE):
        return ParaType.CHAPTER

    # 9. Heading by numbered sub-section pattern (most specific first)
    if _matches_any(text, _H3_RE):
        return ParaType.HEADING3
    if _matches_any(text, _H2_RE):
        return ParaType.HEADING2
    if _matches_any(text, _H1_RE):
        return ParaType.HEADING1

    # 10. List detection
    if _BULLET_RE.match(text):
        return ParaType.BULLET
    if _NUMBER_RE.match(text):
        return ParaType.NUMBERED

    # 11. Equation heuristic
    if _EQ_CHARS.search(text) and len(text) < 120:
        return ParaType.EQUATION

    # 12. Bibliography heuristic — paragraph after a "References" or "Bibliography" heading
    if idx > 0:
        prev_text = paras[idx - 1].text.strip().lower()
        if prev_text in ("references", "bibliography", "reference"):
            return ParaType.BIBLIOGRAPHY

    # 13. Default → body
    return ParaType.BODY


def print_analysis(analysis: DocAnalysis) -> None:
    """Print a formatted analysis report to stdout."""
    print("\n" + "=" * 60)
    print("DOCUMENT ANALYSIS")
    print("=" * 60)
    print(f"  Paragraphs       : {analysis.para_count}")
    print(f"  Tables           : {analysis.table_count}")
    print(f"  Images           : {analysis.image_count}")
    print(f"  Chapters         : {analysis.chapter_count}")
    print(f"  Headings (H1-H3) : {analysis.heading_count}")
    print(f"  Table captions   : {analysis.table_cap_count}")
    print(f"  Figure captions  : {analysis.figure_cap_count}")
    print(f"  Lists            : {analysis.list_count}")
    print(f"  Equations        : {analysis.equation_count}")
    print(f"  Blank paragraphs : {analysis.blank_count}")
    print(f"  Front matter     : {analysis.front_matter_count}")
    if analysis.chapters:
        print("\n  Chapters detected:")
        for c in analysis.chapters:
            print(f"    • {c}")
    if analysis.table_captions:
        print("\n  Table captions detected:")
        for c in analysis.table_captions[:10]:
            print(f"    • {c}")
    if analysis.figures:
        print("\n  Figures detected:")
        for f in analysis.figures[:10]:
            print(f"    • {f}")
    print("=" * 60 + "\n")
