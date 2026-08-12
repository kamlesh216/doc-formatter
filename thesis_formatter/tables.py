"""
tables.py — Table formatting.

For every table:
- Center the table on the page
- Apply Table Grid border style
- Repeat header row on multi-page tables
- Set font and size for all cells
- Center cell content vertically
- Cap table width at MAX_TABLE_WIDTH
"""
from __future__ import annotations

import logging
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.table import Table, _Cell

from . import config
from .utils import apply_font_to_para_runs, style_exists

log = logging.getLogger(__name__)

_BORDER_TYPES = ("top", "left", "bottom", "right", "insideH", "insideV")


def format_tables(doc: Document) -> int:
    """Format all tables in the document. Returns count processed."""
    count = 0
    for table in doc.tables:
        _format_table(doc, table)
        count += 1
    log.info("Formatted %d table(s).", count)
    return count


# ---------------------------------------------------------------------------
# Per-table
# ---------------------------------------------------------------------------

def _format_table(doc: Document, table: Table) -> None:
    # 1. Style
    if style_exists(doc, "Table Grid"):
        try:
            table.style = "Table Grid"
        except Exception:
            pass

    # 2. Autofit + max width
    table.autofit = True
    _set_table_width(table)

    # 3. Center table on page
    _center_table(table)

    # 4. Borders
    _apply_table_borders(table)

    # 5. Repeat header row
    if table.rows:
        _set_repeat_header(table.rows[0])

    # 6. Cell content
    for row in table.rows:
        for cell in row.cells:
            _format_cell(cell)


def _format_cell(cell: _Cell) -> None:
    # Vertical center
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tcPr.append(vAlign)
    vAlign.set(qn("w:val"), "center")

    # Para formatting
    for para in cell.paragraphs:
        if para.text.strip():
            # font
            apply_font_to_para_runs(para, config.TABLE_FONT, config.TABLE_FONT_SIZE)
            # spacing — tight for tables
            pf = para.paragraph_format
            if pf.space_after is None:
                pf.space_after = Pt(2)
            if pf.space_before is None:
                pf.space_before = Pt(2)


def _center_table(table: Table) -> None:
    """Set table alignment to center via tblPr/jc."""
    tbl  = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tblPr.append(jc)
    jc.set(qn("w:val"), "center")


def _set_table_width(table: Table) -> None:
    """Set preferred table width to MAX_TABLE_WIDTH (in twips = EMU/635)."""
    max_twips = int(config.MAX_TABLE_WIDTH / 914400 * 1440)  # EMU → twips
    tbl  = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"),    str(max_twips))


def _apply_table_borders(table: Table) -> None:
    """Apply uniform borders to the entire table."""
    tbl   = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders")
        tblPr.append(tblBorders)

    for side in _BORDER_TYPES:
        elem = tblBorders.find(qn(f"w:{side}"))
        if elem is None:
            elem = OxmlElement(f"w:{side}")
            tblBorders.append(elem)
        elem.set(qn("w:val"),   "single")
        elem.set(qn("w:sz"),    str(config.TABLE_BORDER_SZ))
        elem.set(qn("w:space"), "0")
        elem.set(qn("w:color"), config.TABLE_BORDER_CLR)


def _set_repeat_header(row) -> None:
    """Mark the first row of a table as a repeating header."""
    tr    = row._tr
    trPr  = tr.find(qn("w:trPr"))
    if trPr is None:
        trPr = OxmlElement("w:trPr")
        tr.insert(0, trPr)
    tblHeader = trPr.find(qn("w:tblHeader"))
    if tblHeader is None:
        tblHeader = OxmlElement("w:tblHeader")
        trPr.append(tblHeader)
