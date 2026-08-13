"""
tables.py — Production-quality table formatting.

Fixes:
1. Automatically detects and removes 100% empty trailing columns (ghost columns).
2. Fits every table strictly within MAX_TABLE_WIDTH so no table overflows page margins.
3. Scales grid/cell widths proportionally if total width exceeds MAX_TABLE_WIDTH.
4. Clears cell noWrap settings so long text wraps cleanly inside cells.
5. Centers tables on the page.
6. Applies clean Table Grid borders.
7. Repeats header rows on multi-page tables.
8. Vertically centers cell content and formats text fonts/spacing.
"""
from __future__ import annotations

import logging
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
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
# Per-table formatting pipeline
# ---------------------------------------------------------------------------

def _format_table(doc: Document, table: Table) -> None:
    # Step 1: Remove empty trailing columns (ghost columns)
    _remove_empty_trailing_columns(table)

    # Step 2: Apply Table Grid style if available
    if style_exists(doc, "Table Grid"):
        try:
            table.style = "Table Grid"
        except Exception:
            pass

    # Step 3: Enable autofit and remove fixed layout flags
    table.autofit = True
    _ensure_autofit_layout(table)

    # Step 4: Scale grid and cell widths if table exceeds MAX_TABLE_WIDTH
    _scale_table_width_if_needed(table)

    # Step 5: Center table on the page
    _center_table(table)

    # Step 6: Apply uniform table borders
    _apply_table_borders(table)

    # Step 7: Repeat header row on multi-page tables
    if table.rows:
        _set_repeat_header(table.rows[0])

    # Step 8: Format individual cells (vertical align, wrap text, font, spacing)
    for row in table.rows:
        for cell in row.cells:
            _format_cell(cell)


# ---------------------------------------------------------------------------
# 1. Remove 100% empty trailing columns
# ---------------------------------------------------------------------------

def _remove_empty_trailing_columns(table: Table) -> None:
    """Remove trailing columns that contain no text across all rows."""
    if not table.rows:
        return

    num_cols = max(len(r.cells) for r in table.rows)
    empty_col_indices = []

    # Check from rightmost column backwards
    for c_idx in range(num_cols - 1, -1, -1):
        texts = [
            r.cells[c_idx].text.strip()
            for r in table.rows
            if c_idx < len(r.cells)
        ]
        if not any(texts):
            empty_col_indices.append(c_idx)
        else:
            break  # stop at first non-empty column from right

    if not empty_col_indices:
        return

    log.info("Removing %d empty trailing column(s) from table.", len(empty_col_indices))

    for c_idx in empty_col_indices:
        for r in table.rows:
            if c_idx < len(r.cells):
                cell_elem = r.cells[c_idx]._tc
                r._tr.remove(cell_elem)

        # Also remove gridCol from tblGrid if present
        tblGrid = table._tbl.tblGrid
        if tblGrid is not None:
            grid_cols = tblGrid.findall(qn("w:gridCol"))
            if c_idx < len(grid_cols):
                tblGrid.remove(grid_cols[c_idx])


# ---------------------------------------------------------------------------
# 2. Autofit & width scaling
# ---------------------------------------------------------------------------

def _ensure_autofit_layout(table: Table) -> None:
    """Remove fixed layout flag and set tblLayout to autofit."""
    tblPr = table._tbl.tblPr
    if tblPr is not None:
        tblLayout = tblPr.find(qn("w:tblLayout"))
        if tblLayout is not None:
            tblLayout.set(qn("w:type"), "autofit")


def _scale_table_width_if_needed(table: Table) -> None:
    """
    Check total table grid width. If it exceeds MAX_TABLE_WIDTH, scale all
    gridCol and cell tcW values proportionally so total width <= MAX_TABLE_WIDTH.
    """
    max_twips = int(config.MAX_TABLE_WIDTH / 914400 * 1440)  # EMU → dxa (twips)

    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    # Set table preferred width (tblW)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(max_twips))

    # Inspect tblGrid
    tblGrid = tbl.tblGrid
    if tblGrid is not None:
        grid_cols = tblGrid.findall(qn("w:gridCol"))
        col_widths = []
        for gc in grid_cols:
            try:
                col_widths.append(int(gc.get(qn("w:w"), 0)))
            except (TypeError, ValueError):
                col_widths.append(0)

        total_grid_w = sum(col_widths)

        # Scale down if total_grid_w exceeds max_twips
        if total_grid_w > max_twips and total_grid_w > 0:
            scale = max_twips / total_grid_w
            for gc in grid_cols:
                w_orig = int(gc.get(qn("w:w"), 0))
                w_new = max(int(w_orig * scale), 200)
                gc.set(qn("w:w"), str(w_new))

    # Scale or clear fixed cell widths (tcW) so cells wrap nicely
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            # Remove w:noWrap so text wraps inside cell
            noWrap = tcPr.find(qn("w:noWrap"))
            if noWrap is not None:
                tcPr.remove(noWrap)
            # Remove or adjust hardcoded rigid cell width if it exceeds max_twips
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is not None:
                try:
                    w_val = int(tcW.get(qn("w:w"), 0))
                    if w_val > max_twips:
                        tcW.set(qn("w:w"), str(max_twips))
                except (TypeError, ValueError):
                    pass


# ---------------------------------------------------------------------------
# 3. Center table
# ---------------------------------------------------------------------------

def _center_table(table: Table) -> None:
    """Set table alignment to center via tblPr/jc."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tblPr.append(jc)
    jc.set(qn("w:val"), "center")


# ---------------------------------------------------------------------------
# 4. Table borders
# ---------------------------------------------------------------------------

def _apply_table_borders(table: Table) -> None:
    """Apply uniform borders to the entire table."""
    tbl = table._tbl
    tblPr = tbl.tblPr
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
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:sz"), str(config.TABLE_BORDER_SZ))
        elem.set(qn("w:space"), "0")
        elem.set(qn("w:color"), config.TABLE_BORDER_CLR)


# ---------------------------------------------------------------------------
# 5. Header row repeat & cell formatting
# ---------------------------------------------------------------------------

def _set_repeat_header(row) -> None:
    """Mark the first row of a table as a repeating header."""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = trPr.find(qn("w:tblHeader"))
    if tblHeader is None:
        tblHeader = OxmlElement("w:tblHeader")
        trPr.append(tblHeader)


def _format_cell(cell: _Cell) -> None:
    """Format an individual table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Vertical alignment -> center
    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tcPr.append(vAlign)
    vAlign.set(qn("w:val"), "center")

    # Format cell text paragraphs
    for para in cell.paragraphs:
        if para.text.strip():
            apply_font_to_para_runs(para, config.TABLE_FONT, config.TABLE_FONT_SIZE)
            pf = para.paragraph_format
            if pf.space_after is None:
                pf.space_after = Pt(2)
            if pf.space_before is None:
                pf.space_before = Pt(2)
