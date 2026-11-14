"""Bilingual Contract Layout Engine (.docx & .xlsx).

Merges original document bytes + Cloud Translation document bytes into 3 legal bilingual layouts:
1. side_by_side   : 2-column comparison table (Source Left | Target Right).
2. alternating    : Interleaved paragraphs/rows (Original followed immediately by Translation).
3. replace_column : Replaces target column in an existing multi-column bilingual contract
                    (e.g., converting a VI-EN 2-column contract into VI-KO by replacing col 1).
"""

from __future__ import annotations

import copy
import io
from itertools import zip_longest
from typing import Literal

import docx
from docx.document import Document as DocxDocument
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

LayoutMode = Literal["side_by_side", "alternating", "replace_column"]

LANG_LABELS: dict[str, str] = {
    "vi": "TIẾNG VIỆT (ORIGINAL/VI)",
    "ko": "한국어 (KOREAN/KO)",
    "en": "ENGLISH (EN)",
}


def build_bilingual_document(
    original_bytes: bytes,
    translated_bytes: bytes,
    extension: str,
    layout_mode: str,
    source_lang: str,
    target_lang: str,
    replace_col_index: int = 1,
) -> bytes:
    """Dispatches bilingual document assembly based on file extension and layout mode."""
    ext = extension.lower()
    if ext == ".docx":
        return _build_bilingual_docx(
            original_bytes=original_bytes,
            translated_bytes=translated_bytes,
            layout_mode=layout_mode,
            source_lang=source_lang,
            target_lang=target_lang,
            replace_col_index=replace_col_index,
        )
    if ext == ".xlsx":
        return _build_bilingual_xlsx(
            original_bytes=original_bytes,
            translated_bytes=translated_bytes,
            layout_mode=layout_mode,
            source_lang=source_lang,
            target_lang=target_lang,
            replace_col_index=replace_col_index,
        )
    raise ValueError(f"Unsupported file extension for bilingual layout: {extension}")


def build_trilingual_document(
    original_bytes: bytes,
    translated_en_bytes: bytes,
    translated_ko_bytes: bytes,
    extension: str,
    layout_mode: str,
    source_lang: str = "vi",
) -> bytes:
    """Assembles a 3-language (VI / EN / KO) document for alternating or side_by_side layouts."""
    ext = extension.lower()
    if ext == ".docx":
        if layout_mode == "side_by_side":
            out_doc = _docx_trilingual_side_by_side(
                original_bytes, translated_en_bytes, translated_ko_bytes, source_lang
            )
        elif layout_mode == "alternating":
            out_doc = _docx_trilingual_alternating(
                original_bytes, translated_en_bytes, translated_ko_bytes
            )
        else:
            raise ValueError(f"Unsupported trilingual docx layout_mode: {layout_mode}")
        buf = io.BytesIO()
        out_doc.save(buf)
        return buf.getvalue()

    if ext == ".xlsx":
        return _build_trilingual_xlsx(
            original_bytes=original_bytes,
            translated_en_bytes=translated_en_bytes,
            translated_ko_bytes=translated_ko_bytes,
            layout_mode=layout_mode,
            source_lang=source_lang,
        )

    raise ValueError(f"Unsupported file extension for trilingual layout: {extension}")


# ==============================================================================
# DOCX HELPERS & LAYOUT BUILDERS
# ==============================================================================


def _set_table_borders(table: docx.table.Table, color: str = "CBD5E1", sz: str = "4") -> None:
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), sz)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)
        borders.append(border)
    tbl_pr.append(borders)


def _set_cell_shading(cell: docx.table._Cell, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def _mark_row_as_header(row: docx.table._Row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tr_pr.append(tbl_header)


def _prevent_row_split(row: docx.table._Row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def _copy_runs_with_basic_style(
    src_para: Paragraph,
    dst_para: Paragraph,
    override_color: RGBColor | None = None,
) -> None:
    dst_para.alignment = src_para.alignment
    if not src_para.runs:
        if src_para.text:
            r = dst_para.add_run(src_para.text)
            if override_color:
                r.font.color.rgb = override_color
        return

    for run in src_para.runs:
        if not run.text:
            continue
        new_run = dst_para.add_run(run.text)
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        if run.font.size:
            new_run.font.size = run.font.size
        else:
            new_run.font.size = Pt(10.5)
        if override_color is not None:
            new_run.font.color.rgb = override_color
        elif run.font.color and run.font.color.rgb:
            new_run.font.color.rgb = run.font.color.rgb


def _insert_paragraph_after(paragraph: Paragraph, text: str = "") -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def _iter_body_blocks(doc: DocxDocument) -> list[tuple[str, object]]:
    """Yields ordered ('p', Paragraph) or ('tbl', Table) blocks from document body."""
    para_map = {p._p: p for p in doc.paragraphs}
    tbl_map = {t._tbl: t for t in doc.tables}
    blocks: list[tuple[str, object]] = []
    for child in doc.element.body.iterchildren():
        if child in para_map:
            blocks.append(("p", para_map[child]))
        elif child in tbl_map:
            blocks.append(("tbl", tbl_map[child]))
    return blocks


def _flatten_table_text(tbl: docx.table.Table) -> list[str]:
    lines: list[str] = []
    for row in tbl.rows:
        seen_tc: set[int] = set()
        cells_text: list[str] = []
        for cell in row.cells:
            if id(cell._tc) not in seen_tc:
                seen_tc.add(id(cell._tc))
                cleaned = " ".join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                cells_text.append(cleaned)
        row_line = " | ".join(c for c in cells_text if c)
        if row_line:
            lines.append(row_line)
    return lines


def _build_bilingual_docx(
    original_bytes: bytes,
    translated_bytes: bytes,
    layout_mode: str,
    source_lang: str,
    target_lang: str,
    replace_col_index: int = 1,
) -> bytes:
    if layout_mode == "side_by_side":
        out_doc = _docx_side_by_side(original_bytes, translated_bytes, source_lang, target_lang)
    elif layout_mode == "alternating":
        out_doc = _docx_alternating(original_bytes, translated_bytes)
    elif layout_mode == "replace_column":
        out_doc = _docx_replace_column(original_bytes, translated_bytes, replace_col_index)
    else:
        raise ValueError(f"Unknown layout_mode: {layout_mode}")

    buf = io.BytesIO()
    out_doc.save(buf)
    return buf.getvalue()


def _docx_side_by_side(
    original_bytes: bytes,
    translated_bytes: bytes,
    source_lang: str,
    target_lang: str,
) -> DocxDocument:
    orig_doc = docx.Document(io.BytesIO(original_bytes))
    trans_doc = docx.Document(io.BytesIO(translated_bytes))

    out_doc = docx.Document()
    for section in out_doc.sections:
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(1.6)
        section.right_margin = Cm(1.6)

    table = out_doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    _set_table_borders(table)

    # Header Row
    hdr_row = table.rows[0]
    _mark_row_as_header(hdr_row)
    src_label = LANG_LABELS.get(source_lang, source_lang.upper())
    tgt_label = LANG_LABELS.get(target_lang, target_lang.upper())

    for idx, label in enumerate((src_label, tgt_label)):
        cell = hdr_row.cells[idx]
        _set_cell_shading(cell, "1E293B")
        p = cell.paragraphs[0]
        r = p.add_run(label)
        r.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    orig_blocks = _iter_body_blocks(orig_doc)
    trans_blocks = _iter_body_blocks(trans_doc)

    for (o_item, t_item) in zip_longest(orig_blocks, trans_blocks, fillvalue=(None, None)):
        o_kind, o_obj = o_item if o_item else (None, None)
        t_kind, t_obj = t_item if t_item else (None, None)

        if o_kind == "p" or t_kind == "p":
            o_text = o_obj.text.strip() if isinstance(o_obj, Paragraph) else ""
            t_text = t_obj.text.strip() if isinstance(t_obj, Paragraph) else ""
            if not o_text and not t_text:
                continue

            row = table.add_row()
            _prevent_row_split(row)
            left_p = row.cells[0].paragraphs[0]
            right_p = row.cells[1].paragraphs[0]

            if isinstance(o_obj, Paragraph):
                _copy_runs_with_basic_style(o_obj, left_p)
            elif o_obj is not None:
                left_p.add_run("\n".join(_flatten_table_text(o_obj)))

            if isinstance(t_obj, Paragraph):
                _copy_runs_with_basic_style(t_obj, right_p)
            elif t_obj is not None:
                right_p.add_run("\n".join(_flatten_table_text(t_obj)))

        elif o_kind == "tbl" or t_kind == "tbl":
            o_lines = _flatten_table_text(o_obj) if o_obj is not None else []
            t_lines = _flatten_table_text(t_obj) if t_obj is not None else []
            for o_line, t_line in zip_longest(o_lines, t_lines, fillvalue=""):
                if not o_line and not t_line:
                    continue
                row = table.add_row()
                _prevent_row_split(row)
                _set_cell_shading(row.cells[0], "F8FAFC")
                _set_cell_shading(row.cells[1], "F8FAFC")
                r_left = row.cells[0].paragraphs[0].add_run(o_line)
                r_left.font.size = Pt(9.5)
                r_right = row.cells[1].paragraphs[0].add_run(t_line)
                r_right.font.size = Pt(9.5)

    return out_doc


def _docx_alternating(original_bytes: bytes, translated_bytes: bytes) -> DocxDocument:
    orig_doc = docx.Document(io.BytesIO(original_bytes))
    trans_doc = docx.Document(io.BytesIO(translated_bytes))
    accent_color = RGBColor(0x1E, 0x3A, 0x5F)

    # Interleave top-level paragraphs in reverse order so insertion indices stay stable
    pairs = list(zip(orig_doc.paragraphs, trans_doc.paragraphs))
    for orig_p, trans_p in reversed(pairs):
        o_txt = orig_p.text.strip()
        t_txt = trans_p.text.strip()
        if not o_txt and not t_txt:
            continue
        if t_txt and t_txt != o_txt:
            inserted = _insert_paragraph_after(orig_p)
            inserted.paragraph_format.left_indent = Cm(0.5)
            inserted.paragraph_format.space_after = Pt(6)
            _copy_runs_with_basic_style(trans_p, inserted, override_color=accent_color)

    # Interleave table cell text cleanly without duplicating merged cells
    for orig_tbl, trans_tbl in zip(orig_doc.tables, trans_doc.tables):
        seen_tc: set[int] = set()
        for o_row, t_row in zip(orig_tbl.rows, trans_tbl.rows):
            for o_cell, t_cell in zip(o_row.cells, t_row.cells):
                if id(o_cell._tc) in seen_tc:
                    continue
                seen_tc.add(id(o_cell._tc))
                o_val = o_cell.text.strip()
                t_val = t_cell.text.strip()
                if t_val and t_val != o_val:
                    new_p = o_cell.add_paragraph()
                    run = new_p.add_run(t_val)
                    run.font.size = Pt(9.5)
                    run.font.color.rgb = accent_color

    return orig_doc


def _docx_replace_column(
    original_bytes: bytes,
    translated_bytes: bytes,
    replace_col_index: int = 1,
) -> DocxDocument:
    """Replaces target column (default col 1, e.g. EN -> KO) in bilingual tables while keeping col 0 (e.g. VI)."""
    orig_doc = docx.Document(io.BytesIO(original_bytes))
    trans_doc = docx.Document(io.BytesIO(translated_bytes))

    replaced_any_table = False
    for orig_tbl, trans_tbl in zip(orig_doc.tables, trans_doc.tables):
        for o_row, t_row in zip(orig_tbl.rows, trans_tbl.rows):
            if len(o_row.cells) <= replace_col_index or len(t_row.cells) <= replace_col_index:
                continue
            target_o_cell = o_row.cells[replace_col_index]
            # Skip if horizontally merged with the kept left column
            if replace_col_index > 0 and id(target_o_cell._tc) == id(o_row.cells[replace_col_index - 1]._tc):
                continue

            # Source for replacement: if trans_doc translated the whole row, prefer replace_col_index
            # (e.g., EN column translated to KO, or VI column 0 translated to KO if user uploaded VI-EN with source=vi).
            donor_cell = t_row.cells[replace_col_index]
            if not donor_cell.text.strip() and len(t_row.cells) > 0:
                donor_cell = t_row.cells[0]

            # Clear existing paragraphs in target_o_cell except the first, then copy runs
            for extra_p in target_o_cell.paragraphs[1:]:
                p_elem = extra_p._element
                p_elem.getparent().remove(p_elem)
            first_p = target_o_cell.paragraphs[0]
            first_p.clear()

            for idx, donor_p in enumerate(donor_cell.paragraphs):
                dst_p = first_p if idx == 0 else target_o_cell.add_paragraph()
                _copy_runs_with_basic_style(donor_p, dst_p)

            replaced_any_table = True

    # Fallback: if the contract had no multi-column tables at all, build a clean 2-col side-by-side doc
    if not replaced_any_table and not orig_doc.tables:
        return _docx_side_by_side(original_bytes, translated_bytes, "orig", "trans")

    return orig_doc


# ==============================================================================
# XLSX HELPERS & LAYOUT BUILDERS
# ==============================================================================


def _copy_cell_style(src_cell: openpyxl.cell.Cell, dst_cell: openpyxl.cell.Cell) -> None:
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def _build_bilingual_xlsx(
    original_bytes: bytes,
    translated_bytes: bytes,
    layout_mode: str,
    source_lang: str,
    target_lang: str,
    replace_col_index: int = 1,
) -> bytes:
    orig_wb = openpyxl.load_workbook(io.BytesIO(original_bytes))
    trans_wb = openpyxl.load_workbook(io.BytesIO(translated_bytes))

    if layout_mode == "replace_column":
        target_col_1based = replace_col_index + 1
        for o_ws, t_ws in zip(orig_wb.worksheets, trans_wb.worksheets):
            max_r = max(o_ws.max_row or 1, t_ws.max_row or 1)
            for r in range(1, max_r + 1):
                t_val = t_ws.cell(row=r, column=target_col_1based).value
                if t_val is None and target_col_1based > 1:
                    t_val = t_ws.cell(row=r, column=1).value
                if t_val is not None:
                    o_ws.cell(row=r, column=target_col_1based).value = t_val
        out = io.BytesIO()
        orig_wb.save(out)
        return out.getvalue()

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    trans_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    hdr_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    hdr_font = Font(bold=True, color="FFFFFF")

    for o_ws, t_ws in zip(orig_wb.worksheets, trans_wb.worksheets):
        ws = out_wb.create_sheet(title=o_ws.title[:31])
        max_r = max(o_ws.max_row or 1, t_ws.max_row or 1)
        max_c = max(o_ws.max_column or 1, t_ws.max_column or 1)

        if layout_mode == "side_by_side":
            # Row 1: language group banner
            for c in range(1, max_c + 1):
                left_col = (c - 1) * 2 + 1
                right_col = left_col + 1
                c1 = ws.cell(row=1, column=left_col, value=f"Col {get_column_letter(c)} ({source_lang.upper()})")
                c2 = ws.cell(row=1, column=right_col, value=f"Col {get_column_letter(c)} ({target_lang.upper()})")
                c1.fill, c1.font = hdr_fill, hdr_font
                c2.fill, c2.font = hdr_fill, hdr_font
                ws.column_dimensions[get_column_letter(left_col)].width = 28
                ws.column_dimensions[get_column_letter(right_col)].width = 28

            for r in range(1, max_r + 1):
                dst_r = r + 1
                for c in range(1, max_c + 1):
                    o_cell = o_ws.cell(row=r, column=c)
                    t_cell = t_ws.cell(row=r, column=c)
                    d_left = ws.cell(row=dst_r, column=(c - 1) * 2 + 1, value=o_cell.value)
                    d_right = ws.cell(row=dst_r, column=(c - 1) * 2 + 2, value=t_cell.value)
                    _copy_cell_style(o_cell, d_left)
                    _copy_cell_style(t_cell, d_right)
                    d_left.alignment = Alignment(wrap_text=True, vertical="top")
                    d_right.alignment = Alignment(wrap_text=True, vertical="top")

        elif layout_mode == "alternating":
            dst_r = 1
            for r in range(1, max_r + 1):
                has_diff = False
                for c in range(1, max_c + 1):
                    o_cell = o_ws.cell(row=r, column=c)
                    t_cell = t_ws.cell(row=r, column=c)
                    d_orig = ws.cell(row=dst_r, column=c, value=o_cell.value)
                    _copy_cell_style(o_cell, d_orig)
                    if t_cell.value is not None and str(t_cell.value).strip() != str(o_cell.value or "").strip():
                        has_diff = True
                dst_r += 1

                if has_diff:
                    for c in range(1, max_c + 1):
                        t_cell = t_ws.cell(row=r, column=c)
                        d_trans = ws.cell(row=dst_r, column=c, value=t_cell.value)
                        _copy_cell_style(t_cell, d_trans)
                        d_trans.fill = trans_fill
                        d_trans.alignment = Alignment(wrap_text=True, vertical="top")
                    dst_r += 1

            for c in range(1, max_c + 1):
                col_letter = get_column_letter(c)
                orig_w = o_ws.column_dimensions[col_letter].width
                ws.column_dimensions[col_letter].width = max(orig_w or 22, 22)

    out_buf = io.BytesIO()
    out_wb.save(out_buf)
    return out_buf.getvalue()


def _docx_trilingual_side_by_side(
    original_bytes: bytes,
    translated_en_bytes: bytes,
    translated_ko_bytes: bytes,
    source_lang: str = "vi",
) -> DocxDocument:
    orig_doc = docx.Document(io.BytesIO(original_bytes))
    en_doc = docx.Document(io.BytesIO(translated_en_bytes))
    ko_doc = docx.Document(io.BytesIO(translated_ko_bytes))

    out_doc = docx.Document()
    for section in out_doc.sections:
        section.top_margin = Cm(1.6)
        section.bottom_margin = Cm(1.6)
        section.left_margin = Cm(1.2)
        section.right_margin = Cm(1.2)

    table = out_doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    _set_table_borders(table)

    hdr_row = table.rows[0]
    _mark_row_as_header(hdr_row)
    labels = (
        LANG_LABELS.get(source_lang, source_lang.upper()),
        LANG_LABELS.get("en", "ENGLISH (EN)"),
        LANG_LABELS.get("ko", "한국어 (KOREAN/KO)"),
    )
    for idx, label in enumerate(labels):
        cell = hdr_row.cells[idx]
        _set_cell_shading(cell, "1E293B")
        p = cell.paragraphs[0]
        r = p.add_run(label)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    orig_blocks = _iter_body_blocks(orig_doc)
    en_blocks = _iter_body_blocks(en_doc)
    ko_blocks = _iter_body_blocks(ko_doc)

    for o_item, e_item, k_item in zip_longest(orig_blocks, en_blocks, ko_blocks, fillvalue=(None, None)):
        o_kind, o_obj = o_item if o_item else (None, None)
        e_kind, e_obj = e_item if e_item else (None, None)
        k_kind, k_obj = k_item if k_item else (None, None)

        if "p" in (o_kind, e_kind, k_kind):
            o_text = o_obj.text.strip() if isinstance(o_obj, Paragraph) else ""
            e_text = e_obj.text.strip() if isinstance(e_obj, Paragraph) else ""
            k_text = k_obj.text.strip() if isinstance(k_obj, Paragraph) else ""
            if not o_text and not e_text and not k_text:
                continue

            row = table.add_row()
            _prevent_row_split(row)
            for col_idx, obj in enumerate((o_obj, e_obj, k_obj)):
                dst_p = row.cells[col_idx].paragraphs[0]
                if isinstance(obj, Paragraph):
                    _copy_runs_with_basic_style(obj, dst_p)
                elif obj is not None:
                    dst_p.add_run("\n".join(_flatten_table_text(obj)))
        elif "tbl" in (o_kind, e_kind, k_kind):
            o_lines = _flatten_table_text(o_obj) if o_obj is not None else []
            e_lines = _flatten_table_text(e_obj) if e_obj is not None else []
            k_lines = _flatten_table_text(k_obj) if k_obj is not None else []
            for o_line, e_line, k_line in zip_longest(o_lines, e_lines, k_lines, fillvalue=""):
                if not o_line and not e_line and not k_line:
                    continue
                row = table.add_row()
                _prevent_row_split(row)
                for col_idx, line in enumerate((o_line, e_line, k_line)):
                    _set_cell_shading(row.cells[col_idx], "F8FAFC")
                    r = row.cells[col_idx].paragraphs[0].add_run(line)
                    r.font.size = Pt(9)

    return out_doc


def _docx_trilingual_alternating(
    original_bytes: bytes,
    translated_en_bytes: bytes,
    translated_ko_bytes: bytes,
) -> DocxDocument:
    orig_doc = docx.Document(io.BytesIO(original_bytes))
    en_doc = docx.Document(io.BytesIO(translated_en_bytes))
    ko_doc = docx.Document(io.BytesIO(translated_ko_bytes))

    en_color = RGBColor(0x1D, 0x4E, 0xD8)
    ko_color = RGBColor(0x1E, 0x3A, 0x5F)

    # Reverse order so insertions after orig_p yield: orig (VI) -> EN -> KO
    triplets = list(zip(orig_doc.paragraphs, en_doc.paragraphs, ko_doc.paragraphs))
    for orig_p, en_p, ko_p in reversed(triplets):
        o_txt = orig_p.text.strip()
        e_txt = en_p.text.strip()
        k_txt = ko_p.text.strip()
        if not o_txt and not e_txt and not k_txt:
            continue
        if k_txt and k_txt != o_txt:
            inserted_ko = _insert_paragraph_after(orig_p)
            inserted_ko.paragraph_format.left_indent = Cm(0.5)
            inserted_ko.paragraph_format.space_after = Pt(6)
            _copy_runs_with_basic_style(ko_p, inserted_ko, override_color=ko_color)
        if e_txt and e_txt != o_txt:
            inserted_en = _insert_paragraph_after(orig_p)
            inserted_en.paragraph_format.left_indent = Cm(0.5)
            inserted_en.paragraph_format.space_after = Pt(2)
            _copy_runs_with_basic_style(en_p, inserted_en, override_color=en_color)

    for orig_tbl, en_tbl, ko_tbl in zip(orig_doc.tables, en_doc.tables, ko_doc.tables):
        seen_tc: set[int] = set()
        for o_row, e_row, k_row in zip(orig_tbl.rows, en_tbl.rows, ko_tbl.rows):
            for o_cell, e_cell, k_cell in zip(o_row.cells, e_row.cells, k_row.cells):
                if id(o_cell._tc) in seen_tc:
                    continue
                seen_tc.add(id(o_cell._tc))
                o_val = o_cell.text.strip()
                e_val = e_cell.text.strip()
                k_val = k_cell.text.strip()
                if e_val and e_val != o_val:
                    p_en = o_cell.add_paragraph()
                    r_en = p_en.add_run(e_val)
                    r_en.font.size = Pt(9.5)
                    r_en.font.color.rgb = en_color
                if k_val and k_val != o_val:
                    p_ko = o_cell.add_paragraph()
                    r_ko = p_ko.add_run(k_val)
                    r_ko.font.size = Pt(9.5)
                    r_ko.font.color.rgb = ko_color

    return orig_doc


def _build_trilingual_xlsx(
    original_bytes: bytes,
    translated_en_bytes: bytes,
    translated_ko_bytes: bytes,
    layout_mode: str,
    source_lang: str = "vi",
) -> bytes:
    orig_wb = openpyxl.load_workbook(io.BytesIO(original_bytes))
    en_wb = openpyxl.load_workbook(io.BytesIO(translated_en_bytes))
    ko_wb = openpyxl.load_workbook(io.BytesIO(translated_ko_bytes))

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    en_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    ko_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    hdr_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    hdr_font = Font(bold=True, color="FFFFFF")

    for o_ws, e_ws, k_ws in zip(orig_wb.worksheets, en_wb.worksheets, ko_wb.worksheets):
        ws = out_wb.create_sheet(title=o_ws.title[:31])
        max_r = max(o_ws.max_row or 1, e_ws.max_row or 1, k_ws.max_row or 1)
        max_c = max(o_ws.max_column or 1, e_ws.max_column or 1, k_ws.max_column or 1)

        if layout_mode == "side_by_side":
            for c in range(1, max_c + 1):
                base_col = (c - 1) * 3 + 1
                for offset, lang_code in enumerate((source_lang.upper(), "EN", "KO")):
                    hdr_c = ws.cell(row=1, column=base_col + offset, value=f"Col {get_column_letter(c)} ({lang_code})")
                    hdr_c.fill, hdr_c.font = hdr_fill, hdr_font
                    ws.column_dimensions[get_column_letter(base_col + offset)].width = 24

            for r in range(1, max_r + 1):
                dst_r = r + 1
                for c in range(1, max_c + 1):
                    base_col = (c - 1) * 3 + 1
                    for offset, src_ws in enumerate((o_ws, e_ws, k_ws)):
                        src_cell = src_ws.cell(row=r, column=c)
                        dst_cell = ws.cell(row=dst_r, column=base_col + offset, value=src_cell.value)
                        _copy_cell_style(src_cell, dst_cell)
                        dst_cell.alignment = Alignment(wrap_text=True, vertical="top")

        elif layout_mode == "alternating":
            dst_r = 1
            for r in range(1, max_r + 1):
                en_diff = False
                ko_diff = False
                for c in range(1, max_c + 1):
                    o_cell = o_ws.cell(row=r, column=c)
                    e_cell = e_ws.cell(row=r, column=c)
                    k_cell = k_ws.cell(row=r, column=c)
                    d_orig = ws.cell(row=dst_r, column=c, value=o_cell.value)
                    _copy_cell_style(o_cell, d_orig)
                    o_str = str(o_cell.value or "").strip()
                    if e_cell.value is not None and str(e_cell.value).strip() != o_str:
                        en_diff = True
                    if k_cell.value is not None and str(k_cell.value).strip() != o_str:
                        ko_diff = True
                dst_r += 1

                if en_diff:
                    for c in range(1, max_c + 1):
                        e_cell = e_ws.cell(row=r, column=c)
                        d_en = ws.cell(row=dst_r, column=c, value=e_cell.value)
                        _copy_cell_style(e_cell, d_en)
                        d_en.fill = en_fill
                        d_en.alignment = Alignment(wrap_text=True, vertical="top")
                    dst_r += 1

                if ko_diff:
                    for c in range(1, max_c + 1):
                        k_cell = k_ws.cell(row=r, column=c)
                        d_ko = ws.cell(row=dst_r, column=c, value=k_cell.value)
                        _copy_cell_style(k_cell, d_ko)
                        d_ko.fill = ko_fill
                        d_ko.alignment = Alignment(wrap_text=True, vertical="top")
                    dst_r += 1

            for c in range(1, max_c + 1):
                col_letter = get_column_letter(c)
                orig_w = o_ws.column_dimensions[col_letter].width
                ws.column_dimensions[col_letter].width = max(orig_w or 22, 22)
        else:
            raise ValueError(f"Unsupported trilingual xlsx layout_mode: {layout_mode}")

    out_buf = io.BytesIO()
    out_wb.save(out_buf)
    return out_buf.getvalue()
