"""Bilingual & Trilingual Contract Layout Engine (.docx & .xlsx).

Merges original document bytes + Cloud Translation document bytes into 3 legal layouts:
1. alternating    : Interleaved paragraphs/rows (Original followed immediately by Translation).
                    For vi-ko on bilingual VI+EN originals: keeps VI, overwrites EN with VI->KO.
                    For vi-en-kr on bilingual VI+EN originals: keeps VI and original EN untouched, adds KO.
2. replace_column : Replaces target column in an existing multi-column bilingual contract
                    (e.g., vi-ko replaces EN col with KO; vi-en-kr keeps VI & EN cols + adds KO col).
3. side_by_side   : Comparison table (2 cols for bilingual, 3 cols VI | EN | KO for trilingual).
"""

from __future__ import annotations

import copy
import io
from itertools import zip_longest
from typing import Literal, Optional

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

from app.en_checker import (
    _table_column_languages,
    _xlsx_sheet_structure,
    detect_segment_lang_heuristic,
    pair_paragraph_blocks,
    resolve_heuristic_fallback,
)

LayoutMode = Literal["side_by_side", "alternating", "replace_column"]

LANG_LABELS: dict[str, str] = {
    "vi": "TIẾNG VIỆT (ORIGINAL/VI)",
    "ko": "한국어 (KOREAN/KO)",
    "en": "ENGLISH (EN)",
}

_KO_ACCENT_COLOR = RGBColor(0x1E, 0x3A, 0x5F)
_EN_ACCENT_COLOR = RGBColor(0x1D, 0x4E, 0xD8)


def build_bilingual_document(
    original_bytes: bytes,
    translated_bytes: bytes,
    extension: str,
    layout_mode: str,
    source_lang: str,
    target_lang: str,
    replace_col_index: int = 1,
    original_has_en: Optional[bool] = None,
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
            original_has_en=original_has_en,
        )
    if ext == ".xlsx":
        return _build_bilingual_xlsx(
            original_bytes=original_bytes,
            translated_bytes=translated_bytes,
            layout_mode=layout_mode,
            source_lang=source_lang,
            target_lang=target_lang,
            replace_col_index=replace_col_index,
            original_has_en=original_has_en,
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
    """Assembles a 3-language (VI / EN / KO) document for VI-only originals."""
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
        elif layout_mode == "replace_column":
            # VI-only original has no EN column to replace; fall back to clean 3-column table
            out_doc = _docx_trilingual_side_by_side(
                original_bytes, translated_en_bytes, translated_ko_bytes, source_lang
            )
        else:
            raise ValueError(f"Unsupported trilingual docx layout_mode: {layout_mode}")
        buf = io.BytesIO()
        out_doc.save(buf)
        return buf.getvalue()

    if ext == ".xlsx":
        effective_mode = "side_by_side" if layout_mode == "replace_column" else layout_mode
        return _build_trilingual_xlsx(
            original_bytes=original_bytes,
            translated_en_bytes=translated_en_bytes,
            translated_ko_bytes=translated_ko_bytes,
            layout_mode=effective_mode,
            source_lang=source_lang,
        )

    raise ValueError(f"Unsupported file extension for trilingual layout: {extension}")


# =============================================================================
# TRILINGUAL (VI + EN ORIGINAL) ASSEMBLY
# =============================================================================


def _classify_paragraphs_local(paragraphs: list[Paragraph]) -> list[str]:
    res: list[str] = []
    for p in paragraphs:
        txt = p.text.strip()
        if not txt:
            res.append("_empty")
        else:
            res.append(resolve_heuristic_fallback(txt, detect_segment_lang_heuristic(txt)))
    return res


def build_trilingual_from_bilingual_document(
    original_bytes: bytes,
    translated_ko_bytes: bytes,
    extension: str,
    layout_mode: str,
    translated_en_bytes: Optional[bytes] = None,
    source_lang: str = "vi",
    replace_col_index: int = 1,
) -> bytes:
    """Assembles VI / EN (original untouched) / KO from a bilingual VI+EN original."""
    ext = extension.lower()
    if ext == ".docx":
        if layout_mode == "side_by_side":
            out_doc = _docx_trilingual_from_bilingual_side_by_side(
                original_bytes=original_bytes,
                translated_ko_bytes=translated_ko_bytes,
                translated_en_bytes=translated_en_bytes,
                source_lang=source_lang,
            )
        elif layout_mode in ("alternating", "replace_column"):
            out_doc = _docx_trilingual_from_bilingual_in_place(
                original_bytes=original_bytes,
                translated_ko_bytes=translated_ko_bytes,
                translated_en_bytes=translated_en_bytes,
                layout_mode=layout_mode,
                replace_col_index=replace_col_index,
            )
        else:
            raise ValueError(f"Unsupported trilingual docx layout_mode: {layout_mode}")
        buf = io.BytesIO()
        out_doc.save(buf)
        return buf.getvalue()

    if ext == ".xlsx":
        return _build_trilingual_from_bilingual_xlsx(
            original_bytes=original_bytes,
            translated_ko_bytes=translated_ko_bytes,
            translated_en_bytes=translated_en_bytes,
            layout_mode=layout_mode,
            source_lang=source_lang,
            replace_col_index=replace_col_index,
        )

    raise ValueError(f"Unsupported file extension for trilingual layout: {extension}")


# =============================================================================
# DOCX HELPERS & LAYOUT BUILDERS
# =============================================================================


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


def _replace_paragraph_with_source(
    dst_para: Paragraph,
    src_para: Paragraph,
    override_color: RGBColor | None = None,
) -> None:
    dst_para.clear()
    _copy_runs_with_basic_style(src_para, dst_para, override_color=override_color)


def _replace_paragraph_text(
    dst_para: Paragraph,
    new_text: str,
    override_color: RGBColor | None = None,
) -> None:
    """Replaces text inside dst_para while preserving basic run formatting from its first run."""
    first_bold = dst_para.runs[0].bold if dst_para.runs else None
    first_italic = dst_para.runs[0].italic if dst_para.runs else None
    first_size = dst_para.runs[0].font.size if dst_para.runs else Pt(10.5)
    dst_para.clear()
    r = dst_para.add_run(new_text)
    r.bold = first_bold
    r.italic = first_italic
    r.font.size = first_size or Pt(10.5)
    if override_color is not None:
        r.font.color.rgb = override_color


def _remove_paragraph(paragraph: Paragraph) -> None:
    elem = paragraph._p
    parent = elem.getparent()
    if parent is not None:
        parent.remove(elem)


def _insert_paragraph_after(paragraph: Paragraph, text: str = "") -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def _copy_cell_paragraphs(
    donor_cell: docx.table._Cell,
    target_cell: docx.table._Cell,
    override_color: RGBColor | None = None,
) -> None:
    for extra_p in target_cell.paragraphs[1:]:
        _remove_paragraph(extra_p)
    first_p = target_cell.paragraphs[0]
    first_p.clear()
    for idx, donor_p in enumerate(donor_cell.paragraphs):
        dst_p = first_p if idx == 0 else target_cell.add_paragraph()
        _copy_runs_with_basic_style(donor_p, dst_p, override_color=override_color)


def _set_cell_text_preserve_first(
    target_cell: docx.table._Cell,
    new_text: str,
    override_color: RGBColor | None = None,
) -> None:
    for extra_p in target_cell.paragraphs[1:]:
        _remove_paragraph(extra_p)
    first_p = target_cell.paragraphs[0]
    _replace_paragraph_text(first_p, new_text, override_color=override_color)


def _append_column_to_table(
    orig_tbl: docx.table.Table,
    ko_texts_by_row: list[Optional[docx.table._Cell]],
) -> None:
    """Appends a 3rd KO column to an existing DOCX table, handling merged header rows safely."""
    tbl_el = orig_tbl._tbl
    tbl_grid = tbl_el.tblGrid
    expected_cols = len(tbl_grid.gridCol_lst) if tbl_grid is not None else 2

    if tbl_grid is not None:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), "2800")
        tbl_grid.append(grid_col)

    for r_idx, row in enumerate(orig_tbl.rows):
        tr = row._tr
        tcs = tr.tc_lst
        # Full-width merged row (e.g., section header spanning all columns)
        if len(tcs) == 1 and expected_cols > 1:
            tc_pr = tcs[0].get_or_add_tcPr()
            grid_span = tc_pr.find(qn("w:gridSpan"))
            if grid_span is not None:
                curr = int(grid_span.get(qn("w:val"), str(expected_cols)))
                grid_span.set(qn("w:val"), str(curr + 1))
                continue

        new_tc = OxmlElement("w:tc")
        # Copy tcPr borders/shading structure if available from last tc (strip gridSpan/vMerge)
        if tcs:
            ref_pr = tcs[-1].find(qn("w:tcPr"))
            if ref_pr is not None:
                cloned_pr = copy.deepcopy(ref_pr)
                for tag in ("w:gridSpan", "w:vMerge", "w:hMerge"):
                    child = cloned_pr.find(qn(tag))
                    if child is not None:
                        cloned_pr.remove(child)
                new_tc.append(cloned_pr)

        p_el = OxmlElement("w:p")
        new_tc.append(p_el)
        tr.append(new_tc)

        donor_cell = ko_texts_by_row[r_idx] if r_idx < len(ko_texts_by_row) else None
        if donor_cell is not None:
            new_cell = docx.table._Cell(new_tc, orig_tbl)
            _copy_cell_paragraphs(donor_cell, new_cell, override_color=_KO_ACCENT_COLOR)


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
    original_has_en: Optional[bool] = None,
) -> bytes:
    if layout_mode == "side_by_side":
        out_doc = _docx_side_by_side(original_bytes, translated_bytes, source_lang, target_lang)
    elif layout_mode == "alternating":
        if source_lang == "vi" and target_lang == "ko":
            out_doc = _docx_alternating_vi_ko(
                original_bytes, translated_bytes, original_has_en=original_has_en
            )
        else:
            out_doc = _docx_alternating(original_bytes, translated_bytes)
    elif layout_mode == "replace_column":
        out_doc = _docx_replace_column(
            original_bytes,
            translated_bytes,
            replace_col_index=replace_col_index,
            source_lang=source_lang,
        )
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
            _copy_runs_with_basic_style(trans_p, inserted, override_color=_KO_ACCENT_COLOR)

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
                    run.font.color.rgb = _KO_ACCENT_COLOR

    return orig_doc


def _overwrite_cell_en_paragraphs_with_ko(
    o_cell: docx.table._Cell,
    t_cell: docx.table._Cell,
) -> None:
    """For a cell containing stacked VI+EN paragraphs, keeps VI and overwrites EN with VI->KO."""
    c_langs = _classify_paragraphs_local(o_cell.paragraphs)
    blocks = pair_paragraph_blocks(c_langs)
    has_en = any(en_idxs for _, en_idxs in blocks)
    if not has_en:
        o_val = o_cell.text.strip()
        t_val = t_cell.text.strip()
        if t_val and t_val != o_val:
            new_p = o_cell.add_paragraph()
            run = new_p.add_run(t_val)
            run.font.size = Pt(9.5)
            run.font.color.rgb = _KO_ACCENT_COLOR
        return

    for vi_idxs, en_idxs in reversed(blocks):
        if vi_idxs and en_idxs:
            leading_en = en_idxs[0] < vi_idxs[0]
            if len(vi_idxs) == 1 and not leading_en:
                v_i = vi_idxs[0]
                t_p = t_cell.paragraphs[v_i] if v_i < len(t_cell.paragraphs) else None
                for extra_e in reversed(en_idxs[1:]):
                    _remove_paragraph(o_cell.paragraphs[extra_e])
                if t_p is not None and t_p.text.strip():
                    _replace_paragraph_with_source(
                        o_cell.paragraphs[en_idxs[0]], t_p, override_color=_KO_ACCENT_COLOR
                    )
            else:
                for e_i in reversed(en_idxs):
                    _remove_paragraph(o_cell.paragraphs[e_i])
                for v_i in reversed(vi_idxs):
                    t_p = t_cell.paragraphs[v_i] if v_i < len(t_cell.paragraphs) else None
                    if t_p is not None and t_p.text.strip():
                        ins = _insert_paragraph_after(o_cell.paragraphs[v_i])
                        _copy_runs_with_basic_style(t_p, ins, override_color=_KO_ACCENT_COLOR)
        elif vi_idxs:
            for v_i in reversed(vi_idxs):
                t_p = t_cell.paragraphs[v_i] if v_i < len(t_cell.paragraphs) else None
                if t_p is not None and t_p.text.strip() and t_p.text.strip() != o_cell.paragraphs[v_i].text.strip():
                    ins = _insert_paragraph_after(o_cell.paragraphs[v_i])
                    _copy_runs_with_basic_style(t_p, ins, override_color=_KO_ACCENT_COLOR)


def _docx_alternating_vi_ko(
    original_bytes: bytes,
    translated_ko_bytes: bytes,
    original_has_en: Optional[bool] = None,
) -> DocxDocument:
    """Alternating layout for vi-ko: if original is VI+EN, keeps VI and overwrites EN with VI->KO."""
    orig_doc = docx.Document(io.BytesIO(original_bytes))
    trans_doc = docx.Document(io.BytesIO(translated_ko_bytes))

    para_langs = _classify_paragraphs_local(orig_doc.paragraphs)
    if original_has_en is None:
        has_para_en = "en" in para_langs
        has_tbl_en = any(_table_column_languages(t)[1] is not None for t in orig_doc.tables)
        if not has_tbl_en:
            for tbl in orig_doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        if "en" in _classify_paragraphs_local(cell.paragraphs):
                            has_tbl_en = True
                            break
                    if has_tbl_en:
                        break
        original_has_en = has_para_en or has_tbl_en

    if not original_has_en:
        return _docx_alternating(original_bytes, translated_ko_bytes)

    blocks = pair_paragraph_blocks(para_langs)
    for vi_idxs, en_idxs in reversed(blocks):
        if vi_idxs and en_idxs:
            leading_en = en_idxs[0] < vi_idxs[0]
            if len(vi_idxs) == 1 and not leading_en:
                v_i = vi_idxs[0]
                t_p = trans_doc.paragraphs[v_i] if v_i < len(trans_doc.paragraphs) else None
                for extra_e in reversed(en_idxs[1:]):
                    _remove_paragraph(orig_doc.paragraphs[extra_e])
                if t_p is not None and t_p.text.strip():
                    e_p = orig_doc.paragraphs[en_idxs[0]]
                    e_p.paragraph_format.left_indent = Cm(0.5)
                    e_p.paragraph_format.space_after = Pt(6)
                    _replace_paragraph_with_source(e_p, t_p, override_color=_KO_ACCENT_COLOR)
            else:
                for e_i in reversed(en_idxs):
                    _remove_paragraph(orig_doc.paragraphs[e_i])
                for v_i in reversed(vi_idxs):
                    t_p = trans_doc.paragraphs[v_i] if v_i < len(trans_doc.paragraphs) else None
                    if t_p is not None and t_p.text.strip():
                        inserted = _insert_paragraph_after(orig_doc.paragraphs[v_i])
                        inserted.paragraph_format.left_indent = Cm(0.5)
                        inserted.paragraph_format.space_after = Pt(6)
                        _copy_runs_with_basic_style(t_p, inserted, override_color=_KO_ACCENT_COLOR)
        elif vi_idxs:
            for v_i in reversed(vi_idxs):
                orig_p = orig_doc.paragraphs[v_i]
                t_p = trans_doc.paragraphs[v_i] if v_i < len(trans_doc.paragraphs) else None
                if t_p is not None and t_p.text.strip() and t_p.text.strip() != orig_p.text.strip():
                    inserted = _insert_paragraph_after(orig_p)
                    inserted.paragraph_format.left_indent = Cm(0.5)
                    inserted.paragraph_format.space_after = Pt(6)
                    _copy_runs_with_basic_style(t_p, inserted, override_color=_KO_ACCENT_COLOR)
        elif en_idxs:
            # Standalone EN paragraph in a VI+EN doc without adjacent VI: overwrite with trans if translated, else keep clean
            for e_i in reversed(en_idxs):
                orig_e = orig_doc.paragraphs[e_i]
                t_p = trans_doc.paragraphs[e_i] if e_i < len(trans_doc.paragraphs) else None
                if t_p is not None and t_p.text.strip() and t_p.text.strip() != orig_e.text.strip():
                    _replace_paragraph_with_source(orig_e, t_p, override_color=_KO_ACCENT_COLOR)

    for orig_tbl, trans_tbl in zip(orig_doc.tables, trans_doc.tables):
        vi_col, en_col = _table_column_languages(orig_tbl)
        if vi_col is not None and en_col is not None:
            seen_en_tc: set[int] = set()
            for o_row, t_row in zip(orig_tbl.rows, trans_tbl.rows):
                if vi_col >= len(o_row.cells) or en_col >= len(o_row.cells):
                    continue
                v_cell = o_row.cells[vi_col]
                e_cell = o_row.cells[en_col]
                if id(v_cell._tc) == id(e_cell._tc):
                    # Merged full-width cell: process stacked paragraphs inside it
                    if id(v_cell._tc) not in seen_en_tc:
                        seen_en_tc.add(id(v_cell._tc))
                        t_cell = t_row.cells[vi_col] if vi_col < len(t_row.cells) else t_row.cells[0]
                        _overwrite_cell_en_paragraphs_with_ko(v_cell, t_cell)
                    continue
                if id(e_cell._tc) in seen_en_tc:
                    continue
                seen_en_tc.add(id(e_cell._tc))
                donor_cell = t_row.cells[vi_col] if vi_col < len(t_row.cells) else t_row.cells[0]
                if donor_cell.text.strip():
                    _copy_cell_paragraphs(donor_cell, e_cell, override_color=_KO_ACCENT_COLOR)
        else:
            seen_tc: set[int] = set()
            for o_row, t_row in zip(orig_tbl.rows, trans_tbl.rows):
                for o_cell, t_cell in zip(o_row.cells, t_row.cells):
                    if id(o_cell._tc) in seen_tc:
                        continue
                    seen_tc.add(id(o_cell._tc))
                    _overwrite_cell_en_paragraphs_with_ko(o_cell, t_cell)

    return orig_doc


def _docx_replace_column(
    original_bytes: bytes,
    translated_bytes: bytes,
    replace_col_index: int = 1,
    source_lang: str = "vi",
) -> DocxDocument:
    """Replaces target column (default col 1, e.g. EN -> KO) in bilingual tables while keeping VI col."""
    orig_doc = docx.Document(io.BytesIO(original_bytes))
    trans_doc = docx.Document(io.BytesIO(translated_bytes))

    replaced_any_table = False
    for orig_tbl, trans_tbl in zip(orig_doc.tables, trans_doc.tables):
        vi_col, en_col = _table_column_languages(orig_tbl)
        for o_row, t_row in zip(orig_tbl.rows, trans_tbl.rows):
            if len(o_row.cells) <= replace_col_index or len(t_row.cells) <= replace_col_index:
                continue
            target_o_cell = o_row.cells[replace_col_index]
            if replace_col_index > 0 and id(target_o_cell._tc) == id(o_row.cells[replace_col_index - 1]._tc):
                continue

            # When source_lang='vi' and replace_col_index holds EN while vi_col holds VI,
            # take the VI->KO translation from t_row.cells[vi_col] rather than untranslated EN cell.
            if source_lang == "vi" and vi_col is not None and vi_col != replace_col_index and vi_col < len(t_row.cells):
                donor_cell = t_row.cells[vi_col]
            else:
                donor_cell = t_row.cells[replace_col_index]
                if not donor_cell.text.strip() and len(t_row.cells) > 0:
                    donor_cell = t_row.cells[0]

            _copy_cell_paragraphs(donor_cell, target_o_cell)
            replaced_any_table = True

    if not replaced_any_table and not orig_doc.tables:
        return _docx_side_by_side(original_bytes, translated_bytes, source_lang, "ko")

    return orig_doc


# =============================================================================
# XLSX HELPERS & LAYOUT BUILDERS
# =============================================================================


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
    original_has_en: Optional[bool] = None,
) -> bytes:
    orig_wb = openpyxl.load_workbook(io.BytesIO(original_bytes))
    trans_wb = openpyxl.load_workbook(io.BytesIO(translated_bytes))

    if layout_mode == "replace_column":
        target_col_1based = replace_col_index + 1
        for o_ws, t_ws in zip(orig_wb.worksheets, trans_wb.worksheets):
            col_pair, _ = _xlsx_sheet_structure(o_ws)
            vi_c = col_pair[0] if col_pair else None
            max_r = max(o_ws.max_row or 1, t_ws.max_row or 1)
            for r in range(1, max_r + 1):
                if source_lang == "vi" and vi_c is not None and vi_c != target_col_1based:
                    t_val = t_ws.cell(row=r, column=vi_c).value
                else:
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
            col_pair, row_langs = _xlsx_sheet_structure(o_ws) if (source_lang == "vi" and target_lang == "ko") else (None, [])
            is_vi_ko_with_en = (
                source_lang == "vi"
                and target_lang == "ko"
                and (original_has_en is True or col_pair is not None or ("en" in row_langs and "vi" in row_langs))
            )

            if is_vi_ko_with_en and col_pair is not None:
                vi_c, en_c = col_pair
                for r in range(1, max_r + 1):
                    for c in range(1, max_c + 1):
                        o_cell = o_ws.cell(row=r, column=c)
                        if c == en_c:
                            t_val = t_ws.cell(row=r, column=vi_c).value
                            dst_c = ws.cell(row=r, column=c, value=t_val if t_val is not None else o_cell.value)
                            _copy_cell_style(o_cell, dst_c)
                            dst_c.fill = trans_fill
                        else:
                            dst_c = ws.cell(row=r, column=c, value=o_cell.value)
                            _copy_cell_style(o_cell, dst_c)
                        dst_c.alignment = Alignment(wrap_text=True, vertical="top")
            elif is_vi_ko_with_en and ("en" in row_langs and "vi" in row_langs):
                dst_r = 1
                blocks = pair_paragraph_blocks(row_langs)
                handled_rows: set[int] = set()
                for vi_idxs, en_idxs in blocks:
                    if vi_idxs and en_idxs:
                        for idx_pos, v_i in enumerate(vi_idxs):
                            v_r = v_i + 1
                            handled_rows.add(v_r)
                            for c in range(1, max_c + 1):
                                o_cell = o_ws.cell(row=v_r, column=c)
                                d_orig = ws.cell(row=dst_r, column=c, value=o_cell.value)
                                _copy_cell_style(o_cell, d_orig)
                            dst_r += 1
                            for c in range(1, max_c + 1):
                                t_cell = t_ws.cell(row=v_r, column=c)
                                d_trans = ws.cell(row=dst_r, column=c, value=t_cell.value)
                                _copy_cell_style(t_cell, d_trans)
                                d_trans.fill = trans_fill
                                d_trans.alignment = Alignment(wrap_text=True, vertical="top")
                            dst_r += 1
                        for e_i in en_idxs:
                            handled_rows.add(e_i + 1)
                # Keep any remaining neutral/unhandled rows in order if sheet had mixed rows
                if not handled_rows:
                    for r in range(1, max_r + 1):
                        for c in range(1, max_c + 1):
                            o_cell = o_ws.cell(row=r, column=c)
                            _copy_cell_style(o_cell, ws.cell(row=r, column=c, value=o_cell.value))
            else:
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


# =============================================================================
# TRILINGUAL (VI-ONLY ORIGINAL) BUILDERS
# =============================================================================


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
            _copy_runs_with_basic_style(ko_p, inserted_ko, override_color=_KO_ACCENT_COLOR)
        if e_txt and e_txt != o_txt:
            inserted_en = _insert_paragraph_after(orig_p)
            inserted_en.paragraph_format.left_indent = Cm(0.5)
            inserted_en.paragraph_format.space_after = Pt(2)
            _copy_runs_with_basic_style(en_p, inserted_en, override_color=_EN_ACCENT_COLOR)

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
                    r_en.font.color.rgb = _EN_ACCENT_COLOR
                if k_val and k_val != o_val:
                    p_ko = o_cell.add_paragraph()
                    r_ko = p_ko.add_run(k_val)
                    r_ko.font.size = Pt(9.5)
                    r_ko.font.color.rgb = _KO_ACCENT_COLOR

    return orig_doc


# =============================================================================
# TRILINGUAL (VI + EN BILINGUAL ORIGINAL -> VI / EN(original) / KO) BUILDERS
# =============================================================================


def _docx_trilingual_from_bilingual_in_place(
    original_bytes: bytes,
    translated_ko_bytes: bytes,
    translated_en_bytes: Optional[bytes] = None,
    layout_mode: str = "alternating",
    replace_col_index: int = 1,
) -> DocxDocument:
    """Keeps VI and original EN untouched, and adds KO (VI->KO) for alternating & replace_column."""
    orig_doc = docx.Document(io.BytesIO(original_bytes))
    ko_doc = docx.Document(io.BytesIO(translated_ko_bytes))
    en_trans_doc = docx.Document(io.BytesIO(translated_en_bytes)) if translated_en_bytes else None

    para_langs = _classify_paragraphs_local(orig_doc.paragraphs)
    blocks = pair_paragraph_blocks(para_langs)

    # If replace_column on a pure-paragraph doc with no tables and no EN paras, fallback to 3-col side-by-side
    if layout_mode == "replace_column" and not orig_doc.tables and not any(e for _, e in blocks):
        if en_trans_doc is not None:
            return _docx_trilingual_side_by_side(original_bytes, translated_en_bytes, translated_ko_bytes, "vi")

    for vi_idxs, en_idxs in reversed(blocks):
        if vi_idxs and en_idxs:
            leading_en = en_idxs[0] < vi_idxs[0]
            if len(vi_idxs) == len(en_idxs) and not leading_en:
                for v_i, e_i in reversed(list(zip(vi_idxs, en_idxs))):
                    e_p = orig_doc.paragraphs[e_i]
                    for r in e_p.runs:
                        if not r.font.color or not r.font.color.rgb:
                            r.font.color.rgb = _EN_ACCENT_COLOR
                    e_p.paragraph_format.space_after = Pt(2)
                    t_ko = ko_doc.paragraphs[v_i] if v_i < len(ko_doc.paragraphs) else None
                    if t_ko is not None and t_ko.text.strip():
                        ins_ko = _insert_paragraph_after(e_p)
                        ins_ko.paragraph_format.left_indent = Cm(0.5)
                        ins_ko.paragraph_format.space_after = Pt(6)
                        _copy_runs_with_basic_style(t_ko, ins_ko, override_color=_KO_ACCENT_COLOR)
            elif not leading_en:
                # Keep all original EN paragraphs untouched; append KO after the last EN paragraph
                last_e = orig_doc.paragraphs[en_idxs[-1]]
                last_e.paragraph_format.space_after = Pt(2)
                ko_joined = "\n".join(
                    ko_doc.paragraphs[v_i].text.strip()
                    for v_i in vi_idxs
                    if v_i < len(ko_doc.paragraphs) and ko_doc.paragraphs[v_i].text.strip()
                )
                if ko_joined:
                    ins_ko = _insert_paragraph_after(last_e, ko_joined)
                    ins_ko.paragraph_format.left_indent = Cm(0.5)
                    ins_ko.paragraph_format.space_after = Pt(6)
                    for r in ins_ko.runs:
                        r.font.color.rgb = _KO_ACCENT_COLOR
            else:
                # Leading EN before VI: normalize order to VI -> EN (original) -> KO
                orig_en_joined = "\n".join(orig_doc.paragraphs[e].text.strip() for e in en_idxs if orig_doc.paragraphs[e].text.strip())
                for e_i in reversed(en_idxs):
                    _remove_paragraph(orig_doc.paragraphs[e_i])
                last_v = orig_doc.paragraphs[vi_idxs[-1]]
                ko_joined = "\n".join(
                    ko_doc.paragraphs[v_i].text.strip()
                    for v_i in vi_idxs
                    if v_i < len(ko_doc.paragraphs) and ko_doc.paragraphs[v_i].text.strip()
                )
                if ko_joined:
                    ins_ko = _insert_paragraph_after(last_v, ko_joined)
                    ins_ko.paragraph_format.left_indent = Cm(0.5)
                    ins_ko.paragraph_format.space_after = Pt(6)
                    for r in ins_ko.runs:
                        r.font.color.rgb = _KO_ACCENT_COLOR
                if orig_en_joined:
                    ins_en = _insert_paragraph_after(last_v, orig_en_joined)
                    ins_en.paragraph_format.left_indent = Cm(0.5)
                    ins_en.paragraph_format.space_after = Pt(2)
                    for r in ins_en.runs:
                        r.font.color.rgb = _EN_ACCENT_COLOR
        elif vi_idxs:
            for v_i in reversed(vi_idxs):
                orig_p = orig_doc.paragraphs[v_i]
                o_txt = orig_p.text.strip()
                t_ko = ko_doc.paragraphs[v_i] if v_i < len(ko_doc.paragraphs) else None
                t_en = en_trans_doc.paragraphs[v_i] if (en_trans_doc and v_i < len(en_trans_doc.paragraphs)) else None
                if t_ko is not None and t_ko.text.strip() and t_ko.text.strip() != o_txt:
                    ins_ko = _insert_paragraph_after(orig_p)
                    ins_ko.paragraph_format.left_indent = Cm(0.5)
                    ins_ko.paragraph_format.space_after = Pt(6)
                    _copy_runs_with_basic_style(t_ko, ins_ko, override_color=_KO_ACCENT_COLOR)
                if t_en is not None and t_en.text.strip() and t_en.text.strip() != o_txt:
                    ins_en = _insert_paragraph_after(orig_p)
                    ins_en.paragraph_format.left_indent = Cm(0.5)
                    ins_en.paragraph_format.space_after = Pt(2)
                    _copy_runs_with_basic_style(t_en, ins_en, override_color=_EN_ACCENT_COLOR)

    # Tables
    for t_idx, (orig_tbl, ko_tbl) in enumerate(zip(orig_doc.tables, ko_doc.tables)):
        vi_col, _en_col = _table_column_languages(orig_tbl)
        en_trans_tbl = en_trans_doc.tables[t_idx] if (en_trans_doc and t_idx < len(en_trans_doc.tables)) else None
        if vi_col is not None and _en_col is not None:
            ko_cells_by_row: list[Optional[docx.table._Cell]] = []
            for k_row in ko_tbl.rows:
                ko_cell = k_row.cells[vi_col] if vi_col < len(k_row.cells) else (k_row.cells[0] if k_row.cells else None)
                ko_cells_by_row.append(ko_cell)
            _append_column_to_table(orig_tbl, ko_cells_by_row)
        else:
            seen_tc: set[int] = set()
            for r_idx, (o_row, k_row) in enumerate(zip(orig_tbl.rows, ko_tbl.rows)):
                for c_idx, (o_cell, k_cell) in enumerate(zip(o_row.cells, k_row.cells)):
                    if id(o_cell._tc) in seen_tc:
                        continue
                    seen_tc.add(id(o_cell._tc))
                    c_langs = _classify_paragraphs_local(o_cell.paragraphs)
                    c_blocks = pair_paragraph_blocks(c_langs)
                    has_cell_en = any(e for _, e in c_blocks)
                    if has_cell_en:
                        for vi_idxs, en_idxs in reversed(c_blocks):
                            if vi_idxs and en_idxs:
                                if len(vi_idxs) == len(en_idxs) and en_idxs[0] > vi_idxs[0]:
                                    for v_i, e_i in reversed(list(zip(vi_idxs, en_idxs))):
                                        e_p = o_cell.paragraphs[e_i]
                                        t_ko = k_cell.paragraphs[v_i] if v_i < len(k_cell.paragraphs) else None
                                        if t_ko is not None and t_ko.text.strip():
                                            ins_ko = _insert_paragraph_after(e_p)
                                            _copy_runs_with_basic_style(t_ko, ins_ko, override_color=_KO_ACCENT_COLOR)
                                elif en_idxs[0] > vi_idxs[0]:
                                    last_e = o_cell.paragraphs[en_idxs[-1]]
                                    ko_joined = "\n".join(
                                        k_cell.paragraphs[v_i].text.strip()
                                        for v_i in vi_idxs
                                        if v_i < len(k_cell.paragraphs) and k_cell.paragraphs[v_i].text.strip()
                                    )
                                    if ko_joined:
                                        ins_ko = _insert_paragraph_after(last_e, ko_joined)
                                        for r in ins_ko.runs:
                                            r.font.color.rgb = _KO_ACCENT_COLOR
                                else:
                                    orig_en_joined = "\n".join(o_cell.paragraphs[e].text.strip() for e in en_idxs if o_cell.paragraphs[e].text.strip())
                                    for e_i in reversed(en_idxs):
                                        _remove_paragraph(o_cell.paragraphs[e_i])
                                    last_v = o_cell.paragraphs[vi_idxs[-1]]
                                    ko_joined = "\n".join(
                                        k_cell.paragraphs[v_i].text.strip()
                                        for v_i in vi_idxs
                                        if v_i < len(k_cell.paragraphs) and k_cell.paragraphs[v_i].text.strip()
                                    )
                                    if ko_joined:
                                        ins_ko = _insert_paragraph_after(last_v, ko_joined)
                                        for r in ins_ko.runs:
                                            r.font.color.rgb = _KO_ACCENT_COLOR
                                    if orig_en_joined:
                                        ins_en = _insert_paragraph_after(last_v, orig_en_joined)
                                        for r in ins_en.runs:
                                            r.font.color.rgb = _EN_ACCENT_COLOR
                    else:
                        o_val = o_cell.text.strip()
                        k_val = k_cell.text.strip()
                        if en_trans_tbl and r_idx < len(en_trans_tbl.rows) and c_idx < len(en_trans_tbl.rows[r_idx].cells):
                            e_val = en_trans_tbl.rows[r_idx].cells[c_idx].text.strip()
                            if e_val and e_val != o_val:
                                p_en = o_cell.add_paragraph()
                                r_en = p_en.add_run(e_val)
                                r_en.font.size = Pt(9.5)
                                r_en.font.color.rgb = _EN_ACCENT_COLOR
                        if k_val and k_val != o_val:
                            p_ko = o_cell.add_paragraph()
                            r_ko = p_ko.add_run(k_val)
                            r_ko.font.size = Pt(9.5)
                            r_ko.font.color.rgb = _KO_ACCENT_COLOR

    return orig_doc


def _docx_trilingual_from_bilingual_side_by_side(
    original_bytes: bytes,
    translated_ko_bytes: bytes,
    translated_en_bytes: Optional[bytes] = None,
    source_lang: str = "vi",
) -> DocxDocument:
    """Builds a 3-column comparison table (VI | EN(original) | KO) from a bilingual VI+EN doc."""
    orig_doc = docx.Document(io.BytesIO(original_bytes))
    ko_doc = docx.Document(io.BytesIO(translated_ko_bytes))
    en_trans_doc = docx.Document(io.BytesIO(translated_en_bytes)) if translated_en_bytes else None

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

    # Pre-compute paragraph block mapping
    para_to_idx = {p._p: idx for idx, p in enumerate(orig_doc.paragraphs)}
    tbl_to_idx = {t._tbl: idx for idx, t in enumerate(orig_doc.tables)}
    para_langs = _classify_paragraphs_local(orig_doc.paragraphs)
    p_blocks = pair_paragraph_blocks(para_langs)

    # Map starting paragraph index -> list of (vi_text, en_text, ko_text) rows
    block_by_first_para: dict[int, list[tuple[str, str, str]]] = {}
    consumed_paras: set[int] = set()
    for b_idx, (vi_idxs, en_idxs) in enumerate(p_blocks):
        all_idxs = sorted(set(vi_idxs + en_idxs))
        if not all_idxs:
            continue
        first_idx = all_idxs[0]
        consumed_paras.update(all_idxs)
        rows_list: list[tuple[str, str, str]] = []
        if vi_idxs and en_idxs:
            if len(vi_idxs) == len(en_idxs):
                for v_i, e_i in zip(vi_idxs, en_idxs):
                    v_txt = orig_doc.paragraphs[v_i].text.strip()
                    e_txt = orig_doc.paragraphs[e_i].text.strip()
                    k_txt = ko_doc.paragraphs[v_i].text.strip() if v_i < len(ko_doc.paragraphs) else ""
                    rows_list.append((v_txt, e_txt, k_txt))
            else:
                v_txt = "\n".join(orig_doc.paragraphs[i].text.strip() for i in vi_idxs if orig_doc.paragraphs[i].text.strip())
                e_txt = "\n".join(orig_doc.paragraphs[i].text.strip() for i in en_idxs if orig_doc.paragraphs[i].text.strip())
                k_txt = "\n".join(
                    ko_doc.paragraphs[i].text.strip()
                    for i in vi_idxs
                    if i < len(ko_doc.paragraphs) and ko_doc.paragraphs[i].text.strip()
                )
                rows_list.append((v_txt, e_txt, k_txt))
        elif vi_idxs:
            for v_i in vi_idxs:
                v_txt = orig_doc.paragraphs[v_i].text.strip()
                e_txt = (
                    en_trans_doc.paragraphs[v_i].text.strip()
                    if (en_trans_doc and v_i < len(en_trans_doc.paragraphs))
                    else ""
                )
                k_txt = ko_doc.paragraphs[v_i].text.strip() if v_i < len(ko_doc.paragraphs) else ""
                rows_list.append((v_txt, e_txt, k_txt))
        block_by_first_para[first_idx] = rows_list

    for kind, obj in _iter_body_blocks(orig_doc):
        if kind == "p":
            p_idx = para_to_idx.get(obj._p)
            if p_idx is None:
                continue
            if p_idx in block_by_first_para:
                for v_txt, e_txt, k_txt in block_by_first_para[p_idx]:
                    if not v_txt and not e_txt and not k_txt:
                        continue
                    row = table.add_row()
                    _prevent_row_split(row)
                    row.cells[0].paragraphs[0].add_run(v_txt)
                    row.cells[1].paragraphs[0].add_run(e_txt)
                    row.cells[2].paragraphs[0].add_run(k_txt)
            elif p_idx not in consumed_paras:
                txt = obj.text.strip()
                if txt:
                    k_txt = ko_doc.paragraphs[p_idx].text.strip() if p_idx < len(ko_doc.paragraphs) else txt
                    row = table.add_row()
                    _prevent_row_split(row)
                    row.cells[0].paragraphs[0].add_run(txt)
                    row.cells[1].paragraphs[0].add_run(txt)
                    row.cells[2].paragraphs[0].add_run(k_txt)
        elif kind == "tbl":
            t_idx = tbl_to_idx.get(obj._tbl)
            if t_idx is None:
                continue
            ko_tbl = ko_doc.tables[t_idx] if t_idx < len(ko_doc.tables) else None
            vi_col, en_col = _table_column_languages(obj)
            if vi_col is not None and en_col is not None:
                seen_pairs: set[tuple[int, int]] = set()
                for r_idx, row_o in enumerate(obj.rows):
                    if vi_col >= len(row_o.cells) or en_col >= len(row_o.cells):
                        continue
                    v_cell = row_o.cells[vi_col]
                    e_cell = row_o.cells[en_col]
                    if (id(v_cell._tc), id(e_cell._tc)) in seen_pairs:
                        continue
                    seen_pairs.add((id(v_cell._tc), id(e_cell._tc)))
                    v_txt = v_cell.text.strip()
                    e_txt = e_cell.text.strip()
                    k_txt = ""
                    if ko_tbl and r_idx < len(ko_tbl.rows) and vi_col < len(ko_tbl.rows[r_idx].cells):
                        k_txt = ko_tbl.rows[r_idx].cells[vi_col].text.strip()
                    if not v_txt and not e_txt and not k_txt:
                        continue
                    new_row = table.add_row()
                    _prevent_row_split(new_row)
                    for c_i, val in enumerate((v_txt, e_txt, k_txt)):
                        _set_cell_shading(new_row.cells[c_i], "F8FAFC")
                        r = new_row.cells[c_i].paragraphs[0].add_run(val)
                        r.font.size = Pt(9)
            else:
                o_lines = _flatten_table_text(obj)
                k_lines = _flatten_table_text(ko_tbl) if ko_tbl is not None else []
                for o_line, k_line in zip_longest(o_lines, k_lines, fillvalue=""):
                    if not o_line and not k_line:
                        continue
                    new_row = table.add_row()
                    _prevent_row_split(new_row)
                    for c_i, val in enumerate((o_line, o_line, k_line)):
                        _set_cell_shading(new_row.cells[c_i], "F8FAFC")
                        r = new_row.cells[c_i].paragraphs[0].add_run(val)
                        r.font.size = Pt(9)

    return out_doc


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


def _build_trilingual_from_bilingual_xlsx(
    original_bytes: bytes,
    translated_ko_bytes: bytes,
    translated_en_bytes: Optional[bytes] = None,
    layout_mode: str = "alternating",
    source_lang: str = "vi",
    replace_col_index: int = 1,
) -> bytes:
    """Assembles VI / EN(original) / KO Excel workbook from a bilingual VI+EN original."""
    orig_wb = openpyxl.load_workbook(io.BytesIO(original_bytes))
    ko_wb = openpyxl.load_workbook(io.BytesIO(translated_ko_bytes))
    en_wb = openpyxl.load_workbook(io.BytesIO(translated_en_bytes)) if translated_en_bytes else None

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    en_fill = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
    ko_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    hdr_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    hdr_font = Font(bold=True, color="FFFFFF")

    for s_idx, (o_ws, k_ws) in enumerate(zip(orig_wb.worksheets, ko_wb.worksheets)):
        ws = out_wb.create_sheet(title=o_ws.title[:31])
        e_ws = en_wb.worksheets[s_idx] if (en_wb and s_idx < len(en_wb.worksheets)) else None
        col_pair, row_langs = _xlsx_sheet_structure(o_ws)
        max_r = max(o_ws.max_row or 1, k_ws.max_row or 1)
        max_c = max(o_ws.max_column or 1, k_ws.max_column or 1)

        if col_pair is not None:
            vi_c, en_c = col_pair
            if layout_mode == "side_by_side":
                for col_idx, code in enumerate((source_lang.upper(), "EN", "KO"), start=1):
                    hc = ws.cell(row=1, column=col_idx, value=code)
                    hc.fill, hc.font = hdr_fill, hdr_font
                    ws.column_dimensions[get_column_letter(col_idx)].width = 28
                for r in range(1, max_r + 1):
                    dst_r = r + 1
                    v_cell = o_ws.cell(row=r, column=vi_c)
                    e_cell = o_ws.cell(row=r, column=en_c)
                    k_cell = k_ws.cell(row=r, column=vi_c)
                    d1 = ws.cell(row=dst_r, column=1, value=v_cell.value)
                    d2 = ws.cell(row=dst_r, column=2, value=e_cell.value)
                    d3 = ws.cell(row=dst_r, column=3, value=k_cell.value)
                    _copy_cell_style(v_cell, d1)
                    _copy_cell_style(e_cell, d2)
                    _copy_cell_style(k_cell, d3)
                    for d in (d1, d2, d3):
                        d.alignment = Alignment(wrap_text=True, vertical="top")
            else:
                # Keep original VI & EN columns untouched, insert KO column right after EN
                ko_insert_c = en_c + 1
                for r in range(1, max_r + 1):
                    for c in range(1, max_c + 1):
                        dst_c = c if c < ko_insert_c else c + 1
                        o_cell = o_ws.cell(row=r, column=c)
                        d_cell = ws.cell(row=r, column=dst_c, value=o_cell.value)
                        _copy_cell_style(o_cell, d_cell)
                        d_cell.alignment = Alignment(wrap_text=True, vertical="top")
                    k_cell = k_ws.cell(row=r, column=vi_c)
                    d_ko = ws.cell(row=r, column=ko_insert_c, value=k_cell.value)
                    _copy_cell_style(k_cell, d_ko)
                    d_ko.fill = ko_fill
                    d_ko.alignment = Alignment(wrap_text=True, vertical="top")
        elif "en" in row_langs and "vi" in row_langs:
            blocks = pair_paragraph_blocks(row_langs)
            if layout_mode == "side_by_side":
                for c in range(1, max_c + 1):
                    base_col = (c - 1) * 3 + 1
                    for offset, lang_code in enumerate((source_lang.upper(), "EN", "KO")):
                        hc = ws.cell(row=1, column=base_col + offset, value=f"Col {get_column_letter(c)} ({lang_code})")
                        hc.fill, hc.font = hdr_fill, hdr_font
                        ws.column_dimensions[get_column_letter(base_col + offset)].width = 24
                dst_r = 2
                for vi_idxs, en_idxs in blocks:
                    for sub_idx, v_i in enumerate(vi_idxs):
                        v_r = v_i + 1
                        e_r = en_idxs[sub_idx] + 1 if sub_idx < len(en_idxs) else None
                        for c in range(1, max_c + 1):
                            base_col = (c - 1) * 3 + 1
                            v_val = o_ws.cell(row=v_r, column=c).value
                            e_val = o_ws.cell(row=e_r, column=c).value if e_r else (e_ws.cell(row=v_r, column=c).value if e_ws else "")
                            k_val = k_ws.cell(row=v_r, column=c).value
                            ws.cell(row=dst_r, column=base_col, value=v_val)
                            ws.cell(row=dst_r, column=base_col + 1, value=e_val)
                            ws.cell(row=dst_r, column=base_col + 2, value=k_val)
                        dst_r += 1
            else:
                dst_r = 1
                for vi_idxs, en_idxs in blocks:
                    for sub_idx, v_i in enumerate(vi_idxs):
                        v_r = v_i + 1
                        e_r = en_idxs[sub_idx] + 1 if sub_idx < len(en_idxs) else None
                        for c in range(1, max_c + 1):
                            o_cell = o_ws.cell(row=v_r, column=c)
                            d_v = ws.cell(row=dst_r, column=c, value=o_cell.value)
                            _copy_cell_style(o_cell, d_v)
                        dst_r += 1
                        if e_r is not None or e_ws is not None:
                            for c in range(1, max_c + 1):
                                src_c = o_ws.cell(row=e_r, column=c) if e_r else (e_ws.cell(row=v_r, column=c) if e_ws else None)
                                val = src_c.value if src_c else ""
                                d_e = ws.cell(row=dst_r, column=c, value=val)
                                if src_c:
                                    _copy_cell_style(src_c, d_e)
                                d_e.fill = en_fill
                            dst_r += 1
                        for c in range(1, max_c + 1):
                            k_cell = k_ws.cell(row=v_r, column=c)
                            d_k = ws.cell(row=dst_r, column=c, value=k_cell.value)
                            _copy_cell_style(k_cell, d_k)
                            d_k.fill = ko_fill
                        dst_r += 1
        else:
            # Sheet without EN: fall back to VI / EN / KO
            for r in range(1, max_r + 1):
                for c in range(1, max_c + 1):
                    ws.cell(row=r, column=c, value=o_ws.cell(row=r, column=c).value)

    out_buf = io.BytesIO()
    out_wb.save(out_buf)
    return out_buf.getvalue()
