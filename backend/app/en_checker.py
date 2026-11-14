"""Heuristic VI / EN / KO Language Detection Service (zero LLM/API calls).

Used for:
1. Zero-cost per-paragraph / per-cell / per-column language detection using
   character and stopword heuristics.
2. Detecting whether a .docx or .xlsx contract already contains both VI and EN
   so that:
   - `vi-ko` can keep VI and overwrite EN with VI->KO.
   - `vi-en-kr` can keep original VI and EN untouched and append VI->KO.
"""

from __future__ import annotations

import io
import re
from typing import Literal, Optional

import docx
import openpyxl
from pydantic import BaseModel

LangCode = Literal["vi", "en", "ko", "neutral", "unknown"]

_HANGUL_RE = re.compile(r"[\uAC00-\uD7AF\u1100-\u11FF\u3130-\u318F]")
_VI_DIACRITIC_RE = re.compile(
    r"[àáảãạăằắẳẵặâầấẩẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵ"
    r"ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬĐÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴ]"
)
_WORD_RE = re.compile(r"[A-Za-zÀ-ỹ]+")
_ASCII_WORD_RE = re.compile(r"^[A-Za-z]+$")

_EN_STOPWORDS = frozenset({
    "the", "and", "of", "to", "in", "for", "with", "by", "on", "at", "from", "as",
    "or", "is", "are", "was", "were", "be", "been", "shall", "will", "may", "must",
    "not", "no", "this", "that", "these", "those", "which", "who", "whom", "whose",
    "where", "when", "if", "unless", "until", "upon", "under", "between", "among",
    "against", "within", "without", "article", "section", "clause", "chapter",
    "party", "parties", "agreement", "contract", "hereby", "herein", "hereof",
    "thereto", "whereas", "therefore", "effective", "date", "term", "terms",
    "payment", "price", "amount", "days", "month", "months", "year", "years",
    "force", "majeure", "confidential", "confidentiality", "liability", "damages",
    "penalty", "breach", "termination", "dispute", "arbitration", "governing",
    "law", "signature", "representative", "director", "company", "limited",
    "address", "tax", "number", "appendix", "schedule", "annex", "each", "any",
    "all", "such", "other", "both", "either", "subject", "scope", "work",
    "services", "service", "obligations", "rights", "value", "total", "percent",
})

_VI_STOPWORDS = frozenset({
    "điều", "khoản", "chương", "mục", "phần", "hợp", "đồng", "bên", "của", "và",
    "hoặc", "là", "có", "không", "được", "phải", "này", "đó", "các", "những",
    "cho", "về", "từ", "đến", "với", "theo", "tại", "trong", "ngoài", "trên",
    "dưới", "trước", "sau", "khi", "nếu", "trừ", "trường", "hợp", "thanh", "toán",
    "giá", "trị", "thời", "hạn", "ngày", "tháng", "năm", "vi", "phạm", "phạt",
    "bồi", "thường", "trách", "nhiệm", "bất", "khả", "kháng", "bảo", "mật",
    "chấm", "dứt", "tranh", "chấp", "luật", "đại", "diện", "giám", "đốc",
    "công", "ty", "hữu", "địa", "chỉ", "mã", "số", "thuế", "phụ", "lục",
    "ký", "tên", "cam", "kết", "thỏa", "thuận", "cung", "cấp", "dịch", "vụ",
    # Common unaccented legal Vietnamese tokens in case text lacks diacritics
    "dieu", "khoan", "hop", "dong", "ben", "cua", "khong", "duoc", "phai",
    "thanh", "toan", "ngay", "thang", "nam", "cong", "ty",
})


def detect_segment_lang_heuristic(text: str) -> LangCode:
    """Zero-cost character + stopword heuristic for a single segment."""
    s = (text or "").strip()
    if not s:
        return "neutral"

    hangul_count = len(_HANGUL_RE.findall(s))
    if hangul_count >= 1:
        return "ko"

    words = _WORD_RE.findall(s)
    if not words:
        return "neutral"

    vi_diacritics = len(_VI_DIACRITIC_RE.findall(s))
    lower_words = [w.lower() for w in words]
    en_hits = sum(1 for w in lower_words if w in _EN_STOPWORDS)
    vi_hits = sum(1 for w in lower_words if w in _VI_STOPWORDS)
    ascii_words = [w for w in words if _ASCII_WORD_RE.match(w)]

    # Clear Vietnamese: diacritics present and not dominated by English stopwords
    if vi_diacritics >= 2 and en_hits <= 1:
        return "vi"
    if vi_diacritics >= 1 and vi_hits >= 1 and en_hits == 0:
        return "vi"
    # Mixed or English sentence containing a Vietnamese proper name (e.g. "Mr. Nguyễn Văn A shall...")
    if vi_diacritics >= 1 and en_hits >= 2:
        if en_hits > vi_hits * 2 and len(ascii_words) > vi_diacritics:
            return "en"
        return "unknown"

    # No Vietnamese diacritics
    if vi_diacritics == 0:
        if vi_hits >= 2 and en_hits == 0:
            return "vi"
        if en_hits >= 1 and len(ascii_words) >= 1:
            return "en"
        if len(ascii_words) >= 3 and vi_hits == 0:
            return "en"
        if len(ascii_words) == 1 and len(ascii_words[0]) <= 3 and len(words) == 1:
            return "neutral"
        return "unknown"

    return "unknown"


def resolve_heuristic_fallback(text: str, lang: LangCode) -> Literal["vi", "en", "ko", "neutral"]:
    """Resolves 'unknown' purely locally without any LLM call."""
    if lang != "unknown":
        return lang
    s = (text or "").strip()
    if not s:
        return "neutral"
    if _VI_DIACRITIC_RE.search(s):
        return "vi"
    words = _WORD_RE.findall(s)
    ascii_words = [w for w in words if _ASCII_WORD_RE.match(w)]
    if len(ascii_words) >= 2:
        return "en"
    return "neutral"


class DocumentLangSummary(BaseModel):
    has_vi_and_en: bool = False
    vi_count: int = 0
    en_count: int = 0
    paired_vi_en_count: int = 0
    unmatched_vi_count: int = 0
    bilingual_tables_count: int = 0


def _table_column_languages(
    tbl: docx.table.Table,
    resolved_cell_langs: Optional[dict[tuple[int, int], str]] = None,
) -> tuple[Optional[int], Optional[int]]:
    """Detects (vi_col_idx, en_col_idx) if a DOCX table is a bilingual column table."""
    if not tbl.rows:
        return None, None
    max_cols = max((len(r.cells) for r in tbl.rows), default=0)
    if max_cols < 2:
        return None, None

    col_counts: list[dict[str, int]] = [{"vi": 0, "en": 0} for _ in range(max_cols)]
    for r_idx, row in enumerate(tbl.rows):
        seen_tc: set[int] = set()
        for c_idx, cell in enumerate(row.cells):
            tc_id = id(cell._tc)
            if tc_id in seen_tc:
                continue
            seen_tc.add(tc_id)
            txt = cell.text.strip()
            if not txt:
                continue
            if resolved_cell_langs and (r_idx, c_idx) in resolved_cell_langs:
                lang = resolved_cell_langs[(r_idx, c_idx)]
            else:
                lang = resolve_heuristic_fallback(txt, detect_segment_lang_heuristic(txt))
            if lang in ("vi", "en"):
                col_counts[c_idx][lang] += 1

    vi_col: Optional[int] = None
    en_col: Optional[int] = None
    best_vi = 0
    best_en = 0
    for c_idx, counts in enumerate(col_counts):
        if counts["vi"] > counts["en"] and counts["vi"] > best_vi:
            best_vi = counts["vi"]
            vi_col = c_idx
        if counts["en"] > counts["vi"] and counts["en"] > best_en:
            best_en = counts["en"]
            en_col = c_idx

    if vi_col is not None and en_col is not None and vi_col != en_col:
        return vi_col, en_col
    return None, None


def _xlsx_sheet_structure(
    ws: openpyxl.worksheet.worksheet.Worksheet,
) -> tuple[Optional[tuple[int, int]], list[str]]:
    """Returns ((vi_col_1based, en_col_1based) or None, row_langs)."""
    max_r = ws.max_row or 0
    max_c = ws.max_column or 0
    if max_r == 0 or max_c == 0:
        return None, []

    col_counts: list[dict[str, int]] = [{"vi": 0, "en": 0} for _ in range(max_c)]
    row_langs: list[str] = []

    for r in range(1, max_r + 1):
        row_vi = 0
        row_en = 0
        for c in range(1, max_c + 1):
            val = ws.cell(row=r, column=c).value
            if val is None:
                continue
            s = str(val).strip()
            if not s:
                continue
            lang = resolve_heuristic_fallback(s, detect_segment_lang_heuristic(s))
            if lang == "vi":
                col_counts[c - 1]["vi"] += 1
                row_vi += 1
            elif lang == "en":
                col_counts[c - 1]["en"] += 1
                row_en += 1
        if row_vi > row_en and row_vi > 0:
            row_langs.append("vi")
        elif row_en > row_vi and row_en > 0:
            row_langs.append("en")
        else:
            row_langs.append("neutral")

    vi_col: Optional[int] = None
    en_col: Optional[int] = None
    best_vi = 0
    best_en = 0
    if max_c >= 2:
        for c_idx, counts in enumerate(col_counts):
            if counts["vi"] > counts["en"] and counts["vi"] > best_vi:
                best_vi = counts["vi"]
                vi_col = c_idx + 1
            if counts["en"] > counts["vi"] and counts["en"] > best_en:
                best_en = counts["en"]
                en_col = c_idx + 1

    col_pair = (vi_col, en_col) if (vi_col is not None and en_col is not None and vi_col != en_col) else None
    return col_pair, row_langs


def pair_paragraph_blocks(langs: list[str]) -> list[tuple[list[int], list[int]]]:
    """Pairs consecutive VI paragraph indices with immediately following EN paragraph indices.

    Also handles rare EN-before-VI ordering when an EN block precedes a VI block with no preceding VI.
    Returns list of (vi_indices, en_indices) in document order.
    """
    n = len(langs)
    paired: list[tuple[list[int], list[int]]] = []
    i = 0
    while i < n:
        if langs[i] == "vi":
            vi_idxs = [i]
            j = i + 1
            while j < n and (langs[j] == "vi" or langs[j] == "_empty"):
                if langs[j] == "vi":
                    vi_idxs.append(j)
                j += 1
            en_idxs: list[int] = []
            k = j
            while k < n and (langs[k] == "en" or langs[k] == "_empty"):
                if langs[k] == "en":
                    en_idxs.append(k)
                k += 1
            paired.append((vi_idxs, en_idxs))
            i = k if en_idxs else j
        elif langs[i] == "en":
            en_idxs = [i]
            j = i + 1
            while j < n and (langs[j] == "en" or langs[j] == "_empty"):
                if langs[j] == "en":
                    en_idxs.append(j)
                j += 1
            # Check if VI immediately follows this leading EN block
            vi_idxs = []
            k = j
            if k < n and langs[k] == "vi":
                while k < n and (langs[k] == "vi" or langs[k] == "_empty"):
                    if langs[k] == "vi":
                        vi_idxs.append(k)
                    k += 1
                paired.append((vi_idxs, en_idxs))
                i = k
            else:
                paired.append(([], en_idxs))
                i = j
        else:
            i += 1
    return paired


class EnCheckerService:
    """Provides purely heuristic VI/EN/KO language detection for contract layout routing."""

    def classify_texts(self, texts: list[str]) -> list[str]:
        """Classifies texts locally using character and stopword heuristics."""
        resolved: list[str] = []
        for t in texts:
            if not (t or "").strip():
                resolved.append("_empty")
            else:
                lang = detect_segment_lang_heuristic(t)
                resolved.append(resolve_heuristic_fallback(t, lang))
        return resolved

    async def inspect_document_languages(
        self,
        file_bytes: bytes,
        extension: str,
    ) -> DocumentLangSummary:
        """Detects whether a .docx or .xlsx contract is bilingual (VI + EN) vs VI-only."""
        ext = extension.lower()
        vi_count = 0
        en_count = 0
        paired_count = 0
        unmatched_vi = 0
        bilingual_tables = 0

        if ext == ".docx":
            doc = docx.Document(io.BytesIO(file_bytes))
            para_texts = [p.text for p in doc.paragraphs]
            para_langs = self.classify_texts(para_texts)
            vi_count += sum(1 for l in para_langs if l == "vi")
            en_count += sum(1 for l in para_langs if l == "en")

            for vi_idxs, en_idxs in pair_paragraph_blocks(para_langs):
                if vi_idxs and en_idxs:
                    paired_count += len(vi_idxs)
                elif vi_idxs and not en_idxs:
                    unmatched_vi += len(vi_idxs)

            for tbl in doc.tables:
                vi_col, en_col = _table_column_languages(tbl)
                if vi_col is not None and en_col is not None:
                    bilingual_tables += 1
                    for row in tbl.rows:
                        if vi_col < len(row.cells) and en_col < len(row.cells):
                            v_txt = row.cells[vi_col].text.strip()
                            e_txt = row.cells[en_col].text.strip()
                            v_lang = resolve_heuristic_fallback(v_txt, detect_segment_lang_heuristic(v_txt))
                            e_lang = resolve_heuristic_fallback(e_txt, detect_segment_lang_heuristic(e_txt))
                            if v_lang == "vi":
                                vi_count += 1
                                if e_lang == "en":
                                    en_count += 1
                                    paired_count += 1
                                else:
                                    unmatched_vi += 1
                else:
                    seen_tc: set[int] = set()
                    for row in tbl.rows:
                        for cell in row.cells:
                            if id(cell._tc) in seen_tc:
                                continue
                            seen_tc.add(id(cell._tc))
                            c_langs = self.classify_texts([p.text for p in cell.paragraphs])
                            vi_count += sum(1 for l in c_langs if l == "vi")
                            en_count += sum(1 for l in c_langs if l == "en")
                            for vi_idxs, en_idxs in pair_paragraph_blocks(c_langs):
                                if vi_idxs and en_idxs:
                                    paired_count += len(vi_idxs)
                                elif vi_idxs:
                                    unmatched_vi += len(vi_idxs)

        elif ext == ".xlsx":
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
            try:
                for ws in wb.worksheets:
                    col_pair, row_langs = _xlsx_sheet_structure(ws)
                    if col_pair is not None:
                        bilingual_tables += 1
                        vi_c, en_c = col_pair
                        for r in range(1, (ws.max_row or 0) + 1):
                            v_val = ws.cell(row=r, column=vi_c).value
                            e_val = ws.cell(row=r, column=en_c).value
                            v_s = str(v_val or "").strip()
                            e_s = str(e_val or "").strip()
                            v_l = resolve_heuristic_fallback(v_s, detect_segment_lang_heuristic(v_s)) if v_s else "neutral"
                            e_l = resolve_heuristic_fallback(e_s, detect_segment_lang_heuristic(e_s)) if e_s else "neutral"
                            if v_l == "vi":
                                vi_count += 1
                                if e_l == "en":
                                    en_count += 1
                                    paired_count += 1
                                else:
                                    unmatched_vi += 1
                    else:
                        vi_count += sum(1 for l in row_langs if l == "vi")
                        en_count += sum(1 for l in row_langs if l == "en")
                        for vi_idxs, en_idxs in pair_paragraph_blocks(row_langs):
                            if vi_idxs and en_idxs:
                                paired_count += len(vi_idxs)
                            elif vi_idxs:
                                unmatched_vi += len(vi_idxs)
            finally:
                wb.close()

        has_vi_and_en = (
            vi_count >= 1
            and en_count >= 1
            and (
                bilingual_tables >= 1
                or paired_count >= 1
                or en_count >= 2
            )
            and not (vi_count >= 8 and en_count <= 1 and bilingual_tables == 0)
        )
        return DocumentLangSummary(
            has_vi_and_en=has_vi_and_en,
            vi_count=vi_count,
            en_count=en_count,
            paired_vi_en_count=paired_count,
            unmatched_vi_count=unmatched_vi,
            bilingual_tables_count=bilingual_tables,
        )


_en_checker_singleton: Optional[EnCheckerService] = None


def get_en_checker_service() -> EnCheckerService:
    global _en_checker_singleton
    if _en_checker_singleton is None:
        _en_checker_singleton = EnCheckerService()
    return _en_checker_singleton
