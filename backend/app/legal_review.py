"""AI Legal Review Service (Warning Flags + Marginal Comments Only).

Architectural Rule:
- Inspects bilingual contract pairs (Source vs Translation) for:
  1. Monetary amounts / percentages / unit mismatches (amount_mismatch)
  2. Clause / Article numbering jumps or omissions (clause_numbering)
  3. Dates, deadlines, notice periods, contract term discrepancies (date_term_mismatch)
  4. Penalty, liquidated damages (위약금 / 손해배상), liability cap risks (penalty_liability_risk)
  5. Legal terminology distortion (terminology_warning)
- STRICT INVARIANT: NEVER modifies contract body text!
  Outputs warning flags to UI AND attaches marginal comments (Word Review Comments / Excel Cell Comments)
  for human legal counsel to review and decide.
"""

from __future__ import annotations

import io
import json
import logging
import os
import re
from typing import Any, Literal, Optional

import docx
import openpyxl
from openpyxl.comments import Comment as ExcelComment
from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)

LEGAL_SYSTEM_PROMPT = """Bạn là Luật sư Thẩm định Hợp đồng Song ngữ Quốc tế (Việt - Hàn - Anh).
Nhiệm vụ của bạn là ĐỐI CHIẾU bản gốc và bản dịch hợp đồng để phát hiện sai lệch hoặc rủi ro pháp lý.
TUYỆT ĐỐI KHÔNG tự ý viết lại toàn bộ hợp đồng; chỉ gắn cờ cảnh báo và viết đề xuất dạng Comment bên lề.

Hãy kiểm tra kỹ 5 nhóm lỗi:
1. "amount_mismatch": Sai lệch con số, giá trị hợp đồng, đơn vị tiền tệ (VND/USD/KRW/원), tỷ lệ %, thuế suất.
2. "clause_numbering": Nhảy số Điều/Khoản (Article/Clause/제...조), bỏ sót khoản mục khi dịch.
3. "date_term_mismatch": Sai lệch ngày tháng, thời hạn thanh toán, thời hạn thông báo chấm dứt, hiệu lực.
4. "penalty_liability_risk": Điều khoản phạt vi phạm (위약금), bồi thường thiệt hại (손해배상), giới hạn trách nhiệm, bất khả kháng (불가항력) dịch sai sắc thái nghĩa vụ ("shall/must" vs "may").
5. "terminology_warning": Thuật ngữ pháp lý dịch chưa chuẩn xác.

Trả về DUY NHẤT JSON hợp lệ theo cấu trúc:
{
  "summary": "Tóm tắt ngắn gọn kết quả rà soát pháp lý (1-2 câu tiếng Việt)",
  "issues": [
    {
      "clause_ref": "Điều 5.2 / 제5조 제2항",
      "category": "amount_mismatch | clause_numbering | date_term_mismatch | penalty_liability_risk | terminology_warning",
      "severity": "high | medium | low",
      "original_excerpt": "Trích đoạn ngắn câu gốc (<= 120 ký tự) dùng để định vị gắn comment",
      "translated_excerpt": "Trích đoạn ngắn câu dịch tương ứng (<= 120 ký tự)",
      "warning_flag": "Mô tả ngắn gọn lỗi/rủi ro phát hiện",
      "suggested_comment": "Đề xuất chỉnh sửa cụ thể để luật sư tham khảo gắn vào Comment lề phải"
    }
  ]
}
Nếu bản dịch hoàn toàn chính xác, trả về "issues": []."""


class LegalIssue(BaseModel):
    clause_ref: str = "General"
    category: str = "terminology_warning"
    severity: Literal["high", "medium", "low"] = "medium"
    original_excerpt: str = ""
    translated_excerpt: str = ""
    warning_flag: str
    suggested_comment: str


class LegalReviewResult(BaseModel):
    enabled: bool = True
    provider: str = "heuristic"
    model: str = "rule-engine"
    summary: str = ""
    issues: list[LegalIssue] = Field(default_factory=list)


class LegalReviewService:
    """Runs LLM + deterministic numeric checks and attaches marginal comments without altering contract text."""

    def __init__(self) -> None:
        self._preferred_provider = os.getenv("LEGAL_REVIEW_PROVIDER", "anthropic").strip().lower()
        self._anthropic_model = os.getenv("ANTHROPIC_LEGAL_MODEL", "claude-sonnet-4-5")
        self._gemini_model = os.getenv("GEMINI_LEGAL_MODEL", "gemini-2.5-pro")

    def _extract_paired_segments(
        self,
        original_bytes: bytes,
        translated_bytes: bytes,
        extension: str,
        max_pairs: int = 120,
    ) -> list[dict[str, Any]]:
        pairs: list[dict[str, Any]] = []
        if extension == ".docx":
            o_doc = docx.Document(io.BytesIO(original_bytes))
            t_doc = docx.Document(io.BytesIO(translated_bytes))
            for idx, (op, tp) in enumerate(zip(o_doc.paragraphs, t_doc.paragraphs), start=1):
                os_txt, ts_txt = op.text.strip(), tp.text.strip()
                if os_txt or ts_txt:
                    pairs.append({"idx": idx, "src": os_txt, "tgt": ts_txt})
                if len(pairs) >= max_pairs:
                    return pairs

            for o_tbl, t_tbl in zip(o_doc.tables, t_doc.tables):
                seen_tc: set[int] = set()
                for o_row, t_row in zip(o_tbl.rows, t_tbl.rows):
                    for o_cell, t_cell in zip(o_row.cells, t_row.cells):
                        if id(o_cell._tc) in seen_tc:
                            continue
                        seen_tc.add(id(o_cell._tc))
                        os_txt, ts_txt = o_cell.text.strip(), t_cell.text.strip()
                        if os_txt or ts_txt:
                            pairs.append({"idx": len(pairs) + 1, "src": os_txt, "tgt": ts_txt})
                        if len(pairs) >= max_pairs:
                            return pairs
        elif extension == ".xlsx":
            o_wb = openpyxl.load_workbook(io.BytesIO(original_bytes), read_only=True, data_only=True)
            t_wb = openpyxl.load_workbook(io.BytesIO(translated_bytes), read_only=True, data_only=True)
            try:
                for o_ws, t_ws in zip(o_wb.worksheets, t_wb.worksheets):
                    for r_idx, (o_row, t_row) in enumerate(
                        zip(o_ws.iter_rows(values_only=True), t_ws.iter_rows(values_only=True)),
                        start=1,
                    ):
                        src_line = " | ".join(str(x).strip() for x in o_row if x is not None and str(x).strip())
                        tgt_line = " | ".join(str(x).strip() for x in t_row if x is not None and str(x).strip())
                        if src_line or tgt_line:
                            pairs.append({"idx": r_idx, "src": src_line, "tgt": tgt_line})
                        if len(pairs) >= max_pairs:
                            return pairs
            finally:
                o_wb.close()
                t_wb.close()
        return pairs

    @staticmethod
    def _deterministic_number_checks(pairs: list[dict[str, Any]]) -> list[LegalIssue]:
        """Fast deterministic check catching dropped or altered Arabic numerals across segments."""
        num_re = re.compile(r"\d+(?:[.,]\d+)*")
        findings: list[LegalIssue] = []

        for item in pairs:
            src, tgt = item["src"], item["tgt"]
            if not src or not tgt:
                continue
            src_nums = {re.sub(r"[.,]", "", m) for m in num_re.findall(src) if len(re.sub(r"[.,]", "", m)) >= 3}
            tgt_nums = {re.sub(r"[.,]", "", m) for m in num_re.findall(tgt) if len(re.sub(r"[.,]", "", m)) >= 3}
            missing_nums = src_nums - tgt_nums
            if missing_nums and not any(unit in tgt for unit in ("억", "만", "천")):
                findings.append(
                    LegalIssue(
                        clause_ref=f"Đoạn #{item['idx']}",
                        category="amount_mismatch",
                        severity="high",
                        original_excerpt=src[:120],
                        translated_excerpt=tgt[:120],
                        warning_flag=f"Chênh lệch/thiếu chuỗi số {sorted(missing_nums)} giữa câu gốc và bản dịch.",
                        suggested_comment=(
                            f"Kiểm tra lại giá trị số {sorted(missing_nums)} trong đoạn gốc so với bản dịch để tránh sai số tiền/ngày tháng."
                        ),
                    )
                )
            if len(findings) >= 8:
                break
        return findings

    async def _call_llm_review(
        self,
        pairs: list[dict[str, Any]],
        source_lang: str,
        target_lang: str,
    ) -> tuple[str, str, str, list[LegalIssue]]:
        prompt_payload = json.dumps(
            {
                "language_pair": f"{source_lang} -> {target_lang}",
                "segments": pairs[:90],
            },
            ensure_ascii=False,
        )

        anthropic_key = os.getenv("ANTHROPIC_API_KEY", "").strip()
        gemini_key = os.getenv("GEMINI_API_KEY", "").strip() or os.getenv("GOOGLE_API_KEY", "").strip()

        providers_order = (
            ["anthropic", "gemini"]
            if self._preferred_provider == "anthropic"
            else ["gemini", "anthropic"]
        )

        for provider in providers_order:
            if provider == "anthropic" and anthropic_key:
                try:
                    import anthropic

                    client = anthropic.AsyncAnthropic(api_key=anthropic_key)
                    resp = await client.messages.create(
                        model=self._anthropic_model,
                        max_tokens=2500,
                        temperature=0.1,
                        system=LEGAL_SYSTEM_PROMPT,
                        messages=[{"role": "user", "content": prompt_payload}],
                    )
                    raw_text = "".join(
                        block.text for block in resp.content if getattr(block, "type", None) == "text"
                    )
                    parsed = self._parse_llm_json(raw_text)
                    issues = [LegalIssue(**i) for i in parsed.get("issues", [])]
                    summary = parsed.get("summary", "Đã hoàn tất thẩm định pháp lý tự động bằng Claude.")
                    return "anthropic", self._anthropic_model, summary, issues
                except Exception as exc:
                    logger.warning("Anthropic legal review call failed (%s); trying fallback.", type(exc).__name__)

            if provider == "gemini" and gemini_key:
                try:
                    from google import genai
                    from google.genai import types

                    client = genai.Client(api_key=gemini_key)
                    resp = await client.aio.models.generate_content(
                        model=self._gemini_model,
                        contents=prompt_payload,
                        config=types.GenerateContentConfig(
                            system_instruction=LEGAL_SYSTEM_PROMPT,
                            temperature=0.1,
                            response_mime_type="application/json",
                        ),
                    )
                    parsed = self._parse_llm_json(resp.text or "{}")
                    issues = [LegalIssue(**i) for i in parsed.get("issues", [])]
                    summary = parsed.get("summary", "Đã hoàn tất thẩm định pháp lý tự động bằng Gemini.")
                    return "gemini", self._gemini_model, summary, issues
                except Exception as exc:
                    logger.warning("Gemini legal review call failed (%s); trying fallback.", type(exc).__name__)

        return (
            "heuristic",
            "deterministic-number-auditor",
            "Đã chạy kiểm tra đối chiếu số liệu & điều khoản cơ bản (chưa cấu hình ANTHROPIC_API_KEY / GEMINI_API_KEY).",
            [],
        )

    @staticmethod
    def _parse_llm_json(raw_text: str) -> dict[str, Any]:
        cleaned = raw_text.strip()
        fence = "\x60\x60\x60"
        if cleaned.startswith(fence):
            cleaned = re.sub(r"^\x60{3}(?:json)?\s*", "", cleaned)
            cleaned = re.sub(r"\s*\x60{3}$", "", cleaned)
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start != -1 and end != -1 and end > start:
            cleaned = cleaned[start : end + 1]
        return json.loads(cleaned)

    async def review_and_annotate(
        self,
        original_bytes: bytes,
        translated_bytes: bytes,
        merged_bilingual_bytes: bytes,
        extension: str,
        source_lang: str,
        target_lang: str,
    ) -> tuple[bytes, LegalReviewResult]:
        """Executes AI Legal Review and injects non-destructive marginal comments into .docx/.xlsx."""
        pairs = self._extract_paired_segments(original_bytes, translated_bytes, extension)
        det_issues = self._deterministic_number_checks(pairs)
        provider, model_name, summary, llm_issues = await self._call_llm_review(
            pairs=pairs,
            source_lang=source_lang,
            target_lang=target_lang,
        )

        combined: list[LegalIssue] = list(llm_issues)
        existing_excerpts = {i.original_excerpt[:50] for i in combined if i.original_excerpt}
        for d_issue in det_issues:
            if d_issue.original_excerpt[:50] not in existing_excerpts:
                combined.append(d_issue)

        annotated_bytes = merged_bilingual_bytes
        if combined:
            if extension == ".docx":
                annotated_bytes = self._attach_docx_comments(merged_bilingual_bytes, combined)
            elif extension == ".xlsx":
                annotated_bytes = self._attach_xlsx_comments(merged_bilingual_bytes, combined)

        result = LegalReviewResult(
            enabled=True,
            provider=provider,
            model=model_name,
            summary=summary,
            issues=combined,
        )
        return annotated_bytes, result

    @staticmethod
    def _attach_docx_comments(docx_bytes: bytes, issues: list[LegalIssue]) -> bytes:
        """Attaches native Word Review marginal comments WITHOUT changing any contract text."""
        doc = docx.Document(io.BytesIO(docx_bytes))
        if not hasattr(doc, "add_comment"):
            return docx_bytes

        all_paragraphs = list(doc.paragraphs)
        seen_tc: set[int] = set()
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    if id(cell._tc) not in seen_tc:
                        seen_tc.add(id(cell._tc))
                        all_paragraphs.extend(cell.paragraphs)

        for issue in issues:
            comment_body = (
                f"[{issue.severity.upper()} | {issue.category}] ({issue.clause_ref})\n"
                f"Cảnh báo: {issue.warning_flag}\n"
                f"Đề xuất: {issue.suggested_comment}"
            )
            target_para = None
            needle_orig = issue.original_excerpt.strip()[:45]
            needle_trans = issue.translated_excerpt.strip()[:45]

            for p in all_paragraphs:
                p_txt = p.text
                if not p_txt or not p.runs:
                    continue
                if (needle_trans and needle_trans in p_txt) or (needle_orig and needle_orig in p_txt):
                    target_para = p
                    break

            if target_para is None:
                for p in all_paragraphs:
                    if p.runs and p.text.strip():
                        target_para = p
                        break

            if target_para is not None and target_para.runs:
                try:
                    doc.add_comment(
                        runs=target_para.runs,
                        text=comment_body,
                        author="AI Legal Review",
                        initials="ALR",
                    )
                except Exception as exc:
                    logger.debug("Skipped comment attachment on paragraph (%s)", type(exc).__name__)

        out = io.BytesIO()
        doc.save(out)
        return out.getvalue()

    @staticmethod
    def _attach_xlsx_comments(xlsx_bytes: bytes, issues: list[LegalIssue]) -> bytes:
        """Attaches Excel cell sticky-note comments WITHOUT altering cell values/formulas."""
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        if not wb.worksheets:
            return xlsx_bytes

        for issue in issues:
            note_text = (
                f"[{issue.severity.upper()}] {issue.clause_ref}\n"
                f"Cảnh báo: {issue.warning_flag}\n"
                f"Đề xuất: {issue.suggested_comment}"
            )
            placed = False
            needle_orig = issue.original_excerpt.strip()[:35]
            needle_trans = issue.translated_excerpt.strip()[:35]

            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        val_str = str(cell.value or "")
                        if not val_str:
                            continue
                        if (needle_trans and needle_trans in val_str) or (needle_orig and needle_orig in val_str):
                            cell.comment = ExcelComment(text=note_text[:1000], author="AI Legal Review")
                            placed = True
                            break
                    if placed:
                        break
                if placed:
                    break

            if not placed:
                first_ws = wb.worksheets[0]
                anchor = first_ws.cell(row=1, column=1)
                existing = f"{anchor.comment.text}\n---\n" if anchor.comment else ""
                anchor.comment = ExcelComment(text=(existing + note_text)[:1500], author="AI Legal Review")

        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()


_legal_review_singleton: Optional[LegalReviewService] = None


def get_legal_review_service() -> LegalReviewService:
    global _legal_review_singleton
    if _legal_review_singleton is None:
        _legal_review_singleton = LegalReviewService()
    return _legal_review_singleton
