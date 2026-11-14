"""Unit & integration tests for VI+EN overwrite (vi-ko), VI-EN-KR check+fix+KO, and FastAPI endpoints."""

from __future__ import annotations

import io
import os
import unittest
from unittest.mock import AsyncMock, MagicMock, patch

import docx
import openpyxl
from fastapi.testclient import TestClient

os.environ.setdefault("APP_ENV", "development")
os.environ.setdefault("ALLOW_DEV_AUTH_BYPASS", "true")

from app.bilingual_layout import (
    build_bilingual_document,
    build_trilingual_document,
    build_trilingual_from_bilingual_document,
    extract_vi_en_pairs_for_check,
)
from app.en_checker import (
    EnCheckItem,
    EnCheckerService,
    detect_segment_lang_heuristic,
)
from app.main import app


def _make_docx_bytes(paragraphs: list[str], tables: list[list[list[str]]] | None = None) -> bytes:
    doc = docx.Document()
    for p in paragraphs:
        doc.add_paragraph(p)
    if tables:
        for rows_data in tables:
            cols = max(len(r) for r in rows_data)
            tbl = doc.add_table(rows=len(rows_data), cols=cols)
            for r_i, row_vals in enumerate(rows_data):
                for c_i, val in enumerate(row_vals):
                    tbl.cell(r_i, c_i).text = val
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _read_docx_paragraphs(raw: bytes) -> list[str]:
    doc = docx.Document(io.BytesIO(raw))
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def _read_docx_tables(raw: bytes) -> list[list[list[str]]]:
    doc = docx.Document(io.BytesIO(raw))
    out = []
    for tbl in doc.tables:
        t_rows = []
        for row in tbl.rows:
            t_rows.append([c.text.strip() for c in row.cells])
        out.append(t_rows)
    return out


def _make_xlsx_bytes(rows: list[list[str]]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    for r_i, row in enumerate(rows, start=1):
        for c_i, val in enumerate(row, start=1):
            ws.cell(row=r_i, column=c_i, value=val)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestLanguageHeuristic(unittest.TestCase):
    def test_detect_segment_lang_heuristic(self) -> None:
        self.assertEqual(detect_segment_lang_heuristic("Điều 1. Giá trị hợp đồng"), "vi")
        self.assertEqual(detect_segment_lang_heuristic("Article 1. Contract Value and Payment Terms"), "en")
        self.assertEqual(detect_segment_lang_heuristic("제1조. 계약 금액"), "ko")
        self.assertEqual(detect_segment_lang_heuristic("100,000,000"), "neutral")
        self.assertEqual(detect_segment_lang_heuristic(""), "neutral")


class TestViKoAlternating(unittest.TestCase):
    def test_docx_vi_en_paragraphs_overwrite_en_with_ko(self) -> None:
        orig = _make_docx_bytes([
            "Điều 1. Đối tượng hợp đồng",
            "Article 1. Subject of the Contract",
            "Bên A thanh toán cho Bên B trong 15 ngày.",
            "Party A shall pay Party B within 15 days.",
        ])
        # Simulated Cloud Translation (source=vi, target=ko): VI paras become KO, EN paras stay EN/garbled
        trans_ko = _make_docx_bytes([
            "제1조. 계약의 목적",
            "Article 1. Subject of the Contract",
            "갑은 15일 이내에 을에게 지급합니다.",
            "Party A shall pay Party B within 15 days.",
        ])

        out_bytes = build_bilingual_document(
            original_bytes=orig,
            translated_bytes=trans_ko,
            extension=".docx",
            layout_mode="alternating",
            source_lang="vi",
            target_lang="ko",
            original_has_en=True,
        )
        paras = _read_docx_paragraphs(out_bytes)
        self.assertEqual(
            paras,
            [
                "Điều 1. Đối tượng hợp đồng",
                "제1조. 계약의 목적",
                "Bên A thanh toán cho Bên B trong 15 ngày.",
                "갑은 15일 이내에 을에게 지급합니다.",
            ],
        )

    def test_docx_vi_en_table_overwrites_en_col_with_vi_to_ko(self) -> None:
        orig = _make_docx_bytes(
            paragraphs=[],
            tables=[[
                ["Điều 2. Bảo mật", "Article 2. Confidentiality"],
                ["Hai bên cam kết bảo mật thông tin.", "Both parties undertake to keep information confidential."],
            ]],
        )
        trans_ko = _make_docx_bytes(
            paragraphs=[],
            tables=[[
                ["제2조. 비밀유지", "Article 2. Confidentiality"],
                ["양 당사자는 정보를 비밀로 유지할 것을 약속합니다.", "Both parties undertake..."],
            ]],
        )
        out_bytes = build_bilingual_document(
            original_bytes=orig,
            translated_bytes=trans_ko,
            extension=".docx",
            layout_mode="alternating",
            source_lang="vi",
            target_lang="ko",
            original_has_en=True,
        )
        tables = _read_docx_tables(out_bytes)
        self.assertEqual(
            tables[0],
            [
                ["Điều 2. Bảo mật", "제2조. 비밀유지"],
                ["Hai bên cam kết bảo mật thông tin.", "양 당사자는 정보를 비밀로 유지할 것을 약속합니다."],
            ],
        )

    def test_docx_vi_only_alternating_inserts_ko(self) -> None:
        orig = _make_docx_bytes(["Điều 1. Hiệu lực hợp đồng"])
        trans_ko = _make_docx_bytes(["제1조. 계약의 효력"])
        out_bytes = build_bilingual_document(
            original_bytes=orig,
            translated_bytes=trans_ko,
            extension=".docx",
            layout_mode="alternating",
            source_lang="vi",
            target_lang="ko",
            original_has_en=False,
        )
        self.assertEqual(_read_docx_paragraphs(out_bytes), ["Điều 1. Hiệu lực hợp đồng", "제1조. 계약의 효력"])

    def test_xlsx_vi_en_columns_overwrite_en_with_ko(self) -> None:
        orig = _make_xlsx_bytes([
            ["Điều 1. Thanh toán", "Article 1. Payment"],
            ["Thanh toán 100% giá trị.", "Pay 100% of the value."],
        ])
        trans_ko = _make_xlsx_bytes([
            ["제1조. 지급", "Article 1. Payment"],
            ["가치의 100%를 지급합니다.", "Pay 100% of the value."],
        ])
        out_bytes = build_bilingual_document(
            original_bytes=orig,
            translated_bytes=trans_ko,
            extension=".xlsx",
            layout_mode="alternating",
            source_lang="vi",
            target_lang="ko",
            original_has_en=True,
        )
        wb = openpyxl.load_workbook(io.BytesIO(out_bytes))
        ws = wb.active
        self.assertEqual(ws.cell(1, 1).value, "Điều 1. Thanh toán")
        self.assertEqual(ws.cell(1, 2).value, "제1조. 지급")
        self.assertEqual(ws.cell(2, 1).value, "Thanh toán 100% giá trị.")
        self.assertEqual(ws.cell(2, 2).value, "가치의 100%를 지급합니다.")


class TestViEnKrCheckAndLayouts(unittest.IsolatedAsyncioTestCase):
    async def test_gemini_batch_check_fixes_only_mismatched_en(self) -> None:
        checker = EnCheckerService()
        mock_gemini = AsyncMock(
            return_value={
                "results": [
                    {"id": "p:0:0", "is_accurate": True},
                    {
                        "id": "p:1:0",
                        "is_accurate": False,
                        "corrected_en": "Penalty for breach is 8% of the contract value.",
                    },
                ]
            }
        )
        with patch.object(checker, "_call_gemini_json", mock_gemini):
            res = await checker.verify_and_fix_en_batch([
                EnCheckItem(pair_id="p:0:0", vi_text="Điều 1. Phạm vi", en_text="Article 1. Scope"),
                EnCheckItem(
                    pair_id="p:1:0",
                    vi_text="Phạt vi phạm 8% giá trị hợp đồng.",
                    en_text="Penalty for breach is 5% of contract value.",
                ),
            ])
        self.assertEqual(mock_gemini.await_count, 1)
        self.assertEqual(res.stats.checked_count, 2)
        self.assertEqual(res.stats.corrected_count, 1)
        self.assertEqual(res.corrected_map["p:0:0"], "Article 1. Scope")
        self.assertEqual(res.corrected_map["p:1:0"], "Penalty for breach is 8% of the contract value.")

    async def test_docx_vi_en_kr_all_3_layouts_from_bilingual_original(self) -> None:
        orig = _make_docx_bytes(
            paragraphs=[
                "Điều 1. Phạt vi phạm 8% giá trị hợp đồng.",
                "Article 1. Penalty for breach is 5% of contract value.",
            ],
            tables=[[
                ["Điều 2. Bảo mật thông tin.", "Article 2. Confidentiality."],
            ]],
        )
        trans_ko = _make_docx_bytes(
            paragraphs=[
                "제1조. 계약 위반 위약금은 계약 금액의 8%입니다.",
                "Article 1...",
            ],
            tables=[[
                ["제2조. 정보 비밀유지.", "Article 2..."],
            ]],
        )
        pairs = extract_vi_en_pairs_for_check(orig, ".docx")
        self.assertEqual(len(pairs), 2)

        corrected_map = {
            pairs[0].pair_id: "Article 1. Penalty for breach is 8% of the contract value.",
            pairs[1].pair_id: "Article 2. Confidentiality.",
        }

        # 1. alternating
        alt_bytes = build_trilingual_from_bilingual_document(
            original_bytes=orig,
            translated_ko_bytes=trans_ko,
            corrected_en_map=corrected_map,
            extension=".docx",
            layout_mode="alternating",
        )
        self.assertEqual(
            _read_docx_paragraphs(alt_bytes),
            [
                "Điều 1. Phạt vi phạm 8% giá trị hợp đồng.",
                "Article 1. Penalty for breach is 8% of the contract value.",
                "제1조. 계약 위반 위약금은 계약 금액의 8%입니다.",
            ],
        )
        self.assertEqual(
            _read_docx_tables(alt_bytes)[0],
            [["Điều 2. Bảo mật thông tin.", "Article 2. Confidentiality.", "제2조. 정보 비밀유지."]],
        )

        # 2. replace_column
        rep_bytes = build_trilingual_from_bilingual_document(
            original_bytes=orig,
            translated_ko_bytes=trans_ko,
            corrected_en_map=corrected_map,
            extension=".docx",
            layout_mode="replace_column",
        )
        self.assertEqual(
            _read_docx_tables(rep_bytes)[0],
            [["Điều 2. Bảo mật thông tin.", "Article 2. Confidentiality.", "제2조. 정보 비밀유지."]],
        )

        # 3. side_by_side
        sbs_bytes = build_trilingual_from_bilingual_document(
            original_bytes=orig,
            translated_ko_bytes=trans_ko,
            corrected_en_map=corrected_map,
            extension=".docx",
            layout_mode="side_by_side",
        )
        sbs_tables = _read_docx_tables(sbs_bytes)
        self.assertEqual(len(sbs_tables[0]), 3)  # header + para row + table row
        self.assertEqual(
            sbs_tables[0][1],
            [
                "Điều 1. Phạt vi phạm 8% giá trị hợp đồng.",
                "Article 1. Penalty for breach is 8% of the contract value.",
                "제1조. 계약 위반 위약금은 계약 금액의 8%입니다.",
            ],
        )
        self.assertEqual(
            sbs_tables[0][2],
            ["Điều 2. Bảo mật thông tin.", "Article 2. Confidentiality.", "제2조. 정보 비밀유지."],
        )

    def test_vi_en_kr_vi_only_replace_column_builds_3_cols(self) -> None:
        orig = _make_docx_bytes(["Điều 1. Hiệu lực"])
        en_b = _make_docx_bytes(["Article 1. Effectiveness"])
        ko_b = _make_docx_bytes(["제1조. 효력"])
        out_bytes = build_trilingual_document(
            original_bytes=orig,
            translated_en_bytes=en_b,
            translated_ko_bytes=ko_b,
            extension=".docx",
            layout_mode="replace_column",
        )
        tables = _read_docx_tables(out_bytes)
        self.assertEqual(tables[0][1], ["Điều 1. Hiệu lực", "Article 1. Effectiveness", "제1조. 효력"])


class TestFastApiEndpoints(unittest.TestCase):
    def setUp(self) -> None:
        self.client = TestClient(app)

    def test_health_and_config_and_422_empty(self) -> None:
        r_health = self.client.get("/api/health")
        self.assertEqual(r_health.status_code, 200)
        self.assertEqual(r_health.json()["status"], "ok")

        r_cfg = self.client.get("/api/config")
        self.assertEqual(r_cfg.status_code, 200)
        cfg = r_cfg.json()
        self.assertEqual(len(cfg["supported_pairs"]), 7)
        self.assertEqual(cfg["default_layout"], "alternating")
        self.assertEqual([m["id"] for m in cfg["layout_modes"]], ["alternating", "replace_column", "side_by_side"])

        r_tr_empty = self.client.post("/api/translate")
        self.assertEqual(r_tr_empty.status_code, 422)
        self.assertIn("detail", r_tr_empty.json())

        r_ocr_empty = self.client.post("/api/ocr")
        self.assertEqual(r_ocr_empty.status_code, 422)
        self.assertIn("detail", r_ocr_empty.json())

    def test_translate_vi_kr_does_not_call_gemini_en_check(self) -> None:
        orig = _make_docx_bytes([
            "Điều 1. Giá trị hợp đồng là 100 triệu đồng.",
            "Article 1. Contract value is 100 million VND.",
        ])
        trans_ko = _make_docx_bytes([
            "제1조. 계약 금액은 1억 동입니다.",
            "Article 1. Contract value is 100 million VND.",
        ])
        mock_ts = MagicMock()
        mock_ts.inspect_document.return_value = MagicMock(
            filename="c.docx",
            extension=".docx",
            mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            size_bytes=len(orig),
            char_count=100,
            estimated_pages=1,
            requires_async=False,
            pdf_type=None,
            ocr_vision_requests=0,
            model_dump=lambda: {"estimated_pages": 1, "char_count": 100},
        )
        mock_ts.translate_sync.return_value = (trans_ko, "mock-project")
        mock_ts.store_session_artifact.return_value = "file:///tmp/mock.docx"

        mock_en_checker = EnCheckerService()
        mock_gemini = AsyncMock()
        with patch("app.main.get_translation_service", return_value=mock_ts), \
             patch("app.main.get_en_checker_service", return_value=mock_en_checker), \
             patch.object(mock_en_checker, "_call_gemini_json", mock_gemini):
            resp = self.client.post(
                "/api/translate",
                files={"file": ("contract.docx", orig, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                data={"source_lang": "vi", "target_lang": "ko", "layout_mode": "alternating"},
            )
        self.assertEqual(resp.status_code, 200)
        body = resp.json()
        self.assertIsNone(body.get("en_check"))
        self.assertEqual(mock_gemini.await_count, 0)

    def test_translate_vi_en_kr_bilingual_calls_gemini_once_and_skips_vi_en_cloud_translate(self) -> None:
        orig = _make_docx_bytes([
            "Điều 1. Thời hạn là 30 ngày.",
            "Article 1. Term is 15 days.",
        ])
        trans_ko = _make_docx_bytes([
            "제1조. 기간은 30일입니다.",
            "Article 1...",
        ])
        mock_ts = MagicMock()
        mock_ts.inspect_document.return_value = MagicMock(
            filename="c.docx",
            extension=".docx",
            mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            size_bytes=len(orig),
            char_count=80,
            estimated_pages=1,
            requires_async=False,
            pdf_type=None,
            ocr_vision_requests=0,
            model_dump=lambda: {"estimated_pages": 1, "char_count": 80},
        )
        mock_ts.translate_sync.return_value = (trans_ko, "mock-project")
        saved_artifacts: list[bytes] = []

        def _capture_store(**kwargs):
            if kwargs.get("stage") == "final":
                saved_artifacts.append(kwargs["content"])
            return "file:///tmp/mock.docx"

        mock_ts.store_session_artifact.side_effect = _capture_store

        mock_en_checker = EnCheckerService()
        mock_gemini = AsyncMock(
            return_value={
                "results": [
                    {"id": "p:0:0", "is_accurate": False, "corrected_en": "Article 1. Term is 30 days."}
                ]
            }
        )
        with patch("app.main.get_translation_service", return_value=mock_ts), \
             patch("app.main.get_en_checker_service", return_value=mock_en_checker), \
             patch.object(mock_en_checker, "_call_gemini_json", mock_gemini):
            resp = self.client.post(
                "/api/translate",
                files={"file": ("contract.docx", orig, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                data={"source_lang": "vi", "target_lang": "en,ko", "layout_mode": "replace_column"},
            )
        self.assertEqual(resp.status_code, 200)
        body = resp.json()
        self.assertEqual(body["en_check"]["status"], "ok")
        self.assertEqual(body["en_check"]["checked_count"], 1)
        self.assertEqual(body["en_check"]["corrected_count"], 1)
        self.assertEqual(mock_gemini.await_count, 1)
        # Cloud Translation only called once for KO (no wasted VI->EN Cloud Translation call)
        self.assertEqual(mock_ts.translate_sync.call_count, 1)
        self.assertEqual(mock_ts.translate_sync.call_args.kwargs["target_lang"], "ko")
        self.assertEqual(
            _read_docx_paragraphs(saved_artifacts[0]),
            [
                "Điều 1. Thời hạn là 30 ngày.",
                "Article 1. Term is 30 days.",
                "제1조. 기간은 30일입니다.",
            ],
        )

    def test_translate_vi_en_kr_vi_only_skips_gemini_and_translates_both(self) -> None:
        orig = _make_docx_bytes(["Điều 1. Hiệu lực hợp đồng."])
        trans_ko = _make_docx_bytes(["제1조. 계약의 효력."])
        trans_en = _make_docx_bytes(["Article 1. Effectiveness of the contract."])

        mock_ts = MagicMock()
        mock_ts.inspect_document.return_value = MagicMock(
            filename="c.docx",
            extension=".docx",
            mime_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            size_bytes=len(orig),
            char_count=40,
            estimated_pages=1,
            requires_async=False,
            pdf_type=None,
            ocr_vision_requests=0,
            model_dump=lambda: {"estimated_pages": 1, "char_count": 40},
        )
        mock_ts.translate_sync.side_effect = [
            (trans_ko, "mock-project"),
            (trans_en, "mock-project"),
        ]
        mock_ts.store_session_artifact.return_value = "file:///tmp/mock.docx"

        mock_en_checker = EnCheckerService()
        mock_gemini = AsyncMock()
        with patch("app.main.get_translation_service", return_value=mock_ts), \
             patch("app.main.get_en_checker_service", return_value=mock_en_checker), \
             patch.object(mock_en_checker, "_call_gemini_json", mock_gemini):
            resp = self.client.post(
                "/api/translate",
                files={"file": ("contract.docx", orig, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                data={"source_lang": "vi", "target_lang": "en,ko", "layout_mode": "alternating"},
            )
        self.assertEqual(resp.status_code, 200)
        self.assertIsNone(resp.json().get("en_check"))
        self.assertEqual(mock_gemini.await_count, 0)
        self.assertEqual(mock_ts.translate_sync.call_count, 2)


if __name__ == "__main__":
    unittest.main()
