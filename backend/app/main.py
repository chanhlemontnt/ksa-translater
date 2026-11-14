"""FastAPI Entrypoint for Bilingual Contract Translation Web App.

Architecture highlights:
- Single domain: dich.tnt.name.vn (Cloudflare Worker routes /api/* -> Cloud Run).
- Auth: Cloudflare Access JWT (Cf-Access-Jwt-Assertion) + Worker shared secret.
- Sync flow: <= 15 pages AND <= 20,000 extracted chars -> Cloud Translation v3 translateDocument.
- Async flow: > 15 pages OR > 20,000 chars -> batchTranslateDocument via GCS + polling.
- Storage & Security: GCS 7-day lifecycle; Cloud DLP redacts logs/audit metadata ONLY.
  Downloaded contract files are NEVER altered/redacted by DLP (100% genuine legal content).
- Optional AI Legal Review: Adds warning flags + Word/Excel marginal comments without editing text.
"""

from __future__ import annotations

import io
import logging
import os
import time
import uuid
from pathlib import Path
from typing import Any, Literal, Optional
from urllib.parse import quote

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse

from app.auth_cf_access import AuthenticatedUser, require_cf_access
from app.bilingual_layout import build_bilingual_document, build_trilingual_document
from app.dlp_service import get_dlp_service
from app.key_rotator import get_project_rotator
from app.legal_review import LegalReviewResult, get_legal_review_service
from app.translation_service import (
    MIME_DOCX,
    MIME_PDF,
    MIME_XLSX,
    SUPPORTED_LANG_PAIRS,
    SYNC_MAX_CHARS,
    SYNC_MAX_PAGES,
    get_translation_service,
    normalize_lang_pair,
)
from app.vision_ocr_service import get_vision_ocr_service

logger = logging.getLogger("bilingual_contract_app")
logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)

MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "25"))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024
ALLOWED_ORIGINS = [
    origin.strip()
    for origin in os.getenv("ALLOWED_ORIGINS", "https://dich.tnt.name.vn").split(",")
    if origin.strip()
]

app = FastAPI(
    title="Dịch Hợp Đồng Song Ngữ API (Việt - Hàn - Anh)",
    version="1.0.0",
    docs_url="/api/docs" if os.getenv("APP_ENV") == "development" else None,
    redoc_url=None,
    openapi_url="/api/openapi.json" if os.getenv("APP_ENV") == "development" else None,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception) -> JSONResponse:
    """Catch unhandled exceptions and return JSON 500 instead of plain text."""
    logger.exception("Unhandled exception on %s %s", request.method, request.url.path, exc_info=exc)
    return JSONResponse(
        status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
        content={"detail": "Lỗi hệ thống nội bộ. Vui lòng thử lại sau."},
    )


def _safe_output_filename(original_name: str, source_lang: str, target_lang: str, layout_mode: str) -> str:
    stem = Path(original_name).stem or "contract"
    ext = Path(original_name).suffix.lower() or ".docx"
    # PDF inputs are OCR-extracted and assembled into bilingual .docx outputs
    out_ext = ".docx" if ext == ".pdf" else ext
    clean_stem = "".join(ch if ch.isalnum() or ch in ("-", "_", ".", " ") else "_" for ch in stem).strip()
    return f"{clean_stem}_{source_lang}-{target_lang}_{layout_mode}{out_ext}"


@app.get("/healthz")
@app.get("/api/health")
async def health_check() -> dict[str, Any]:
    """Unauthenticated container liveness check for Cloud Run."""
    return {
        "status": "ok",
        "service": "dich-hop-dong-song-ngu-backend",
        "timestamp": int(time.time()),
    }


@app.get("/api/config")
async def get_app_config(
    user: AuthenticatedUser = Depends(require_cf_access),
) -> dict[str, Any]:
    """Returns supported language pairs, layout modes, sync/async thresholds, and sanitized pool status."""
    rotator = get_project_rotator()
    return {
        "user": {"email": get_dlp_service().sanitize_log_text(user.email)},
        "default_layout": "alternating",
        "supported_pairs": [
            {"source": "vi", "target": "ko", "label": "Việt → Hàn (VI → KO)"},
            {"source": "ko", "target": "vi", "label": "Hàn → Việt (KO → VI)"},
            {"source": "en", "target": "ko", "label": "Anh → Hàn (EN → KO)"},
            {"source": "ko", "target": "en", "label": "Hàn → Anh (KO → EN)"},
            {"source": "vi", "target": "en", "label": "Việt → Anh (VI → EN)"},
            {"source": "en", "target": "vi", "label": "Anh → Việt (EN → VI)"},
            {"id": "vi-en-kr", "source": "vi", "target": "en,ko", "label": "Việt → Anh & Hàn (VI → EN + KO)"},
        ],
        "layout_modes": [
            {
                "id": "alternating",
                "label": "Đoạn văn xen kẽ (Alternating)",
                "description": "Mỗi đoạn gốc theo sau ngay bởi đoạn dịch đối ứng (Mặc định).",
            },
            {
                "id": "replace_column",
                "label": "Thay thế ngôn ngữ (Replace Column)",
                "description": "Chuyển hợp đồng song ngữ sẵn có (VD: Việt-Anh → Việt-Hàn) bằng cách thay cột phải.",
            },
            {
                "id": "side_by_side",
                "label": "Bảng đối chiếu (Side-by-side)",
                "description": "Đặt câu gốc cột trái, bản dịch cột phải (2 hoặc 3 cột), giữ định dạng điều khoản.",
            },
        ],
        "thresholds": {
            "sync_max_pages": SYNC_MAX_PAGES,
            "sync_max_chars": SYNC_MAX_CHARS,
            "max_upload_mb": MAX_UPLOAD_MB,
            "retention_days": 7,
        },
        "pool_projects_count": len(rotator.get_pool_summary()),
    }


@app.post("/api/translate")
async def translate_contract(
    request: Request,
    file: UploadFile = File(..., description="Hợp đồng .docx, .pdf hoặc bảng biểu .xlsx"),
    source_lang: str = Form(..., description="Mã ngôn ngữ nguồn: vi | en | ko"),
    target_lang: str = Form(..., description="Mã ngôn ngữ đích: ko | vi | en | en,ko"),
    layout_mode: Literal["side_by_side", "alternating", "replace_column"] = Form("alternating"),
    use_glossary: bool = Form(True),
    enable_legal_review: bool = Form(False),
    replace_col_index: int = Form(1, ge=0, le=10),
    user: AuthenticatedUser = Depends(require_cf_access),
) -> JSONResponse:
    """Main endpoint to upload and translate a legal contract (.docx / .xlsx / .pdf).

    - Small files (<=15 pages AND <=20,000 chars): synchronous translateDocument -> layout -> legal review.
    - Large files (>15 pages OR >20,000 chars): asynchronous batchTranslateDocument on GCS + polling.
    """
    dlp = get_dlp_service()
    translation_svc = get_translation_service()
    legal_svc = get_legal_review_service()

    src, tgt = normalize_lang_pair(source_lang, target_lang)
    if (src, tgt) not in SUPPORTED_LANG_PAIRS:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Cặp ngôn ngữ '{src} -> {tgt}' chưa được hỗ trợ. Hỗ trợ: vi-ko, en-ko, vi-en (2 chiều) và vi-en-kr.",
        )

    is_trilingual = (src == "vi" and tgt == "en,ko")
    effective_tgt = tgt
    if is_trilingual and layout_mode == "replace_column":
        # TODO(vi-en-kr-replace-column): Pending clarification from Anh Thanh on
        # "thay thế khi cột EN bất đồng với VI". Currently treats vi-en-kr + replace_column
        # as VI -> KO overwriting EN column (same as vi-kr).
        is_trilingual = False
        effective_tgt = "ko"

    raw_filename = file.filename or "contract.docx"
    ext = Path(raw_filename).suffix.lower()
    if ext not in (".docx", ".xlsx", ".pdf"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Chỉ hỗ trợ định dạng hợp đồng Word (.docx), Excel (.xlsx) hoặc PDF (.pdf).",
        )

    file_bytes = await file.read()
    if not file_bytes:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="File tải lên rỗng.")
    if len(file_bytes) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"Dung lượng file vượt ngưỡng tối đa {MAX_UPLOAD_MB}MB.",
        )

    job_id = str(uuid.uuid4())
    started_at = time.time()

    # 1. Extract character & page metrics (with Auto-Detect PDF / Vision OCR if .pdf)
    if ext == ".pdf":
        working_bytes, metrics, _ = translation_svc.prepare_pdf_for_translation(
            file_bytes=file_bytes,
            filename=raw_filename,
        )
        working_ext = ".docx"
        orig_mime = MIME_PDF
    else:
        working_bytes = file_bytes
        working_ext = ext
        metrics = translation_svc.inspect_document(file_bytes=file_bytes, filename=raw_filename)
        orig_mime = metrics.mime_type

    output_filename = _safe_output_filename(raw_filename, src, tgt, layout_mode)

    # Log sanitized request metadata via DLP (never log raw filename/email PII unredacted)
    safe_audit = dlp.sanitize_dict(
        {
            "event": "translation_requested",
            "job_id": job_id,
            "user_email": user.email,
            "original_filename": raw_filename,
            "source_lang": src,
            "target_lang": tgt,
            "layout_mode": layout_mode,
            "use_glossary": use_glossary,
            "enable_legal_review": enable_legal_review,
            "char_count": metrics.char_count,
            "estimated_pages": metrics.estimated_pages,
            "requires_async": metrics.requires_async,
            "pdf_type": metrics.pdf_type,
            "ocr_vision_requests": metrics.ocr_vision_requests,
            "timestamp": int(started_at),
        }
    )
    logger.info("Job %s initialized: %s", job_id, safe_audit)
    dlp.save_redacted_audit_log(session_id=job_id, audit_payload=safe_audit)

    # Store original file in session prefix (subject to 7-day GCS lifecycle auto-deletion)
    orig_ref = translation_svc.store_session_artifact(
        session_id=job_id,
        stage="original",
        filename=raw_filename,
        content=file_bytes,
        mime_type=orig_mime,
    )
    layout_source_ref = orig_ref
    if ext == ".pdf":
        layout_source_ref = translation_svc.store_session_artifact(
            session_id=job_id,
            stage="prepared",
            filename=f"{Path(raw_filename).stem}.docx",
            content=working_bytes,
            mime_type=MIME_DOCX,
        )

    # 2A. ASYNC FLOW (> 15 pages OR > 20,000 chars)
    if metrics.requires_async:
        if is_trilingual:
            batch_en = translation_svc.start_batch_async(
                session_id=job_id,
                file_bytes=working_bytes,
                metrics=metrics,
                source_lang=src,
                target_lang="en",
                use_glossary=use_glossary,
                stage_tag="en",
            )
            batch_ko = translation_svc.start_batch_async(
                session_id=job_id,
                file_bytes=working_bytes,
                metrics=metrics,
                source_lang=src,
                target_lang="ko",
                use_glossary=use_glossary,
                stage_tag="ko",
            )
            batch_meta = {
                "multi_target": True,
                "targets": {"en": batch_en, "ko": batch_ko},
                "bucket": batch_ko.get("bucket") or batch_en.get("bucket"),
                "project_id": batch_ko.get("project_id") or batch_en.get("project_id"),
            }
        else:
            batch_meta = translation_svc.start_batch_async(
                session_id=job_id,
                file_bytes=working_bytes,
                metrics=metrics,
                source_lang=src,
                target_lang=effective_tgt,
                use_glossary=use_glossary,
            )
        job_state = {
            "job_id": job_id,
            "mode": "async",
            "status": "processing",
            "owner_sub": user.sub,
            "created_at": int(started_at),
            "source_lang": src,
            "target_lang": tgt,
            "effective_target_lang": effective_tgt,
            "is_trilingual": is_trilingual,
            "layout_mode": layout_mode,
            "use_glossary": use_glossary,
            "enable_legal_review": enable_legal_review,
            "replace_col_index": replace_col_index,
            "extension": working_ext,
            "mime_type": metrics.mime_type,
            "output_filename": output_filename,
            "metrics": metrics.model_dump(),
            "original_ref": layout_source_ref,
            "batch_meta": batch_meta,
            "project_id": batch_meta.get("project_id"),
        }
        translation_svc.save_job_state(job_id=job_id, state=job_state)
        return JSONResponse(
            status_code=status.HTTP_202_ACCEPTED,
            content={
                "job_id": job_id,
                "mode": "async",
                "status": "processing",
                "message": (
                    f"Hợp đồng lớn ({metrics.estimated_pages} trang, {metrics.char_count:,} ký tự) "
                    "đang được xử lý bất đồng bộ qua Batch Translation."
                ),
                "poll_url": f"/api/jobs/{job_id}",
                "download_url": f"/api/jobs/{job_id}/download",
                "metrics": metrics.model_dump(),
            },
        )

    # 2B. SYNC FLOW (<= 15 pages AND <= 20,000 chars)
    if is_trilingual:
        translated_en_bytes, project_en = translation_svc.translate_sync(
            session_id=job_id,
            file_bytes=working_bytes,
            metrics=metrics,
            source_lang=src,
            target_lang="en",
            use_glossary=use_glossary,
        )
        translated_ko_bytes, project_used = translation_svc.translate_sync(
            session_id=job_id,
            file_bytes=working_bytes,
            metrics=metrics,
            source_lang=src,
            target_lang="ko",
            use_glossary=use_glossary,
        )
        final_bytes = build_trilingual_document(
            original_bytes=working_bytes,
            translated_en_bytes=translated_en_bytes,
            translated_ko_bytes=translated_ko_bytes,
            extension=working_ext,
            layout_mode=layout_mode,
            source_lang=src,
        )
        review_target_bytes = translated_ko_bytes
        review_target_lang = "ko"
        project_used = project_used or project_en
    else:
        translated_raw_bytes, project_used = translation_svc.translate_sync(
            session_id=job_id,
            file_bytes=working_bytes,
            metrics=metrics,
            source_lang=src,
            target_lang=effective_tgt,
            use_glossary=use_glossary,
        )
        # 3. Build Bilingual Layout (.docx / .xlsx)
        final_bytes = build_bilingual_document(
            original_bytes=working_bytes,
            translated_bytes=translated_raw_bytes,
            extension=working_ext,
            layout_mode=layout_mode,
            source_lang=src,
            target_lang=effective_tgt,
            replace_col_index=replace_col_index,
        )
        review_target_bytes = translated_raw_bytes
        review_target_lang = effective_tgt

    # 4. Optional AI Legal Review (warning flags + marginal comments ONLY; never mutates contract text)
    legal_result: Optional[LegalReviewResult] = None
    if enable_legal_review:
        final_bytes, legal_result = await legal_svc.review_and_annotate(
            original_bytes=working_bytes,
            translated_bytes=review_target_bytes,
            merged_bilingual_bytes=final_bytes,
            extension=working_ext,
            source_lang=src,
            target_lang=review_target_lang,
        )

    # 5. Store untouched final file for download (DLP NEVER applied to user download artifact!)
    final_ref = translation_svc.store_session_artifact(
        session_id=job_id,
        stage="final",
        filename=output_filename,
        content=final_bytes,
        mime_type=metrics.mime_type,
        bucket_name=project_used and get_project_rotator().get_bucket_for_project(project_used),
    )

    elapsed_ms = int((time.time() - started_at) * 1000)
    completed_state = {
        "job_id": job_id,
        "mode": "sync",
        "status": "completed",
        "owner_sub": user.sub,
        "created_at": int(started_at),
        "completed_at": int(time.time()),
        "elapsed_ms": elapsed_ms,
        "output_filename": output_filename,
        "mime_type": metrics.mime_type,
        "final_ref": final_ref,
        "project_id": project_used,
        "metrics": metrics.model_dump(),
        "legal_review": legal_result.model_dump() if legal_result else None,
    }
    translation_svc.save_job_state(job_id=job_id, state=completed_state)

    dlp.save_redacted_audit_log(
        session_id=job_id,
        audit_payload={
            "event": "translation_completed_sync",
            "job_id": job_id,
            "project_id": project_used,
            "elapsed_ms": elapsed_ms,
            "legal_issues_count": len(legal_result.issues) if legal_result else 0,
        },
    )

    return JSONResponse(
        status_code=status.HTTP_200_OK,
        content={
            "job_id": job_id,
            "mode": "sync",
            "status": "completed",
            "elapsed_ms": elapsed_ms,
            "output_filename": output_filename,
            "download_url": f"/api/jobs/{job_id}/download",
            "metrics": metrics.model_dump(),
            "legal_review": legal_result.model_dump() if legal_result else None,
        },
    )


@app.get("/api/jobs/{job_id}")
async def get_job_status(
    job_id: str,
    user: AuthenticatedUser = Depends(require_cf_access),
) -> dict[str, Any]:
    """Polls status of a translation job. Finalizes bilingual merge + legal review once async batch completes."""
    translation_svc = get_translation_service()
    legal_svc = get_legal_review_service()
    dlp = get_dlp_service()

    state = translation_svc.load_job_state(job_id)
    if not state:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Không tìm thấy phiên dịch thuật.")

    if state.get("owner_sub") and state["owner_sub"] != user.sub:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Không có quyền truy cập phiên này.")

    if state["status"] in ("completed", "failed"):
        return {
            "job_id": job_id,
            "mode": state.get("mode", "async"),
            "status": state["status"],
            "error": state.get("error"),
            "output_filename": state.get("output_filename"),
            "download_url": f"/api/jobs/{job_id}/download" if state["status"] == "completed" else None,
            "metrics": state.get("metrics"),
            "legal_review": state.get("legal_review"),
        }

    # Poll Google Cloud Translation LRO operation(s)
    batch_meta = state["batch_meta"]
    if batch_meta.get("multi_target"):
        targets_meta = batch_meta["targets"]
        check_en = translation_svc.check_batch_operation(targets_meta["en"])
        check_ko = translation_svc.check_batch_operation(targets_meta["ko"])

        if check_en.get("error") or check_ko.get("error"):
            err_msg = check_en.get("error") or check_ko.get("error")
            state["status"] = "failed"
            state["error"] = dlp.sanitize_log_text(str(err_msg))
            translation_svc.save_job_state(job_id=job_id, state=state)
            return {
                "job_id": job_id,
                "mode": "async",
                "status": "failed",
                "error": state["error"],
            }

        if not check_en["done"] or not check_ko["done"]:
            return {
                "job_id": job_id,
                "mode": "async",
                "status": "processing",
                "message": "Đang dịch tài liệu 3 ngôn ngữ (VI → EN & KO) trên Google Cloud...",
                "metrics": state.get("metrics"),
            }

        original_bytes = translation_svc.load_session_artifact(state["original_ref"])
        translated_en_bytes = check_en["translated_bytes"]
        translated_ko_bytes = check_ko["translated_bytes"]
        final_bytes = build_trilingual_document(
            original_bytes=original_bytes,
            translated_en_bytes=translated_en_bytes,
            translated_ko_bytes=translated_ko_bytes,
            extension=state["extension"],
            layout_mode=state["layout_mode"],
            source_lang=state["source_lang"],
        )
        review_target_bytes = translated_ko_bytes
        review_target_lang = "ko"
    else:
        batch_check = translation_svc.check_batch_operation(batch_meta)
        if not batch_check["done"]:
            return {
                "job_id": job_id,
                "mode": "async",
                "status": "processing",
                "message": "Đang dịch tài liệu trên Google Cloud Translation (Batch)...",
                "metrics": state.get("metrics"),
            }

        if batch_check.get("error"):
            state["status"] = "failed"
            state["error"] = dlp.sanitize_log_text(str(batch_check["error"]))
            translation_svc.save_job_state(job_id=job_id, state=state)
            return {
                "job_id": job_id,
                "mode": "async",
                "status": "failed",
                "error": state["error"],
            }

        original_bytes = translation_svc.load_session_artifact(state["original_ref"])
        translated_raw_bytes = batch_check["translated_bytes"]
        eff_tgt = state.get("effective_target_lang") or state["target_lang"]
        final_bytes = build_bilingual_document(
            original_bytes=original_bytes,
            translated_bytes=translated_raw_bytes,
            extension=state["extension"],
            layout_mode=state["layout_mode"],
            source_lang=state["source_lang"],
            target_lang=eff_tgt,
            replace_col_index=state.get("replace_col_index", 1),
        )
        review_target_bytes = translated_raw_bytes
        review_target_lang = eff_tgt

    legal_result: Optional[LegalReviewResult] = None
    if state.get("enable_legal_review"):
        final_bytes, legal_result = await legal_svc.review_and_annotate(
            original_bytes=original_bytes,
            translated_bytes=review_target_bytes,
            merged_bilingual_bytes=final_bytes,
            extension=state["extension"],
            source_lang=state["source_lang"],
            target_lang=review_target_lang,
        )

    final_ref = translation_svc.store_session_artifact(
        session_id=job_id,
        stage="final",
        filename=state["output_filename"],
        content=final_bytes,
        mime_type=state["mime_type"],
        bucket_name=state["batch_meta"].get("bucket"),
    )

    state["status"] = "completed"
    state["completed_at"] = int(time.time())
    state["final_ref"] = final_ref
    state["legal_review"] = legal_result.model_dump() if legal_result else None
    translation_svc.save_job_state(job_id=job_id, state=state)

    dlp.save_redacted_audit_log(
        session_id=job_id,
        audit_payload={
            "event": "translation_completed_async",
            "job_id": job_id,
            "project_id": state.get("project_id"),
            "legal_issues_count": len(legal_result.issues) if legal_result else 0,
        },
    )

    return {
        "job_id": job_id,
        "mode": "async",
        "status": "completed",
        "output_filename": state["output_filename"],
        "download_url": f"/api/jobs/{job_id}/download",
        "metrics": state.get("metrics"),
        "legal_review": state["legal_review"],
    }


@app.get("/api/jobs/{job_id}/download")
async def download_translated_contract(
    job_id: str,
    user: AuthenticatedUser = Depends(require_cf_access),
) -> StreamingResponse:
    """Streams the 100% genuine, unredacted bilingual contract file to the authenticated user."""
    translation_svc = get_translation_service()
    state = translation_svc.load_job_state(job_id)
    if not state or state.get("status") != "completed" or not state.get("final_ref"):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Bản dịch chưa hoàn tất hoặc đã hết hạn lưu trữ (7 ngày).",
        )

    if state.get("owner_sub") and state["owner_sub"] != user.sub:
        raise HTTPException(status_code=status.HTTP_403_FORBIDDEN, detail="Không có quyền tải file này.")

    content_bytes = translation_svc.load_session_artifact(state["final_ref"])
    filename = state.get("output_filename", f"bilingual_{job_id}.docx")
    encoded_filename = quote(filename)
    mime_type = state.get("mime_type", MIME_DOCX if filename.endswith(".docx") else MIME_XLSX)

    return StreamingResponse(
        io.BytesIO(content_bytes),
        media_type=mime_type,
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",
            "Cache-Control": "private, no-store",
        },
    )


@app.post("/api/ocr")
async def ocr_only_endpoint(
    file: UploadFile = File(..., description="File PDF hoặc ảnh (.pdf, .png, .jpg, .jpeg)"),
    user: AuthenticatedUser = Depends(require_cf_access),
) -> JSONResponse:
    """Standalone OCR endpoint: extracts text/tables from PDF or image into .docx without translating."""
    dlp = get_dlp_service()
    translation_svc = get_translation_service()
    ocr_svc = get_vision_ocr_service()

    raw_filename = file.filename or "document.pdf"
    ext = Path(raw_filename).suffix.lower()
    if ext not in (".pdf", ".png", ".jpg", ".jpeg"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Chỉ hỗ trợ OCR file PDF (.pdf) hoặc ảnh (.png, .jpg, .jpeg).",
        )

    file_bytes = await file.read()
    if not file_bytes:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="File tải lên rỗng.")
    if len(file_bytes) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"Dung lượng file vượt ngưỡng tối đa {MAX_UPLOAD_MB}MB.",
        )

    job_id = str(uuid.uuid4())
    started_at = time.time()

    if ext == ".pdf":
        extraction = ocr_svc.extract_pdf(file_bytes=file_bytes, filename=raw_filename)
        orig_mime = MIME_PDF
    else:
        extraction = ocr_svc.extract_image(file_bytes=file_bytes, filename=raw_filename)
        orig_mime = "image/png" if ext == ".png" else "image/jpeg"

    docx_bytes = extraction.to_docx_bytes()
    plain_text = extraction.to_plain_text()

    stem = Path(raw_filename).stem or "document"
    clean_stem = "".join(ch if ch.isalnum() or ch in ("-", "_", ".", " ") else "_" for ch in stem).strip()
    output_filename = f"{clean_stem}_ocr.docx"

    safe_audit = dlp.sanitize_dict(
        {
            "event": "ocr_only_requested",
            "job_id": job_id,
            "user_email": user.email,
            "original_filename": raw_filename,
            "pdf_type": extraction.pdf_type,
            "page_count": extraction.page_count,
            "char_count": extraction.char_count,
            "ocr_vision_requests": extraction.vision_requests_used,
            "timestamp": int(started_at),
        }
    )
    logger.info("OCR Job %s completed: %s", job_id, safe_audit)
    dlp.save_redacted_audit_log(session_id=job_id, audit_payload=safe_audit)

    translation_svc.store_session_artifact(
        session_id=job_id,
        stage="original",
        filename=raw_filename,
        content=file_bytes,
        mime_type=orig_mime,
    )
    final_ref = translation_svc.store_session_artifact(
        session_id=job_id,
        stage="final",
        filename=output_filename,
        content=docx_bytes,
        mime_type=MIME_DOCX,
    )

    elapsed_ms = int((time.time() - started_at) * 1000)
    metrics_payload = {
        "filename": raw_filename,
        "pdf_type": extraction.pdf_type,
        "estimated_pages": extraction.page_count,
        "char_count": extraction.char_count,
        "ocr_vision_requests": extraction.vision_requests_used,
    }
    completed_state = {
        "job_id": job_id,
        "mode": "ocr",
        "status": "completed",
        "owner_sub": user.sub,
        "created_at": int(started_at),
        "completed_at": int(time.time()),
        "elapsed_ms": elapsed_ms,
        "output_filename": output_filename,
        "mime_type": MIME_DOCX,
        "final_ref": final_ref,
        "metrics": metrics_payload,
    }
    translation_svc.save_job_state(job_id=job_id, state=completed_state)

    preview_limit = 6000
    return JSONResponse(
        status_code=status.HTTP_200_OK,
        content={
            "job_id": job_id,
            "mode": "ocr",
            "status": "completed",
            "elapsed_ms": elapsed_ms,
            "output_filename": output_filename,
            "download_url": f"/api/jobs/{job_id}/download",
            "metrics": metrics_payload,
            "text_preview": plain_text[:preview_limit],
            "text_truncated": len(plain_text) > preview_limit,
        },
    )
