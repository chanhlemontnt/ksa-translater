"""Cloud Translation v3 Advanced Service (Document Sync + Batch Async).

Implements:
- Metric extraction (.docx & .xlsx) to route:
  * Small contracts (<= 15 pages AND <= 20,000 chars) -> translate_document (synchronous)
  * Large contracts (> 15 pages OR > 20,000 chars)    -> batch_translate_document (async via GCS)
- Custom Glossary attachment for 3 bilingual pairs (vi-ko, en-ko, vi-en, bidirectional).
- Automatic failover across GCP projects (same billing account) via ProjectPoolRotator on HTTP 429.
- Session storage in GCS (gs://<bucket>/sessions/<session_id>/...) with 7-day lifecycle auto-cleanup,
  plus local /tmp fallback for local development.
"""

from __future__ import annotations

import io
import json
import logging
import math
import os
import time
from pathlib import Path
from typing import Any, Optional

import docx
import openpyxl
from google.api_core import exceptions as gcp_exceptions
from google.cloud import storage
from google.cloud import translate_v3 as translate
from pydantic import BaseModel

from app.dlp_service import get_dlp_service
from app.key_rotator import ProjectEntry, get_project_rotator
from app.vision_ocr_service import PdfExtractionResult, get_vision_ocr_service

logger = logging.getLogger(__name__)

SYNC_MAX_PAGES = int(os.getenv("SYNC_MAX_PAGES", "15"))
SYNC_MAX_CHARS = int(os.getenv("SYNC_MAX_CHARS", "20000"))

MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_PDF = "application/pdf"

SUPPORTED_LANG_PAIRS: set[tuple[str, str]] = {
    ("vi", "ko"),
    ("ko", "vi"),
    ("en", "ko"),
    ("ko", "en"),
    ("vi", "en"),
    ("en", "vi"),
}

LOCAL_FALLBACK_DIR = Path(os.getenv("LOCAL_SESSION_DIR", "/tmp/dich-hop-dong-sessions"))


class DocumentMetrics(BaseModel):
    filename: str
    extension: str
    mime_type: str
    size_bytes: int
    char_count: int
    estimated_pages: int
    requires_async: bool
    pdf_type: Optional[str] = None
    ocr_vision_requests: int = 0


class TranslationService:
    """Orchestrates Cloud Translation API v3 Advanced document workflows."""

    def __init__(self) -> None:
        self._translate_client: Optional[translate.TranslationServiceClient] = None
        self._storage_client: Optional[storage.Client] = None
        self._default_bucket = os.getenv("GCS_SESSION_BUCKET", "").strip()
        self._memory_jobs: dict[str, dict[str, Any]] = {}

    @property
    def translate_client(self) -> translate.TranslationServiceClient:
        if self._translate_client is None:
            self._translate_client = translate.TranslationServiceClient()
        return self._translate_client

    @property
    def storage_client(self) -> storage.Client:
        if self._storage_client is None:
            self._storage_client = storage.Client()
        return self._storage_client

    def inspect_document(self, file_bytes: bytes, filename: str) -> DocumentMetrics:
        """Extracts text length and estimates page count to choose sync vs async path."""
        ext = Path(filename).suffix.lower()
        if ext == ".docx":
            char_count, est_pages = self._inspect_docx(file_bytes)
            mime_type = MIME_DOCX
        elif ext == ".xlsx":
            char_count, est_pages = self._inspect_xlsx(file_bytes)
            mime_type = MIME_XLSX
        else:
            raise ValueError(f"Unsupported extension: {ext}")

        # Sync translateDocument hard limits in Cloud Translation v3 are 20 pages & 30K code points.
        # Buffer at <=15 pages AND <=20,000 chars per finalized architecture spec.
        requires_async = (est_pages > SYNC_MAX_PAGES) or (char_count > SYNC_MAX_CHARS)

        return DocumentMetrics(
            filename=filename,
            extension=ext,
            mime_type=mime_type,
            size_bytes=len(file_bytes),
            char_count=char_count,
            estimated_pages=max(1, est_pages),
            requires_async=requires_async,
        )

    def prepare_pdf_for_translation(
        self,
        file_bytes: bytes,
        filename: str,
    ) -> tuple[bytes, DocumentMetrics, PdfExtractionResult]:
        """Runs Auto-Detect PDF extraction / Vision OCR and converts blocks to .docx bytes."""
        ocr_service = get_vision_ocr_service()
        extraction = ocr_service.extract_pdf(file_bytes=file_bytes, filename=filename)
        docx_bytes = extraction.to_docx_bytes()

        est_pages = max(1, extraction.page_count)
        char_count = extraction.char_count
        requires_async = (est_pages > SYNC_MAX_PAGES) or (char_count > SYNC_MAX_CHARS)

        metrics = DocumentMetrics(
            filename=filename,
            extension=".docx",
            mime_type=MIME_DOCX,
            size_bytes=len(file_bytes),
            char_count=char_count,
            estimated_pages=est_pages,
            requires_async=requires_async,
            pdf_type=extraction.pdf_type,
            ocr_vision_requests=extraction.vision_requests_used,
        )
        return docx_bytes, metrics, extraction

    def _inspect_docx(self, file_bytes: bytes) -> tuple[int, int]:
        doc = docx.Document(io.BytesIO(file_bytes))
        total_chars = 0
        explicit_breaks = 0

        for para in doc.paragraphs:
            total_chars += len(para.text)
            if "w:br" in para._p.xml and 'w:type="page"' in para._p.xml:
                explicit_breaks += 1

        seen_cells: set[int] = set()
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    tc_id = id(cell._tc)
                    if tc_id not in seen_cells:
                        seen_cells.add(tc_id)
                        total_chars += len(cell.text)

        char_based_pages = max(1, math.ceil(total_chars / 2200)) if total_chars > 0 else 1
        estimated_pages = max(explicit_breaks + 1, char_based_pages)
        return total_chars, estimated_pages

    def _inspect_xlsx(self, file_bytes: bytes) -> tuple[int, int]:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        total_chars = 0
        total_rows = 0
        try:
            for sheet in wb.worksheets:
                sheet_non_empty_rows = 0
                for row in sheet.iter_rows(values_only=True):
                    row_has_text = False
                    for val in row:
                        if val is not None:
                            s = str(val).strip()
                            if s:
                                total_chars += len(s)
                                row_has_text = True
                    if row_has_text:
                        sheet_non_empty_rows += 1
                total_rows += sheet_non_empty_rows
        finally:
            wb.close()

        estimated_pages = max(1, math.ceil(total_rows / 55), math.ceil(total_chars / 2500))
        return total_chars, estimated_pages

    def translate_sync(
        self,
        session_id: str,
        file_bytes: bytes,
        metrics: DocumentMetrics,
        source_lang: str,
        target_lang: str,
        use_glossary: bool = True,
    ) -> tuple[bytes, str]:
        """Calls Cloud Translation v3 translate_document synchronously with project rotation failover."""
        rotator = get_project_rotator()
        pool_size = max(1, len(rotator.get_pool_summary()))
        last_exc: Optional[Exception] = None

        for attempt in range(pool_size):
            project: ProjectEntry = rotator.select_project(estimated_chars=metrics.char_count)
            parent = f"projects/{project.project_id}/locations/{project.location}"

            doc_input_config = translate.DocumentInputConfig(
                content=file_bytes,
                mime_type=metrics.mime_type,
            )

            glossary_config: Optional[translate.TranslateDocumentGlossaryConfig] = None
            if use_glossary:
                glossary_path = rotator.resolve_glossary_path(project, source_lang, target_lang)
                if glossary_path:
                    glossary_config = translate.TranslateDocumentGlossaryConfig(
                        glossary=glossary_path,
                        ignore_case=True,
                    )

            req = translate.TranslateDocumentRequest(
                parent=parent,
                source_language_code=source_lang,
                target_language_code=target_lang,
                document_input_config=doc_input_config,
                glossary_config=glossary_config,
                is_translate_native_pdf_only=False,
            )

            try:
                response = self.translate_client.translate_document(request=req, timeout=120.0)
                doc_trans = response.document_translation
                if use_glossary and glossary_config and response.glossary_document_translation.byte_stream_outputs:
                    translated_bytes = bytes(response.glossary_document_translation.byte_stream_outputs[0])
                else:
                    translated_bytes = bytes(doc_trans.byte_stream_outputs[0])

                rotator.record_usage(project.project_id, metrics.char_count)
                return translated_bytes, project.project_id

            except gcp_exceptions.ResourceExhausted as exc:
                logger.warning(
                    "Quota exhausted on GCP project %s (attempt %d/%d); rotating to next project.",
                    project.project_id,
                    attempt + 1,
                    pool_size,
                )
                rotator.mark_exhausted(project.project_id)
                last_exc = exc
                continue

        raise RuntimeError(
            f"All GCP projects in pool exhausted quota for sync translation: {last_exc}"
        )

    def start_batch_async(
        self,
        session_id: str,
        file_bytes: bytes,
        metrics: DocumentMetrics,
        source_lang: str,
        target_lang: str,
        use_glossary: bool = True,
    ) -> dict[str, Any]:
        """Uploads input contract to GCS and triggers batch_translate_document LRO."""
        rotator = get_project_rotator()
        project = rotator.select_project(estimated_chars=metrics.char_count)
        bucket_name = project.gcs_bucket or self._default_bucket
        if not bucket_name:
            raise RuntimeError(
                "GCS_SESSION_BUCKET (or per-project gcs_bucket) is required for batchTranslateDocument (>15 pages / >20k chars)."
            )

        safe_in_name = f"input{metrics.extension}"
        input_blob_path = f"sessions/{session_id}/batch_in/{safe_in_name}"
        output_prefix_path = f"sessions/{session_id}/batch_out/"

        bucket = self.storage_client.bucket(bucket_name)
        in_blob = bucket.blob(input_blob_path)
        in_blob.upload_from_string(file_bytes, content_type=metrics.mime_type)

        input_gcs_uri = f"gs://{bucket_name}/{input_blob_path}"
        output_gcs_prefix = f"gs://{bucket_name}/{output_prefix_path}"

        parent = f"projects/{project.project_id}/locations/{project.location}"
        batch_input = translate.BatchDocumentInputConfig(
            gcs_source=translate.GcsSource(input_uri=input_gcs_uri)
        )
        batch_output = translate.BatchDocumentOutputConfig(
            gcs_destination=translate.GcsDestination(output_uri_prefix=output_gcs_prefix)
        )

        glossaries_map: dict[str, translate.TranslateTextGlossaryConfig] = {}
        if use_glossary:
            glossary_path = rotator.resolve_glossary_path(project, source_lang, target_lang)
            if glossary_path:
                glossaries_map[target_lang] = translate.TranslateTextGlossaryConfig(
                    glossary=glossary_path,
                    ignore_case=True,
                )

        request = translate.BatchTranslateDocumentRequest(
            parent=parent,
            source_language_code=source_lang,
            target_language_codes=[target_lang],
            input_configs=[batch_input],
            output_config=batch_output,
            glossaries=glossaries_map or None,
        )

        operation = self.translate_client.batch_translate_document(request=request)
        rotator.record_usage(project.project_id, metrics.char_count)

        return {
            "operation_name": operation.operation.name,
            "project_id": project.project_id,
            "location": project.location,
            "bucket": bucket_name,
            "input_gcs_uri": input_gcs_uri,
            "output_gcs_prefix": output_gcs_prefix,
            "output_prefix_path": output_prefix_path,
            "extension": metrics.extension,
        }

    def check_batch_operation(self, operation_meta: dict[str, Any]) -> dict[str, Any]:
        """Checks Google LRO status and fetches translated document bytes from GCS when complete."""
        op_name = operation_meta["operation_name"]
        ops_client = self.translate_client._transport.operations_client
        raw_op = ops_client.get_operation(name=op_name)

        if not raw_op.done:
            return {"done": False, "error": None, "translated_bytes": None}

        if raw_op.error and raw_op.error.code != 0:
            return {
                "done": True,
                "error": f"BatchTranslate error ({raw_op.error.code}): {raw_op.error.message}",
                "translated_bytes": None,
            }

        bucket_name = operation_meta["bucket"]
        prefix = operation_meta["output_prefix_path"]
        ext = operation_meta.get("extension", ".docx")

        bucket = self.storage_client.bucket(bucket_name)
        blobs = list(bucket.list_blobs(prefix=prefix))
        matching = [
            b for b in blobs
            if b.name.lower().endswith(ext) and not b.name.endswith("/index.csv")
        ]
        if not matching:
            return {
                "done": True,
                "error": "Batch operation finished but no translated document artifact found in GCS output prefix.",
                "translated_bytes": None,
            }

        # Prefer glossary-translated output if Cloud Translation emitted both standard & glossary outputs
        glossary_blobs = [b for b in matching if "glossary" in b.name.lower()]
        chosen_blob = glossary_blobs[0] if glossary_blobs else matching[0]
        translated_bytes = chosen_blob.download_as_bytes()

        return {
            "done": True,
            "error": None,
            "translated_bytes": translated_bytes,
            "output_blob": f"gs://{bucket_name}/{chosen_blob.name}",
        }

    def store_session_artifact(
        self,
        session_id: str,
        stage: str,
        filename: str,
        content: bytes,
        mime_type: str,
        bucket_name: Optional[str] = None,
    ) -> str:
        """Saves session artifact to GCS (7-day lifecycle rule) or local fallback in dev."""
        target_bucket = bucket_name or self._default_bucket
        clean_name = Path(filename).name
        rel_key = f"sessions/{session_id}/{stage}/{clean_name}"

        if target_bucket:
            bucket = self.storage_client.bucket(target_bucket)
            blob = bucket.blob(rel_key)
            blob.upload_from_string(content, content_type=mime_type)
            return f"gs://{target_bucket}/{rel_key}"

        local_path = LOCAL_FALLBACK_DIR / rel_key
        local_path.parent.mkdir(parents=True, exist_ok=True)
        local_path.write_bytes(content)
        return f"file://{local_path}"

    def load_session_artifact(self, ref: str) -> bytes:
        """Loads session artifact bytes from GCS URI or local fallback URI."""
        if ref.startswith("gs://"):
            without_scheme = ref[len("gs://"):]
            bucket_name, blob_path = without_scheme.split("/", 1)
            bucket = self.storage_client.bucket(bucket_name)
            return bucket.blob(blob_path).download_as_bytes()
        if ref.startswith("file://"):
            return Path(ref[len("file://"):]).read_bytes()
        raise ValueError(f"Unsupported artifact reference: {ref}")

    def save_job_state(self, job_id: str, state: dict[str, Any]) -> None:
        """Persists sanitized job state JSON to GCS + in-memory cache."""
        dlp = get_dlp_service()
        # Keep structural refs intact while sanitizing any free-text message/error fields
        persisted = dict(state)
        if "error" in persisted and persisted["error"]:
            persisted["error"] = dlp.sanitize_log_text(str(persisted["error"]))
        persisted["updated_at"] = int(time.time())

        self._memory_jobs[job_id] = persisted
        raw_json = json.dumps(persisted, ensure_ascii=False).encode("utf-8")

        if self._default_bucket:
            bucket = self.storage_client.bucket(self._default_bucket)
            bucket.blob(f"sessions/{job_id}/state.json").upload_from_string(
                raw_json, content_type="application/json"
            )
        else:
            local_state = LOCAL_FALLBACK_DIR / f"sessions/{job_id}/state.json"
            local_state.parent.mkdir(parents=True, exist_ok=True)
            local_state.write_bytes(raw_json)

    def load_job_state(self, job_id: str) -> Optional[dict[str, Any]]:
        """Loads job state from memory or GCS/local persistence."""
        if self._default_bucket:
            bucket = self.storage_client.bucket(self._default_bucket)
            blob = bucket.blob(f"sessions/{job_id}/state.json")
            if blob.exists():
                data = json.loads(blob.download_as_bytes().decode("utf-8"))
                self._memory_jobs[job_id] = data
                return data
        else:
            local_state = LOCAL_FALLBACK_DIR / f"sessions/{job_id}/state.json"
            if local_state.exists():
                data = json.loads(local_state.read_bytes().decode("utf-8"))
                self._memory_jobs[job_id] = data
                return data
        return self._memory_jobs.get(job_id)


_translation_service_singleton: Optional[TranslationService] = None


def get_translation_service() -> TranslationService:
    global _translation_service_singleton
    if _translation_service_singleton is None:
        _translation_service_singleton = TranslationService()
    return _translation_service_singleton
