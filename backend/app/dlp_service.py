"""Cloud DLP Data-at-Rest & Audit Log Sanitization Service.

CRITICAL ARCHITECTURAL INVARIANT:
- Cloud DLP redaction applies EXCLUSIVELY to system logs, metadata, and GCS audit records
  (data-at-rest observability).
- NEVER pass user-downloadable contract files (.docx / .xlsx) through this service!
  Downloaded contracts must retain 100% original legal names, tax IDs, bank accounts, and amounts.
"""

from __future__ import annotations

import json
import logging
import os
import re
import time
from typing import Any, Optional

from google.cloud import dlp_v2
from google.cloud import storage

from app.key_rotator import get_project_rotator

logger = logging.getLogger(__name__)

# Deterministic local regex patterns as defense-in-depth (and zero-cost fallback)
_LOCAL_REDACTION_PATTERNS: list[tuple[re.Pattern[str], str]] = [
    # Email addresses
    (re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"), "[REDACTED_EMAIL]"),
    # Vietnam CCCD (12 digits) / CMND (9 digits) & Korean RRN (6-7 digits)
    (re.compile(r"\b\d{6}-[1-4]\d{6}\b"), "[REDACTED_KR_ID]"),
    (re.compile(r"\b0\d{11}\b"), "[REDACTED_VN_CCCD]"),
    # Vietnam Tax ID (10 or 10-3 digits) & Korean Business Reg No (XXX-XX-XXXXX)
    (re.compile(r"\b\d{3}-\d{2}-\d{5}\b"), "[REDACTED_KR_BIZ_ID]"),
    (re.compile(r"\b\d{10}-\d{3}\b"), "[REDACTED_VN_MST]"),
    # Phone numbers (VN +84 / 09x / 03x & KR +82 / 010-xxxx-xxxx)
    (re.compile(r"(?:\+?84|0)(?:3|5|7|8|9)\d{8}\b"), "[REDACTED_PHONE]"),
    (re.compile(r"(?:\+?82[-\s]?|0)1[016789][-\s]?\d{3,4}[-\s]?\d{4}\b"), "[REDACTED_KR_PHONE]"),
    # Monetary amounts with currency markers (VND, USD, KRW, 원, VNĐ)
    (
        re.compile(
            r"\b\d{1,3}(?:[.,]\d{3})+(?:\.\d+)?\s*(?:VND|VNĐ|USD|KRW|원|동)\b",
            re.IGNORECASE,
        ),
        "[REDACTED_AMOUNT]",
    ),
    # Bank account numbers near banking keywords
    (
        re.compile(
            r"(?i)(số tài khoản|stk|account\s*no\.?|계좌번호)\s*[:\-]?\s*([0-9\-]{8,20})"
        ),
        r"\1: [REDACTED_BANK_ACCT]",
    ),
]


class DlpProtectionService:
    """Inspects and redacts sensitive contract PII in logs and stored audit records."""

    def __init__(self) -> None:
        self._enabled = os.getenv("DLP_ENABLED", "true").lower() == "true"
        self._dlp_client: Optional[dlp_v2.DlpServiceClient] = None
        self._storage_client: Optional[storage.Client] = None
        self._default_bucket = os.getenv("GCS_SESSION_BUCKET", "").strip()

    @property
    def dlp_client(self) -> dlp_v2.DlpServiceClient:
        if self._dlp_client is None:
            self._dlp_client = dlp_v2.DlpServiceClient()
        return self._dlp_client

    @property
    def storage_client(self) -> storage.Client:
        if self._storage_client is None:
            self._storage_client = storage.Client()
        return self._storage_client

    @staticmethod
    def _apply_local_regex_mask(text: str) -> str:
        masked = text
        for pattern, replacement in _LOCAL_REDACTION_PATTERNS:
            masked = pattern.sub(replacement, masked)
        return masked

    def sanitize_log_text(self, text: str) -> str:
        """Redacts sensitive PII from log/audit strings using Cloud DLP + local regex fallback."""
        if not text:
            return text

        # Always run fast local regex redaction first
        pre_masked = self._apply_local_regex_mask(text)
        if not self._enabled or len(pre_masked.strip()) < 4:
            return pre_masked

        try:
            rotator = get_project_rotator()
            summary = rotator.get_pool_summary()
            project_id = (
                summary[0]["project_id"]
                if summary
                else os.getenv("GCP_PROJECT_ID", "").strip()
            )
            if not project_id:
                return pre_masked

            parent = f"projects/{project_id}/locations/global"

            inspect_config = {
                "info_types": [
                    {"name": "PERSON_NAME"},
                    {"name": "EMAIL_ADDRESS"},
                    {"name": "PHONE_NUMBER"},
                    {"name": "STREET_ADDRESS"},
                    {"name": "PASSPORT"},
                    {"name": "GOVERNMENT_ID"},
                    {"name": "IBAN_CODE"},
                    {"name": "FINANCIAL_ACCOUNT_NUMBER"},
                ],
                "custom_info_types": [
                    {
                        "info_type": {"name": "CONTRACT_PENALTY_OR_VALUE"},
                        "regex": {
                            "pattern": r"(?i)(phạt vi phạm|위약금|손해배상|liquidated damages|giá trị hợp đồng)[^.\n]{0,80}"
                        },
                    },
                    {
                        "info_type": {"name": "VN_IDENTITY_OR_TAX_ID"},
                        "regex": {"pattern": r"\b\d{9,13}\b"},
                    },
                ],
                "min_likelihood": dlp_v2.Likelihood.POSSIBLE,
            }

            deidentify_config = {
                "info_type_transformations": {
                    "transformations": [
                        {
                            "primitive_transformation": {
                                "replace_with_info_type_config": {}
                            }
                        }
                    ]
                }
            }

            response = self.dlp_client.deidentify_content(
                request={
                    "parent": parent,
                    "inspect_config": inspect_config,
                    "deidentify_config": deidentify_config,
                    "item": {"value": pre_masked[:30000]},
                },
                timeout=8.0,
            )
            return response.item.value
        except Exception as exc:
            logger.debug("Cloud DLP fallback to local regex mask (%s)", type(exc).__name__)
            return pre_masked

    def sanitize_dict(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Sanitizes string values in a metadata/audit dictionary while keeping non-PII IDs."""
        passthrough_keys = {
            "event",
            "job_id",
            "source_lang",
            "target_lang",
            "layout_mode",
            "project_id",
            "status",
            "mode",
        }
        sanitized: dict[str, Any] = {}
        for k, v in payload.items():
            if k in passthrough_keys or isinstance(v, (bool, int, float)) or v is None:
                sanitized[k] = v
            elif isinstance(v, str):
                sanitized[k] = self.sanitize_log_text(v)
            elif isinstance(v, dict):
                sanitized[k] = self.sanitize_dict(v)
            elif isinstance(v, list):
                sanitized[k] = [
                    self.sanitize_log_text(item) if isinstance(item, str) else item
                    for item in v
                ]
            else:
                sanitized[k] = self.sanitize_log_text(str(v))
        return sanitized

    def save_redacted_audit_log(
        self,
        session_id: str,
        audit_payload: dict[str, Any],
        bucket_name: Optional[str] = None,
    ) -> Optional[str]:
        """Writes DLP-redacted audit record to GCS (auto-deleted after 7 days by lifecycle rule)."""
        redacted = self.sanitize_dict(audit_payload)
        redacted["recorded_at"] = int(time.time())
        target_bucket = bucket_name or self._default_bucket
        if not target_bucket:
            return None

        try:
            ts_ms = int(time.time() * 1000)
            blob_key = f"sessions/{session_id}/audit/{ts_ms}_{redacted.get('event', 'log')}.json"
            bucket = self.storage_client.bucket(target_bucket)
            bucket.blob(blob_key).upload_from_string(
                json.dumps(redacted, ensure_ascii=False, indent=2).encode("utf-8"),
                content_type="application/json",
            )
            return f"gs://{target_bucket}/{blob_key}"
        except Exception as exc:
            logger.warning("Could not persist redacted audit log to GCS: %s", type(exc).__name__)
            return None


_dlp_singleton: Optional[DlpProtectionService] = None


def get_dlp_service() -> DlpProtectionService:
    global _dlp_singleton
    if _dlp_singleton is None:
        _dlp_singleton = DlpProtectionService()
    return _dlp_singleton
