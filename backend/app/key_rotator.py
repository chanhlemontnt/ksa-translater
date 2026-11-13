"""GCP Project Pool Rotator (Single Real Billing Account, Multi-Project Quota Pooling).

Architecture compliance:
- Rotates across multiple Google Cloud Projects owned by ONE real Google Account / Billing Account
  (100% compliant with Google Cloud ToS; Cloud Run service account is granted
  roles/cloudtranslate.user on each project in the pool via IAM).
- NEVER hardcodes credentials, API keys, or real project IDs in source code.
- Reads pool configuration from:
  1) GCP_PROJECT_POOL_JSON (JSON array string in env / Cloud Run Secret Manager)
  2) GCP_PROJECT_POOL_FILE (path to mounted JSON file)
  3) Fallback single project from GCP_PROJECT_ID / GOOGLE_CLOUD_PROJECT env var.
- Tracks estimated monthly character consumption per project (default soft cap 480,000 chars
  below the 500,000 chars/month free tier threshold) and handles cooldowns on HTTP 429.
"""

from __future__ import annotations

import datetime as dt
import json
import logging
import os
import threading
import time
from pathlib import Path
from typing import Any, Optional

from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)

DEFAULT_MONTHLY_SOFT_LIMIT = int(os.getenv("GCP_PROJECT_MONTHLY_CHAR_LIMIT", "480000"))
DEFAULT_VISION_MONTHLY_LIMIT = int(os.getenv("GCP_VISION_MONTHLY_LIMIT", "950"))
DEFAULT_LOCATION = os.getenv("GCP_TRANSLATE_LOCATION", "us-central1")

# Canonical glossary IDs per bidirectional language pair if not explicitly overridden per project
CANONICAL_GLOSSARY_IDS: dict[str, str] = {
    "vi-ko": "legal-glossary-vi-ko",
    "ko-vi": "legal-glossary-vi-ko",
    "en-ko": "legal-glossary-en-ko",
    "ko-en": "legal-glossary-en-ko",
    "vi-en": "legal-glossary-vi-en",
    "en-vi": "legal-glossary-vi-en",
}


class ProjectEntry(BaseModel):
    """Configuration & runtime state for a single GCP project in the translation pool."""

    project_id: str
    location: str = DEFAULT_LOCATION
    gcs_bucket: Optional[str] = None
    glossaries: dict[str, str] = Field(default_factory=dict)
    enable_default_glossary_naming: bool = False
    monthly_char_limit: int = DEFAULT_MONTHLY_SOFT_LIMIT
    used_chars: int = 0
    usage_month: str = Field(default_factory=lambda: dt.datetime.now(dt.timezone.utc).strftime("%Y-%m"))
    cooldown_until: float = 0.0
    enabled: bool = True
    # Google Cloud Vision API quota tracking (Always Free: 1,000 requests/month per project)
    vision_monthly_limit: int = DEFAULT_VISION_MONTHLY_LIMIT
    vision_used_requests: int = 0
    vision_usage_month: str = Field(default_factory=lambda: dt.datetime.now(dt.timezone.utc).strftime("%Y-%m"))
    vision_cooldown_until: float = 0.0
    vision_enabled: bool = True
    vision_credentials_file: Optional[str] = None


class ProjectPoolRotator:
    """Thread-safe round-robin rotator across GCP projects sharing one real billing account."""

    def __init__(self, projects: Optional[list[ProjectEntry]] = None) -> None:
        self._lock = threading.Lock()
        self._projects: list[ProjectEntry] = projects if projects is not None else self._load_from_env()
        self._cursor: int = 0
        self._vision_cursor: int = 0
        self._vision_clients: dict[str, Any] = {}

    @staticmethod
    def _current_utc_month() -> str:
        return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m")

    @classmethod
    def _load_from_env(cls) -> list[ProjectEntry]:
        raw_json = os.getenv("GCP_PROJECT_POOL_JSON", "").strip()
        pool_file = os.getenv("GCP_PROJECT_POOL_FILE", "").strip()

        parsed_items: list[dict[str, Any]] = []
        if raw_json:
            data = json.loads(raw_json)
            if isinstance(data, dict) and "projects" in data:
                parsed_items = list(data["projects"])
            elif isinstance(data, list):
                parsed_items = data
            else:
                raise ValueError("GCP_PROJECT_POOL_JSON must be a JSON list or {'projects': [...]}.")
        elif pool_file and Path(pool_file).exists():
            data = json.loads(Path(pool_file).read_text(encoding="utf-8"))
            parsed_items = data["projects"] if isinstance(data, dict) and "projects" in data else list(data)
        else:
            fallback_pid = (
                os.getenv("GCP_PROJECT_ID", "").strip()
                or os.getenv("GOOGLE_CLOUD_PROJECT", "").strip()
            )
            if fallback_pid:
                parsed_items = [
                    {
                        "project_id": fallback_pid,
                        "location": DEFAULT_LOCATION,
                        "gcs_bucket": os.getenv("GCS_SESSION_BUCKET", "").strip() or None,
                        "enable_default_glossary_naming": os.getenv("USE_DEFAULT_GLOSSARY_IDS", "false").lower()
                        == "true",
                    }
                ]

        entries = [ProjectEntry(**item) for item in parsed_items if item.get("project_id")]
        if not entries:
            logger.warning(
                "No GCP projects configured in GCP_PROJECT_POOL_JSON / GCP_PROJECT_ID yet. "
                "Set environment variables before invoking Cloud Translation."
            )
        return entries

    def _reset_month_if_needed(self, entry: ProjectEntry) -> None:
        curr_month = self._current_utc_month()
        if entry.usage_month != curr_month:
            entry.usage_month = curr_month
            entry.used_chars = 0
            entry.cooldown_until = 0.0
        if entry.vision_usage_month != curr_month:
            entry.vision_usage_month = curr_month
            entry.vision_used_requests = 0
            entry.vision_cooldown_until = 0.0

    def select_project(self, estimated_chars: int = 0) -> ProjectEntry:
        """Selects the next healthy GCP project with sufficient monthly quota buffer."""
        with self._lock:
            if not self._projects:
                raise RuntimeError(
                    "GCP project pool is empty. Configure GCP_PROJECT_POOL_JSON or GCP_PROJECT_ID."
                )

            now = time.time()
            n = len(self._projects)

            # Pass 1: Active, not in cooldown, and within monthly soft character limit
            for offset in range(n):
                idx = (self._cursor + offset) % n
                candidate = self._projects[idx]
                self._reset_month_if_needed(candidate)
                if not candidate.enabled or candidate.cooldown_until > now:
                    continue
                if candidate.used_chars + estimated_chars <= candidate.monthly_char_limit:
                    self._cursor = (idx + 1) % n
                    return candidate.model_copy(deep=True)

            # Pass 2: Active and not in 429 cooldown (even if soft char estimate exceeded)
            for offset in range(n):
                idx = (self._cursor + offset) % n
                candidate = self._projects[idx]
                if candidate.enabled and candidate.cooldown_until <= now:
                    logger.warning(
                        "All projects reached soft monthly limit; selecting %s as overflow.",
                        candidate.project_id,
                    )
                    self._cursor = (idx + 1) % n
                    return candidate.model_copy(deep=True)

            # Pass 3: Pick project with earliest cooldown expiry as last resort
            active_projects = [p for p in self._projects if p.enabled]
            if not active_projects:
                raise RuntimeError("All configured GCP projects are disabled.")
            earliest = min(active_projects, key=lambda p: p.cooldown_until)
            return earliest.model_copy(deep=True)

    def record_usage(self, project_id: str, chars: int) -> None:
        """Increments in-memory monthly character counter for observability & rotation."""
        if chars <= 0:
            return
        with self._lock:
            for entry in self._projects:
                if entry.project_id == project_id:
                    self._reset_month_if_needed(entry)
                    entry.used_chars += chars
                    break

    def mark_exhausted(self, project_id: str, cooldown_seconds: int = 3600) -> None:
        """Places a project into temporary cooldown upon receiving HTTP 429 ResourceExhausted."""
        with self._lock:
            for entry in self._projects:
                if entry.project_id == project_id:
                    entry.cooldown_until = time.time() + cooldown_seconds
                    break

    def get_bucket_for_project(self, project_id: str) -> Optional[str]:
        with self._lock:
            for entry in self._projects:
                if entry.project_id == project_id:
                    return entry.gcs_bucket or os.getenv("GCS_SESSION_BUCKET", "").strip() or None
        return os.getenv("GCS_SESSION_BUCKET", "").strip() or None

    def resolve_glossary_path(
        self,
        project: ProjectEntry,
        source_lang: str,
        target_lang: str,
    ) -> Optional[str]:
        """Resolves full Cloud Translation v3 glossary resource name for a language pair."""
        pair_key = f"{source_lang}-{target_lang}"
        reverse_key = f"{target_lang}-{source_lang}"

        explicit = project.glossaries.get(pair_key) or project.glossaries.get(reverse_key)
        if explicit:
            if explicit.startswith("projects/"):
                return explicit
            return f"projects/{project.project_id}/locations/{project.location}/glossaries/{explicit}"

        if project.enable_default_glossary_naming and pair_key in CANONICAL_GLOSSARY_IDS:
            gid = CANONICAL_GLOSSARY_IDS[pair_key]
            return f"projects/{project.project_id}/locations/{project.location}/glossaries/{gid}"

        return None

    # ------------------------------------------------------------------
    # Google Cloud Vision API Client Rotation (1,000 req/month free/project)
    # ------------------------------------------------------------------

    def select_vision_project(self, estimated_requests: int = 1) -> ProjectEntry:
        """Selects the next healthy GCP project with available Vision API monthly free quota."""
        with self._lock:
            if not self._projects:
                raise RuntimeError(
                    "GCP project pool is empty. Configure GCP_PROJECT_POOL_JSON or GCP_PROJECT_ID."
                )

            now = time.time()
            n = len(self._projects)

            # Pass 1: Active, Vision-enabled, not in cooldown, within soft monthly request limit
            for offset in range(n):
                idx = (self._vision_cursor + offset) % n
                candidate = self._projects[idx]
                self._reset_month_if_needed(candidate)
                if (
                    not candidate.enabled
                    or not candidate.vision_enabled
                    or candidate.vision_cooldown_until > now
                ):
                    continue
                if candidate.vision_used_requests + estimated_requests <= candidate.vision_monthly_limit:
                    self._vision_cursor = (idx + 1) % n
                    return candidate.model_copy(deep=True)

            # Pass 2: Active, Vision-enabled, not in cooldown (overflow past soft limit)
            for offset in range(n):
                idx = (self._vision_cursor + offset) % n
                candidate = self._projects[idx]
                if candidate.enabled and candidate.vision_enabled and candidate.vision_cooldown_until <= now:
                    logger.warning(
                        "All projects reached Vision soft monthly limit; selecting %s as overflow.",
                        candidate.project_id,
                    )
                    self._vision_cursor = (idx + 1) % n
                    return candidate.model_copy(deep=True)

            # Pass 3: Earliest Vision cooldown expiry as last resort
            active_projects = [p for p in self._projects if p.enabled and p.vision_enabled]
            if not active_projects:
                raise RuntimeError("All configured GCP projects have Vision API disabled.")
            earliest = min(active_projects, key=lambda p: p.vision_cooldown_until)
            return earliest.model_copy(deep=True)

    def get_vision_client(self, estimated_requests: int = 1) -> tuple[Any, ProjectEntry]:
        """Allocates a Cloud Vision ImageAnnotatorClient bound to the rotated GCP project."""
        from google.api_core.client_options import ClientOptions
        from google.cloud import vision

        project = self.select_vision_project(estimated_requests=estimated_requests)
        with self._lock:
            cached = self._vision_clients.get(project.project_id)
            if cached is not None:
                return cached, project

            client_options = ClientOptions(quota_project_id=project.project_id)
            if project.vision_credentials_file and Path(project.vision_credentials_file).exists():
                from google.oauth2 import service_account

                creds = service_account.Credentials.from_service_account_file(
                    project.vision_credentials_file
                )
                client = vision.ImageAnnotatorClient(
                    credentials=creds,
                    client_options=client_options,
                )
            else:
                client = vision.ImageAnnotatorClient(client_options=client_options)

            self._vision_clients[project.project_id] = client
            return client, project

    def record_vision_usage(self, project_id: str, requests_count: int = 1) -> None:
        """Increments in-memory monthly Vision API request counter for a GCP project."""
        if requests_count <= 0:
            return
        with self._lock:
            for entry in self._projects:
                if entry.project_id == project_id:
                    self._reset_month_if_needed(entry)
                    entry.vision_used_requests += requests_count
                    break

    def mark_vision_exhausted(self, project_id: str, cooldown_seconds: int = 3600) -> None:
        """Places a project's Vision API into temporary cooldown upon HTTP 429 ResourceExhausted."""
        with self._lock:
            for entry in self._projects:
                if entry.project_id == project_id:
                    entry.vision_cooldown_until = time.time() + cooldown_seconds
                    break

    def get_pool_summary(self) -> list[dict[str, Any]]:
        """Returns non-sensitive pool telemetry (never exposes secrets)."""
        now = time.time()
        with self._lock:
            summary = []
            for entry in self._projects:
                self._reset_month_if_needed(entry)
                summary.append(
                    {
                        "project_id": entry.project_id,
                        "location": entry.location,
                        "enabled": entry.enabled,
                        "in_cooldown": entry.cooldown_until > now,
                        "used_chars": entry.used_chars,
                        "monthly_char_limit": entry.monthly_char_limit,
                        "usage_month": entry.usage_month,
                        "vision_enabled": entry.vision_enabled,
                        "vision_in_cooldown": entry.vision_cooldown_until > now,
                        "vision_used_requests": entry.vision_used_requests,
                        "vision_monthly_limit": entry.vision_monthly_limit,
                        "vision_usage_month": entry.vision_usage_month,
                    }
                )
            return summary


_rotator_singleton: Optional[ProjectPoolRotator] = None


def get_project_rotator() -> ProjectPoolRotator:
    global _rotator_singleton
    if _rotator_singleton is None:
        _rotator_singleton = ProjectPoolRotator()
    return _rotator_singleton
