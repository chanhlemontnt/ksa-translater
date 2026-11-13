"""Cloudflare Access Zero Trust JWT + Cloudflare Worker Shared Secret Auth.

Defense-in-depth verification for Cloud Run Backend:
1. Verifies X-Worker-Shared-Secret header injected by the Cloudflare Worker reverse proxy
   (dich.tnt.name.vn/api/* -> Cloud Run), blocking direct internet bypass to *.run.app.
2. Cryptographically verifies Cf-Access-Jwt-Assertion (RS256) against Cloudflare Access JWKS
   (https://<team>.cloudflareaccess.com/cdn-cgi/access/certs) and validates AUD + issuer + expiry.
"""

from __future__ import annotations

import hmac
import json
import logging
import os
import time
from typing import Any, Optional

import httpx
import jwt
from jwt.algorithms import RSAAlgorithm
from fastapi import HTTPException, Request, status
from pydantic import BaseModel

logger = logging.getLogger(__name__)

JWKS_CACHE_TTL_SECONDS = 3600


class AuthenticatedUser(BaseModel):
    email: str
    sub: str
    auth_mode: str = "cloudflare_access"


class _JwksKeyCache:
    """Caches Cloudflare Access public signing keys with TTL and automatic refresh on key rotation."""

    def __init__(self) -> None:
        self._keys_by_kid: dict[str, Any] = {}
        self._fetched_at: float = 0.0

    async def get_public_key(self, team_domain: str, kid: str) -> Any:
        now = time.time()
        if kid in self._keys_by_kid and (now - self._fetched_at) < JWKS_CACHE_TTL_SECONDS:
            return self._keys_by_kid[kid]

        await self._refresh(team_domain)
        if kid not in self._keys_by_kid:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Không tìm thấy khóa ký JWT hợp lệ từ Cloudflare Access (unknown kid).",
            )
        return self._keys_by_kid[kid]

    async def _refresh(self, team_domain: str) -> None:
        clean_domain = team_domain.replace("https://", "").replace("http://", "").strip("/")
        certs_url = f"https://{clean_domain}/cdn-cgi/access/certs"
        async with httpx.AsyncClient(timeout=8.0) as client:
            resp = await client.get(certs_url)
            resp.raise_for_status()
            payload = resp.json()

        new_map: dict[str, Any] = {}
        for jwk_dict in payload.get("keys", []):
            k_id = jwk_dict.get("kid")
            if k_id:
                new_map[k_id] = RSAAlgorithm.from_jwk(json.dumps(jwk_dict))

        self._keys_by_kid = new_map
        self._fetched_at = time.time()


_jwks_cache = _JwksKeyCache()


async def require_cf_access(request: Request) -> AuthenticatedUser:
    """FastAPI dependency enforcing Worker Shared Secret + Cloudflare Access JWT."""
    app_env = os.getenv("APP_ENV", "production").strip().lower()
    allow_dev_bypass = os.getenv("ALLOW_DEV_AUTH_BYPASS", "false").strip().lower() == "true"

    if app_env == "development" and allow_dev_bypass:
        dev_email = request.headers.get("X-Dev-User-Email", "dev@local.test")
        return AuthenticatedUser(email=dev_email, sub=f"dev:{dev_email}", auth_mode="dev_bypass")

    # 1. Verify Cloudflare Worker -> Cloud Run Shared Secret
    expected_secret = os.getenv("WORKER_SHARED_SECRET", "").strip()
    if not expected_secret:
        logger.error("WORKER_SHARED_SECRET is not configured in production environment.")
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Backend security configuration incomplete (missing WORKER_SHARED_SECRET).",
        )

    incoming_secret = request.headers.get("X-Worker-Shared-Secret", "")
    if not incoming_secret or not hmac.compare_digest(incoming_secret, expected_secret):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Truy cập trực tiếp bị từ chối: yêu cầu đi qua cổng dich.tnt.name.vn.",
        )

    # 2. Verify Cloudflare Access JWT Assertion
    cf_token = (
        request.headers.get("Cf-Access-Jwt-Assertion", "").strip()
        or request.cookies.get("CF_Authorization", "").strip()
    )
    if not cf_token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Thiếu chứng thực Cloudflare Access (Cf-Access-Jwt-Assertion).",
        )

    team_domain = os.getenv("CF_ACCESS_TEAM_DOMAIN", "").strip()
    expected_aud = os.getenv("CF_ACCESS_AUD", "").strip()
    if not team_domain or not expected_aud:
        logger.error("CF_ACCESS_TEAM_DOMAIN or CF_ACCESS_AUD is missing.")
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Backend security configuration incomplete (missing CF_ACCESS_TEAM_DOMAIN / CF_ACCESS_AUD).",
        )

    try:
        unverified_header = jwt.get_unverified_header(cf_token)
    except jwt.InvalidTokenError as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Header JWT không hợp lệ: {exc}",
        ) from exc

    kid = unverified_header.get("kid")
    if not kid:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="JWT thiếu trường 'kid'.",
        )

    public_key = await _jwks_cache.get_public_key(team_domain=team_domain, kid=kid)
    clean_domain = team_domain.replace("https://", "").replace("http://", "").strip("/")
    expected_issuer = f"https://{clean_domain}"
    aud_list = [a.strip() for a in expected_aud.split(",") if a.strip()]

    try:
        claims = jwt.decode(
            cf_token,
            key=public_key,
            algorithms=["RS256"],
            audience=aud_list,
            issuer=expected_issuer,
            options={"require": ["exp", "iat", "aud", "iss"]},
        )
    except jwt.ExpiredSignatureError as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Phiên đăng nhập Cloudflare Access đã hết hạn. Vui lòng tải lại trang.",
        ) from exc
    except jwt.InvalidTokenError as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=f"Xác thực Cloudflare Access thất bại: {exc}",
        ) from exc

    email = str(claims.get("email") or claims.get("common_name") or "authenticated-service")
    sub = str(claims.get("sub") or email)
    return AuthenticatedUser(email=email, sub=sub, auth_mode="cloudflare_access")
