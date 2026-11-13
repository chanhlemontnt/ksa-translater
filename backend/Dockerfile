# Production Container for Google Cloud Run (Always Free Tier Optimized)
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=8080

WORKDIR /app

# Install minimal native dependencies required by lxml / cryptography
RUN apt-get update && apt-get install -y --no-install-recommends \
    libxml2 \
    libxslt1.1 \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Run as non-root user for Zero-Trust hardening
RUN groupadd --gid 10001 appgroup && \
    useradd --uid 10001 --gid appgroup --shell /sbin/nologin --create-home appuser

COPY requirements.txt ./
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY app ./app
RUN mkdir -p /tmp/dich-hop-dong-sessions && chown -R appuser:appgroup /app /tmp/dich-hop-dong-sessions

USER appuser

EXPOSE 8080

# Single worker with async event loop scales cleanly to zero on Cloud Run Always Free
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8080} --proxy-headers --forwarded-allow-ips=*"]
