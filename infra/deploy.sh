#!/usr/bin/env bash
# ==============================================================================
# Script Triển Khai Hạ Tầng Serverless: Web App Dịch Hợp Đồng Song Ngữ
# Stack: Google Cloud Run + GCS (Lifecycle 7 ngày) + Cloud Translation v3 + DLP
# Lưu ý: Toàn bộ các Project trong POOL_PROJECT_IDS PHẢI thuộc CÙNG 1 Google
#        Account / Billing Account thật của anh Thanh (tuân thủ 100% Google ToS).
# ==============================================================================
set -euo pipefail

# ------------------------------------------------------------------------------
# [BƯỚC CẦN ĐIỀN THẬT] — Anh Thanh cập nhật các biến dưới đây trước khi chạy
# ------------------------------------------------------------------------------
PRIMARY_PROJECT_ID="REPLACE_WITH_PRIMARY_GCP_PROJECT_ID"      # VD: tnt-dich-hopdong-01
# Danh sách các GCP Project xoay vòng quota (cách nhau dấu cách, cùng 1 Billing thật).
# Tuần đầu test thực tế 1 project để theo dõi Cloud Billing report trước khi mở rộng:
POOL_PROJECT_IDS=("${PRIMARY_PROJECT_ID}")                    # VD: ("tnt-dich-01" "tnt-dich-02" "tnt-dich-03")

REGION="us-central1"                                          # Region thuộc gói Always Free của Cloud Run & GCS
GCS_BUCKET_NAME="${PRIMARY_PROJECT_ID}-contract-sessions"     # Tên bucket duy nhất toàn cầu
SERVICE_NAME="dich-hop-dong-api"
SA_NAME="dich-hop-dong-runner"                                # Service Account chạy Cloud Run
SA_EMAIL="${SA_NAME}@${PRIMARY_PROJECT_ID}.iam.gserviceaccount.com"

# Cấu hình Bảo mật Cloudflare Access & Cloudflare Worker (dich.tnt.name.vn)
CF_ACCESS_TEAM_DOMAIN="REPLACE_WITH_TEAM.cloudflareaccess.com"
CF_ACCESS_AUD="REPLACE_WITH_CLOUDFLARE_ACCESS_AUD_TAG"
WORKER_SHARED_SECRET="REPLACE_WITH_STRONG_32B_HEX_SECRET"     # Tạo bằng: openssl rand -hex 32

# (Tuỳ chọn) API Key cho AI Legal Review — để rỗng nếu chưa bật
LEGAL_REVIEW_PROVIDER="anthropic"                             # anthropic | gemini
ANTHROPIC_API_KEY="${ANTHROPIC_API_KEY:-}"
GEMINI_API_KEY="${GEMINI_API_KEY:-}"
# ------------------------------------------------------------------------------

if [[ "${PRIMARY_PROJECT_ID}" == "REPLACE_"* ]]; then
  echo "❌ Vui lòng mở infra/deploy.sh và điền PRIMARY_PROJECT_ID, CF_ACCESS_TEAM_DOMAIN, CF_ACCESS_AUD, WORKER_SHARED_SECRET thật trước khi chạy!"
  exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

echo "=== [1/6] Bật API cần thiết trên Primary Project (${PRIMARY_PROJECT_ID}) ==="
gcloud config set project "${PRIMARY_PROJECT_ID}"
gcloud services enable \
  run.googleapis.com \
  cloudbuild.googleapis.com \
  artifactregistry.googleapis.com \
  translate.googleapis.com \
  storage.googleapis.com \
  dlp.googleapis.com

echo "=== [2/6] Khởi tạo Service Account (${SA_EMAIL}) ==="
if ! gcloud iam service-accounts describe "${SA_EMAIL}" >/dev/null 2>&1; then
  gcloud iam service-accounts create "${SA_NAME}" \
    --display-name="Cloud Run Bilingual Contract Translator SA"
fi

echo "=== [3/6] Bật Translation API & Cấp quyền IAM cho các Project trong Pool ==="
for PID in "${POOL_PROJECT_IDS[@]}"; do
  echo "  -> Cấu hình project pool: ${PID}"
  gcloud services enable translate.googleapis.com dlp.googleapis.com --project="${PID}"
  gcloud projects add-iam-policy-binding "${PID}" \
    --member="serviceAccount:${SA_EMAIL}" \
    --role="roles/cloudtranslate.user" >/dev/null
  gcloud projects add-iam-policy-binding "${PID}" \
    --member="serviceAccount:${SA_EMAIL}" \
    --role="roles/dlp.user" >/dev/null
done

echo "=== [4/6] Tạo GCS Bucket (${GCS_BUCKET_NAME}) & Gắn Lifecycle Tự Xoá Sau 7 Ngày ==="
if ! gcloud storage buckets describe "gs://${GCS_BUCKET_NAME}" >/dev/null 2>&1; then
  gcloud storage buckets create "gs://${GCS_BUCKET_NAME}" \
    --project="${PRIMARY_PROJECT_ID}" \
    --location="${REGION}" \
    --uniform-bucket-level-access
fi

# Cấp quyền đọc/ghi bucket cho Cloud Run Service Account
gcloud storage buckets add-iam-policy-binding "gs://${GCS_BUCKET_NAME}" \
  --member="serviceAccount:${SA_EMAIL}" \
  --role="roles/storage.objectAdmin" >/dev/null

# Áp dụng chính sách tự động xoá mọi file phiên dịch thuật (sessions/) sau 7 ngày
LIFECYCLE_FILE="$(mktemp)"
cat > "${LIFECYCLE_FILE}" <<EOF
{
  "rule": [
    {
      "action": { "type": "Delete" },
      "condition": {
        "age": 7,
        "matchesPrefix": ["sessions/"]
      }
    }
  ]
}
EOF
gcloud storage buckets update "gs://${GCS_BUCKET_NAME}" --lifecycle-file="${LIFECYCLE_FILE}"
rm -f "${LIFECYCLE_FILE}"
echo "  -> Đã kích hoạt GCS Lifecycle tự động xoá dữ liệu phiên sau 7 ngày."

echo "=== [5/6] Upload Từ điển Thuật ngữ Pháp lý mẫu lên GCS ==="
if [[ -d "${SCRIPT_DIR}/glossaries" ]]; then
  gcloud storage cp "${SCRIPT_DIR}/glossaries/"*.csv "gs://${GCS_BUCKET_NAME}/glossaries/"
fi

# Tạo chuỗi JSON cấu hình Pool Project cho KeyRotator
POOL_JSON="["
FIRST=1
for PID in "${POOL_PROJECT_IDS[@]}"; do
  [[ "${FIRST}" -eq 0 ]] && POOL_JSON+=","
  POOL_JSON+="{\"project_id\":\"${PID}\",\"location\":\"${REGION}\",\"gcs_bucket\":\"${GCS_BUCKET_NAME}\",\"enable_default_glossary_naming\":false}"
  FIRST=0
done
POOL_JSON+="]"

echo "=== [6/6] Build & Deploy Backend lên Cloud Run (Scale-to-Zero Always Free) ==="
gcloud run deploy "${SERVICE_NAME}" \
  --source="${PROJECT_ROOT}/backend" \
  --project="${PRIMARY_PROJECT_ID}" \
  --region="${REGION}" \
  --service-account="${SA_EMAIL}" \
  --allow-unauthenticated \
  --min-instances=0 \
  --max-instances=3 \
  --cpu=1 \
  --memory=512Mi \
  --timeout=300 \
  --set-env-vars="^|^APP_ENV=production|GCP_PROJECT_ID=${PRIMARY_PROJECT_ID}|GCS_SESSION_BUCKET=${GCS_BUCKET_NAME}|GCP_PROJECT_POOL_JSON=${POOL_JSON}|CF_ACCESS_TEAM_DOMAIN=${CF_ACCESS_TEAM_DOMAIN}|CF_ACCESS_AUD=${CF_ACCESS_AUD}|WORKER_SHARED_SECRET=${WORKER_SHARED_SECRET}|DLP_ENABLED=true|LEGAL_REVIEW_PROVIDER=${LEGAL_REVIEW_PROVIDER}|ANTHROPIC_API_KEY=${ANTHROPIC_API_KEY}|GEMINI_API_KEY=${GEMINI_API_KEY}"

CLOUD_RUN_URL="$(gcloud run services describe "${SERVICE_NAME}" --project="${PRIMARY_PROJECT_ID}" --region="${REGION}" --format="value(status.url)")"
echo ""
echo "✅ HOÀN TẤT DEPLOY CLOUD RUN!"
echo "   Cloud Run Origin URL : ${CLOUD_RUN_URL}"
echo "   Bước tiếp theo: Điền CLOUD_RUN_ORIGIN=${CLOUD_RUN_URL} vào worker-proxy/wrangler.toml và chạy npx wrangler deploy."
