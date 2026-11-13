# ⚖️ Web App Dịch Thuật Song Ngữ Hợp Đồng (Việt • Hàn • Anh)

Hệ thống dịch thuật tài liệu pháp lý / hợp đồng song ngữ tự động trên tên miền duy nhất **`https://dich.tnt.name.vn`**, thiết kế theo kiến trúc Serverless tối ưu gói Always Free.

---

## 1. Kiến Trúc Tổng Thể

```text
Người dùng ──► Cloudflare Access (Zero Trust: Google SSO / Email OTP)
                 │
                 ▼
       https://dich.tnt.name.vn (Cloudflare Worker Router)
          ├── GET /*        ──► Cloudflare Pages (Frontend: index.html + app.js)
          └── /api/*        ──► Google Cloud Run (FastAPI Backend + X-Worker-Shared-Secret)
                                   ├── Cloud Translation API v3 Advanced + Custom Glossary
                                   │     • <= 15 trang & <= 20.000 ký tự: translateDocument (Sync)
                                   │     • >  15 trang | >  20.000 ký tự: batchTranslateDocument (Async)
                                   ├── KeyRotator: Xoay vòng nhiều GCP Project (cùng 1 Billing thật)
                                   ├── Bilingual Layout Engine (python-docx + openpyxl):
                                   │     • side_by_side | alternating | replace_column
                                   ├── AI Legal Review (Claude Sonnet / Gemini):
                                   │     • Gắn cờ cảnh báo + chèn Comment lề phải (KHÔNG tự sửa hợp đồng)
                                   └── GCS Bucket (US region, Lifecycle tự xoá sau 7 ngày)
                                         + Cloud DLP (CHỈ che PII trong log/audit, KHÔNG che file tải về)
```

---

## 2. Cấu Trúc Thư Mục Mã Nguồn

- `backend/requirements.txt` — Thư viện Python 3.12 cho Cloud Run.
- `backend/Dockerfile` — Container tối giản chạy quyền non-root, tự động scale-to-zero.
- `backend/app/main.py` — FastAPI API (`/api/config`, `/api/translate`, `/api/jobs/{id}`, `/api/jobs/{id}/download`).
- `backend/app/translation_service.py` — Đo số trang/ký tự và điều phối `translateDocument` (sync) vs `batchTranslateDocument` (async).
- `backend/app/key_rotator.py` — Xoay vòng nhiều GCP Project thuộc cùng 1 tài khoản Billing thật, tự động chuyển project khi gặp HTTP 429.
- `backend/app/bilingual_layout.py` — Dựng hợp đồng song ngữ 2 cột đối chiếu, đoạn văn xen kẽ, hoặc thay thế cột ngôn ngữ (`.docx` & `.xlsx`).
- `backend/app/dlp_service.py` — Quét & che PII (CCCD, MST, SĐT, STK, số tiền...) bằng Cloud DLP + Regex **chỉ cho log/audit**, giữ nguyên 100% file người dùng tải về.
- `backend/app/legal_review.py` — Gọi Claude Sonnet / Gemini rà soát sai lệch số liệu, nhảy điều khoản, phạt vi phạm và gắn Comment bên lề Word/Excel.
- `backend/app/auth_cf_access.py` — Xác thực chữ ký JWT `Cf-Access-Jwt-Assertion` của Cloudflare Access + `X-Worker-Shared-Secret`.
- `worker-proxy/worker.js` & `wrangler.toml` — Cloudflare Worker hợp nhất 1 domain `dich.tnt.name.vn`.
- `frontend/index.html` & `frontend/app.js` — Giao diện kéo thả file, chọn cặp ngôn ngữ, bố cục, polling bất đồng bộ và hiển thị cảnh báo pháp lý.
- `infra/deploy.sh` & `infra/glossaries/*.csv` — Script triển khai Cloud Run + GCS Lifecycle 7 ngày + 3 bộ từ điển thuật ngữ pháp lý mẫu.

---

## 3. Hướng Dẫn Triển Khai Thực Tế Từng Bước (Dành Cho Anh Thanh)

### Bước 1: Tạo khóa bí mật chung (Worker <-> Cloud Run)
```bash
openssl rand -hex 32
```
Lưu lại chuỗi 64 ký tự hex này để điền vào cả `infra/deploy.sh` và Cloudflare Worker Secret.

### Bước 2: Thiết lập Cloudflare Access (Zero Trust)
1. Vào **Cloudflare Zero Trust Dashboard** -> **Access** -> **Applications** -> **Add an application (Self-hosted)**.
2. Đặt Application Domain: `dich.tnt.name.vn` (bảo vệ toàn bộ cả `/` và `/api/*`).
3. Tạo Policy cho phép danh sách email nội bộ (Email OTP hoặc Google Workspace SSO).
4. Copy **Application Audience (AUD) Tag** và **Team Domain** (`<team>.cloudflareaccess.com`).

### Bước 3: Deploy Backend lên Google Cloud Run + GCS Lifecycle 7 ngày
1. Mở `infra/deploy.sh`, điền các giá trị thật:
   - `PRIMARY_PROJECT_ID` (tuần đầu chỉ để 1 project trong `POOL_PROJECT_IDS` để kiểm chứng thực tế Cloud Billing report của Document Translation).
   - `CF_ACCESS_TEAM_DOMAIN`, `CF_ACCESS_AUD`, `WORKER_SHARED_SECRET`.
   - `ANTHROPIC_API_KEY` hoặc `GEMINI_API_KEY` (nếu dùng AI Legal Review).
2. Chạy lệnh triển khai:
```bash
chmod +x infra/deploy.sh
./infra/deploy.sh
```
3. Ghi lại địa chỉ `Cloud Run Origin URL` (`https://dich-hop-dong-api-xxxx-uc.a.run.app`) in ra ở cuối script.

### Bước 4: (Tuỳ chọn) Khởi tạo 3 Custom Glossaries trên Cloud Translation v3
Sau khi `deploy.sh` đã upload 3 file CSV lên `gs://<BUCKET>/glossaries/`, có thể đăng ký Glossary bằng REST/CLI hoặc bật `enable_default_glossary_naming: true` trong `GCP_PROJECT_POOL_JSON` sau khi tạo 3 glossary ID: `legal-glossary-vi-ko`, `legal-glossary-en-ko`, `legal-glossary-vi-en`.

### Bước 5: Deploy Giao diện lên Cloudflare Pages
```bash
npx wrangler pages project create dich-hop-dong-fe --production-branch=main
npx wrangler pages deploy frontend --project-name=dich-hop-dong-fe
```
Ghi lại URL Pages (`https://dich-hop-dong-fe.pages.dev`).

### Bước 6: Deploy Cloudflare Worker gắn vào `dich.tnt.name.vn`
1. Mở `worker-proxy/wrangler.toml`, điền `CLOUD_RUN_ORIGIN` và `PAGES_ORIGIN` thật.
2. Nạp `WORKER_SHARED_SECRET` an toàn vào Worker:
```bash
cd worker-proxy
npx wrangler secret put WORKER_SHARED_SECRET
npx wrangler deploy
```
3. Đảm bảo trong DNS của `tnt.name.vn` có bản ghi `dich` (Proxied - đám mây cam bật) để Worker nhận lưu lượng `dich.tnt.name.vn/*`.

---

## 4. Lưu Ý Vận Hành & Kiểm Soát Chi Phí Tuần Đầu
- **Theo dõi Billing tuần đầu:** Trong 7 ngày đầu tiên, chỉ cấu hình 1 project trong `POOL_PROJECT_IDS` và kiểm tra mục **Cloud Billing -> Reports** (lọc theo SKU *Cloud Translation API - Document Translation*) để xác nhận chính xác cách trừ hạn mức 500.000 ký tự miễn phí/tháng trước khi thêm các project tiếp theo vào `POOL_PROJECT_IDS`.
- **Tính toàn vẹn pháp lý:** Cloud DLP chỉ che dữ liệu nhạy cảm trong log/audit; hợp đồng tải xuống giữ nguyên 100% nội dung gốc và AI Legal Review chỉ gắn Comment bên lề để người dùng chủ động xét duyệt.
