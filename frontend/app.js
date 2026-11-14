/**
 * Frontend Controller for Bilingual Contract Translation Web App (dich.tnt.name.vn)
 * Calls unified same-origin /api/* endpoints proxied by Cloudflare Worker -> Cloud Run.
 */
(() => {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const selectedFilePill = document.getElementById('selectedFilePill');
  const langPairSelect = document.getElementById('langPairSelect');
  const layoutGroup = document.getElementById('layoutGroup');
  const replaceColWrapper = document.getElementById('replaceColWrapper');
  const replaceColIndex = document.getElementById('replaceColIndex');
  const useGlossary = document.getElementById('useGlossary');
  const enableLegalReview = document.getElementById('enableLegalReview');
  const translateForm = document.getElementById('translateForm');
  const submitBtn = document.getElementById('submitBtn');

  const statusCard = document.getElementById('statusCard');
  const statusTitle = document.getElementById('statusTitle');
  const statusDesc = document.getElementById('statusDesc');
  const progressBar = document.getElementById('progressBar');
  const metricsInfo = document.getElementById('metricsInfo');
  const downloadContainer = document.getElementById('downloadContainer');
  const legalReviewContainer = document.getElementById('legalReviewContainer');
  const authBadge = document.getElementById('authBadge');

  const tabTranslateBtn = document.getElementById('tabTranslateBtn');
  const tabOcrBtn = document.getElementById('tabOcrBtn');
  const translatePanel = document.getElementById('translatePanel');
  const ocrPanel = document.getElementById('ocrPanel');

  const ocrForm = document.getElementById('ocrForm');
  const ocrDropzone = document.getElementById('ocrDropzone');
  const ocrFileInput = document.getElementById('ocrFileInput');
  const ocrFilePill = document.getElementById('ocrFilePill');
  const ocrSubmitBtn = document.getElementById('ocrSubmitBtn');
  const ocrStatusCard = document.getElementById('ocrStatusCard');
  const ocrStatusTitle = document.getElementById('ocrStatusTitle');
  const ocrStatusDesc = document.getElementById('ocrStatusDesc');
  const ocrProgressBar = document.getElementById('ocrProgressBar');
  const ocrMetricsInfo = document.getElementById('ocrMetricsInfo');
  const ocrDownloadContainer = document.getElementById('ocrDownloadContainer');
  const ocrTextPreviewContainer = document.getElementById('ocrTextPreviewContainer');
  const ocrTextPreview = document.getElementById('ocrTextPreview');

  let currentFile = null;
  let currentOcrFile = null;
  let pollTimer = null;

  function switchTab(tab) {
    const isOcr = tab === 'ocr';
    tabTranslateBtn.classList.toggle('active', !isOcr);
    tabOcrBtn.classList.toggle('active', isOcr);
    translatePanel.classList.toggle('hidden', isOcr);
    ocrPanel.classList.toggle('hidden', !isOcr);
  }
  tabTranslateBtn.addEventListener('click', () => switchTab('translate'));
  tabOcrBtn.addEventListener('click', () => switchTab('ocr'));

  function formatErrorDetail(detail, fallback) {
    if (Array.isArray(detail)) {
      const msgs = detail
        .map((item) => {
          if (item && typeof item === 'object') {
            const loc = Array.isArray(item.loc) ? item.loc.join('.') : '';
            const msg = item.msg || JSON.stringify(item);
            return loc ? `${loc}: ${msg}` : msg;
          }
          return String(item);
        })
        .filter(Boolean);
      return msgs.length ? msgs.join('; ') : fallback;
    }
    if (detail && typeof detail === 'object') {
      return detail.message || JSON.stringify(detail);
    }
    return detail ? String(detail) : fallback;
  }

  async function readApiResponse(res) {
    const rawText = await res.text();
    let data = null;
    let isJson = false;
    if (rawText) {
      try {
        data = JSON.parse(rawText);
        isJson = true;
      } catch (_) {
        isJson = false;
      }
    }
    if (!res.ok) {
      if (isJson && data && typeof data === 'object' && data.detail !== undefined) {
        throw new Error(formatErrorDetail(data.detail, `HTTP ${res.status}`));
      }
      const snippet = (rawText || res.statusText || 'Lỗi không xác định').trim().slice(0, 200);
      throw new Error(`HTTP ${res.status}: ${snippet}`);
    }
    if (!isJson || data === null) {
      const snippet = (rawText || '').trim().slice(0, 200);
      throw new Error(`HTTP ${res.status}: ${snippet || 'Phản hồi không phải JSON hợp lệ'}`);
    }
    return data;
  }

  // Load backend config & verify Cloudflare Access session
  async function initConfig() {
    try {
      const res = await fetch('/api/config', { credentials: 'include' });
      if (res.ok) {
        const cfg = await readApiResponse(res);
        if (cfg.user && cfg.user.email) {
          authBadge.textContent = '🔒 Cloudflare Access • Ready';
        }
      }
    } catch (_) {
      // Silent on initial load; error surfaces clearly upon action if offline
    }
  }

  function setFile(file) {
    if (!file) return;
    const lower = file.name.toLowerCase();
    if (!lower.endsWith('.docx') && !lower.endsWith('.xlsx') && !lower.endsWith('.pdf')) {
      alert('Vui lòng chỉ chọn file hợp đồng Word (.docx), Excel (.xlsx) hoặc PDF (.pdf).');
      return;
    }
    if (file.size > 25 * 1024 * 1024) {
      alert('Dung lượng file vượt quá giới hạn 25MB.');
      return;
    }
    currentFile = file;
    const kb = (file.size / 1024).toFixed(1);
    const icon = lower.endsWith('.pdf') ? '📕' : lower.endsWith('.xlsx') ? '📗' : '📘';
    selectedFilePill.textContent = `${icon} ${file.name} (${kb} KB)`;
    selectedFilePill.classList.remove('hidden');
  }

  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) setFile(e.target.files[0]);
  });
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
  });

  // Layout radio selection highlight + toggle replace_column input
  layoutGroup.addEventListener('change', () => {
    const selected = document.querySelector('input[name="layout_mode"]:checked')?.value;
    document.querySelectorAll('.layout-option').forEach((lbl) => {
      const radio = lbl.querySelector('input[type="radio"]');
      lbl.classList.toggle('active', Boolean(radio && radio.checked));
    });
    replaceColWrapper.classList.toggle('hidden', selected !== 'replace_column');
  });

  function escapeHtml(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function renderCompleted(data) {
    progressBar.style.width = '100%';
    statusTitle.textContent = '✅ Dịch & dàn trang song ngữ hoàn tất!';
    statusDesc.textContent = `Chế độ xử lý: ${data.mode === 'async' ? 'Batch Async (Hợp đồng lớn)' : 'Đồng bộ tốc độ cao (Sync)'}. File tải về giữ nguyên 100% dữ liệu gốc.`;

    if (data.metrics) {
      let pdfNote = '';
      if (data.metrics.pdf_type === 'text_pdf') {
        pdfNote = ' • Text PDF (0 quota OCR)';
      } else if (data.metrics.pdf_type === 'scanned_pdf') {
        pdfNote = ` • Scanned PDF (Vision OCR: ${data.metrics.ocr_vision_requests || 0} trang)`;
      }
      metricsInfo.textContent = `📊 Thống kê: ~${data.metrics.estimated_pages} trang • ${Number(data.metrics.char_count || 0).toLocaleString('vi-VN')} ký tự trích xuất${pdfNote}`;
    }

    if (data.download_url) {
      const fname = escapeHtml(data.output_filename || 'hop-dong-song-ngu.docx');
      downloadContainer.innerHTML = `<a class="download-btn" href="${escapeHtml(data.download_url)}" download="${fname}">⬇️ Tải xuống: ${fname}</a>`;
    }

    legalReviewContainer.innerHTML = '';
    if (data.legal_review && data.legal_review.enabled) {
      const lr = data.legal_review;
      const issues = Array.isArray(lr.issues) ? lr.issues : [];
      let html = `<div style="border-top:1px solid #e2e8f0; padding-top:14px;">`;
      html += `<strong>⚖️ Báo cáo Thẩm định Pháp lý AI (${escapeHtml(lr.provider)} / ${escapeHtml(lr.model)})</strong>`;
      html += `<p class="meta-note" style="margin-top:4px;">${escapeHtml(lr.summary)} (Các đề xuất dưới đây đã được đính kèm dưới dạng Comment bên lề trong file tải về, không tự ý sửa nội dung gốc).</p>`;
      if (issues.length === 0) {
        html += `<div class="issue-card low">✅ Không phát hiện sai lệch số tiền, nhảy điều khoản hay rủi ro dịch thuật nghiêm trọng.</div>`;
      } else {
        for (const item of issues) {
          const sev = escapeHtml(item.severity || 'medium');
          html += `<div class="issue-card ${sev}">`;
          html += `<div><strong>[${sev.toUpperCase()}] ${escapeHtml(item.clause_ref)}</strong> — <em>${escapeHtml(item.category)}</em></div>`;
          html += `<div style="margin-top:4px;">⚠️ <strong>Cảnh báo:</strong> ${escapeHtml(item.warning_flag)}</div>`;
          html += `<div style="margin-top:4px;">💡 <strong>Comment đề xuất:</strong> ${escapeHtml(item.suggested_comment)}</div>`;
          html += `</div>`;
        }
      }
      html += `</div>`;
      legalReviewContainer.innerHTML = html;
    }
  }

  async function pollAsyncJob(pollUrl) {
    let tick = 0;
    pollTimer = setInterval(async () => {
      tick += 1;
      const pct = Math.min(92, 35 + tick * 5);
      progressBar.style.width = `${pct}%`;
      try {
        const res = await fetch(pollUrl, { credentials: 'include' });
        const data = await readApiResponse(res);

        if (data.status === 'completed') {
          clearInterval(pollTimer);
          pollTimer = null;
          submitBtn.disabled = false;
          renderCompleted(data);
        } else if (data.status === 'failed') {
          clearInterval(pollTimer);
          pollTimer = null;
          submitBtn.disabled = false;
          statusTitle.textContent = '❌ Dịch thuật thất bại';
          statusDesc.textContent = data.error || 'Đã xảy ra lỗi trong quá trình xử lý Batch.';
        } else {
          statusDesc.textContent = data.message || `Đang dịch hợp đồng lớn trên Google Cloud... (lần kiểm tra #${tick})`;
        }
      } catch (err) {
        clearInterval(pollTimer);
        pollTimer = null;
        submitBtn.disabled = false;
        statusTitle.textContent = '❌ Mất kết nối khi kiểm tra tiến độ';
        statusDesc.textContent = err.message;
      }
    }, 3500);
  }

  translateForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentFile) {
      alert('Vui lòng chọn file hợp đồng (.docx, .xlsx hoặc .pdf) trước khi dịch.');
      return;
    }
    if (pollTimer) {
      clearInterval(pollTimer);
      pollTimer = null;
    }

    let srcLang;
    let tgtLang;
    if (langPairSelect.value === 'vi-en-kr') {
      srcLang = 'vi';
      tgtLang = 'en,ko';
    } else {
      [srcLang, tgtLang] = langPairSelect.value.split('-');
    }
    const layoutMode = document.querySelector('input[name="layout_mode"]:checked')?.value || 'alternating';

    const formData = new FormData();
    formData.append('file', currentFile);
    formData.append('source_lang', srcLang);
    formData.append('target_lang', tgtLang);
    formData.append('layout_mode', layoutMode);
    formData.append('use_glossary', useGlossary.checked ? 'true' : 'false');
    formData.append('enable_legal_review', enableLegalReview.checked ? 'true' : 'false');
    formData.append('replace_col_index', String(replaceColIndex.value || 1));

    submitBtn.disabled = true;
    statusCard.classList.add('visible');
    statusTitle.textContent = '⏳ Đang tải lên & phân tích dung lượng hợp đồng...';
    statusDesc.textContent = 'Hệ thống tự động chọn luồng Đồng bộ (<=15 trang & <=20.000 ký tự) hoặc Batch Async.';
    progressBar.style.width = '28%';
    metricsInfo.textContent = '';
    downloadContainer.innerHTML = '';
    legalReviewContainer.innerHTML = '';

    try {
      const res = await fetch('/api/translate', {
        method: 'POST',
        body: formData,
        credentials: 'include',
      });
      const data = await readApiResponse(res);

      if (res.status === 202 || data.status === 'processing') {
        statusTitle.textContent = '🔄 Đang dịch hợp đồng lớn (Batch Async)...';
        statusDesc.textContent = data.message || 'Vui lòng giữ tab mở, hệ thống tự động cập nhật sau mỗi vài giây.';
        progressBar.style.width = '38%';
        if (data.metrics) {
          metricsInfo.textContent = `📊 Thống kê: ~${data.metrics.estimated_pages} trang • ${Number(data.metrics.char_count || 0).toLocaleString('vi-VN')} ký tự`;
        }
        pollAsyncJob(data.poll_url || `/api/jobs/${data.job_id}`);
      } else {
        submitBtn.disabled = false;
        renderCompleted(data);
      }
    } catch (err) {
      submitBtn.disabled = false;
      progressBar.style.width = '0%';
      statusTitle.textContent = '❌ Không thể thực hiện dịch thuật';
      statusDesc.textContent = err.message || String(err);
    }
  });

  // Tab 2: OCR Only Dropzone & Submit
  function setOcrFile(file) {
    if (!file) return;
    const lower = file.name.toLowerCase();
    if (
      !lower.endsWith('.pdf') &&
      !lower.endsWith('.png') &&
      !lower.endsWith('.jpg') &&
      !lower.endsWith('.jpeg')
    ) {
      alert('Vui lòng chỉ chọn file PDF (.pdf) hoặc ảnh (.png, .jpg, .jpeg).');
      return;
    }
    if (file.size > 25 * 1024 * 1024) {
      alert('Dung lượng file vượt quá giới hạn 25MB.');
      return;
    }
    currentOcrFile = file;
    const kb = (file.size / 1024).toFixed(1);
    const icon = lower.endsWith('.pdf') ? '📕' : '🖼️';
    ocrFilePill.textContent = `${icon} ${file.name} (${kb} KB)`;
    ocrFilePill.classList.remove('hidden');
  }

  ocrDropzone.addEventListener('click', () => ocrFileInput.click());
  ocrFileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) setOcrFile(e.target.files[0]);
  });
  ocrDropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    ocrDropzone.classList.add('dragover');
  });
  ocrDropzone.addEventListener('dragleave', () => ocrDropzone.classList.remove('dragover'));
  ocrDropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    ocrDropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) setOcrFile(e.dataTransfer.files[0]);
  });

  ocrForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!currentOcrFile) {
      alert('Vui lòng chọn file PDF hoặc ảnh (.pdf, .png, .jpg, .jpeg) trước khi OCR.');
      return;
    }

    const formData = new FormData();
    formData.append('file', currentOcrFile);

    ocrSubmitBtn.disabled = true;
    ocrStatusCard.classList.add('visible');
    ocrStatusTitle.textContent = '⏳ Đang trích xuất văn bản (OCR)...';
    ocrStatusDesc.textContent = 'Đang xử lý PDF/ảnh bằng Google Cloud Vision OCR & tạo file Word (.docx).';
    ocrProgressBar.style.width = '45%';
    ocrMetricsInfo.textContent = '';
    ocrDownloadContainer.innerHTML = '';
    ocrTextPreviewContainer.classList.add('hidden');
    ocrTextPreview.textContent = '';

    try {
      const res = await fetch('/api/ocr', {
        method: 'POST',
        body: formData,
        credentials: 'include',
      });
      const data = await readApiResponse(res);

      ocrSubmitBtn.disabled = false;
      ocrProgressBar.style.width = '100%';
      ocrStatusTitle.textContent = '✅ Trích xuất OCR hoàn tất!';
      ocrStatusDesc.textContent = 'File Word (.docx) đã sẵn sàng tải xuống (không dịch thuật).';

      if (data.metrics) {
        const m = data.metrics;
        const note = m.pdf_type === 'text_pdf'
          ? 'Text PDF (0 quota OCR)'
          : `Vision OCR: ${m.ocr_vision_requests || 0} trang`;
        ocrMetricsInfo.textContent = `📊 Thống kê: ~${m.estimated_pages} trang • ${Number(m.char_count || 0).toLocaleString('vi-VN')} ký tự • ${note}`;
      }

      if (data.download_url) {
        const fname = escapeHtml(data.output_filename || 'ocr-output.docx');
        ocrDownloadContainer.innerHTML = `<a class="download-btn" href="${escapeHtml(data.download_url)}" download="${fname}">⬇️ Tải xuống: ${fname}</a>`;
      }

      if (data.text_preview) {
        ocrTextPreview.textContent = data.text_preview + (data.text_truncated ? '\n\n... (đã rút gọn; tải file .docx để xem đầy đủ)' : '');
        ocrTextPreviewContainer.classList.remove('hidden');
      }
    } catch (err) {
      ocrSubmitBtn.disabled = false;
      ocrProgressBar.style.width = '0%';
      ocrStatusTitle.textContent = '❌ Không thể thực hiện OCR';
      ocrStatusDesc.textContent = err.message || String(err);
    }
  });

  initConfig();
})();
