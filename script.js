/**
 * DocuScan — Azure Document Intelligence Frontend
 * ─────────────────────────────────────────────────
 * Pure vanilla JS · no frameworks · no backend
 * Uses Azure Document Intelligence prebuilt-layout model
 * Polling-based async processing via Operation-Location header
 */

'use strict';

/* ═══════════════════════════════════════════════════════════════
   CONSTANTS & CONFIG
═══════════════════════════════════════════════════════════════ */
const AZURE_API_VERSION = '2024-11-30';
const POLL_INTERVAL_MS  = 2000;   // how often to poll for results
const MAX_POLL_ATTEMPTS = 90;     // 90 × 2s = 3 min max wait
const MODEL_ID          = 'prebuilt-layout';

/* localStorage keys */
const STORAGE_ENDPOINT  = 'docuscan_endpoint';
const STORAGE_KEY       = 'docuscan_key';
const STORAGE_THEME     = 'docuscan_theme';

/* ═══════════════════════════════════════════════════════════════
   STATE
═══════════════════════════════════════════════════════════════ */
const State = {
  endpoint:     '',          // Azure Document Intelligence endpoint
  azureKey:     '',          // Azure Document Intelligence key
  activeTab:    'file',      // file | url | camera
  currentFile:  null,        // File object
  currentUrl:   '',          // URL string
  cameraStream: null,        // MediaStream
  isProcessing: false,       // guard against double-clicks
  extractedText: '',         // last successful extraction
};

/* ═══════════════════════════════════════════════════════════════
   DOM HELPERS
═══════════════════════════════════════════════════════════════ */
const el   = id  => document.getElementById(id);
const qs   = sel => document.querySelector(sel);
const qsa  = sel => document.querySelectorAll(sel);

/* ═══════════════════════════════════════════════════════════════
   INIT
═══════════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
  loadPersistedData();
  initThemeSwitcher();
  initCredentials();
  initTabs();
  initFileUpload();
  initCamera();
  initAnalyze();
  initResults();
  initParticles();
  setupAccessibility();
});

/* ═══════════════════════════════════════════════════════════════
   PERSIST / LOAD
═══════════════════════════════════════════════════════════════ */
function loadPersistedData() {
  State.endpoint = localStorage.getItem(STORAGE_ENDPOINT) || '';
  State.azureKey = localStorage.getItem(STORAGE_KEY)      || '';
  const savedTheme = localStorage.getItem(STORAGE_THEME) || 'default';

  el('azureEndpoint').value = State.endpoint;
  el('azureKey').value      = State.azureKey;
  applyTheme(savedTheme);

  // If credentials already saved, collapse credentials card
  if (State.endpoint && State.azureKey) {
    setTimeout(() => collapseCredentials(true), 600);
    showCredStatus('saved', '✓ Credentials loaded');
  }
}

function saveCredentials() {
  const endpoint = el('azureEndpoint').value.trim();
  const key      = el('azureKey').value.trim();

  if (!endpoint || !key) {
    toast('Please enter both endpoint and key.', 'error');
    return;
  }
  if (!endpoint.startsWith('https://')) {
    toast('Endpoint must start with https://', 'error');
    return;
  }

  State.endpoint = endpoint.replace(/\/$/, ''); // strip trailing slash
  State.azureKey = key;

  localStorage.setItem(STORAGE_ENDPOINT, State.endpoint);
  localStorage.setItem(STORAGE_KEY,      State.azureKey);

  showCredStatus('saved', '✓ Credentials saved');
  toast('Azure Document Intelligence credentials saved.', 'success');
  collapseCredentials(true);
}

function clearCredentials() {
  State.endpoint = '';
  State.azureKey = '';
  localStorage.removeItem(STORAGE_ENDPOINT);
  localStorage.removeItem(STORAGE_KEY);
  el('azureEndpoint').value = '';
  el('azureKey').value      = '';
  showCredStatus('cleared', '✕ Cleared');
  toast('Credentials cleared from local storage.', 'info');
}

function showCredStatus(cls, text) {
  const s = el('credStatus');
  s.className = 'cred-status ' + cls;
  s.textContent = text;
  clearTimeout(s._timer);
  s._timer = setTimeout(() => { s.textContent = ''; s.className = 'cred-status'; }, 4000);
}

/* ═══════════════════════════════════════════════════════════════
   CREDENTIALS CARD UI
═══════════════════════════════════════════════════════════════ */
function initCredentials() {
  el('saveCredBtn').addEventListener('click',  saveCredentials);
  el('clearCredBtn').addEventListener('click', clearCredentials);
  el('collapseCredBtn').addEventListener('click', () => {
    const body    = el('credBody');
    const btn     = el('collapseCredBtn');
    const isOpen  = !body.classList.contains('collapsed');
    collapseCredentials(isOpen);
  });

  // Toggle API key visibility
  el('revealKeyBtn').addEventListener('click', () => {
    const inp    = el('azureKey');
    const isPass = inp.type === 'password';
    inp.type     = isPass ? 'text' : 'password';
    qs('.eye-show').style.display = isPass ? 'none'  : '';
    qs('.eye-hide').style.display = isPass ? ''      : 'none';
  });
}

function collapseCredentials(collapse) {
  const body = el('credBody');
  const btn  = el('collapseCredBtn');
  body.classList.toggle('collapsed', collapse);
  btn.classList.toggle('collapsed',  collapse);
  btn.setAttribute('aria-expanded', String(!collapse));
}

/* ═══════════════════════════════════════════════════════════════
   THEME SWITCHER
═══════════════════════════════════════════════════════════════ */
function initThemeSwitcher() {
  qsa('.theme-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const theme = btn.dataset.theme;
      applyTheme(theme);
      localStorage.setItem(STORAGE_THEME, theme);
    });
  });
}

function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  qsa('.theme-btn').forEach(b => b.classList.toggle('active', b.dataset.theme === theme));
}

/* ═══════════════════════════════════════════════════════════════
   TABS
═══════════════════════════════════════════════════════════════ */
function initTabs() {
  qsa('.tab').forEach(tab => {
    tab.addEventListener('click', () => switchTab(tab.dataset.tab));
  });
}

function switchTab(tabId) {
  qsa('.tab').forEach(t => {
    const active = t.dataset.tab === tabId;
    t.classList.toggle('active', active);
    t.setAttribute('aria-selected', String(active));
  });
  qsa('.tab-panel').forEach(p => {
    p.classList.toggle('active', p.id === `panel-${tabId}`);
  });
  State.activeTab = tabId;

  // Stop camera if switching away
  if (tabId !== 'camera' && State.cameraStream) {
    stopCamera();
  }
}

/* ═══════════════════════════════════════════════════════════════
   FILE UPLOAD & DRAG-AND-DROP
═══════════════════════════════════════════════════════════════ */
function initFileUpload() {
  const dropZone = el('dropZone');
  const fileInput = el('fileInput');

  // Browse button opens file picker
  el('browseBtn').addEventListener('click', e => {
    e.stopPropagation();
    fileInput.click();
  });

  // Click on drop zone also opens file picker
  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
  });

  fileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) setFile(file);
    fileInput.value = ''; // reset so same file can be re-selected
  });

  // Drag events
  dropZone.addEventListener('dragenter', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', e => {
    if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over');
  });
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) setFile(file);
  });

  // Also allow drop anywhere on window (convenience)
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => {
    e.preventDefault();
    if (State.activeTab !== 'file') switchTab('file');
    const file = e.dataTransfer.files[0];
    if (file) setFile(file);
  });

  // Remove button
  el('clearDocBtn').addEventListener('click', clearDocument);
}

/** Accept a file, validate it, and show preview */
function setFile(file) {
  const allowed = ['image/png','image/jpeg','image/jpg','image/webp','image/tiff','image/bmp','application/pdf'];
  if (!allowed.includes(file.type) && !file.name.match(/\.(pdf|png|jpe?g|webp|tiff?|bmp)$/i)) {
    toast('Unsupported file type. Please use PDF or an image file.', 'error');
    return;
  }
  if (file.size > 50 * 1024 * 1024) {
    toast('File exceeds 50 MB limit.', 'error');
    return;
  }

  State.currentFile = file;
  State.currentUrl  = '';

  el('previewFilename').textContent = `${file.name} · ${formatBytes(file.size)}`;
  el('previewArea').hidden = false;

  const img = el('previewImage');
  const pdf = el('previewPdf');

  if (file.type === 'application/pdf') {
    img.hidden = true;
    pdf.hidden = false;
    pdf.src = URL.createObjectURL(file);
  } else {
    pdf.hidden = true;
    img.hidden = false;
    img.src = URL.createObjectURL(file);
  }

  hideResults();
  hideError();
}

function clearDocument() {
  State.currentFile = null;
  State.currentUrl  = '';
  el('previewArea').hidden = true;
  el('previewImage').src   = '';
  el('previewPdf').src     = '';
  el('fileInput').value    = '';
  hideResults();
  hideError();
}

/* ═══════════════════════════════════════════════════════════════
   CAMERA CAPTURE
═══════════════════════════════════════════════════════════════ */
function initCamera() {
  el('startCameraBtn').addEventListener('click', startCamera);
  el('captureBtn').addEventListener('click',    capturePhoto);
  el('stopCameraBtn').addEventListener('click', stopCamera);
}

async function startCamera() {
  try {
    State.cameraStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: 'environment', width: { ideal: 1920 }, height: { ideal: 1080 } }
    });
    const video = el('cameraVideo');
    video.srcObject = State.cameraStream;
    video.classList.add('active');
    el('captureBtn').disabled   = false;
    el('captureBtn').removeAttribute('aria-disabled');
    el('stopCameraBtn').disabled = false;
    el('stopCameraBtn').removeAttribute('aria-disabled');
    el('startCameraBtn').disabled = true;
  } catch (err) {
    toast('Camera access denied. Please allow camera permissions.', 'error');
    console.error('Camera error:', err);
  }
}

function capturePhoto() {
  const video  = el('cameraVideo');
  const canvas = el('cameraCanvas');
  canvas.width  = video.videoWidth;
  canvas.height = video.videoHeight;
  canvas.getContext('2d').drawImage(video, 0, 0);
  canvas.toBlob(blob => {
    const file = new File([blob], `capture-${Date.now()}.jpg`, { type: 'image/jpeg' });
    stopCamera();
    switchTab('file');
    setFile(file);
    toast('Photo captured!', 'success');
  }, 'image/jpeg', 0.95);
}

function stopCamera() {
  if (State.cameraStream) {
    State.cameraStream.getTracks().forEach(t => t.stop());
    State.cameraStream = null;
  }
  const video = el('cameraVideo');
  video.srcObject = null;
  video.classList.remove('active');
  el('captureBtn').disabled     = true;
  el('captureBtn').setAttribute('aria-disabled', 'true');
  el('stopCameraBtn').disabled  = true;
  el('startCameraBtn').disabled = false;
}

/* ═══════════════════════════════════════════════════════════════
   AZURE DOCUMENT INTELLIGENCE — MAIN FLOW
═══════════════════════════════════════════════════════════════ */
function initAnalyze() {
  el('analyzeBtn').addEventListener('click', runAnalysis);
}

async function runAnalysis() {
  // ── Guard ──────────────────────────────────────────────────
  if (State.isProcessing) return;

  // ── Validate credentials ───────────────────────────────────
  const endpoint = el('azureEndpoint').value.trim() || State.endpoint;
  const key      = el('azureKey').value.trim()      || State.azureKey;

  if (!endpoint || !key) {
    toast('Please enter your Azure Document Intelligence endpoint and key in the credentials panel.', 'error');
    collapseCredentials(false);
    el('azureEndpoint').focus();
    return;
  }

  // ── Validate input ─────────────────────────────────────────
  const tab = State.activeTab;
  let hasInput = false;

  if (tab === 'file'   && State.currentFile) hasInput = true;
  if (tab === 'url'    && el('urlInput').value.trim()) hasInput = true;
  if (tab === 'camera' && State.currentFile) hasInput = true;

  if (!hasInput) {
    toast(tab === 'url'
      ? 'Please enter a document URL.'
      : 'Please upload or capture a document first.',
    'error');
    return;
  }

  State.isProcessing = true;
  State.endpoint     = endpoint.replace(/\/$/, '');
  State.azureKey     = key;

  hideError();
  hideResults();
  showProgress(true);
  setStep('upload', 'active');
  setProgressBar(10);
  el('analyzeBtn').disabled = true;

  try {
    // STEP 1: Submit document → get Operation-Location URL
    setStatus('uploading', 'Uploading document…');
    const operationUrl = await submitDocument(tab);

    setStep('upload', 'done');
    setStep('process', 'active');
    setProgressBar(35);
    setStatus('processing', 'Processing with Azure Document Intelligence…');

    // STEP 2: Poll until result is ready
    const result = await pollForResult(operationUrl);

    setStep('process', 'done');
    setStep('extract', 'active');
    setProgressBar(80);
    setStatus('extracting', 'Extracting text and tables…');

    // Small deliberate pause so user sees "Extracting" step
    await delay(500);

    // STEP 3: Parse and render the result
    renderResult(result);

    setStep('extract', 'done');
    setStep('done', 'done');
    setProgressBar(100);
    setStatus('completed', '✓ Analysis complete');
    toast('Document analyzed successfully!', 'success');

  } catch (err) {
    console.error('Analysis error:', err);
    showError(err.message || 'An unexpected error occurred.');
    setStatus('error', '✕ Error');
    resetSteps();
    setProgressBar(0);
  } finally {
    State.isProcessing = false;
    el('analyzeBtn').disabled = false;
    setTimeout(() => showProgress(false), 1500);
  }
}

/* ─── SUBMIT DOCUMENT ──────────────────────────────────────── */
/**
 * Submits the document to Azure Document Intelligence.
 * Returns the operation URL to poll.
 */
async function submitDocument(tab) {
  // Azure Doc Intelligence analyze endpoint (v4 GA)
  const analyzeUrl =
    `${State.endpoint}/documentintelligence/documentModels/${MODEL_ID}:analyze` +
    `?api-version=${AZURE_API_VERSION}`;

  let body, contentType;

  if (tab === 'url') {
    // Send URL as JSON body
    const docUrl = el('urlInput').value.trim();
    if (!docUrl.startsWith('http')) throw new Error('Please enter a valid URL starting with http.');
    body        = JSON.stringify({ urlSource: docUrl });
    contentType = 'application/json';

  } else {
    // Send file as binary body
    const file = State.currentFile;
    if (!file) throw new Error('No file selected.');
    body        = await file.arrayBuffer();
    contentType = file.type || 'application/octet-stream';
  }

  const response = await fetch(analyzeUrl, {
    method: 'POST',
    headers: {
      'Content-Type': contentType,
      'Ocp-Apim-Subscription-Key': State.azureKey,
    },
    body,
  });

  if (response.status === 202) {
    // Async accepted — get poll URL from header
    const opUrl = response.headers.get('Operation-Location') ||
                  response.headers.get('operation-location');
    if (!opUrl) throw new Error('Azure returned 202 but no Operation-Location header.');
    return opUrl;
  }

  if (response.status === 200) {
    // Synchronous result (unusual but handle it)
    const json = await response.json();
    return { _syncResult: json };
  }

  // Error response
  let errText = `Azure Document Intelligence returned status ${response.status}.`;
  try {
    const errJson = await response.json();
    const msg = errJson?.error?.message || errJson?.message;
    if (msg) errText += ' ' + msg;
  } catch {}

  if (response.status === 401) errText = 'Authentication failed. Check your Azure Document Intelligence key.';
  if (response.status === 403) errText = 'Access forbidden. Verify the endpoint and key are correct.';
  if (response.status === 404) errText = 'Resource not found. Check your Azure Document Intelligence endpoint URL.';

  throw new Error(errText);
}

/* ─── POLL FOR RESULT ──────────────────────────────────────── */
/**
 * Polls the Operation-Location URL until Azure finishes processing.
 * Returns the analyzeResult object.
 */
async function pollForResult(operationUrlOrSyncResult) {
  // Handle sync result
  if (operationUrlOrSyncResult?._syncResult) {
    return operationUrlOrSyncResult._syncResult?.analyzeResult || operationUrlOrSyncResult._syncResult;
  }

  const operationUrl = operationUrlOrSyncResult;
  let attempts = 0;

  while (attempts < MAX_POLL_ATTEMPTS) {
    await delay(POLL_INTERVAL_MS);
    attempts++;

    const response = await fetch(operationUrl, {
      headers: { 'Ocp-Apim-Subscription-Key': State.azureKey },
    });

    if (!response.ok) {
      throw new Error(`Polling failed with status ${response.status}.`);
    }

    const json   = await response.json();
    const status = (json.status || '').toLowerCase();

    if (status === 'succeeded') {
      return json.analyzeResult;
    }

    if (status === 'failed') {
      const msg = json?.error?.message || json?.error?.innererror?.message || 'Azure processing failed.';
      throw new Error(msg);
    }

    // 'running' or 'notStarted' — keep polling
    // Update progress bar to reflect polling progress
    const pct = 35 + Math.min(40, Math.floor((attempts / MAX_POLL_ATTEMPTS) * 40));
    setProgressBar(pct);
  }

  throw new Error('Analysis timed out after 3 minutes. Please try again.');
}

/* ═══════════════════════════════════════════════════════════════
   RENDER RESULT
═══════════════════════════════════════════════════════════════ */
function renderResult(analyzeResult) {
  if (!analyzeResult) {
    throw new Error('Azure returned an empty result.');
  }

  // ── Full text ──────────────────────────────────────────────
  const fullText = analyzeResult.content || '';
  State.extractedText = fullText;

  el('textOutput').textContent = fullText || '(No text detected)';

  const wordCount = fullText.trim() ? fullText.trim().split(/\s+/).length : 0;
  el('textStats').textContent =
    `${wordCount.toLocaleString()} words · ${fullText.length.toLocaleString()} characters`;

  // ── Tables ─────────────────────────────────────────────────
  const tables = analyzeResult.tables || [];
  if (tables.length > 0) {
    el('tablesCard').hidden = false;
    el('tableStats').textContent = `${tables.length} table${tables.length > 1 ? 's' : ''} detected`;
    renderTables(tables);
  } else {
    el('tablesCard').hidden = true;
  }

  // Show results section
  el('resultsSection').hidden = false;
  el('resultsSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/**
 * Renders Azure table objects into HTML <table> elements.
 * Azure tables contain a flat `cells` array with rowIndex/columnIndex.
 */
function renderTables(tables) {
  const container = el('tablesOutput');
  container.innerHTML = '';

  tables.forEach((table, tableIdx) => {
    const rowCount = table.rowCount || 0;
    const colCount = table.columnCount || 0;
    if (rowCount === 0 || colCount === 0) return;

    // Build a 2D grid, tracking which cells are headers
    const grid       = Array.from({ length: rowCount }, () => Array(colCount).fill(null));
    const isHeader   = Array.from({ length: rowCount }, () => Array(colCount).fill(false));
    const spans      = Array.from({ length: rowCount }, () => Array(colCount).fill({ r: 1, c: 1 }));

    (table.cells || []).forEach(cell => {
      const r = cell.rowIndex;
      const c = cell.columnIndex;
      if (r < rowCount && c < colCount) {
        grid[r][c]     = cell.content || '';
        isHeader[r][c] = cell.kind === 'columnHeader' || cell.kind === 'rowHeader';
        spans[r][c]    = { r: cell.rowSpan || 1, c: cell.columnSpan || 1 };
      }
    });

    // Build HTML table
    const wrap = document.createElement('div');
    wrap.className = 'table-wrap';

    const label = document.createElement('div');
    label.className = 'table-label';
    label.textContent = `Table ${tableIdx + 1} — ${rowCount} rows × ${colCount} columns`;
    wrap.appendChild(label);

    const tableEl = document.createElement('table');
    tableEl.className = 'doc-table';

    for (let r = 0; r < rowCount; r++) {
      const row = document.createElement('tr');
      for (let c = 0; c < colCount; c++) {
        if (grid[r][c] === null) continue; // merged cell

        const tag  = isHeader[r][c] ? 'th' : 'td';
        const cell = document.createElement(tag);
        cell.textContent = grid[r][c] || '';

        const { r: rs, c: cs } = spans[r][c];
        if (rs > 1) cell.rowSpan = rs;
        if (cs > 1) cell.colSpan = cs;

        // Mark cells consumed by this span as null so we skip them
        for (let dr = 0; dr < rs; dr++) {
          for (let dc = 0; dc < cs; dc++) {
            if (dr === 0 && dc === 0) continue;
            if (r + dr < rowCount && c + dc < colCount) {
              grid[r + dr][c + dc] = null;
            }
          }
        }

        row.appendChild(cell);
      }
      tableEl.appendChild(row);
    }

    wrap.appendChild(tableEl);
    container.appendChild(wrap);
  });
}

/* ═══════════════════════════════════════════════════════════════
   RESULTS ACTIONS (copy / download)
═══════════════════════════════════════════════════════════════ */
function initResults() {
  el('copyTextBtn').addEventListener('click', copyText);
  el('downloadTextBtn').addEventListener('click', downloadText);
  el('dismissErrorBtn').addEventListener('click', hideError);
}

async function copyText() {
  if (!State.extractedText) { toast('No text to copy.', 'error'); return; }
  try {
    await navigator.clipboard.writeText(State.extractedText);
    toast('Text copied to clipboard!', 'success');
  } catch {
    // Fallback for older browsers
    const ta = document.createElement('textarea');
    ta.value = State.extractedText;
    ta.style.position = 'fixed';
    ta.style.opacity  = '0';
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    toast('Text copied!', 'success');
  }
}

function downloadText() {
  if (!State.extractedText) { toast('No text to download.', 'error'); return; }
  const filename = State.currentFile
    ? State.currentFile.name.replace(/\.[^.]+$/, '') + '_extracted.txt'
    : `docuscan_${Date.now()}.txt`;
  const blob = new Blob([State.extractedText], { type: 'text/plain;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
  toast(`Downloaded as "${filename}"`, 'success');
}

/* ═══════════════════════════════════════════════════════════════
   PROGRESS & STATUS UI
═══════════════════════════════════════════════════════════════ */
function showProgress(show) {
  el('progressTrack').hidden = !show;
  if (show) el('progressTrack').removeAttribute('aria-hidden');
  else      el('progressTrack').setAttribute('aria-hidden', 'true');
}

function setProgressBar(pct) {
  const bar = el('progressBar');
  bar.style.width           = Math.min(100, pct) + '%';
  bar.setAttribute('aria-valuenow', String(pct));
}

/** state: 'active' | 'done' | '' */
function setStep(stepId, state) {
  const step = el(`step-${stepId}`);
  if (!step) return;
  step.classList.remove('active', 'done');
  if (state) step.classList.add(state);

  // Activate connector before this step
  const steps   = ['upload', 'process', 'extract', 'done'];
  const idx     = steps.indexOf(stepId);
  const connectors = qsa('.progress-connector');
  if (idx > 0 && connectors[idx - 1]) {
    connectors[idx - 1].classList.remove('active', 'done');
    if (state === 'done')   connectors[idx - 1].classList.add('done');
    if (state === 'active') connectors[idx - 1].classList.add('active');
  }
}

function resetSteps() {
  qsa('.progress-step').forEach(s => s.classList.remove('active', 'done'));
  qsa('.progress-connector').forEach(c => c.classList.remove('active', 'done'));
}

function setStatus(cls, text) {
  const chip = el('statusChip');
  chip.className = `status-chip ${cls} visible`;

  // Add spinner for in-progress states
  const inProgress = ['uploading', 'processing', 'extracting'].includes(cls);
  chip.innerHTML = inProgress
    ? `<span class="spinner" aria-hidden="true"></span>${escHtml(text)}`
    : escHtml(text);
}

/* ═══════════════════════════════════════════════════════════════
   SHOW / HIDE HELPERS
═══════════════════════════════════════════════════════════════ */
function hideResults() {
  el('resultsSection').hidden = true;
  el('textOutput').textContent = '';
  el('tablesOutput').innerHTML = '';
}

function showError(message) {
  el('errorCard').hidden    = false;
  el('errorMessage').textContent = message;
  el('errorCard').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function hideError() {
  el('errorCard').hidden = true;
}

/* ═══════════════════════════════════════════════════════════════
   TOAST NOTIFICATIONS
═══════════════════════════════════════════════════════════════ */
function toast(message, type = 'info') {
  const icons   = { success: '✓', error: '✕', info: 'ℹ' };
  const container = el('toastContainer');

  const t = document.createElement('div');
  t.className = `toast ${type}`;
  t.innerHTML = `<span class="toast-icon">${icons[type] || 'ℹ'}</span><span>${escHtml(message)}</span>`;
  container.appendChild(t);

  // Auto-remove after 4s
  setTimeout(() => {
    t.style.opacity   = '0';
    t.style.transform = 'translateX(20px)';
    t.style.transition = 'opacity 0.3s, transform 0.3s';
    setTimeout(() => t.remove(), 300);
  }, 4000);
}

/* ═══════════════════════════════════════════════════════════════
   PARTICLE BACKGROUND (canvas)
═══════════════════════════════════════════════════════════════ */
function initParticles() {
  const canvas = el('particleCanvas');
  if (!canvas) return;
  const ctx    = canvas.getContext('2d');
  let particles = [];
  let animId;

  function resize() {
    canvas.width  = window.innerWidth;
    canvas.height = window.innerHeight;
  }
  resize();
  window.addEventListener('resize', resize);

  // Create particles
  function createParticles() {
    particles = [];
    const count = Math.min(60, Math.floor((window.innerWidth * window.innerHeight) / 20000));
    for (let i = 0; i < count; i++) {
      particles.push({
        x:    Math.random() * canvas.width,
        y:    Math.random() * canvas.height,
        r:    Math.random() * 1.5 + 0.3,
        vx:   (Math.random() - 0.5) * 0.3,
        vy:   (Math.random() - 0.5) * 0.3,
        alpha: Math.random() * 0.5 + 0.1,
      });
    }
  }
  createParticles();

  function draw() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Get accent color from CSS variable
    const accentColor = getComputedStyle(document.documentElement)
      .getPropertyValue('--accent').trim() || '#f4a261';

    particles.forEach(p => {
      // Move
      p.x += p.vx;
      p.y += p.vy;

      // Wrap
      if (p.x < 0)           p.x = canvas.width;
      if (p.x > canvas.width) p.x = 0;
      if (p.y < 0)           p.y = canvas.height;
      if (p.y > canvas.height) p.y = 0;

      // Draw
      ctx.beginPath();
      ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
      ctx.fillStyle = accentColor;
      ctx.globalAlpha = p.alpha;
      ctx.fill();
    });

    ctx.globalAlpha = 1;

    // Draw connecting lines between nearby particles
    for (let i = 0; i < particles.length; i++) {
      for (let j = i + 1; j < particles.length; j++) {
        const dx   = particles[i].x - particles[j].x;
        const dy   = particles[i].y - particles[j].y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist < 120) {
          ctx.beginPath();
          ctx.moveTo(particles[i].x, particles[i].y);
          ctx.lineTo(particles[j].x, particles[j].y);
          ctx.strokeStyle = accentColor;
          ctx.globalAlpha = (1 - dist / 120) * 0.12;
          ctx.lineWidth   = 0.5;
          ctx.stroke();
          ctx.globalAlpha = 1;
        }
      }
    }

    animId = requestAnimationFrame(draw);
  }

  // Pause when tab is hidden (battery savings)
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) {
      cancelAnimationFrame(animId);
    } else {
      draw();
    }
  });

  draw();
}

/* ═══════════════════════════════════════════════════════════════
   ACCESSIBILITY
═══════════════════════════════════════════════════════════════ */
function setupAccessibility() {
  // Keyboard support for drop zone already in initFileUpload
  // Make sure all interactive elements have visible focus
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      if (!el('errorCard').hidden) hideError();
    }
  });
}

/* ═══════════════════════════════════════════════════════════════
   UTILITIES
═══════════════════════════════════════════════════════════════ */
/** Format bytes into human-readable string */
function formatBytes(bytes) {
  if (bytes < 1024)        return bytes + ' B';
  if (bytes < 1048576)     return (bytes / 1024).toFixed(1) + ' KB';
  if (bytes < 1073741824)  return (bytes / 1048576).toFixed(1) + ' MB';
  return (bytes / 1073741824).toFixed(2) + ' GB';
}

/** Escape HTML special chars */
function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}

/** Promise-based delay */
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
