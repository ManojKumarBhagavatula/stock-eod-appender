// ── State ──────────────────────────────────────────────────────────────────
let activeMode       = 0;   // 0=EOD, 1=Live, 2=Date, 3=Range
let selectedFile     = null;
let downloadBlob     = null;
let downloadFilename = '';

const modeConfig = [
  {
    endpoint:  '/upload',
    btnText:   '📅 &nbsp;Fetch Today\'s EOD Prices',
    btnCls:    'b0',
    strip:     '<strong>Today\'s EOD</strong> — Appends today\'s end-of-day closing prices as a new column. Best used after 3:30 PM IST.',
    stripCls:  'c0',
    loadTitle: 'Fetching today\'s EOD prices…',
  },
  {
    endpoint:  '/get-current',
    btnText:   '⚡ &nbsp;Fetch Live Prices Now',
    btnCls:    'b1',
    strip:     '<strong>Live Price</strong> — Fetches the latest available market price for each stock right now.',
    stripCls:  'c1',
    loadTitle: 'Fetching live prices…',
  },
  {
    endpoint:  '/get-date',
    btnText:   '🗓️ &nbsp;Fetch Closing Price for Selected Date',
    btnCls:    'b2',
    strip:     '<strong>Specific Date</strong> — Fetches the closing price for every stock on the date you pick.',
    stripCls:  'c2',
    loadTitle: 'Fetching prices for selected date…',
  },
  {
    endpoint:  '/get-range',
    btnText:   '📆 &nbsp;Fetch Prices for Date Range',
    btnCls:    'b3',
    strip:     '<strong>Date Range</strong> — Appends one column per trading day between your From and To dates. Max 90 days.',
    stripCls:  'c3',
    loadTitle: 'Fetching range prices (this may take longer)…',
  },
];

// ── DOM refs ───────────────────────────────────────────────────────────────
const fileInput      = document.getElementById('fileInput');
const dropzone       = document.getElementById('dropzone');
const filePill       = document.getElementById('filePill');
const fetchBtn       = document.getElementById('fetchBtn');
const loadingBox     = document.getElementById('loadingBox');
const loadingTitle   = document.getElementById('loadingTitle');
const progressFill   = document.getElementById('progressFill');
const resultBox      = document.getElementById('resultBox');
const resultHeader   = document.getElementById('resultHeader');
const resultBody     = document.getElementById('resultBody');
const modeStrip      = document.getElementById('modeStrip');
const dateSection    = document.getElementById('dateSection');
const singleDateWrap = document.getElementById('singleDateWrap');
const rangeDateWrap  = document.getElementById('rangeDateWrap');
const singleDate     = document.getElementById('singleDate');
const fromDate       = document.getElementById('fromDate');
const toDate         = document.getElementById('toDate');
const singleWarn     = document.getElementById('singleWarn');
const step2label     = document.getElementById('step2label');

// Set max dates to today
const todayStr = new Date().toISOString().split('T')[0];
singleDate.max = todayStr;
fromDate.max   = todayStr;
toDate.max     = todayStr;

// ── Tab switching ──────────────────────────────────────────────────────────
function switchTab(idx) {
  activeMode = idx;
  document.querySelectorAll('.tab-btn').forEach((b, i) => {
    b.classList.toggle('active', i === idx);
  });

  const cfg = modeConfig[idx];
  modeStrip.innerHTML = cfg.strip;
  modeStrip.className = 'mode-strip ' + cfg.stripCls;

  fetchBtn.className = 'btn ' + cfg.btnCls;
  fetchBtn.innerHTML = cfg.btnText;

  dateSection.style.display    = idx >= 2 ? 'block' : 'none';
  singleDateWrap.style.display = idx === 2 ? 'block' : 'none';
  rangeDateWrap.style.display  = idx === 3 ? 'block' : 'none';

  step2label.textContent = idx >= 2 ? 'Step 3 — Fetch prices' : 'Step 2 — Fetch prices';

  resultBox.style.display  = 'none';
  loadingBox.style.display = 'none';
  validateBtn();
}

// ── File select ────────────────────────────────────────────────────────────
fileInput.addEventListener('change', e => {
  const f = e.target.files[0];
  if (f) selectFile(f);
});

function selectFile(f) {
  selectedFile = f;
  filePill.style.display = 'block';
  filePill.textContent   = '✅  ' + f.name + '  (' + (f.size / 1024).toFixed(1) + ' KB)';
  resultBox.style.display = 'none';
  validateBtn();
}

dropzone.addEventListener('dragover',  e => { e.preventDefault(); dropzone.classList.add('drag-over'); });
dropzone.addEventListener('dragleave', ()  => dropzone.classList.remove('drag-over'));
dropzone.addEventListener('drop', e => {
  e.preventDefault(); dropzone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f && (f.name.endsWith('.xlsx') || f.name.endsWith('.xlsm'))) selectFile(f);
  else alert('Please upload an .xlsx or .xlsm file.');
});

// ── Weekend warnings ───────────────────────────────────────────────────────
function isWeekend(dateStr) {
  if (!dateStr) return false;
  const d = new Date(dateStr);
  return d.getUTCDay() === 0 || d.getUTCDay() === 6;
}

singleDate.addEventListener('change', () => {
  singleWarn.style.display = isWeekend(singleDate.value) ? 'block' : 'none';
  validateBtn();
});

fromDate.addEventListener('change', validateBtn);
toDate.addEventListener('change',   validateBtn);

function validateBtn() {
  let ok = !!selectedFile;
  if (activeMode === 2) ok = ok && !!singleDate.value;
  if (activeMode === 3) ok = ok && !!fromDate.value && !!toDate.value && fromDate.value <= toDate.value;
  fetchBtn.disabled = !ok;
}

// ── Progress bar ───────────────────────────────────────────────────────────
let progressTimer;

function startFakeProgress() {
  let pct = 5;
  progressFill.style.width      = pct + '%';
  progressFill.style.background = getComputedStyle(document.documentElement).getPropertyValue('--accent2');
  progressTimer = setInterval(() => {
    const step = pct < 40 ? 2 : pct < 70 ? 1 : pct < 88 ? 0.4 : 0.05;
    pct = Math.min(pct + step, 92);
    progressFill.style.width = pct + '%';
  }, 1000);
}

function stopProgress(ok) {
  clearInterval(progressTimer);
  progressFill.style.width = ok ? '100%' : '0%';
}

// ── Fetch ──────────────────────────────────────────────────────────────────
async function startFetch() {
  if (!selectedFile) return;
  if (activeMode === 2 && isWeekend(singleDate.value)) {
    showError('Selected date is a weekend — markets are closed. Please pick a weekday.');
    return;
  }

  fetchBtn.disabled        = true;
  loadingBox.style.display = 'block';
  resultBox.style.display  = 'none';
  loadingTitle.textContent = modeConfig[activeMode].loadTitle;
  startFakeProgress();

  const form     = new FormData();
  const endpoint = modeConfig[activeMode].endpoint;
  form.append('file', selectedFile);
  if (activeMode === 2) form.append('target_date', singleDate.value);
  if (activeMode === 3) { form.append('from_date', fromDate.value); form.append('to_date', toDate.value); }

  try {
    const resp = await fetch(endpoint, { method: 'POST', body: form });
    stopProgress(resp.ok);

    if (resp.ok) {
      const blob        = await resp.blob();
      const dDate       = resp.headers.get('X-Stats-Date')         || '';
      const fetched     = resp.headers.get('X-Stats-Fetched')      || '?';
      const failed      = resp.headers.get('X-Stats-Failed')       || '?';
      const total       = resp.headers.get('X-Stats-Total')        || '?';
      const cols        = resp.headers.get('X-Stats-Columns')      || '1';
      const tradingDays = resp.headers.get('X-Stats-Trading-Days') || '';
      const rangeFrom   = resp.headers.get('X-Stats-From')         || '';
      const rangeTo     = resp.headers.get('X-Stats-To')           || '';

      downloadBlob     = blob;
      const prefix     = ['eod', 'live', 'date', 'range'][activeMode];
      downloadFilename = `stocks_${prefix}_${dDate.replace(/[^a-z0-9]/gi, '_')}.xlsx`;

      showSuccess(dDate, fetched, failed, total, cols, tradingDays, rangeFrom, rangeTo);
    } else {
      const err = await resp.json().catch(() => ({ detail: 'Unknown server error.' }));
      showError(err.detail || 'Something went wrong.');
    }
  } catch (e) {
    stopProgress(false);
    showError('Could not connect to the server. Make sure the app is running.');
  }

  loadingBox.style.display = 'none';
  fetchBtn.disabled        = false;
}

// ── NSE Holiday list ───────────────────────────────────────────────────────
// Format: 'YYYY-MM-DD': 'Occasion name'
const NSE_HOLIDAYS = {
  // 2024
  '2024-01-26': 'Republic Day',
  '2024-03-08': 'Mahashivratri',
  '2024-03-25': 'Holi',
  '2024-03-29': 'Good Friday',
  '2024-04-11': 'Id-Ul-Fitr (Ramzan Eid)',
  '2024-04-14': 'Dr. Baba Saheb Ambedkar Jayanti',
  '2024-04-17': 'Shri Ram Navami',
  '2024-04-21': 'Shri Mahavir Jayanti',
  '2024-05-01': 'Maharashtra Day',
  '2024-06-17': 'Bakri Id',
  '2024-07-17': 'Muharram',
  '2024-08-15': 'Independence Day',
  '2024-10-02': 'Mahatma Gandhi Jayanti',
  '2024-11-01': 'Diwali Laxmi Pujan (Muhurat Trading)',
  '2024-11-15': 'Gurunanak Jayanti',
  '2024-12-25': 'Christmas',
  // 2025
  '2025-01-26': 'Republic Day',
  '2025-02-26': 'Mahashivratri',
  '2025-03-14': 'Holi',
  '2025-03-31': 'Id-Ul-Fitr (Ramzan Eid)',
  '2025-04-10': 'Shri Ram Navami',
  '2025-04-14': 'Dr. Baba Saheb Ambedkar Jayanti',
  '2025-04-18': 'Good Friday',
  '2025-05-01': 'Maharashtra Day',
  '2025-08-15': 'Independence Day',
  '2025-08-27': 'Ganesh Chaturthi',
  '2025-10-02': 'Mahatma Gandhi Jayanti / Dussehra',
  '2025-10-20': 'Diwali Laxmi Pujan (Muhurat Trading)',
  '2025-10-21': 'Diwali Balipratipada',
  '2025-10-23': 'Diwali Balipratipada (extra)',
  '2025-11-05': 'Prakash Gurpurb Sri Guru Nanak Dev Ji',
  '2025-12-25': 'Christmas',
  // 2026
  '2026-01-26': 'Republic Day',
  '2026-03-20': 'Holi',
  '2026-04-03': 'Good Friday',
  '2026-04-14': 'Dr. Baba Saheb Ambedkar Jayanti / Visu',
  '2026-04-15': 'Shri Ram Navami',
  '2026-05-01': 'Maharashtra Day / Id-Ul-Fitr',
  '2026-08-15': 'Independence Day',
  '2026-10-02': 'Mahatma Gandhi Jayanti',
  '2026-10-26': 'Diwali Muhurat Trading',
  '2026-11-25': 'Gurunanak Jayanti',
  '2026-12-25': 'Christmas',
};

// ── Calendar builder ───────────────────────────────────────────────────────
function buildCalendar(tradingDaysStr, rangeFrom, rangeTo) {
  if (!rangeFrom || !rangeTo) return '';

  const tradingSet = new Set(tradingDaysStr ? tradingDaysStr.split(',') : []);

  const [fy, fm, fd] = rangeFrom.split('-').map(Number);
  const [ty, tm, td] = rangeTo.split('-').map(Number);
  const start = new Date(fy, fm - 1, fd);
  const end   = new Date(ty, tm - 1, td);

  const days = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    days.push(new Date(d));
  }

  let html = '<div class="cal-wrap"><div class="cal-label">📅 Date Breakdown</div><div class="cal-grid">';

  for (const d of days) {
    const dow     = d.getDay();
    const iso     = d.toISOString().split('T')[0];
    const display = d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short' });
    const dayName = d.toLocaleDateString('en-IN', { weekday: 'short' });

    if (dow === 0 || dow === 6) {
      html += `<div class="cal-day weekend" title="Weekend — market closed">
        <span class="cd-date">${display}</span>
        <span class="cd-day">${dayName}</span>
        <span class="cd-tag wk">Weekend</span>
      </div>`;
    } else if (tradingSet.has(iso)) {
      html += `<div class="cal-day trading" title="Market open — data fetched">
        <span class="cd-date">${display}</span>
        <span class="cd-day">${dayName}</span>
        <span class="cd-tag tr">✓ Fetched</span>
      </div>`;
    } else {
      const reason = NSE_HOLIDAYS[iso] || 'NSE Market Holiday';
      html += `<div class="cal-day holiday" title="${reason}">
        <span class="cd-date">${display}</span>
        <span class="cd-day">${dayName}</span>
        <span class="cd-tag hol">🏛 ${reason}</span>
      </div>`;
    }
  }

  html += '</div></div>';
  return html;
}

// ── Result display ─────────────────────────────────────────────────────────
function showSuccess(dDate, fetched, failed, total, cols, tradingDays, rangeFrom, rangeTo) {
  const failedNum = parseInt(failed, 10);
  resultBox.className     = 'result-box success';
  resultBox.style.display = 'block';
  resultHeader.innerHTML  = '✅ &nbsp;Done! Prices fetched successfully.';

  const colRow = parseInt(cols) > 1
    ? `<div class="stat-row"><span>Trading day columns added</span><span class="stat-val g">${cols}</span></div>`
    : `<div class="stat-row"><span>Column added</span><span class="stat-val g">${dDate}</span></div>`;

  const calHtml = activeMode === 3
    ? buildCalendar(tradingDays, rangeFrom, rangeTo)
    : '';

  resultBody.innerHTML =
    colRow +
    `<div class="stat-row"><span>Total stocks processed</span><span class="stat-val w">${total}</span></div>` +
    `<div class="stat-row"><span>Prices fetched</span><span class="stat-val g">${fetched}</span></div>` +
    `<div class="stat-row"><span>Could not fetch (N/A)</span><span class="stat-val ${failedNum > 0 ? 'y' : 'g'}">${failed}</span></div>` +
    calHtml +
    `<button class="btn-dl" onclick="triggerDownload()">⬇ &nbsp;Download Updated Excel File</button>`;
}

function showError(msg) {
  resultBox.className     = 'result-box error';
  resultBox.style.display = 'block';
  resultHeader.innerHTML  = '❌ &nbsp;Something went wrong';
  resultBody.innerHTML    = `<div style="color:var(--red); font-size:0.83rem;">${msg}</div>`;
}

// ── Download ───────────────────────────────────────────────────────────────
function triggerDownload() {
  if (!downloadBlob) return;
  const url = URL.createObjectURL(downloadBlob);
  const a   = document.createElement('a');
  a.href = url; a.download = downloadFilename; a.click();
  URL.revokeObjectURL(url);
}