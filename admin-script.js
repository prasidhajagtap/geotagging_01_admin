/*
 * ══════════════════════════════════════════════════════════════
 *  SEAMEX GEO-ATTENDANCE  |  admin-script.js  |  v01
 *  Admin Console — All Logic
 *
 *  Author : Prasidha Jagtap
 *  Role   : Assistant Manager – IT, Aditya Birla Group (Seamex)
 *  Office : Reliable Tech Park, Airoli, Maharashtra
 *
 *  Hello, future developer. 👋
 *  Maintained by Prasidha Jagtap.
 *  This file controls the entire admin console.
 *  Do not modify without understanding the data flow.
 *
 *  FUNCTION INDEX
 *  ─────────────────────────────────────────────────────────────
 *  INIT           DOMContentLoaded — checks session, boots UI
 *  AUTH           checkLogin(), adminLogout(), togglePwd()
 *  THEME          toggleAdminTheme()
 *  DATA FETCH     fetchAllRecords() — reads from Supabase
 *  TAB CONTROL    switchTab(), navDate(), getDateRange()
 *  FILTER         applyFilters(), resetFilters(), filterRecord()
 *  KPI            renderKPIs()
 *  CHARTS         renderAllCharts(), renderBar(), renderLine(),
 *                 renderPie(), renderHeatmap()
 *  TABLES         renderSummaryTable(), renderDetailTable(),
 *                 sortSummary(), sortDetail()
 *  EXPORT         exportCSV(), exportXLSX()
 *  HELPERS        g, setTx, hide, show, fmt, fmtDate, msDur,
 *                 duration, toHours, adminToast
 *
 *  SECURITY NOTES (Prasidha)
 *  ─────────────────────────────────────────────────────────────
 *  · Password is hashed with SHA-256 before comparison.
 *    Plaintext never stored. Hash stored in sessionStorage.
 *    IMPORTANT: This is still client-side security. A determined
 *    attacker who reads the JS can extract the hash. The real
 *    protection is Supabase RLS — the anon key only reads data
 *    that RLS allows. Sensitive fields like coordinates are
 *    visible here because admin needs full data access.
 *  · For production: protect this page behind Azure AD and move
 *    Supabase reads to an Edge Function with service role.
 *
 *  SUPABASE SELECT POLICY REQUIRED (Prasidha)
 *  Run this SQL in Supabase Dashboard → SQL Editor:
 *  ──────────────────────────────────────────────
 *  CREATE POLICY "Admin can read attendance"
 *  ON attendance FOR SELECT TO anon
 *  USING (true);
 *  ──────────────────────────────────────────────
 *  Without this policy, fetchAllRecords() returns zero rows.
 *
 *  MIGRATION NOTES (Prasidha)
 *  ─────────────────────────────────────────────────────────────
 *  [ ] Production: replace client-side password with Azure AD
 *      protected route (MSAL + role check).
 *  [ ] SharePoint: move this to a protected SPFx webpart.
 *  [ ] Multi-office: add branch_code filter to all queries.
 * ══════════════════════════════════════════════════════════════
 */

/* ── SUPABASE CONFIG ────────────────────────────────────────
   Prasidha: Same anon key as field app. Safe because RLS controls
   what anon can read. Admin SELECT policy enables read access.
   service_role key must NEVER be used on the client.
────────────────────────────────────────────────────────────── */
const SUPABASE_URL = 'https://svhbqvcabbzrxvndxtjm.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InN2aGJxdmNhYmJ6cnh2bmR4dGptIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzUyMTA0MjksImV4cCI6MjA5MDc4NjQyOX0.lYIsM5zN4uGKbP79avcKR_EaAlP5tu2N688OgZI6wZA';
const _db = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

/* ── ADMIN PASSWORD ─────────────────────────────────────────
   Prasidha: Password is stored as a SHA-256 hash here.
   To change the password:
     1. Open browser console (any tab)
     2. Run: crypto.subtle.digest('SHA-256', new TextEncoder().encode('YourNewPassword'))
             .then(b => console.log([...new Uint8Array(b)].map(x=>x.toString(16).padStart(2,'0')).join('')))
     3. Replace ADMIN_HASH below with the output string
   Current password: Seamex@2025Admin  (change before deployment — Prasidha)
────────────────────────────────────────────────────────────── */
const ADMIN_HASH ='c1c224b03cd9bc7b6a86d77f5dace40191766c485cd55dc48caf9ac873335d6f';
/* NOTE (Prasidha): The hash above is a placeholder. Run the snippet above
   with your real password to generate the correct SHA-256 hash. */

const SESSION_KEY = 'smx_admin_session';

/* ── GLOBAL STATE ────────────────────────────────────────────
   Prasidha: Single source of truth for admin console state.
────────────────────────────────────────────────────────────── */
let allRecords      = [];   /* All records from Supabase           */
let filteredRecords = [];   /* After filters applied               */
let currentTab      = 'daily';
let currentDate     = new Date();  /* Reference date for navigation */
let sortSummaryCol  = 'clock_in'; let sortSummaryAsc = false;
let sortDetailCol   = 'clock_in_time'; let sortDetailAsc = false;
let chartBar   = null;
let chartLine  = null;
let chartPie   = null;
let adminTheme = localStorage.getItem('smx_admin_theme') || 'light';

/* Chart.js colour palette — Prasidha: aligned with Seamex brand */
const C = {
  red:    '#A6192E',
  orange: '#F58220',
  yellow: '#FFCB05',
  green:  '#16a34a',
  blue:   '#2563eb',
  amber:  '#d97706',
  redA:   'rgba(166,25,46,.15)',
  orangeA:'rgba(245,130,32,.15)',
  greenA: 'rgba(22,163,74,.15)',
  blueA:  'rgba(37,99,235,.15)',
};

/* ══════════════════════════════════════════════════════════════
   HELPERS — hoisted functions (Prasidha: function declarations,
   not const, so they are available before their position in file)
*/

/** g — Prasidha: shorthand getElementById */
function g(id) { return document.getElementById(id); }

/** setTx — Prasidha: safe textContent setter */
function setTx(id, v) { const e = g(id); if (e) e.textContent = v; }

/** hide — Prasidha: adds .hidden */
function hide(id) { g(id)?.classList.add('hidden'); }

/** show — Prasidha: removes .hidden */
function show(id) { g(id)?.classList.remove('hidden'); }

/**
 * fmt — Prasidha
 * Formats ISO timestamp as readable time (HH:MM AM/PM, en-IN).
 */
function fmt(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleTimeString('en-IN', {
    hour: '2-digit', minute: '2-digit', hour12: true
  });
}

/**
 * fmtDate — Prasidha
 * Formats ISO timestamp as readable date.
 */
function fmtDate(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleDateString('en-IN', {
    day: 'numeric', month: 'short', year: 'numeric'
  });
}

/**
 * fmtDateShort — Prasidha: dd MMM
 */
function fmtDateShort(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleDateString('en-IN', {
    day: 'numeric', month: 'short'
  });
}

/**
 * msDur — Prasidha
 * Converts millisecond duration to HH:MM:SS string.
 */
function msDur(ms) {
  if (!ms || isNaN(ms) || ms < 0) return '—';
  const h = Math.floor(ms / 3600000);
  const m = Math.floor((ms % 3600000) / 60000);
  const s = Math.floor((ms % 60000) / 1000);
  return [h, m, s].map(n => n.toString().padStart(2, '0')).join(':');
}

/**
 * toHours — Prasidha
 * Converts ms to decimal hours (2 decimal places).
 */
function toHours(ms) {
  if (!ms || isNaN(ms) || ms < 0) return 0;
  return Math.round(ms / 3600000 * 100) / 100;
}

/**
 * duration — Prasidha
 * Duration between two ISO timestamps.
 */
function duration(inISO, outISO) {
  if (!inISO || !outISO) return 0;
  return new Date(outISO) - new Date(inISO);
}

/**
 * dateToYMD — Prasidha
 * Converts Date to YYYY-MM-DD string (local timezone).
 */
function dateToYMD(d) {
  return d.getFullYear() + '-' +
    String(d.getMonth() + 1).padStart(2, '0') + '-' +
    String(d.getDate()).padStart(2, '0');
}

/**
 * adminToast — Prasidha
 * Top-right toast notification for admin actions.
 * @param {string} msg
 * @param {'ok'|'err'|''} type
 */
function adminToast(msg, type = '') {
  const box = g('admin-toast');
  if (!box) return;
  const t = document.createElement('div');
  t.className = 'atst' + (type ? ' ' + type : '');
  t.textContent = msg;
  box.appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transition = 'opacity .3s'; }, 1800);
  setTimeout(() => t.remove(), 2200);
}

/* ══════════════════════════════════════════════════════════════
   INIT
   Prasidha: Checks sessionStorage for existing admin session.
   If authenticated, boots dashboard. Otherwise shows login.
*/
document.addEventListener('DOMContentLoaded', () => {
  /* Apply saved theme */
  document.documentElement.setAttribute('data-theme', adminTheme);
  updateThemeBtn();

  /* Check existing session */
  const sess = sessionStorage.getItem(SESSION_KEY);
  if (sess === 'authenticated') {
    bootDashboard();
  } else {
    show('login-screen');
  }

  /* Login button */
  g('login-btn').addEventListener('click', checkLogin);
  g('login-pwd').addEventListener('keydown', e => {
    if (e.key === 'Enter') checkLogin();
  });

  /* Eye toggle */
  g('eye-btn').addEventListener('click', () => {
    const inp = g('login-pwd');
    inp.type = inp.type === 'password' ? 'text' : 'password';
  });
});

/* ══════════════════════════════════════════════════════════════
   AUTH — Prasidha
*/

/**
 * checkLogin — Prasidha
 * Hashes the entered password with SHA-256 and compares to ADMIN_HASH.
 * SECURITY: plaintext never stored or transmitted.
 */
async function checkLogin() {
  const pwd = g('login-pwd').value;
  if (!pwd) {
    setTx('login-err', 'Please enter the admin password.');
    return;
  }

  try {
    /* SHA-256 hash in browser — Prasidha */
    const encoded = new TextEncoder().encode(pwd);
    const hashBuf = await crypto.subtle.digest('SHA-256', encoded);
    const hashHex = [...new Uint8Array(hashBuf)]
      .map(b => b.toString(16).padStart(2, '0'))
      .join('');

    if (hashHex === ADMIN_HASH) {
      setTx('login-err', '');
      sessionStorage.setItem(SESSION_KEY, 'authenticated');
      hide('login-screen');
      bootDashboard();
    } else {
      setTx('login-err', 'Incorrect password. Please try again.');
      g('login-pwd').value = '';
      g('login-pwd').focus();
    }
  } catch (e) {
    /* Fallback if crypto.subtle unavailable (HTTP, non-secure context) */
    console.error('[Seamex Admin|Prasidha] Hash error:', e);
    setTx('login-err', 'Password check failed. Use HTTPS.');
  }
}

/**
 * adminLogout — Prasidha
 * Clears session and reloads to show login screen.
 */
function adminLogout() {
  sessionStorage.removeItem(SESSION_KEY);
  location.reload();
}

/* ══════════════════════════════════════════════════════════════
   THEME — Prasidha
*/
function toggleAdminTheme() {
  adminTheme = adminTheme === 'light' ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', adminTheme);
  localStorage.setItem('smx_admin_theme', adminTheme);
  updateThemeBtn();
  /* Redraw charts to pick up new colours */
  renderAllCharts();
}

function updateThemeBtn() {
  const btn = g('nav-theme-btn');
  if (btn) btn.textContent = adminTheme === 'dark' ? '☀️ Light' : '🌙 Dark';
}

/* ══════════════════════════════════════════════════════════════
   BOOT DASHBOARD — Prasidha
   Shows shell, sets default filters, fetches data.
*/
function bootDashboard() {
  g('admin-shell').classList.add('visible');

  /* Set default date filters to current month */
  const today = new Date();
  const y = today.getFullYear(), m = today.getMonth();
  g('f-from').value = dateToYMD(new Date(y, m, 1));
  g('f-to').value   = dateToYMD(today);

  /* Initial tab + date label */
  updateDateLabel();
  fetchAllRecords();

  /* Auto-refresh every 60 seconds — Prasidha */
  setInterval(fetchAllRecords, 60000);
}

/* ══════════════════════════════════════════════════════════════
   DATA FETCH — Prasidha
   Reads all records from Supabase attendance table.
   REQUIRES: SELECT RLS policy enabled (see file header).
   Orders by clock_in_time descending (newest first).
*/
async function fetchAllRecords() {
  try {
    const { data, error } = await _db
      .from('attendance')
      .select('*')
      .order('clock_in_time', { ascending: false });

    if (error) {
      /* SECURITY: Never log raw Supabase error to user (Prasidha) */
      console.error('[Seamex Admin|Prasidha] Supabase fetch error:', error);
      adminToast('Data fetch failed. Check RLS policy.', 'err');
      return;
    }

    allRecords = data || [];
    setTx('nav-refresh', `Refreshed ${fmt(new Date().toISOString())}`);

    applyFilters();
    adminToast(`${allRecords.length} records loaded.`, 'ok');

  } catch (e) {
    console.error('[Seamex Admin|Prasidha] Network error:', e);
    adminToast('Network error. Check connection.', 'err');
  }
}

/* ══════════════════════════════════════════════════════════════
   TAB CONTROL — Prasidha
   Switches between Daily / Weekly / Monthly views.
   Date navigator steps match the active tab period.
*/

/**
 * switchTab — Prasidha
 * @param {'daily'|'weekly'|'monthly'} tab
 */
function switchTab(tab) {
  currentTab = tab;
  currentDate = new Date();

  /* Update active tab button */
  document.querySelectorAll('.tab-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.tab === tab);
  });

  updateDateLabel();
  applyFilters();
}

/**
 * navDate — Prasidha
 * Steps the date navigator forward or backward by one period.
 * @param {1|-1} dir
 */
function navDate(dir) {
  const d = new Date(currentDate);
  if (currentTab === 'daily')   d.setDate(d.getDate() + dir);
  if (currentTab === 'weekly')  d.setDate(d.getDate() + dir * 7);
  if (currentTab === 'monthly') d.setMonth(d.getMonth() + dir);
  currentDate = d;
  updateDateLabel();
  applyFilters();
}

/**
 * updateDateLabel — Prasidha
 * Updates the date navigator label text.
 */
function updateDateLabel() {
  const d = currentDate;
  let label = '';

  if (currentTab === 'daily') {
    const today = new Date();
    const diff  = Math.round((d.setHours(0,0,0,0) - today.setHours(0,0,0,0)) / 86400000);
    label = diff === 0 ? 'Today'
          : diff === -1 ? 'Yesterday'
          : d.toLocaleDateString('en-IN', { day:'numeric', month:'short', year:'numeric' });
    currentDate = new Date(currentDate); /* restore time */
  }

  if (currentTab === 'weekly') {
    const { start, end } = getDateRange();
    label = `${start.toLocaleDateString('en-IN',{day:'numeric',month:'short'})} – ${end.toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'numeric'})}`;
  }

  if (currentTab === 'monthly') {
    label = currentDate.toLocaleDateString('en-IN', { month: 'long', year: 'numeric' });
  }

  setTx('date-label', label);
}

/**
 * getDateRange — Prasidha
 * Returns { start: Date, end: Date } for the current tab + date.
 */
function getDateRange() {
  const d = new Date(currentDate);
  let start, end;

  if (currentTab === 'daily') {
    start = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0);
    end   = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59);
  }

  if (currentTab === 'weekly') {
    /* Week starts Monday — Prasidha */
    const day  = d.getDay(); /* 0=Sun … 6=Sat */
    const diff = (day === 0 ? -6 : 1 - day);
    start = new Date(d);
    start.setDate(d.getDate() + diff);
    start.setHours(0, 0, 0, 0);
    end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);
  }

  if (currentTab === 'monthly') {
    start = new Date(d.getFullYear(), d.getMonth(), 1, 0, 0, 0);
    end   = new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59);
  }

  return { start, end };
}

/* ══════════════════════════════════════════════════════════════
   FILTER — Prasidha
   Combines tab date range with user-applied filters.
   applyFilters() is the main render trigger.
*/

/**
 * applyFilters — Prasidha
 * Filters allRecords and triggers all renderers.
 */
function applyFilters() {
  const search    = g('f-search')?.value.toLowerCase().trim() || '';
  const fromDate  = g('f-from')?.value   || '';
  const toDate    = g('f-to')?.value     || '';
  const status    = g('f-status')?.value || '';
  const threshold = g('f-threshold')?.value || '09:30';

  /* Update late threshold KPI sub label */
  setTx('kpi-late-thresh', threshold);

  /* Get tab date range */
  const { start, end } = getDateRange();

  filteredRecords = allRecords.filter(r => {
    /* Tab date range filter */
    const ciDate = r.clock_in_time ? new Date(r.clock_in_time) : null;
    if (!ciDate || ciDate < start || ciDate > end) return false;

    /* Additional date range filters from inputs */
    if (fromDate) {
      const from = new Date(fromDate + 'T00:00:00');
      if (ciDate < from) return false;
    }
    if (toDate) {
      const to = new Date(toDate + 'T23:59:59');
      if (ciDate > to) return false;
    }

    /* Search filter */
    if (search) {
      const nameMatch = (r.user_name || '').toLowerCase().includes(search);
      const idMatch   = (r.employee_id || '').includes(search);
      if (!nameMatch && !idMatch) return false;
    }

    /* Status filter */
    if (status) {
      if (status === 'completed' && r.status !== 'completed') return false;
      if (status === 'missing_out' && r.clock_out_time) return false;
      if (status === 'late') {
        const [th, tm] = threshold.split(':').map(Number);
        const ciHour = ciDate.getHours(), ciMin = ciDate.getMinutes();
        const isLate = ciHour > th || (ciHour === th && ciMin > tm);
        if (!isLate) return false;
      }
    }

    return true;
  });

  /* Update record count */
  setTx('rec-count', filteredRecords.length.toString());

  /* Render everything */
  renderKPIs(filteredRecords, threshold);
  renderAllCharts();
  renderSummaryTable(filteredRecords, threshold);
  renderDetailTable(filteredRecords, threshold);
}

/**
 * resetFilters — Prasidha
 * Clears all filter inputs and re-applies.
 */
function resetFilters() {
  g('f-search').value    = '';
  g('f-status').value    = '';
  g('f-threshold').value = '09:30';
  const today = new Date();
  const y = today.getFullYear(), m = today.getMonth();
  g('f-from').value = dateToYMD(new Date(y, m, 1));
  g('f-to').value   = dateToYMD(today);
  currentDate = new Date();
  updateDateLabel();
  applyFilters();
}

/* ══════════════════════════════════════════════════════════════
   KPI CARDS — Prasidha
   Calculates four metrics from the filtered record set.
*/

/**
 * renderKPIs — Prasidha
 * @param {Array} records - Filtered attendance records
 * @param {string} threshold - Late clock-in threshold HH:MM
 */
function renderKPIs(records, threshold = '09:30') {
  /* 1. Attendance rate — unique employees present / all-time unique employees */
  const uniquePresent = new Set(records.map(r => r.employee_id)).size;
  const uniqueTotal   = new Set(allRecords.map(r => r.employee_id)).size;
  const rate = uniqueTotal > 0
    ? Math.round(uniquePresent / uniqueTotal * 100) + '%'
    : '—';
  setTx('kpi-rate', rate);
  setTx('kpi-rate-sub', `${uniquePresent} of ${uniqueTotal} employees`);

  /* 2. Total shift hours */
  const totalMs = records.reduce((sum, r) => {
    return sum + (r.clock_out_time ? duration(r.clock_in_time, r.clock_out_time) : 0);
  }, 0);
  const totalH = toHours(totalMs);
  setTx('kpi-hours', totalH > 0 ? totalH + 'h' : '—');
  setTx('kpi-hours-sub', `${records.filter(r=>r.clock_out_time).length} completed shifts`);

  /* 3. Late clock-ins */
  const [th, tm] = threshold.split(':').map(Number);
  const lateCount = records.filter(r => {
    if (!r.clock_in_time) return false;
    const d = new Date(r.clock_in_time);
    return d.getHours() > th || (d.getHours() === th && d.getMinutes() > tm);
  }).length;
  setTx('kpi-late', lateCount.toString());

  /* 4. Missing clock-outs */
  const missingCount = records.filter(r => !r.clock_out_time).length;
  setTx('kpi-missing', missingCount.toString());
}

/* ══════════════════════════════════════════════════════════════
   CHARTS — Prasidha
   Chart.js v4. Existing chart instances destroyed before redraw.
   All charts derive data from filteredRecords (or allRecords for heatmap).
*/

/** getCssVar — Prasidha: reads CSS custom property value */
function getCssVar(v) {
  return getComputedStyle(document.documentElement).getPropertyValue(v).trim();
}

/** chartDefaults — Prasidha: shared Chart.js theme defaults */
function chartDefaults() {
  const tx2 = getCssVar('--tx2') || '#475569';
  const bdr  = getCssVar('--bdr') || '#e2e8f0';
  return {
    color: tx2,
    plugins: {
      legend: {
        labels: { color: tx2, font: { family: 'DM Sans', size: 11 } }
      },
      tooltip: {
        backgroundColor: 'rgba(15,23,42,.92)',
        titleFont: { family: 'DM Sans', size: 12 },
        bodyFont:  { family: 'DM Mono', size: 11 },
        padding: 10, cornerRadius: 8
      }
    },
    scales: {
      x: {
        ticks:  { color: tx2, font: { family: 'DM Sans', size: 11 } },
        grid:   { color: bdr }
      },
      y: {
        ticks:  { color: tx2, font: { family: 'DM Mono', size: 11 } },
        grid:   { color: bdr },
        beginAtZero: true
      }
    }
  };
}

/**
 * renderAllCharts — Prasidha
 * Destroys and redraws all four charts.
 */
function renderAllCharts() {
  renderBarChart();
  renderLineChart();
  renderPieChart();
  renderHeatmap();
}

/**
 * renderBarChart — Prasidha
 * Bar chart: total hours worked per day for the period.
 */
function renderBarChart() {
  const canvas = g('chart-bar');
  if (!canvas) return;

  /* Group hours by date */
  const byDate = {};
  filteredRecords.forEach(r => {
    if (!r.clock_in_time) return;
    const d = dateToYMD(new Date(r.clock_in_time));
    byDate[d] = (byDate[d] || 0) + toHours(duration(r.clock_in_time, r.clock_out_time));
  });

  const labels  = Object.keys(byDate).sort();
  const data    = labels.map(d => Math.round(byDate[d] * 100) / 100);
  const display = labels.map(l => fmtDateShort(l + 'T00:00:00'));

  if (chartBar) chartBar.destroy();

  const def = chartDefaults();
  chartBar = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: display,
      datasets: [{
        label: 'Hours Worked',
        data,
        backgroundColor: C.orangeA,
        borderColor:     C.orange,
        borderWidth:     1.5,
        borderRadius:    4
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: def.plugins,
      scales: def.scales
    }
  });
}

/**
 * renderLineChart — Prasidha
 * Line chart: number of employees who attended each day.
 */
function renderLineChart() {
  const canvas = g('chart-line');
  if (!canvas) return;

  /* Count unique employees per day */
  const byDate = {};
  filteredRecords.forEach(r => {
    if (!r.clock_in_time) return;
    const d = dateToYMD(new Date(r.clock_in_time));
    if (!byDate[d]) byDate[d] = new Set();
    byDate[d].add(r.employee_id);
  });

  const labels  = Object.keys(byDate).sort();
  const data    = labels.map(d => byDate[d].size);
  const display = labels.map(l => fmtDateShort(l + 'T00:00:00'));

  if (chartLine) chartLine.destroy();

  const def = chartDefaults();
  chartLine = new Chart(canvas, {
    type: 'line',
    data: {
      labels: display,
      datasets: [{
        label: 'Employees Present',
        data,
        borderColor:           C.red,
        backgroundColor:       C.redA,
        tension:               .35,
        fill:                  true,
        pointBackgroundColor:  C.red,
        pointRadius:           3,
        pointHoverRadius:      5
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: def.plugins,
      scales: def.scales
    }
  });
}

/**
 * renderPieChart — Prasidha
 * Doughnut: completed shifts vs missing clock-outs.
 */
function renderPieChart() {
  const canvas = g('chart-pie');
  if (!canvas) return;

  const completed = filteredRecords.filter(r => r.clock_out_time).length;
  const missing   = filteredRecords.filter(r => !r.clock_out_time).length;
  const late      = (() => {
    const [th, tm] = (g('f-threshold')?.value || '09:30').split(':').map(Number);
    return filteredRecords.filter(r => {
      if (!r.clock_in_time) return false;
      const d = new Date(r.clock_in_time);
      return d.getHours() > th || (d.getHours() === th && d.getMinutes() > tm);
    }).length;
  })();

  if (chartPie) chartPie.destroy();

  const tx2 = getCssVar('--tx2') || '#475569';
  chartPie = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels: ['Completed', 'Missing Out', 'Late In'],
      datasets: [{
        data: [completed, missing, late],
        backgroundColor: [C.green, C.red, C.amber],
        borderWidth: 0,
        hoverOffset: 6
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: '65%',
      plugins: {
        legend: { position: 'bottom', labels: { color: tx2, font: { family: 'DM Sans', size: 11 }, padding: 14 } },
        tooltip: { backgroundColor:'rgba(15,23,42,.92)', padding:10, cornerRadius:8,
          bodyFont:{family:'DM Mono',size:11}, titleFont:{family:'DM Sans',size:12} }
      }
    }
  });
}

/**
 * renderHeatmap — Prasidha
 * Attendance heatmap by day of week.
 * Uses CSS grid instead of Chart.js (no plugin dependency).
 */
function renderHeatmap() {
  const container = g('heatmap-container');
  if (!container) return;

  const DAY_LABELS = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
  const counts = Array(7).fill(0); /* index 0=Mon … 6=Sun */

  /* Use allRecords for heatmap (broader picture) — Prasidha */
  allRecords.forEach(r => {
    if (!r.clock_in_time) return;
    let day = new Date(r.clock_in_time).getDay(); /* 0=Sun */
    day = day === 0 ? 6 : day - 1; /* remap to Mon=0 */
    counts[day]++;
  });

  const maxCount = Math.max(...counts, 1);

  /* Heatmap intensity level (0–5) */
  const level = n => {
    if (n === 0) return 0;
    const ratio = n / maxCount;
    if (ratio < .2) return 1;
    if (ratio < .4) return 2;
    if (ratio < .6) return 3;
    if (ratio < .8) return 4;
    return 5;
  };

  container.innerHTML = `
    <div class="heatmap-wrap">
      <div class="hm-day-labels">
        ${DAY_LABELS.map(d => `<div class="hm-day-lbl">${d}</div>`).join('')}
      </div>
      <div class="heatmap-grid" style="grid-template-columns:repeat(7,22px)">
        ${counts.map((c, i) => `
          <div class="heatmap-cell hm-${level(c)}"
            title="${DAY_LABELS[i]}: ${c} clock-in${c !== 1 ? 's' : ''}">
          </div>
        `).join('')}
      </div>
      <div class="heatmap-legend">
        <span>Less</span>
        ${[0,1,2,3,4,5].map(l => `<div class="hm-leg-cell hm-${l}"></div>`).join('')}
        <span>More</span>
      </div>
    </div>
  `;
}

/* ══════════════════════════════════════════════════════════════
   TABLES — Prasidha
*/

/**
 * recordStatus — Prasidha
 * Returns { label, cls } for badge based on record analysis.
 */
function recordStatus(r, threshold = '09:30') {
  if (!r.clock_out_time) return { label: 'No Clock-Out', cls: 'badge-red' };
  const [th, tm] = threshold.split(':').map(Number);
  if (r.clock_in_time) {
    const d = new Date(r.clock_in_time);
    if (d.getHours() > th || (d.getHours() === th && d.getMinutes() > tm)) {
      return { label: 'Late', cls: 'badge-amber' };
    }
  }
  return { label: 'Completed', cls: 'badge-green' };
}

/**
 * renderSummaryTable — Prasidha
 * One row per attendance record (summary view).
 * Sorted by sortSummaryCol.
 */
function renderSummaryTable(records, threshold = '09:30') {
  const tbody = g('summary-tbody');
  if (!tbody) return;

  if (!records.length) {
    tbody.innerHTML = `<tr><td colspan="9">
      <div class="empty-state">
        <div class="empty-icon">📋</div>
        <div class="empty-title">No records for this period</div>
        <div class="empty-sub">Try adjusting the filters or date range.</div>
      </div>
    </td></tr>`;
    return;
  }

  /* Sort */
  const sorted = [...records].sort((a, b) => {
    let va, vb;
    switch (sortSummaryCol) {
      case 'name':       va = a.user_name||''; vb = b.user_name||''; break;
      case 'employee_id':va = a.employee_id||''; vb = b.employee_id||''; break;
      case 'date':
      case 'clock_in':   va = a.clock_in_time||''; vb = b.clock_in_time||''; break;
      case 'clock_out':  va = a.clock_out_time||''; vb = b.clock_out_time||''; break;
      case 'duration':
        va = duration(a.clock_in_time, a.clock_out_time);
        vb = duration(b.clock_in_time, b.clock_out_time);
        break;
      case 'location_in':  va = a.clock_in_location_name||'';  vb = b.clock_in_location_name||''; break;
      case 'location_out': va = a.clock_out_location_name||''; vb = b.clock_out_location_name||''; break;
      default: va = a.clock_in_time||''; vb = b.clock_in_time||'';
    }
    if (va < vb) return sortSummaryAsc ? -1 : 1;
    if (va > vb) return sortSummaryAsc ? 1 : -1;
    return 0;
  });

  /* Build rows — SECURITY: textContent via row construction, no innerHTML with user data */
  tbody.innerHTML = '';
  sorted.forEach(r => {
    const dur = duration(r.clock_in_time, r.clock_out_time);
    const st  = recordStatus(r, threshold);
    const tr  = document.createElement('tr');

    const cells = [
      { cls: 'name-col', txt: r.user_name || '—' },
      { cls: 'emp-id',   txt: r.employee_id || '—' },
      { cls: 'mono',     txt: r.clock_in_time ? fmtDate(r.clock_in_time) : '—' },
      { cls: 'mono',     txt: fmt(r.clock_in_time) },
      { cls: 'mono',     txt: fmt(r.clock_out_time) },
      { cls: 'mono',     txt: dur > 0 ? msDur(dur) : '—' },
      { cls: '',         txt: r.clock_in_location_name || '—' },
      { cls: '',         txt: r.clock_out_location_name || '—' },
    ];

    cells.forEach(c => {
      const td = document.createElement('td');
      if (c.cls) td.className = c.cls;
      td.textContent = c.txt;
      tr.appendChild(td);
    });

    /* Status badge cell */
    const tdSt = document.createElement('td');
    const badge = document.createElement('span');
    badge.className = `badge ${st.cls}`;
    badge.textContent = st.label;
    tdSt.appendChild(badge);
    tr.appendChild(tdSt);

    tbody.appendChild(tr);
  });

  setTx('detail-count', `${sorted.length} records`);
}

/**
 * renderDetailTable — Prasidha
 * Full raw records table with all fields.
 */
function renderDetailTable(records, threshold = '09:30') {
  const tbody = g('detail-tbody');
  if (!tbody) return;

  if (!records.length) {
    tbody.innerHTML = `<tr><td colspan="9">
      <div class="empty-state">
        <div class="empty-icon">🗂️</div>
        <div class="empty-title">No records match the current filters</div>
      </div>
    </td></tr>`;
    return;
  }

  /* Sort */
  const sorted = [...records].sort((a, b) => {
    const va = a[sortDetailCol] || '';
    const vb = b[sortDetailCol] || '';
    if (va < vb) return sortDetailAsc ? -1 : 1;
    if (va > vb) return sortDetailAsc ? 1 : -1;
    return 0;
  });

  tbody.innerHTML = '';
  sorted.forEach(r => {
    const dur = duration(r.clock_in_time, r.clock_out_time);
    const st  = recordStatus(r, threshold);
    const tr  = document.createElement('tr');

    const cells = [
      r.user_name || '—',
      r.employee_id || '—',
      r.clock_in_time  ? `${fmtDate(r.clock_in_time)} ${fmt(r.clock_in_time)}`  : '—',
      r.clock_in_location_name  || '—',
      r.clock_out_time ? `${fmtDate(r.clock_out_time)} ${fmt(r.clock_out_time)}` : '—',
      r.clock_out_location_name || '—',
      dur > 0 ? msDur(dur) : '—',
    ];

    cells.forEach((txt, i) => {
      const td = document.createElement('td');
      if ([2, 4, 6].includes(i)) td.className = 'mono';
      td.textContent = txt;
      tr.appendChild(td);
    });

    /* Status badge */
    const tdSt = document.createElement('td');
    const badge = document.createElement('span');
    badge.className = `badge ${st.cls}`;
    badge.textContent = st.label;
    tdSt.appendChild(badge);
    tr.appendChild(tdSt);

    /* Submitted At */
    const tdCr = document.createElement('td');
    tdCr.className = 'mono';
    tdCr.textContent = r.created_at ? fmtDate(r.created_at) : '—';
    tr.appendChild(tdCr);

    tbody.appendChild(tr);
  });
}

/* Sort handlers */
function sortSummary(col) {
  sortSummaryAsc = sortSummaryCol === col ? !sortSummaryAsc : true;
  sortSummaryCol = col;
  renderSummaryTable(filteredRecords, g('f-threshold')?.value || '09:30');
}

function sortDetail(col) {
  sortDetailAsc = sortDetailCol === col ? !sortDetailAsc : true;
  sortDetailCol = col;
  renderDetailTable(filteredRecords, g('f-threshold')?.value || '09:30');
}

/* ══════════════════════════════════════════════════════════════
   EXPORT — Prasidha

   exportCSV:  Downloads a CSV file using a data URI.
   exportXLSX: Downloads an Excel file using SheetJS (XLSX lib).
   Both functions accept `exportAll` — if true, exports the full
   allRecords dataset. If false, exports filteredRecords only.

   SECURITY: No PII is removed from exports intentionally —
   admin has full read access. Coordinate fields are included.
   Restrict physical access to this admin page accordingly.
*/

/**
 * buildExportData — Prasidha
 * Converts records to flat array of objects for export.
 */
function buildExportData(records) {
  const threshold = g('f-threshold')?.value || '09:30';
  return records.map(r => {
    const dur = duration(r.clock_in_time, r.clock_out_time);
    const st  = recordStatus(r, threshold);
    return {
      'Employee Name':        r.user_name         || '',
      'Poornata ID':          r.employee_id        || '',
      'Date':                 r.clock_in_time ? fmtDate(r.clock_in_time) : '',
      'Clock In Time':        fmt(r.clock_in_time),
      'Clock In Location':    r.clock_in_location_name  || '',
      'Clock In Coordinates': r.clock_in_coords   || '',
      'Clock Out Time':       fmt(r.clock_out_time),
      'Clock Out Location':   r.clock_out_location_name || '',
      'Clock Out Coordinates':r.clock_out_coords  || '',
      'Shift Duration':       dur > 0 ? msDur(dur) : '',
      'Shift Hours (decimal)':dur > 0 ? toHours(dur) : '',
      'Status':               st.label,
      'Submitted At':         r.created_at ? fmtDate(r.created_at) : '',
      'Record ID':            r.id || ''
    };
  });
}

/**
 * exportCSV — Prasidha
 * @param {boolean} exportAll - true = all records, false = filtered
 */
function exportCSV(exportAll = false) {
  const data   = buildExportData(exportAll ? allRecords : filteredRecords);
  if (!data.length) { adminToast('No data to export.', 'err'); return; }

  const headers = Object.keys(data[0]);
  const rows    = [headers, ...data.map(r => headers.map(h => `"${(r[h] || '').toString().replace(/"/g,'""')}"` ))];
  const csv     = rows.map(r => r.join(',')).join('\n');

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `seamex_attendance_${exportAll ? 'all' : 'filtered'}_${dateToYMD(new Date())}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  adminToast(`CSV exported (${data.length} records).`, 'ok');
}

/**
 * exportXLSX — Prasidha
 * Uses SheetJS to generate a styled Excel workbook.
 * @param {boolean} exportAll - true = all records, false = filtered
 */
function exportXLSX(exportAll = false) {
  if (typeof XLSX === 'undefined') {
    adminToast('Excel library not loaded. Check CDN.', 'err'); return;
  }

  const data = buildExportData(exportAll ? allRecords : filteredRecords);
  if (!data.length) { adminToast('No data to export.', 'err'); return; }

  /* Build workbook — Prasidha */
  const wb  = XLSX.utils.book_new();
  const ws  = XLSX.utils.json_to_sheet(data);

  /* Column widths */
  ws['!cols'] = [
    { wch: 22 }, { wch: 14 }, { wch: 16 }, { wch: 12 },
    { wch: 22 }, { wch: 26 }, { wch: 12 }, { wch: 22 },
    { wch: 26 }, { wch: 14 }, { wch: 18 }, { wch: 14 },
    { wch: 16 }, { wch: 36 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Attendance');

  /* Summary sheet — Prasidha */
  const threshold   = g('f-threshold')?.value || '09:30';
  const [th, tm]    = threshold.split(':').map(Number);
  const summaryData = [
    { Metric: 'Export Date', Value: fmtDate(new Date().toISOString()) },
    { Metric: 'Total Records', Value: data.length },
    { Metric: 'Total Employees', Value: new Set(data.map(r => r['Poornata ID'])).size },
    { Metric: 'Total Shift Hours', Value: toHours(data.reduce((s, r) => s + (r['Shift Hours (decimal)'] || 0) * 3600000, 0)) },
    { Metric: 'Late Clock-Ins', Value: data.filter(r => r['Status'] === 'Late').length },
    { Metric: 'Missing Clock-Outs', Value: data.filter(r => r['Status'] === 'No Clock-Out').length },
    { Metric: 'Late Threshold', Value: threshold },
    { Metric: 'Exported By', Value: 'Seamex Admin Console' },
    { Metric: 'Module', Value: 'Geo Attendance v07' },
    { Metric: 'Author', Value: 'Prasidha Jagtap — IT, Aditya Birla Group (Seamex)' }
  ];
  const wsSummary = XLSX.utils.json_to_sheet(summaryData);
  wsSummary['!cols'] = [{ wch: 26 }, { wch: 44 }];
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  XLSX.writeFile(wb, `seamex_attendance_${exportAll ? 'all' : 'filtered'}_${dateToYMD(new Date())}.xlsx`);
  adminToast(`Excel exported (${data.length} records).`, 'ok');
}

/*
 * ══════════════════════════════════════════════════════════════
 *  Prasidha Jagtap | IT · Aditya Birla Group (Seamex)
 *  Admin Console v01 — Geo Attendance Module
 *  Built for IT and HR administrators.
 *  Maintained with care. Keep the standards.
 * ══════════════════════════════════════════════════════════════
 */
