// AFAB x SAM comparison viewer — reads data.json + codes.json produced by the
// backend (GitHub Actions). No computation here; display only.

const COLS = [
  { key: 'Commission no.', label: 'Commission' },
  { key: 'Model(WINGS)', label: 'Model (WINGS)' },
  { key: 'Vehicle', label: 'Vehicle' },
  { key: 'Type', label: 'Type' },
  { key: 'Cab', label: 'Cab' },
  { key: 'PTO', label: 'PTO' },
  { key: 'Changeability Date', label: 'Changeability' },
  { key: 'Until Dealine', label: 'Changeability D-Day', dday: true },
  { key: 'Only_in_SAM', label: 'Only in SAM', count: 'sam' },
  { key: 'Only_in_WINGS', label: 'Only in WINGS', count: 'win' },
  { key: 'Mandatory Codes', label: 'Mandatory', count: 'mand' },
  { key: 'SAM Status', label: 'Status', status: true },
];

const NUMERIC_KEYS = new Set(['Until Dealine', 'Baumuster',
  'Only_in_SAM', 'Only_in_WINGS', 'Mandatory Codes']);

// Friendly labels for the detail card / Excel export (data keys stay unchanged).
const META_LABELS = {
  'Changeability Date': 'Changeability',
  'Until Dealine': 'Changeability D-Day',
};
// Keys whose value should render as a D-day (D-9 / D-Day / D+9).
const DDAY_KEYS = new Set(['Until Dealine']);

// current month "YYYY-MM" (browser today) — used for the "이번달 이후" filter.
const _now = new Date();
const CUR_MONTH = _now.getFullYear() + '-' + String(_now.getMonth() + 1).padStart(2, '0');

let DATA = { rows: [] };
let CODES = { options: {}, mandatory: {} };
let sortKey = null, sortDir = 1;

const $ = (s) => document.querySelector(s);

// ---- HTML escaping (rows come from .docx; never trust raw) ----
function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ---- init ----
function initTheme() {
  const saved = localStorage.getItem('theme');
  if (saved) document.documentElement.dataset.theme = saved;
  syncThemeBtn();
}
function syncThemeBtn() {
  $('#themeBtn').textContent = document.documentElement.dataset.theme === 'dark' ? '☀' : '🌙';
}
function toggleTheme() {
  const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  document.documentElement.dataset.theme = next;
  localStorage.setItem('theme', next);
  syncThemeBtn();
}

async function load() {
  const [d, c] = await Promise.all([
    fetch('data.json?_=' + Date.now()).then((r) => r.json()),
    fetch('codes.json?_=' + Date.now()).then((r) => r.json()).catch(() => CODES),
  ]);
  DATA = d; CODES = c;
  renderMeta();
  renderSummary();
  fillVehicleFilter();
  fillProductionFilter();
  renderHead();
  render();
}

function prodMonth(r) {
  return String(r['Production date'] || '').slice(0, 7);  // YYYY-MM
}
function changeMonth(r) {
  const s = String(r['Changeability Date'] || '');
  return /^\d{4}-\d{2}/.test(s) ? s.slice(0, 7) : '';
}

function renderMeta() {
  const when = DATA.generated_at ? new Date(DATA.generated_at).toLocaleString('ko-KR') : '-';
  $('#meta').textContent = `생성: ${when}  ·  WINGS: ${DATA.wings_file || '-'}`;
}

function dashStats(rows) {
  return {
    total: rows.length,
    mismatch: rows.filter((r) => r['SAM Status'] === 'Mismatch').length,
    match: rows.filter((r) => r['SAM Status'] === 'Match').length,
    mand: rows.filter((r) => countOf(r['Mandatory Codes']) > 0).length,
  };
}

function within2weeks(r) {
  const n = Number(r['Until Dealine']);
  return !Number.isNaN(n) && n >= 0 && n <= 14;
}

function tile(cls, n, label) {
  return `<div class="tile ${cls}"><div class="n">${esc(n)}</div><div class="l">${label}</div></div>`;
}

function renderSummary() {
  const all = dashStats(DATA.rows);
  const soon = dashStats(DATA.rows.filter(within2weeks));
  $('#summary').innerHTML = `
    <div class="dash-row">
      <div class="dash-cap">전체 현황</div>
      <div class="tiles">
        ${tile('t-total', all.total, '전체 Commission')}
        ${tile('t-miss', all.mismatch, '미스매치')}
        ${tile('t-match', all.match, '매칭')}
        ${tile('t-mand', all.mand, 'Mandatory 누락')}
      </div>
    </div>
    <div class="dash-row">
      <div class="dash-cap">2주 이내 (Changeability D-14)</div>
      <div class="tiles">
        ${tile('t-total2', soon.total, '전체 Commission')}
        ${tile('t-miss2', soon.mismatch, '미스매치')}
        ${tile('t-match2', soon.match, '매칭')}
        ${tile('t-mand2', soon.mand, 'Mandatory 누락')}
      </div>
    </div>`;
}

function fmtMonth(ym) {
  const s = String(ym);
  return s.length === 6 ? `${s.slice(0, 4)}-${s.slice(4)}` : s;
}

function fillVehicleFilter() {
  const sel = $('#vehicleFilter');
  const vehicles = [...new Set(DATA.rows.map((r) => r.Vehicle).filter(Boolean))].sort();
  sel.innerHTML = '<option value="">전체 차종</option>' +
    vehicles.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
}

function fillProductionFilter() {
  const sel = $('#productionFilter');
  const months = [...new Set(DATA.rows.map(prodMonth).filter(Boolean))].sort().reverse();
  sel.innerHTML = '<option value="">전체 생산월</option>' +
    months.map((m) => `<option value="${esc(m)}">${esc(m)}</option>`).join('');
}

function renderHead() {
  $('#grid thead').innerHTML =
    '<tr>' + COLS.map((c) => {
      const arrow = sortKey === c.key ? `<span class="arrow">${sortDir === 1 ? '▲' : '▼'}</span>` : '';
      return `<th data-k="${esc(c.key)}">${c.label}${arrow}</th>`;
    }).join('') + '</tr>';
  $('#grid thead').querySelectorAll('th').forEach((th) =>
    th.addEventListener('click', () => {
      const k = th.dataset.k;
      sortDir = sortKey === k ? -sortDir : 1;
      sortKey = k;
      renderHead();
      render();
    })
  );
}

// Plain-text D-day: future = "D-9", today = "D-Day", passed = "D+9" (days since).
function ddayText(v) {
  const s = String(v ?? '').trim();
  if (!s) return '';
  if (s.toLowerCase() === 'passed') return '지남';
  const n = Number(s);
  if (!Number.isNaN(n)) {
    if (n < 0) return `D+${-n}`;
    if (n === 0) return 'D-Day';
    return `D-${n}`;
  }
  return s;
}

function ddayHtml(v) {
  const s = String(v ?? '').trim();
  if (!s) return '<span class="empty-cell">—</span>';
  if (s.toLowerCase() === 'passed') return '<span class="dday passed">지남</span>';
  const n = Number(s);
  if (!Number.isNaN(n)) {
    if (n < 0) return `<span class="dday passed">D+${-n}</span>`;
    const cls = n <= 14 ? 'soon' : 'ok';
    return `<span class="dday ${cls}">${n === 0 ? 'D-Day' : 'D-' + n}</span>`;
  }
  return esc(s);
}

// highlight current search term inside escaped text
function hl(text) {
  const q = $('#search').value.trim();
  const safe = esc(text);
  if (!q) return safe;
  try {
    const re = new RegExp('(' + q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
    return safe.replace(re, '<mark>$1</mark>');
  } catch { return safe; }
}

// Table cell: just the count, color-coded, clickable (opens the row drawer).
function countCell(csv, cls) {
  const n = countOf(csv);
  if (!n) return '<span class="cbadge ok" title="이상없음 (차이·누락 없음)">✓</span>';
  return `<span class="cbadge ${cls}">${n}</span>`;
}

function filtered() {
  const q = $('#search').value.trim().toLowerCase();
  const status = $('#statusFilter').value;
  const vehicle = $('#vehicleFilter').value;
  const prod = $('#productionFilter').value;
  const upcoming = $('#upcomingOnly').checked;
  const mmOnly = $('#mismatchOnly').checked;
  const ptoOnly = $('#ptoOnly').checked;
  let rows = DATA.rows.filter((r) => {
    if (status && r['SAM Status'] !== status) return false;
    if (vehicle && r.Vehicle !== vehicle) return false;
    if (prod && prodMonth(r) !== prod) return false;
    // "이번달 이후만": keep rows whose changeability month >= current month.
    if (upcoming) {
      const cm = changeMonth(r);
      if (!cm || cm < CUR_MONTH) return false;
    }
    if (ptoOnly && !String(r.PTO || '').trim()) return false;
    if (mmOnly && !(r['Only_in_SAM'] || r['Only_in_WINGS'])) return false;
    if (q) {
      const hay = Object.values(r).join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
  if (sortKey) {
    const isCount = COLS.some((c) => c.key === sortKey && c.count);
    const numeric = NUMERIC_KEYS.has(sortKey);
    rows = [...rows].sort((a, b) => {
      let cmp;
      if (isCount) {
        cmp = countOf(a[sortKey]) - countOf(b[sortKey]);
      } else {
        const av = a[sortKey] ?? '', bv = b[sortKey] ?? '';
        const an = Number(av), bn = Number(bv);
        const bothNum = numeric && !Number.isNaN(an) && !Number.isNaN(bn) && av !== '' && bv !== '';
        cmp = bothNum ? an - bn : String(av).localeCompare(String(bv));
      }
      return cmp * sortDir;
    });
  }
  return rows;
}

function render() {
  const rows = filtered();
  $('#count').textContent = `${rows.length} / ${DATA.rows.length} 건`;
  const tb = $('#grid tbody');
  const msg = $('#statusMsg');

  if (!DATA.rows.length) {
    tb.innerHTML = '';
    msg.textContent = '표시할 데이터가 없습니다. data.json 을 먼저 빌드하세요.';
    msg.classList.remove('hidden');
    return;
  }
  if (!rows.length) {
    tb.innerHTML = '';
    msg.textContent = '필터 조건에 맞는 항목이 없습니다.';
    msg.classList.remove('hidden');
    return;
  }
  msg.classList.add('hidden');

  tb.innerHTML = rows.map((r, i) => {
    const tds = COLS.map((c) => {
      let v = r[c.key];
      v = v == null ? '' : String(v);
      if (c.status) return `<td><span class="status ${esc(v).replace(/\s+/g, '')}">${esc(v)}</span></td>`;
      if (c.dday) return `<td class="num">${ddayHtml(v)}</td>`;
      if (c.count) return `<td class="num">${countCell(v, c.count)}</td>`;
      return `<td>${hl(v)}</td>`;
    }).join('');
    return `<tr data-i="${i}">${tds}</tr>`;
  }).join('');
  tb.querySelectorAll('tr').forEach((tr) =>
    tr.addEventListener('click', () => openDrawer(rows[Number(tr.dataset.i)]))
  );
}

function describe(code) {
  return CODES.options[code] || CODES.mandatory[code] || '';
}

function countOf(csv) {
  return (csv || '').split(',').map((c) => c.trim()).filter(Boolean).length;
}

function splitCodes(csv) {
  return (csv || '').split(',').map((c) => c.trim()).filter(Boolean);
}

let DRAWER_ROW = null;

function codeRowsHtml(codes, extraClass) {
  const cls = extraClass ? ' ' + extraClass : '';
  return codes.map((c) =>
    `<div class="code-row${cls}"><span class="c">${esc(c)}</span><span class="d">${esc(describe(c) || '—')}</span></div>`).join('');
}

// One side of a two-column code block: title + count + code/description rows.
// opts.diff=true → column represents a discrepancy (0 = good, mark with ✅).
// opts.okCodes → codes to show as "verified OK" when the discrepancy is 0.
// opts.okLabel / opts.okHint → texts shown in the 0 case.
function codeColHtml(title, csv, opts) {
  opts = opts || {};
  const codes = splitCodes(csv);
  const n = codes.length;
  let badge, body;
  if (n === 0) {
    if (opts.diff) {
      badge = '<span class="badge ok">✓ 0</span>';
      const ok = opts.okCodes || [];
      body = `<div class="ok-mark">${opts.okLabel || '✅ 이상없음 — 누락·차이 없음'}</div>`;
      if (ok.length) body += `<div class="code-list ok-list">${codeRowsHtml(ok, 'ok-row')}</div>`;
      else if (opts.okHint) body += `<div class="hint">${esc(opts.okHint)}</div>`;
    } else {
      badge = '<span class="badge">0</span>';
      body = '<div class="none">없음</div>';
    }
  } else {
    badge = `<span class="badge ${opts.diff ? 'warn' : ''}">${n}</span>`;
    body = `<div class="code-list">${codeRowsHtml(codes)}</div>`;
  }
  return `<div class="code-col"><h4>${esc(title)} ${badge}</h4>${body}</div>`;
}

// Full Code List: SAM vs WINGS aligned by code. Union of both sides, sorted, so
// the SAME code always sits on the SAME row. A code present on only one side
// leaves the other cell blank and shades it light blue (--accent-soft) to flag
// the gap (e.g. A0B only in WINGS -> blank, shaded SAM cell).
function alignedFullHtml(samCsv, wingsCsv) {
  const sam = new Set(splitCodes(samCsv));
  const wings = new Set(splitCodes(wingsCsv));
  const union = [...new Set([...sam, ...wings])].sort();
  const cell = (code, present) => present
    ? `<span class="c">${esc(code)}</span><span class="d">${esc(describe(code) || '—')}</span>`
    : '';
  const rows = union.map((code) => {
    const inS = sam.has(code), inW = wings.has(code);
    return `<div class="acode-row">`
      + `<div class="acode-cell${inS ? '' : ' miss'}">${cell(code, inS)}</div>`
      + `<div class="acode-cell${inW ? '' : ' miss'}">${cell(code, inW)}</div>`
      + `</div>`;
  }).join('');
  return `<div class="acode">
      <div class="acode-head">
        <h4>All SAM Codes <span class="badge">${sam.size}</span></h4>
        <h4>All WINGS Codes <span class="badge">${wings.size}</span></h4>
      </div>
      <div class="acode-body">${rows}</div>
    </div>`;
}

function openDrawer(r) {
  if (!r) return;
  DRAWER_ROW = r;
  $('#drawerTitle').textContent = `${r['Commission no.']}  ·  ${r['Model(WINGS)'] || ''}`;
  $('#drawerSub').textContent = r['Compared SAM file name'] || '';
  const meta = ['Vehicle', 'Type', 'Cab', 'PTO', 'Production date', 'Changeability Date',
    'Until Dealine', 'SAM Baumuster', 'SAM now',
    'Order status financial', 'SAM Status', 'FIN']
    .filter((k) => r[k] !== undefined && r[k] !== '')
    .map((k) => {
      const val = DDAY_KEYS.has(k) ? ddayText(r[k]) : r[k];
      return `<div class="k">${esc(META_LABELS[k] || k)}</div><div>${esc(val)}</div>`;
    }).join('');

  // Codes verified OK (present in both SAM & WINGS) — shown when a diff is 0.
  const allSam = new Set(splitCodes(r['_all_sam_codes']));
  const allWings = new Set(splitCodes(r['_all_wings_codes']));
  const mandSet = CODES.mandatory || {};
  const matched = [...allSam].filter((c) => allWings.has(c)).sort();
  const matchedMand = matched.filter((c) => c in mandSet);

  $('#drawerBody').innerHTML = `
    <div class="kv">${meta}</div>
    <div class="drawer-actions">
      <button id="drawerXls" class="icon-btn primary">⬇ 이 차량 Excel</button>
    </div>
    <div class="tabs">
      <button class="tab active" data-tab="diff">🔍 Difference Codes</button>
      <button class="tab" data-tab="full">📄 Full Code List</button>
    </div>
    <div class="tab-pane" data-pane="diff">
      <div class="code-cols">
        ${codeColHtml('Codes Only in SAM', r['Only_in_SAM'], { diff: true })}
        ${codeColHtml('Codes Only in WINGS', r['Only_in_WINGS'], { diff: true })}
      </div>
      <div class="code-cols" style="margin-top:18px">
        ${codeColHtml('Mandatory', r['Mandatory Codes'], { diff: true, okCodes: matchedMand, okLabel: '✅ 이상없음 — 아래 필수코드가 양쪽 모두 반영됨' })}
        ${codeColHtml('Factory Control', r['Factory Control Codes'])}
      </div>
    </div>
    <div class="tab-pane hidden" data-pane="full">
      ${alignedFullHtml(r['_all_sam_codes'], r['_all_wings_codes'])}
    </div>
  `;
  $('#drawerBody').querySelectorAll('.tab').forEach((t) =>
    t.addEventListener('click', () => switchTab(t.dataset.tab)));
  $('#drawerXls').addEventListener('click', () => exportRowXls(r));
  $('#drawer').classList.remove('hidden');
  $('#backdrop').classList.remove('hidden');
}

function switchTab(name) {
  $('#drawerBody').querySelectorAll('.tab').forEach((t) =>
    t.classList.toggle('active', t.dataset.tab === name));
  $('#drawerBody').querySelectorAll('.tab-pane').forEach((p) =>
    p.classList.toggle('hidden', p.dataset.pane !== name));
}

function closeDrawer() {
  $('#drawer').classList.add('hidden');
  $('#backdrop').classList.add('hidden');
}

// ---- CSV export of the currently filtered view ----
function exportCsv() {
  const rows = filtered();
  if (!rows.length) return;
  const keys = COLS.map((c) => c.key);
  const headers = COLS.map((c) => c.label);
  const cell = (v) => {
    const s = v == null ? '' : String(v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const lines = [headers.map(cell).join(',')];
  for (const r of rows) lines.push(keys.map((k) => cell(r[k])).join(','));
  const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `afab_sam_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ---- Excel export (dependency-free: HTML-table .xls that Excel opens natively) ----
const _today = () => new Date().toISOString().slice(0, 10);

function downloadXls(filename, sheetHtml) {
  const html =
    `<html xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:x="urn:schemas-microsoft-com:office:excel" ` +
    `xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8">` +
    `<style>td,th{border:1px solid #ccc;padding:3px 6px;mso-number-format:"\\@";} ` +
    `th{background:#f0f2f5;font-weight:bold;}</style></head><body>${sheetHtml}</body></html>`;
  const blob = new Blob(['﻿' + html], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function xc(v) {  // escaped cell text
  return esc(v == null ? '' : String(v));
}

// Export the currently filtered table to Excel.
function exportTableXls() {
  const rows = filtered();
  if (!rows.length) return;
  const head = '<tr>' + COLS.map((c) => `<th>${xc(c.label)}</th>`).join('') + '</tr>';
  const body = rows.map((r) => '<tr>' + COLS.map((c) => {
    if (c.count) return `<td>${countOf(r[c.key])}</td>`;       // numeric count
    return `<td>${xc(r[c.key])}</td>`;
  }).join('') + '</tr>').join('');
  downloadXls(`afab_sam_${_today()}.xls`, `<table>${head}${body}</table>`);
}

// Export one vehicle's full comparison (the detail view) to Excel.
function exportRowXls(r) {
  if (!r) return;
  const metaKeys = ['Commission no.', 'Model(WINGS)', 'Vehicle', 'Type', 'Cab', 'PTO',
    'Production date', 'Changeability Date', 'Until Dealine',
    'SAM Baumuster', 'SAM now', 'SAM Status', 'Compared SAM file name'];
  const metaTbl = '<table><tr><th colspan="2">Vehicle Info</th></tr>' +
    metaKeys.filter((k) => r[k] !== undefined && r[k] !== '')
      .map((k) => {
        const val = DDAY_KEYS.has(k) ? ddayText(r[k]) : r[k];
        return `<tr><th>${xc(META_LABELS[k] || k)}</th><td>${xc(val)}</td></tr>`;
      }).join('') + '</table>';

  const section = (title, csv) => {
    const codes = splitCodes(csv);
    const rowsHtml = codes.length
      ? codes.map((c) => `<tr><td>${xc(c)}</td><td>${xc(describe(c))}</td></tr>`).join('')
      : '<tr><td colspan="2"></td></tr>';
    return `<table><tr><th colspan="2">${xc(title)} (${codes.length})</th></tr>` +
      `<tr><th>Code</th><th>Description</th></tr>${rowsHtml}</table>`;
  };

  const sheet = metaTbl + '<br/>' +
    section('Codes Only in SAM', r['Only_in_SAM']) + '<br/>' +
    section('Codes Only in WINGS', r['Only_in_WINGS']) + '<br/>' +
    section('Mandatory', r['Mandatory Codes']) + '<br/>' +
    section('Factory Control', r['Factory Control Codes']) + '<br/>' +
    section('All SAM Codes', r['_all_sam_codes']) + '<br/>' +
    section('All WINGS Codes', r['_all_wings_codes']);
  const name = String(r['Commission no.'] || 'vehicle').replace(/[^\w.-]/g, '_');
  downloadXls(`afab_sam_${name}.xls`, sheet);
}

// ====================== Rules editor (model-conversion rules) ======================
// Matching is now data-driven: each SAM file carries both its numbers (파일제목=SAM now,
// 본문 Vehicle type=SAM Baumuster). The only editable rule is a manual alias override.
const RULE_MAPS = [];
const RULE_LISTS = [
  ['reverse_aliases', '매칭 별칭 (수동)', '파일 제목이 다른 세대 번호를 쓸 때만: 번호 → 추가 매칭 번호 (쉼표 구분)'],
];
let RULES_RAW = null;

function mapRowHtml(k, v) {
  return `<div class="rule-row">
    <input type="text" class="k-in" value="${esc(k)}" placeholder="키" />
    <span class="arrow-i">→</span>
    <input type="text" class="v-in" value="${esc(v)}" placeholder="값" />
    <button class="del" title="삭제">×</button></div>`;
}

function groupHtml(key, title, desc, type, obj) {
  const rows = Object.entries(obj || {}).map(([k, v]) =>
    mapRowHtml(k, type === 'list' ? (Array.isArray(v) ? v.join(', ') : v) : v)).join('');
  return `<details class="rule-group" open>
    <summary>${title} <span class="desc">${desc}</span></summary>
    <div class="rule-table" data-group="${key}" data-type="${type}">
      ${rows}<button class="rule-add">+ 행 추가</button>
    </div></details>`;
}

function renderRules(data) {
  RULES_RAW = data;
  let html = '';
  for (const [key, title, desc] of RULE_MAPS) html += groupHtml(key, title, desc, 'map', data[key]);
  for (const [key, title, desc] of RULE_LISTS) html += groupHtml(key, title, desc, 'list', data[key]);
  $('#rulesBody').innerHTML = html;
}

function collectRules() {
  const out = {};
  if (RULES_RAW && RULES_RAW._comment) out._comment = RULES_RAW._comment;
  $('#rulesBody').querySelectorAll('.rule-table').forEach((tbl) => {
    const g = tbl.dataset.group, type = tbl.dataset.type;
    const obj = {};
    tbl.querySelectorAll('.rule-row').forEach((row) => {
      const k = row.querySelector('.k-in').value.trim();
      const v = row.querySelector('.v-in').value.trim();
      if (!k) return;
      obj[k] = type === 'list' ? v.split(',').map((s) => s.trim()).filter(Boolean) : v;
    });
    out[g] = obj;
  });
  return out;
}

async function openRules() {
  try {
    const data = await fetch('rules.json?_=' + Date.now()).then((r) => r.json());
    renderRules(data);
  } catch (e) {
    $('#rulesBody').innerHTML = `<div class="none">rules.json 로드 실패: ${esc(e.message)}</div>`;
  }
  $('#rulesModal').classList.remove('hidden');
  $('#rulesBackdrop').classList.remove('hidden');
}
function closeRules() {
  $('#rulesModal').classList.add('hidden');
  $('#rulesBackdrop').classList.add('hidden');
}
// ---- xlsx <-> rules bridge (sheet names must match backend/rules.py & generator) ----
const ALIAS_SHEET = '매칭_별칭(수동)';

function rulesToSheets(data) {
  // Returns [ [sheetName, aoa], ... ]; aoa = array-of-arrays incl. header row.
  const sheets = [];
  // 1) recognition table from the currently-loaded grid (informational; matches backend cols)
  const ref = [['차종(Vehicle)', 'WINGS 모델', 'SAM Baumuster(원본)', 'SAM now(수정)', '매칭상태', 'PTO', 'SAM 파일']];
  const seen = new Set();
  for (const r of DATA.rows || []) {
    const w = String(r['Model(WINGS)'] || '').trim();
    if (!w) continue;
    const samf = String(r['Compared SAM file name'] || '').split(/[\\/]/).pop();
    const row = [String(r.Vehicle || '').trim(), w,
      String(r['SAM Baumuster'] || '').trim(), String(r['SAM now'] || '').trim(),
      String(r['SAM Status'] || '').trim(), String(r.PTO || '').trim(), samf];
    const key = row.join('|');
    if (seen.has(key)) continue;
    seen.add(key); ref.push(row);
  }
  sheets.push(['인식모델_대조표', ref]);
  // 2) editable manual alias sheet
  const aoa = [['번호(원본/WINGS)', '추가로 매칭할 번호(쉼표 구분)']];
  for (const [k, v] of Object.entries(data.reverse_aliases || {}))
    aoa.push([k, Array.isArray(v) ? v.join(', ') : v]);
  sheets.push([ALIAS_SHEET, aoa]);
  return sheets;
}

function sheetsToRules(wb) {
  const out = {};
  const ws = wb.Sheets[ALIAS_SHEET];
  if (ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false }).slice(1);
    const obj = {};
    for (const r of rows) {
      const k = String(r[0] ?? '').trim();
      if (k) obj[k] = String(r[1] ?? '').split(',').map((s) => s.trim()).filter(Boolean);
    }
    out.reverse_aliases = obj;
  }
  return out;
}

function downloadRules() {
  const data = collectRules();
  if (window.XLSX && !window.__XLSX_BLOCKED) {
    const wb = XLSX.utils.book_new();
    for (const [name, aoa] of rulesToSheets(data))
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
    XLSX.writeFile(wb, 'model_mapping.xlsx');
    return;
  }
  // Fallback: xlsx lib unavailable (e.g. CDN blocked) -> emit JSON so nothing is lost.
  alert('엑셀 라이브러리를 불러오지 못해 JSON(rules.json)으로 저장합니다.\n엑셀 편집은 model_rules/model_mapping.xlsx 를 직접 열어 진행하세요.');
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'rules.json';
  a.click();
  URL.revokeObjectURL(a.href);
}

function uploadRulesFile(file) {
  if (!file) return;
  if (!window.XLSX || window.__XLSX_BLOCKED) {
    alert('엑셀 라이브러리를 불러오지 못했습니다(사내망 차단 가능). model_rules/model_mapping.xlsx 를 엑셀에서 직접 편집하세요.');
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      renderRules(sheetsToRules(wb));
    } catch (err) {
      alert('엑셀 읽기 실패: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

$('#rulesBtn').addEventListener('click', openRules);
$('#rulesClose').addEventListener('click', closeRules);
$('#rulesBackdrop').addEventListener('click', closeRules);
$('#rulesReload').addEventListener('click', openRules);
$('#rulesDownload').addEventListener('click', downloadRules);
$('#rulesUploadInput').addEventListener('change', (e) => {
  uploadRulesFile(e.target.files[0]);
  e.target.value = '';
});
$('#rulesBody').addEventListener('click', (e) => {
  if (e.target.classList.contains('del')) e.target.closest('.rule-row').remove();
  if (e.target.classList.contains('rule-add')) {
    e.target.insertAdjacentHTML('beforebegin', mapRowHtml('', ''));
  }
});

// ---- events ----
$('#drawerClose').addEventListener('click', closeDrawer);
$('#backdrop').addEventListener('click', closeDrawer);
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') { closeDrawer(); closeRules(); } });
$('#themeBtn').addEventListener('click', toggleTheme);
$('#exportBtn').addEventListener('click', exportCsv);
$('#exportXlsBtn').addEventListener('click', exportTableXls);
['#search', '#statusFilter', '#vehicleFilter', '#productionFilter',
  '#upcomingOnly', '#mismatchOnly', '#ptoOnly'].forEach((s) =>
  $(s).addEventListener('input', render));

initTheme();
load().catch((e) => {
  $('#meta').textContent = 'data.json 로드 실패: ' + e.message;
  const msg = $('#statusMsg');
  msg.textContent = 'data.json 을 불러올 수 없습니다: ' + e.message;
  msg.classList.remove('hidden');
});
