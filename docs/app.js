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
  'Category': 'Category (tractor/rigid/tipper)',
};
// Keys whose value should render as a D-day (D-9 / D-Day / D+9).
const DDAY_KEYS = new Set(['Until Dealine']);

// current month "YYYY-MM" (browser today) — used for the "이번달 이후" filter.
const _now = new Date();
const CUR_MONTH = _now.getFullYear() + '-' + String(_now.getMonth() + 1).padStart(2, '0');

let DATA = { rows: [] };
let CODES = { options: {}, mandatory: {} };
let sortKey = null, sortDir = 1;
let restrictSoon = false;  // "2주 이내 전체 Commission" 타일로 목록을 2주 이내로 제한
let tileMandatory = false; // Mandatory 누락 타일: mandatory 코드가 있는 행만
let activeTile = 't-total'; // 현재 활성 대시보드 타일 (하이라이트용)

const $ = (s) => document.querySelector(s);

// ====================== i18n (한국어 / English) ======================
const I18N = {
  ko: {
    'meta.loading': '불러오는 중…',
    'meta.generated': '생성: {when}  ·  WINGS: {file}',
    'btn.rules': '⚙ 규칙',
    'btn.rules.title': '모델 변환 규칙 보기/편집',
    'btn.xls.title': '현재 필터된 결과를 Excel로 저장',
    'btn.csv.title': '현재 필터된 결과를 CSV로 저장',
    'btn.theme.title': '라이트/다크 전환',
    'search.ph': 'Commission no. / 모델 / 코드 검색…',
    'filter.allStatus': '전체 상태',
    'filter.allVehicle': '전체 차종',
    'filter.allProd': '전체 생산월',
    'filter.prod.title': '생산월(Requested delivery) 선택',
    'chk.upcoming': '이번달 이후만',
    'chk.upcoming.title': 'Changeability 가 이번 달 이후인 항목만',
    'chk.pto': 'PTO만',
    'chk.mismatch': '불일치만',
    'count': '{n} / {total} 건',
    'dash.overall': '전체 현황 (이번 생산월 이후)',
    'dash.soon': '2주 이내 (Changeability D-14)',
    'tile.total': '전체 Commission',
    'tile.mismatch': '미스매치',
    'tile.match': '매칭',
    'tile.mand': 'Mandatory 누락',
    'dday.passed': '지남',
    'cell.ok.title': '이상없음 (차이·누락 없음)',
    'msg.noData': '표시할 데이터가 없습니다. data.json 을 먼저 빌드하세요.',
    'msg.noRows': '필터 조건에 맞는 항목이 없습니다.',
    'msg.loadFail': 'data.json 을 불러올 수 없습니다: {err}',
    'meta.loadFail': 'data.json 로드 실패: {err}',
    'code.none': '없음',
    'code.okDiff': '✅ 이상없음 — 누락·차이 없음',
    'code.okMand': '✅ 이상없음 — 아래 필수코드가 양쪽 모두 반영됨',
    'drawer.xls': '⬇ 이 차량 Excel',
    'modal.title': '모델 인식 (model_mapping.xlsx)',
    'modal.note':
      '모델 매칭은 이제 <b>SAM 파일에서 자동</b>으로 이뤄집니다 — 파일 제목 = <b>SAM now(수정)</b>, ' +
      '본문 <code>Vehicle type</code> = <b>SAM Baumuster(원본)</b>. WINGS 번호가 둘 중 하나라도 같으면 매칭됩니다. ' +
      '인식 결과는 <code>model_rules/model_mapping.xlsx</code> 의 <b>인식모델_대조표</b> 시트에서 확인하세요(빌드 시 자동 갱신). ' +
      '<b>잘못 매칭될 때 수정</b>은 엑셀에서: <b>수동매핑</b> 시트(WINGS 모델 → SAM Baumuster/now 강제, 또는 SAM 파일 지정으로 특정 파일과 강제 비교), ' +
      '<b>매칭_별칭(수동)</b> 시트(파일 제목이 다른 세대 번호일 때 연결). 수정 후 <code>python backend/build_data.py</code> 재실행하면 반영·보존됩니다.',
    'modal.reload': '↺ 서버 규칙 다시 불러오기',
    'modal.upload': '⬆ 엑셀 불러오기',
    'modal.upload.title': 'PC의 model_mapping.xlsx 불러오기',
    'modal.download': '⬇ 엑셀 저장 (model_mapping.xlsx)',
    'alert.xlsxJson':
      '엑셀 라이브러리를 불러오지 못해 JSON(rules.json)으로 저장합니다.\n엑셀 편집은 model_rules/model_mapping.xlsx 를 직접 열어 진행하세요.',
    'alert.xlsxBlocked':
      '엑셀 라이브러리를 불러오지 못했습니다(사내망 차단 가능). model_rules/model_mapping.xlsx 를 엑셀에서 직접 편집하세요.',
    'alert.xlsxReadFail': '엑셀 읽기 실패: ',
    'rules.loadFail': 'rules.json 로드 실패: ',
  },
  en: {
    'meta.loading': 'Loading…',
    'meta.generated': 'Generated: {when}  ·  WINGS: {file}',
    'btn.rules': '⚙ Rules',
    'btn.rules.title': 'View / edit model-conversion rules',
    'btn.xls.title': 'Export the current filtered result to Excel',
    'btn.csv.title': 'Export the current filtered result to CSV',
    'btn.theme.title': 'Toggle light / dark',
    'search.ph': 'Search commission no. / model / code…',
    'filter.allStatus': 'All statuses',
    'filter.allVehicle': 'All vehicles',
    'filter.allProd': 'All production months',
    'filter.prod.title': 'Select production month (Requested delivery)',
    'chk.upcoming': 'This month onward',
    'chk.upcoming.title': 'Only rows whose Changeability is this month or later',
    'chk.pto': 'PTO only',
    'chk.mismatch': 'Mismatch only',
    'count': '{n} / {total} rows',
    'dash.overall': 'Overall (this production month onward)',
    'dash.soon': 'Within 2 weeks (Changeability D-14)',
    'tile.total': 'Total commissions',
    'tile.mismatch': 'Mismatch',
    'tile.match': 'Match',
    'tile.mand': 'Mandatory missing',
    'dday.passed': 'Passed',
    'cell.ok.title': 'No issue (no difference / omission)',
    'msg.noData': 'No data to show. Build data.json first.',
    'msg.noRows': 'No items match the current filters.',
    'msg.loadFail': 'Could not load data.json: {err}',
    'meta.loadFail': 'Failed to load data.json: {err}',
    'code.none': 'None',
    'code.okDiff': '✅ No issue — nothing missing or different',
    'code.okMand': '✅ No issue — the mandatory codes below are present on both sides',
    'drawer.xls': '⬇ Export this vehicle',
    'modal.title': 'Model recognition (model_mapping.xlsx)',
    'modal.note':
      'Model matching is now done <b>automatically from the SAM file</b> — file title = <b>SAM now (revised)</b>, ' +
      'body <code>Vehicle type</code> = <b>SAM Baumuster (original)</b>. A match is made when either WINGS number agrees. ' +
      'Check the recognition result in the <b>인식모델_대조표</b> sheet of <code>model_rules/model_mapping.xlsx</code> (auto-refreshed at build time). ' +
      'To <b>fix a wrong match</b>, edit the Excel: the <b>수동매핑</b> sheet (force a WINGS model → SAM Baumuster/now, or pin a specific SAM file for comparison), ' +
      'and the <b>매칭_별칭(수동)</b> sheet (link a file title that uses a different generation number). After editing, re-run <code>python backend/build_data.py</code> to apply and preserve the changes.',
    'modal.reload': '↺ Reload server rules',
    'modal.upload': '⬆ Load Excel',
    'modal.upload.title': 'Load model_mapping.xlsx from your PC',
    'modal.download': '⬇ Save Excel (model_mapping.xlsx)',
    'alert.xlsxJson':
      'The Excel library could not be loaded, so saving as JSON (rules.json).\nTo edit in Excel, open model_rules/model_mapping.xlsx directly.',
    'alert.xlsxBlocked':
      'The Excel library could not be loaded (corporate network may block it). Edit model_rules/model_mapping.xlsx directly in Excel.',
    'alert.xlsxReadFail': 'Failed to read Excel: ',
    'rules.loadFail': 'Failed to load rules.json: ',
  },
};

let LANG = localStorage.getItem('lang') || 'ko';

function t(key, params) {
  let s = (I18N[LANG] && I18N[LANG][key]) || (I18N.ko[key]) || key;
  if (params) for (const [k, v] of Object.entries(params)) s = s.replaceAll('{' + k + '}', v);
  return s;
}

// Apply translations to all static elements carrying data-i18n* attributes.
function applyStaticI18n() {
  document.documentElement.lang = LANG;
  document.querySelectorAll('[data-i18n]').forEach((el) => { el.textContent = t(el.dataset.i18n); });
  document.querySelectorAll('[data-i18n-html]').forEach((el) => { el.innerHTML = t(el.dataset.i18nHtml); });
  document.querySelectorAll('[data-i18n-ph]').forEach((el) => { el.placeholder = t(el.dataset.i18nPh); });
  document.querySelectorAll('[data-i18n-title]').forEach((el) => { el.title = t(el.dataset.i18nTitle); });
  const lb = $('#langBtn');
  if (lb) lb.textContent = LANG === 'ko' ? '🌐 EN' : '🌐 한국어';
}

function toggleLang() {
  LANG = LANG === 'ko' ? 'en' : 'ko';
  localStorage.setItem('lang', LANG);
  applyStaticI18n();
  renderMeta();
  renderSummary();
  fillVehicleFilter();
  fillProductionFilter();
  render();
  if (DRAWER_ROW && !$('#drawer').classList.contains('hidden')) openDrawer(DRAWER_ROW);
}

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
  const locale = LANG === 'ko' ? 'ko-KR' : 'en-GB';
  const when = DATA.generated_at ? new Date(DATA.generated_at).toLocaleString(locale) : '-';
  $('#meta').textContent = t('meta.generated', { when, file: DATA.wings_file || '-' });
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

// "전체 현황" counts only vehicles produced this month or later (current
// production month onward), not the entire WINGS file.
function overallRows() {
  return DATA.rows.filter((r) => {
    const pm = prodMonth(r);
    return pm && pm >= CUR_MONTH;
  });
}

function tile(cls, n, label, action) {
  return `<div class="tile ${cls} clickable" data-tile="${cls}" data-action="${action}">` +
    `<div class="n">${esc(n)}</div><div class="l">${label}</div></div>`;
}

// Each tile applies a filter/sort combination. soon = restrict to Changeability
// D-14; status = SAM Status filter; mand = only rows with mandatory codes missing;
// sort = [column key, direction] applied to the grid (null = clear sort).
const TILE_ACTIONS = {
  't-total':  { soon: false, status: '',         mand: false, sort: null },
  't-miss':   { soon: false, status: 'Mismatch', mand: false, sort: ['Until Dealine', 1] },
  't-match':  { soon: false, status: 'Match',    mand: false, sort: null },
  't-mand':   { soon: false, status: '',         mand: true,  sort: ['Mandatory Codes', -1] },
  't-total2': { soon: true,  status: '',         mand: false, sort: ['Until Dealine', 1] },
  't-miss2':  { soon: true,  status: 'Mismatch', mand: false, sort: ['Until Dealine', 1] },
  't-match2': { soon: true,  status: 'Match',    mand: false, sort: null },
  't-mand2':  { soon: true,  status: '',         mand: true,  sort: ['Mandatory Codes', -1] },
};

function applyTile(id) {
  const cfg = TILE_ACTIONS[id];
  if (!cfg) return;
  activeTile = id;
  restrictSoon = cfg.soon;
  tileMandatory = cfg.mand;
  $('#statusFilter').value = cfg.status;
  $('#mismatchOnly').checked = false;
  if (cfg.sort) { sortKey = cfg.sort[0]; sortDir = cfg.sort[1]; }
  else { sortKey = null; sortDir = 1; }
  renderHead();
  syncTileActive();
  render();
}

function renderSummary() {
  const all = dashStats(overallRows());
  const soon = dashStats(DATA.rows.filter(within2weeks));
  $('#summary').innerHTML = `
    <div class="dash-row">
      <div class="dash-cap">${t('dash.overall')}</div>
      <div class="tiles">
        ${tile('t-total', all.total, t('tile.total'), 't-total')}
        ${tile('t-miss', all.mismatch, t('tile.mismatch'), 't-miss')}
        ${tile('t-match', all.match, t('tile.match'), 't-match')}
        ${tile('t-mand', all.mand, t('tile.mand'), 't-mand')}
      </div>
    </div>
    <div class="dash-row">
      <div class="dash-cap">${t('dash.soon')}</div>
      <div class="tiles">
        ${tile('t-total2', soon.total, t('tile.total'), 't-total2')}
        ${tile('t-miss2', soon.mismatch, t('tile.mismatch'), 't-miss2')}
        ${tile('t-match2', soon.match, t('tile.match'), 't-match2')}
        ${tile('t-mand2', soon.mand, t('tile.mand'), 't-mand2')}
      </div>
    </div>`;
  $('#summary').querySelectorAll('.tile[data-action]').forEach((el) =>
    el.addEventListener('click', () => applyTile(el.dataset.action)));
  syncTileActive();
}

function syncTileActive() {
  $('#summary').querySelectorAll('.tile').forEach((el) =>
    el.classList.toggle('active', el.dataset.tile === activeTile));
}

function fmtMonth(ym) {
  const s = String(ym);
  return s.length === 6 ? `${s.slice(0, 4)}-${s.slice(4)}` : s;
}

function fillVehicleFilter() {
  const sel = $('#vehicleFilter');
  const prev = sel.value;
  const vehicles = [...new Set(DATA.rows.map((r) => r.Vehicle).filter(Boolean))].sort();
  sel.innerHTML = `<option value="">${esc(t('filter.allVehicle'))}</option>` +
    vehicles.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  sel.value = prev;
}

function fillProductionFilter() {
  const sel = $('#productionFilter');
  const prev = sel.value;
  const months = [...new Set(DATA.rows.map(prodMonth).filter(Boolean))].sort().reverse();
  sel.innerHTML = `<option value="">${esc(t('filter.allProd'))}</option>` +
    months.map((m) => `<option value="${esc(m)}">${esc(m)}</option>`).join('');
  sel.value = prev;
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
  if (s.toLowerCase() === 'passed') return t('dday.passed');
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
  if (s.toLowerCase() === 'passed') return `<span class="dday passed">${esc(t('dday.passed'))}</span>`;
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
  if (!n) return `<span class="cbadge ok" title="${esc(t('cell.ok.title'))}">✓</span>`;
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
    if (restrictSoon && !within2weeks(r)) return false;
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
    if (tileMandatory && countOf(r['Mandatory Codes']) === 0) return false;
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
  $('#count').textContent = t('count', { n: rows.length, total: DATA.rows.length });
  const tb = $('#grid tbody');
  const msg = $('#statusMsg');

  if (!DATA.rows.length) {
    tb.innerHTML = '';
    msg.textContent = t('msg.noData');
    msg.classList.remove('hidden');
    return;
  }
  if (!rows.length) {
    tb.innerHTML = '';
    msg.textContent = t('msg.noRows');
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
      body = `<div class="ok-mark">${opts.okLabel || t('code.okDiff')}</div>`;
      if (ok.length) body += `<div class="code-list ok-list">${codeRowsHtml(ok, 'ok-row')}</div>`;
      else if (opts.okHint) body += `<div class="hint">${esc(opts.okHint)}</div>`;
    } else {
      badge = '<span class="badge">0</span>';
      body = `<div class="none">${esc(t('code.none'))}</div>`;
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
  const meta = ['Vehicle', 'Category', 'Type', 'Cab', 'PTO', 'Production date', 'Changeability Date',
    'Until Dealine', 'SAM Baumuster', 'SAM now',
    'Order status financial', 'SAM Status', 'FIN']
    .filter((k) => r[k] !== undefined && r[k] !== '')
    .map((k) => {
      let val;
      if (DDAY_KEYS.has(k)) val = ddayHtml(r[k]);                    // colored D-Day (same as grid)
      else if (k === 'SAM Status') {
        const s = String(r[k]);
        val = `<span class="status ${esc(s).replace(/\s+/g, '')}">${esc(s)}</span>`;  // colored badge
      } else val = esc(r[k]);
      return `<div class="kv-item"><div class="k">${esc(META_LABELS[k] || k)}</div><div class="v">${val}</div></div>`;
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
      <button id="drawerXls" class="icon-btn primary">${esc(t('drawer.xls'))}</button>
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
        ${codeColHtml('Mandatory', r['Mandatory Codes'], { diff: true, okCodes: matchedMand, okLabel: t('code.okMand') })}
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

// ---- Minimal, dependency-free .xlsx writer (real OOXML, with cell shading) ----
// The bundled SheetJS community build cannot write fills, so we emit the package
// ourselves: a STORED (uncompressed) zip of the OOXML parts. Works even when the
// CDN is blocked. Style indices used below:
//   1 title(bold) · 2 header(bold+grey+border) · 3 cell(border) ·
//   4 cell shaded/diff(border) · 5 label(bold+border)
function _colName(n) {
  let s = ''; n++;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function _crc32(bytes) {
  let c = ~0;
  for (let i = 0; i < bytes.length; i++) {
    c ^= bytes[i];
    for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xEDB88320 & -(c & 1));
  }
  return (~c) >>> 0;
}
function _zipStore(files) {
  const enc = new TextEncoder();
  const parts = [], central = [];
  let offset = 0;
  const u16 = (n) => [n & 255, (n >> 8) & 255];
  const u32 = (n) => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >> 24) & 255];
  for (const f of files) {
    const nameB = enc.encode(f.name);
    const data = f.data;
    const crc = _crc32(data);
    const local = new Uint8Array([].concat(
      u32(0x04034b50), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(crc), u32(data.length), u32(data.length), u16(nameB.length), u16(0)));
    parts.push(local, nameB, data);
    central.push({ head: new Uint8Array([].concat(
      u32(0x02014b50), u16(20), u16(20), u16(0), u16(0), u16(0), u16(0),
      u32(crc), u32(data.length), u32(data.length),
      u16(nameB.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset))), name: nameB });
    offset += local.length + nameB.length + data.length;
  }
  const cdStart = offset;
  for (const c of central) { parts.push(c.head, c.name); offset += c.head.length + c.name.length; }
  const eocd = new Uint8Array([].concat(
    u32(0x06054b50), u16(0), u16(0), u16(central.length), u16(central.length),
    u32(offset - cdStart), u32(cdStart), u16(0)));
  parts.push(eocd);
  return new Blob(parts, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
const _XLSX_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
  '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="4"><fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FFF0F2F5"/></patternFill></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFCE4E6"/></patternFill></fill></fills>' +
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
  '<border><left style="thin"><color rgb="FFD0D0D0"/></left><right style="thin"><color rgb="FFD0D0D0"/></right>' +
  '<top style="thin"><color rgb="FFD0D0D0"/></top><bottom style="thin"><color rgb="FFD0D0D0"/></bottom><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="6">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
  '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +
  '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>' +
  '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

function _cellXml(ref, cell) {
  const s = cell.s || 0;
  const v = cell.v;
  if (v === '' || v == null) return `<c r="${ref}" s="${s}"/>`;
  return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(String(v))}</t></is></c>`;
}
function _sheetXml(rows, merges, cols) {
  const colsXml = (cols && cols.length)
    ? '<cols>' + cols.map((c) => `<col min="${c.min}" max="${c.max}" width="${c.w}" customWidth="1"/>`).join('') + '</cols>'
    : '';
  const body = rows.map((row, ri) => {
    const cells = row.map((cell, ci) => cell == null ? '' : _cellXml(_colName(ci) + (ri + 1), cell)).join('');
    return `<row r="${ri + 1}">${cells}</row>`;
  }).join('');
  const mg = (merges && merges.length)
    ? `<mergeCells count="${merges.length}">` + merges.map((m) => `<mergeCell ref="${m}"/>`).join('') + '</mergeCells>'
    : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    colsXml + `<sheetData>${body}</sheetData>${mg}</worksheet>`;
}
function writeXlsx(filename, sheetName, rows, merges, cols) {
  const enc = new TextEncoder();
  const file = (name, str) => ({ name, data: enc.encode(str) });
  const files = [
    file('[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'),
    file('_rels/.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
    file('xl/workbook.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets><sheet name="${esc(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    file('xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),
    file('xl/styles.xml', _XLSX_STYLES),
    file('xl/worksheets/sheet1.xml', _sheetXml(rows, merges, cols)),
  ];
  const blob = _zipStore(files);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

// Export one vehicle's comparison to .xlsx: vehicle info spread across columns,
// then a SAM ↔ WINGS code table aligned by code. A code present on only one side
// leaves the other two cells blank and shaded (style 4) to flag the difference.
function exportRowXls(r) {
  if (!r) return;
  const metaKeys = ['Commission no.', 'Model(WINGS)', 'Vehicle', 'Category', 'Type', 'Cab', 'PTO',
    'Production date', 'Changeability Date', 'Until Dealine',
    'SAM Baumuster', 'SAM now', 'SAM Status', 'Compared SAM file name'];
  const rows = [];
  const merges = [];

  rows.push([{ v: `${r['Commission no.']}  ·  ${r['Model(WINGS)'] || ''}`, s: 1 }]);
  rows.push([{ v: r['Compared SAM file name'] || '', s: 0 }]);
  rows.push([]);

  // Vehicle info: two label→value pairs per row (uses the full width).
  const infoKeys = metaKeys.filter((k) => r[k] !== undefined && r[k] !== '');
  for (let i = 0; i < infoKeys.length; i += 2) {
    const cells = [];
    for (let j = 0; j < 2; j++) {
      const k = infoKeys[i + j];
      if (!k) { cells.push(null, null); continue; }
      const val = DDAY_KEYS.has(k) ? ddayText(r[k]) : r[k];
      cells.push({ v: META_LABELS[k] || k, s: 5 }, { v: String(val), s: 3 });
    }
    rows.push(cells);
  }
  rows.push([]);

  // Code comparison, aligned by code (union sorted).
  const sam = new Set(splitCodes(r['_all_sam_codes']));
  const wings = new Set(splitCodes(r['_all_wings_codes']));
  const titleRow = rows.length + 1;
  rows.push([{ v: `Code Comparison — SAM (${sam.size}) ↔ WINGS (${wings.size});  shaded = missing on that side`, s: 1 }]);
  merges.push(`A${titleRow}:D${titleRow}`);
  rows.push([{ v: 'SAM Code', s: 2 }, { v: 'SAM Description', s: 2 },
    { v: 'WINGS Code', s: 2 }, { v: 'WINGS Description', s: 2 }]);
  const union = [...new Set([...sam, ...wings])].sort();
  for (const code of union) {
    const inS = sam.has(code), inW = wings.has(code);
    rows.push([
      inS ? { v: code, s: 3 } : { v: '', s: 4 },
      inS ? { v: describe(code) || '', s: 3 } : { v: '', s: 4 },
      inW ? { v: code, s: 3 } : { v: '', s: 4 },
      inW ? { v: describe(code) || '', s: 3 } : { v: '', s: 4 },
    ]);
  }

  const cols = [{ min: 1, max: 1, w: 16 }, { min: 2, max: 2, w: 52 },
    { min: 3, max: 3, w: 16 }, { min: 4, max: 4, w: 52 }];
  const name = String(r['Commission no.'] || 'vehicle').replace(/[^\w.-]/g, '_');
  writeXlsx(`afab_sam_${name}.xlsx`, 'Comparison', rows, merges, cols);
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
    $('#rulesBody').innerHTML = `<div class="none">${esc(t('rules.loadFail') + e.message)}</div>`;
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
  alert(t('alert.xlsxJson'));
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
    alert(t('alert.xlsxBlocked'));
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      renderRules(sheetsToRules(wb));
    } catch (err) {
      alert(t('alert.xlsxReadFail') + err.message);
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
$('#langBtn').addEventListener('click', toggleLang);
$('#exportBtn').addEventListener('click', exportCsv);
$('#exportXlsBtn').addEventListener('click', exportTableXls);
// Manual edits to the filter controls become the single source of truth: drop the
// tile-only hidden flags (2-week restriction, mandatory-only) and clear the tile
// highlight so what the user sees in the controls is exactly what is applied.
function onManualFilter() {
  restrictSoon = false;
  tileMandatory = false;
  activeTile = null;
  syncTileActive();
  render();
}
['#search', '#statusFilter', '#vehicleFilter', '#productionFilter',
  '#upcomingOnly', '#mismatchOnly', '#ptoOnly'].forEach((s) =>
  $(s).addEventListener('input', onManualFilter));

initTheme();
applyStaticI18n();
load().catch((e) => {
  $('#meta').textContent = t('meta.loadFail', { err: e.message });
  const msg = $('#statusMsg');
  msg.textContent = t('msg.loadFail', { err: e.message });
  msg.classList.remove('hidden');
});
