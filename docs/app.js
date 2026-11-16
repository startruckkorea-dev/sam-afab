// AFAB x SAM comparison viewer — reads data.json + codes.json from SharePoint
// ('05. output'), falling back to the copy deployed with the site. 관리자가
// "데이터 빌드"를 누르면 lib/pipeline.js 가 이 브라우저에서 계산해 그 결과를
// SharePoint 에 저장한다. 모델 매칭 / 코드 관리 뷰는 SharePoint Excel 을
// Graph 로 직접 읽고 쓴다 (graph.js + auth.js).

const COLS = [
  { key: 'Commission no.', label: 'Commission' },
  { key: 'Vehicle', label: 'Model' },
  { key: 'Model(WINGS)', label: 'Type' },
  { key: 'Type', label: 'Axle' },
  { key: 'Cab', label: 'Cab', pair: 'cab' },
  { key: 'MY', label: 'MY' },
  { key: 'PTO', label: 'PTO' },
  { key: 'Production date', label: 'Production' },
  { key: 'Changeability Date', label: 'Changeability' },
  { key: 'Until Dealine', label: 'Changeability D-Day', dday: true },
  { key: 'Only_in_SAM', label: 'Only in SAM', count: 'sam' },
  { key: 'Only_in_WINGS', label: 'Only in WINGS', count: 'win' },
  { key: 'Mandatory Codes', label: 'Mandatory', count: 'mand' },
  { key: 'SAM Status', label: 'Status', status: true },
];

const NUMERIC_KEYS = new Set(['Until Dealine', 'Baumuster',
  'Only_in_SAM', 'Only_in_WINGS', 'Mandatory Codes']);

const META_LABELS = {
  'Vehicle': 'Model',
  'Model(WINGS)': 'Type',
  'Type': 'Axle',
  'MY': 'MY',
  'Changeability Date': 'Changeability',
  'Until Dealine': 'Changeability D-Day',
  'Category': 'Category (tractor/rigid/tipper)',
};
const DDAY_KEYS = new Set(['Until Dealine']);

// Model Year (MY): recognized from SAM codes V8Q..V8Z = "Model year 0..9".
// The digit is the last digit of the model year (e.g. V8W → 6 → 2026).
const MY_CODE_DIGIT = {
  V8Q: 0, V8R: 1, V8S: 2, V8T: 3, V8U: 4,
  V8V: 5, V8W: 6, V8X: 7, V8Y: 8, V8Z: 9,
};
function computeMY(r) {
  // SAM 쪽 코드가 1순위. SAM 이 없는(No SAM) 차량도 MY 로 묶을 수 있게 WINGS 코드로 보완한다.
  for (const key of ['_all_sam_codes', '_all_wings_codes']) {
    for (const c of splitCodes(r && r[key])) {
      if (c in MY_CODE_DIGIT) return String(2020 + MY_CODE_DIGIT[c]);
    }
  }
  return '';
}

const _now = new Date();
const CUR_MONTH = _now.getFullYear() + '-' + String(_now.getMonth() + 1).padStart(2, '0');
const CUR_DATE = CUR_MONTH + '-' + String(_now.getDate()).padStart(2, '0');
// 올해 1월 — 데이터에 그 달 차가 없어도 시작월 후보로는 늘 띄운다.
const YEAR_START = CUR_MONTH.slice(0, 4) + '-01';

// 대시보드 타일과 표가 다루는 생산월의 하한(YYYY-MM). 기본은 이번 달이라
// 지난 차량은 보이지 않는다. 시작월 드롭다운으로 과거까지 넓히면
// 그때부터 지난 차량도 함께 비교된다.
let DASH_FROM = CUR_MONTH;

let DATA = { rows: [] };
let CODES = { options: {}, mandatory: {} };
let sortKey = null, sortDir = 1;
let restrictSoon = false;
let tileMandatory = false;
let tileSamUpdate = false;
let activeTile = 't-total';
// '다음 Changeability 마감' 카드의 '목록 보기' — 이 날짜(YYYY-MM-DD)의 Changeability 만 남긴다.
let tileDate = null;

const $ = (s) => document.querySelector(s);

// ====================== i18n (한국어 / English) ======================
// 문구 사전은 i18n/ko.json · en.json 에 있다. 첫 렌더 전에 boot() 가 둘 다 읽어 온다.
// 기본은 한국어, 키가 없으면 한국어로, 그래도 없으면 키 자체로 떨어진다.
const I18N = { ko: {}, en: {} };
const I18N_VER = '20261116i';
async function loadI18n() {
  await Promise.all(['ko', 'en'].map(async (l) => {
    try {
      const r = await fetch('i18n/' + l + '.json?v=' + I18N_VER);
      if (r.ok) I18N[l] = await r.json();
    } catch (e) { console.warn('[i18n] ' + l + '.json 로드 실패:', e.message); }
  }));
}

// 선택 언어는 이 브라우저에 기억한다. 저장소가 막힌 환경(사생활 보호 창 등)에서는 기본값.
function readLang() {
  try { const v = localStorage.getItem('lang'); return v === 'en' ? 'en' : 'ko'; } catch { return 'ko'; }
}
let LANG = readLang();

// ====================== 접속 권한 (auth.js 가 SharePoint 명단에서 확정) ======================
// 로컬 개발처럼 로그인 게이트가 없는 환경에서는 auth.js 가 'admin' 을 넣어준다.
const ROLE = (window.MB_AUTH && window.MB_AUTH.role) || 'admin';
const IS_ADMIN = ROLE === 'admin';
const LOCKED_VIEWS = ['matching', 'codes'];

// 영어 단수형 — n 이 1 이고 '<key>.one' 문구가 있으면 그걸 쓴다(한국어는 따로 두지 않는다).
function tn(key, n, params) {
  const one = key + '.one';
  return t(n === 1 && I18N[LANG] && I18N[LANG][one] ? one : key, params);
}

function t(key, params) {
  let s = (I18N[LANG] && I18N[LANG][key]) || (I18N.ko[key]) || key;
  if (params) for (const [k, v] of Object.entries(params)) s = s.replaceAll('{' + k + '}', v);
  return s;
}

function applyStaticI18n() {
  document.documentElement.lang = LANG;
  document.querySelectorAll('[data-i18n]').forEach((el) => { el.textContent = t(el.dataset.i18n); });
  document.querySelectorAll('[data-i18n-html]').forEach((el) => { el.innerHTML = t(el.dataset.i18nHtml); });
  document.querySelectorAll('[data-i18n-ph]').forEach((el) => { el.placeholder = t(el.dataset.i18nPh); });
  document.querySelectorAll('[data-i18n-title]').forEach((el) => { el.title = t(el.dataset.i18nTitle); });
  document.querySelectorAll('[data-i18n-aria]').forEach((el) => { el.setAttribute('aria-label', t(el.dataset.i18nAria)); });
  document.querySelectorAll('.lang-toggle button[data-lang]').forEach((b) => {
    const on = b.dataset.lang === LANG;
    b.classList.toggle('active', on);
    b.setAttribute('aria-pressed', on ? 'true' : 'false');
  });
  document.body.classList.toggle('lang-en', LANG === 'en');
  applyRoleUi();
}

// Read 권한이면 편집성 기능(모델 매칭 · 코드 관리 · 데이터 다시 계산)을 비활성 표시한다.
// 탭은 숨기지 않고 회색으로 두어 기능의 존재는 보이게 한다. 실제 차단은
// switchView() / runBuild() 의 가드가 담당한다.
function applyRoleUi() {
  if (IS_ADMIN) return;
  LOCKED_VIEWS.forEach((v) => {
    const el = document.querySelector('.nav-link[data-view="' + v + '"]');
    if (!el) return;
    el.classList.add('disabled');
    el.setAttribute('aria-disabled', 'true');
    el.title = t('role.noPerm');
  });
  const bb = $('#buildBtn');
  if (bb) { bb.disabled = true; bb.title = t('role.noPerm'); }
}

function setLang(lang) {
  if (lang === LANG) return;
  LANG = lang === 'en' ? 'en' : 'ko';
  try { localStorage.setItem('lang', LANG); } catch { /* 저장 불가 — 이번 세션만 */ }
  applyStaticI18n();
  renderMeta();
  renderSummary();
  fillFilters();
  render();
  if (VIEW_INIT.history) { histFillMy(); histRender(); }
  if (VIEW_INIT.samcmp) scmpRefresh();
  if (VIEW_INIT.exceptions) excRender();
  if (typeof codeEditor !== 'undefined') codeEditor.refresh();
  if (typeof matchEditor !== 'undefined') matchEditor.refresh();
  if (VIEW_INIT.matching) renderMatchSummary();
  if (DRAWER_ROW && !$('#drawer').classList.contains('hidden')) openDrawer(DRAWER_ROW);
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ====================== 뷰 전환 (대시보드 / 생산월 이력관리 / 모델 매칭 / 코드 관리) ======================
let CUR_VIEW = 'dashboard';
const VIEW_INIT = { history: false, samcmp: false, exceptions: false, matching: false, codes: false, manual: false };

function switchView(view) {
  if (!['dashboard', 'history', 'samcmp', 'exceptions', 'matching', 'codes', 'manual'].includes(view)) return;
  if (!IS_ADMIN && LOCKED_VIEWS.includes(view)) { alert(t('role.denied')); return; }
  // 저장 안 한 편집이 있으면 이탈 확인 (모델 매칭 / 코드 관리)
  if (view !== CUR_VIEW) {
    const leavingEd = CUR_VIEW === 'codes' ? codeEditor : (CUR_VIEW === 'matching' ? matchEditor : null);
    if (leavingEd && leavingEd.S.dirty) {
      if (!confirm(t('codes.confirmLeave'))) return;
      leavingEd.S.dirty = false;
    }
  }
  CUR_VIEW = view;
  document.querySelectorAll('.view').forEach((el) =>
    el.classList.toggle('active', el.id === 'view-' + view));
  document.querySelectorAll('.nav-link').forEach((el) => {
    el.classList.toggle('active', el.dataset.view === view);
    if (el.dataset.view === view) el.setAttribute('aria-current', 'page');
    else el.removeAttribute('aria-current');
  });
  // Export 는 대시보드 필터에 대한 기능이라 다른 뷰에서는 숨긴다.
  const exportBtn = $('#exportBtn');
  if (exportBtn) exportBtn.classList.toggle('hidden', view !== 'dashboard');
  if (view === 'history' && !VIEW_INIT.history) { VIEW_INIT.history = true; initHistory(); }
  if (view === 'samcmp' && !VIEW_INIT.samcmp) { VIEW_INIT.samcmp = true; initSamCmp(); }
  if (view === 'exceptions' && !VIEW_INIT.exceptions) { VIEW_INIT.exceptions = true; initExceptions(); }
  if (view === 'matching' && !VIEW_INIT.matching) { VIEW_INIT.matching = true; initMatching(); }
  if (view === 'codes' && !VIEW_INIT.codes) { VIEW_INIT.codes = true; initCodes(); }
  if (view === 'manual' && !VIEW_INIT.manual) { VIEW_INIT.manual = true; initManual(); }
}

// ====================== 사용자 매뉴얼 ======================
// 본문은 index.html 에 그대로 들어 있다(정적 문서). 여기서는 왼쪽 목차 이동과
// 스크롤에 따른 현재 위치 표시만 담당한다.
function initManual() {
  const body = $('#manBody');
  const toc = $('#manToc');
  if (!body || !toc) return;
  const links = [...toc.querySelectorAll('.man-tl')];
  const secOf = (l) => document.getElementById(l.dataset.sec);

  links.forEach((l) => l.addEventListener('click', () => {
    const sec = secOf(l);
    if (sec) body.scrollTo({ top: sec.offsetTop - 12, behavior: 'smooth' });
  }));

  // 현재 보고 있는 절 — 화면 위쪽(1/3 지점)을 지난 마지막 섹션.
  function syncActive() {
    const line = body.scrollTop + body.clientHeight / 3;
    let cur = links[0];
    for (const l of links) {
      const sec = secOf(l);
      if (sec && sec.offsetTop <= line) cur = l;
    }
    links.forEach((l) => l.classList.toggle('sel', l === cur));
    const sel = toc.querySelector('.man-tl.sel');
    if (sel) sel.scrollIntoView({ block: 'nearest' });
  }
  let ticking = false;
  body.addEventListener('scroll', () => {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(() => { ticking = false; syncActive(); });
  });
  syncActive();
}

// op-status 배지 (info/ok/err)
function setStatus(el, msg, type) {
  if (typeof el === 'string') el = $(el);
  if (!el) return;
  el.textContent = msg || '';
  el.className = 'op-status' + (msg ? ' show ' + (type || 'info') : '');
}

// ====================== 대시보드 ======================
// 데이터 출처: SharePoint '05. output' 의 빌드 결과가 1순위, 저장소에 함께 배포된
// 사본이 2순위. 관리자가 "데이터 빌드"를 누르면 브라우저에서 계산해 SharePoint 에
// 저장하므로, 다른 사용자는 재계산 없이 그 결과만 내려받는다.
let DATA_SOURCE = '';

async function fetchJson(name, fallback) {
  if (window.Graph && Graph.available()) {
    try {
      const j = await Graph.downloadJson('output', name);
      DATA_SOURCE = 'sharepoint';
      return j;
    } catch (e) {
      console.warn('[data] SharePoint ' + name + ' 없음/실패 — 사본 사용:', e.message);
    }
  }
  try {
    const j = await fetch(name + '?_=' + Date.now()).then((r) => {
      if (!r.ok) throw new Error(r.status);
      return r.json();
    });
    if (!DATA_SOURCE) DATA_SOURCE = 'bundled';
    return j;
  } catch (e) {
    if (fallback !== undefined) return fallback;
    throw e;
  }
}

async function load() {
  DATA_SOURCE = '';
  const d = await fetchJson('data.json');
  const c = await fetchJson('codes.json', CODES);
  await Promise.all([loadExceptions(), loadMatchIgnore()]);   // 실패해도 대시보드는 뜬다(미적용 = 원래 판정 그대로)
  applyData(d, c);
}

// 새로 만든(또는 내려받은) 데이터로 대시보드를 다시 그린다.
function applyData(d, c) {
  DATA = d; CODES = c || CODES;
  (DATA.rows || []).forEach((r) => { r.MY = computeMY(r); });
  applyExceptions();
  renderMeta();
  renderSummary();
  renderModelPanel();
  fillFilters();
  renderHead();
  render();
  if (VIEW_INIT.history) { histFillMy(); histRender(); }
  if (VIEW_INIT.samcmp) scmpRefresh();
  if (VIEW_INIT.exceptions) excRender();
  if (VIEW_INIT.matching) {
    renderMatchSummary();
    renderMatchProblems();
    matchEditor.refresh();
  }
}

function prodMonth(r) { return String(r['Production date'] || '').slice(0, 7); }
function changeMonth(r) {
  const s = String(r['Changeability Date'] || '');
  return /^\d{4}-\d{2}/.test(s) ? s.slice(0, 7) : '';
}

// 시각은 언어와 상관없이 ISO 형식 KST(UTC+9)로 보여 준다 — 'YYYY-MM-DD HH:mm'.
function fmtKst(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  return new Date(d.getTime() + 9 * 3600e3).toISOString().slice(0, 16).replace('T', ' ');
}
// WINGS 리포트 날짜 = 파일명 앞의 'YYYY-MM-DD'. 그런 접두어가 없으면 파일명 그대로.
function wingsReportDate(file) {
  const m = /^(\d{4}-\d{2}-\d{2})/.exec(String(file || '').trim());
  return m ? m[1] : String(file || '-');
}
function fmtNum(n) { return Number(n).toLocaleString('en-US'); }

function renderMeta() {
  const when = DATA.generated_at ? fmtKst(DATA.generated_at) : '';
  const src = DATA_SOURCE === 'sharepoint' ? t('meta.srcSp')
    : (DATA_SOURCE === 'bundled' ? t('meta.srcBundled') : '');
  const parts = [];
  if (when) parts.push(t('meta.asOf', { when }));
  if (DATA.wings_file) parts.push(t('meta.wings', { date: wingsReportDate(DATA.wings_file) }));
  if (src) parts.push(src);
  const meta = $('#meta');
  meta.textContent = parts.join('  ·  ') || '-';
  meta.title = DATA.wings_file ? 'WINGS: ' + DATA.wings_file : '';

  // 최신성 — 마지막 계산(generated_at) 이후 24시간 미만 초록, 7일까지 주황, 그 이상 빨강.
  const pill = $('#freshPill');
  const ts = DATA.generated_at ? new Date(DATA.generated_at).getTime() : NaN;
  if (!pill) return;
  if (Number.isNaN(ts)) { pill.className = 'pill hidden'; return; }
  const hours = Math.max(0, (Date.now() - ts) / 3600e3);
  const days = Math.floor(hours / 24);
  pill.className = 'pill ' + (hours < 24 ? 'ok' : (days <= 7 ? 'warn' : 'crit'));
  pill.textContent = hours < 24 ? t('fresh.recent') : t('fresh.days', { n: days });
  pill.title = t('fresh.title');
}

function dashStats(rows) {
  return {
    total: rows.length,
    mismatch: rows.filter((r) => r['SAM Status'] === 'Mismatch').length,
    match: rows.filter((r) => r['SAM Status'] === 'Match').length,
    nosam: rows.filter((r) => r['SAM Status'] === 'No SAM').length,
    samupd: rows.filter((r) => r['SAM Update']).length,   // 대체 SAM 매칭(총계와 무관, 겹침)
    mand: rows.filter((r) => countOf(r['Mandatory Codes']) > 0).length,
  };
}

function within2weeks(r) {
  const n = Number(r['Until Dealine']);
  return !Number.isNaN(n) && n >= 0 && n <= 14;
}

function overallRows() {
  return DATA.rows.filter((r) => { const pm = prodMonth(r); return pm && pm >= DASH_FROM; });
}

function tile(cls, n, label, action, row) {
  return `<button type="button" class="tile ${cls} clickable" data-tile="${cls}" data-action="${action}"` +
    ` aria-pressed="false" aria-label="${esc(t('tile.aria', { row, label, n: fmtNum(n) }))}">` +
    `<span class="n">${esc(fmtNum(n))}</span><span class="l">${esc(label)}</span></button>`;
}

const TILE_ACTIONS = {
  't-total':  { soon: false, status: '',         mand: false, samupd: false, sort: null },
  't-miss':   { soon: false, status: 'Mismatch', mand: false, samupd: false, sort: ['Until Dealine', 1] },
  't-match':  { soon: false, status: 'Match',    mand: false, samupd: false, sort: null },
  't-nosam':  { soon: false, status: 'No SAM',   mand: false, samupd: false, sort: null },
  't-samupd': { soon: false, status: '',         mand: false, samupd: true,  sort: null },
  't-mand':   { soon: false, status: '',         mand: true,  samupd: false, sort: ['Mandatory Codes', -1] },
  't-total2': { soon: true,  status: '',         mand: false, samupd: false, sort: ['Until Dealine', 1] },
  't-miss2':  { soon: true,  status: 'Mismatch', mand: false, samupd: false, sort: ['Until Dealine', 1] },
  't-match2': { soon: true,  status: 'Match',    mand: false, samupd: false, sort: null },
  't-nosam2': { soon: true,  status: 'No SAM',   mand: false, samupd: false, sort: null },
  't-samupd2':{ soon: true,  status: '',         mand: false, samupd: true,  sort: null },
  't-mand2':  { soon: true,  status: '',         mand: true,  samupd: false, sort: ['Mandatory Codes', -1] },
};

function applyTile(id) {
  const cfg = TILE_ACTIONS[id];
  if (!cfg) return;
  activeTile = id;
  tileDate = null;
  restrictSoon = cfg.soon;
  tileMandatory = cfg.mand;
  tileSamUpdate = !!cfg.samupd;
  // 타일 숫자는 '이번 생산월 이후' 전체를 센 값이므로, 눌렀을 때 기존 드롭다운·검색
  // 필터를 모두 풀고 타일 조건만 남긴다(그래야 표 건수 = 타일 숫자).
  for (const f of FILTER_FIELDS) $(f.id).value = '';
  $('#search').value = '';
  $('#upcomingOnly').checked = true;
  $('#statusFilter').value = cfg.status;
  fillFilters();                       // 연동 옵션 재계산 + 커스텀 드롭다운 라벨 갱신
  if (cfg.sort) { sortKey = cfg.sort[0]; sortDir = cfg.sort[1]; }
  else { sortKey = null; sortDir = 1; }
  renderHead();
  syncTileActive();
  render();
}

function renderSummary() {
  const all = dashStats(overallRows());
  const soon = dashStats(DATA.rows.filter(within2weeks));
  // 줄마다 같은 6칸 — 접미사 '2' 가 '2주 이내' 줄의 타일(TILE_ACTIONS 키와 같다).
  const row = (title, sub, s, sfx) => `
    <div class="dash-row">
      <div class="row-head"><span class="row-title">${esc(title)}</span><span class="row-sub">${esc(sub)}</span></div>
      <div class="tiles">
        ${tile('t-total' + sfx, s.total, t('tile.total'), 't-total' + sfx, title)}
        ${tile('t-match' + sfx, s.match, t('tile.match'), 't-match' + sfx, title)}
        ${tile('t-miss' + sfx, s.mismatch, t('tile.mismatch'), 't-miss' + sfx, title)}
        ${tile('t-nosam' + sfx, s.nosam, t('tile.nosam'), 't-nosam' + sfx, title)}
        ${tile('t-mand' + sfx, s.mand, t('tile.mand'), 't-mand' + sfx, title)}
        ${tile('t-samupd' + sfx, s.samupd, t('tile.samupd'), 't-samupd' + sfx, title)}
      </div>
    </div>`;
  const el = $('#summary');
  el.innerHTML = nearestCardHtml()
    + row(t('dash.overall'), t('dash.overall.sub', { from: DASH_FROM }), all, '')
    + row(t('dash.soon'), t('dash.soon.sub'), soon, '2');
  el.removeAttribute('aria-busy');
  // 같은 타일을 다시 누르면 해제(= 기본 상태 '전체 Commission').
  el.querySelectorAll('.tile[data-action]').forEach((b) =>
    b.addEventListener('click', () => {
      const id = b.dataset.action;
      applyTile(activeTile === id && id !== 't-total' ? 't-total' : id);
    }));
  const nb = $('#nearList');
  if (nb) nb.addEventListener('click', () =>
    (tileDate === nb.dataset.date ? applyTile('t-total') : applyNearDate(nb.dataset.date)));
  syncTileActive();
}

// 다음 Changeability 마감 카드 — 날짜 · D-Day · 해당일 건수 + '목록 보기'.
function nearestCardHtml() {
  const near = nearestChangeability(overallRows());
  const cap = `<svg class="ic near-ic" aria-hidden="true"><use href="#i-calendar"/></svg>`
    + `<span class="near-cap">${esc(t('dash.near'))}</span>`;
  if (!near) return `<div class="near-card">${cap}<span class="near-none">${esc(t('dash.near.none'))}</span></div>`;
  return `<div class="near-card">${cap}
      <span class="near-sep" aria-hidden="true">·</span>
      <span class="near-date">${esc(near.date)}</span>
      ${ddayHtml(near.dday)}
      <span class="near-sep" aria-hidden="true">·</span>
      <span class="near-count"><b>${esc(t('dash.near.n', { n: fmtNum(near.count) }))}</b> · ${esc(t('dash.near.count'))}</span>
      <button type="button" class="btn link near-btn" id="nearList" data-date="${esc(near.date)}" aria-pressed="false">${esc(t('dash.near.list'))}</button>
    </div>`;
}

// '목록 보기' — 카드의 마감일에 해당하는 Commission 만 남긴다. 카드 숫자는 시작월 이후
// 차량에서 센 것이므로, 타일처럼 다른 필터를 다 풀고(시작월 조건만 유지) 날짜 조건만 건다
// → 표 건수 = 카드 건수.
function applyNearDate(date) {
  applyTile('t-total');
  activeTile = null;
  tileDate = date;
  syncTileActive();
  render();
}

function syncTileActive() {
  $('#summary').querySelectorAll('.tile').forEach((el) => {
    const on = el.dataset.tile === activeTile;
    el.classList.toggle('active', on);
    el.setAttribute('aria-pressed', on ? 'true' : 'false');
  });
  const nb = $('#nearList');
  if (nb) {
    const on = !!tileDate && tileDate === nb.dataset.date;
    nb.closest('.near-card').classList.toggle('active', on);
    nb.setAttribute('aria-pressed', on ? 'true' : 'false');
  }
}

// 모델 정보 조합(Model · Type · Axle · Cab · MY)을 1개 모델로 보고 대수를 집계한다.
function modelCombos(rows) {
  const m = new Map();
  rows.forEach((r) => {
    const parts = [r.Vehicle, r['Model(WINGS)'], r.Type, r.Cab, r.MY]
      .map((v) => String(v ?? '').trim());
    if (parts.every((p) => !p)) return;
    const key = parts.join('');
    m.set(key, (m.get(key) || 0) + 1);
  });
  return [...m.entries()]
    .map(([k, c]) => ({ parts: k.split(''), count: c }))
    .sort((a, b) => b.count - a.count || a.parts.join(' ').localeCompare(b.parts.join(' ')));
}

function daysFromToday(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const base = new Date(CUR_DATE + 'T00:00:00');
  return Math.round((d - base) / 86400000);
}

// 오늘과 가장 가까운 Changeability Date(우선 미래, 없으면 과거 최신)와 해당일 Commission 수.
function nearestChangeability(rows) {
  const dates = rows
    .map((r) => String(r['Changeability Date'] || ''))
    .filter((s) => /^\d{4}-\d{2}-\d{2}/.test(s))
    .map((s) => s.slice(0, 10));
  if (!dates.length) return null;
  const upcoming = dates.filter((d) => d >= CUR_DATE).sort();
  const chosen = upcoming.length ? upcoming[0] : [...dates].sort().slice(-1)[0];
  const count = rows.filter(
    (r) => String(r['Changeability Date'] || '').slice(0, 10) === chosen).length;
  return { date: chosen, count, dday: daysFromToday(chosen) };
}

function ddayLabel(n) {
  if (n > 0) return 'D-' + n;
  if (n === 0) return 'D-Day';
  return 'D+' + (-n);
}

// 모델별 대수 — 예전의 '전체'·'필터' 두 목록을 한 패널에 두고 토글로 바꾼다.
//   all      = 시작월 이후 전체(overallRows)   filtered = 지금 표에 보이는 행
// 막대를 누르면 Model·Type·Axle·Cab·MY 드롭다운에 그 조합을 넣는다(기존 필터 그대로 사용).
let MODEL_SCOPE = 'all';
let MODEL_EXPANDED = false;
let LAST_ROWS = [];
const MODEL_TOP = 6;
// modelCombos 의 parts 순서(Vehicle · Model(WINGS) · Type · Cab · MY)에 맞춘 드롭다운.
const COMBO_FILTER_IDS = ['#modelFilter', '#typeFilter', '#axleFilter', '#cabFilter', '#myFilter'];

function comboSelected(c) {
  return c.parts.every((p, k) => p === $(COMBO_FILTER_IDS[k]).value);
}

function renderModelPanel() {
  const el = $('#modelPanel');
  if (!el) return;
  const combos = modelCombos(MODEL_SCOPE === 'filtered' ? LAST_ROWS : overallRows());
  const units = combos.reduce((s, c) => s + c.count, 0);
  const max = combos.length ? combos[0].count : 1;
  const shown = MODEL_EXPANDED ? combos : combos.slice(0, MODEL_TOP);
  const rest = combos.slice(MODEL_TOP);
  const restUnits = rest.reduce((s, c) => s + c.count, 0);

  const bars = shown.map((c, i) => {
    const label = c.parts.filter(Boolean).join(' · ');
    const on = comboSelected(c);
    return `<button type="button" class="mbar${on ? ' active' : ''}" data-i="${i}" aria-pressed="${on}"`
      + ` aria-label="${esc(t('dash.models.aria', { label, n: c.count }))}" title="${esc(label)}">`
      + `<span class="mbar-label">${esc(label)}</span><span class="mbar-n">${fmtNum(c.count)}</span>`
      + `<span class="mbar-track" aria-hidden="true"><span class="mbar-fill" style="width:${Math.max(2, c.count / max * 100).toFixed(1)}%"></span></span>`
      + `</button>`;
  }).join('');
  const more = rest.length
    ? `<button type="button" class="btn link mbar-more" aria-expanded="${MODEL_EXPANDED}">${esc(MODEL_EXPANDED
      ? t('dash.models.less') : t('dash.models.more', { n: rest.length, u: fmtNum(restUnits) }))}</button>`
    : '';
  const seg = (scope, key) => `<button type="button" data-scope="${scope}" aria-pressed="${MODEL_SCOPE === scope}"`
    + ` class="${MODEL_SCOPE === scope ? 'active' : ''}">${esc(t(key))}</button>`;

  el.innerHTML = `
    <div class="panel-card">
      <div class="panel-head">
        <h3 class="panel-title">${esc(t('dash.models'))}
          <span class="panel-sub">${esc(t('dash.models.sum', { n: combos.length, u: fmtNum(units) }))}</span></h3>
        <div class="seg" role="group" aria-label="${esc(t('dash.models'))}">
          ${seg('all', 'dash.models.all')}${seg('filtered', 'dash.models.filtered')}
        </div>
      </div>
      <div class="mbar-list${MODEL_EXPANDED ? ' expanded' : ''}">${bars || '<div class="mc-empty">—</div>'}</div>
      ${more}
    </div>`;

  el.querySelectorAll('.seg button').forEach((b) => b.addEventListener('click', () => {
    MODEL_SCOPE = b.dataset.scope;
    renderModelPanel();
  }));
  const mb = el.querySelector('.mbar-more');
  if (mb) mb.addEventListener('click', () => { MODEL_EXPANDED = !MODEL_EXPANDED; renderModelPanel(); });
  el.querySelectorAll('.mbar').forEach((b) =>
    b.addEventListener('click', () => applyModelCombo(shown[Number(b.dataset.i)])));
}

function applyModelCombo(c) {
  if (!c) return;
  const on = comboSelected(c);
  if (MODEL_SCOPE === 'all') {
    // '전체' 숫자는 시작월 이후 전체에서 센 값 — 타일처럼 다른 필터를 풀어야 표 건수 = 막대 숫자.
    applyTile('t-total');
    COMBO_FILTER_IDS.forEach((id, k) => { $(id).value = on ? '' : (c.parts[k] || ''); });
    onFilterChange();
  } else {
    // '필터 적용' — 지금 걸린 조건(타일 포함) 안에서 이 모델로 좁힌다.
    COMBO_FILTER_IDS.forEach((id, k) => { $(id).value = on ? '' : (c.parts[k] || ''); });
    fillFilters();
    render();
  }
}

// 지금 걸린 필터를 칩으로 — ×로 하나씩 풀고, '필터 초기화'로 전부 푼다.
function tileLabel(id) {
  const base = String(id || '').replace(/2$/, '');
  return t({ 't-total': 'tile.total', 't-match': 'tile.match', 't-miss': 'tile.mismatch',
    't-nosam': 'tile.nosam', 't-mand': 'tile.mand', 't-samupd': 'tile.samupd' }[base] || base);
}
function renderChips() {
  const el = $('#chips');
  if (!el) return;
  const chips = [];
  if (activeTile && activeTile !== 't-total') {
    const row = TILE_ACTIONS[activeTile] && TILE_ACTIONS[activeTile].soon ? t('dash.soon') : t('dash.overall');
    chips.push({ k: 'tile', label: `${row} · ${tileLabel(activeTile)}` });
  }
  if (tileDate) chips.push({ k: 'date', label: `${t('chip.date')}: ${tileDate}` });
  for (const f of FILTER_FIELDS) {
    const v = $(f.id).value;
    if (v) chips.push({ k: f.id, label: `${t('export.f.' + f.key)}: ${v}` });
  }
  const st = $('#statusFilter').value;
  // 타일이 넣은 상태값은 타일 칩이 이미 말해 준다.
  if (st && !(activeTile && TILE_ACTIONS[activeTile] && TILE_ACTIONS[activeTile].status === st)) {
    chips.push({ k: '#statusFilter', label: `${t('export.f.status')}: ${st}` });
  }
  const q = $('#search').value.trim();
  if (q) chips.push({ k: 'search', label: `${t('export.f.search')}: ${q}` });
  if (!$('#upcomingOnly').checked) chips.push({ k: 'upcoming', label: t('chip.allMonths') });

  el.innerHTML = chips.map((c) =>
    `<span class="chip">${esc(c.label)}<button type="button" class="chip-x" data-k="${esc(c.k)}"`
    + ` aria-label="${esc(t('chip.remove', { f: c.label }))}"><svg class="ic" aria-hidden="true"><use href="#i-x"/></svg></button></span>`
  ).join('') + (chips.length
    ? `<button type="button" class="btn link chip-reset">${esc(t('chip.reset'))}</button>` : '');

  el.querySelectorAll('.chip-x').forEach((b) => b.addEventListener('click', () => {
    const k = b.dataset.k;
    if (k === 'tile' || k === 'date') { applyTile('t-total'); return; }
    if (k === 'search') { $('#search').value = ''; onManualFilter(); return; }
    if (k === 'upcoming') { $('#upcomingOnly').checked = true; onFilterChange(); return; }
    $(k).value = '';
    onFilterChange();
  }));
  const rb = el.querySelector('.chip-reset');
  if (rb) rb.addEventListener('click', () => applyTile('t-total'));
}

function fillSelectFilter(id, allKey, values) {
  const sel = $(id);
  const prev = sel.value;
  sel.innerHTML = `<option value="">${esc(t(allKey))}</option>` +
    values.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  sel.value = prev;
}

// 각 필터 드롭다운 정의 (id, facet key, 라벨, 값 추출, 내림차순 여부).
const FILTER_FIELDS = [
  { id: '#productionFilter', key: 'prod',  allKey: 'filter.allProd',  get: prodMonth,               desc: true },
  { id: '#myFilter',         key: 'my',    allKey: 'filter.allMY',    get: (r) => r.MY },
  { id: '#modelFilter',      key: 'model', allKey: 'filter.allModel', get: (r) => r.Vehicle },
  { id: '#typeFilter',       key: 'type',  allKey: 'filter.allType',  get: (r) => r['Model(WINGS)'] },
  { id: '#axleFilter',       key: 'axle',  allKey: 'filter.allAxle',  get: (r) => r.Type },
  { id: '#cabFilter',        key: 'cab',   allKey: 'filter.allCab',   get: (r) => r.Cab },
];

// 드롭다운/상태/생산월 체크박스 필터를 모두 만족하는지. skip 에 해당하는 필드는 건너뛴다
// (그 필드의 선택가능 옵션을 계산할 때, 자기 자신은 제외하고 다른 필터만 반영하기 위함).
function matchesFilters(r, skip) {
  if (skip !== 'upcoming' && $('#upcomingOnly').checked) {
    const pm = prodMonth(r);
    if (!pm || pm < DASH_FROM) return false;
  }
  if (skip !== 'status') {
    const st = $('#statusFilter').value;
    if (st && r['SAM Status'] !== st) return false;
  }
  for (const f of FILTER_FIELDS) {
    if (skip === f.key) continue;
    const val = $(f.id).value;
    if (val && String(f.get(r) ?? '') !== val) return false;
  }
  return true;
}

// 특정 필드의 선택가능 값 = 다른 필터를 만족하는 행들의 그 필드 distinct 값(연동 필터).
function facetValues(field) {
  const set = new Set();
  for (const r of DATA.rows) {
    if (!matchesFilters(r, field.key)) continue;
    const v = field.get(r);
    if (v !== undefined && v !== null && v !== '') set.add(String(v));
  }
  const arr = [...set].sort();
  return field.desc ? arr.reverse() : arr;
}

function fillFilters() {
  fillFromMonth();
  for (const f of FILTER_FIELDS) fillSelectFilter(f.id, f.allKey, facetValues(f));
  enhanceFilterSelects();
}

// 시작월 드롭다운 — 데이터에 있는 모든 생산월(내림차순)에 이번 달과 올해 1월을
// 항상 더한다. 그래야 해당 월에 차가 없어도 기준으로 고를 수 있다.
function fillFromMonth() {
  const sel = $('#fromMonth');
  if (!sel) return;
  const set = new Set([CUR_MONTH, YEAR_START]);
  for (const r of (DATA.rows || [])) { const pm = prodMonth(r); if (pm) set.add(String(pm)); }
  const months = [...set].sort().reverse();
  if (!months.includes(DASH_FROM)) DASH_FROM = CUR_MONTH;
  sel.innerHTML = months.map((m) =>
    `<option value="${esc(m)}">${esc(t('filter.from.opt', { m }))}</option>`).join('');
  sel.value = DASH_FROM;
}

// 시작월 변경 — 타일 숫자·연동 필터·표가 모두 같은 범위를 보도록 다시 계산한다.
function setDashFrom(month) {
  DASH_FROM = month || CUR_MONTH;
  renderSummary();
  renderModelPanel();
  fillFilters();
  onManualFilter();
}

// 필터 select 들을 커스텀 드롭다운으로 감싼다(아래로 열림, ~10개 표시 후 스크롤).
const FILTER_SELECT_IDS = ['#productionFilter', '#myFilter', '#modelFilter',
  '#typeFilter', '#axleFilter', '#cabFilter', '#statusFilter', '#fromMonth'];
function enhanceFilterSelects() {
  FILTER_SELECT_IDS.forEach((id) => { const el = $(id); if (el) enhanceSelect(el); });
}

function enhanceSelect(sel) {
  if (sel._csel) { sel._csel.refresh(); return; }
  const wrap = document.createElement('div');
  wrap.className = 'csel';
  sel.parentNode.insertBefore(wrap, sel);
  wrap.appendChild(sel);
  sel.classList.add('csel-native');
  sel.tabIndex = -1;
  sel.setAttribute('aria-hidden', 'true');

  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'csel-btn';
  const panel = document.createElement('div');
  panel.className = 'csel-panel';
  wrap.appendChild(btn);
  wrap.appendChild(panel);

  function renderPanel() {
    panel.innerHTML = '';
    Array.prototype.forEach.call(sel.options, (o) => {
      const it = document.createElement('div');
      it.className = 'csel-opt' + (o.value === sel.value ? ' sel' : '');
      it.textContent = o.textContent;
      it.dataset.value = o.value;
      it.addEventListener('click', () => {
        sel.value = o.value;
        sel.dispatchEvent(new Event('input', { bubbles: true }));
        refresh();
        close();
      });
      panel.appendChild(it);
    });
  }
  function refresh() {
    const cur = sel.options[sel.selectedIndex];
    btn.textContent = cur ? cur.textContent : '';
    Array.prototype.forEach.call(panel.children, (c) =>
      c.classList.toggle('sel', c.dataset.value === sel.value));
  }
  function open() {
    document.querySelectorAll('.csel.open').forEach((w) => { if (w !== wrap) w.classList.remove('open'); });
    wrap.classList.add('open');
    const s = panel.querySelector('.csel-opt.sel');
    if (s) s.scrollIntoView({ block: 'nearest' });
  }
  function close() { wrap.classList.remove('open'); }

  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    wrap.classList.contains('open') ? close() : open();
  });
  sel.addEventListener('change', refresh);
  sel._csel = { refresh: () => { renderPanel(); refresh(); } };
  renderPanel();
  refresh();
}
if (!window._cselOutside) {
  window._cselOutside = true;
  document.addEventListener('click', (e) => {
    document.querySelectorAll('.csel.open').forEach((w) => {
      if (!w.contains(e.target)) w.classList.remove('open');
    });
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') document.querySelectorAll('.csel.open').forEach((w) => w.classList.remove('open'));
  });
}

function renderHead() {
  // 열 제목은 두 언어 모두 영어(시스템 용어). 정렬 상태는 aria-sort 로도 알린다.
  $('#grid thead').innerHTML =
    '<tr>' + COLS.map((c) => {
      const on = sortKey === c.key;
      const arrow = on ? `<span class="arrow" aria-hidden="true">${sortDir === 1 ? '▲' : '▼'}</span>` : '';
      const sort = on ? (sortDir === 1 ? 'ascending' : 'descending') : 'none';
      return `<th scope="col" data-k="${esc(c.key)}" aria-sort="${sort}" tabindex="0">${c.label}${arrow}</th>`;
    }).join('') + '</tr>';
  $('#grid thead').querySelectorAll('th').forEach((th) =>
    th.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); th.click(); }
    }));
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

function hl(text) {
  const q = $('#search').value.trim();
  const safe = esc(text);
  if (!q) return safe;
  try {
    const re = new RegExp('(' + q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
    return safe.replace(re, '<mark>$1</mark>');
  } catch { return safe; }
}

// ✓ = 0개(이상 없음), ✕ n = n개. 색만으로 구분하지 않도록 기호와 aria-label 을 함께 둔다.
function countCell(n, cls) {
  if (!n) return `<span class="mark ok" role="img" aria-label="${esc(t('cell.ok.title'))}" title="${esc(t('cell.ok.title'))}">✓</span>`;
  const label = t('cell.' + cls + '.aria', { n });
  return `<span class="mark bad" role="img" aria-label="${esc(label)}" title="${esc(label)}">✕ ${n}</span>`;
}

// SAM Status 값 → CSS 클래스(공백/비ASCII 안전).
const STATUS_CLASS = {
  'Match': 'Match', 'Mismatch': 'Mismatch', 'No SAM': 'NoSAM', 'SAM update요청': 'SAMupdate',
};
function statusClass(v) {
  return STATUS_CLASS[v] || esc(v).replace(/[^A-Za-z0-9_-]/g, '');
}
// 상태 배지 — 값(Match / Mismatch / No SAM)은 데이터 그대로, 앞에 기호를 붙인다.
const STATUS_ICON = { Match: '✓', Mismatch: '✕', NoSAM: '✕' };
function statusBadge(v) {
  const cls = statusClass(v);
  const ic = STATUS_ICON[cls];
  return `<span class="status ${cls}">${ic ? `<span aria-hidden="true">${ic}</span> ` : ''}${esc(v)}</span>`;
}

function filtered() {
  const q = $('#search').value.trim().toLowerCase();
  let rows = DATA.rows.filter((r) => {
    if (restrictSoon && !within2weeks(r)) return false;
    if (tileMandatory && countOf(r['Mandatory Codes']) === 0) return false;
    if (tileSamUpdate && !r['SAM Update']) return false;
    if (tileDate && String(r['Changeability Date'] || '').slice(0, 10) !== tileDate) return false;
    if (!matchesFilters(r, null)) return false;   // 드롭다운 + 상태 + 생산월 체크박스
    if (q) {
      const hay = Object.values(r).join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
  if (sortKey) {
    const countKind = (COLS.find((c) => c.key === sortKey) || {}).count;
    const numeric = NUMERIC_KEYS.has(sortKey);
    rows = [...rows].sort((a, b) => {
      let cmp;
      if (countKind) {
        cmp = diffCount(a, countKind) - diffCount(b, countKind);
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
  $('#count').textContent = t('count', { n: fmtNum(rows.length), total: fmtNum(DATA.rows.length) });
  LAST_ROWS = rows;
  renderModelPanel();
  renderChips();
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
    msg.innerHTML = `${esc(t('msg.noRows'))} <button type="button" class="btn link" id="emptyReset">${esc(t('chip.reset'))}</button>`;
    $('#emptyReset').addEventListener('click', () => applyTile('t-total'));
    msg.classList.remove('hidden');
    return;
  }
  msg.classList.add('hidden');

  tb.innerHTML = rows.map((r, i) => {
    const tds = COLS.map((c) => {
      let v = r[c.key];
      v = v == null ? '' : String(v);
      if (c.status) {
        const upd = r['SAM Update'] ? `<span class="status SAMupdate upd">${esc(t('tile.samupd'))}</span>` : '';
        const exc = r._accepted
          ? `<span class="status exc upd" title="${esc(t('exc.chip.title'))}">${esc(t('exc.chip'))}</span>`
          : (r._accN ? `<span class="status exc part upd" title="${esc(t('exc.chipPart.title'))}">${esc(t('exc.chipPart', { n: r._accN }))}</span>` : '');
        return `<td><div class="status-cell">${statusBadge(v)}${exc}${upd}</div></td>`;
      }
      // Commission — 상세 창을 여는 버튼(행 클릭과 같은 동작, 키보드로도 열린다).
      if (c.key === 'Commission no.') {
        return `<td class="c-comm"><button type="button" class="comm-link" aria-label="${esc(t('table.open', { c: v }))}">${hl(v)}</button></td>`;
      }
      if (c.dday) return `<td class="num">${ddayHtml(v)}</td>`;
      if (c.count) return `<td class="num">${countCell(diffCount(r, c.count), c.count)}</td>`;
      if (c.pair && pairMismatch(r, c.pair)) {
        return `<td class="pair-diff" title="${esc(pairTitle(r, c.pair))}">${hl(v)}<span class="ne">≠</span></td>`;
      }
      return `<td>${hl(v)}</td>`;
    }).join('');
    return `<tr data-i="${i}">${tds}</tr>`;
  }).join('');
  tb.querySelectorAll('tr').forEach((tr) =>
    tr.addEventListener('click', () => openDrawer(rows[Number(tr.dataset.i)]))
  );
}

function describe(code) { return CODES.options[code] || CODES.mandatory[code] || ''; }
function countOf(csv) { return (csv || '').split(',').map((c) => c.trim()).filter(Boolean).length; }
function splitCodes(csv) { return (csv || '').split(',').map((c) => c.trim()).filter(Boolean); }

// 페인트·타이어는 Only_in_* CSV 에 들어가지 않고 별도 그룹으로 비교된다(상세 창의 🎨/🛞 섹션).
// 목록·정렬·Export 의 '차이 개수'에는 그 그룹 차이도 함께 센다 — 타이어만 다른 차량이
// Mismatch 인데 목록에는 ✓ 로 보이던 문제를 없애기 위함.
// 캡은 compare 가 WINGS/SAM 양쪽 값을 남긴다(_cab_wings / _cab_sam). 목록엔 한 값만
// 들어가므로, 양쪽이 다르면 셀에 ≠ 를 붙이고 툴팁에 두 값을 보여준다.
// PTO 는 따로 대조하지 않는다 — PTO 코드 차이는 일반 코드 비교에 그대로 나온다.
function pairSides(r, kind) {
  return { wings: splitCodes(r['_' + kind + '_wings']), sam: splitCodes(r['_' + kind + '_sam']) };
}
function pairMismatch(r, kind) {
  if (!r || !r['Compared SAM file name']) return false;   // 비교 대상 자체가 없으면 불일치 아님
  const { wings, sam } = pairSides(r, kind);
  if (!wings.length || !sam.length) return false;         // 한쪽에 신호가 없으면 비교 안 함
  return wings.join(',') !== sam.join(',');
}
function pairTitle(r, kind) {
  const { wings, sam } = pairSides(r, kind);
  return `WINGS: ${wings.join(', ') || '—'}  /  SAM: ${sam.join(', ') || '—'}`;
}

const DIFF_GROUPS = ['_paint', '_tyre'];
// 원본 차이 — 'Mismatch 예외' 승인 여부와 상관없이 compare 가 낸 그대로.
function rawDiffCodes(r, kind) {
  if (kind === 'mand') return splitCodes(r['Mandatory Codes']);
  const mine = kind === 'sam' ? '_sam' : '_wings';
  const other = kind === 'sam' ? '_wings' : '_sam';
  const out = splitCodes(kind === 'sam' ? r['Only_in_SAM'] : r['Only_in_WINGS']);
  for (const g of DIFF_GROUPS) {
    const a = splitCodes(r[g + mine]);
    const b = new Set(splitCodes(r[g + other]));
    // 한쪽 그룹이 비어 있으면 비교하지 않는다 (compare.js 의 paint/tyre 판정과 같은 규칙).
    if (!a.length || !b.size) continue;
    for (const c of a) if (!b.has(c)) out.push(c);
  }
  return out;
}
// 화면·Export 가 쓰는 차이 — 관리자가 '문제없음'으로 승인한 코드는 뺀다.
function diffCodes(r, kind) {
  const raw = rawDiffCodes(r, kind);
  return kind === 'mand' ? raw : raw.filter((c) => !isAccepted(r, kind, c));
}
function diffCount(r, kind) { return diffCodes(r, kind).length; }

// ====================== Mismatch 예외 (문제없는 차이) ======================
// Admin 이 '문제없음'으로 표시한 차이 코드는 더 이상 차이로 세지 않는다. 한 차량의 차이
// 코드가 모두 표시되면 대시보드에서 Match 로 분류된다. 별도 승인 단계는 없다 — 등록·노트
// 수정·해제는 Admin 만 할 수 있고, 모든 변경은 SharePoint
// '06. exceptions/exceptions.json' 의 log 에 쌓인다(지우지 않는다). 원래 판정은 r._orig_status.
//
//   단위(key) = Commission | 구분(sam = SAM에만, win = WINGS에만) | 코드
//   상태      = approved(문제없음) ⇄ revoked(해제)
//   (예전 승인 절차 시절의 pending/rejected/cancelled 기록은 '미반영'으로만 보여 준다.)
//
// 빌드 결과(data.json)는 건드리지 않는다 — 다시 계산 없이 화면을 여는 순간 반영된다.
const EXC_FILE = 'exceptions.json';
const EXC_LOCAL_KEY = 'sam-afab-exceptions';   // 로그인 없는 로컬 개발용 저장소
let EXC = { version: 1, items: {}, log: [], updated_at: '' };
let EXC_STATE = 'idle';   // 'ok' | 'none'(아직 파일 없음) | 'error' | 'local'
let EXC_ERR = '';

function excKey(comm, side, code) { return `${String(comm).trim()}|${side}|${code}`; }
function excItem(r, side, code) { return r ? EXC.items[excKey(r['Commission no.'], side, code)] : undefined; }
function isAccepted(r, side, code) { const it = excItem(r, side, code); return !!it && it.status === 'approved'; }
function excLocal() { return !(window.Graph && Graph.available()); }

function excUser() {
  const a = window.MB_AUTH && typeof window.MB_AUTH.account === 'function' ? window.MB_AUTH.account() : null;
  const id = String((a && a.username) || 'local').toLowerCase();
  return { id, name: (a && (a.name || a.username)) || id };
}

function excNormalize(d) {
  d = d && typeof d === 'object' ? d : {};
  return {
    version: 1,
    items: d.items && typeof d.items === 'object' ? d.items : {},
    log: Array.isArray(d.log) ? d.log : [],
    updated_at: d.updated_at || '',
  };
}

async function excRead() {
  if (excLocal()) {
    let d = null;
    try { d = JSON.parse(localStorage.getItem(EXC_LOCAL_KEY) || 'null'); } catch { /* 저장소 막힘 */ }
    return { data: d, etag: null };
  }
  return Graph.readJsonVersioned('exceptions', EXC_FILE);
}

async function loadExceptions() {
  try {
    const { data } = await excRead();
    EXC = excNormalize(data);
    EXC_STATE = excLocal() ? 'local' : (data ? 'ok' : 'none');
    EXC_ERR = '';
  } catch (e) {
    console.warn('[exceptions] 읽기 실패 — 예외 미적용:', e.message);
    EXC_STATE = 'error';
    EXC_ERR = e.message;
  }
}

// 차량별 상태를 예외 기준으로 다시 매긴다. 원래 판정은 _orig_status 에 보존.
function excRowCodes(r) {
  return [...rawDiffCodes(r, 'sam').map((c) => ['sam', c]), ...rawDiffCodes(r, 'win').map((c) => ['win', c])];
}
function applyExceptions() {
  for (const r of DATA.rows || []) {
    if (r._orig_status === undefined) r._orig_status = r['SAM Status'];
    r._accepted = false;
    r._accN = 0;
    if (r._orig_status !== 'Mismatch') { r['SAM Status'] = r._orig_status; continue; }
    const codes = excRowCodes(r);
    r._accN = codes.filter(([s, c]) => isAccepted(r, s, c)).length;
    r._accepted = codes.length > 0 && r._accN === codes.length;
    r['SAM Status'] = r._accepted ? 'Match' : 'Mismatch';
  }
}

// 모든 변경은 Admin 만. accept = 문제없음 등록(해제됐거나 예전 요청이던 것도 다시 등록),
// edit = 노트 수정, revoke = 해제.
const EXC_STATUS_OF = { accept: 'approved', edit: 'approved', revoke: 'revoked',
  // 예전 승인 절차 기록(변경 이력 표시용)
  request: 'pending', approve: 'approved', reject: 'rejected', cancel: 'cancelled' };
function excCan(action, it) {
  if (!IS_ADMIN) return false;
  const st = it ? it.status : '';
  switch (action) {
    case 'accept': return st !== 'approved';
    case 'edit':
    case 'revoke': return st === 'approved';
    default: return false;
  }
}

function rowByComm(comm) {
  const k = String(comm).trim();
  return (DATA.rows || []).find((r) => String(r['Commission no.']).trim() === k) || null;
}

// 읽고 → 바꾸고 → eTag 조건부로 쓴다. 그 사이 누가 고쳤으면 다시 읽어서 처음부터.
// targets = [{ comm, side, code }]. 지금 상태에서 할 수 없는 항목은 건너뛴다.
async function excMutate(action, targets, text) {
  const u = excUser();
  for (let attempt = 0; attempt < 4; attempt++) {
    const { data, etag } = await excRead();
    const doc = excNormalize(data);
    const now = new Date().toISOString();
    let done = 0;
    for (const tg of targets) {
      const key = excKey(tg.comm, tg.side, tg.code);
      const it = doc.items[key];
      if (!excCan(action, it)) continue;
      if (action === 'accept') {
        const r = rowByComm(tg.comm) || {};
        doc.items[key] = {
          commission: String(tg.comm).trim(), side: tg.side, code: tg.code, status: 'approved',
          note: text || '', by: u.id, by_name: u.name, at: now, revoke_note: '',
          // 그 차량이 다음 빌드에서 빠져도 무엇에 대한 예외였는지 알 수 있게 남긴다.
          vehicle: r['Vehicle'] || (it && it.vehicle) || '', type: r['Model(WINGS)'] || (it && it.type) || '',
          axle: r['Type'] || (it && it.axle) || '', production: r['Production date'] || (it && it.production) || '',
          sam_file: r['Compared SAM file name'] || (it && it.sam_file) || '',
        };
      } else if (action === 'edit') {
        it.note = text || '';
        it.by = u.id; it.by_name = u.name; it.at = now;
      } else if (action === 'revoke') {
        it.status = 'revoked';
        it.revoke_note = text || '';
        it.by = u.id; it.by_name = u.name; it.at = now;
      }
      doc.log.push({ at: now, by: u.id, by_name: u.name, action, key,
        commission: String(tg.comm).trim(), side: tg.side, code: tg.code, text: text || '' });
      done++;
    }
    if (!done) { EXC = doc; return 0; }
    doc.updated_at = now;
    if (excLocal()) {
      try { localStorage.setItem(EXC_LOCAL_KEY, JSON.stringify(doc)); } catch { /* 이번 세션만 */ }
      EXC = doc;
      return done;
    }
    try {
      if (!etag) await Graph.ensureFolder('exceptions');
      await Graph.writeJsonVersioned('exceptions', EXC_FILE, doc, etag);
      EXC = doc;
      EXC_STATE = 'ok';
      return done;
    } catch (e) {
      if (e.conflict && attempt < 3) continue;   // 다른 사람이 먼저 저장 — 최신본에 다시 적용
      throw e;
    }
  }
  return 0;
}

// 노트 기본값 — 노트 수정이면 지금 노트, 예전 요청이 있던 코드면 그 사유를 미리 채운다.
function excNoteDefault(action, targets) {
  if (targets.length !== 1) return '';
  const tg = targets[0];
  const it = EXC.items[excKey(tg.comm, tg.side, tg.code)];
  if (!it) return '';
  if (action === 'edit') return it.note || '';
  if (action === 'accept') return it.note || it.reason || '';
  return '';
}

// 버튼 동작 공통 — 노트 입력 → 저장 → 화면 전체 갱신.
let EXC_BUSY = false;
async function excAction(action, targets) {
  if (EXC_BUSY || !targets.length) return;
  if (!IS_ADMIN) { alert(t('role.denied')); return; }
  let text = prompt(t('exc.ask.' + action, { n: targets.length }), excNoteDefault(action, targets));
  if (text === null) return;
  text = text.trim();
  if (!text) { alert(t('exc.ask.noteReq')); return; }
  EXC_BUSY = true;
  document.body.classList.add('exc-busy');
  try {
    const n = await excMutate(action, targets, text);
    if (n < targets.length) alert(t('exc.skipped', { n: targets.length - n }));
  } catch (e) {
    console.error(e);
    alert(t(e.status === 403 || e.status === 401 ? 'exc.err.perm' : 'exc.err.save', { err: e.message }));
  } finally {
    EXC_BUSY = false;
    document.body.classList.remove('exc-busy');
    excRefreshAll();
  }
}

// 예외가 바뀌면 상태·타일·표·상세 창·예외 화면을 모두 다시 그린다.
function excRefreshAll() {
  applyExceptions();
  renderSummary();
  fillFilters();
  render();
  if (VIEW_INIT.history) histRender();
  if (VIEW_INIT.exceptions) excRender();
  if (DRAWER_ROW && !$('#drawer').classList.contains('hidden')) openDrawer(DRAWER_ROW);
}

function excSideLabel(side) { return t(side === 'sam' ? 'exc.side.sam' : 'exc.side.win'); }
// approved / revoked 외(예전 승인 절차의 pending·rejected·cancelled)는 모두 '미반영'.
function excStatusKey(it) {
  if (!it) return 'none';
  return it.status === 'approved' || it.status === 'revoked' ? it.status : 'legacy';
}
function excStateHtml(it) {
  const st = excStatusKey(it);
  return `<span class="exc-st ${st}">${esc(t('exc.st.' + st))}</span>`;
}
function excWho(name, at) {
  return [name, at ? fmtKst(at) : ''].filter(Boolean).join(' · ');
}
// 예전 기록은 requested_*/decided_* 필드를 쓴다 — 마지막 처리자로 읽을 때 함께 본다.
function excLastBy(it) {
  return { name: it.by_name || it.decided_by_name || it.requested_by_name || '',
    at: it.at || it.decided_at || it.requested_at || '' };
}
// 한 코드에 대해 Admin 이 누를 수 있는 버튼들 (Read 는 없음).
function excButtonsHtml(it, comm, side, code) {
  const btn = (a, cls) => excCan(a, it)
    ? `<button type="button" class="btn xs ${cls || ''}" data-exc="${a}" data-comm="${esc(comm)}" data-side="${side}" data-code="${esc(code)}">${esc(t('exc.act.' + a))}</button>`
    : '';
  return btn('accept', 'ok') + btn('edit') + btn('revoke', 'danger');
}
function excTargetOf(b) { return { comm: b.dataset.comm, side: b.dataset.side, code: b.dataset.code }; }

// 상세 창 — 원래 Mismatch 인 차량에만, 차이 코드마다 예외 상태(와 Admin 버튼)를 보여 준다.
function excSecHtml(r) {
  if (r._orig_status !== 'Mismatch') return '';
  const comm = String(r['Commission no.']).trim();
  const codes = excRowCodes(r);
  if (!codes.length) return '';
  const rows = codes.map(([side, code]) => {
    const it = excItem(r, side, code);
    const st = excStatusKey(it);
    let meta = '';
    if (st === 'approved') {
      const w = excLastBy(it);
      meta = `${esc(excWho(w.name, w.at))}` + (it.note ? ` — “${esc(it.note)}”` : '');
    } else if (st === 'revoked') {
      const w = excLastBy(it);
      meta = `${esc(t('exc.st.revoked'))} ${esc(excWho(w.name, w.at))}`
        + (it.revoke_note ? ` — “${esc(it.revoke_note)}”` : '');
    } else if (st === 'legacy') {
      meta = `${esc(t('exc.meta.legacy'))} ${esc(excWho(it.requested_by_name, it.requested_at))}`
        + (it.reason ? ` — “${esc(it.reason)}”` : '');
    }
    const chip = side === 'sam' ? 'crit' : 'accent';
    return `<div class="exc-row${st === 'approved' ? ' accepted' : ''}" data-q="${qAttr(code, describe(code))}">`
      + `<span class="exc-side">${esc(excSideLabel(side))}</span>`
      + `<span class="cchip ${chip}">${esc(code)}</span>`
      + `<span class="cdesc">${esc(describe(code) || '—')}</span>`
      + excStateHtml(it)
      + `<span class="exc-acts">${excButtonsHtml(it, comm, side, code)}</span>`
      + (meta ? `<div class="exc-meta">${meta}</div>` : '')
      + `</div>`;
  }).join('');
  const open = codes.filter(([s, c]) => excCan('accept', excItem(r, s, c)));
  const all = open.length > 1
    ? `<button type="button" class="btn xs ok" data-exc-all="accept">${esc(t('exc.act.acceptAll', { n: open.length }))}</button>`
    : '';
  const n = r._accN || 0;
  return `<section class="dsec exc-sec">
      <header class="dsec-head">
        <h3>${esc(t('exc.sec.title'))}</h3><span class="cnt ${n === codes.length ? 'zero' : 'crit'}">${n}/${codes.length}</span>
        ${all}
        <span class="dsec-hint">${esc(t(IS_ADMIN ? 'exc.sec.hint' : 'exc.sec.hintRead'))}</span>
      </header>
      <div class="dsec-body">${rows}</div>
    </section>`;
}

function wireExcSec(root, r) {
  const sec = root.querySelector('.exc-sec');
  if (!sec) return;
  sec.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-exc]');
    if (b) { excAction(b.dataset.exc, [excTargetOf(b)]); return; }
    const all = e.target.closest('button[data-exc-all]');
    if (all) {
      const comm = String(r['Commission no.']).trim();
      const targets = excRowCodes(r).filter(([s, c]) => excCan('accept', excItem(r, s, c)))
        .map(([side, code]) => ({ comm, side, code }));
      excAction('accept', targets);
    }
  });
}

let DRAWER_ROW = null;
let DRAWER_TAB = 'diff';
let DRAWER_RETURN = null;   // 닫으면 포커스를 돌려줄 요소

// 코드 검색(상세 창 오른쪽 위)이 훑는 문자열 — 코드 + 설명, 소문자.
function qAttr(...parts) { return esc(parts.filter(Boolean).join(' ').toLowerCase()); }

// ---- 전체 코드 목록 탭 — 구조(좌 SAM / 우 WINGS 정렬, Paint·Tyre 블록)는 그대로 둔다 ----
// heads: 좌/우 머리글 — 기본은 상세 창의 SAM / WINGS. 'SAM 파일 비교'는 A / B 로 쓴다.
function alignedFullHtml(samCsv, wingsCsv, heads) {
  const sam = new Set(splitCodes(samCsv));
  const wings = new Set(splitCodes(wingsCsv));
  const union = [...new Set([...sam, ...wings])].sort();
  const cell = (code, present) => present
    ? `<span class="c">${esc(code)}</span><span class="d">${esc(describe(code) || '—')}</span>`
    : '';
  const rows = union.map((code) => {
    const inS = sam.has(code), inW = wings.has(code);
    return `<div class="acode-row" data-q="${qAttr(code, describe(code))}">`
      + `<div class="acode-cell${inS ? '' : ' miss'}">${cell(code, inS)}</div>`
      + `<div class="acode-cell${inW ? '' : ' miss'}">${cell(code, inW)}</div>`
      + `</div>`;
  }).join('');
  return `<div class="acode">
      <div class="acode-head">
        <h4>${esc(heads ? heads[0] : 'All SAM Codes')} <span class="badge">${sam.size}</span></h4>
        <h4>${esc(heads ? heads[1] : 'All WINGS Codes')} <span class="badge">${wings.size}</span></h4>
      </div>
      <div class="acode-body">${rows}</div>
    </div>`;
}

// Aligned SAM-vs-WINGS block for a named category (Paint / Tyre). Same two-column
// layout as alignedFullHtml but with its own heading; returns '' when there is no
// data on either side so empty groups don't clutter the chart.
function alignedGroupHtml(title, samCsv, wingsCsv, kind, diffOnly, heads) {
  const sam = new Set(splitCodes(samCsv));
  const wings = new Set(splitCodes(wingsCsv));
  let union = [...new Set([...sam, ...wings])].sort();
  // diffOnly (Difference Codes tab): keep only codes present on exactly one side,
  // and hide the whole section when the two sides are identical.
  if (diffOnly) union = union.filter((c) => sam.has(c) !== wings.has(c));
  if (!union.length) return '';
  const desc = (code) => kind === 'paint' ? `MB ${code}`
    : (kind === 'tyre' ? '' : (describe(code) || '—'));
  const cell = (code, present) => present
    ? `<span class="c">${esc(code)}</span><span class="d">${esc(desc(code))}</span>`
    : '';
  const rows = union.map((code) => {
    const inS = sam.has(code), inW = wings.has(code);
    return `<div class="acode-row" data-q="${qAttr(code, desc(code))}">`
      + `<div class="acode-cell${inS ? '' : ' miss'}">${cell(code, inS)}</div>`
      + `<div class="acode-cell${inW ? '' : ' miss'}">${cell(code, inW)}</div>`
      + `</div>`;
  }).join('');
  // 차이 탭에서는 '한쪽에만 있는 코드' 수(=실제로 표시된 줄 수)를, 전체 탭에서는 그룹 전체 수를 뱃지에 쓴다.
  const nS = diffOnly ? union.filter((c) => sam.has(c)).length : sam.size;
  const nW = diffOnly ? union.filter((c) => wings.has(c)).length : wings.size;
  return `<div class="acode" style="margin-bottom:16px">
      <div class="acode-head">
        <h4>${esc(title)} · ${esc(heads ? heads[0] : 'SAM')} <span class="badge">${nS}</span></h4>
        <h4>${esc(title)} · ${esc(heads ? heads[1] : 'WINGS')} <span class="badge">${nW}</span></h4>
      </div>
      <div class="acode-body">${rows}</div>
    </div>`;
}

// ---- 차이 코드 탭 — 2×2 섹션 ----
function okLine(text) {
  return `<div class="ok-line"><svg class="ic" aria-hidden="true"><use href="#i-check-circle"/></svg>${esc(text)}</div>`;
}
function codeItemHtml(code, chip, desc, extra) {
  return `<div class="crow${extra ? ' ' + extra : ''}" data-q="${qAttr(code, desc)}">`
    + `<span class="cchip ${chip}">${esc(code)}</span><span class="cdesc">${esc(desc || '—')}</span></div>`;
}
function markHtml(side, has) {
  const label = `${side} ${t(has ? 'mark.yes' : 'mark.no')}`;
  return `<span class="mark ${has ? 'ok' : 'bad'}" role="img" aria-label="${esc(label)}" title="${esc(label)}">${has ? '✓' : '✕'}</span>`;
}
function diffSecHtml(o) {
  return `<section class="dsec">
      <header class="dsec-head">
        <h3>${esc(o.title)}</h3><span class="cnt ${o.n ? o.cls : 'zero'}">${o.n}</span>
        <span class="dsec-hint">${esc(o.hint)}</span>
      </header>
      <div class="dsec-body">${o.body}</div>
    </section>`;
}
// 한쪽에만 있는 옵션코드 목록(SAM에만 / WINGS에만).
// r·side 를 주면 'Mismatch 예외'로 승인된 코드는 흐리게(취소선) 표시한다.
function onlySecHtml(title, hint, cls, csv, r, side) {
  const codes = splitCodes(csv);
  const body = codes.length
    ? codes.map((c) => codeItemHtml(c, cls, describe(c), r && isAccepted(r, side, c) ? 'accepted' : '')).join('')
    : okLine(t('code.okDiff'));
  return diffSecHtml({ title, hint, cls, n: codes.length, body });
}
// Mandatory — 빠진 필수코드(한쪽에만 있음)와, 양쪽 모두 들어간 필수코드를 한 표로.
function mandSecHtml(r) {
  const allSam = new Set(splitCodes(r['_all_sam_codes']));
  const allWings = new Set(splitCodes(r['_all_wings_codes']));
  // 누락 판정은 compare.js 가 표기를 맞춘 코드(AJxxx=Jxxx 등)로 했으므로, 어느 쪽에
  // 있는지도 같은 기준으로 본다.
  const canon = (window.Compare && Compare._canonCode) || ((c) => c);
  const samC = new Set([...allSam].map(canon));
  const wingsC = new Set([...allWings].map(canon));
  const missing = splitCodes(r['Mandatory Codes']);
  const mandSet = CODES.mandatory || {};
  const matched = [...allSam].filter((c) => allWings.has(c) && c in mandSet && !missing.includes(c)).sort();

  const row = (c, chip, sHas, wHas) =>
    `<div class="mrow" data-q="${qAttr(c, describe(c))}"><span class="cchip ${chip}">${esc(c)}</span>`
    + `<span class="cdesc">${esc(describe(c) || '—')}</span>`
    + `<span class="mcell">${markHtml('SAM', sHas)}</span><span class="mcell">${markHtml('WINGS', wHas)}</span></div>`;
  const rows = missing.map((c) => row(c, 'crit', samC.has(c) || allSam.has(c), wingsC.has(c) || allWings.has(c)))
    .concat(matched.map((c) => row(c, 'ok', true, true))).join('');
  const head = rows
    ? `<div class="mrow mhead" aria-hidden="true"><span>${esc(t('col.code'))}</span><span>${esc(t('col.desc'))}</span><span class="mcell">SAM</span><span class="mcell">WINGS</span></div>`
    : '';
  const body = (missing.length ? '' : okLine(t('code.okMand'))) + head + rows;
  return diffSecHtml({ title: t('sec.mand'), hint: t('sec.mand.hint'), cls: 'crit', n: missing.length, body });
}
// Factory Control — 참고용. 처음 5개만 보이고 나머지는 '더 보기'.
const FC_TOP = 5;
function fcSecHtml(r) {
  const codes = splitCodes(r['Factory Control Codes']);
  const items = codes.map((c, i) => codeItemHtml(c, 'accent', describe(c), i >= FC_TOP ? 'more' : '')).join('');
  const more = codes.length > FC_TOP
    ? `<button type="button" class="btn link sec-more" aria-expanded="false" data-n="${codes.length - FC_TOP}">${esc(t('sec.more', { n: codes.length - FC_TOP }))}</button>`
    : '';
  const body = codes.length ? `<div class="fc-list">${items}</div>${more}` : okLine(t('code.okFc'));
  return diffSecHtml({ title: t('sec.fc'), hint: t('sec.fc.hint'), cls: 'accent', n: codes.length, body });
}

// ---- 상단 요약 카드 ----
function fieldHtml(label, valueHtml, cls) {
  return `<div class="fld${cls ? ' ' + cls : ''}"><div class="fld-l">${esc(label)}</div><div class="fld-v">${valueHtml}</div></div>`;
}
function fv(r, k) {
  const v = r[k];
  return (v == null || String(v).trim() === '') ? '<span class="empty-cell">—</span>' : esc(v);
}
function ddayPassed(v) {
  const s = String(v ?? '').trim();
  if (!s) return false;
  if (s.toLowerCase() === 'passed') return true;
  const n = Number(s);
  return !Number.isNaN(n) && n < 0;
}

function summaryCardsHtml(r) {
  const cab = (r.Cab && pairMismatch(r, 'cab'))
    ? `<span class="pair-diff" title="${esc(pairTitle(r, 'cab'))}">${esc(r.Cab)}<span class="ne">≠</span></span>`
    : fv(r, 'Cab');
  const fin = String(r.FIN || '').trim();
  const finHtml = fin
    ? `<span class="mono">${esc(fin)}</span><button type="button" class="copy-btn" id="copyFin" data-v="${esc(fin)}"`
      + ` aria-label="${esc(t('copy.fin'))}" title="${esc(t('copy.fin'))}"><svg class="ic" aria-hidden="true"><use href="#i-copy"/></svg></button>`
    : '<span class="empty-cell">—</span>';
  const vehicle = `<section class="scard sc-vehicle"><h3 class="scard-t">${esc(t('g.vehicle'))}</h3><div class="fgrid g4">`
    + fieldHtml('Model', fv(r, 'Vehicle')) + fieldHtml('Category', fv(r, 'Category'))
    + fieldHtml('Axle', fv(r, 'Type')) + fieldHtml('Cab', cab)
    + fieldHtml('MY', fv(r, 'MY')) + fieldHtml('PTO', fv(r, 'PTO'))
    + fieldHtml('FIN', finHtml, 'span2') + `</div></section>`;

  const passed = ddayPassed(r['Until Dealine']);
  const ofs = String(r['Order status financial'] || '').trim();
  const schedule = `<section class="scard"><h3 class="scard-t">${esc(t('g.schedule'))}</h3><div class="fgrid g2">`
    + fieldHtml('Production', fv(r, 'Production date')) + fieldHtml('Changeability', fv(r, 'Changeability Date'))
    + fieldHtml('Changeability D-Day', ddayHtml(r['Until Dealine'])
      + (passed ? ` <span class="fld-note">${esc(t('dday.passedNote'))}</span>` : ''), ofs ? '' : 'span2')
    + (ofs ? fieldHtml('Order status financial', esc(ofs)) : '')
    + `</div></section>`;

  // SAM 대조 — 값·차이 표시·상태만. 'SAM now' 가 Baumuster 와 다르다고 오류는 아니므로
  // 고치라는 문구 없이 회색 칩만 붙인다.
  const bm = String(r['SAM Baumuster'] || '').trim();
  const now = String(r['SAM now'] || '').trim();
  const differs = bm && now && bm !== now;
  const nowHtml = fv(r, 'SAM now') + (differs
    ? ` <span class="nchip" title="${esc(t('sam.differsTitle'))}">${esc(t('sam.differs'))}</span>` : '');
  const st = String(r['SAM Status'] || '');
  const upd = r['SAM Update'] ? t('fileUpdate.yes') : t('fileUpdate.no');
  const sam = `<section class="scard"><h3 class="scard-t">${esc(t('g.sam'))}</h3><div class="fgrid g2">`
    + fieldHtml('SAM Baumuster', fv(r, 'SAM Baumuster')) + fieldHtml('SAM now', nowHtml)
    + fieldHtml('SAM Status', (st ? statusBadge(st) : '<span class="empty-cell">—</span>')
      + (r._accepted ? ` <span class="nchip" title="${esc(t('exc.chip.title'))}">${esc(t('exc.chip.orig'))}</span>` : ''))
    + fieldHtml(t('f.samUpdate'), `<span class="npill${r['SAM Update'] ? ' on' : ''}">${esc(upd)}</span>`)
    + `</div></section>`;
  return `<div class="scards">${vehicle}${schedule}${sam}</div>`;
}

function openDrawer(r) {
  if (!r) return;
  const drawer = $('#drawer');
  const wasOpen = !drawer.classList.contains('hidden');
  if (!wasOpen) DRAWER_RETURN = document.activeElement;
  if (DRAWER_ROW !== r) DRAWER_TAB = 'diff';
  const q = wasOpen && DRAWER_ROW === r && $('#codeQ') ? $('#codeQ').value : '';
  DRAWER_ROW = r;

  // 제목 한 줄: Commission · Model Type · Axle · Cab · MY  + 상태 · D-Day 배지.
  // Column semantics: Model=Vehicle, Type=Model(WINGS), Axle=Type.
  const parts = [[r['Vehicle'], r['Model(WINGS)']].filter(Boolean).join(' '), r['Type'], r['Cab'],
    r['MY'] ? 'MY' + r['MY'] : '']
    .map((x) => (x == null ? '' : String(x).trim())).filter(Boolean);
  const st = String(r['SAM Status'] || '');
  $('#drawerTitle').innerHTML = `<span class="mono dh-comm">${esc(r['Commission no.'])}</span>`
    + (parts.length ? `<span class="dh-sep" aria-hidden="true">·</span><span>${esc(parts.join(' · '))}</span>` : '')
    + `<span class="dh-badges">${st ? statusBadge(st) : ''}${String(r['Until Dealine'] ?? '').trim() ? ddayHtml(r['Until Dealine']) : ''}</span>`;
  const file = r['Compared SAM file name'] || '';
  $('#drawerSub').innerHTML = `<svg class="ic" aria-hidden="true"><use href="#i-file"/></svg>`
    + (file ? `<span title="${esc(t('drawer.fileTitle'))}">${esc(file)}</span>` : `<span class="empty-cell">${esc(t('drawer.noFile'))}</span>`);

  const nDiff = diffCount(r, 'sam') + diffCount(r, 'win');
  const nFull = new Set([...splitCodes(r['_all_sam_codes']), ...splitCodes(r['_all_wings_codes'])]).size;
  const tab = (name, label, n) => `<button type="button" role="tab" class="dtab" id="dtab-${name}" data-tab="${name}"`
    + ` aria-controls="dpane-${name}" aria-selected="false">${esc(label)} <span class="cnt">${fmtNum(n)}</span></button>`;

  $('#drawerBody').innerHTML = `
    ${summaryCardsHtml(r)}
    <div class="dtabs-bar">
      <div class="dtabs" role="tablist">${tab('diff', t('tab.diff'), nDiff)}${tab('full', t('tab.full'), nFull)}</div>
      <label class="search-box sm">
        <svg class="ic" aria-hidden="true"><use href="#i-search"/></svg>
        <input type="search" id="codeQ" placeholder="${esc(t('code.searchPh'))}" aria-label="${esc(t('code.searchPh'))}" />
      </label>
    </div>
    <div class="dpane" id="dpane-diff" role="tabpanel" aria-labelledby="dtab-diff" data-pane="diff">
      ${excSecHtml(r)}
      ${alignedGroupHtml('🎨 Paint', r['_paint_sam'], r['_paint_wings'], 'paint', true)}
      ${alignedGroupHtml('🛞 Tyre', r['_tyre_sam'], r['_tyre_wings'], 'tyre', true)}
      <div class="dgrid">
        ${onlySecHtml(t('sec.samOnly'), t('sec.samOnly.hint'), 'crit', r['Only_in_SAM'], r, 'sam')}
        ${onlySecHtml(t('sec.winOnly'), t('sec.winOnly.hint'), 'accent', r['Only_in_WINGS'], r, 'win')}
        ${mandSecHtml(r)}
        ${fcSecHtml(r)}
      </div>
      <div class="q-none hidden">${esc(t('code.noHit'))}</div>
    </div>
    <div class="dpane hidden" id="dpane-full" role="tabpanel" aria-labelledby="dtab-full" data-pane="full">
      <!-- 캡·PTO 는 3자리 옵션코드라 아래 전체 목록에 이미 들어 있다. 여기서 또 세우면
           같은 코드가 두 번 나오므로 전체 탭에서는 빼고, 차이 탭에만 남긴다. -->
      ${alignedGroupHtml('🎨 Paint', r['_paint_sam'], r['_paint_wings'], 'paint')}
      ${alignedGroupHtml('🛞 Tyre', r['_tyre_sam'], r['_tyre_wings'], 'tyre')}
      ${alignedFullHtml(r['_all_sam_codes'], r['_all_wings_codes'])}
      <div class="q-none hidden">${esc(t('code.noHit'))}</div>
    </div>
  `;
  const body = $('#drawerBody');
  wireExcSec(body, r);
  body.querySelectorAll('.dtab').forEach((tb) => tb.addEventListener('click', () => switchTab(tb.dataset.tab)));
  body.querySelector('.dtabs').addEventListener('keydown', (e) => {
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    const next = DRAWER_TAB === 'diff' ? 'full' : 'diff';
    switchTab(next);
    $('#dtab-' + next).focus();
  });
  body.querySelectorAll('.sec-more').forEach((b) => b.addEventListener('click', () => {
    const list = b.previousElementSibling;
    const open = !list.classList.contains('open');
    list.classList.toggle('open', open);
    b.setAttribute('aria-expanded', open ? 'true' : 'false');
    b.textContent = open ? t('sec.less') : t('sec.more', { n: b.dataset.n });
  }));
  const cf = $('#copyFin');
  if (cf) cf.addEventListener('click', () => {
    const done = () => { cf.classList.add('done'); cf.title = t('copy.done'); setTimeout(() => { cf.classList.remove('done'); cf.title = t('copy.fin'); }, 1400); };
    if (navigator.clipboard) navigator.clipboard.writeText(cf.dataset.v).then(done).catch(() => {});
  });
  const qi = $('#codeQ');
  qi.value = q;
  qi.addEventListener('input', filterDrawerCodes);
  switchTab(DRAWER_TAB);

  drawer.classList.remove('hidden');
  $('#backdrop').classList.remove('hidden');
  if (!wasOpen) {
    drawer.scrollTop = 0;
    $('#drawerClose').focus();
  }
}

function switchTab(name) {
  DRAWER_TAB = name;
  $('#drawerBody').querySelectorAll('.dtab').forEach((tb) => {
    const on = tb.dataset.tab === name;
    tb.classList.toggle('active', on);
    tb.setAttribute('aria-selected', on ? 'true' : 'false');
    tb.tabIndex = on ? 0 : -1;
  });
  $('#drawerBody').querySelectorAll('.dpane').forEach((p) =>
    p.classList.toggle('hidden', p.dataset.pane !== name));
  filterDrawerCodes();
}

// 상세 창 코드 검색 — 보이는 탭의 줄(코드·설명)만 거른다. 섹션 숫자는 그대로 둔다.
function filterDrawerCodes() {
  const qi = $('#codeQ');
  const q = qi ? qi.value.trim().toLowerCase() : '';
  const pane = $('#drawerBody .dpane:not(.hidden)');
  if (!pane) return;
  let hits = 0, total = 0;
  pane.querySelectorAll('[data-q]').forEach((el) => {
    total++;
    const hit = !q || el.dataset.q.includes(q);
    el.classList.toggle('q-hide', !hit);
    if (hit) hits++;
  });
  pane.classList.toggle('searching', !!q);
  const none = pane.querySelector('.q-none');
  if (none) none.classList.toggle('hidden', !q || hits > 0 || !total);
}

function closeDrawer() {
  const drawer = $('#drawer');
  if (drawer.classList.contains('hidden')) return;
  drawer.classList.add('hidden');
  $('#backdrop').classList.add('hidden');
  if (DRAWER_RETURN && document.contains(DRAWER_RETURN)) DRAWER_RETURN.focus();
  DRAWER_RETURN = null;
}

// 상세 창이 열려 있는 동안 Tab 포커스를 창 안에 가둔다.
function trapDrawerFocus(e) {
  const drawer = $('#drawer');
  if (e.key !== 'Tab' || drawer.classList.contains('hidden')) return;
  const f = [...drawer.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])')]
    .filter((el) => !el.disabled && el.offsetParent !== null);
  if (!f.length) return;
  const first = f[0], last = f[f.length - 1];
  if (!drawer.contains(document.activeElement)) { e.preventDefault(); first.focus(); }
  else if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
  else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
}

const _today = () => new Date().toISOString().slice(0, 10);
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
  // cell.n = 숫자 셀(엑셀에서 정렬·집계 가능). 그 외는 문자열로 넣는다.
  if (cell.n && typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
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

function exportRowXls(r) {
  if (!r) return;
  const metaKeys = ['Commission no.', 'Model(WINGS)', 'Vehicle', 'Category', 'Type', 'Cab', 'MY', 'PTO',
    'Production date', 'Changeability Date', 'Until Dealine',
    'SAM Baumuster', 'SAM now', 'SAM Status', 'SAM Update', 'Compared SAM file name'];
  const rows = [];
  const merges = [];
  rows.push([{ v: `${r['Commission no.']}  ·  ${r['Model(WINGS)'] || ''}`, s: 1 }]);
  rows.push([{ v: r['Compared SAM file name'] || '', s: 0 }]);
  rows.push([]);
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

// ====================== 대시보드 Export (현재 필터/정렬 결과 → xlsx) ======================
// 화면에 보이는 목록을 그대로 내보낸다. 코드 컬럼은 개수와 실제 코드 목록을 함께 담는다.
const EXPORT_COLS = [
  { label: 'Commission',          w: 16, get: (r) => r['Commission no.'] },
  { label: 'Model',               w: 14, get: (r) => r['Vehicle'] },
  { label: 'Type',                w: 18, get: (r) => r['Model(WINGS)'] },
  { label: 'Axle',                w: 10, get: (r) => r['Type'] },
  { label: 'Cab',                 w: 12, get: (r) => r['Cab'] },
  { label: 'MY',                  w: 8,  get: (r) => r['MY'] },
  { label: 'PTO',                 w: 10, get: (r) => r['PTO'] },
  { label: 'Production',          w: 14, get: (r) => r['Production date'] },
  { label: 'Changeability',       w: 14, get: (r) => r['Changeability Date'] },
  { label: 'Changeability D-Day', w: 18, get: (r) => ddayText(r['Until Dealine']) },
  { label: 'Status',              w: 14, get: (r) => r['SAM Status'] },
  { label: 'SAM Update',          w: 12, get: (r) => (r['SAM Update'] ? 'Y' : '') },
  { label: 'Mismatch exception',  w: 22, get: (r) => (r._accepted ? 'Approved (was Mismatch)'
    : (r._accN ? `${r._accN} code(s) approved` : '')) },
  { label: 'Only in SAM (n)',     w: 14, num: true, get: (r) => diffCount(r, 'sam') },
  { label: 'Only in SAM',         w: 40, get: (r) => diffCodes(r, 'sam').join(',') },
  { label: 'Only in WINGS (n)',   w: 16, num: true, get: (r) => diffCount(r, 'win') },
  { label: 'Only in WINGS',       w: 40, get: (r) => diffCodes(r, 'win').join(',') },
  { label: 'Mandatory (n)',       w: 14, num: true, get: (r) => countOf(r['Mandatory Codes']) },
  { label: 'Mandatory Codes',     w: 40, get: (r) => r['Mandatory Codes'] },
  { label: 'SAM Baumuster',       w: 16, get: (r) => r['SAM Baumuster'] },
  { label: 'Compared SAM file',   w: 34, get: (r) => r['Compared SAM file name'] },
];

// 지금 적용된 필터를 사람이 읽을 수 있는 한 줄로. (시트 상단에 적어둔다)
function exportFilterSummary() {
  const parts = [];
  for (const f of FILTER_FIELDS) {
    const v = $(f.id).value;
    if (v) parts.push(`${t('export.f.' + f.key)}=${v}`);
  }
  const st = $('#statusFilter').value;
  if (st) parts.push(`${t('export.f.status')}=${st}`);
  const q = $('#search').value.trim();
  if (q) parts.push(`${t('export.f.search')}="${q}"`);
  if ($('#upcomingOnly').checked) parts.push(`${t('chk.upcoming')}: ${DASH_FROM}~`);
  if (restrictSoon) parts.push(t('dash.soon'));
  if (tileMandatory) parts.push(t('tile.mand'));
  if (tileSamUpdate) parts.push(t('tile.samupd'));
  if (sortKey) parts.push(`${t('export.f.sort')}: ${sortKey} ${sortDir === 1 ? '▲' : '▼'}`);
  return parts.length ? parts.join('  ·  ') : t('export.noFilter');
}

function exportFilteredXlsx() {
  const rows = filtered();
  if (!rows.length) { alert(t('export.noRows')); return; }

  const d = new Date();
  const p2 = (n) => String(n).padStart(2, '0');
  const stamp = `${d.getFullYear()}${p2(d.getMonth() + 1)}${p2(d.getDate())}_${p2(d.getHours())}${p2(d.getMinutes())}`;
  const ts = `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}`;
  const lastCol = _colName(EXPORT_COLS.length - 1);

  const sheet = [];
  const merges = [];
  sheet.push([{ v: t('export.sheetTitle', { n: rows.length, total: DATA.rows.length }), s: 1 }]);
  sheet.push([{ v: t('export.filters', { f: exportFilterSummary() }), s: 0 }]);
  sheet.push([{ v: t('export.exportedAt', { ts }), s: 0 }]);
  sheet.push([]);
  for (let i = 1; i <= 3; i++) merges.push(`A${i}:${lastCol}${i}`);

  sheet.push(EXPORT_COLS.map((c) => ({ v: c.label, s: 2 })));
  for (const r of rows) {
    sheet.push(EXPORT_COLS.map((c) => {
      const v = c.get(r);
      if (c.num) return { v: Number(v) || 0, n: true, s: 3 };
      return { v: v == null ? '' : String(v), s: 3 };
    }));
  }

  const cols = EXPORT_COLS.map((c, i) => ({ min: i + 1, max: i + 1, w: c.w }));
  writeXlsx(`afab_sam_export_${stamp}.xlsx`, 'Filtered', sheet, merges, cols);
}

// ====================== 생산월 이력관리 (같은 모델의 코드 변경 히스토리) ======================
// 같은 MY · 같은 모델(Model · Type · Axle · Cab)의 commission 을 생산월 순으로 늘어놓고,
// 달마다 어떤 코드가 추가(＋)/삭제(−)됐는지 보여준다. 각 모델의 첫 생산월이 '기본'이 되고,
// 그 뒤 달은 직전 생산월(기본) 또는 첫 생산월과 비교한다 — 비교 기준은 화면에서 전환.
const HIST_NO_MY = '-';
const HIST = { my: '', src: 'wings', basis: 'prev', q: '', rows: [], last: null,
  cat: '', onlyChanged: false, pop: null, fallbackWarned: new Set() };

// 모델을 이루는 항목은 대시보드 목록과 똑같이 Model · Type · Axle · Cab · MY · PTO.
// MY 는 화면 위 드롭다운으로 이미 고른 값이라 모델 키에서는 빼고, 표에만 함께 보여준다.
const HIST_COLS = ['Vehicle', 'Model(WINGS)', 'Type', 'Cab', 'MY', 'PTO'];
function colLabel(key) {
  const c = COLS.find((x) => x.key === key);
  return c ? c.label : key;
}
function histModelKey(r) {
  return HIST_COLS.filter((k) => k !== 'MY')
    .map((k) => String(r[k] ?? '').trim()).filter(Boolean).join(' ');
}
function histMyOf(r) { return String(r.MY || '') || HIST_NO_MY; }
function histMyLabel(my) {
  return my === HIST_NO_MY ? t('hist.noMy') : 'MY' + String(my).slice(-2);
}
const HIST_SRCS = ['wings', 'sam'];
// 도색·타이어(CTT)는 compare 가 3자리 옵션코드와 섞지 않고 따로 남긴다(_paint_*, _tyre_*).
// 이력관리도 같은 갈래로 나눠 줄을 따로 세운다 — 성격이 다른 코드가 한 줄에 섞이지 않게.
const HIST_KINDS = [
  { key: 'code', field: { wings: '_all_wings_codes', sam: '_all_sam_codes' } },
  { key: 'paint', field: { wings: '_paint_wings', sam: '_paint_sam' } },
  { key: 'tyre', field: { wings: '_tyre_wings', sam: '_tyre_sam' } },
];
const HIST_KIND_ICON = { code: '', paint: '🎨', tyre: '🛞' };
function histCodes(r, src, kind) {
  const k = HIST_KINDS.find((x) => x.key === (kind || 'code'));
  return splitCodes(r[k.field[(src || HIST.src) === 'sam' ? 'sam' : 'wings']]);
}
function histBasisLabel() { return t('hist.basis.' + HIST.basis); }
function histSrcLabel() { return t('hist.src.' + HIST.src); }

// MY 드롭다운 — 데이터에 있는 MY 만(최근 순). 현재 선택은 가능한 한 유지한다.
function histFillMy() {
  const sel = $('#histMy');
  if (!sel) return;
  const set = new Set();
  for (const r of DATA.rows) if (prodMonth(r) && histModelKey(r)) set.add(histMyOf(r));
  const list = [...set].filter((v) => v !== HIST_NO_MY).sort().reverse();
  if (set.has(HIST_NO_MY)) list.push(HIST_NO_MY);
  if (!list.includes(HIST.my)) HIST.my = list[0] || '';
  sel.innerHTML = list.map((v) =>
    `<option value="${esc(v)}">${esc(histMyLabel(v))}</option>`).join('');
  sel.value = HIST.my;
  enhanceSelect(sel);
  // 언어를 바꾸면 코드 기준 드롭다운의 라벨도 다시 그려야 한다.
  const src = $('#histSrc');
  if (src && src._csel) src._csel.refresh();
}

// 같은 달 안에서 갈린 코드 목록 — 너무 길면 줄인다.
function histVariedText(varied) {
  const head = varied.slice(0, 8).join(', ');
  return t('hist.varied',
    { n: varied.length, codes: varied.length > 8 ? head + ' …' : head });
}

// 화면에 그릴 구조를 만든다.
//   { months: [YYYY-MM…], models: [{ key, months, cells }], total }
// 한 달에 commission 이 여러 대면 그 달의 코드 집합은 합집합으로 본다. 합집합에는 있지만
// 그 달 모든 commission 에 다 들어있지는 않은 코드는 '같은 달 안에서 사양이 갈린' 것이라
// 따로(varied) 표시해 준다 — 합집합 때문에 없는 변경이 생긴 것처럼 보이지 않게.
function histBuild() {
  const q = HIST.q.trim().toLowerCase();
  const models = new Map();
  const monthSet = new Set();
  HIST.rows = [];
  for (const r of DATA.rows) {
    const month = prodMonth(r);
    const key = histModelKey(r);
    if (!month || !key || histMyOf(r) !== HIST.my) continue;
    if (q && !key.toLowerCase().includes(q)
      && !String(r['Commission no.'] ?? '').toLowerCase().includes(q)) continue;
    monthSet.add(month);
    if (!models.has(key)) models.set(key, new Map());
    const byMonth = models.get(key);
    if (!byMonth.has(month)) byMonth.set(month, []);
    byMonth.get(month).push(r);
  }

  // 두 소스 × 세 갈래(코드·도색·타이어)를 늘 함께 계산해 둔다 — 화면이 줄로 나눠 쓴다.
  const aggregate = (rs) => {
    const per = {};
    for (const src of HIST_SRCS) {
      per[src] = {};
      for (const kind of HIST_KINDS) {
        const sets = rs.map((r) => new Set(histCodes(r, src, kind.key)));
        const codes = new Set();
        sets.forEach((st) => st.forEach((c) => codes.add(c)));
        per[src][kind.key] = {
          codes,
          varied: [...codes].filter((c) => !sets.every((st) => st.has(c))).sort(),
        };
      }
    }
    return per;
  };

  // 이전 연식(지금 MY 보다 앞선 가장 가까운 MY)의 같은 모델 — 그 연식의 마지막 생산월.
  // 연식만 바뀌어 넘어온 모델의 기준월을 이 달과 비교한다. 검색어와 무관하게 찾는다.
  const prevMy = new Map();
  if (HIST.my !== HIST_NO_MY && models.size) {
    const found = new Map(); // key → { my, month, rows }
    for (const r of DATA.rows) {
      const my = histMyOf(r);
      if (my === HIST_NO_MY || !(my < HIST.my)) continue;
      const key = histModelKey(r);
      const month = prodMonth(r);
      if (!month || !models.has(key)) continue;
      const f = found.get(key);
      if (!f || my > f.my || (my === f.my && month > f.month)) found.set(key, { my, month, rows: [r] });
      else if (my === f.my && month === f.month) f.rows.push(r);
    }
    for (const [key, f] of found) {
      prevMy.set(key, { my: f.my, month: f.month, cell: aggregate(f.rows) });
    }
  }

  const out = [];
  let total = 0;
  for (const [key, byMonth] of models) {
    const seq = [...byMonth.keys()].sort();
    const cells = new Map();
    for (const m of seq) {
      const rs = byMonth.get(m);
      total += rs.length;
      const per = aggregate(rs);
      // 그 달의 WINGS↔SAM 차이는 compare 가 이미 정리해 둔 Only_in_* 를 쓴다
      // (Factory Control·필수코드는 빠져 있어 대시보드의 미스매치 기준과 같다).
      const onlyW = new Set(), onlyS = new Set();
      for (const r of rs) {
        splitCodes(r['Only_in_WINGS']).forEach((c) => onlyW.add(c));
        splitCodes(r['Only_in_SAM']).forEach((c) => onlyS.add(c));
      }
      // 대표(샘플) commission 이 매번 같도록 번호순으로 정렬해 둔다.
      rs.sort((a, b) => String(a['Commission no.']).localeCompare(String(b['Commission no.'])));
      cells.set(m, {
        rows: rs, wings: per.wings, sam: per.sam,
        onlyW: [...onlyW].sort(), onlyS: [...onlyS].sort(),
      });
    }
    const diff = (cell, refCell) => {
      for (const src of HIST_SRCS) {
        for (const kind of HIST_KINDS) {
          const ref = refCell[src][kind.key].codes;
          const cur = cell[src][kind.key];
          cur.added = [...cur.codes].filter((c) => !ref.has(c)).sort();
          cur.removed = [...ref].filter((c) => !cur.codes.has(c)).sort();
        }
      }
    };
    seq.forEach((m, i) => {
      const cell = cells.get(m);
      if (i === 0) {
        // 기준월은 '기준'으로 두되, 같은 모델이 이전 연식에 있었다면 그 연식의
        // 마지막 생산월과 비교한 변경(연식 변경분)을 함께 보여 준다.
        cell.base = true;
        const prev = prevMy.get(key);
        if (prev) {
          cell.prevMy = { my: prev.my, month: prev.month };
          diff(cell, prev.cell);
        } else {
          for (const src of HIST_SRCS) {
            for (const kind of HIST_KINDS) { cell[src][kind.key].added = []; cell[src][kind.key].removed = []; }
          }
        }
        return;
      }
      const refMonth = HIST.basis === 'first' ? seq[0] : seq[i - 1];
      cell.ref = refMonth;
      diff(cell, cells.get(refMonth));
    });
    // 모델 표(대시보드와 같은 항목)를 그릴 때 쓰는 대표 행 — 같은 키면 항목 값이 같다.
    out.push({ key, months: seq, cells, sample: cells.get(seq[0]).rows[0] });
  }
  out.sort((a, b) => a.key.localeCompare(b.key));
  return { months: [...monthSet].sort(), models: out, total };
}

// 그 달의 WINGS↔SAM 대조 — 배지는 바로 아래 두 줄(WINGS·SAM)의 변경 내역이 같은지를 말한다.
// 화면에 보이는 것과 배지가 늘 맞아떨어지게 하려는 것. 다만 '변경 내역이 같다'와 '그 달 코드가
// 같다'는 다른 이야기라(기준이 되는 달이 이미 갈려 있으면 변화만 같을 수 있다) 그 달 전체 코드의
// 차이(Only_in_*, 대시보드 미스매치와 같은 기준)는 툴팁 둘째 줄에 함께 적어 둔다.
function histXOf(cell, kinds) {
  const ks = kinds || HIST_KINDS.map((k) => k.key);
  const sig = (src) => ks.reduce((a, k) => {
    const d = cell[src][k], p = k === 'code' ? '' : HIST_KIND_ICON[k] + ' ';
    return a.concat(d.added.map((c) => '＋' + p + c), d.removed.map((c) => '−' + p + c));
  }, []);
  let same, first;
  if (cell.base && !cell.prevMy) {
    same = true;
    first = t('hist.xBase');
  } else {
    const a = sig('wings'), b = sig('sam');
    const onlyW = a.filter((x) => !b.includes(x));
    const onlyS = b.filter((x) => !a.includes(x));
    same = !onlyW.length && !onlyS.length;
    first = same ? t('hist.xDeltaSame')
      : t('hist.xDelta', { w: onlyW.join(', ') || '—', s: onlyS.join(', ') || '—' });
  }
  // 옵션코드는 compare 가 정리해 둔 Only_in_*(Factory Control·필수코드 제외 — 대시보드
  // 미스매치와 같은 기준), 도색·타이어는 그 달 값끼리 직접 비교한다.
  const absW = [...cell.onlyW], absS = [...cell.onlyS];
  for (const k of ks) {
    if (k === 'code') continue;
    const w = cell.wings[k].codes, s = cell.sam[k].codes;
    [...w].filter((c) => !s.has(c)).sort().forEach((c) => absW.push(HIST_KIND_ICON[k] + ' ' + c));
    [...s].filter((c) => !w.has(c)).sort().forEach((c) => absS.push(HIST_KIND_ICON[k] + ' ' + c));
  }
  const second = (absW.length || absS.length)
    ? t('hist.xTitle', { w: absW.join(', ') || '—', s: absS.join(', ') || '—' })
    : t('hist.xNone');
  return { same, tip: first + '\n' + second };
}

// 그 달 아래 붙는 WINGS↔SAM 일치/불일치 띠 — 월 머리글 바로 다음 줄.
function histXHeadHtml(cell, kinds) {
  if (!cell) return `<th class="hist-xcell off"></th>`;
  const x = histXOf(cell, kinds);
  return `<th class="hist-xcell"><span class="hist-x ${x.same ? 'ok' : 'bad'}" ` +
    `title="${esc(x.tip)}">${esc(t(x.same ? 'hist.xOk' : 'hist.xBad'))}</span></th>`;
}

// ---- 화면 표시 전용 집계 — histBuild() 가 만든 added/removed 를 세기만 한다 ----
// 지금 보고 있는 소스(WINGS / SAM / 대조)만 센다 — 왼쪽 목록의 '변경 n개월'과 같은 기준.
function histShownSrcs() { return HIST_SRCS.filter((src) => HIST.src === 'both' || HIST.src === src); }
function histCellDelta(cell) {
  const out = { add: 0, del: 0 };
  if (!cell || (cell.base && !cell.prevMy)) return out;
  for (const src of histShownSrcs()) {
    for (const k of HIST_KINDS) {
      out.add += cell[src][k.key].added.length;
      out.del += cell[src][k.key].removed.length;
    }
  }
  return out;
}
function histChangedMonths(g) {
  return g.months.filter((m) => { const d = histCellDelta(g.cells.get(m)); return d.add + d.del > 0; }).length;
}
function histUnits(g) { return g.months.reduce((s, m) => s + g.cells.get(m).rows.length, 0); }

// 차종 구분(Tractor / Rigid / Tipper) — 상세 창의 Category 와 같은 필드.
// 값이 없을 때만 Baumuster 표기 끝으로 가른다: LS·S → Tractor, K → Tipper, 나머지 Rigid.
const HIST_CATS = ['tractor', 'rigid', 'tipper'];
function histCategoryOf(g) {
  const c = String((g.sample && g.sample.Category) || '').trim().toLowerCase();
  if (HIST_CATS.includes(c)) return c;
  const bm = String((g.sample && g.sample['Model(WINGS)']) || '').trim().toUpperCase();
  const cat = /\s(LS|S)$/.test(bm) ? 'tractor' : (/\sK$/.test(bm) ? 'tipper' : 'rigid');
  if (!HIST.fallbackWarned.has(g.key)) {
    HIST.fallbackWarned.add(g.key);
    console.info('[history] Category 없음 — Baumuster 규칙으로 분류:', g.key, '→', cat);
  }
  return cat;
}

// 모델 목록(왼쪽) — 지금 탭(카테고리)의 모델만. 대수 · 생산월 수 · 변경 월 배지.
function histModelListHtml(list) {
  if (!list.length) return `<div class="hist-empty">${esc(t('hist.catEmpty'))}</div>`;
  return list.map((g) => {
    const chg = histChangedMonths(g);
    const on = g.key === HIST.model;
    const badge = chg
      ? `<span class="hm-badge chg">${esc(tn('hist.changedN', chg, { n: chg }))}</span>`
      : `<span class="hm-badge">${esc(t('hist.noChange'))}</span>`;
    return `<button type="button" class="hm-item${on ? ' sel' : ''}" aria-pressed="${on}" ` +
      `data-key="${esc(g.key)}" title="${esc(g.key)}">` +
      `<span class="hm-line"><span class="hm-name">${esc(g.key)}</span>${badge}</span>` +
      `<span class="hm-sub">${esc(t('hist.modelSub', {
        n: tn('hist.unitsOf', histUnits(g), { n: histUnits(g) }),
        m: tn('hist.monthsOf', g.months.length, { n: g.months.length }) }))}</span></button>`;
  }).join('');
}

// 이 모델이 실제로 값을 가진 갈래만 줄로 세운다 — 도색·타이어가 없는 모델에 빈 줄을 만들지 않게.
function histKindsOf(g) {
  return HIST_KINDS.filter((k) => k.key === 'code' || g.months.some((m) =>
    HIST_SRCS.some((src) => g.cells.get(m)[src][k.key].codes.size)));
}

function histGapCellHtml() {
  return `<td class="hist-cell gap" title="${esc(t('hist.gapTitle'))}">` +
    `<span class="hist-none">—</span></td>`;
}
function histColCls(cell) {
  if (!cell) return '';
  if (cell.base) return ' base';
  const d = histCellDelta(cell);
  return d.add + d.del ? ' chg' : '';
}

// 생산 대수 줄 — 그 달 대수 알약. 누르면 그 달 Commission 목록이 뜬다.
function histUnitsCellHtml(cell, month) {
  if (!cell) return histGapCellHtml();
  const n = cell.rows.length;
  const first = cell.rows[0]['Commission no.'];
  const tip = n > 1 ? t('hist.unitsTitle', { first, n: n - 1 }) : String(first);
  return `<td class="hist-cell units${histColCls(cell)}"><button type="button" class="units-pill" ` +
    `data-m="${esc(month)}" aria-haspopup="true" title="${esc(tip)}">${esc(t('hist.unitsN', { n }))}</button></td>`;
}

// 한 소스(WINGS 또는 SAM)의 그 달 칸 — 갈래(코드·도색·타이어)별로 한 칸씩.
// 기준월은 '기준', 변경은 ＋추가 먼저 −삭제 나중, 변경 없음은 옅은 —.
function histSrcCellHtml(cell, month, src, kindKey) {
  if (!cell) return histGapCellHtml();
  const d = cell[src][kindKey];
  const desc = (c) => kindKey === 'paint' ? 'MB ' + c
    : (kindKey === 'tyre' ? c : (describe(c) || c));
  const line = (c, kind) =>
    `<span class="hcode ${kind}" title="${esc(desc(c))}">` +
    `${kind === 'add' ? '+' : '−'}${esc(c)}</span>`;
  const deltaHtml = () => {
    const lines = d.added.map((c) => line(c, 'add'))
      .concat(d.removed.map((c) => line(c, 'del')));
    return lines.length
      ? `<div class="hist-codes">${lines.join('')}</div>`
      : `<span class="hist-dash" role="img" aria-label="${esc(t('hist.same'))}">—</span>`;
  };
  let body, title;
  if (cell.base) {
    title = t('hist.baseTitle', { m: month });
    body = `<div class="hist-base">${esc(t('hist.base'))}</div>`;
    if (cell.prevMy) {
      const ref = histMyLabel(cell.prevMy.my) + ' ' + cell.prevMy.month;
      title += '\n' + t('hist.cellTitle', { ref, cur: month });
      body += `<div class="hist-prevmy">${esc(t('hist.vsPrevMy', { ref }))}</div>` + deltaHtml();
    }
  } else {
    title = t('hist.cellTitle', { ref: cell.ref, cur: month });
    body = deltaHtml();
  }
  // 같은 달 commission 끼리 갈린 코드
  const warn = d.varied.length
    ? `<div class="hist-warn" title="${esc(histVariedText(d.varied))}">⚠ ${d.varied.length}</div>`
    : '';
  return `<td class="hist-cell${histColCls(cell)}" title="${esc(title)}">${body}${warn}</td>`;
}

// 오른쪽 위 요약 줄 — 모델 이름·사양은 왼쪽에 이미 있으니 숫자만.
function histSummaryHtml(g) {
  const tot = g.months.reduce((a, m) => {
    const d = histCellDelta(g.cells.get(m));
    return { add: a.add + d.add, del: a.del + d.del };
  }, { add: 0, del: 0 });
  const chg = histChangedMonths(g);
  const stat = (label, value, cls) =>
    `<div class="hs-stat${cls ? ' ' + cls : ''}"><span class="hs-l">${esc(label)}</span><span class="hs-v">${value}</span></div>`;
  const note = HIST.basis === 'first'
    ? t('hist.note.first', { m: g.months[0] }) : t('hist.note.prev');
  return `<div class="hist-sum">` +
    stat(t('hist.st.units'), esc(t('hist.unitsN', { n: fmtNum(histUnits(g)) }))) +
    stat(t('hist.st.months'), esc(t('hist.st.monthsV', { n: g.months.length }))) +
    stat(t('hist.st.chg'), esc(t('hist.st.monthsV', { n: chg })), chg ? 'warn' : '') +
    stat(t('hist.st.items'), `<span class="hs-add">+${tot.add}</span> <span class="hs-sep">/</span> <span class="hs-del">−${tot.del}</span>`) +
    `<span class="hs-note">${esc(histSrcLabel())} · ${esc(note)}</span></div>`;
}

// 선택한 모델 1종의 월별 히스토리 표(가로 스크롤).
// 머리글은 연도(한 번만) → 월 번호 + 기준/변경 표시. 두 소스를 함께 볼 때는 그 아래에
// 그 달의 WINGS↔SAM 일치/불일치를 한 줄로 깔아 준다.
function histDetailHtml(g, months) {
  const both = HIST.src === 'both';
  const srcs = both ? HIST_SRCS : [HIST.src === 'sam' ? 'sam' : 'wings'];
  const kinds = histKindsOf(g);
  const years = [];
  for (const m of months) {
    const y = m.slice(0, 4);
    if (years.length && years[years.length - 1].y === y) years[years.length - 1].n++;
    else years.push({ y, n: 1 });
  }
  let head = `<tr class="hist-yrow"><td class="hist-rowhead corner" rowspan="2"></td>` +
    years.map((y) => `<th scope="colgroup" colspan="${y.n}">${esc(y.y)}</th>`).join('') + `</tr>`;
  head += `<tr class="hist-mrow">` + months.map((m) => {
    const cell = g.cells.get(m);
    const d = histCellDelta(cell);
    const tag = cell && cell.base
      ? `<span class="hm-tag base">${esc(t('hist.base'))}` +
        (d.add + d.del ? ` <span class="chg" title="${esc(t('hist.monthChgPrevMy', { my: histMyLabel(cell.prevMy.my), n: d.add + d.del }))}">● ${d.add + d.del}</span>` : '') +
        `</span>`
      : (d.add + d.del ? `<span class="hm-tag chg" title="${esc(t('hist.monthChg', { n: d.add + d.del }))}">● ${d.add + d.del}</span>`
        : `<span class="hm-tag"></span>`);
    return `<th scope="col" class="${histColCls(cell).trim()}" title="${esc(m)}">` +
      `<span class="hm-num">${esc(m.slice(5, 7))}</span>${tag}</th>`;
  }).join('') + `</tr>`;
  if (both) {
    const ks = kinds.map((k) => k.key);
    head += `<tr class="hist-xrow"><td class="hist-rowhead corner"></td>` +
      months.map((m) => histXHeadHtml(g.cells.get(m), ks)).join('') + `</tr>`;
  }
  let body = `<tr class="r-units"><th scope="row" class="hist-rowhead"><span class="hkind">${esc(t('hist.units'))}</span></th>` +
    months.map((m) => histUnitsCellHtml(g.cells.get(m), m)).join('') + `</tr>`;
  // 갈래를 먼저 묶고 그 안에서 WINGS·SAM 을 위아래로 — 같은 성격의 코드끼리 맞대어 보게.
  body += kinds.map((k) => srcs.map((src, si) =>
    `<tr class="${si === 0 ? 'ktop' : ''}">` +
    `<th scope="row" class="hist-rowhead"><span class="hkind">${esc(t('hist.kind.' + k.key))}</span>` +
    `<span class="hsub">${src === 'sam' ? 'SAM' : 'WINGS'}</span></th>` +
    months.map((m) => histSrcCellHtml(g.cells.get(m), m, src, k.key)).join('') + `</tr>`
  ).join('')).join('');
  const srcName = t('hist.srcName.' + (HIST.src === 'sam' ? 'sam' : (both ? 'both' : 'wings')));
  return histSummaryHtml(g) +
    `<div class="hist-scroll"><table class="hist-grid">` +
    `<thead>${head}</thead><tbody>${body}</tbody></table></div>` +
    `<div class="hist-foot"><span>${esc(t('hist.foot'))}</span><span>${esc(t('hist.footSrc', { src: srcName }))}</span></div>`;
}

// 대수 알약을 누르면 뜨는 그 달 Commission 목록 — 번호를 누르면 상세 창.
function histOpenPopover(btn, g) {
  histClosePopover();
  const cell = g.cells.get(btn.dataset.m);
  if (!cell) return;
  const pop = document.createElement('div');
  pop.className = 'hist-pop';
  pop.setAttribute('role', 'dialog');
  pop.setAttribute('aria-label', t('hist.popTitle', { m: btn.dataset.m, n: cell.rows.length }));
  pop.innerHTML = `<div class="hp-head">${esc(t('hist.popTitle', { m: btn.dataset.m, n: cell.rows.length }))}</div>` +
    `<div class="hp-list">` + cell.rows.map((r, i) =>
      `<button type="button" class="hp-item" data-i="${i}"><span class="mono">${esc(r['Commission no.'])}</span>` +
      `${statusBadge(String(r['SAM Status'] || ''))}</button>`).join('') + `</div>`;
  document.body.appendChild(pop);
  const rc = btn.getBoundingClientRect();
  const w = pop.offsetWidth, h = pop.offsetHeight;
  pop.style.left = Math.max(8, Math.min(rc.left, window.innerWidth - w - 8)) + 'px';
  pop.style.top = (rc.bottom + h + 8 > window.innerHeight ? Math.max(8, rc.top - h - 6) : rc.bottom + 6) + 'px';
  pop.querySelectorAll('.hp-item').forEach((b) => b.addEventListener('click', () => {
    histClosePopover();
    openDrawer(cell.rows[Number(b.dataset.i)]);
  }));
  HIST.pop = { el: pop, btn };
  btn.setAttribute('aria-expanded', 'true');
  const firstItem = pop.querySelector('.hp-item');
  if (firstItem) firstItem.focus();
}
function histClosePopover(restore) {
  if (!HIST.pop) return;
  HIST.pop.el.remove();
  HIST.pop.btn.setAttribute('aria-expanded', 'false');
  if (restore && document.contains(HIST.pop.btn)) HIST.pop.btn.focus();
  HIST.pop = null;
}
document.addEventListener('click', (e) => {
  if (HIST.pop && !HIST.pop.el.contains(e.target) && e.target !== HIST.pop.btn) histClosePopover();
});
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') histClosePopover(true); });

function histRender() {
  histClosePopover();
  const el = $('#histBody');
  const listEl = $('#histModels');
  const catEl = $('#histCats');
  if (!el || !listEl) return;
  const countEl = $('#histCount');
  const empty = (msg) => {
    listEl.innerHTML = '';
    if (catEl) catEl.innerHTML = '';
    el.innerHTML = `<div class="hist-empty">${esc(msg)}</div>`;
  };
  if (!DATA.rows.length) {
    HIST.last = null;
    if (countEl) countEl.textContent = '';
    empty(t('msg.noData'));
    return;
  }
  const b = histBuild();
  HIST.last = b;
  if (countEl) {
    countEl.textContent = t('hist.count',
      { models: b.models.length, months: b.months.length, n: fmtNum(b.total) });
  }
  if (!b.models.length) { empty(t('hist.emptyMy')); return; }

  // 카테고리 탭 — 모델은 정확히 한 탭에만 들어간다. 지금 탭이 비면(MY·검색이 바뀌어서)
  // 모델이 있는 첫 탭으로 옮긴다 — 검색어가 다른 카테고리에만 걸려도 그쪽으로 넘어간다.
  const byCat = {};
  HIST_CATS.forEach((c) => { byCat[c] = []; });
  b.models.forEach((g) => byCat[histCategoryOf(g)].push(g));
  if (!byCat[HIST.cat] || !byCat[HIST.cat].length) HIST.cat = HIST_CATS.find((c) => byCat[c].length);
  if (catEl) {
    catEl.innerHTML = HIST_CATS.map((c) => {
      const on = c === HIST.cat;
      return `<button type="button" role="tab" class="hm-tab${on ? ' active' : ''}" data-cat="${c}" ` +
        `aria-selected="${on}" tabindex="${on ? 0 : -1}"${byCat[c].length ? '' : ' disabled'}>` +
        `${esc(t('hist.cat.' + c))} <span class="cnt">${byCat[c].length}</span></button>`;
    }).join('');
    catEl.querySelectorAll('.hm-tab').forEach((tb) => tb.addEventListener('click', () => {
      HIST.cat = tb.dataset.cat;
      histRender();
      const a = $('#histCats .hm-tab.active');
      if (a) a.focus();
    }));
  }

  let list = byCat[HIST.cat];
  if (HIST.onlyChanged) list = list.filter((g) => histChangedMonths(g) > 0);
  listEl.innerHTML = histModelListHtml(list);
  if (!list.length) {
    el.innerHTML = `<div class="hist-empty">${esc(t('hist.catEmpty'))}</div>`;
    return;
  }
  // 선택이 없거나 지금 목록에 없으면 볼거리가 가장 많은 모델(생산월이 긴 쪽)을 고른다.
  if (!list.some((g) => g.key === HIST.model)) {
    HIST.model = [...list].sort((x, y) => y.months.length - x.months.length)[0].key;
    listEl.innerHTML = histModelListHtml(list);
  }
  listEl.querySelectorAll('.hm-item').forEach((btn) =>
    btn.addEventListener('click', () => {
      HIST.model = btn.dataset.key;
      histRender();
      const sel = $('#histModels').querySelector('.hm-item.sel');
      if (sel) { sel.scrollIntoView({ block: 'nearest' }); sel.focus(); }
    }));

  const g = list.find((x) => x.key === HIST.model);
  // 그 모델이 실제로 생산된 달만 열로 세운다 — 빈 달을 지우면 한 화면에 더 많은 달이 들어온다.
  el.innerHTML = histDetailHtml(g, g.months);
  // 머리글 줄(연도 · 월 · 일치/불일치)이 위에 붙어 있게 — 높이는 글꼴·언어에 따라 달라져
  // CSS 로 고정할 수 없으니 그린 뒤 재서 넣는다.
  const grid = el.querySelector('.hist-grid');
  const yr = grid.querySelector('.hist-yrow');
  const mr = grid.querySelector('.hist-mrow');
  if (yr) grid.style.setProperty('--h1', yr.offsetHeight + 'px');
  if (yr && mr) grid.style.setProperty('--h2', (yr.offsetHeight + mr.offsetHeight) + 'px');
  el.querySelectorAll('.units-pill').forEach((btn) =>
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      if (HIST.pop && HIST.pop.btn === btn) { histClosePopover(); return; }
      histOpenPopover(btn, g);
    }));
}

// 지금 보이는 히스토리를 그대로 엑셀로. 한 셀에 commission + 변경 코드 줄바꿈으로 담는다.
function histExportXlsx() {
  const b = HIST.last;
  if (!b || !b.models.length) { alert(t('hist.emptyMy')); return; }
  const d = new Date();
  const p2 = (n) => String(n).padStart(2, '0');
  const stamp = `${d.getFullYear()}${p2(d.getMonth() + 1)}${p2(d.getDate())}_${p2(d.getHours())}${p2(d.getMinutes())}`;
  const ts = `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}`;
  const lastCol = _colName(b.months.length);

  const sheet = [];
  const merges = [];
  sheet.push([{ v: t('hist.exportTitleRow',
    { my: histMyLabel(HIST.my), src: histSrcLabel(), basis: histBasisLabel() }), s: 1 }]);
  sheet.push([{ v: t('export.exportedAt', { ts }), s: 0 }]);
  sheet.push([]);
  for (let i = 1; i <= 2; i++) merges.push(`A${i}:${lastCol}${i}`);

  sheet.push([{ v: t('hist.model'), s: 2 }].concat(b.months.map((m) => ({ v: m, s: 2 }))));
  for (const g of b.models) {
    sheet.push([{ v: g.key, s: 3 }].concat(b.months.map((m) => {
      const cell = g.cells.get(m);
      if (!cell) return { v: '', s: 4 };
      const both = HIST.src === 'both';
      const srcs = both ? HIST_SRCS : [HIST.src === 'sam' ? 'sam' : 'wings'];
      const kinds = histKindsOf(g);
      const lines = cell.rows.map((r) => 'Commission ' + r['Commission no.']);
      if (cell.base) lines.push(t('hist.base'));
      if (cell.base && cell.prevMy) {
        lines.push(t('hist.vsPrevMy', { ref: histMyLabel(cell.prevMy.my) + ' ' + cell.prevMy.month }));
      }
      if (!cell.base || cell.prevMy) {
        for (const k of kinds) {
          for (const src of srcs) {
            const tag = (kinds.length > 1 ? t('hist.kind.' + k.key) + ' ' : '')
              + (both ? (src === 'sam' ? 'SAM ' : 'WINGS ') : '');
            const d = cell[src][k.key];
            if (d.added.length || d.removed.length) {
              d.added.forEach((c) => lines.push(tag + '+' + c));
              d.removed.forEach((c) => lines.push(tag + '-' + c));
            } else lines.push(tag + t('hist.same'));
          }
        }
      }
      for (const k of kinds) {
        for (const src of srcs) {
          const v = cell[src][k.key].varied;
          if (v.length) lines.push(histVariedText(v));
        }
      }
      // 크로스체크(그 달 WINGS↔SAM 차이)도 함께 내보낸다.
      if (both) {
        const x = histXOf(cell, kinds.map((k) => k.key));
        lines.push(t(x.same ? 'hist.xOk' : 'hist.xBad') + ' — ' + x.tip.split('\n').join(' / '));
      }
      return { v: lines.join('\n'), s: 3 };
    })));
  }

  const cols = [{ min: 1, max: 1, w: 32 }].concat(
    b.months.map((m, i) => ({ min: i + 2, max: i + 2, w: 24 })));
  writeXlsx(`afab_code_history_${histMyLabel(HIST.my)}_${stamp}.xlsx`,
    'History', sheet, merges, cols);
}

function initHistory() {
  histFillMy();
  const src = $('#histSrc');
  if (src) { src.value = HIST.src; enhanceSelect(src); }

  $('#histMy').addEventListener('input', (e) => { HIST.my = e.target.value; histRender(); });
  if (src) src.addEventListener('input', (e) => { HIST.src = e.target.value; histRender(); });
  $('#histSearch').addEventListener('input', (e) => { HIST.q = e.target.value; histRender(); });
  $('#histBasis').querySelectorAll('button[data-basis]').forEach((btn) =>
    btn.addEventListener('click', () => {
      HIST.basis = btn.dataset.basis;
      $('#histBasis').querySelectorAll('button').forEach((b2) => {
        b2.classList.toggle('active', b2 === btn);
        b2.setAttribute('aria-pressed', b2 === btn ? 'true' : 'false');
      });
      histRender();
    }));
  const oc = $('#histOnlyChanged');
  if (oc) oc.addEventListener('change', () => { HIST.onlyChanged = oc.checked; histRender(); });
  const ex = $('#histExport');
  if (ex) ex.addEventListener('click', histExportXlsx);

  histRender();
}

// ====================== SAM 파일 비교 (SAM 문서 ↔ SAM 문서) ======================
// MY 폴더 · 생산월 폴더 · 모델로 SAM 견적서 두 개(A=기준, B=비교)를 골라, 상세 창의
// SAM/WINGS 좌우 대조와 같은 모양으로 코드·도색·타이어를 맞대어 본다.
// 데이터는 빌드가 남긴 data.json 의 sam_files(읽은 문서 전체). 예전 data.json 에는 없으므로
// 그때는 주문에 붙은 문서만 행(Compared SAM file name)에서 모아 쓴다.
const SCMP = { a: { my: '', month: '', id: '', key: '' }, b: { my: '', month: '', id: '', key: '' },
  tab: 'diff', list: [], byId: new Map(), fromRows: false };
const SCMP_SIDES = ['a', 'b'];

// 파일명 → 화면용 모델 이름. 'Internal quotation_TR01-1  Actros-L 2863 LS 6x2 G5F 2026-08_(1).docx'
//   → label 'Actros-L 2863 LS 6x2 G5F', tag 'TR01-1'
function scmpLabelOf(file) {
  let s = String(file || '').split(/[\\/]/).pop().replace(/\.docx$/i, '');
  s = s.replace(/^Internal[\s_]*quotation[\s_]*/i, '');
  let tag = '';
  const m = /^([A-Z]{2}\d{2}(?:-\d+)?)[\s_]+/.exec(s);
  if (m) { tag = m[1]; s = s.slice(m[0].length); }
  s = s.replace(/[\s_]*\(\d+\)\s*$/, '').replace(/[\s_]+\d{4}-\d{2}(?:-\d{2})?\s*$/, '');
  s = s.replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
  return { label: s || String(file || ''), tag };
}
// 같은 모델인지 — 다른 달로 옮겨 가도 같은 모델을 계속 고르게 하는 키.
function scmpModelKey(e) { return e.label.toLowerCase().replace(/[\s-]+/g, '') + (e.pto ? '|pto' : ''); }

function scmpCatalog() {
  const out = [];
  const set = (csv) => new Set(splitCodes(csv));
  const src = Array.isArray(DATA.sam_files) ? DATA.sam_files : null;
  SCMP.fromRows = !src;
  if (src) {
    src.forEach((d) => {
      const my = /^MY\d{2}$/i.test(d.my || '') ? String(2000 + Number(d.my.slice(2)))
        : (computeMY({ _all_sam_codes: d.codes }) || HIST_NO_MY);
      out.push({ id: (d.folder || '') + '|' + d.file, folder: d.folder || '', file: d.file, my, month: d.month || '',
        now: d.model_now || '', bm: d.model_baumuster || '', pto: !!d.pto,
        codes: set(d.codes), paint: set(d.paint), tyre: set(d.tyre), ...scmpLabelOf(d.file) });
    });
  } else {
    const seen = new Set();
    for (const r of DATA.rows || []) {
      const f = String(r['Compared SAM file name'] || '');
      if (!f || seen.has(f)) continue;
      seen.add(f);
      const fm = /(\d{4}-\d{2})/.exec(f.split(/[\\/]/).pop());
      out.push({ id: 'row|' + f, folder: '', file: f, my: r.MY || HIST_NO_MY, month: fm ? fm[1] : prodMonth(r),
        now: r['SAM now'] || '', bm: r['SAM Baumuster'] || '', pto: /(^|[^A-Za-z])PTO([^A-Za-z]|$)/i.test(f),
        codes: set(r['_all_sam_codes']), paint: set(r['_paint_sam']), tyre: set(r['_tyre_sam']), ...scmpLabelOf(f) });
    }
  }
  // 같은 폴더에 이름이 같은 모델이 둘이면 태그(TR01-1 등)로 가른다.
  const groups = new Map();
  const gk = (e) => e.my + '|' + e.month + '|' + e.label;
  for (const e of out) {
    e.key = scmpModelKey(e);
    groups.set(gk(e), (groups.get(gk(e)) || 0) + 1);
  }
  for (const e of out) e.text = e.label + (groups.get(gk(e)) > 1 ? ' · ' + (e.tag || e.file) : '');
  out.sort((x, y) => (y.month || '').localeCompare(x.month || '') || x.text.localeCompare(y.text));
  return out;
}

// 처음 열 때: B = 가장 최근 문서 중 이전 생산월에도 같은 모델이 있는 것, A = 그 모델의 직전 문서.
function scmpDefaults() {
  const L = SCMP.list;
  if (!L.length) return;
  const b = L.find((e) => L.some((x) => x.key === e.key && x.month < e.month)) || L[0];
  const a = L.filter((x) => x.key === b.key && x.month < b.month)[0] || L.find((x) => x !== b) || b;
  SCMP.a = { my: a.my, month: a.month, id: a.id, key: a.key };
  SCMP.b = { my: b.my, month: b.month, id: b.id, key: b.key };
}

function scmpOptions(sel, items, value) {
  sel.innerHTML = items.map((o) => `<option value="${esc(o.v)}">${esc(o.t)}</option>`).join('');
  sel.value = value;
  enhanceSelect(sel);
}
// 한 쪽(A/B)의 MY → 생산월 → 모델 드롭다운을 채운다. 위 단계가 바뀌면 같은 모델을 유지하려 한다.
function scmpFillSide(side) {
  const S = SCMP[side], up = side.toUpperCase(), L = SCMP.list;
  const mySel = $('#scmpMy' + up), mSel = $('#scmpMonth' + up), fSel = $('#scmpFile' + up);
  if (!mySel) return;
  const mys = [...new Set(L.map((e) => e.my))].filter((v) => v !== HIST_NO_MY).sort().reverse();
  if (L.some((e) => e.my === HIST_NO_MY)) mys.push(HIST_NO_MY);
  if (!mys.includes(S.my)) S.my = mys[0] || '';
  const months = [...new Set(L.filter((e) => e.my === S.my).map((e) => e.month))].sort().reverse();
  if (!months.includes(S.month)) {
    // 같은 모델이 있는 가장 최근 달을 먼저 찾는다.
    S.month = months.find((m) => L.some((e) => e.my === S.my && e.month === m && e.key === S.key)) || months[0] || '';
  }
  const files = L.filter((e) => e.my === S.my && e.month === S.month).sort((x, y) => x.text.localeCompare(y.text));
  if (!files.some((e) => e.id === S.id)) {
    const same = files.find((e) => e.key === S.key);
    S.id = (same || files[0] || {}).id || '';
  }
  const cur = SCMP.byId.get(S.id);
  if (cur) S.key = cur.key;
  scmpOptions(mySel, mys.map((v) => ({ v, t: histMyLabel(v) })), S.my);
  scmpOptions(mSel, months.map((v) => ({ v, t: t('scmp.monthOpt', { m: v || '—' }) })), S.month);
  scmpOptions(fSel, files.map((e) => ({ v: e.id, t: e.text })), S.id);
  const fn = $('#scmpFileName' + up);
  if (fn) {
    fn.textContent = cur ? (cur.folder ? cur.folder + ' / ' : '') + cur.file.split(/[\\/]/).pop() : '';
    fn.title = fn.textContent;
  }
}

// 데이터가 바뀌면(빌드·언어 전환) 목록을 다시 만들고 화면을 새로 그린다.
function scmpRefresh() {
  SCMP.list = scmpCatalog();
  SCMP.byId = new Map(SCMP.list.map((e) => [e.id, e]));
  if (!SCMP.byId.has(SCMP.a.id) || !SCMP.byId.has(SCMP.b.id)) scmpDefaults();
  SCMP_SIDES.forEach(scmpFillSide);
  const note = $('#scmpNote');
  if (note) {
    const on = SCMP.fromRows && SCMP.list.length > 0;
    note.textContent = on ? t('scmp.oldData') : '';
    note.classList.toggle('hidden', !on);
  }
  scmpRender();
}

// 비교 계산 — 갈래(코드·도색·타이어)별 A에만 / B에만 / 공통.
function scmpDiff(A, B) {
  const out = {};
  for (const k of ['codes', 'paint', 'tyre']) {
    const a = A[k], b = B[k];
    out[k] = {
      onlyA: [...a].filter((c) => !b.has(c)).sort(),
      onlyB: [...b].filter((c) => !a.has(c)).sort(),
      common: [...a].filter((c) => b.has(c)).sort(),
    };
  }
  return out;
}
function scmpHead(side, e) { return `${side.toUpperCase()} · ${histMyLabel(e.my)} ${e.month}`; }

function scmpCardHtml(side, e) {
  const v = (x) => (x == null || String(x).trim() === '') ? '<span class="empty-cell">—</span>' : esc(x);
  return `<section class="scard scmp-card ${side}"><h3 class="scard-t"><span class="scmp-tag ${side}">${side.toUpperCase()}</span>` +
    `<span class="scmp-card-name">${esc(e.label)}</span></h3><div class="fgrid g4">` +
    fieldHtml('MY', v(e.my === HIST_NO_MY ? '' : e.my)) +
    fieldHtml(t('scmp.f.month'), v(e.month)) +
    fieldHtml('SAM now', v(e.now)) +
    fieldHtml('SAM Baumuster', v(e.bm)) +
    fieldHtml('PTO', e.pto ? 'PTO' : v('')) +
    fieldHtml(t('scmp.f.codes'), esc(fmtNum(e.codes.size))) +
    fieldHtml(t('scmp.f.file'), `<span class="scmp-fname" title="${esc(e.file)}">${esc(e.file.split(/[\\/]/).pop())}</span>`, 'span2') +
    `</div></section>`;
}

function scmpRender() {
  const el = $('#scmpBody');
  if (!el) return;
  if (!SCMP.list.length) {
    el.innerHTML = `<div class="hist-empty">${esc(t(DATA.rows && DATA.rows.length ? 'scmp.empty' : 'msg.noData'))}</div>`;
    return;
  }
  const A = SCMP.byId.get(SCMP.a.id), B = SCMP.byId.get(SCMP.b.id);
  if (!A || !B) { el.innerHTML = ''; return; }
  const d = scmpDiff(A, B);
  const heads = [scmpHead('a', A), scmpHead('b', B)];
  const nA = d.codes.onlyA.length + d.paint.onlyA.length + d.tyre.onlyA.length;
  const nB = d.codes.onlyB.length + d.paint.onlyB.length + d.tyre.onlyB.length;
  const nUnion = new Set([...A.codes, ...B.codes]).size;
  const stat = (label, value, cls) =>
    `<div class="hs-stat${cls ? ' ' + cls : ''}"><span class="hs-l">${esc(label)}</span><span class="hs-v">${value}</span></div>`;
  const sum = `<div class="hist-sum">` +
    stat(t('scmp.st.common'), esc(fmtNum(d.codes.common.length))) +
    stat(t('scmp.st.onlyA'), `<span class="hs-del">−${nA}</span>`) +
    stat(t('scmp.st.onlyB'), `<span class="hs-add">+${nB}</span>`) +
    stat(t('scmp.st.result'), esc(t(nA + nB ? 'scmp.st.diff' : 'scmp.st.same')), nA + nB ? 'warn' : '') +
    `<span class="hs-note">${esc(A.id === B.id ? t('scmp.sameFile') : t('scmp.note'))}</span></div>`;

  const sec = (title, hint, cls, codes) => {
    const body = codes.length ? codes.map((c) => codeItemHtml(c, cls, describe(c))).join('') : okLine(t('scmp.none'));
    return diffSecHtml({ title, hint, cls, n: codes.length, body });
  };
  const tab = (name, label, n) => `<button type="button" role="tab" class="dtab${SCMP.tab === name ? ' active' : ''}" ` +
    `data-tab="${name}" aria-selected="${SCMP.tab === name}">${esc(label)} <span class="cnt">${fmtNum(n)}</span></button>`;
  const join = (s) => [...s].sort().join(',');

  el.innerHTML = sum +
    `<div class="scmp-inner">` +
    `<div class="scards scmp-cards">${scmpCardHtml('a', A)}${scmpCardHtml('b', B)}</div>` +
    `<div class="dtabs-bar"><div class="dtabs" role="tablist">${tab('diff', t('scmp.tab.diff'), nA + nB)}${tab('full', t('scmp.tab.full'), nUnion)}</div>` +
    `<label class="search-box sm"><svg class="ic" aria-hidden="true"><use href="#i-search"/></svg>` +
    `<input type="search" id="scmpQ" placeholder="${esc(t('code.searchPh'))}" aria-label="${esc(t('code.searchPh'))}" /></label></div>` +
    `<div class="dpane scmp-pane${SCMP.tab === 'diff' ? '' : ' hidden'}" data-pane="diff">` +
      alignedGroupHtml('🎨 Paint', join(A.paint), join(B.paint), 'paint', true, heads) +
      alignedGroupHtml('🛞 Tyre', join(A.tyre), join(B.tyre), 'tyre', true, heads) +
      `<div class="dgrid">` +
        sec(t('scmp.onlyA'), t('scmp.onlyA.hint'), 'crit', d.codes.onlyA) +
        sec(t('scmp.onlyB'), t('scmp.onlyB.hint'), 'ok', d.codes.onlyB) +
      `</div><div class="q-none hidden">${esc(t('code.noHit'))}</div></div>` +
    `<div class="dpane scmp-pane${SCMP.tab === 'full' ? '' : ' hidden'}" data-pane="full">` +
      alignedGroupHtml('🎨 Paint', join(A.paint), join(B.paint), 'paint', false, heads) +
      alignedGroupHtml('🛞 Tyre', join(A.tyre), join(B.tyre), 'tyre', false, heads) +
      alignedFullHtml(join(A.codes), join(B.codes), heads) +
      `<div class="q-none hidden">${esc(t('code.noHit'))}</div></div>` +
    `</div>`;

  el.querySelectorAll('.dtab').forEach((b) => b.addEventListener('click', () => {
    SCMP.tab = b.dataset.tab;
    el.querySelectorAll('.dtab').forEach((x) => {
      const on = x === b;
      x.classList.toggle('active', on);
      x.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    el.querySelectorAll('.scmp-pane').forEach((p) => p.classList.toggle('hidden', p.dataset.pane !== SCMP.tab));
    scmpFilter();
  }));
  $('#scmpQ').addEventListener('input', scmpFilter);
}

// 코드 검색 — 보이는 탭의 줄(코드·설명)만 거른다(상세 창과 같은 방식).
function scmpFilter() {
  const qi = $('#scmpQ');
  const q = qi ? qi.value.trim().toLowerCase() : '';
  const pane = $('#scmpBody .scmp-pane:not(.hidden)');
  if (!pane) return;
  let hits = 0, total = 0;
  pane.querySelectorAll('[data-q]').forEach((x) => {
    total++;
    const hit = !q || x.dataset.q.includes(q);
    x.classList.toggle('q-hide', !hit);
    if (hit) hits++;
  });
  pane.classList.toggle('searching', !!q);
  const none = pane.querySelector('.q-none');
  if (none) none.classList.toggle('hidden', !q || hits > 0 || !total);
}

// 엑셀 — 갈래 · 코드 · 설명 · A · B · 구분 을 한 표로.
function scmpExportXlsx() {
  const A = SCMP.byId.get(SCMP.a.id), B = SCMP.byId.get(SCMP.b.id);
  if (!A || !B) return;
  const rows = [], merges = [];
  rows.push([{ v: `${t('scmp.title')} — A ${histMyLabel(A.my)} ${A.month} ${A.label}  ↔  B ${histMyLabel(B.my)} ${B.month} ${B.label}`, s: 1 }]);
  merges.push('A1:F1');
  rows.push([{ v: 'A', s: 5 }, { v: (A.folder ? A.folder + ' / ' : '') + A.file, s: 3 }]);
  rows.push([{ v: 'B', s: 5 }, { v: (B.folder ? B.folder + ' / ' : '') + B.file, s: 3 }]);
  merges.push('B2:F2', 'B3:F3');
  rows.push([]);
  rows.push(['Group', 'Code', 'Description', 'A', 'B', 'Result'].map((v) => ({ v, s: 2 })));
  for (const [k, label] of [['codes', 'Code'], ['paint', 'Paint'], ['tyre', 'Tyre']]) {
    const union = [...new Set([...A[k], ...B[k]])].sort();
    for (const c of union) {
      const inA = A[k].has(c), inB = B[k].has(c);
      const res = inA && inB ? t('scmp.x.common') : (inA ? t('scmp.x.onlyA') : t('scmp.x.onlyB'));
      const s = inA && inB ? 3 : 4;
      const desc = k === 'codes' ? (describe(c) || '') : (k === 'paint' ? 'MB ' + c : '');
      rows.push([{ v: label, s }, { v: c, s }, { v: desc, s },
        { v: inA ? 'O' : '', s }, { v: inB ? 'O' : '', s }, { v: res, s }]);
    }
  }
  const cols = [{ min: 1, max: 1, w: 9 }, { min: 2, max: 2, w: 14 }, { min: 3, max: 3, w: 54 },
    { min: 4, max: 5, w: 6 }, { min: 6, max: 6, w: 12 }];
  const name = (e) => `${histMyLabel(e.my)}_${e.month}_${e.label}`.replace(/[^\w.-]+/g, '_');
  writeXlsx(`sam_compare_${name(A)}__${name(B)}.xlsx`, 'SAM Compare', rows, merges, cols);
}

function initSamCmp() {
  SCMP_SIDES.forEach((side) => {
    const up = side.toUpperCase();
    const S = () => SCMP[side];
    $('#scmpMy' + up).addEventListener('input', (e) => { S().my = e.target.value; S().month = ''; scmpFillSide(side); scmpRender(); });
    $('#scmpMonth' + up).addEventListener('input', (e) => { S().month = e.target.value; scmpFillSide(side); scmpRender(); });
    $('#scmpFile' + up).addEventListener('input', (e) => {
      S().id = e.target.value;
      const cur = SCMP.byId.get(S().id);
      if (cur) S().key = cur.key;
      scmpFillSide(side);
      scmpRender();
    });
  });
  $('#scmpSwap').addEventListener('click', () => {
    [SCMP.a, SCMP.b] = [SCMP.b, SCMP.a];
    SCMP_SIDES.forEach(scmpFillSide);
    scmpRender();
  });
  $('#scmpExport').addEventListener('click', scmpExportXlsx);
  scmpRefresh();
}

// ====================== 공용 시트 에디터 (모델 매칭 · 코드 관리 공유) ======================
const MODEL_MAPPING_FILE = 'model_mapping.xlsx';
function xlsxReady() { return window.XLSX && !window.__XLSX_BLOCKED; }

// ---- 값 도우미: 시트 셀에는 문자열/불리언/숫자가 섞여 들어온다 ----
function isTrue(v) {
  if (typeof v === 'boolean') return v;
  return /^(true|1|y|yes|예|o|on)$/i.test(String(v ?? '').trim());
}
function isBoolCell(v) {
  if (typeof v === 'boolean') return true;
  return /^(true|false)$/i.test(String(v ?? '').trim());
}
function cellText(v) { return v == null ? '' : String(v); }

// 시트 한 장의 열(컬럼) 성격을 표본으로 추정한다 — 편집 폼의 입력 위젯과
// 목록 셀 표현(토글/색상칩/말줄임)을 고르는 데 쓴다.
function inferCols(aoa) {
  const header = aoa[0] || [];
  const ncol = aoa.reduce((m, r) => Math.max(m, (r || []).length), header.length || 1);
  const sample = aoa.slice(1, 301);
  const cols = [];
  for (let c = 0; c < ncol; c++) {
    const raw = sample.map((r) => (r || [])[c]);
    const vals = raw.map(cellText).map((s) => s.trim()).filter((s) => s !== '');
    const name = cellText(header[c]).trim();
    const lname = name.toLowerCase();
    let type = 'text';
    if (vals.length && raw.filter((v) => cellText(v).trim() !== '').every(isBoolCell)) type = 'bool';
    else if (/hex|colou?r|컬러|색상/.test(lname) || (vals.length && vals.every((v) => /^#[0-9a-f]{3,8}$/i.test(v)))) type = 'color';
    else if (vals.length && vals.reduce((a, v) => a + v.length, 0) / vals.length > 55) type = 'long';
    const distinct = [...new Set(vals)].sort();
    const numeric = type === 'text' && vals.length > 0 && vals.every((v) => /^-?\d+(\.\d+)?$/.test(v));
    // 코드성 열(code / 코드 / 접두어 …)은 목록에서 모노스페이스 + 강조색으로 보여준다.
    const mono = type === 'text' && !numeric && /(^|[_\s(])code|코드|prefix|접두어|baumuster/i.test(lname);
    cols.push({ c, type, numeric, mono, distinct, label: name || t('rec.colFallback', { n: c + 1 }) });
  }
  return cols;
}

// 목록 위 "필터" 드롭다운에 쓸 열: 값 종류가 적고 짧은 텍스트 열(카테고리 성격).
function pickFilterCol(cols, nrows) {
  for (const col of cols) {
    if (col.type !== 'text') continue;
    const d = col.distinct;
    if (d.length < 2 || d.length > 25) continue;
    if (nrows > 4 && d.length > Math.max(2, nrows / 2)) continue;
    if (d.some((v) => v.length > 24)) continue;
    return col;
  }
  return null;
}

// ---- 레코드 편집 모달 (코드 관리 · 모델 매칭 공용) ----
const RecModal = (function () {
  let cur = null;   // { cols, values, onSave, onDelete }

  function close() { cur = null; $('#recModal').classList.add('hidden'); $('#recBackdrop').classList.add('hidden'); }

  function fieldHtml(col, v, i) {
    const id = 'rmf' + i;
    let input;
    if (col.type === 'bool') {
      input = `<label class="switch"><input type="checkbox" id="${id}" data-i="${i}"${isTrue(v) ? ' checked' : ''} /><span class="track"></span></label>`;
    } else if (col.type === 'color') {
      const s = cellText(v).trim();
      const hex = /^#[0-9a-f]{6}$/i.test(s) ? s : '#ffffff';
      input = `<div class="color-field"><input type="color" class="rm-swatch" data-for="${id}" value="${esc(hex)}" />` +
        `<input type="text" id="${id}" data-i="${i}" value="${esc(cellText(v))}" placeholder="#1a1a1a" /></div>`;
    } else if (col.type === 'long') {
      input = `<textarea id="${id}" data-i="${i}" rows="3">${esc(cellText(v))}</textarea>`;
    } else {
      const useList = col.distinct.length > 1 && col.distinct.length <= 60;
      const dl = useList
        ? `<datalist id="${id}-dl">${col.distinct.map((d) => `<option value="${esc(d)}"></option>`).join('')}</datalist>` : '';
      input = `<input type="text" id="${id}" data-i="${i}" value="${esc(cellText(v))}"${useList ? ` list="${id}-dl"` : ''} />${dl}`;
    }
    return `<div class="rm-field${col.type === 'long' ? ' wide' : ''}">` +
      `<label for="${id}">${esc(col.label)}</label>${input}</div>`;
  }

  function open(o) {
    cur = o;
    $('#recModalTitle').textContent = o.title;
    $('#recModalBody').innerHTML =
      `<div class="rm-fields">${o.cols.map((col, i) => fieldHtml(col, o.values[col.c], i)).join('')}</div>`;
    $('#recModalDel').classList.toggle('hidden', !o.onDelete);
    $('#recModal').classList.remove('hidden');
    $('#recBackdrop').classList.remove('hidden');
    // id·정렬순서 같은 숫자 칸은 건너뛰고 첫 텍스트 칸에 포커스
    const idx = o.cols.findIndex((col) => col.type !== 'bool' && !col.numeric);
    const first = $('#recModalBody').querySelector(
      idx >= 0 ? `[data-i="${idx}"]` : 'input[type="text"], textarea');
    if (first) first.focus();
  }

  function collect() {
    const out = {};
    $('#recModalBody').querySelectorAll('[data-i]').forEach((inp) => {
      const col = cur.cols[+inp.dataset.i];
      out[col.c] = col.type === 'bool' ? inp.checked : inp.value;
    });
    return out;
  }

  // 색상 피커 ↔ 텍스트 입력 동기화
  document.addEventListener('input', (e) => {
    const sw = e.target.closest('.rm-swatch');
    if (sw) { const tx = document.getElementById(sw.dataset.for); if (tx) tx.value = sw.value; return; }
    const tx = e.target.closest('.color-field input[type="text"]');
    if (tx) {
      const sw2 = tx.parentElement.querySelector('.rm-swatch');
      if (sw2 && /^#[0-9a-f]{6}$/i.test(tx.value.trim())) sw2.value = tx.value.trim();
    }
  });
  $('#recModalClose').addEventListener('click', close);
  $('#recModalCancel').addEventListener('click', close);
  $('#recBackdrop').addEventListener('click', close);
  $('#recModalSave').addEventListener('click', () => { const o = cur; const v = collect(); close(); o.onSave(v); });
  $('#recModalDel').addEventListener('click', () => {
    if (!cur.onDelete || !confirm(t('rec.delConfirm'))) return;
    const o = cur; close(); o.onDelete();
  });
  document.addEventListener('keydown', (e) => {
    if (cur && e.key === 'Escape') close();
  });
  return { open, close, isOpen: () => !!cur };
})();

// 하나의 xlsx(다중 시트)를 편집 가능한 그리드로 다루는 재사용 에디터.
// 두 가지 편집 모드를 제공한다:
//   record — 행 = 한 건의 레코드. 목록 + 검색/필터/페이지 + 폼 모달로 등록·편집·삭제 (기본)
//   grid   — 기존 스프레드시트식 셀 직접 편집 (열 추가·헤더 변경 등 구조 수정용)
// cfg: { folderKey, fixedFile?, onSaved?, els:{...} }
const REC_PAGE_SIZE = 50;
function makeSheetEditor(cfg) {
  const S = { file: cfg.fixedFile || null, sheets: {}, order: [], sheet: null, dirty: false,
              mode: localStorage.getItem('editMode.' + cfg.folderKey) || 'record',
              cols: [], filterCol: null, page: 1 };
  const el = (k) => (cfg.els[k] ? document.getElementById(cfg.els[k]) : null);
  const stSel = '#' + cfg.els.status;
  // 자동 생성 시트(cfg.readOnlySheets)는 고쳐 봐야 다시 만들 때 덮어써지므로 편집을 막는다.
  const readOnly = () => (cfg.readOnlySheets || []).indexOf(S.sheet) !== -1;

  function setDirty(v) { S.dirty = v; const d = el('dirty'); if (d) d.classList.toggle('hidden', !v); }
  function curAoa() { return S.sheets[S.sheet] || []; }

  function renderTabs() {
    const tabs = el('tabs');
    tabs.classList.toggle('hidden', !!cfg.hideSingleTab && S.order.length <= 1);
    tabs.innerHTML = S.order.map((n) =>
      `<button class="sheet-tab${n === S.sheet ? ' active' : ''}" data-sheet="${esc(n)}">${esc(n)}</button>`).join('');
    tabs.querySelectorAll('.sheet-tab').forEach((b) =>
      b.addEventListener('click', () => {
        S.sheet = b.dataset.sheet;
        const s = el('search'); if (s) s.value = '';
        renderTabs(); refreshSheet();
      }));
  }

  // 시트가 바뀌었을 때: 열 성격 재추정 → 필터 드롭다운 재구성 → 현재 모드로 그리기
  function refreshSheet() {
    S.cols = inferCols(curAoa());
    S.page = 1;
    buildFilter();
    renderAll();
  }

  function buildFilter() {
    const sel = el('filter'); if (!sel) return;
    const aoa = curAoa();
    S.filterCol = S.mode === 'record' ? pickFilterCol(S.cols, Math.max(0, aoa.length - 1)) : null;
    if (!S.filterCol) { sel.classList.add('hidden'); sel.innerHTML = ''; return; }
    sel.classList.remove('hidden');
    sel.innerHTML = `<option value="">${esc(t('rec.allOf', { col: S.filterCol.label }))}</option>` +
      S.filterCol.distinct.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  }

  // 검색어·필터를 통과한 데이터 행 인덱스(헤더 제외)
  function matchedRows() {
    const aoa = curAoa();
    const s = el('search');
    const q = s ? s.value.trim().toLowerCase() : '';
    const fsel = el('filter');
    const fv = (S.filterCol && fsel && !fsel.classList.contains('hidden')) ? fsel.value : '';
    const out = [];
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      if (fv && cellText(row[S.filterCol.c]).trim() !== fv) continue;
      if (q && !row.map(cellText).join(' ').toLowerCase().includes(q)) continue;
      out.push(r);
    }
    return out;
  }

  function recCellHtml(col, v, r) {
    if (col.type === 'bool') {
      return `<label class="switch sm" title="${esc(isTrue(v) ? t('rec.on') : t('rec.off'))}">` +
        `<input type="checkbox" class="rec-toggle" data-r="${r}" data-c="${col.c}"${isTrue(v) ? ' checked' : ''} />` +
        `<span class="track"></span></label>`;
    }
    const s = cellText(v);
    if (!s.trim()) return '<span class="empty-cell">—</span>';
    if (col.type === 'color') {
      const hex = /^#[0-9a-f]{3,8}$/i.test(s.trim()) ? s.trim() : '';
      return (hex ? `<span class="swatch" style="background:${esc(hex)}"></span>` : '') +
        `<span class="mono">${esc(s)}</span>`;
    }
    return `<span class="rec-text" title="${esc(s)}">${esc(s)}</span>`;
  }

  function renderRecords() {
    const aoa = curAoa();
    const table = el('recs');
    const thead = table.querySelector('thead'), tbody = table.querySelector('tbody');
    const pager = el('pager'), cnt = el('count');
    if (!aoa.length) {
      thead.innerHTML = '';
      tbody.innerHTML = `<tr><td class="none">${esc(t('rec.emptySheet'))}</td></tr>`;
      if (pager) pager.classList.add('hidden');
      if (cnt) cnt.textContent = '';
      return;
    }
    const rows = matchedRows();
    const pages = Math.max(1, Math.ceil(rows.length / REC_PAGE_SIZE));
    if (S.page > pages) S.page = pages;
    const slice = rows.slice((S.page - 1) * REC_PAGE_SIZE, S.page * REC_PAGE_SIZE);

    const ro = readOnly();
    thead.innerHTML = '<tr><th class="rownum">#</th>' +
      S.cols.map((col) => `<th${col.type === 'bool' ? ' class="tight"' : (col.mono ? ' class="code-col-cell"' : '')}>${esc(col.label)}</th>`).join('') +
      (ro ? '' : `<th class="rowact">${esc(t('rec.actions'))}</th>`) + '</tr>';

    if (!slice.length) {
      tbody.innerHTML = `<tr><td class="none" colspan="${S.cols.length + (ro ? 1 : 2)}">${esc(t('rec.empty'))}</td></tr>`;
    } else {
      tbody.innerHTML = slice.map((r) => {
        const row = aoa[r] || [];
        return `<tr data-r="${r}"><td class="rownum">${r}</td>` +
          S.cols.map((col) => `<td${col.type === 'bool' ? ' class="tight"' : (col.mono ? ' class="code-col-cell"' : '')}>${recCellHtml(col, row[col.c], r)}</td>`).join('') +
          (ro ? '' : `<td class="rowact"><button type="button" class="link-btn rec-edit" data-r="${r}">${esc(t('btn.edit'))}</button>` +
            `<button type="button" class="link-btn danger rec-del" data-r="${r}">${esc(t('btn.delete'))}</button></td>`) + '</tr>';
      }).join('');
    }

    if (cnt) cnt.textContent = t('rec.count', { n: rows.length, total: Math.max(0, aoa.length - 1) });
    if (pager) {
      if (pages <= 1) { pager.classList.add('hidden'); pager.innerHTML = ''; }
      else {
        pager.classList.remove('hidden');
        pager.innerHTML =
          `<button type="button" class="icon-btn pg-prev"${S.page <= 1 ? ' disabled' : ''}>${esc(t('rec.prev'))}</button>` +
          `<span class="pg-label">${esc(t('rec.page', { page: S.page, pages }))}</span>` +
          `<button type="button" class="icon-btn pg-next"${S.page >= pages ? ' disabled' : ''}>${esc(t('rec.next'))}</button>`;
      }
    }
  }

  // 폼 모달로 한 건 편집 / 신규 등록
  function editRecord(r, prefill) {
    const aoa = curAoa();
    const isNew = r == null;
    const row = isNew ? (prefill || []) : (aoa[r] || []);
    RecModal.open({
      title: isNew ? t('rec.new') : t('rec.edit'),
      cols: S.cols,
      values: row,
      onDelete: isNew ? null : () => { aoa.splice(r, 1); setDirty(true); renderAll(); },
      onSave: (vals) => {
        const target = isNew ? [] : row;   // 미리 채운 값은 폼에만 — 저장은 폼 값으로
        for (const col of S.cols) {
          while (target.length <= col.c) target.push('');
          // 원래 숫자였던 칸(id·정렬순서 등)은 숫자로 되돌려 저장한다 — 폼 입력은 항상 문자열이므로.
          const nv = vals[col.c];
          const wasNum = typeof target[col.c] === 'number';
          target[col.c] = (typeof nv === 'string' && nv.trim() !== '' && (wasNum || col.numeric)
            && /^-?\d+(\.\d+)?$/.test(nv.trim())) ? Number(nv.trim()) : nv;
        }
        if (isNew) aoa.push(target);
        setDirty(true);
        if (isNew) { S.page = Math.max(1, Math.ceil(matchedRows().length / REC_PAGE_SIZE)); }
        renderAll();
      },
    });
  }

  function onRecClick(e) {
    const tg = e.target.closest('.rec-toggle');
    if (tg) {
      const aoa = curAoa(); const r = +tg.dataset.r, c = +tg.dataset.c;
      const row = aoa[r] || (aoa[r] = []);
      while (row.length <= c) row.push('');
      row[c] = tg.checked;
      setDirty(true);
      return;
    }
    const del = e.target.closest('.rec-del');
    if (del) {
      if (!confirm(t('rec.delConfirm'))) return;
      curAoa().splice(+del.dataset.r, 1); setDirty(true); renderAll(); return;
    }
    const ed = e.target.closest('.rec-edit');
    if (ed) { editRecord(+ed.dataset.r); return; }
    const tr = e.target.closest('tr[data-r]');
    if (tr) editRecord(+tr.dataset.r);
  }

  function setMode(mode) {
    S.mode = mode;
    localStorage.setItem('editMode.' + cfg.folderKey, mode);
    const ms = el('mode');
    if (ms) ms.querySelectorAll('button').forEach((b) => b.classList.toggle('active', b.dataset.mode === mode));
    const show = (k, on) => { const e2 = el(k); if (e2) e2.classList.toggle('hidden', !on); };
    show('recWrap', mode === 'record');
    show('gridWrap', mode === 'grid');
    show('addRec', mode === 'record' && !readOnly());
    show('addRow', mode === 'grid' && !readOnly());
    if (mode !== 'record') { const pg = el('pager'); if (pg) { pg.classList.add('hidden'); pg.innerHTML = ''; } }
    const cnt = el('count'); if (cnt && mode !== 'record') cnt.textContent = '';
    buildFilter();
    renderAll();
  }

  function renderAll() {
    if (S.mode === 'record') { renderRecords(); }
    else { renderGrid(); applySearch(); }
  }

  function renderGrid() {
    const aoa = curAoa();
    const grid = el('grid');
    if (!aoa.length) {
      grid.querySelector('thead').innerHTML = '';
      grid.querySelector('tbody').innerHTML = `<tr><td class="none">— 빈 시트 —</td></tr>`;
      return;
    }
    const ncol = aoa.reduce((m, r) => Math.max(m, r.length), 1);
    const header = aoa[0] || [];
    const ed = readOnly() ? 'false' : 'true';
    let thead = '<tr><th class="rownum">#</th>';
    for (let c = 0; c < ncol; c++)
      thead += `<th><div class="cell-edit hdr" contenteditable="${ed}" data-r="0" data-c="${c}">${esc(header[c] ?? '')}</div></th>`;
    thead += '<th class="rowact"></th></tr>';
    grid.querySelector('thead').innerHTML = thead;
    let body = '';
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      body += `<tr data-r="${r}"><td class="rownum">${r}</td>`;
      for (let c = 0; c < ncol; c++)
        body += `<td><div class="cell-edit" contenteditable="${ed}" data-r="${r}" data-c="${c}">${esc(row[c] ?? '')}</div></td>`;
      body += `<td class="rowact">${ed === 'true' ? `<button class="row-del" data-r="${r}" title="행 삭제">🗑</button>` : ''}</td></tr>`;
    }
    grid.querySelector('tbody').innerHTML = body;
  }

  function onInput(e) {
    if (readOnly()) return;
    const cell = e.target.closest('.cell-edit'); if (!cell) return;
    const r = +cell.dataset.r, c = +cell.dataset.c;
    const aoa = curAoa();
    while (aoa.length <= r) aoa.push([]);
    while (aoa[r].length <= c) aoa[r].push('');
    aoa[r][c] = cell.textContent;
    if (!S.dirty) setDirty(true);
  }
  function onClick(e) {
    const del = e.target.closest('.row-del'); if (!del) return;
    if (!confirm(t('codes.delRow'))) return;
    curAoa().splice(+del.dataset.r, 1);
    setDirty(true); renderGrid(); applySearch();
  }
  function addRow() {
    const aoa = curAoa();
    if (!aoa.length) aoa.push(['']);
    const ncol = aoa.reduce((m, r) => Math.max(m, r.length), 1);
    aoa.push(new Array(ncol).fill(''));
    setDirty(true); renderGrid();
    const gs = el('grid').closest('.grid-scroll'); if (gs) gs.scrollTop = gs.scrollHeight;
  }
  function applySearch() {
    const s = el('search'); if (!s) return;
    const q = s.value.trim().toLowerCase();
    el('grid').querySelector('tbody').querySelectorAll('tr').forEach((tr) => {
      tr.style.display = (!q || tr.textContent.toLowerCase().includes(q)) ? '' : 'none';
    });
  }
  function onSearch() {
    if (S.mode === 'record') { S.page = 1; renderRecords(); }
    else applySearch();
  }

  async function open(filename, isRevert) {
    if (!window.Graph || !Graph.available()) { setStatus(stSel, t('op.needLogin'), 'err'); return; }
    if (!xlsxReady()) { setStatus(stSel, t('alert.xlsxBlocked'), 'err'); return; }
    setStatus(stSel, t('codes.loadingFile'), 'info');
    try {
      const buf = await Graph.download(cfg.folderKey, filename);
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
      // 폐기한 시트(cfg.dropSheets)는 열지도, 다시 저장하지도 않는다 — 저장하면 파일에서 사라진다.
      const drop = new Set(cfg.dropSheets || []);
      S.file = filename; S.sheets = {}; S.order = wb.SheetNames.filter((n) => !drop.has(n));
      for (const n of S.order)
        S.sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, blankrows: false, defval: '' });
      S.sheet = S.order[0] || null;
      // onLoaded 가 워크북을 바꿨으면(예: 예전 형식 정리) true 또는 안내 문구를 돌려준다.
      const changed = cfg.onLoaded ? cfg.onLoaded(S) : false;
      setDirty(!!changed);
      const emp = el('empty'); if (emp) emp.classList.add('hidden');
      const mn = el('main'); if (mn) mn.classList.remove('hidden');
      const sb = el('search'); if (sb) sb.value = '';
      renderTabs();
      S.cols = inferCols(curAoa());
      S.page = 1;
      setMode(S.mode);
      if (changed) setStatus(stSel, typeof changed === 'string' ? changed : t('op.loaded'), 'info');
      else setStatus(stSel, isRevert ? t('op.loaded') : '', isRevert ? 'ok' : 'info');
    } catch (e) { setStatus(stSel, t('op.fail', { err: e.message }), 'err'); }
  }

  function workbook() {
    const wb = XLSX.utils.book_new();
    for (const n of S.order) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(S.sheets[n] || [[]]), n);
    return wb;
  }
  async function saveSp() {
    if (!window.Graph || !Graph.available()) { setStatus(stSel, t('op.needLogin'), 'err'); return; }
    if (!S.file || !xlsxReady()) { setStatus(stSel, t('alert.xlsxBlocked'), 'err'); return; }
    setStatus(stSel, t('codes.saving'), 'info');
    try {
      const buf = XLSX.write(workbook(), { bookType: 'xlsx', type: 'array' });
      await Graph.upload(cfg.folderKey, S.file, buf);
      setDirty(false); setStatus(stSel, t('codes.saved'), 'ok');
      if (cfg.onSaved) cfg.onSaved();
    } catch (e) { setStatus(stSel, t('op.fail', { err: e.message }), 'err'); }
  }
  function downloadLocal() { if (S.file && xlsxReady()) XLSX.writeFile(workbook(), S.file); }

  function wire() {
    el('grid').addEventListener('input', onInput);
    el('grid').addEventListener('click', onClick);
    const bind = (k, fn) => { const e = el(k); if (e) e.addEventListener('click', fn); };
    bind('addRow', addRow);
    bind('addRec', () => editRecord(null));
    bind('revert', () => { if (S.file) open(S.file, true); });
    bind('saveSp', saveSp);
    bind('download', downloadLocal);
    const s = el('search'); if (s) s.addEventListener('input', onSearch);
    const recs = el('recs'); if (recs) recs.addEventListener('click', onRecClick);
    const f = el('filter');
    if (f) f.addEventListener('change', () => { S.page = 1; renderRecords(); });
    const ms = el('mode');
    if (ms) ms.addEventListener('click', (e) => {
      const b = e.target.closest('button[data-mode]'); if (b) setMode(b.dataset.mode);
    });
    const pg = el('pager');
    if (pg) pg.addEventListener('click', (e) => {
      if (e.target.closest('.pg-prev')) { S.page = Math.max(1, S.page - 1); renderRecords(); }
      else if (e.target.closest('.pg-next')) { S.page += 1; renderRecords(); }
      const rw = el('recWrap'); if (rw) rw.scrollTop = 0;
    });
  }
  // 언어 전환 등으로 라벨을 다시 그려야 할 때
  function refresh() { if (!S.sheet) return; S.cols = inferCols(curAoa()); buildFilter(); renderAll(); }
  // 밖에서 미리 채운 새 레코드 창을 연다(모델 매칭 '규칙 추가').
  function addRecord(prefill) {
    if (S.mode !== 'record') setMode('record');
    if (!S.cols.length) S.cols = inferCols(curAoa());
    editRecord(null, prefill);
  }
  return { S, open, saveSp, wire, refresh, addRecord };
}

// ---- 매칭 문제 목록: 어디가 안 붙었는지를 한 표에 모은다 ----
// 예전에는 '미매칭_진단' 시트 · 진단 탭 · 대조표에 흩어져 있던 것을, 데이터(rows +
// diagnostics)에서 바로 만든다. 심각도: problem(고쳐야 함) / info(참고 — 주문이 아직 없는 등).
//   nosam   SAM 을 못 찾은 WINGS 주문 (모델·캡·PTO·생산월로 묶음)
//   unused  읽혔지만 그 달 주문 어디에도 안 붙은 SAM 문서
//   skipped 읽지 못하고 건너뛴 SAM 파일
//   month   생산월 폴더별 읽은 파일·인식 번호 (참고)
const MP_KINDS = ['nosam', 'unused', 'skipped', 'month'];
function modelNumOf(model) {
  const m = /(\d{4})/.exec(String(model || ''));
  return m ? m[1] : '';
}
// SAM 폴더 이름 'MY27/2026-04' · '2026-04 생산' → 생산월 'YYYY-MM' 과 MY '2027' 로 나눈다.
function folderProd(name) {
  const m = /(\d{4})-(\d{2})/.exec(String(name || ''));
  return m ? m[1] + '-' + m[2] : '';
}
function folderMy(name) {
  const m = /(?:^|\/)MY(\d{2})(?![0-9])/i.exec(String(name || ''));
  return m ? '20' + m[1] : '';
}
// 안 붙은 SAM 문서가 '왜' 안 붙었는지 — 그 생산월 주문 중 같은 번호(번호+끝글자) 주문을 본다.
//   none   그 번호 주문이 그 달에 없음 → 고칠 것 아님(참고)
//   pto    같은 번호 주문은 있지만 PTO 여부가 전부 다름 → 그 주문들이 다른 견적서에 붙었으면 참고, 아니면 확인
//   other  같은 번호·같은 PTO 주문이 모두 다른 문서에 붙음(캡·Baumuster 가 더 맞음) → 참고
//   nosam  같은 조건 주문인데 SAM 을 못 찾은 것이 있음 → 확인 필요
// 예전처럼 '그 달 전체 주문 수'를 쓰면 모든 문서가 같은 숫자로 보여 원인을 가를 수 없었다.
function unusedCause(u) {
  const norm = (window.SamParse && SamParse.normalizeModel)
    || ((x) => String(x || '').toUpperCase().replace(/\d[X]\d/g, '').replace(/DNA/g, '').replace(/[^A-Z0-9]/g, ''));
  const prod = folderProd(u.month);
  const keys = new Set(u.keys || []);
  const nums = new Set([...keys].map((k) => k.slice(0, 4)));
  const inMonth = (DATA.rows || []).filter((r) => prodMonth(r) === prod);
  const same = inMonth.filter((r) => keys.has(norm(r['Model(WINGS)'])));
  if (!same.length) {
    const near = [...new Set(inMonth.map((r) => norm(r['Model(WINGS)']))
      .filter((k) => nums.has(k.slice(0, 4)) && !keys.has(k)))];
    return { cause: 'none', sev: 'info', n: 0,
      why: near.length ? t('mp.un.noneNear', { near: near.join(', ') }) : t('mp.un.none') };
  }
  const samePto = same.filter((r) => (String(r.PTO || '') === 'PTO') === !!u.pto);
  if (!samePto.length) {
    // 그 달 PTO판·비PTO판이 같이 있고 주문이 한쪽뿐이면, 다른 쪽 문서가 남는 건 정상이다.
    // 주문이 모두 (다른 견적서에) 붙었으면 참고, 못 붙은 주문이 있으면 확인 필요.
    const allMatched = same.every((r) => r['Compared SAM file name']);
    return { cause: 'pto', sev: allMatched ? 'info' : 'problem', n: same.length,
      why: t((u.pto ? 'mp.un.ptoDocPto' : 'mp.un.ptoDocNon') + (allMatched ? 'Ok' : ''), { n: same.length }) };
  }
  const noSam = samePto.filter((r) => !r['Compared SAM file name']);
  if (noSam.length) {
    return { cause: 'nosam', sev: 'problem', n: samePto.length,
      why: t('mp.un.nosam', { n: noSam.length }) };
  }
  const files = [...new Set(samePto.map((r) => r['Compared SAM file name']))];
  return { cause: 'other', sev: 'info', n: samePto.length,
    why: t('mp.un.other', { n: samePto.length, files: files.slice(0, 3).join(', ') + (files.length > 3 ? ' …' : '') }) };
}

function matchProblems() {
  const g = (DATA && DATA.diagnostics) || null;
  const months = (g && g.sam_months) || [];
  // 같은 생산월이 연식(MY) 폴더마다 따로 있을 수 있다 — 인식 번호는 생산월 단위로 합친다.
  const monthByYm = new Map();
  for (const m of months) {
    const k = String(m.yyyymm);
    const prev = monthByYm.get(k);
    monthByYm.set(k, prev ? { keys: [...new Set([...(prev.keys || []), ...(m.keys || [])])] } : m);
  }
  const out = [];

  const groups = new Map();
  for (const r of DATA.rows || []) {
    if (String(r._orig_status || r['SAM Status'] || '') !== 'No SAM') continue;
    const ym = String(r['Production date'] || '').slice(0, 7);
    const my = String(r.MY || '');
    const k = [r['Model(WINGS)'], r['Cab'], r['PTO'], ym, my].join('|');
    if (!groups.has(k)) groups.set(k, { r, ym, my, n: 0, comms: [] });
    const gg = groups.get(k);
    gg.n += 1;
    gg.comms.push(String(r['Commission no.']));
  }
  for (const gg of [...groups.values()].sort((a, b) => b.n - a.n)) {
    const mf = monthByYm.get(gg.ym.replace('-', ''));
    const num = modelNumOf(gg.r['Model(WINGS)']);
    let why;
    if (!g) why = t('mp.why.nosamNoDiag');
    else if (!mf) why = t('mp.why.noFolder', { ym: gg.ym || '?' });
    else {
      const keys = (mf.keys || []).filter((k) => /^\d{4}/.test(k));
      const near = num ? keys.filter((k) => k.slice(0, 2) === num.slice(0, 2)) : [];
      why = t('mp.why.noDoc', { num: num || gg.r['Model(WINGS)'],
        keys: (near.length ? near : keys).slice(0, 12).join(', ') || '—' });
    }
    // MY 를 나누기 전의 key(legacyKey)로 무시해 둔 것도 계속 인정한다.
    const legacyKey = ['nosam', gg.r['Model(WINGS)'] || '', gg.r['Cab'] || '', gg.r['PTO'] || '', gg.ym].join('|');
    out.push({ key: gg.my ? legacyKey + '|MY' + gg.my : legacyKey, legacyKey,
      kind: 'nosam', sev: 'problem', model: gg.r['Model(WINGS)'] || '', cab: gg.r['Cab'] || '',
      pto: gg.r['PTO'] || '', ym: gg.ym, my: gg.my, n: gg.n, file: '', why, num, comms: gg.comms,
      sample: gg.r });
  }
  for (const u of ((g && g.sam_unused) || [])) {
    const c = unusedCause(u);
    const keys = u.keys || [];
    out.push({ key: ['unused', u.month, u.file].join('|'),
      kind: 'unused', sev: c.sev, cause: c.cause, model: keys.join(', '),
      modelTitle: keys.length > 1 ? t('mp.keysTitle2', { a: keys[0], b: keys[1] })
        : (keys.length ? t('mp.keysTitle1', { a: keys[0] }) : ''),
      cab: '', pto: u.pto ? 'PTO' : '', ym: folderProd(u.month), my: folderMy(u.month), n: c.n, file: u.file,
      why: c.why });
  }
  for (const sk of ((g && g.sam_skipped) || [])) {
    out.push({ key: ['skipped', sk.month, sk.file].join('|'),
      kind: 'skipped', sev: 'problem', model: '', cab: '', pto: '', ym: folderProd(sk.month), my: folderMy(sk.month), n: '',
      file: sk.file, why: sk.reason });
  }
  for (const m of months) {
    const orders = ((g && g.orders_by_month) || {})[m.yyyymm] || 0;
    out.push({ key: ['month', m.name].join('|'),
      kind: 'month', sev: 'info', model: '', cab: '', pto: '', ym: folderProd(m.name), my: folderMy(m.name), n: orders, file: m.name,
      why: t('match.diagMonthWhy', { parsed: m.parsed, files: m.files, keys: m.keys.length })
        + (m.keys.length ? ' — ' + m.keys.join(', ') : '') });
  }
  return out;
}

// ---- 매칭 문제 '무시' 목록 ----
// WINGS 는 SAM 이 없는 과거 모델까지 매칭하려 하므로, 고칠 수 없는 미매칭이 많다.
// Admin 이 골라 '무시'하면 고칠 것 수·배지에서 빠진다(목록에서는 '무시됨' 칩으로 계속 볼 수 있음).
// 공유되어야 하므로 SharePoint '06. exceptions/match-ignore.json' 에 둔다 — Mismatch 예외와
// 같은 방식(eTag 조건부 저장 + log 누적).
//   key = nosam|모델|캡|PTO|생산월[|MY연식]  ·  unused|생산월폴더|파일  ·  skipped|생산월폴더|파일
const MI_FILE = 'match-ignore.json';
const MI_LOCAL_KEY = 'sam-afab-match-ignore';
let MI = { version: 1, items: {}, log: [], updated_at: '' };

function miItem(p) { return MI.items[p.key] || (p.legacyKey ? MI.items[p.legacyKey] : undefined); }
async function miRead() {
  if (excLocal()) {
    let d = null;
    try { d = JSON.parse(localStorage.getItem(MI_LOCAL_KEY) || 'null'); } catch { /* 저장소 막힘 */ }
    return { data: d, etag: null };
  }
  return Graph.readJsonVersioned('exceptions', MI_FILE);
}
async function loadMatchIgnore() {
  try { MI = excNormalize((await miRead()).data); }
  catch (e) { console.warn('[match-ignore] 읽기 실패 — 무시 목록 미적용:', e.message); }
}
// action = 'ignore' | 'unignore'. probs = 문제 줄(matchProblems 항목)들.
async function miMutate(action, probs, note) {
  const u = excUser();
  for (let attempt = 0; attempt < 4; attempt++) {
    const { data, etag } = await miRead();
    const doc = excNormalize(data);
    const now = new Date().toISOString();
    let done = 0;
    for (const p of probs) {
      const has = !!doc.items[p.key] || !!(p.legacyKey && doc.items[p.legacyKey]);
      if (action === 'ignore' && !has) {
        doc.items[p.key] = { key: p.key, kind: p.kind, model: p.model, cab: p.cab, pto: p.pto, ym: p.ym,
          file: p.file, note: note || '', by: u.id, by_name: u.name, at: now };
      } else if (action === 'unignore' && (has || (p.legacyKey && doc.items[p.legacyKey]))) {
        delete doc.items[p.key];
        if (p.legacyKey) delete doc.items[p.legacyKey];
      } else continue;
      doc.log.push({ at: now, by: u.id, by_name: u.name, action, key: p.key, text: note || '' });
      done++;
    }
    if (!done) { MI = doc; return 0; }
    doc.updated_at = now;
    if (excLocal()) {
      try { localStorage.setItem(MI_LOCAL_KEY, JSON.stringify(doc)); } catch { /* 이번 세션만 */ }
      MI = doc;
      return done;
    }
    try {
      if (!etag) await Graph.ensureFolder('exceptions');
      await Graph.writeJsonVersioned('exceptions', MI_FILE, doc, etag);
      MI = doc;
      return done;
    } catch (e) {
      if (e.conflict && attempt < 3) continue;
      throw e;
    }
  }
  return 0;
}

// ---- 예전 워크북 → '모델명변경' 시트 한 장 ----
// 규칙은 '모델명 변경'(From → To) 한 가지만 남았다. 예전 워크북의 표시치환은 그대로 옮기고,
// 폐기한 대체번호·SAM별칭·옵션 덕분에 매칭되던 모델은 지금 데이터에서 찾아 자동으로 추가한다
// — 그래야 규칙을 지워도 지금 붙어 있는 차가 No SAM 으로 떨어지지 않는다.
// 지금 데이터는 예전 규칙으로 계산된 것이어야 한다: 그래서 이 정리는 빌드 '전에' 한다.
function prettyModel(v) {
  const s = String(v || '').trim();
  const m = /^(\d{4})\s*([A-Z]+)$/i.exec(s);
  return m ? `${m[1]} ${m[2].toUpperCase()}` : s;
}
// 매칭된 차 중 WINGS 모델명이 붙은 SAM 문서의 두 번호(파일명·본문)와 다른 것 → From → To(SAM now).
function neededRenames() {
  const norm = (window.SamParse && SamParse.normalizeModel) || ((x) => String(x || '').toUpperCase().replace(/[^A-Z0-9]/g, ''));
  const tally = new Map();
  for (const r of DATA.rows || []) {
    if (!r['Compared SAM file name']) continue;
    const from = String(r['Model(WINGS)'] || '').trim();
    if (!from) continue;
    const docKeys = new Set([norm(r['SAM now']), norm(r['SAM Baumuster'])].filter(Boolean));
    if (!docKeys.size || docKeys.has(norm(from))) continue;
    const to = prettyModel(r['SAM now'] || r['SAM Baumuster']);
    if (!tally.has(from)) tally.set(from, new Map());
    const m = tally.get(from);
    m.set(to, (m.get(to) || 0) + 1);
  }
  const out = [];
  for (const [from, m] of tally) {
    const tos = [...m.entries()].sort((a, b) => b[1] - a[1]);
    out.push({ from, to: tos[0][0], n: tos[0][1], alt: tos.slice(1).map(([to, n]) => `${to}(${n})`) });
  }
  return out.sort((a, b) => b.n - a.n);
}
// 예전 규칙(시트들) → 모델명 변경 AOA + 무엇을 더하고 뺐는지.
function migrateRules(names, getAoa) {
  const rules = RefData.rulesFromSheets(names, getAoa);
  const map = rules.model_rename || {};
  const notes = {};
  const added = [], removed = [];
  for (const nd of neededRenames()) {
    // 표시치환 X→From 이 매칭을 깨던 경우(필요한 것이 From→X): 그 표시치환을 지운다.
    const back = Object.keys(map).find((x) => map[x] === nd.from && x === nd.to);
    if (back) { delete map[back]; removed.push(`${back} → ${nd.from}`); continue; }
    if (map[nd.from] === nd.to) continue;
    map[nd.from] = nd.to;
    notes[nd.from] = t('matching.autoNote', { n: nd.n }) + (nd.alt.length ? ' · ' + nd.alt.join(', ') : '');
    added.push(`${nd.from} → ${nd.to} (${nd.n})`);
  }
  return { aoa: RefData.rulesToAoa({ model_rename: map }, notes), added, removed };
}
function migrationMsg(r) {
  return t('matching.migrated', { n: r.aoa.length - 1, added: r.added.length, removed: r.removed.length })
    + (r.added.length ? ' ' + t('matching.migAdded', { list: r.added.slice(0, 8).join(', ') + (r.added.length > 8 ? ' …' : '') }) : '')
    + (r.removed.length ? ' ' + t('matching.migRemoved', { list: r.removed.join(', ') }) : '');
}
// 편집기에서 예전 워크북을 열면 한 장으로 바꿔 보여 준다(저장해야 파일에 반영). 안내 문구를 돌려준다.
function unifyRuleSheets(S) {
  const RS = RefData.RULE_SHEET;
  if (S.order.length === 1 && S.order[0] === RS) return false;
  const r = S.sheets[RS] ? { aoa: S.sheets[RS], added: [], removed: [] }
    : migrateRules(S.order, (n) => S.sheets[n]);
  S.sheets = { [RS]: r.aoa };
  S.order = [RS];
  S.sheet = RS;
  return migrationMsg(r);
}

// ---- pto-codes.xlsx 가 없으면 기본값으로 만들어 둔다 ----
// PTO 판정 기준은 이 워크북 하나다. 파일이 없으면 앱은 내장 기본값으로 조용히 돌지만,
// 그러면 코드 관리 화면에 뜨지 않아 사람이 검토할 수가 없다. 그래서 한 번은 만들어 준다.
// 이미 있으면 손대지 않는다 — 사람이 분류한 내용을 덮어쓰면 안 되기 때문.
const PTO_CODES_FILE = 'pto-codes.xlsx';
async function ensurePtoCodesWorkbook() {
  const items = await Graph.list('code');
  if (items.some((it) => !it.isFolder && it.name.toLowerCase() === PTO_CODES_FILE)) return null;
  const d = (window.RefData && RefData.PTO_DEFAULTS) || { codes: [], excepts: [] };
  const desc = (CODES && CODES.options) || {};
  const aoa = [['유형(Type)', '코드(Code)', '설명(Description)', '비고(Note)']];
  for (const c of d.codes) aoa.push(['PTO코드(pto)', c, desc[c] || '', 'PTO 판정에 사용']);
  for (const c of d.excepts) {
    aoa.push(['제외(except)', c, desc[c] || '', 'PTO 관련이지만 판정에서 제외 — 준비·제어 코드']);
  }
  aoa.push([]);
  aoa.push(['※ WINGS 주문·SAM 견적서 양쪽 모두 여기 "PTO코드" 를 하나라도 가지면 PTO 로 봅니다. '
    + 'PTO 차량과 비PTO 차량은 서로 비교하지 않으므로, 이 목록이 곧 매칭 기준입니다. '
    + '준비/제어 코드처럼 PTO 가 없는 차에도 붙는 코드는 "제외(except)" 로 두세요.']);
  // 검토용 원본 — 설명에 PTO 가 들어간 코드 전부. 어느 것을 목록에 넣을지 판단할 근거이고,
  // 새 코드가 생기면 '미분류' 로 드러난다. 편집해도 판정에는 쓰이지 않는다.
  const listed = new Set([...d.codes, ...d.excepts]);
  const mentionsPto = (t2) => {
    const x = String(t2 || '');
    return x.indexOf('PTO') !== -1
      || /(?<![A-Za-z])pto(?![A-Za-z])/i.test(x)
      || /power[ -]*take[ -]*off/i.test(x);
  };
  const ref = [['코드(Code)', '설명(Description)', '현재 분류']];
  for (const c of Object.keys(desc).sort()) {
    if (!mentionsPto(desc[c])) continue;
    ref.push([c, desc[c], d.codes.indexOf(c) !== -1 ? 'PTO코드'
      : (d.excepts.indexOf(c) !== -1 ? '제외' : '미분류 — 검토 필요')]);
  }
  ref.push([]);
  ref.push(['※ 확인용 목록입니다(편집해도 판정에 쓰이지 않음). PTO 시트에 넣을지 여기서 판단하세요.']);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'PTO');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ref), '설명에_PTO포함(Ref)');
  await Graph.upload('code', PTO_CODES_FILE, XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
  return { name: PTO_CODES_FILE, codes: d.codes.length, excepts: d.excepts.length,
    review: Math.max(0, ref.length - 3 - listed.size) };
}

// ---- 빌드 직전에 model_mapping.xlsx 를 '모델명변경' 한 장짜리로 정리한다 ----
// 이미 한 장짜리면 손대지 않는다. 예전 워크북이면 지금 데이터(예전 규칙으로 계산된 결과)로
// 필요한 모델명 변경을 더해 저장한다 — 그 뒤에 도는 빌드가 새 규칙을 읽는다.
async function migrateModelMapping() {
  const RS = RefData.RULE_SHEET;
  const items = await Graph.list('model_rules');
  const found = items.find((it) => !it.isFolder
    && it.name.toLowerCase() === MODEL_MAPPING_FILE.toLowerCase());
  const name = found ? found.name : MODEL_MAPPING_FILE;
  let r;
  if (found) {
    const wb = XLSX.read(new Uint8Array(await Graph.download('model_rules', name)), { type: 'array' });
    if (wb.SheetNames.length === 1 && wb.SheetNames[0] === RS) return { name, changed: false };
    const get = (n) => XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, blankrows: false, defval: '' });
    r = wb.SheetNames.indexOf(RS) !== -1 ? { aoa: get(RS), added: [], removed: [] }
      : migrateRules(wb.SheetNames, get);
  } else {
    r = { aoa: RefData.rulesToAoa(RefData.loadRules(null)), added: [], removed: [] };
  }
  const out = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(out, XLSX.utils.aoa_to_sheet(r.aoa), RS);
  await Graph.upload('model_rules', name, XLSX.write(out, { bookType: 'xlsx', type: 'array' }));
  return { name, changed: true, msg: migrationMsg(r) };
}

// ---- 모델 매칭: model_mapping.xlsx 의 '모델명변경' 시트 편집 ----
const matchEditor = makeSheetEditor({
  folderKey: 'model_rules', fixedFile: MODEL_MAPPING_FILE,
  // 예전 워크북이면 열면서 한 장으로 합친다(저장해야 파일에 반영).
  onLoaded: unifyRuleSheets,
  hideSingleTab: true,
  els: { tabs: 'matchTabs', grid: 'matchGrid', addRow: 'matchAddRow', revert: 'matchLoadSp',
         saveSp: 'matchSaveSp', download: 'matchDownload', dirty: 'matchDirty',
         status: 'matchStatus', search: 'matchSearch', main: 'matchMain', empty: 'matchEmpty',
         recs: 'matchRecs', recWrap: 'matchRecWrap', gridWrap: 'matchGridWrap', pager: 'matchPager',
         mode: 'matchMode', filter: 'matchFilter', count: 'matchCount', addRec: 'matchAddRec' },
});
const MATCH_PANES = ['problem', 'result', 'rules'];
let MP_FILTER = 'problem';   // 문제 탭의 칩 필터: problem(고칠 것) | all | nosam | unused | skipped | month
function initMatching() {
  const link = $('#matchFolderLink'); if (window.Graph) link.href = Graph.folders.model_rules.shareUrl;
  matchEditor.wire();
  matchEditor.open(MODEL_MAPPING_FILE);
  for (const id of MS_FIELDS.map((f) => f.id).concat(['#msStatus', '#msSearch'])) {
    const el = $(id);
    if (el) el.addEventListener('input', renderMatchSummary);
  }
  const pane = $('#matchPane');
  if (pane) {
    pane.addEventListener('click', (e) => {
      const b = e.target.closest('button[data-pane]');
      if (b) showMatchPane(b.dataset.pane);
    });
  }
  for (const id of ['#mpSearch', '#mpFrom', '#mpTo', '#mpMy', '#mpModel']) {
    const el = $(id);
    if (el) el.addEventListener('input', renderMatchProblems);
  }
  const mpI = $('#mpIgnore'); if (mpI) mpI.addEventListener('click', () => mpBulk('ignore'));
  const mpU = $('#mpUnignore'); if (mpU) mpU.addEventListener('click', () => mpBulk('unignore'));
  const chips = $('#mpChips');
  if (chips) chips.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-mp]');
    if (b) { MP_FILTER = b.dataset.mp; renderMatchProblems(); }
  });
  const grid = $('#matchProbGrid');
  if (grid) grid.addEventListener('change', (e) => {
    if (e.target.id === 'mpAll') {
      for (const p of MP_SHOWN) {
        if (p.kind === 'month') continue;
        if (e.target.checked) MP_SEL.add(p.key); else MP_SEL.delete(p.key);
      }
      grid.querySelectorAll('input[data-sel]').forEach((cb) => { cb.checked = e.target.checked; });
      mpSyncBulk();
      return;
    }
    const cb = e.target.closest('input[data-sel]');
    if (cb) { if (cb.checked) MP_SEL.add(cb.dataset.sel); else MP_SEL.delete(cb.dataset.sel); mpSyncBulk(); }
  });
  if (grid) grid.addEventListener('click', (e) => {
    const add = e.target.closest('button[data-add-rule]');
    if (add) { addRuleFor(add.dataset.addRule); return; }
    const open = e.target.closest('button[data-open]');
    if (open) { const r = rowByComm(open.dataset.open); if (r) openDrawer(r); }
  });
  let saved = 'problem';
  try { saved = localStorage.getItem('matchPane') || 'problem'; } catch { /* 기본 */ }
  showMatchPane(saved);
  renderMatchSummary();
}

// 문제 / 매칭 결과 / 매칭 규칙은 한 번에 하나만 — 보이는 쪽이 화면 높이를 다 쓴다.
function showMatchPane(name) {
  const which = MATCH_PANES.indexOf(name) !== -1 ? name : 'problem';
  try { localStorage.setItem('matchPane', which); } catch { /* 이번 세션만 */ }
  const sum = $('#matchSum'), ed = $('#matchEditor'), pb = $('#matchProb');
  if (sum) sum.classList.toggle('hidden', which !== 'result');
  if (ed) ed.classList.toggle('hidden', which !== 'rules');
  if (pb) pb.classList.toggle('hidden', which !== 'problem');
  if (which === 'problem') renderMatchProblems();
  const pane = $('#matchPane');
  if (pane) {
    pane.querySelectorAll('button[data-pane]').forEach((b) =>
      b.classList.toggle('active', b.dataset.pane === which));
  }
}

// 문제 줄의 '규칙 추가' — 규칙 탭으로 넘어가 From(WINGS 모델명)을 미리 채운 편집 창을 연다.
function addRuleFor(model) {
  showMatchPane('rules');
  if (!matchEditor.S.file || !matchEditor.S.sheet) { alert(t('matching.loading')); return; }
  matchEditor.addRecord([model || '', '', '']);
}

// ---- 매칭 결과: 지금 데이터의 '모델 → 실제로 비교된 SAM 파일' ----
// 한 줄 = 모델 하나(대시보드와 같은 Model · Type · Axle · Cab · MY · PTO).
// 매칭은 차 한 대씩 그 차의 생산월 폴더를 보므로, 같은 모델이라도 생산월에 따라 다른
// 견적서(월별 개정본)에 붙는다. 그 문서들은 한 줄 안에 접어 두고 펼쳐서 본다.
function tallyStatus(acc, status) {
  if (status === 'Match') acc.match++;
  else if (status === 'Mismatch') acc.mismatch++;
  else acc.noSam++;
}
function matchSummaryRows() {
  const m = new Map();
  for (const r of DATA.rows) {
    const key = HIST_COLS.map((k) => String(r[k] ?? '').trim()).join('');
    let g = m.get(key);
    if (!g) {
      g = { key: key, sample: r, n: 0, match: 0, mismatch: 0, noSam: 0, files: new Map() };
      m.set(key, g);
    }
    const st = String(r['SAM Status'] || '');
    g.n++;
    tallyStatus(g, st);
    const f = String(r['Compared SAM file name'] || '');
    let fe = g.files.get(f);
    if (!fe) { fe = { file: f, n: 0, match: 0, mismatch: 0, noSam: 0 }; g.files.set(f, fe); }
    fe.n++;
    tallyStatus(fe, st);
  }
  const name = (g) => HIST_COLS.map((k) => String(g.sample[k] ?? '')).join(' ');
  const out = [...m.values()];
  for (const g of out) {
    g.fileList = [...g.files.values()]
      .sort((a, b) => b.n - a.n || a.file.localeCompare(b.file));
  }
  return out.sort((a, b) => b.n - a.n || name(a).localeCompare(name(b)));
}

// 매칭 결과 위의 필터 — 대시보드와 같은 항목(+상태)으로 좁혀 본다.
const MS_FIELDS = [
  { id: '#msModel', key: 'Vehicle', allKey: 'filter.allModel' },
  { id: '#msType', key: 'Model(WINGS)', allKey: 'filter.allType' },
  { id: '#msAxle', key: 'Type', allKey: 'filter.allAxle' },
  { id: '#msCab', key: 'Cab', allKey: 'filter.allCab' },
  { id: '#msMy', key: 'MY', allKey: 'filter.allMY' },
  { id: '#msPto', key: 'PTO', allKey: 'match.allPto' },
];
// 펼쳐 둔 모델(그룹 키) — 다시 그려도 유지된다.
const MS_OPEN = new Set();
function msSearchTerm() {
  const el = $('#msSearch');
  return el ? el.value.trim() : '';
}
// 검색어를 셀에서 노랗게 짚어 준다(대시보드 목록의 hl 과 같은 방식, 검색창만 다름).
function msHl(text) {
  const q = msSearchTerm();
  const safe = esc(text);
  if (!q) return safe;
  try {
    const re = new RegExp('(' + q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
    return safe.replace(re, '<mark>$1</mark>');
  } catch { return safe; }
}
function msFilterGroups(groups) {
  const q = msSearchTerm().toLowerCase();
  const st = $('#msStatus') ? $('#msStatus').value : '';
  return groups.filter((g) => {
    for (const f of MS_FIELDS) {
      const sel = $(f.id);
      const v = sel ? sel.value : '';
      if (v && String(g.sample[f.key] ?? '').trim() !== v) return false;
    }
    if (st === 'Match' && !g.match) return false;
    if (st === 'Mismatch' && !g.mismatch) return false;
    if (st === 'No SAM' && !g.noSam) return false;
    if (q) {
      const hay = HIST_COLS.map((k) => String(g.sample[k] ?? ''))
        .concat(g.fileList.map((f) => f.file)).join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}
// 드롭다운 값은 (필터 전) 전체 그룹에서 뽑고, 고른 값은 유지한다.
function msFillFilters(groups) {
  for (const f of MS_FIELDS) {
    const sel = $(f.id);
    if (!sel) continue;
    const vals = [...new Set(groups.map((g) => String(g.sample[f.key] ?? '').trim())
      .filter(Boolean))].sort();
    const prev = sel.value;
    sel.innerHTML = `<option value="">${esc(t(f.allKey))}</option>` +
      vals.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
    sel.value = vals.includes(prev) ? prev : '';
    enhanceSelect(sel);
  }
  const st = $('#msStatus'); if (st) enhanceSelect(st);
}

function renderMatchSummary() {
  const table = $('#matchSumGrid');
  if (!table) return;
  const cntEl = $('#matchSumCount');
  const all = DATA.rows.length ? matchSummaryRows() : [];
  msFillFilters(all);
  const groups = msFilterGroups(all);
  if (cntEl) {
    cntEl.textContent = all.length
      ? t(groups.length === all.length ? 'match.count' : 'match.countOf',
        { n: groups.length, total: all.length })
      : '';
  }
  if (!groups.length) {
    table.querySelector('thead').innerHTML = '';
    table.querySelector('tbody').innerHTML =
      `<tr><td class="none">${esc(t(all.length ? 'match.noHit' : 'msg.noData'))}</td></tr>`;
    return;
  }
  table.querySelector('thead').innerHTML = '<tr>'
    + HIST_COLS.map((k) => `<th>${esc(colLabel(k))}</th>`).join('')
    + `<th>${esc(t('match.samFile'))}</th><th>${esc(t('match.units'))}</th>`
    + `<th>${esc(t('match.status'))}</th></tr>`;
  table.querySelector('tbody').innerHTML = groups.map((g, i) => {
    const r = g.sample;
    const cells = HIST_COLS.map((k) => {
      const v = String(r[k] ?? '').trim();
      if (k === 'Cab' && pairMismatch(r, 'cab')) {
        return `<td class="pair-diff" title="${esc(pairTitle(r, 'cab'))}">${msHl(v)}<span class="ne">≠</span></td>`;
      }
      return `<td>${msHl(v)}</td>`;
    }).join('');
    const open = MS_OPEN.has(g.key);
    const multi = g.fileList.length > 1;
    let rows = `<tr class="ms-row${multi ? ' multi' : ''}${open ? ' open' : ''}" data-i="${i}">${cells}`
      + `<td class="ms-file">${msFileCellHtml(g, open)}</td>`
      + `<td class="num">${g.n}</td><td>${msChipsHtml(g)}</td></tr>`;
    // 문서가 여럿일 때만 펼침 — 어느 견적서에 몇 대가 붙었는지 한 줄씩.
    if (open && multi) {
      rows += g.fileList.map((f) => `<tr class="ms-sub">`
        + `<td colspan="${HIST_COLS.length}"></td>`
        + `<td class="ms-file">${msFileNameHtml(f.file)}</td>`
        + `<td class="num">${f.n}</td><td>${msChipsHtml(f)}</td></tr>`).join('');
    }
    return rows;
  }).join('');
  table.querySelectorAll('.ms-row').forEach((tr) => tr.addEventListener('click', () => {
    const g = groups[Number(tr.dataset.i)];
    if (!g || g.fileList.length < 2) return;
    if (MS_OPEN.has(g.key)) MS_OPEN.delete(g.key); else MS_OPEN.add(g.key);
    renderMatchSummary();
  }));
}

// ---- 문제 탭 ----
// 고칠 것(problem, 무시 안 함)을 먼저 본다. 필터(생산월 범위·모델)로 좁히고, 체크한 줄을
// 한꺼번에 '무시'하거나 되돌린다. 무시한 줄은 고칠 것 수·탭 배지에서 빠진다.
const MP_SEL = new Set();   // 체크한 문제 key
function mpYm(p) { return String(p.ym || '').slice(0, 7); }
function mpFilterVals() {
  const v = (id) => (($(id) || {}).value || '');
  return { from: v('#mpFrom'), to: v('#mpTo'), model: v('#mpModel'), my: v('#mpMy'),
    q: v('#mpSearch').trim().toLowerCase() };
}
// 칩(종류/무시됨) 말고 드롭다운·검색만 적용 — 칩 숫자도 이 결과로 센다.
function mpApplyFilters(list) {
  const f = mpFilterVals();
  return list.filter((p) => {
    const ym = mpYm(p);
    if (f.from && (!ym || ym < f.from)) return false;
    if (f.to && (!ym || ym > f.to)) return false;
    if (f.model && p.model !== f.model) return false;
    if (f.my && p.my !== f.my) return false;
    if (f.q && ![t('mp.kind.' + p.kind), p.model, p.ym, p.my, p.file, p.why, (p.comms || []).join(' ')]
      .join(' ').toLowerCase().includes(f.q)) return false;
    return true;
  });
}
function mpFillFilters(all) {
  const months = [...new Set(all.map(mpYm).filter((x) => /^\d{4}-\d{2}$/.test(x)))].sort();
  const models = [...new Set(all.filter((p) => p.kind === 'nosam' || p.kind === 'unused')
    .map((p) => p.model).filter(Boolean))].sort();
  const fill = (id, allKey, vals) => {
    const sel = $(id);
    if (!sel) return;
    const prev = sel.value;
    sel.innerHTML = `<option value="">${esc(t(allKey))}</option>`
      + vals.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
    sel.value = vals.includes(prev) ? prev : '';
    enhanceSelect(sel);
  };
  fill('#mpFrom', 'mp.f.from', months);
  fill('#mpTo', 'mp.f.to', months);
  fill('#mpModel', 'filter.allModel', models);
  fill('#mpMy', 'filter.allMY', [...new Set(all.map((p) => p.my).filter(Boolean))].sort());
}

function mpIsTodo(p) { return p.sev === 'problem' && !p.ignored; }
function mpChipKeep(p) {
  switch (MP_FILTER) {
    case 'all': return true;
    case 'problem': return mpIsTodo(p);
    case 'ignored': return p.ignored;
    default: return p.kind === MP_FILTER && !p.ignored;
  }
}

let MP_SHOWN = [];
function renderMatchProblems() {
  const table = $('#matchProbGrid');
  if (!table) return;
  const all = (DATA.rows.length || DATA.diagnostics) ? matchProblems() : [];
  for (const p of all) p.ignored = !!miItem(p);

  const badge = $('#matchPaneBadge');
  const nTodoAll = all.filter(mpIsTodo).length;
  if (badge) { badge.textContent = nTodoAll ? fmtNum(nTodoAll) : ''; badge.classList.toggle('hidden', !nTodoAll); }

  mpFillFilters(all);
  const base = mpApplyFilters(all);
  const live = base.filter((p) => !p.ignored);
  const nTodo = base.filter(mpIsTodo).length;
  const units = live.filter((p) => p.kind === 'nosam').reduce((a, p) => a + p.n, 0);
  const cnt = (k) => live.filter((p) => p.kind === k).length;

  const chip = (k, label, n, cls) => `<button type="button" class="mp-chip ${cls || ''}${MP_FILTER === k ? ' active' : ''}" data-mp="${k}">`
    + `<b>${esc(fmtNum(n))}</b> ${esc(label)}</button>`;
  const chips = $('#mpChips');
  if (chips) {
    chips.innerHTML = chip('problem', t('mp.f.problem'), nTodo, nTodo ? 'bad' : 'ok')
      + chip('nosam', t('mp.f.nosam', { units: fmtNum(units) }), cnt('nosam'))
      + chip('unused', t('mp.f.unused'), cnt('unused'))
      + chip('skipped', t('mp.f.skipped'), cnt('skipped'))
      + chip('month', t('mp.f.month'), cnt('month'))
      + chip('ignored', t('mp.f.ignored'), base.filter((p) => p.ignored).length, 'muted')
      + chip('all', t('mp.f.all'), base.length);
  }

  const shown = base.filter(mpChipKeep);
  MP_SHOWN = shown;
  // 화면에서 사라진 줄의 체크는 버린다 — 안 보이는 줄이 같이 무시되면 안 된다.
  const visible = new Set(shown.map((p) => p.key));
  for (const k of [...MP_SEL]) if (!visible.has(k)) MP_SEL.delete(k);

  const c = $('#matchProbCount');
  if (c) c.textContent = t('match.diagCount', { n: fmtNum(shown.length), total: fmtNum(all.length) });
  mpSyncBulk();

  if (!all.length) {
    table.querySelector('thead').innerHTML = '';
    table.querySelector('tbody').innerHTML = `<tr><td class="none">${esc(t(DATA.diagnostics ? 'mp.none' : 'match.diagOld'))}</td></tr>`;
    return;
  }
  const allOn = shown.length > 0 && shown.every((p) => MP_SEL.has(p.key));
  table.querySelector('thead').innerHTML = '<tr>'
    + `<th class="mp-ck"><input type="checkbox" id="mpAll" aria-label="${esc(t('mp.selAll'))}"${allOn ? ' checked' : ''} /></th>`
    + ['mp.col.kind', 'mp.col.model', 'Cab', 'PTO', 'mp.col.month', 'MY', 'mp.col.n', 'match.samFile', 'match.diagWhy', '']
      .map((k) => `<th>${k ? esc(k.indexOf('.') !== -1 ? t(k) : k) : ''}</th>`).join('') + '</tr>';
  table.querySelector('tbody').innerHTML = shown.length
    ? shown.map((p) => {
      const first = p.comms && p.comms[0];
      const units2 = p.kind === 'nosam' && first
        ? `<button type="button" class="link-btn" data-open="${esc(first)}" title="${esc(p.comms.join(', '))}">${esc(fmtNum(p.n))}</button>`
        : esc(String(p.n));
      const act = p.kind === 'nosam' && IS_ADMIN && !p.ignored
        ? `<button type="button" class="btn xs" data-add-rule="${esc(p.model)}">${esc(t('mp.addRule'))}</button>` : '';
      const ig = p.ignored ? miItem(p) : null;
      const why = esc(p.why) + (ig ? `<div class="mp-ig">${esc(t('mp.ignoredBy', { who: excWho(ig.by_name, ig.at) }))}`
        + (ig.note ? ` — “${esc(ig.note)}”` : '') + `</div>` : '');
      const canSel = p.kind !== 'month';
      return `<tr class="mp-${p.sev}${p.ignored ? ' mp-ignored' : ''}">`
        + `<td class="mp-ck">${canSel ? `<input type="checkbox" data-sel="${esc(p.key)}"${MP_SEL.has(p.key) ? ' checked' : ''} />` : ''}</td>`
        + `<td class="nowrap"><span class="mp-kind ${p.kind}">${esc(t('mp.kind.' + p.kind))}</span></td>`
        + `<td${p.modelTitle ? ` title="${esc(p.modelTitle)}" class="mp-keys"` : ''}>${esc(p.model)}</td><td>${esc(p.cab)}</td><td>${esc(p.pto)}</td>`
        + `<td class="nowrap">${esc(p.ym || '—')}</td><td class="nowrap">${esc(p.my || '—')}</td><td class="num">${units2}</td>`
        + `<td class="ms-file"><span title="${esc(p.file)}">${esc(p.file)}</span></td>`
        + `<td class="mp-why">${why}</td><td class="nowrap">${act}</td></tr>`;
    }).join('')
    : `<tr><td class="none" colspan="11">${esc(t(MP_FILTER === 'problem' ? 'mp.allGood' : 'match.noHit'))}</td></tr>`;
}

// 선택 막대 — 몇 개 골랐는지, 무시/해제 가능 여부.
function mpSyncBulk() {
  const sel = MP_SHOWN.filter((p) => MP_SEL.has(p.key));
  const nIg = sel.filter((p) => p.ignored).length;
  const n = $('#mpSelCount');
  if (n) n.textContent = sel.length ? t('mp.selected', { n: fmtNum(sel.length) }) : '';
  const bi = $('#mpIgnore'), bu = $('#mpUnignore');
  if (bi) bi.disabled = !IS_ADMIN || sel.length - nIg === 0;
  if (bu) bu.disabled = !IS_ADMIN || nIg === 0;
}

let MP_BUSY = false;
async function mpBulk(action) {
  if (MP_BUSY || !IS_ADMIN) return;
  const sel = MP_SHOWN.filter((p) => MP_SEL.has(p.key) && (action === 'ignore' ? !p.ignored : p.ignored));
  if (!sel.length) return;
  let note = '';
  if (action === 'ignore') {
    note = prompt(t('mp.ask.ignore', { n: sel.length }), t('mp.ask.ignoreDefault'));
    if (note === null) return;
    note = note.trim();
  } else if (!confirm(t('mp.ask.unignore', { n: sel.length }))) return;
  MP_BUSY = true;
  document.body.classList.add('exc-busy');
  try {
    await miMutate(action, sel, note);
    MP_SEL.clear();
  } catch (e) {
    console.error(e);
    alert(t(e.status === 403 || e.status === 401 ? 'exc.err.perm' : 'exc.err.save', { err: e.message }));
  } finally {
    MP_BUSY = false;
    document.body.classList.remove('exc-busy');
    renderMatchProblems();
  }
}

// 상태 뱃지(Match / Mismatch / No SAM 대수) — 모델 줄과 문서 줄이 같이 쓴다.
function msChipsHtml(o) {
  return [
    o.match ? `<span class="status Match">${o.match}</span>` : '',
    o.mismatch ? `<span class="status Mismatch">${o.mismatch}</span>` : '',
    o.noSam ? `<span class="status NoSAM">${o.noSam}</span>` : '',
  ].join(' ');
}
function msFileNameHtml(file) {
  return file
    ? `<span title="${esc(file)}">${msHl(file)}</span>`
    : `<span class="none">${esc(t('match.noFile'))}</span>`;
}
// 문서가 하나면 그대로, 여럿이면 '문서 N개'로 접는다(클릭하면 펼침).
function msFileCellHtml(g, open) {
  if (g.fileList.length <= 1) return msFileNameHtml(g.fileList.length ? g.fileList[0].file : '');
  const names = g.fileList.map((f) => `${f.file || t('match.noFile')} (${f.n})`).join('\n');
  return `<button type="button" class="ms-more" title="${esc(names)}">`
    + `<span class="caret">${open ? '▾' : '▸'}</span>`
    + `${esc(t('match.files', { n: g.fileList.length }))}</button>`;
}

// ---- 코드 관리: 04. code 폴더의 모든 xlsx ----
const codeEditor = makeSheetEditor({
  folderKey: 'code',
  els: { tabs: 'sheetTabs', grid: 'editGrid', addRow: 'codeAddRow', revert: 'codeReloadFile',
         saveSp: 'codeSaveSp', download: 'codeDownload', dirty: 'codeDirty',
         status: 'codeStatus', search: 'codeSearch', main: 'codeMain', empty: 'codeEmpty',
         recs: 'codeRecs', recWrap: 'codeRecWrap', gridWrap: 'codeGridWrap', pager: 'codePager',
         mode: 'codeMode', filter: 'codeFilter', count: 'codeCount', addRec: 'codeAddRec' },
  onSaved: () => loadCodeFileList(),
});
let CODE_FILES = [];
async function initCodes() {
  const link = $('#codeFolderLink'); if (window.Graph) link.href = Graph.folders.code.shareUrl;
  codeEditor.wire();
  $('#codeRefresh').addEventListener('click', loadCodeFileList);
  await loadCodeFileList();
}
async function loadCodeFileList() {
  const box = $('#codeFileList');
  const markActive = (name) =>
    box.querySelectorAll('.file-item').forEach((x) => x.classList.toggle('active', x.dataset.file === name));
  if (!window.Graph || !Graph.available()) { box.innerHTML = `<span class="none">${esc(t('op.needLogin'))}</span>`; return; }
  box.innerHTML = `<span class="none">${esc(t('codes.pickFile'))}</span>`;
  try {
    // PTO 코드 목록이 아직 없으면 여기서 만들어 준다 — 검토하러 온 화면이 바로 이 화면이라,
    // 다시 계산(1~3분)을 기다렸다 오게 할 이유가 없다. 있으면 손대지 않는다.
    try { await ensurePtoCodesWorkbook(); } catch (e) { console.error(e); }
    const items = await Graph.list('code');
    // code-overview.xlsx 는 코드 설명용 문서라 비교 로직이 참조하지 않는다 — 목록에서 뺀다.
    CODE_FILES = items.filter((i) => !i.isFolder && /\.xlsx?$/i.test(i.name)
      && !/^code-overview\.xlsx?$/i.test(i.name));
    if (!CODE_FILES.length) { box.innerHTML = `<span class="none">${esc(t('codes.noXlsx'))}</span>`; return; }
    box.innerHTML = CODE_FILES.map((f) => {
      const kb = f.size ? Math.max(1, Math.round(f.size / 1024)) + ' KB' : '';
      return `<button class="file-item${f.name === codeEditor.S.file ? ' active' : ''}" data-file="${esc(f.name)}"><span class="fi-name">📄 ${esc(f.name)}</span><span class="fi-meta">${esc(kb)}</span></button>`;
    }).join('');
    box.querySelectorAll('.file-item').forEach((b) => b.addEventListener('click', () => {
      if (codeEditor.S.dirty && b.dataset.file !== codeEditor.S.file && !confirm(t('codes.confirmLeave'))) return;
      localStorage.setItem('codeFile', b.dataset.file);
      codeEditor.open(b.dataset.file).then(() => markActive(b.dataset.file));
    }));
    // 모델 매칭처럼 바로 편집 화면이 보이도록 — 마지막에 보던 파일(없으면 첫 파일)을 자동으로 연다.
    if (!codeEditor.S.file) {
      const last = localStorage.getItem('codeFile');
      const pick = CODE_FILES.some((f) => f.name === last) ? last : CODE_FILES[0].name;
      await codeEditor.open(pick);
      markActive(pick);
    }
  } catch (e) { box.innerHTML = `<span class="none err">${esc(t('codes.listFail', { err: e.message }))}</span>`; }
}

// ====================== 데이터 빌드 (브라우저에서 직접 계산) ======================
// GitHub Actions·PAT·앱 전용 시크릿 없이, 로그인한 관리자의 위임 권한(Graph)으로
// SharePoint 원본을 읽어 이 브라우저에서 파이프라인(docs/lib/*)을 돌린다.
//   01. SAM_files(최신 생산월) · 02. WINGS_data(최신 파일) · 03. model_rules · 04. code
//     → data.json / codes.json → 05. output 에 저장
// 다른 사용자는 05. output 만 읽으므로 재계산이 필요 없다.

// Graph 를 pipeline.js 의 소스 어댑터 형태로 감싼다.
const GRAPH_SOURCE = {
  list: (key) => Graph.list(key),
  download: (key, name) => Graph.download(key, name),
  // SAM 추출 결과 캐시(05. output/sam-cache.json) — 바뀌지 않은 SAM 파일은 다시 받지 않는다.
  readCache: (name) => Graph.downloadJson('output', name),
  writeCache: async (name, obj) => {
    await Graph.ensureFolder('output');
    await Graph.uploadJson('output', name, obj);
  },
  optionCodes: (function () {
    let cache = null;
    return async function () {
      if (!cache) cache = await fetch('lib/optioncodes.json?_=' + Date.now()).then((r) => r.json());
      return cache;
    };
  })(),
};

// 빌드 로그 패널 (진행 상황 표시)
function buildLog() {
  let box = $('#buildLog');
  if (!box) {
    box = document.createElement('div');
    box.id = 'buildLog';
    box.className = 'build-log';
    box.innerHTML = `<div class="bl-head"><b>${esc(t('build.title'))}</b>`
      + `<button class="bl-close" type="button">×</button></div><div class="bl-body"></div>`;
    document.body.appendChild(box);
    box.querySelector('.bl-close').addEventListener('click', () => box.remove());
  }
  const body = box.querySelector('.bl-body');
  return {
    line(msg, cls) {
      const d = document.createElement('div');
      if (cls) d.className = cls;
      d.textContent = msg;
      body.appendChild(d);
      body.scrollTop = body.scrollHeight;
    },
    close() { box.remove(); },
  };
}

// ====================== Mismatch 예외 관리 화면 ======================
// 대시보드에서는 문제없음 처리된 차이가 사라지므로, 무엇이 왜 Match 가 됐는지는 여기서 본다.
// 탭: 문제없음 처리 / 해제·미반영 / 변경 이력(전체 log).
let EXC_TAB = 'approved';
const EXC_TAB_FILTER = {
  approved: (it) => it.status === 'approved',
  closed: (it) => it.status !== 'approved',
};

function initExceptions() {
  document.querySelectorAll('#excTabs button[data-tab]').forEach((b) =>
    b.addEventListener('click', () => { EXC_TAB = b.dataset.tab; excRender(); }));
  $('#excSearch').addEventListener('input', excRender);
  $('#excReload').addEventListener('click', async () => {
    const b = $('#excReload');
    b.disabled = true;
    try { await loadExceptions(); } finally { b.disabled = false; }
    excRefreshAll();
  });
  $('#excExport').addEventListener('click', excExport);
  $('#excBody').addEventListener('click', (e) => {
    const b = e.target.closest('button[data-exc]');
    if (b) { e.stopPropagation(); excAction(b.dataset.exc, [excTargetOf(b)]); return; }
    const c = e.target.closest('button[data-open]');
    if (c) {
      const r = rowByComm(c.dataset.open);
      if (r) openDrawer(r);
    }
  });
  excRender();
}

// 차량이 지금 데이터에서 어떤 상태인지 — 예외의 효과를 바로 보이게 한다.
function excVehicleState(it) {
  const r = rowByComm(it.commission);
  if (!r) return { cls: 'gone', text: t('exc.veh.gone') };
  if (r._orig_status !== 'Mismatch') return { cls: 'na', text: t('exc.veh.nowMatch') };
  const stillDiff = rawDiffCodes(r, it.side).includes(it.code);
  if (!stillDiff) return { cls: 'na', text: t('exc.veh.codeGone') };
  if (r._accepted) return { cls: 'ok', text: t('exc.veh.match') };
  const left = excRowCodes(r).length - (r._accN || 0);
  return { cls: 'miss', text: t('exc.veh.left', { n: left }) };
}

function excNoteOf(it) {
  if (it.status === 'approved') return it.note || '';
  if (it.status === 'revoked') return it.revoke_note || it.note || '';
  return it.reason || it.note || '';
}

function excItemsFor(tab) {
  const keep = EXC_TAB_FILTER[tab] || (() => false);
  const q = ($('#excSearch').value || '').trim().toLowerCase();
  return Object.values(EXC.items)
    .filter(keep)
    .filter((it) => !q || [it.commission, it.code, it.vehicle, it.type, excNoteOf(it), it.note,
      excLastBy(it).name, describe(it.code)].join(' ').toLowerCase().includes(q))
    .sort((a, b) => String(excLastBy(b).at).localeCompare(String(excLastBy(a).at)));
}
function excLogFor() {
  const q = ($('#excSearch').value || '').trim().toLowerCase();
  return EXC.log.slice().reverse()
    .filter((l) => !q || [l.commission, l.code, l.by_name, l.by, l.text, t('exc.act.' + l.action)]
      .join(' ').toLowerCase().includes(q));
}

function excNoticeHtml() {
  if (EXC_STATE === 'error') return `<div class="note warn">${esc(t('exc.notice.error', { err: EXC_ERR }))}</div>`;
  if (EXC_STATE === 'local') return `<div class="note">${esc(t('exc.notice.local'))}</div>`;
  return '';
}

function excRender() {
  if (!VIEW_INIT.exceptions) return;
  const all = Object.values(EXC.items);
  const counts = {
    approved: all.filter(EXC_TAB_FILTER.approved).length,
    closed: all.filter(EXC_TAB_FILTER.closed).length,
    log: EXC.log.length,
  };
  document.querySelectorAll('#excTabs button[data-tab]').forEach((b) => {
    const on = b.dataset.tab === EXC_TAB;
    b.classList.toggle('active', on);
    b.setAttribute('aria-pressed', on ? 'true' : 'false');
    const c = b.querySelector('.cnt');
    if (c) c.textContent = fmtNum(counts[b.dataset.tab] || 0);
  });
  $('#excNotice').innerHTML = excNoticeHtml();
  $('#excUpdated').textContent = EXC.updated_at ? t('exc.updated', { when: fmtKst(EXC.updated_at) }) : '';

  const body = $('#excBody');
  const empty = $('#excEmpty');
  let html = '';
  let n = 0;
  if (EXC_TAB === 'log') {
    const logs = excLogFor();
    n = logs.length;
    html = `<table class="exc-table"><thead><tr>`
      + ['exc.col.at', 'exc.col.by', 'exc.col.action', 'exc.col.comm', 'exc.col.side', 'exc.col.code', 'exc.col.note']
        .map((k) => `<th>${esc(t(k))}</th>`).join('')
      + `</tr></thead><tbody>`
      + logs.map((l) => `<tr>
          <td class="mono nowrap">${esc(fmtKst(l.at))}</td>
          <td title="${esc(l.by)}">${esc(l.by_name || l.by)}</td>
          <td><span class="exc-st ${esc(EXC_STATUS_OF[l.action] || '')}">${esc(t('exc.act.' + l.action))}</span></td>
          <td>${excCommHtml(l.commission)}</td>
          <td>${esc(excSideLabel(l.side))}</td>
          <td><span class="cchip ${l.side === 'sam' ? 'crit' : 'accent'}">${esc(l.code)}</span></td>
          <td class="wrap">${esc(l.text || '')}</td>
        </tr>`).join('')
      + `</tbody></table>`;
  } else {
    const items = excItemsFor(EXC_TAB);
    n = items.length;
    html = `<table class="exc-table"><thead><tr>`
      + ['exc.col.comm', 'exc.col.model', 'exc.col.prod', 'exc.col.side', 'exc.col.code', 'exc.col.desc',
        'exc.col.status', 'exc.col.note', 'exc.col.by', 'exc.col.vehicle', '']
        .map((k) => `<th>${k ? esc(t(k)) : ''}</th>`).join('')
      + `</tr></thead><tbody>`
      + items.map((it) => {
        const v = excVehicleState(it);
        const w = excLastBy(it);
        const model = [it.vehicle, it.type, it.axle].filter(Boolean).join(' ');
        return `<tr>
          <td>${excCommHtml(it.commission)}</td>
          <td class="nowrap">${esc(model || '—')}</td>
          <td class="mono nowrap">${esc(String(it.production || '').slice(0, 10) || '—')}</td>
          <td class="nowrap">${esc(excSideLabel(it.side))}</td>
          <td><span class="cchip ${it.side === 'sam' ? 'crit' : 'accent'}">${esc(it.code)}</span></td>
          <td class="wrap">${esc(describe(it.code) || '—')}</td>
          <td>${excStateHtml(it)}</td>
          <td class="wrap">${esc(excNoteOf(it))}</td>
          <td class="nowrap">${esc(w.name)}<div class="sub mono">${esc(fmtKst(w.at))}</div></td>
          <td><span class="exc-veh ${v.cls}">${esc(v.text)}</span></td>
          <td class="exc-acts nowrap">${excButtonsHtml(it, it.commission, it.side, it.code)}</td>
        </tr>`;
      }).join('')
      + `</tbody></table>`;
  }
  $('#excCount').textContent = tn('exc.count', n, { n: fmtNum(n) });
  body.innerHTML = n ? html : '';
  empty.textContent = n ? '' : t('exc.empty.' + EXC_TAB);
  empty.classList.toggle('hidden', !!n);
}

function excCommHtml(comm) {
  return rowByComm(comm)
    ? `<button type="button" class="comm-link" data-open="${esc(comm)}">${esc(comm)}</button>`
    : `<span class="mono">${esc(comm)}</span>`;
}

// 지금 탭(검색 포함)을 엑셀로 — 감사·보고용.
function excExport() {
  const isLog = EXC_TAB === 'log';
  const H = (v) => ({ v, s: 2 });
  const C = (v) => ({ v: v == null ? '' : String(v), s: 3 });
  const rows = [];
  if (isLog) {
    rows.push(['Time (KST)', 'User', 'User ID', 'Action', 'Commission', 'Side', 'Code', 'Description', 'Note'].map(H));
    excLogFor().forEach((l) => rows.push([fmtKst(l.at), l.by_name, l.by, t('exc.act.' + l.action), l.commission,
      excSideLabel(l.side), l.code, describe(l.code), l.text].map(C)));
  } else {
    rows.push(['Commission', 'Model', 'Production', 'Side', 'Code', 'Description', 'Status',
      'Note', 'By', 'At (KST)', 'Vehicle now'].map(H));
    excItemsFor(EXC_TAB).forEach((it) => {
      const w = excLastBy(it);
      rows.push([it.commission, [it.vehicle, it.type, it.axle].filter(Boolean).join(' '),
        String(it.production || '').slice(0, 10), excSideLabel(it.side), it.code, describe(it.code),
        t('exc.st.' + excStatusKey(it)), excNoteOf(it), w.name, fmtKst(w.at), excVehicleState(it).text].map(C));
    });
  }
  const widths = isLog ? [18, 18, 28, 14, 14, 12, 10, 40, 50] : [14, 20, 12, 12, 10, 40, 14, 50, 18, 18, 22];
  const cols = widths.map((w, i) => ({ min: i + 1, max: i + 1, w }));
  writeXlsx(`mismatch_exceptions_${EXC_TAB}_${CUR_DATE}.xlsx`, 'Exceptions', rows, [], cols);
}

let BUILDING = false;
async function runBuild() {
  if (BUILDING) return;
  if (!IS_ADMIN) { alert(t('role.denied')); return; }
  if (!window.Graph || !Graph.available()) { alert(t('build.needLogin')); return; }
  if (!window.Pipeline) { alert(t('build.fail') + ' pipeline.js'); return; }
  if (!window.XLSX) { alert(t('alert.xlsxBlocked')); return; }
  if (!confirm(t('build.confirm'))) return;

  BUILDING = true;
  const btn = $('#buildBtn');
  const lbl = btn ? btn.querySelector('.lbl') : null;
  if (btn) { btn.disabled = true; btn.classList.add('busy'); btn.setAttribute('aria-busy', 'true'); }
  if (lbl) lbl.textContent = t('build.running');
  const log = buildLog();
  log.line(t('build.step.ref'));
  try {
    // 예전 규칙 워크북이면 먼저 '모델명변경' 한 장으로 정리한다 — 지금 데이터(예전 규칙의 결과)로
    // 필요한 모델명 변경을 찾아야 하므로 반드시 빌드 전에.
    log.line(t('build.step.mapping'));
    try {
      const mm = await migrateModelMapping();
      log.line(mm.changed ? t('build.mapping.done', { file: mm.name }) + ' ' + mm.msg : t('build.mapping.same', { file: mm.name }), 'ok');
    } catch (e) {
      console.error(e);
      log.line(t('build.mapping.fail', { err: e.message }), 'err');
    }
    const { data, codes } = await Pipeline.build(GRAPH_SOURCE, { log: (m) => log.line(m) });

    log.line(t('build.step.upload'));
    await Graph.ensureFolder('output');
    await Graph.uploadJson('output', 'data.json', data);
    await Graph.uploadJson('output', 'codes.json', codes);

    DATA_SOURCE = 'sharepoint';
    applyData(data, codes);        // MY 등 파생값을 채운 뒤라야 대조표를 만들 수 있다

    // 모델 매핑 워크북도 이번 결과로 갱신 — 인식 결과(대조표)와 미매칭 진단이
    // 항상 지금 데이터와 같은 것을 가리키게 한다. 실패해도 빌드는 성공으로 둔다.
    try {
      const p = await ensurePtoCodesWorkbook();
      if (p) log.line(t('build.pto.created', { file: p.name, n: p.codes, ex: p.excepts }), 'ok');
    } catch (e) {
      console.error(e);
      log.line(t('build.pto.fail', { err: e.message }), 'err');
    }

    if (matchEditor.S.file && !matchEditor.S.dirty) matchEditor.open(MODEL_MAPPING_FILE);

    const msg = t('build.done', {
      rows: data.summary.total, match: data.summary.matched, mismatch: data.summary.mismatched });
    log.line(msg, 'ok');
  } catch (e) {
    console.error(e);
    log.line(t('build.fail') + ' ' + e.message, 'err');
  } finally {
    BUILDING = false;
    if (btn) { btn.disabled = false; btn.classList.remove('busy'); btn.removeAttribute('aria-busy'); }
    if (lbl) lbl.textContent = t('build.btn');
  }
}

// ====================== 이벤트 & 초기화 ======================
$('#drawerClose').addEventListener('click', closeDrawer);
$('#backdrop').addEventListener('click', closeDrawer);
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') closeDrawer();
  else trapDrawerFocus(e);
});
$('#drawerXls').addEventListener('click', () => { if (DRAWER_ROW) exportRowXls(DRAWER_ROW); });
document.querySelectorAll('.lang-toggle button[data-lang]').forEach((b) =>
  b.addEventListener('click', () => setLang(b.dataset.lang)));

// 뷰 전환 (네비 + 브랜드)
document.querySelectorAll('[data-view]').forEach((el) =>
  el.addEventListener('click', (e) => { e.preventDefault(); switchView(el.dataset.view); }));

// 데이터 빌드 (관리자) — 이 브라우저에서 계산 후 SharePoint 에 저장
const _buildBtn = $('#buildBtn');
if (_buildBtn) _buildBtn.addEventListener('click', runBuild);

// Export — 대시보드의 현재 필터/정렬 결과를 엑셀로 저장
const _exportBtn = $('#exportBtn');
if (_exportBtn) _exportBtn.addEventListener('click', exportFilteredXlsx);

// 저장 안 한 편집이 있으면 페이지 이탈 경고 (모델 매칭 / 코드 관리)
window.addEventListener('beforeunload', (e) => {
  if (codeEditor.S.dirty || matchEditor.S.dirty) { e.preventDefault(); e.returnValue = ''; }
});

function onManualFilter() {
  tileDate = null;
  restrictSoon = false;
  tileMandatory = false;
  tileSamUpdate = false;
  activeTile = null;
  syncTileActive();
  render();
}
// 드롭다운/상태/체크박스 변경: 연동 필터(다른 드롭다운의 선택가능 옵션)를 다시 계산.
function onFilterChange() {
  fillFilters();
  onManualFilter();
}
['#statusFilter', '#productionFilter', '#myFilter', '#modelFilter',
  '#typeFilter', '#axleFilter', '#cabFilter',
  '#upcomingOnly'].forEach((s) =>
  $(s).addEventListener('input', onFilterChange));
$('#search').addEventListener('input', onManualFilter);

// 시작 생산월 — 이 달부터 집계·표시한다. 과거로 내리면 지난 차량도 함께 비교된다.
const _fromMonth = $('#fromMonth');
if (_fromMonth) _fromMonth.addEventListener('input', () => setDashFrom(_fromMonth.value));

// 문구 사전을 먼저 읽고(실패해도 HTML 의 한국어가 남는다) 데이터를 불러온다.
async function boot() {
  await loadI18n();
  applyStaticI18n();
  enhanceFilterSelects();
  try {
    await load();
  } catch (e) {
    $('#meta').textContent = t('meta.loadFail', { err: e.message });
    $('#grid tbody').innerHTML = '';
    const msg = $('#statusMsg');
    msg.textContent = t('msg.loadFail', { err: e.message });
    msg.classList.remove('hidden');
  }
}
boot();
