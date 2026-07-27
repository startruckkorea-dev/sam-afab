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
  { key: 'Cab', label: 'Cab' },
  { key: 'MY', label: 'MY' },
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
  for (const c of splitCodes(r && r['_all_sam_codes'])) {
    if (c in MY_CODE_DIGIT) return String(2020 + MY_CODE_DIGIT[c]);
  }
  return '';
}

const _now = new Date();
const CUR_MONTH = _now.getFullYear() + '-' + String(_now.getMonth() + 1).padStart(2, '0');
const CUR_DATE = CUR_MONTH + '-' + String(_now.getDate()).padStart(2, '0');

let DATA = { rows: [] };
let CODES = { options: {}, mandatory: {} };
let sortKey = null, sortDir = 1;
let restrictSoon = false;
let tileMandatory = false;
let activeTile = 't-total';

const $ = (s) => document.querySelector(s);

// ====================== i18n (한국어 / English) ======================
const I18N = {
  ko: {
    'nav.dashboard': '대시보드',
    'nav.matching': '모델 매칭',
    'nav.codes': '코드 관리',
    'meta.loading': '불러오는 중…',
    'meta.generated': '생성: {when}  ·  WINGS: {file}',
    'search.ph': 'Commission no. / 모델 / 코드 검색…',
    'filter.allStatus': '전체 상태',
    'filter.allProd': '전체 생산월',
    'filter.allMY': '전체 MY',
    'filter.allModel': '전체 Model',
    'filter.allType': '전체 Type',
    'filter.allAxle': '전체 Axle',
    'filter.allCab': '전체 Cab',
    'filter.prod.title': '생산월(Requested delivery) 선택',
    'chk.upcoming': 'Display only production from this month',
    'chk.upcoming.title': 'Changeability 가 이번 달 이후인 항목만',
    'count': '{n} / {total} 건',
    'dash.overall': '전체 현황 (이번 생산월 이후)',
    'dash.soon': '2주 이내 (Changeability D-14)',
    'tile.total': '전체 Commission',
    'tile.mismatch': '미스매치',
    'tile.match': '매칭',
    'tile.mand': 'Mandatory 누락',
    'dash.models': '모델별 대수',
    'dash.models.kinds': '종',
    'dash.near': '오늘과 가장 가까운 Changeability',
    'dash.near.count': '해당일 Commission',
    'unit.ea': '대',
    'unit.case': '건',
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
    // 모델 매칭
    'matching.title': '모델 매칭',
    'matching.sub': 'SAM ↔ WINGS 모델 인식 규칙 — 모두 03. model_rules/model_mapping.xlsx 에 저장',
    'matching.loading': 'model_mapping.xlsx 를 불러오는 중…',
    'matching.note':
      '모델 매칭과 관련된 모든 규칙(정규화·이전/현재 모델·차종 키워드·매칭 별칭·수동매핑·옵션)이 ' +
      '이 워크북의 시트로 관리됩니다. 시트 탭을 바꿔 편집한 뒤 <b>SharePoint에 저장</b>하면 ' +
      '<code>model_mapping.xlsx</code> 에 반영되고, 다음 <b>데이터 빌드</b> 때 적용됩니다. ' +
      '<code>인식모델_대조표</code> 시트는 로컬 도구가 만드는 확인용 보기입니다(브라우저 빌드는 이 시트를 갱신하지 않음).',
    // 코드 관리
    'codes.title': '코드 관리',
    'codes.sub': "SharePoint 04. code 폴더의 Excel 파일을 웹에서 직접 편집·저장",
    'codes.pickFile': '파일 목록을 불러오는 중…',
    'codes.empty': '왼쪽에서 편집할 Excel 파일을 선택하세요.',
    'codes.searchPh': '이 시트에서 검색…',
    'codes.unsaved': '● 저장되지 않은 변경',
    'codes.loadingFile': '파일 여는 중…',
    'codes.saving': 'SharePoint에 저장 중…',
    'codes.saved': '저장 완료 — SharePoint에 반영되었습니다.',
    'codes.listFail': '파일 목록을 불러오지 못했습니다: {err}',
    'codes.noXlsx': '이 폴더에 Excel(.xlsx) 파일이 없습니다.',
    'codes.confirmLeave': '저장하지 않은 변경이 있습니다. 이동하면 사라집니다. 계속할까요?',
    'codes.delRow': '이 행을 삭제할까요?',
    // 버튼
    'btn.openFolder': '📂 SharePoint 폴더 열기',
    'btn.refresh': '↻ 목록 새로고침',
    'btn.loadSp': '⭳ SharePoint에서 불러오기',
    'btn.saveSp': '⭱ SharePoint에 저장',
    'btn.addRow': '＋ 행 추가',
    'btn.addRec': '＋ 등록',
    'btn.reloadFile': '↺ 되돌리기',
    'btn.download': '⬇ 다운로드',
    'btn.delete': '삭제',
    'btn.cancel': '취소',
    'btn.apply': '저장',
    'btn.edit': '편집',
    // 레코드(카드) 편집
    'mode.record': '🗂 카드 편집',
    'mode.grid': '▦ 표 편집',
    'rec.actions': '액션',
    'rec.count': '{n} / {total} 건',
    'rec.empty': '표시할 항목이 없습니다.',
    'rec.emptySheet': '— 빈 시트 —',
    'rec.new': '새 항목 등록',
    'rec.edit': '항목 편집',
    'rec.page': '{page} / {pages} 페이지',
    'rec.prev': '‹ 이전',
    'rec.next': '다음 ›',
    'rec.allOf': '전체 · {col}',
    'rec.blank': '(비어 있음)',
    'rec.on': '예',
    'rec.off': '아니오',
    'rec.colFallback': '열 {n}',
    'rec.delConfirm': '이 항목을 삭제할까요?',
    'rec.hint': '행을 클릭하면 편집 창이 열립니다. 열(컬럼) 자체를 바꾸려면 표 편집으로 전환하세요.',
    'modal.upload': '⬆ 엑셀 불러오기(로컬)',
    'modal.upload.title': 'PC의 model_mapping.xlsx 불러오기',
    'modal.download': '⬇ 엑셀 저장(로컬)',
    'op.loading': '처리 중…',
    'op.needLogin': 'SharePoint 연동은 회사 계정 로그인 후 사용할 수 있습니다.',
    'op.loaded': '불러왔습니다.',
    'op.saved': '저장 완료 — SharePoint에 반영되었습니다.',
    'op.fail': '실패: {err}',
    'nav.build': '⟳ 데이터 빌드',
    'build.btn': '⟳ 데이터 빌드',
    'build.running': '⟳ 빌드 중…',
    'build.title': '데이터 빌드',
    'build.confirm': 'SharePoint 의 최신 WINGS(02) 와 최신 생산월 SAM(01) 으로 비교를 다시 계산합니다.\n'
      + '이 브라우저에서 1~3분 정도 걸리고, 결과는 SharePoint(05. output)에 저장돼 모든 사용자에게 반영됩니다.\n\n계속할까요?',
    'build.needLogin': '데이터 빌드는 회사 Microsoft 365 계정 로그인이 필요합니다.',
    'build.step.ref': '참조 워크북(코드·매칭 규칙) 읽는 중…',
    'build.step.upload': '결과를 SharePoint 에 저장하는 중…',
    'build.done': '빌드 완료 — {rows}행 (일치 {match} / 불일치 {mismatch}). SharePoint 에 저장했습니다.',
    'build.fail': '빌드 실패:',
    'build.close': '닫기',
    'meta.srcSp': '출처: SharePoint 최신 빌드',
    'meta.srcBundled': '출처: 배포 사본(빌드 결과 없음)',
    'alert.xlsxBlocked':
      '엑셀 라이브러리를 불러오지 못했습니다(사내망 차단 가능). 잠시 후 다시 시도하세요.',
    'alert.xlsxReadFail': '엑셀 읽기 실패: ',
    'rules.loadFail': 'rules.json 로드 실패: ',
  },
  en: {
    'nav.dashboard': 'Dashboard',
    'nav.matching': 'Model Matching',
    'nav.codes': 'Code Manager',
    'meta.loading': 'Loading…',
    'meta.generated': 'Generated: {when}  ·  WINGS: {file}',
    'search.ph': 'Search commission no. / model / code…',
    'filter.allStatus': 'All statuses',
    'filter.allProd': 'All production months',
    'filter.allMY': 'All MY',
    'filter.allModel': 'All models',
    'filter.allType': 'All types',
    'filter.allAxle': 'All axles',
    'filter.allCab': 'All cabs',
    'filter.prod.title': 'Select production month (Requested delivery)',
    'chk.upcoming': 'Display only production from this month',
    'chk.upcoming.title': 'Only rows whose Changeability is this month or later',
    'count': '{n} / {total} rows',
    'dash.overall': 'Overall (this production month onward)',
    'dash.soon': 'Within 2 weeks (Changeability D-14)',
    'tile.total': 'Total commissions',
    'tile.mismatch': 'Mismatch',
    'tile.match': 'Match',
    'tile.mand': 'Mandatory missing',
    'dash.models': 'Units by model',
    'dash.models.kinds': 'kinds',
    'dash.near': 'Nearest changeability date',
    'dash.near.count': 'Commissions on that date',
    'unit.ea': 'ea',
    'unit.case': 'ea',
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
    'matching.title': 'Model Matching',
    'matching.sub': 'SAM ↔ WINGS recognition rules — all stored in 03. model_rules/model_mapping.xlsx',
    'matching.loading': 'Loading model_mapping.xlsx…',
    'matching.note':
      'Every model-matching rule (normalization, previous/current model, vehicle keywords, aliases, ' +
      'manual map, options) lives as a sheet in this workbook. Switch sheet tabs to edit, then ' +
      '<b>Save to SharePoint</b> to write it back to <code>model_mapping.xlsx</code>; it takes effect on the next ' +
      '<b>Build data</b>. The <code>인식모델_대조표</code> sheet is a verification view produced by the local Python tool (the browser build does not refresh it).',
    'codes.title': 'Code Manager',
    'codes.sub': 'Edit & save the Excel files in the SharePoint 04. code folder, right here',
    'codes.pickFile': 'Loading file list…',
    'codes.empty': 'Pick an Excel file on the left to edit.',
    'codes.searchPh': 'Search this sheet…',
    'codes.unsaved': '● Unsaved changes',
    'codes.loadingFile': 'Opening file…',
    'codes.saving': 'Saving to SharePoint…',
    'codes.saved': 'Saved — written back to SharePoint.',
    'codes.listFail': 'Could not load the file list: {err}',
    'codes.noXlsx': 'No Excel (.xlsx) files in this folder.',
    'codes.confirmLeave': 'You have unsaved changes. Leaving will discard them. Continue?',
    'codes.delRow': 'Delete this row?',
    'btn.openFolder': '📂 Open SharePoint folder',
    'btn.refresh': '↻ Refresh list',
    'btn.loadSp': '⭳ Load from SharePoint',
    'btn.saveSp': '⭱ Save to SharePoint',
    'btn.addRow': '＋ Add row',
    'btn.addRec': '＋ New',
    'btn.reloadFile': '↺ Revert',
    'btn.download': '⬇ Download',
    'btn.delete': 'Delete',
    'btn.cancel': 'Cancel',
    'btn.apply': 'Save',
    'btn.edit': 'Edit',
    'mode.record': '🗂 Form editing',
    'mode.grid': '▦ Table editing',
    'rec.actions': 'Actions',
    'rec.count': '{n} / {total} records',
    'rec.empty': 'No records to show.',
    'rec.emptySheet': '— empty sheet —',
    'rec.new': 'New record',
    'rec.edit': 'Edit record',
    'rec.page': 'Page {page} / {pages}',
    'rec.prev': '‹ Prev',
    'rec.next': 'Next ›',
    'rec.allOf': 'All · {col}',
    'rec.blank': '(blank)',
    'rec.on': 'Yes',
    'rec.off': 'No',
    'rec.colFallback': 'Column {n}',
    'rec.delConfirm': 'Delete this record?',
    'rec.hint': 'Click a row to open the edit form. To change the columns themselves, switch to table editing.',
    'modal.upload': '⬆ Load Excel (local)',
    'modal.upload.title': 'Load model_mapping.xlsx from your PC',
    'modal.download': '⬇ Save Excel (local)',
    'op.loading': 'Working…',
    'op.needLogin': 'SharePoint sync is available after signing in with your company account.',
    'op.loaded': 'Loaded.',
    'op.saved': 'Saved — written back to SharePoint.',
    'op.fail': 'Failed: {err}',
    'nav.build': '⟳ Build data',
    'build.btn': '⟳ Build data',
    'build.running': '⟳ Building…',
    'build.title': 'Build data',
    'build.confirm': 'This recomputes the comparison from the newest WINGS export (02) and the newest '
      + 'production-month SAM folder (01) on SharePoint.\nIt runs in this browser (1–3 minutes) and the result is '
      + 'saved to SharePoint (05. output) for everyone.\n\nContinue?',
    'build.needLogin': 'Building requires signing in with your company Microsoft 365 account.',
    'build.step.ref': 'Reading reference workbooks (codes, matching rules)…',
    'build.step.upload': 'Saving the result to SharePoint…',
    'build.done': 'Build complete — {rows} rows ({match} match / {mismatch} mismatch). Saved to SharePoint.',
    'build.fail': 'Build failed:',
    'build.close': 'Close',
    'meta.srcSp': 'Source: latest SharePoint build',
    'meta.srcBundled': 'Source: bundled copy (no build result yet)',
    'alert.xlsxBlocked': 'The Excel library could not be loaded (corporate network may block it). Try again shortly.',
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
  renderDashSide();
  fillFilters();
  render();
  if (typeof codeEditor !== 'undefined') codeEditor.refresh();
  if (typeof matchEditor !== 'undefined') matchEditor.refresh();
  if (DRAWER_ROW && !$('#drawer').classList.contains('hidden')) openDrawer(DRAWER_ROW);
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ====================== 뷰 전환 (대시보드 / 모델 매칭 / 코드 관리) ======================
let CUR_VIEW = 'dashboard';
const VIEW_INIT = { matching: false, codes: false };

function switchView(view) {
  if (!['dashboard', 'matching', 'codes'].includes(view)) return;
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
  document.querySelectorAll('.nav-link').forEach((el) =>
    el.classList.toggle('active', el.dataset.view === view));
  if (view === 'matching' && !VIEW_INIT.matching) { VIEW_INIT.matching = true; initMatching(); }
  if (view === 'codes' && !VIEW_INIT.codes) { VIEW_INIT.codes = true; initCodes(); }
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
  applyData(d, c);
}

// 새로 만든(또는 내려받은) 데이터로 대시보드를 다시 그린다.
function applyData(d, c) {
  DATA = d; CODES = c || CODES;
  (DATA.rows || []).forEach((r) => { r.MY = computeMY(r); });
  renderMeta();
  renderSummary();
  renderDashSide();
  fillFilters();
  renderHead();
  render();
}

function prodMonth(r) { return String(r['Production date'] || '').slice(0, 7); }
function changeMonth(r) {
  const s = String(r['Changeability Date'] || '');
  return /^\d{4}-\d{2}/.test(s) ? s.slice(0, 7) : '';
}

function renderMeta() {
  const locale = LANG === 'ko' ? 'ko-KR' : 'en-GB';
  const when = DATA.generated_at ? new Date(DATA.generated_at).toLocaleString(locale) : '-';
  const src = DATA_SOURCE === 'sharepoint' ? t('meta.srcSp')
    : (DATA_SOURCE === 'bundled' ? t('meta.srcBundled') : '');
  $('#meta').textContent = t('meta.generated', { when, file: DATA.wings_file || '-' })
    + (src ? '  ·  ' + src : '');
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

function overallRows() {
  return DATA.rows.filter((r) => { const pm = prodMonth(r); return pm && pm >= CUR_MONTH; });
}

function tile(cls, n, label, action) {
  return `<div class="tile ${cls} clickable" data-tile="${cls}" data-action="${action}">` +
    `<div class="n">${esc(n)}</div><div class="l">${label}</div></div>`;
}

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
    ${nearestCardHtml()}
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

// 오늘과 가장 가까운 Changeability 카드(왼쪽 상단).
function nearestCardHtml() {
  const near = nearestChangeability(overallRows());
  const body = near
    ? `<div class="near-date">${esc(near.date)}<span class="near-dd">${ddayLabel(near.dday)}</span></div>
       <div class="near-count"><b>${near.count}</b> <span>${esc(t('unit.case'))} · ${esc(t('dash.near.count'))}</span></div>`
    : `<div class="mc-empty">—</div>`;
  return `<div class="near-card">
      <div class="dash-cap">${esc(t('dash.near'))}</div>
      ${body}
    </div>`;
}

function syncTileActive() {
  $('#summary').querySelectorAll('.tile').forEach((el) =>
    el.classList.toggle('active', el.dataset.tile === activeTile));
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

function renderDashSide() {
  const rows = overallRows();
  const combos = modelCombos(rows);
  const totalUnits = combos.reduce((s, c) => s + c.count, 0);

  const listHtml = combos.length
    ? combos.map((c) => `
      <div class="mc-row">
        <div class="mc-label" title="${esc(c.parts.filter(Boolean).join(' · '))}">${esc(c.parts.filter(Boolean).join(' · '))}</div>
        <div class="mc-count">${c.count}</div>
      </div>`).join('')
    : `<div class="mc-empty">—</div>`;

  $('#dashSide').innerHTML = `
    <div class="side-card">
      <div class="side-cap">${esc(t('dash.models'))}
        <span class="side-sub">${combos.length}${esc(t('dash.models.kinds'))} · ${totalUnits}${esc(t('unit.ea'))}</span>
      </div>
      <div class="mc-list">${listHtml}</div>
    </div>`;
}

function fillSelectFilter(id, allKey, values) {
  const sel = $(id);
  const prev = sel.value;
  sel.innerHTML = `<option value="">${esc(t(allKey))}</option>` +
    values.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  sel.value = prev;
}

function uniqueSorted(fn, reverse) {
  const arr = [...new Set(DATA.rows.map(fn).filter((v) => v !== undefined && v !== null && v !== ''))]
    .map(String).sort();
  return reverse ? arr.reverse() : arr;
}

function fillFilters() {
  fillSelectFilter('#productionFilter', 'filter.allProd', uniqueSorted(prodMonth, true));
  fillSelectFilter('#myFilter', 'filter.allMY', uniqueSorted((r) => r.MY));
  fillSelectFilter('#modelFilter', 'filter.allModel', uniqueSorted((r) => r.Vehicle));
  fillSelectFilter('#typeFilter', 'filter.allType', uniqueSorted((r) => r['Model(WINGS)']));
  fillSelectFilter('#axleFilter', 'filter.allAxle', uniqueSorted((r) => r.Type));
  fillSelectFilter('#cabFilter', 'filter.allCab', uniqueSorted((r) => r.Cab));
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

function countCell(csv, cls) {
  const n = countOf(csv);
  if (!n) return `<span class="cbadge ok" title="${esc(t('cell.ok.title'))}">✓</span>`;
  return `<span class="cbadge ${cls}">${n}</span>`;
}

function filtered() {
  const q = $('#search').value.trim().toLowerCase();
  const status = $('#statusFilter').value;
  const prod = $('#productionFilter').value;
  const my = $('#myFilter').value;
  const model = $('#modelFilter').value;
  const type = $('#typeFilter').value;
  const axle = $('#axleFilter').value;
  const cab = $('#cabFilter').value;
  const upcoming = $('#upcomingOnly').checked;
  let rows = DATA.rows.filter((r) => {
    if (restrictSoon && !within2weeks(r)) return false;
    if (status && r['SAM Status'] !== status) return false;
    if (prod && prodMonth(r) !== prod) return false;
    if (my && String(r.MY ?? '') !== my) return false;
    if (model && r.Vehicle !== model) return false;
    if (type && r['Model(WINGS)'] !== type) return false;
    if (axle && r.Type !== axle) return false;
    if (cab && r.Cab !== cab) return false;
    if (upcoming) {
      const cm = changeMonth(r);
      if (!cm || cm < CUR_MONTH) return false;
    }
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

function describe(code) { return CODES.options[code] || CODES.mandatory[code] || ''; }
function countOf(csv) { return (csv || '').split(',').map((c) => c.trim()).filter(Boolean).length; }
function splitCodes(csv) { return (csv || '').split(',').map((c) => c.trim()).filter(Boolean); }

let DRAWER_ROW = null;

function codeRowsHtml(codes, extraClass) {
  const cls = extraClass ? ' ' + extraClass : '';
  return codes.map((c) =>
    `<div class="code-row${cls}"><span class="c">${esc(c)}</span><span class="d">${esc(describe(c) || '—')}</span></div>`).join('');
}

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

// Aligned SAM-vs-WINGS block for a named category (Paint / Tyre). Same two-column
// layout as alignedFullHtml but with its own heading; returns '' when there is no
// data on either side so empty groups don't clutter the chart.
function alignedGroupHtml(title, samCsv, wingsCsv, kind, diffOnly) {
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
    return `<div class="acode-row">`
      + `<div class="acode-cell${inS ? '' : ' miss'}">${cell(code, inS)}</div>`
      + `<div class="acode-cell${inW ? '' : ' miss'}">${cell(code, inW)}</div>`
      + `</div>`;
  }).join('');
  return `<div class="acode" style="margin-bottom:16px">
      <div class="acode-head">
        <h4>${esc(title)} · SAM <span class="badge">${sam.size}</span></h4>
        <h4>${esc(title)} · WINGS <span class="badge">${wings.size}</span></h4>
      </div>
      <div class="acode-body">${rows}</div>
    </div>`;
}

function openDrawer(r) {
  if (!r) return;
  DRAWER_ROW = r;
  // Title: Commission · Model · Type · Axle · Cab · PTO · MY (skip blanks).
  // Column semantics: Model=Vehicle, Type=Model(WINGS), Axle=Type.
  const _titleParts = [r['Vehicle'], r['Model(WINGS)'], r['Type'], r['Cab'], r['PTO'],
    r['MY'] ? 'MY' + r['MY'] : '']
    .map((x) => (x == null ? '' : String(x).trim())).filter(Boolean);
  $('#drawerTitle').textContent = `${r['Commission no.']}  ·  ${_titleParts.join('  ·  ')}`;
  $('#drawerSub').textContent = r['Compared SAM file name'] || '';
  const meta = ['Vehicle', 'Category', 'Type', 'Cab', 'MY', 'PTO', 'Production date', 'Changeability Date',
    'Until Dealine', 'SAM Baumuster', 'SAM now',
    'Order status financial', 'SAM Status', 'FIN']
    .filter((k) => r[k] !== undefined && r[k] !== '')
    .map((k) => {
      let val;
      if (DDAY_KEYS.has(k)) val = ddayHtml(r[k]);
      else if (k === 'SAM Status') {
        const s = String(r[k]);
        val = `<span class="status ${esc(s).replace(/\s+/g, '')}">${esc(s)}</span>`;
      } else val = esc(r[k]);
      return `<div class="kv-item"><div class="k">${esc(META_LABELS[k] || k)}</div><div class="v">${val}</div></div>`;
    }).join('');

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
      ${alignedGroupHtml('🎨 Paint', r['_paint_sam'], r['_paint_wings'], 'paint', true)}
      ${alignedGroupHtml('🛞 Tyre', r['_tyre_sam'], r['_tyre_wings'], 'tyre', true)}
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
      ${alignedGroupHtml('🎨 Paint', r['_paint_sam'], r['_paint_wings'], 'paint')}
      ${alignedGroupHtml('🛞 Tyre', r['_tyre_sam'], r['_tyre_wings'], 'tyre')}
      ${alignedFullHtml(r['_all_sam_codes'], r['_all_wings_codes'])}
    </div>
  `;
  $('#drawerBody').querySelectorAll('.tab').forEach((tb) =>
    tb.addEventListener('click', () => switchTab(tb.dataset.tab)));
  $('#drawerXls').addEventListener('click', () => exportRowXls(r));
  $('#drawer').classList.remove('hidden');
  $('#backdrop').classList.remove('hidden');
}

function switchTab(name) {
  $('#drawerBody').querySelectorAll('.tab').forEach((tb) =>
    tb.classList.toggle('active', tb.dataset.tab === name));
  $('#drawerBody').querySelectorAll('.tab-pane').forEach((p) =>
    p.classList.toggle('hidden', p.dataset.pane !== name));
}

function closeDrawer() {
  $('#drawer').classList.add('hidden');
  $('#backdrop').classList.add('hidden');
}

// ---- Dependency-free .xlsx writer (per-vehicle drawer export) ----
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
    'SAM Baumuster', 'SAM now', 'SAM Status', 'Compared SAM file name'];
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

  function setDirty(v) { S.dirty = v; const d = el('dirty'); if (d) d.classList.toggle('hidden', !v); }
  function curAoa() { return S.sheets[S.sheet] || []; }

  function renderTabs() {
    const tabs = el('tabs');
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

    thead.innerHTML = '<tr><th class="rownum">#</th>' +
      S.cols.map((col) => `<th${col.type === 'bool' ? ' class="tight"' : (col.mono ? ' class="code-col-cell"' : '')}>${esc(col.label)}</th>`).join('') +
      `<th class="rowact">${esc(t('rec.actions'))}</th></tr>`;

    if (!slice.length) {
      tbody.innerHTML = `<tr><td class="none" colspan="${S.cols.length + 2}">${esc(t('rec.empty'))}</td></tr>`;
    } else {
      tbody.innerHTML = slice.map((r) => {
        const row = aoa[r] || [];
        return `<tr data-r="${r}"><td class="rownum">${r}</td>` +
          S.cols.map((col) => `<td${col.type === 'bool' ? ' class="tight"' : (col.mono ? ' class="code-col-cell"' : '')}>${recCellHtml(col, row[col.c], r)}</td>`).join('') +
          `<td class="rowact"><button type="button" class="link-btn rec-edit" data-r="${r}">${esc(t('btn.edit'))}</button>` +
          `<button type="button" class="link-btn danger rec-del" data-r="${r}">${esc(t('btn.delete'))}</button></td></tr>`;
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
  function editRecord(r) {
    const aoa = curAoa();
    const isNew = r == null;
    const row = isNew ? [] : (aoa[r] || []);
    RecModal.open({
      title: isNew ? t('rec.new') : t('rec.edit'),
      cols: S.cols,
      values: row,
      onDelete: isNew ? null : () => { aoa.splice(r, 1); setDirty(true); renderAll(); },
      onSave: (vals) => {
        const target = isNew ? [] : row;
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
    show('addRec', mode === 'record');
    show('addRow', mode === 'grid');
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
    let thead = '<tr><th class="rownum">#</th>';
    for (let c = 0; c < ncol; c++)
      thead += `<th><div class="cell-edit hdr" contenteditable="true" data-r="0" data-c="${c}">${esc(header[c] ?? '')}</div></th>`;
    thead += '<th class="rowact"></th></tr>';
    grid.querySelector('thead').innerHTML = thead;
    let body = '';
    for (let r = 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      body += `<tr data-r="${r}"><td class="rownum">${r}</td>`;
      for (let c = 0; c < ncol; c++)
        body += `<td><div class="cell-edit" contenteditable="true" data-r="${r}" data-c="${c}">${esc(row[c] ?? '')}</div></td>`;
      body += `<td class="rowact"><button class="row-del" data-r="${r}" title="행 삭제">🗑</button></td></tr>`;
    }
    grid.querySelector('tbody').innerHTML = body;
  }

  function onInput(e) {
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
      S.file = filename; S.sheets = {}; S.order = wb.SheetNames.slice();
      for (const n of wb.SheetNames)
        S.sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, blankrows: false, defval: '' });
      S.sheet = S.order[0] || null;
      setDirty(false);
      const emp = el('empty'); if (emp) emp.classList.add('hidden');
      const mn = el('main'); if (mn) mn.classList.remove('hidden');
      const sb = el('search'); if (sb) sb.value = '';
      renderTabs();
      S.cols = inferCols(curAoa());
      S.page = 1;
      setMode(S.mode);
      setStatus(stSel, isRevert ? t('op.loaded') : '', isRevert ? 'ok' : 'info');
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
  return { S, open, saveSp, wire, refresh };
}

// ---- 모델 매칭: model_mapping.xlsx (03. model_rules) 전체 시트 편집 ----
const matchEditor = makeSheetEditor({
  folderKey: 'model_rules', fixedFile: MODEL_MAPPING_FILE,
  els: { tabs: 'matchTabs', grid: 'matchGrid', addRow: 'matchAddRow', revert: 'matchLoadSp',
         saveSp: 'matchSaveSp', download: 'matchDownload', dirty: 'matchDirty',
         status: 'matchStatus', search: 'matchSearch', main: 'matchMain', empty: 'matchEmpty',
         recs: 'matchRecs', recWrap: 'matchRecWrap', gridWrap: 'matchGridWrap', pager: 'matchPager',
         mode: 'matchMode', filter: 'matchFilter', count: 'matchCount', addRec: 'matchAddRec' },
});
function initMatching() {
  const link = $('#matchFolderLink'); if (window.Graph) link.href = Graph.folders.model_rules.shareUrl;
  matchEditor.wire();
  matchEditor.open(MODEL_MAPPING_FILE);
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
  if (!window.Graph || !Graph.available()) { box.innerHTML = `<div class="none">${esc(t('op.needLogin'))}</div>`; return; }
  box.innerHTML = `<div class="none">${esc(t('codes.pickFile'))}</div>`;
  try {
    const items = await Graph.list('code');
    CODE_FILES = items.filter((i) => !i.isFolder && /\.xlsx?$/i.test(i.name));
    if (!CODE_FILES.length) { box.innerHTML = `<div class="none">${esc(t('codes.noXlsx'))}</div>`; return; }
    box.innerHTML = CODE_FILES.map((f) => {
      const kb = f.size ? Math.max(1, Math.round(f.size / 1024)) + ' KB' : '';
      return `<button class="file-item${f.name === codeEditor.S.file ? ' active' : ''}" data-file="${esc(f.name)}"><span class="fi-name">📄 ${esc(f.name)}</span><span class="fi-meta">${esc(kb)}</span></button>`;
    }).join('');
    box.querySelectorAll('.file-item').forEach((b) => b.addEventListener('click', () => {
      if (codeEditor.S.dirty && b.dataset.file !== codeEditor.S.file && !confirm(t('codes.confirmLeave'))) return;
      codeEditor.open(b.dataset.file).then(() =>
        box.querySelectorAll('.file-item').forEach((x) => x.classList.toggle('active', x.dataset.file === b.dataset.file)));
    }));
  } catch (e) { box.innerHTML = `<div class="none err">${esc(t('codes.listFail', { err: e.message }))}</div>`; }
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

let BUILDING = false;
async function runBuild() {
  if (BUILDING) return;
  if (!window.Graph || !Graph.available()) { alert(t('build.needLogin')); return; }
  if (!window.Pipeline) { alert(t('build.fail') + ' pipeline.js'); return; }
  if (!window.XLSX) { alert(t('alert.xlsxBlocked')); return; }
  if (!confirm(t('build.confirm'))) return;

  BUILDING = true;
  const btn = $('#buildBtn');
  const orig = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; btn.textContent = t('build.running'); }
  const log = buildLog();
  log.line(t('build.step.ref'));
  try {
    const { data, codes } = await Pipeline.build(GRAPH_SOURCE, { log: (m) => log.line(m) });

    log.line(t('build.step.upload'));
    await Graph.ensureFolder('output');
    await Graph.uploadJson('output', 'data.json', data);
    await Graph.uploadJson('output', 'codes.json', codes);

    DATA_SOURCE = 'sharepoint';
    applyData(data, codes);
    const msg = t('build.done', {
      rows: data.summary.total, match: data.summary.matched, mismatch: data.summary.mismatched });
    log.line(msg, 'ok');
  } catch (e) {
    console.error(e);
    log.line(t('build.fail') + ' ' + e.message, 'err');
  } finally {
    BUILDING = false;
    if (btn) { btn.disabled = false; btn.textContent = orig || t('build.btn'); }
  }
}

// ====================== 이벤트 & 초기화 ======================
$('#drawerClose').addEventListener('click', closeDrawer);
$('#backdrop').addEventListener('click', closeDrawer);
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeDrawer(); });
$('#langBtn').addEventListener('click', toggleLang);

// 뷰 전환 (네비 + 브랜드)
document.querySelectorAll('[data-view]').forEach((el) =>
  el.addEventListener('click', (e) => { e.preventDefault(); switchView(el.dataset.view); }));

// 데이터 빌드 (관리자) — 이 브라우저에서 계산 후 SharePoint 에 저장
const _buildBtn = $('#buildBtn');
if (_buildBtn) _buildBtn.addEventListener('click', runBuild);

// 저장 안 한 편집이 있으면 페이지 이탈 경고 (모델 매칭 / 코드 관리)
window.addEventListener('beforeunload', (e) => {
  if (codeEditor.S.dirty || matchEditor.S.dirty) { e.preventDefault(); e.returnValue = ''; }
});

function onManualFilter() {
  restrictSoon = false;
  tileMandatory = false;
  activeTile = null;
  syncTileActive();
  render();
}
['#search', '#statusFilter', '#productionFilter', '#myFilter', '#modelFilter',
  '#typeFilter', '#axleFilter', '#cabFilter',
  '#upcomingOnly'].forEach((s) =>
  $(s).addEventListener('input', onManualFilter));

applyStaticI18n();
load().catch((e) => {
  $('#meta').textContent = t('meta.loadFail', { err: e.message });
  const msg = $('#statusMsg');
  msg.textContent = t('msg.loadFail', { err: e.message });
  msg.classList.remove('hidden');
});
