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
  { key: 'PTO', label: 'PTO', pair: 'pto' },
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

let DATA = { rows: [] };
let CODES = { options: {}, mandatory: {} };
let sortKey = null, sortDir = 1;
let restrictSoon = false;
let tileMandatory = false;
let tileSamUpdate = false;
let activeTile = 't-total';

const $ = (s) => document.querySelector(s);

// ====================== i18n (한국어 / English) ======================
const I18N = {
  ko: {
    'nav.dashboard': '대시보드',
    'nav.history': '생산월 이력관리',
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
    'tile.samupd': 'SAM update요청',
    'tile.nosam': 'No SAM',
    'dash.models': '모델별 대수',
    'dash.models.filtered': '모델별 대수 (필터)',
    'dash.models.kinds': '종',
    'dash.near': 'Upcoming changeability',
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
    // 생산월 이력관리
    'hist.title': '생산월 이력관리',
    'hist.sub': '같은 MY·같은 모델의 WINGS commission 을 생산월 순으로 세워 두고, 달마다 추가(＋)·삭제(−)된 코드를 보여줍니다.',
    'hist.src.wings': 'WINGS 코드',
    'hist.src.sam': 'SAM 코드',
    'hist.src.both': 'WINGS + SAM 대조',
    'hist.xTitle': '이 달 WINGS↔SAM 차이 — WINGS 에만: {w} / SAM 에만: {s}',
    'hist.xNone': '이 달 WINGS 와 SAM 의 코드가 같습니다.',
    'hist.xOk': '일치',
    'hist.xBad': '불일치',
    'hist.basis.title': '각 달의 코드를 무엇과 비교할지',
    'hist.basis.prev': '이전 생산월 대비',
    'hist.basis.first': '첫 생산월 대비',
    'hist.searchPh': '모델 / Commission 검색…',
    'hist.added': '추가',
    'hist.removed': '삭제',
    'hist.export': '⬇ Export',
    'hist.exportTitle': '지금 보이는 히스토리를 엑셀(.xlsx)로 저장',
    'hist.model': '모델',
    'hist.models': '모델 목록',
    'hist.modelSub': '{n}대 · {m}개월',
    'hist.chgMonths': '개월 변경',
    'hist.more': '외 {n}대',
    'hist.prodMonth': '생산월',
    'hist.base': '기본',
    'hist.same': '변경 없음',
    'hist.count': '모델 {models}종 · {months}개월 · Commission {n}건',
    'hist.emptyMy': '이 조건에 해당하는 commission 이 없습니다.',
    'hist.noMy': 'MY 미상',
    'hist.varied': '⚠ 같은 달 commission 간 코드 차이 {n}개: {codes}',
    'hist.cellTitle': '{ref} → {cur} 비교',
    'hist.baseTitle': '{m} — 이 모델의 첫 생산월(기준)',
    'hist.gapTitle': '이 달에는 생산이 없습니다.',
    'hist.exportTitleRow': '코드 히스토리 — {my} · {src} · {basis}',
    // 모델 매칭
    'matching.title': '모델 매칭',
    'matching.sub': 'SAM ↔ WINGS 모델 인식 — 규칙은 model_mapping.xlsx',
    'matching.paneResult': '매칭 결과',
    'matching.paneSheet': '대조표',
    'matching.loading': 'model_mapping.xlsx 를 불러오는 중…',
    'matching.result': '매칭 결과',
    'matching.resultSub': '모델 → 비교된 SAM 파일. 캡·PTO 는 코드로 판정하며 WINGS/SAM 이 다르면 ≠.',
    'match.count': '모델 {n}종',
    'match.countOf': '모델 {n}종 / 전체 {total}종',
    'match.allPto': '전체 PTO',
    'match.searchPh': '모델 / SAM 파일 검색…',
    'match.noHit': '이 조건에 맞는 모델이 없습니다.',
    'match.samFile': '비교된 SAM 파일',
    'match.units': '대수',
    'match.status': '상태',
    'match.noFile': '매칭된 SAM 없음',
    'match.files': '문서 {n}개',
    'matching.note':
      '<code>인식모델_대조표</code>는 열 때마다 지금 데이터로 다시 만들어지는 확인용 보기입니다. ' +
      '규칙 시트를 고쳐 <b>SharePoint에 저장</b>하면 다음 <b>데이터 다시 계산</b> 때 적용됩니다.',
    // 코드 관리
    'codes.title': '코드 관리',
    'codes.sub': "SharePoint 04. code 폴더의 Excel 파일을 웹에서 직접 편집·저장",
    'codes.note':
      '비교에 쓰이는 모든 기준표(필수코드·공장관리코드·캡·모델 카테고리 등)가 ' +
      '<code>04. code</code> 폴더의 Excel 파일로 관리됩니다. 아래에서 파일을 고르고 시트 탭을 바꿔 ' +
      '편집한 뒤 <b>SharePoint에 저장</b>하면 원본 파일에 반영되고, 다음 <b>데이터 다시 계산</b> 때 적용됩니다.',
    'codes.pickFile': '파일 목록을 불러오는 중…',
    'codes.empty': '위에서 편집할 Excel 파일을 선택하세요.',
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
    'export.btn': '⬇ Export',
    'export.title': '현재 필터가 적용된 목록을 엑셀(.xlsx)로 저장',
    'export.sheetTitle': 'SAM × AFAB 비교 — {n} / {total} 건',
    'export.filters': '필터: {f}',
    'export.noFilter': '(필터 없음 — 전체)',
    'export.exportedAt': '내보낸 시각: {ts}',
    'export.noRows': '내보낼 행이 없습니다. 필터를 확인하세요.',
    'export.f.prod': '생산월',
    'export.f.my': 'MY',
    'export.f.model': 'Model',
    'export.f.type': 'Type',
    'export.f.axle': 'Axle',
    'export.f.cab': 'Cab',
    'export.f.status': '상태',
    'export.f.search': '검색',
    'export.f.sort': '정렬',
    'nav.build': '⟳ 데이터 다시 계산',
    'build.btn': '⟳ 데이터 다시 계산',
    'build.running': '⟳ 빌드 중…',
    'build.title': '데이터 다시 계산',
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
    'nav.history': 'Month History',
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
    'tile.samupd': 'SAM update needed',
    'tile.nosam': 'No SAM',
    'dash.models': 'Units by model',
    'dash.models.filtered': 'Units by model (filtered)',
    'dash.models.kinds': 'kinds',
    'dash.near': 'Upcoming changeability',
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
    'hist.title': 'Production-month History',
    'hist.sub': 'Lines up the WINGS commissions of one model (same MY) by production month and shows which codes were added (＋) or removed (−) each month.',
    'hist.src.wings': 'WINGS codes',
    'hist.src.sam': 'SAM codes',
    'hist.src.both': 'WINGS + SAM cross-check',
    'hist.xTitle': 'WINGS↔SAM this month — only in WINGS: {w} / only in SAM: {s}',
    'hist.xNone': 'WINGS and SAM carry the same codes this month.',
    'hist.xOk': 'match',
    'hist.xBad': 'mismatch',
    'hist.basis.title': 'What each month is compared against',
    'hist.basis.prev': 'vs previous month',
    'hist.basis.first': 'vs first month',
    'hist.searchPh': 'Search model / commission…',
    'hist.added': 'added',
    'hist.removed': 'removed',
    'hist.export': '⬇ Export',
    'hist.exportTitle': 'Download the history shown here as Excel (.xlsx)',
    'hist.model': 'Model',
    'hist.models': 'Models',
    'hist.modelSub': '{n} units · {m} months',
    'hist.chgMonths': ' months changed',
    'hist.more': '+{n} more',
    'hist.prodMonth': 'Production month',
    'hist.base': 'Baseline',
    'hist.same': 'No change',
    'hist.count': '{models} models · {months} months · {n} commissions',
    'hist.emptyMy': 'No commissions match this selection.',
    'hist.noMy': 'MY unknown',
    'hist.varied': '⚠ {n} code(s) differ between commissions of this month: {codes}',
    'hist.cellTitle': 'Compared {ref} → {cur}',
    'hist.baseTitle': '{m} — first production month of this model (baseline)',
    'hist.gapTitle': 'No production in this month.',
    'hist.exportTitleRow': 'Code history — {my} · {src} · {basis}',
    'matching.title': 'Model Matching',
    'matching.sub': 'SAM ↔ WINGS model recognition — rules live in model_mapping.xlsx',
    'matching.paneResult': 'Matching result',
    'matching.paneSheet': 'Recognition table',
    'matching.loading': 'Loading model_mapping.xlsx…',
    'matching.result': 'Matching result',
    'matching.resultSub': 'Model → the SAM file it was compared against. Cab and PTO come from the codes; ≠ marks a difference.',
    'match.count': '{n} models',
    'match.countOf': '{n} of {total} models',
    'match.allPto': 'All PTO',
    'match.searchPh': 'Search model / SAM file…',
    'match.noHit': 'No model matches this selection.',
    'match.samFile': 'Compared SAM file',
    'match.units': 'Units',
    'match.status': 'Status',
    'match.noFile': 'No SAM matched',
    'match.files': '{n} documents',
    'matching.note':
      '<code>인식모델_대조표</code> is a verification view, rebuilt from the current data every ' +
      'time this page opens. Edit a rule sheet and <b>Save to SharePoint</b>; it applies on the ' +
      'next <b>Recompute Data</b>.',
    'codes.title': 'Code Manager',
    'codes.sub': 'Edit & save the Excel files in the SharePoint 04. code folder, right here',
    'codes.note':
      'Every reference table used by the comparison (mandatory codes, factory control codes, cab, ' +
      'model category …) lives as an Excel file in the <code>04. code</code> folder. Pick a file below, ' +
      'switch sheet tabs to edit, then <b>Save to SharePoint</b> — the original file is updated and the ' +
      'change applies on the next <b>Rebuild data</b>.',
    'codes.pickFile': 'Loading file list…',
    'codes.empty': 'Pick an Excel file above to edit.',
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
    'export.btn': '⬇ Export',
    'export.title': 'Download the currently filtered list as Excel (.xlsx)',
    'export.sheetTitle': 'SAM × AFAB comparison — {n} / {total} rows',
    'export.filters': 'Filters: {f}',
    'export.noFilter': '(no filter — all rows)',
    'export.exportedAt': 'Exported at: {ts}',
    'export.noRows': 'Nothing to export — check the filters.',
    'export.f.prod': 'Production month',
    'export.f.my': 'MY',
    'export.f.model': 'Model',
    'export.f.type': 'Type',
    'export.f.axle': 'Axle',
    'export.f.cab': 'Cab',
    'export.f.status': 'Status',
    'export.f.search': 'Search',
    'export.f.sort': 'Sort',
    'nav.build': '⟳ Recompute Data',
    'build.btn': '⟳ Recompute Data',
    'build.running': '⟳ Building…',
    'build.title': 'Recompute Data',
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
  if (VIEW_INIT.history) { histFillMy(); histRender(); }
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
const VIEW_INIT = { history: false, matching: false, codes: false };

function switchView(view) {
  if (!['dashboard', 'history', 'matching', 'codes'].includes(view)) return;
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
  // Export 는 대시보드 필터에 대한 기능이라 다른 뷰에서는 숨긴다.
  const exportBtn = $('#exportBtn');
  if (exportBtn) exportBtn.classList.toggle('hidden', view !== 'dashboard');
  if (view === 'history' && !VIEW_INIT.history) { VIEW_INIT.history = true; initHistory(); }
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
  if (VIEW_INIT.history) { histFillMy(); histRender(); }
  if (VIEW_INIT.matching) {
    renderMatchSummary();
    // 대조표는 이 데이터로 만드는 보기라, 다시 계산했으면 열려 있는 시트도 새로 그린다.
    refreshRecognitionSheet(matchEditor.S);
    matchEditor.refresh();
  }
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
  return DATA.rows.filter((r) => { const pm = prodMonth(r); return pm && pm >= CUR_MONTH; });
}

function tile(cls, n, label, action) {
  return `<div class="tile ${cls} clickable" data-tile="${cls}" data-action="${action}">` +
    `<div class="n">${esc(n)}</div><div class="l">${label}</div></div>`;
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
  $('#summary').innerHTML = `
    ${nearestCardHtml()}
    <div class="dash-row">
      <div class="dash-cap">${t('dash.overall')}</div>
      <div class="tiles">
        ${tile('t-total', all.total, t('tile.total'), 't-total')}
        ${tile('t-match', all.match, t('tile.match'), 't-match')}
        ${tile('t-miss', all.mismatch, t('tile.mismatch'), 't-miss')}
        ${tile('t-nosam', all.nosam, t('tile.nosam'), 't-nosam')}
        ${tile('t-mand', all.mand, t('tile.mand'), 't-mand')}
        ${tile('t-samupd', all.samupd, t('tile.samupd'), 't-samupd')}
      </div>
    </div>
    <div class="dash-row">
      <div class="dash-cap">${t('dash.soon')}</div>
      <div class="tiles">
        ${tile('t-total2', soon.total, t('tile.total'), 't-total2')}
        ${tile('t-match2', soon.match, t('tile.match'), 't-match2')}
        ${tile('t-miss2', soon.mismatch, t('tile.mismatch'), 't-miss2')}
        ${tile('t-nosam2', soon.nosam, t('tile.nosam'), 't-nosam2')}
        ${tile('t-mand2', soon.mand, t('tile.mand'), 't-mand2')}
        ${tile('t-samupd2', soon.samupd, t('tile.samupd'), 't-samupd2')}
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
    ? `<div class="near-body">
         <div class="near-date">${esc(near.date)}<span class="near-dd">${ddayLabel(near.dday)}</span></div>
         <div class="near-count"><b>${near.count}</b> <span>${esc(t('unit.case'))} · ${esc(t('dash.near.count'))}</span></div>
       </div>`
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

// 필터와 연동되는 모델별 대수 (왼쪽 전체 리스트와 동일한 형식). render() 가 매 필터 변경 시 호출.
function renderFilteredModels(rows) {
  const el = $('#dashFiltered');
  if (!el) return;
  const combos = modelCombos(rows || []);
  const total = combos.reduce((s, c) => s + c.count, 0);
  const listHtml = combos.length
    ? combos.map((c) => {
      const label = c.parts.filter(Boolean).join(' · ');
      return `<div class="mc-row">
        <div class="mc-label" title="${esc(label)}">${esc(label)}</div>
        <div class="mc-count">${c.count}</div>
      </div>`;
    }).join('')
    : `<div class="mc-empty">—</div>`;
  el.innerHTML = `
    <div class="side-card">
      <div class="side-cap">${esc(t('dash.models.filtered'))}
        <span class="side-sub">${combos.length}${esc(t('dash.models.kinds'))} · ${total}${esc(t('unit.ea'))}</span>
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
    if (!pm || pm < CUR_MONTH) return false;
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
  for (const f of FILTER_FIELDS) fillSelectFilter(f.id, f.allKey, facetValues(f));
  enhanceFilterSelects();
}

// 필터 select 들을 커스텀 드롭다운으로 감싼다(아래로 열림, ~10개 표시 후 스크롤).
const FILTER_SELECT_IDS = ['#productionFilter', '#myFilter', '#modelFilter',
  '#typeFilter', '#axleFilter', '#cabFilter', '#statusFilter'];
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

function countCell(n, cls) {
  if (!n) return `<span class="cbadge ok" title="${esc(t('cell.ok.title'))}">✓</span>`;
  return `<span class="cbadge ${cls}">${n}</span>`;
}

// SAM Status 값 → CSS 클래스(공백/비ASCII 안전).
const STATUS_CLASS = {
  'Match': 'Match', 'Mismatch': 'Mismatch', 'No SAM': 'NoSAM', 'SAM update요청': 'SAMupdate',
};
function statusClass(v) {
  return STATUS_CLASS[v] || esc(v).replace(/[^A-Za-z0-9_-]/g, '');
}

function filtered() {
  const q = $('#search').value.trim().toLowerCase();
  let rows = DATA.rows.filter((r) => {
    if (restrictSoon && !within2weeks(r)) return false;
    if (tileMandatory && countOf(r['Mandatory Codes']) === 0) return false;
    if (tileSamUpdate && !r['SAM Update']) return false;
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
  $('#count').textContent = t('count', { n: rows.length, total: DATA.rows.length });
  renderFilteredModels(rows);
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
      if (c.status) {
        const upd = r['SAM Update'] ? `<span class="status SAMupdate upd">${esc(t('tile.samupd'))}</span>` : '';
        return `<td><div class="status-cell"><span class="status ${statusClass(v)}">${esc(v)}</span>${upd}</div></td>`;
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
// 캡·PTO 는 compare 가 WINGS/SAM 양쪽 값을 남긴다(_cab_wings/_cab_sam, _pto_wings/_pto_sam).
// 목록엔 한 값만 들어가므로, 양쪽이 다르면 셀에 ≠ 를 붙이고 툴팁에 두 값을 보여준다.
// 캡은 양쪽 다 신호가 있을 때만, PTO 는 '있다/없다'가 갈릴 때 불일치로 본다.
function pairSides(r, kind) {
  return { wings: splitCodes(r['_' + kind + '_wings']), sam: splitCodes(r['_' + kind + '_sam']) };
}
function pairMismatch(r, kind) {
  if (!r || !r['Compared SAM file name']) return false;   // 비교 대상 자체가 없으면 불일치 아님
  const { wings, sam } = pairSides(r, kind);
  if (kind === 'pto') return (wings.length > 0) !== (sam.length > 0);
  if (!wings.length || !sam.length) return false;
  return wings.join(',') !== sam.join(',');
}
function pairTitle(r, kind) {
  const { wings, sam } = pairSides(r, kind);
  return `WINGS: ${wings.join(', ') || '—'}  /  SAM: ${sam.join(', ') || '—'}`;
}

const DIFF_GROUPS = ['_paint', '_tyre'];
function diffCodes(r, kind) {
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
function diffCount(r, kind) { return diffCodes(r, kind).length; }

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
  // 차이 탭에서는 '한쪽에만 있는 코드' 수(=실제로 표시된 줄 수)를, 전체 탭에서는 그룹 전체 수를 뱃지에 쓴다.
  const nS = diffOnly ? union.filter((c) => sam.has(c)).length : sam.size;
  const nW = diffOnly ? union.filter((c) => wings.has(c)).length : wings.size;
  return `<div class="acode" style="margin-bottom:16px">
      <div class="acode-head">
        <h4>${esc(title)} · SAM <span class="badge">${nS}</span></h4>
        <h4>${esc(title)} · WINGS <span class="badge">${nW}</span></h4>
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
      else if ((k === 'Cab' || k === 'PTO') && pairMismatch(r, k.toLowerCase())) {
        val = `<span class="pair-diff" title="${esc(pairTitle(r, k.toLowerCase()))}">` +
          `${esc(r[k])}<span class="ne">≠</span></span>`;
      } else if (k === 'SAM Status') {
        const s = String(r[k]);
        const upd = r['SAM Update'] ? ` <span class="status SAMupdate">${esc(t('tile.samupd'))}</span>` : '';
        val = `<span class="status ${statusClass(s)}">${esc(s)}</span>${upd}`;
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
      ${alignedGroupHtml('🚪 Cab', r['_cab_sam'], r['_cab_wings'], 'cab', true)}
      ${alignedGroupHtml('⚙️ PTO', r['_pto_sam'], r['_pto_wings'], 'pto', true)}
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
      ${alignedGroupHtml('🚪 Cab', r['_cab_sam'], r['_cab_wings'], 'cab')}
      ${alignedGroupHtml('⚙️ PTO', r['_pto_sam'], r['_pto_wings'], 'pto')}
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
  if ($('#upcomingOnly').checked) parts.push(t('chk.upcoming'));
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
const HIST = { my: '', src: 'wings', basis: 'prev', q: '', rows: [], last: null };

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
function histCodes(r, src) {
  return splitCodes((src || HIST.src) === 'sam' ? r['_all_sam_codes'] : r['_all_wings_codes']);
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

  const out = [];
  let total = 0;
  for (const [key, byMonth] of models) {
    const seq = [...byMonth.keys()].sort();
    const cells = new Map();
    for (const m of seq) {
      const rs = byMonth.get(m);
      total += rs.length;
      // 두 소스를 늘 함께 계산해 둔다 — 'WINGS+SAM' 모드가 한 셀에서 둘을 나란히 쓴다.
      const per = {};
      for (const src of HIST_SRCS) {
        const sets = rs.map((r) => new Set(histCodes(r, src)));
        const codes = new Set();
        sets.forEach((st) => st.forEach((c) => codes.add(c)));
        per[src] = {
          codes,
          varied: [...codes].filter((c) => !sets.every((st) => st.has(c))).sort(),
        };
      }
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
    seq.forEach((m, i) => {
      const cell = cells.get(m);
      if (i === 0) { cell.base = true; return; }
      const refMonth = HIST.basis === 'first' ? seq[0] : seq[i - 1];
      cell.ref = refMonth;
      for (const src of HIST_SRCS) {
        const ref = cells.get(refMonth)[src].codes;
        const cur = cell[src];
        cur.added = [...cur.codes].filter((c) => !ref.has(c)).sort();
        cur.removed = [...ref].filter((c) => !cur.codes.has(c)).sort();
      }
    });
    // 모델 표(대시보드와 같은 항목)를 그릴 때 쓰는 대표 행 — 같은 키면 항목 값이 같다.
    out.push({ key, months: seq, cells, sample: cells.get(seq[0]).rows[0] });
  }
  out.sort((a, b) => a.key.localeCompare(b.key));
  return { months: [...monthSet].sort(), models: out, total };
}

// 'YYYY-MM' → 헤더 표시(연도 + n월). 같은 연도가 이어지면 연도는 위에 한 번만.
function histMonthLabel(m) {
  const [y, mm] = m.split('-');
  const num = Number(mm);
  return { top: y, main: LANG === 'ko' ? num + '월' : m };
}

// 그 달 아래 붙는 WINGS↔SAM 일치/불일치 띠 — 월 머리글 바로 다음 줄.
function histXHeadHtml(cell) {
  if (!cell) return `<th class="hist-xcell off"></th>`;
  if (cell.onlyW.length || cell.onlyS.length) {
    const tip = t('hist.xTitle', {
      w: cell.onlyW.join(', ') || '—', s: cell.onlyS.join(', ') || '—',
    });
    return `<th class="hist-xcell"><span class="hist-x bad" title="${esc(tip)}">` +
      `${esc(t('hist.xBad'))}</span></th>`;
  }
  return `<th class="hist-xcell"><span class="hist-x ok" title="${esc(t('hist.xNone'))}">` +
    `${esc(t('hist.xOk'))}</span></th>`;
}

// commission 줄 — 셀 폭을 좁게 유지하려고 대표 1건만 칩으로 보여주고,
// 나머지는 '외 n대' 로 접는다(툴팁에 전체 번호).
function histCommCellHtml(cell) {
  if (!cell) return histGapCellHtml();
  const r0 = cell.rows[0];
  const ri = HIST.rows.push(r0) - 1;
  const all = cell.rows.map((r) => r['Commission no.']).join(', ');
  const more = cell.rows.length > 1
    ? `<span class="cmore" title="${esc(all)}">${esc(t('hist.more', { n: cell.rows.length - 1 }))}</span>`
    : '';
  return `<td class="hist-cell comm${cell.base ? ' base' : ''}"><div class="hist-comm">` +
    `<button type="button" class="cno" data-ri="${ri}" title="Commission ${esc(r0['Commission no.'])}">` +
    `${esc(r0['Commission no.'])}</button>${more}</div></td>`;
}

function histGapCellHtml() {
  return `<td class="hist-cell gap" title="${esc(t('hist.gapTitle'))}">` +
    `<span class="hist-none">—</span></td>`;
}

// 한 소스(WINGS 또는 SAM)의 그 달 칸 — 변경 코드(＋추가 / −삭제) 또는 '기본'·'변경 없음'.
function histSrcCellHtml(cell, month, src) {
  if (!cell) return histGapCellHtml();
  const d = cell[src];
  let body, cls = '', title;
  if (cell.base) {
    cls = ' base';
    title = t('hist.baseTitle', { m: month });
    body = `<div class="hist-base">${esc(t('hist.base'))}</div>`;
  } else {
    title = t('hist.cellTitle', { ref: cell.ref, cur: month });
    const line = (c, kind) =>
      `<span class="hcode ${kind}" title="${esc(describe(c) || c)}">` +
      `${kind === 'add' ? '＋' : '−'}${esc(c)}</span>`;
    const lines = d.added.map((c) => line(c, 'add'))
      .concat(d.removed.map((c) => line(c, 'del')));
    body = lines.length
      ? `<div class="hist-codes">${lines.join('')}</div>`
      : `<div class="hist-same">${esc(t('hist.same'))}</div>`;
  }
  // 같은 달 commission 끼리 갈린 코드
  const warn = d.varied.length
    ? `<div class="hist-warn" title="${esc(histVariedText(d.varied))}">⚠ ${d.varied.length}</div>`
    : '';
  return `<td class="hist-cell${cls}" title="${esc(title)}">${body}${warn}</td>`;
}

// 모델 목록(왼쪽) — 모델별 대수 / 생산월 수 / 변경이 있었던 달 수.
function histModelListHtml(b) {
  return b.models.map((g) => {
    const units = g.months.reduce((s, m) => s + g.cells.get(m).rows.length, 0);
    const chg = g.months.filter((m) => {
      const c = g.cells.get(m);
      if (c.base) return false;
      return HIST_SRCS.some((src) =>
        (HIST.src === 'both' || HIST.src === src) && (c[src].added.length || c[src].removed.length));
    }).length;
    const sub = t('hist.modelSub', { n: units, m: g.months.length });
    const chgHtml = chg
      ? ` · <b class="chg">${chg}${esc(t('hist.chgMonths'))}</b>` : '';
    return `<button type="button" class="hm-item${g.key === HIST.model ? ' sel' : ''}" ` +
      `data-key="${esc(g.key)}" title="${esc(g.key)}">` +
      `<span class="hm-name">${esc(g.key)}</span>` +
      `<span class="hm-sub">${esc(sub)}${chgHtml}</span></button>`;
  }).join('');
}

// 선택한 모델 1종의 월별 히스토리 표(가로 스크롤).
// 열은 생산월, 줄은 commission → WINGS → SAM. 두 소스를 함께 볼 때는 월 머리글 바로 아래에
// 그 달의 WINGS↔SAM 일치/불일치를 한 줄로 깔아 준다.
function histDetailHtml(g, months) {
  const both = HIST.src === 'both';
  const srcs = both ? HIST_SRCS : [HIST.src === 'sam' ? 'sam' : 'wings'];
  const corner = `<th class="hist-rowhead corner"></th>`;
  let head = `<tr>` + corner + months.map((m) => {
    const lb = histMonthLabel(m);
    const has = g.cells.has(m);
    return `<th title="${esc(m)}"${has ? '' : ' class="off"'}>` +
      `<span class="my">${esc(lb.top)}</span>${esc(lb.main)}</th>`;
  }).join('') + `</tr>`;
  if (both) {
    head += `<tr class="hist-xrow">` + corner +
      months.map((m) => histXHeadHtml(g.cells.get(m))).join('') + `</tr>`;
  }
  let body = `<tr class="r-comm"><th class="hist-rowhead">Commission</th>` +
    months.map((m) => histCommCellHtml(g.cells.get(m))).join('') + `</tr>`;
  body += srcs.map((src) =>
    `<tr><th class="hist-rowhead"><span class="hsrc ${src}" ` +
    `title="${esc(t('hist.src.' + src))}">${src === 'sam' ? 'SAM' : 'WINGS'}</span></th>` +
    months.map((m) => histSrcCellHtml(g.cells.get(m), m, src)).join('') + `</tr>`).join('');
  return `<div class="hist-cap">${histModelTableHtml(g.sample)}` +
    `<span class="sub">${esc(histSrcLabel())} · ${esc(histBasisLabel())}</span></div>` +
    `<div class="hist-scroll"><table class="hist-grid">` +
    `<thead>${head}</thead><tbody>${body}</tbody></table></div>`;
}

// 선택한 모델을 대시보드 목록과 똑같은 항목·라벨로 한 줄 표에 보여준다.
// 캡·PTO 는 WINGS/SAM 이 갈리면 ≠ 로 표시한다(툴팁에 양쪽 값).
function histModelTableHtml(r) {
  if (!r) return '';
  const head = HIST_COLS.map((k) => `<th>${esc(colLabel(k))}</th>`).join('');
  const body = HIST_COLS.map((k) => {
    const v = String(r[k] ?? '').trim();
    const kind = k === 'Cab' ? 'cab' : (k === 'PTO' ? 'pto' : '');
    if (kind && pairMismatch(r, kind)) {
      return `<td class="pair-diff" title="${esc(pairTitle(r, kind))}">${esc(v)}<span class="ne">≠</span></td>`;
    }
    return `<td>${esc(v)}</td>`;
  }).join('');
  return `<table class="hist-model"><thead><tr>${head}</tr></thead>` +
    `<tbody><tr>${body}</tr></tbody></table>`;
}

function histRender() {
  const el = $('#histBody');
  const listEl = $('#histModels');
  if (!el || !listEl) return;
  const countEl = $('#histCount');
  const empty = (msg) => {
    listEl.innerHTML = '';
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
      { models: b.models.length, months: b.months.length, n: b.total });
  }
  if (!b.models.length) { empty(t('hist.emptyMy')); return; }

  // 선택이 없거나 지금 목록에 없으면 볼거리가 가장 많은 모델(생산월이 긴 쪽)을 고른다.
  if (!b.models.some((g) => g.key === HIST.model)) {
    HIST.model = [...b.models].sort((x, y) => y.months.length - x.months.length)[0].key;
  }

  listEl.innerHTML = histModelListHtml(b);
  listEl.querySelectorAll('.hm-item').forEach((btn) =>
    btn.addEventListener('click', () => {
      HIST.model = btn.dataset.key;
      histRender();
      const sel = $('#histModels').querySelector('.hm-item.sel');
      if (sel) sel.scrollIntoView({ block: 'nearest' });
    }));

  const g = b.models.find((x) => x.key === HIST.model);
  // 그 모델이 실제로 생산된 달만 열로 세운다 — 빈 달을 지우면 한 화면에 더 많은 달이 들어온다.
  el.innerHTML = histDetailHtml(g, g.months);
  el.querySelectorAll('.cno').forEach((btn) =>
    btn.addEventListener('click', () => openDrawer(HIST.rows[Number(btn.dataset.ri)])));
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
      const lines = cell.rows.map((r) => 'Commission ' + r['Commission no.']);
      if (cell.base) lines.push(t('hist.base'));
      else {
        for (const src of srcs) {
          const tag = both ? (src === 'sam' ? 'SAM ' : 'WINGS ') : '';
          const d = cell[src];
          if (d.added.length || d.removed.length) {
            d.added.forEach((c) => lines.push(tag + '+' + c));
            d.removed.forEach((c) => lines.push(tag + '-' + c));
          } else lines.push(tag + t('hist.same'));
        }
      }
      for (const src of srcs) {
        if (cell[src].varied.length) lines.push(histVariedText(cell[src].varied));
      }
      // 크로스체크(그 달 WINGS↔SAM 차이)도 함께 내보낸다.
      if (both && (cell.onlyW.length || cell.onlyS.length)) {
        lines.push(t('hist.xBad') + ' — ' + t('hist.xTitle', {
          w: cell.onlyW.join(', ') || '—', s: cell.onlyS.join(', ') || '—',
        }));
      } else if (both) lines.push(t('hist.xOk'));
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
      $('#histBasis').querySelectorAll('button').forEach((b2) =>
        b2.classList.toggle('active', b2 === btn));
      histRender();
    }));
  const ex = $('#histExport');
  if (ex) ex.addEventListener('click', histExportXlsx);

  histRender();
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
      if (cfg.onLoaded) cfg.onLoaded(S);
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

// ---- 인식모델_대조표: data.json 으로 만드는 확인용 시트 ----
// 로컬 파이썬 도구(build_model_rules_xlsx.py)가 만드는 시트지만, 브라우저에서 워크북을 열 때도
// 지금 데이터로 다시 만든다 — 그래야 대시보드·목록·편집 창의 항목과 값이 어긋나지 않는다.
// 컬럼 앞부분은 대시보드 목록과 같은 Model · Type · Axle · Cab · MY · PTO.
const RECOG_SHEET = '인식모델_대조표';
const RECOG_COLS = [
  ['Model(차종)', 'Vehicle'],
  ['Type(WINGS 모델)', 'Model(WINGS)'],
  ['Axle(축)', 'Type'],
  ['Cab(캡)', 'Cab'],
  ['MY', 'MY'],
  ['PTO', 'PTO'],
  ['SAM Baumuster(원본)', 'SAM Baumuster'],
  ['SAM now(수정)', 'SAM now'],
  ['Baumuster', 'Baumuster'],
  ['Subcategory', 'Subcategory (ID)'],
  ['매칭상태', 'SAM Status'],
  ['SAM 파일', 'Compared SAM file name'],
];
function recognitionAoa() {
  const seen = new Map();
  for (const r of DATA.rows) {
    if (!String(r['Model(WINGS)'] ?? '').trim()) continue;
    const row = RECOG_COLS.map(([, key]) => {
      let v = String(r[key] ?? '').trim();
      if (key === 'Compared SAM file name') v = v.split(/[\\/]/).pop();
      return v;
    });
    const k = row.join('');
    if (!seen.has(k)) seen.set(k, row);
  }
  const rows = [...seen.values()].sort((a, b) =>
    a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]) || a[11].localeCompare(b[11]));
  return [RECOG_COLS.map(([label]) => label)].concat(rows);
}
function refreshRecognitionSheet(S) {
  if (!DATA.rows.length || !S.sheets[RECOG_SHEET]) return;
  S.sheets[RECOG_SHEET] = recognitionAoa();
}

// ---- 모델 매칭: model_mapping.xlsx (03. model_rules) 전체 시트 편집 ----
const matchEditor = makeSheetEditor({
  folderKey: 'model_rules', fixedFile: MODEL_MAPPING_FILE,
  onLoaded: refreshRecognitionSheet,
  // 수동 교정 시트는 폐기했다 — 매칭은 SAM 문서의 번호·코드로만 결정한다.
  dropSheets: ['수동매핑', '매칭_별칭(수동)'],
  // 대조표는 데이터로 다시 만드는 보기라 편집해도 남지 않는다 — 편집 UI 자체를 숨긴다.
  readOnlySheets: [RECOG_SHEET],
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
  showMatchPane(localStorage.getItem('matchPane') || 'result');
  renderMatchSummary();
}

// 매칭 결과 / 대조표는 한 번에 하나만 — 보이는 쪽이 화면 높이를 다 쓴다.
function showMatchPane(name) {
  const which = name === 'sheet' ? 'sheet' : 'result';
  localStorage.setItem('matchPane', which);
  const sum = $('#matchSum'), ed = $('#matchEditor');
  if (sum) sum.classList.toggle('hidden', which !== 'result');
  if (ed) ed.classList.toggle('hidden', which !== 'sheet');
  const pane = $('#matchPane');
  if (pane) {
    pane.querySelectorAll('button[data-pane]').forEach((b) =>
      b.classList.toggle('active', b.dataset.pane === which));
  }
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
      const kind = k === 'Cab' ? 'cab' : (k === 'PTO' ? 'pto' : '');
      if (kind && pairMismatch(r, kind)) {
        return `<td class="pair-diff" title="${esc(pairTitle(r, kind))}">${msHl(v)}<span class="ne">≠</span></td>`;
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

// Export — 대시보드의 현재 필터/정렬 결과를 엑셀로 저장
const _exportBtn = $('#exportBtn');
if (_exportBtn) _exportBtn.addEventListener('click', exportFilteredXlsx);

// 저장 안 한 편집이 있으면 페이지 이탈 경고 (모델 매칭 / 코드 관리)
window.addEventListener('beforeunload', (e) => {
  if (codeEditor.S.dirty || matchEditor.S.dirty) { e.preventDefault(); e.returnValue = ''; }
});

function onManualFilter() {
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

applyStaticI18n();
enhanceFilterSelects();
load().catch((e) => {
  $('#meta').textContent = t('meta.loadFail', { err: e.message });
  const msg = $('#statusMsg');
  msg.textContent = t('msg.loadFail', { err: e.message });
  msg.classList.remove('hidden');
});
