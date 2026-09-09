const fs = require('fs'), path = require('path');
const { JSDOM } = require('jsdom');
const ROOT = 'C:/Users/sunghan.cho/Documents/AI2/SAM_AFAB_Github/docs';
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');

const dom = new JSDOM(html, { runScripts: 'outside-only', url: 'http://localhost:8000/' });
const { window } = dom;
global.window = window; global.document = window.document;
window.alert = (m) => { console.log('  [alert]', m); };
window.confirm = () => true;

// 페이지가 실제로 하는 일: data.json / codes.json / rules.json 을 fetch 한다.
const files = { 'data.json': 'data.json', 'codes.json': 'codes.json', 'rules.json': 'rules.json' };
window.fetch = async (u) => {
  const name = Object.keys(files).find((f) => String(u).includes(f));
  if (!name) throw new Error('unexpected fetch ' + u);
  const body = fs.readFileSync(path.join(ROOT, files[name]), 'utf8');
  return { ok: true, status: 200, json: async () => JSON.parse(body), text: async () => body };
};

// app.js 는 auth.js 가 role 을 정한 뒤 로드된다. 여기서는 두 역할을 각각 흉내낸다.
const ROLE = process.argv[2] || 'admin';
window.MB_AUTH = { role: ROLE, getToken: async () => 'x', account: () => ({ username: 'a@b.c' }) };

const code = fs.readFileSync(path.join(ROOT, 'app.js'), 'utf8');
window.eval(code);

(async () => {
  await new Promise((r) => setTimeout(r, 400));   // load() 완료 대기
  const $ = (s) => window.document.querySelector(s);
  const txt = (s) => ($(s) ? $(s).textContent.trim() : '(없음)');
  const fail = [];
  // 기대값은 data.json 에서 직접 센다 (하드코딩하지 않는다)
  const raw = JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
  const expectFrom = (f) => raw.rows.filter((r) => {
    const m = String(r['Production date'] || '').slice(0, 7); return m && m >= f; }).length;
  const check = (ok, label, extra) => { console.log((ok ? '  PASS  ' : '  FAIL  ') + label + (extra ? '  ' + extra : '')); if (!ok) fail.push(label); };

  console.log('\n=== role =', ROLE, '===');
  const total = () => txt('#summary .t-total .n');
  const cap = () => txt('#summary .dash-row .dash-cap');
  const rows = () => window.document.querySelectorAll('#grid tbody tr').length;

  console.log('[기본 상태]', 'caption=' + cap(), 'total=' + total(), 'rows=' + rows());
  check(total() === String(expectFrom('2026-09')), '기본은 이번 달부터 = ' + expectFrom('2026-09'), 'total=' + total());
  check(/2026-09/.test(cap()), '캡션에 시작월 표시', cap());

  // 시작월 드롭다운이 채워졌는가
  const opts = [...$('#fromMonth').options].map((o) => o.value);
  check(opts.length === 12 && opts[0] === '2026-12' && opts[opts.length - 1] === '2026-01',
    '시작월 옵션 12개(내림차순)', opts.join(','));

  // 토글 → 올해 1월부터
  $('#yearStartBtn').dispatchEvent(new window.Event('click'));
  console.log('[토글 ON  ]', 'caption=' + cap(), 'total=' + total(), 'rows=' + rows());
  check(total() === String(expectFrom('2026-01')), '올해 1월부터 = ' + expectFrom('2026-01')
    + ' (과거 ' + (expectFrom('2026-01') - expectFrom('2026-09')) + '대 추가)', 'total=' + total());
  check($('#yearStartBtn').classList.contains('on'), '토글 눌림 표시');
  check($('#fromMonth').value === '2026-01', '드롭다운도 2026-01 로 동기화', $('#fromMonth').value);
  check(/2026-01/.test(cap()), '캡션이 2026-01 로 갱신', cap());

  // 토글 → 다시 이번 달
  $('#yearStartBtn').dispatchEvent(new window.Event('click'));
  check(total() === String(expectFrom('2026-09')) && !$('#yearStartBtn').classList.contains('on'), '토글 OFF 복귀', 'total=' + total());

  // 드롭다운으로 임의 시작월 선택
  $('#fromMonth').value = '2026-05';
  $('#fromMonth').dispatchEvent(new window.Event('input'));
  console.log('[2026-05  ]', 'caption=' + cap(), 'total=' + total(), 'rows=' + rows());
  check(total() === String(expectFrom('2026-05')), '2026-05 부터 = ' + expectFrom('2026-05'), 'total=' + total());
  check(!$('#yearStartBtn').classList.contains('on'), '임의 월 선택 시 토글은 꺼짐');
  check(rows() === expectFrom('2026-05'), '표 건수 = 타일 숫자', 'rows=' + rows());

  // 언어 전환 후에도 시작월 상태가 유지되는가
  $('#langBtn').dispatchEvent(new window.Event('click'));
  check($('#fromMonth').value === '2026-05', 'EN 전환 후 시작월 유지', $('#fromMonth').value);
  check($('#fromMonth').options[0].textContent === 'From 2026-12', 'EN 옵션 라벨 번역', $('#fromMonth').options[0].textContent);
  check(/2026-05/.test(cap()) && /Overall/.test(cap()), 'EN 캡션에 시작월 유지', cap());
  check(total() === String(expectFrom('2026-05')), 'EN 전환 후 집계 동일', 'total=' + total());
  $('#langBtn').dispatchEvent(new window.Event('click'));

  // 권한 UI
  const locked = ['matching', 'codes'].map((v) => $('.nav-link[data-view="' + v + '"]'));
  const dis = locked.every((el) => el.classList.contains('disabled'));
  const build = $('#buildBtn').disabled;
  if (ROLE === 'admin') {
    check(!dis && !build, 'Admin: 탭·재계산 모두 활성');
  } else {
    check(dis && build, 'Read: 탭 회색 + 재계산 비활성');
    const before = window.document.querySelector('#view-codes').classList.contains('active');
    locked[1].dispatchEvent(new window.Event('click', { bubbles: true }));
    check(window.document.querySelector('#view-codes').classList.contains('active') === before,
      'Read: 코드 관리 클릭해도 전환 안 됨');
  }
  console.log(fail.length ? '\n>>> 실패 ' + fail.length + '건' : '\n>>> 전부 통과');
  process.exit(fail.length ? 1 : 0);
})().catch((e) => { console.error('ERROR', e); process.exit(2); });
