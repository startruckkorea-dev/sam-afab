// M365 (Microsoft Entra ID) 로그인 게이트 + Graph 토큰 브로커.
// ------------------------------------------------------------------
// 로그인 게이트는 회사 운영 도메인에서만 작동한다(PROTECTED_HOSTS).
//   - localhost / 127.0.0.1 / 개인 *.github.io  → 로그인 없이 바로 app.js 로드
//   - sam-afab.startruckkorea.com (회사 운영)    → M365 로그인 필수
//
// 또한 SharePoint(모델 매칭 / 코드 관리) 연동을 위해 window.MB_AUTH.getToken()
// 을 노출한다. graph.js 가 이걸로 Sites.ReadWrite.All 토큰을 얻어 SharePoint
// 문서 라이브러리의 Excel 을 읽고/쓴다. 첫 사용 시 팝업으로 증분 동의를 받는다.
//
// ⚠ 이 사이트는 GitHub Pages(정적)라서 로그인 게이트는 "화면 게이트(UX)"다.
//    docs/data.json 자체는 URL 을 알면 로그인 없이도 받아질 수 있다.
//
// 로그인/토큰이 작동하는 도메인은 Entra 앱 등록(SPA)에 Redirect URI 로 등록돼야 한다:
//   - https://sam-afab.startruckkorea.com/   (운영, 끝 슬래시 포함, 플랫폼 유형 SPA)
//   - http://localhost:8000/ 등 (로컬 개발에서 SharePoint 편집을 테스트할 때만)
(function () {
  'use strict';

  var PROTECTED_HOSTS = ['sam-afab.startruckkorea.com'];
  function loginRequired() {
    return PROTECTED_HOSTS.indexOf(window.location.hostname) !== -1;
  }

  var MSAL_CONFIG = {
    auth: {
      clientId: '9b247088-5afb-4622-9c5e-b5f27142761d',
      authority: 'https://login.microsoftonline.com/19cab1f5-21f4-44df-8ac6-96d6ca595203',
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: { cacheLocation: 'localStorage' },
  };
  var LOGIN_REQUEST = { scopes: ['User.Read'] };
  var APP_SCRIPT = 'app.js?v=20260909c';

  var pca = null;
  var activeAccount = null;
  var activeRole = null;   // 'admin' | 'read'
  var activeRoleVia = null; // 'share' | 'path' — 권한 명단을 어느 경로로 읽었는지

  // ---- Graph 토큰 브로커 (graph.js 가 사용) --------------------------
  // 조용히(acquireTokenSilent) 시도하고, 안 되면 팝업으로 증분 동의.
  async function getToken(scopes) {
    if (!pca) throw new Error('로그인이 초기화되지 않았습니다 (Microsoft 라이브러리 차단?).');
    var account = activeAccount || (pca.getAllAccounts()[0] || null);
    if (account) {
      try {
        var r = await pca.acquireTokenSilent({ scopes: scopes, account: account });
        return r.accessToken;
      } catch (e) { /* 동의 필요 등 → 팝업 폴백 */ }
    }
    var rp = await pca.acquireTokenPopup({ scopes: scopes });
    if (rp.account) { activeAccount = rp.account; pca.setActiveAccount(rp.account); }
    return rp.accessToken;
  }

  function exposeAuth() {
    window.MB_AUTH = {
      getToken: getToken,
      account: function () { return activeAccount; },
      pca: pca,
      role: activeRole,
    };
  }

  function setRole(role) {
    activeRole = role;
    if (window.MB_AUTH) {
      window.MB_AUTH.role = role;
      window.MB_AUTH.roleVia = activeRoleVia;
    }
    // 배포 후 "폴더 권한 없는 사람도 되는가" 를 콘솔에서 바로 확인할 수 있게 남긴다.
    // via=share 면 폴더 ACL 이 아니라 공유 링크로 읽은 것이다.
    if (activeRoleVia) console.info('[SAM-AFAB] 접속 권한 =', role, '· 명단 조회 경로 =', activeRoleVia);
  }

  // ---- 접속 권한 명단 (SharePoint: Shared Documents/SAM-AFAB_Access) ---
  // 폴더의 xlsx 를 읽어 로그인 계정의 권한(Admin / Read)을 정한다.
  // 열 위치를 고정하지 않고 헤더 행에서 EMail / UPN / Role 을 이름으로 찾으므로
  // 엑셀에 열이 추가되거나 순서가 바뀌어도 그대로 동작한다.
  var ACCESS_FOLDER = 'access';

  function norm(v) {
    return String(v == null ? '' : v).trim().toLowerCase();
  }

  function findCol(header, names) {
    for (var i = 0; i < header.length; i++) {
      if (names.indexOf(norm(header[i]).replace(/\s+/g, '')) !== -1) return i;
    }
    return -1;
  }

  // 시트 하나(AOA)에서 upn 에 해당하는 권한을 찾는다. 없으면 null.
  function roleFromRows(rows, upn) {
    if (!upn || !rows || rows.length < 2) return null;  // 빈 값이 빈 셀과 일치하면 안 된다
    var head = rows[0] || [];
    var cEmail = findCol(head, ['email', 'e-mail', '메일', '이메일']);
    var cUpn = findCol(head, ['upn']);
    var cRole = findCol(head, ['role', '권한']);
    if (cRole === -1 || (cEmail === -1 && cUpn === -1)) return null;
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r] || [];
      // UPN 열과 EMail 열 중 어느 쪽이든 일치하면 그 행으로 본다.
      // (UPN 열은 현재 비어 있고, 한쪽만 채워도 잠기지 않게 둘 다 본다.)
      var hitUpn = cUpn !== -1 && norm(row[cUpn]) === upn;
      var hitMail = cEmail !== -1 && norm(row[cEmail]) === upn;
      if (!hitUpn && !hitMail) continue;
      var role = norm(row[cRole]);
      if (role === 'admin') return 'admin';
      if (role.indexOf('read') === 0) return 'read';
    }
    return null;
  }

  // 폴더 안에서 쓸 권한 엑셀 하나를 고른다 (Access* 로 시작하는 파일 우선 —
  // mb-truck-spec 의 accessList.js 와 같은 규칙). 엑셀 임시 파일은 건너뛴다.
  function pickAccessXlsx(kids) {
    var pick = null;
    for (var i = 0; i < kids.length; i++) {
      var n = kids[i].name || '';
      if (!kids[i].file || n.indexOf('~$') === 0 || !/\.xlsx$/i.test(n)) continue;
      if (/^access/i.test(n)) return kids[i];
      if (!pick) pick = kids[i];
    }
    return pick;
  }

  // 공유 링크에서 워크북의 base 경로를 얻는다.
  //  - 링크가 파일이면 /shares 경로를 그대로 쓴다 → driveId 조회조차 필요 없다.
  //  - 링크가 폴더면 그 폴더의 children 을 /shares 로 훑는다. 폴더 ACL 이 아니라
  //    링크가 권한이므로, 폴더 접근 권한이 없는 사람도 여기까지 통과한다.
  async function baseFromShare(shareUrl) {
    var sb = window.Graph.shareBase(shareUrl);
    var item = await window.Graph.itemAt(sb);
    if (!item.folder) return sb;
    var pick = pickAccessXlsx(await window.Graph.childrenAt(sb));
    if (!pick) throw new Error('공유된 폴더에 권한 엑셀(.xlsx)이 없습니다.');
    var driveId = pick.parentReference && pick.parentReference.driveId;
    if (!driveId) {
      throw new Error('폴더 공유 링크로는 파일 주소를 얻지 못했습니다. '
        + '폴더 대신 권한 엑셀 "파일"의 공유 링크를 shareUrl 에 넣으세요.');
    }
    return window.Graph.itemBase(driveId, pick.id);
  }

  // 폴더 경로로 직접 찾는 폴백. 이쪽은 폴더 읽기 권한이 있어야 한다.
  async function baseFromPath() {
    var pb = await window.Graph.pathBase(ACCESS_FOLDER);
    var item = await window.Graph.itemAt(pb);
    if (!item.folder) return pb;
    var pick = pickAccessXlsx(await window.Graph.childrenAt(pb));
    if (!pick) throw new Error('SAM-AFAB_Access 폴더에 권한 엑셀(.xlsx)이 없습니다.');
    return await window.Graph.pathBase(ACCESS_FOLDER, pick.name);
  }

  // 권한 엑셀 위치: 공유 링크 우선(폴더 권한 불필요), 실패하면 폴더 경로.
  // 어느 경로로 읽었는지 activeRoleVia 에 남겨 두면 배포 후 확인이 쉽다.
  async function resolveAccessBase() {
    var cfg = (window.Graph.folders && window.Graph.folders[ACCESS_FOLDER]) || {};
    var errs = [];
    if (cfg.shareUrl) {
      try {
        var b = await baseFromShare(cfg.shareUrl);
        activeRoleVia = 'share';
        return b;
      } catch (e) {
        errs.push('공유 링크: ' + ((e && e.message) || e));
      }
    }
    try {
      var b2 = await baseFromPath();
      activeRoleVia = 'path';
      return b2;
    } catch (e) {
      errs.push('폴더 경로: ' + ((e && e.message) || e));
    }
    throw new Error(errs.join('  /  '));
  }

  // 시트를 AOA 로 반환. 워크북 API 라 파일을 통째로 내려받지 않고 SheetJS 도 안 쓴다.
  async function fetchAccessSheets() {
    if (!window.Graph || !window.Graph.available()) {
      throw new Error('SharePoint 연동을 사용할 수 없습니다.');
    }
    return window.Graph.workbookSheets(await resolveAccessBase());
  }

  async function resolveRole(account) {
    var upn = norm(account && account.username);
    if (!upn) return null;
    var sheets = await fetchAccessSheets();
    for (var i = 0; i < sheets.length; i++) {
      var role = roleFromRows(sheets[i], upn);
      if (role) return role;
    }
    return null;
  }

  // ---- 대시보드 로드 (인증 성공/스킵 후 호출) ------------------------
  function loadApp() {
    var s = document.createElement('script');
    s.src = APP_SCRIPT;
    document.body.appendChild(s);
  }

  // ---- 전체화면 오버레이 (로그인 / 권한 없음 / 조회 실패 공용) --------
  function showOverlay(opts) {
    var prev = document.getElementById('authOverlay');
    if (prev) prev.remove();
    var ov = document.createElement('div');
    ov.id = 'authOverlay';
    var card = document.createElement('div');
    card.className = 'auth-card';
    var logo = document.createElement('img');
    logo.className = 'auth-logo';
    logo.src = 'logo.png';
    logo.alt = 'Mercedes-Benz · Trucks you can trust';
    var h1 = document.createElement('h1');
    h1.textContent = 'SAM × AFAB Comparison';
    var p = document.createElement('p');
    p.textContent = opts.desc;
    card.appendChild(logo);
    card.appendChild(h1);
    card.appendChild(p);
    if (opts.errMsg) {
      var err = document.createElement('p');
      err.className = 'auth-err';
      err.textContent = opts.errMsg;
      card.appendChild(err);
    }
    var btn = document.createElement('button');
    btn.id = 'authLoginBtn';
    btn.className = 'auth-btn';
    btn.textContent = opts.btnText;
    if (opts.onClick) {
      btn.addEventListener('click', function () {
        btn.disabled = true;
        if (opts.busyText) btn.textContent = opts.busyText;
        opts.onClick();
      });
    } else {
      btn.disabled = true;
    }
    card.appendChild(btn);
    ov.appendChild(card);
    document.body.appendChild(ov);
  }

  function showLogin(onClick, errMsg) {
    showOverlay({
      desc: '회사 Microsoft 365 계정으로 로그인하세요.',
      errMsg: errMsg,
      btnText: 'Microsoft 계정으로 로그인',
      busyText: '로그인 창으로 이동 중…',
      onClick: onClick,
    });
  }

  // 로그인은 됐지만 권한 명단에 없는 계정.
  function showDenied(account, onLogout) {
    var who = (account && account.username) ? account.username + ' 계정은 ' : '';
    showOverlay({
      desc: who + 'SAM × AFAB 시스템 접속 권한이 없습니다. 관리자에게 권한 등록을 요청하세요.',
      btnText: '다른 계정으로 로그인',
      busyText: '로그아웃 중…',
      onClick: onLogout,
    });
  }

  // 권한 명단 자체를 읽지 못한 경우 — 통과시키지 않고 원인을 보여준다.
  function showAccessError(errMsg, onRetry) {
    showOverlay({
      desc: '접속 권한 명단을 확인하지 못했습니다. 잠시 후 다시 시도하고, 계속 같은 화면이면 '
        + '아래 메시지를 그대로 관리자에게 전달하세요.',
      errMsg: errMsg,
      btnText: '다시 시도',
      busyText: '확인 중…',
      onClick: onRetry,
    });
  }

  // ---- 로그인된 사용자 칩(상단 네비 우측) + 로그아웃 ------------------
  function showUserChip(account, onLogout, role) {
    var slot = document.getElementById('navUser') || document.querySelector('header');
    if (!slot) return;
    var chip = document.createElement('div');
    chip.className = 'auth-chip';
    var user = document.createElement('span');
    user.className = 'auth-user';
    user.textContent = account.name || account.username || '';
    user.title = account.username || '';
    var out = document.createElement('button');
    out.className = 'icon-btn';
    out.title = '로그아웃';
    out.textContent = '⎋ 로그아웃';
    out.addEventListener('click', onLogout);
    chip.appendChild(user);
    // 현재 권한을 이름 옆에 항상 표시한다 (엑셀의 Role 값과 같은 표기).
    if (role === 'admin' || role === 'read') {
      var badge = document.createElement('span');
      badge.className = 'auth-badge ' + role;
      badge.textContent = role === 'admin' ? 'Admin' : 'Read';
      badge.title = role === 'admin'
        ? '전체 권한 — 모델 매칭 · 코드 관리 · 데이터 다시 계산 가능'
        : '조회 전용 — 모델 매칭 · 코드 관리 · 데이터 다시 계산 불가';
      chip.appendChild(badge);
    }
    chip.appendChild(out);
    slot.appendChild(chip);
  }

  // ---- 메인 흐름 -----------------------------------------------------
  async function main() {
    // MSAL 라이브러리 자체가 없으면(CDN·self-host 모두 차단):
    if (typeof msal === 'undefined') {
      if (loginRequired()) {
        showLogin(null, 'Microsoft 로그인 라이브러리를 불러오지 못했습니다 (네트워크/CDN 차단). 관리자에게 문의하세요.');
      } else {
        setRole('admin');
        loadApp();  // 로컬 등: 로그인 없이 대시보드만 (SharePoint 편집은 비활성)
      }
      return;
    }

    pca = new msal.PublicClientApplication(MSAL_CONFIG);
    if (typeof pca.initialize === 'function') {
      try { await pca.initialize(); } catch (e) { /* v2 등 */ }
    }
    exposeAuth();  // 토큰 브로커는 게이트 여부와 무관하게 항상 노출

    // 리다이렉트 응답 처리 (로그인 게이트 도메인에서 loginRedirect 후 복귀)
    try {
      var resp = await pca.handleRedirectPromise();
      if (resp && resp.account) activeAccount = resp.account;
    } catch (e) {
      console.error('MSAL redirect 처리 오류', e);
      if (loginRequired()) {
        showLogin(function () { pca.loginRedirect(LOGIN_REQUEST); },
          '로그인 처리 중 오류가 발생했습니다. 다시 시도하세요: ' + (e && e.message ? e.message : e));
        return;
      }
    }

    if (!activeAccount) {
      var accts = pca.getAllAccounts();
      if (accts && accts.length) activeAccount = accts[0];
    }
    if (activeAccount) pca.setActiveAccount(activeAccount);

    // 보호 대상이 아니면(로컬·개인 github.io) 게이트 없이 로드.
    // 개발 편의를 위해 권한 조회도 건너뛰고 admin 으로 둔다.
    if (!loginRequired()) {
      setRole('admin');
      if (activeAccount) {
        showUserChip(activeAccount, function () { pca.logoutRedirect({ account: activeAccount }); }, 'admin');
      }
      loadApp();
      return;
    }

    // 보호 도메인: 계정이 없으면 로그인 강제.
    if (!activeAccount) {
      showLogin(function () { pca.loginRedirect(LOGIN_REQUEST); });
      return;
    }

    // 접속 권한 명단(SharePoint 엑셀) 확인. 매 로그인마다 조회하므로 캐시가 없고,
    // 엑셀을 고치면 재로그인 시 바로 반영된다.
    var role = null;
    try {
      role = await resolveRole(activeAccount);
    } catch (e) {
      console.error('접속 권한 확인 실패', e);
      // 명단을 못 읽으면 통과시키지 않는다(fail-closed). 원인은 그대로 보여준다.
      showAccessError(e && e.message ? e.message : String(e), function () { window.location.reload(); });
      return;
    }
    if (!role) {
      showDenied(activeAccount, function () { pca.logoutRedirect({ account: activeAccount }); });
      return;
    }

    setRole(role);
    showUserChip(activeAccount, function () { pca.logoutRedirect({ account: activeAccount }); }, role);
    loadApp();
  }

  // 상단 네비게이션(#navUser)이 DOM 에 준비된 뒤 실행.
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', main);
  } else {
    main();
  }
})();
