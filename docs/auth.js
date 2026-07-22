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
  var APP_SCRIPT = 'app.js?v=20260722c';

  var pca = null;
  var activeAccount = null;

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
    };
  }

  // ---- 대시보드 로드 (인증 성공/스킵 후 호출) ------------------------
  function loadApp() {
    var s = document.createElement('script');
    s.src = APP_SCRIPT;
    document.body.appendChild(s);
  }

  // ---- 로그인 오버레이 ----------------------------------------------
  function showLogin(onClick, errMsg) {
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
    p.textContent = '회사 Microsoft 365 계정으로 로그인하세요.';
    card.appendChild(logo);
    card.appendChild(h1);
    card.appendChild(p);
    if (errMsg) {
      var err = document.createElement('p');
      err.className = 'auth-err';
      err.textContent = errMsg;
      card.appendChild(err);
    }
    var btn = document.createElement('button');
    btn.id = 'authLoginBtn';
    btn.className = 'auth-btn';
    btn.textContent = 'Microsoft 계정으로 로그인';
    if (onClick) {
      btn.addEventListener('click', function () {
        btn.disabled = true;
        btn.textContent = '로그인 창으로 이동 중…';
        onClick();
      });
    } else {
      btn.disabled = true;
    }
    card.appendChild(btn);
    ov.appendChild(card);
    document.body.appendChild(ov);
  }

  // ---- 로그인된 사용자 칩(상단 네비 우측) + 로그아웃 ------------------
  function showUserChip(account, onLogout) {
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
    if (!loginRequired()) {
      if (activeAccount) {
        showUserChip(activeAccount, function () { pca.logoutRedirect({ account: activeAccount }); });
      }
      loadApp();
      return;
    }

    // 보호 도메인: 계정이 없으면 로그인 강제.
    if (!activeAccount) {
      showLogin(function () { pca.loginRedirect(LOGIN_REQUEST); });
      return;
    }
    showUserChip(activeAccount, function () { pca.logoutRedirect({ account: activeAccount }); });
    loadApp();
  }

  // 상단 네비게이션(#navUser)이 DOM 에 준비된 뒤 실행.
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', main);
  } else {
    main();
  }
})();
