// M365 (Microsoft Entra ID) 로그인 게이트.
// ------------------------------------------------------------------
// 인증되지 않은 사용자에게는 로그인 화면만 보여주고, 로그인에 성공하면
// 비로소 app.js 를 동적으로 로드해 대시보드를 렌더링한다.
//
// ⚠ 이 사이트는 GitHub Pages(정적)라서 이 로그인은 "화면 게이트(UX)"다.
//    docs/data.json 자체는 URL 을 알면 로그인 없이도 받아질 수 있다.
//    데이터 자체 접근 제어가 필요하면 별도 게이트웨이(호스팅) 가 필요하다.
//
// Entra 앱 등록(SPA)에 아래 Redirect URI 들을 등록해야 동작한다:
//   - https://sam-afab.startruckkorea.com/            (운영)
//   - https://sunghanchostk.github.io/SAM_AFAB_Github/ (개인 테스트)
//   플랫폼 유형은 반드시 "Single-page application(SPA)" 여야 한다(암시적 X, PKCE O).
(function () {
  'use strict';

  var MSAL_CONFIG = {
    auth: {
      clientId: '9b247088-5afb-4622-9c5e-b5f27142761d',
      authority: 'https://login.microsoftonline.com/19cab1f5-21f4-44df-8ac6-96d6ca595203',
      // 현재 페이지 URL(해시/쿼리 제외)을 Redirect URI 로 사용한다.
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: { cacheLocation: 'localStorage' },
  };
  var LOGIN_REQUEST = { scopes: ['User.Read'] };
  var APP_SCRIPT = 'app.js?v=20260721a';

  // ---- 대시보드 로드 (인증 성공 후에만 호출) --------------------------
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

  // ---- 로그인된 사용자 칩(헤더 우측) + 로그아웃 ----------------------
  function showUserChip(account, onLogout) {
    var header = document.querySelector('header');
    if (!header) return;
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
    header.appendChild(chip);
  }

  // ---- 메인 흐름 -----------------------------------------------------
  async function main() {
    if (typeof msal === 'undefined') {
      // MSAL 라이브러리(CDN) 로드 실패 — 회사망 차단 등.
      showLogin(null, 'Microsoft 로그인 라이브러리를 불러오지 못했습니다 (네트워크/CDN 차단). 관리자에게 문의하세요.');
      return;
    }

    var pca = new msal.PublicClientApplication(MSAL_CONFIG);
    // MSAL v3+ 는 initialize() 가 필요, v2 는 없음 — 둘 다 지원.
    if (typeof pca.initialize === 'function') {
      try { await pca.initialize(); } catch (e) { /* v2 등 */ }
    }

    var account = null;
    try {
      var resp = await pca.handleRedirectPromise();
      if (resp && resp.account) account = resp.account;
    } catch (e) {
      console.error('MSAL redirect 처리 오류', e);
      showLogin(function () { pca.loginRedirect(LOGIN_REQUEST); },
        '로그인 처리 중 오류가 발생했습니다. 다시 시도하세요: ' + (e && e.message ? e.message : e));
      return;
    }

    if (!account) {
      var accts = pca.getAllAccounts();
      if (accts && accts.length) account = accts[0];
    }

    if (!account) {
      showLogin(function () { pca.loginRedirect(LOGIN_REQUEST); });
      return;
    }

    pca.setActiveAccount(account);
    showUserChip(account, function () {
      pca.logoutRedirect({ account: account });
    });
    loadApp();
  }

  main();
})();
