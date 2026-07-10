"""
WINGS 자동 다운로드 모듈
로컬 Chrome (WingsAutomation 프로필)을 이용해 WINGS Extended Search에서
Requested delivery date 조건으로 Excel 파일을 자동 다운로드합니다.

사용법:
    from wings_scraper import download_wings_excel
    path = download_wings_excel(["2026-04"], on_status=print)
    path = download_wings_excel(["2026-04", "2026-05"], on_status=print)
"""

import os
import sys
import glob
import json
import asyncio
import tempfile
import subprocess
import threading
import concurrent.futures
from playwright.async_api import async_playwright

WINGS_URL = "https://wings.tsac.daimlertruck.com/sites/main.jsp"

# 전용 Chrome 프로필 디렉터리 (사용자의 메인 Chrome과 충돌 방지)
WINGS_PROFILE_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "Google", "Chrome", "User Data", "WingsAutomation",
)

# 자격 증명 파일 경로
_CREDENTIALS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".wings_credentials")


def _load_credentials() -> tuple:
    """로컬 .wings_credentials 파일에서 이메일/비밀번호를 읽는다."""
    try:
        with open(_CREDENTIALS_FILE, "r", encoding="utf-8") as f:
            lines = f.read().strip().splitlines()
            if len(lines) >= 2:
                return lines[0].strip(), lines[1].strip()
    except FileNotFoundError:
        pass
    return None, None


def _force_show_chrome_window(chrome_pid=None):
    """WingsAutomation Chrome 창을 강제로 앞으로 꺼낸다."""
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32

        # 대상 PID 수집
        pids = set()
        if chrome_pid:
            pids.add(chrome_pid)
            # Chrome 자식 프로세스도 포함 (renderer 등은 창 없지만 browser process는 창 있음)
            try:
                _r = subprocess.run(
                    ['powershell', '-Command',
                     f'Get-CimInstance Win32_Process | Where-Object {{$_.ParentProcessId -eq {chrome_pid}}} | Select-Object -ExpandProperty ProcessId'],
                    capture_output=True, text=True, timeout=5)
                for _l in _r.stdout.strip().splitlines():
                    if _l.strip().isdigit():
                        pids.add(int(_l.strip()))
            except Exception:
                pass
        else:
            # PID 없으면 WingsAutomation 프로필로 검색
            try:
                _r = subprocess.run(
                    ['powershell', '-Command',
                     "Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | "
                     "Where-Object {$_.CommandLine -like '*WingsAutomation*'} | "
                     "Select-Object -ExpandProperty ProcessId"],
                    capture_output=True, text=True, timeout=10)
                for _l in _r.stdout.strip().splitlines():
                    if _l.strip().isdigit():
                        pids.add(int(_l.strip()))
            except Exception:
                pass

        if not pids:
            return

        found = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        def enum_cb(hwnd, _):
            pid = ctypes.c_ulong()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value not in pids:
                return True
            cls = ctypes.create_unicode_buffer(64)
            user32.GetClassNameW(hwnd, cls, 64)
            # Chrome 브라우저 창의 class name
            if 'Chrome_WidgetWin_1' in cls.value:
                found.append(hwnd)
            return True

        user32.EnumWindows(enum_cb, 0)

        for hwnd in found:
            user32.ShowWindow(hwnd, 3)   # SW_SHOWMAXIMIZED
            # keybd_event(Alt) 트릭: 포그라운드 권한 강제 획득 후 SetForegroundWindow
            user32.keybd_event(0x12, 0, 0, 0)    # VK_MENU down
            user32.SetForegroundWindow(hwnd)
            user32.keybd_event(0x12, 0, 0x0002, 0)  # VK_MENU up (KEYEVENTF_KEYUP)
            user32.BringWindowToTop(hwnd)
    except Exception:
        pass


def _release_profile_lock():
    """WingsAutomation 프로필을 점유 중인 Chrome을 종료하고 락 파일을 제거한다."""
    # PowerShell로 WingsAutomation 프로필 사용 중인 Chrome 프로세스 종료
    try:
        subprocess.run(
            ['powershell', '-Command',
             'Get-CimInstance Win32_Process -Filter "Name=\'chrome.exe\'" | '
             'Where-Object {$_.CommandLine -like \'*WingsAutomation*\'} | '
             'ForEach-Object {Stop-Process -Id $_.ProcessId -Force}'],
            capture_output=True, timeout=10,
        )
    except Exception:
        pass

    # Singleton 락 파일 제거 (Chrome이 비정상 종료 시 남기는 파일)
    for fname in ('SingletonLock', 'SingletonCookie', 'SingletonSocket'):
        fpath = os.path.join(WINGS_PROFILE_DIR, fname)
        try:
            if os.path.exists(fpath):
                os.remove(fpath)
        except Exception:
            pass


def _patch_chrome_prefs(profile_dir: str, download_dir: str) -> None:
    """WingsAutomation 프로필의 Preferences를 수정해 자동 다운로드를 허용한다.

    신규 Chrome 프로필은 기본적으로 ▲여러 파일 자동 다운로드 차단
    ▲저장 위치 묻기 프롬프트가 켜져 있어, Export 클릭 후 다운로드 창이
    잠깐 떴다가 사라지는 현상이 발생한다. 프로필의 Default/Preferences
    JSON을 미리 수정해 이 둘을 모두 끈다.
    """
    prefs_path = os.path.join(profile_dir, "Default", "Preferences")
    os.makedirs(os.path.dirname(prefs_path), exist_ok=True)

    try:
        with open(prefs_path, "r", encoding="utf-8") as f:
            prefs = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        prefs = {}

    dl = prefs.setdefault("download", {})
    dl["default_directory"] = os.path.abspath(download_dir)
    dl["prompt_for_download"] = False
    dl["directory_upgrade"] = True

    profile_section = prefs.setdefault("profile", {})
    content_settings = profile_section.setdefault("default_content_setting_values", {})
    content_settings["automatic_downloads"] = 1  # 1 = allow

    try:
        with open(prefs_path, "w", encoding="utf-8") as f:
            json.dump(prefs, f)
    except OSError:
        pass


def _find_chrome_exe() -> str | None:
    """시스템에 설치된 Google Chrome 실행 파일 경로를 찾는다."""
    candidates = [
        os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"),
                     r"Google\Chrome\Application\chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
                     r"Google\Chrome\Application\chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""),
                     r"Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def _are_consecutive(months_sorted: list) -> bool:
    """정렬된 'YYYY-MM' 리스트가 월 단위로 연속인지 확인한다."""
    for i in range(len(months_sorted) - 1):
        y1, m1 = map(int, months_sorted[i].split('-'))
        y2, m2 = map(int, months_sorted[i + 1].split('-'))
        next_m, next_y = (m1 + 1, y1) if m1 < 12 else (1, y1 + 1)
        if not (y2 == next_y and m2 == next_m):
            return False
    return True


async def _wait_filter_rows(page, min_count: int, timeout_ms: int = 15000) -> int:
    """WINGS Extended Search의 FilterCriteriaWidget이 min_count개 이상
    dijit.registry에 등록될 때까지 폴링 대기한다.

    'Remove all filter criteria' 클릭 후 WINGS가 빈 행을 재생성하거나,
    'New criteria' / 복사 버튼으로 행을 추가할 때 등록이 비동기적으로 일어나기 때문에,
    이 대기가 없으면 `_set_filter_row`에서 `rows[idx] === undefined`로 바로 실패한다.
    """
    import time as _t
    deadline = _t.monotonic() + timeout_ms / 1000.0
    last = 0
    while _t.monotonic() < deadline:
        try:
            last = await page.evaluate(
                """() => (typeof dijit !== 'undefined' && dijit.registry)
                    ? dijit.registry.toArray().filter(w =>
                        w.declaredClass === 'com.daimler.wings.view.grid.filter.FilterCriteriaWidget'
                      ).length
                    : 0"""
            )
            if last >= min_count:
                return last
        except Exception:
            pass
        await page.wait_for_timeout(300)
    return last


async def _copy_filter_row(page, row_idx: int):
    """FilterCriteriaWidget row_idx의 복사(📋) 버튼을 클릭해 새 행을 추가한다."""
    copy_bbox = await page.evaluate(
        """idx => {
            const rows = dijit.registry.toArray().filter(w =>
                w.declaredClass === 'com.daimler.wings.view.grid.filter.FilterCriteriaWidget'
            );
            if (!rows[idx]) return null;
            const buttons = dijit.registry.findWidgets(rows[idx].domNode)
                .filter(w => w.declaredClass.includes('Button'));
            let btn = null;
            for (const b of buttons) {
                const html  = (b.domNode.innerHTML || '').toLowerCase();
                const cls   = (b.domNode.className || '').toLowerCase();
                const title = (b.domNode.title || b.label || b.title || '').toLowerCase();
                if (html.includes('copy') || cls.includes('copy') ||
                    html.includes('clone') || html.includes('duplicate') ||
                    title.includes('copy') || title.includes('clone') ||
                    title.includes('duplicate')) {
                    btn = b; break;
                }
            }
            if (!btn && buttons.length >= 3) btn = buttons[2];
            if (!btn) return null;
            const r = btn.domNode.getBoundingClientRect();
            return {x: r.x + r.width / 2, y: r.y + r.height / 2,
                    scrollX: window.scrollX, scrollY: window.scrollY};
        }""",
        row_idx,
    )
    if copy_bbox:
        await page.mouse.click(
            copy_bbox["x"] + copy_bbox["scrollX"],
            copy_bbox["y"] + copy_bbox["scrollY"],
        )
    else:
        await page.click("text=New criteria")
    await page.wait_for_timeout(1000)


async def _set_all_row_connectors(page, connector: str = "or"):
    """필터 행 사이의 모든 and/or 커넥터를 지정한 값으로 변경한다.

    비연속 월 검색 시 'and' → 'or' 변경에 사용한다.
    커넥터 위젯은 현재 값이 'and' 또는 'or'인 dijit 위젯으로 탐지한다.
    """
    conn_bboxes = await page.evaluate(
        """() => {
            const result = [];
            for (const w of dijit.registry.toArray()) {
                try {
                    const val = w.get ? w.get('value') : null;
                    if (val !== 'and' && val !== 'or') continue;
                    if (!w.domNode) continue;
                    const r = w.domNode.getBoundingClientRect();
                    // ▼ 버튼: 오른쪽 끝에서 10px 안쪽
                    result.push({x: r.x + r.width - 10, y: r.y + r.height / 2,
                                  scrollX: window.scrollX, scrollY: window.scrollY,
                                  id: w.id, currentVal: val});
                } catch (e) {}
            }
            return result;
        }"""
    )
    for bbox in conn_bboxes:
        await page.mouse.click(
            bbox["x"] + bbox["scrollX"],
            bbox["y"] + bbox["scrollY"],
        )
        await page.wait_for_timeout(800)
        await _click_popup_item_by_text_playwright(page, connector)
        await page.wait_for_timeout(500)


async def _wings_download_async(months: list, download_dir: str, on_status=None, auth_code_callback=None) -> str:

    months_sorted = sorted(months)
    start_date = months_sorted[0] + "-01"
    end_date = months_sorted[-1] + "-01"
    single = len(months_sorted) == 1
    consecutive = _are_consecutive(months_sorted)

    os.makedirs(WINGS_PROFILE_DIR, exist_ok=True)
    os.makedirs(download_dir, exist_ok=True)

    def status(msg: str):
        if on_status:
            on_status(msg)

    # 이전 세션에서 잠긴 프로필 해제
    _release_profile_lock()

    # 신규/기존 프로필 모두에 자동 다운로드 허용 플래그 주입
    _patch_chrome_prefs(WINGS_PROFILE_DIR, download_dir)

    chrome_exe = _find_chrome_exe()
    if not chrome_exe:
        raise RuntimeError("Google Chrome이 설치되어 있지 않습니다.")

    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            WINGS_PROFILE_DIR,
            executable_path=chrome_exe,
            headless=False,
            accept_downloads=True,
            downloads_path=download_dir,
            args=[
                "--start-maximized",
                "--disable-popup-blocking",
                "--safebrowsing-disable-download-protection",
            ],
            viewport=None,
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        # ── 1. WINGS 접속 ──────────────────────────────────────────────────────
        status("WINGS에 접속 중...")
        await page.goto(WINGS_URL, wait_until="networkidle", timeout=30000)

        # 로그인이 필요한 경우
        login_needed = False
        try:
            login_needed = (
                await page.locator("input[type='password']").count() > 0
                or await page.locator("input[type='email']").count() > 0
                or await page.locator("input[placeholder*='Email'], input[placeholder*='User ID']").count() > 0
                or "login" in page.url.lower()
                or "businessid" in page.url.lower()
                or "microsoftonline" in page.url.lower()
            )
        except Exception:
            pass

        if login_needed:
            email, password = _load_credentials()
            auto_login = email and password

            if auto_login:
                status("자동 로그인 시도 중...")
            else:
                status("로그인이 필요합니다. 브라우저에서 아이디/비밀번호를 입력해 주세요...")

            # ── Daimler Truck Business ID 로그인 ──
            if auto_login:
                # 1) 이메일/User ID 입력 + Continue
                try:
                    # Daimler Business ID 페이지의 입력 필드
                    await page.wait_for_timeout(2000)
                    # JS로 직접 입력 필드 찾기 (placeholder 기반)
                    filled = await page.evaluate(
                        """(email) => {
                            const inputs = document.querySelectorAll('input');
                            for (const inp of inputs) {
                                if (inp.placeholder && (inp.placeholder.includes('Email') || inp.placeholder.includes('User ID'))) {
                                    inp.focus();
                                    inp.value = email;
                                    inp.dispatchEvent(new Event('input', {bubbles: true}));
                                    inp.dispatchEvent(new Event('change', {bubbles: true}));
                                    return true;
                                }
                            }
                            // fallback: 보이는 text/email input
                            for (const inp of inputs) {
                                if ((inp.type === 'text' || inp.type === 'email') && inp.offsetParent !== null) {
                                    inp.focus();
                                    inp.value = email;
                                    inp.dispatchEvent(new Event('input', {bubbles: true}));
                                    inp.dispatchEvent(new Event('change', {bubbles: true}));
                                    return true;
                                }
                            }
                            return false;
                        }""",
                        email,
                    )
                    if filled:
                        await page.wait_for_timeout(500)
                        # Continue 버튼 클릭
                        clicked = await page.evaluate(
                            """() => {
                                const btns = document.querySelectorAll('button, input[type="submit"]');
                                for (const b of btns) {
                                    const txt = (b.textContent || b.value || '').trim();
                                    if (txt === 'Continue' || txt === 'Next' || txt === '계속') {
                                        b.click();
                                        return txt;
                                    }
                                }
                                return null;
                            }"""
                        )
                        status(f"이메일 입력 완료, Continue 클릭: {clicked}")
                        await page.wait_for_timeout(4000)
                except Exception as e:
                    status(f"이메일 입력 실패: {e}")

                # 2) 비밀번호 입력 (페이지 전환 후)
                try:
                    for _ in range(15):  # 최대 7.5초 대기
                        pw_visible = await page.evaluate(
                            """() => {
                                const pw = document.querySelector('input[type="password"]');
                                return pw && pw.offsetParent !== null;
                            }"""
                        )
                        if pw_visible:
                            break
                        await page.wait_for_timeout(500)

                    if pw_visible:
                        await page.evaluate(
                            """(pwd) => {
                                const pw = document.querySelector('input[type="password"]');
                                if (pw) {
                                    pw.focus();
                                    pw.value = pwd;
                                    pw.dispatchEvent(new Event('input', {bubbles: true}));
                                    pw.dispatchEvent(new Event('change', {bubbles: true}));
                                }
                            }""",
                            password,
                        )
                        await page.wait_for_timeout(500)
                        # Sign in / Continue 버튼 클릭
                        await page.evaluate(
                            """() => {
                                const btns = document.querySelectorAll('button, input[type="submit"]');
                                for (const b of btns) {
                                    const txt = (b.textContent || b.value || '').trim().toLowerCase();
                                    if (txt.includes('sign in') || txt.includes('log in') || txt.includes('continue') || txt.includes('submit')
                                        || txt === '로그인' || txt === '다음') {
                                        b.click();
                                        return;
                                    }
                                }
                                // fallback: 첫 번째 submit 버튼
                                const sub = document.querySelector('input[type="submit"], button[type="submit"]');
                                if (sub) sub.click();
                            }"""
                        )
                        status("비밀번호 입력 완료, MFA 방법 선택 대기 (4초)...")
                        await page.wait_for_timeout(4000)
                except Exception as e:
                    status(f"비밀번호 입력 실패: {e}")

            # ── 2FA 및 로그인 완료 대기 루프 ──
            auth_code_used = False
            mfa_method_selected = False
            authenticator_switched = False
            _debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_wings_dl")
            os.makedirs(_debug_dir, exist_ok=True)
            _last_screenshot_text = ""

            async def _dump_screenshot(tag: str):
                """현재 페이지 상태를 스크린샷 + 본문 텍스트로 저장 (디버그용)."""
                try:
                    _png = os.path.join(_debug_dir, f"_login_{tag}.png")
                    await page.screenshot(path=_png, full_page=True)
                    _txt_path = os.path.join(_debug_dir, f"_login_{tag}.txt")
                    body = await page.evaluate("() => document.body ? document.body.innerText : ''")
                    with open(_txt_path, "w", encoding="utf-8") as _fh:
                        _fh.write((body or "")[:5000])
                except Exception:
                    pass

            for _i in range(360):  # 최대 3분 대기
                # 이미 WINGS 메인 페이지에 도달했는지 확인
                try:
                    if await page.locator("text=Extended search").count() > 0:
                        break
                except Exception:
                    pass

                # 현재 화면이 바뀌면 스크린샷 저장 (디버그용)
                try:
                    _cur_text = (await page.evaluate("() => document.body ? document.body.innerText.slice(0, 200) : ''")) or ""
                    if _cur_text and _cur_text != _last_screenshot_text:
                        _last_screenshot_text = _cur_text
                        await _dump_screenshot(f"step{_i:03d}")
                except Exception:
                    pass

                # ── Authenticator 화면 감지 → "다른 방법" 클릭하여 이메일로 전환 ──
                if not authenticator_switched and not mfa_method_selected:
                    try:
                        is_authenticator_screen = await page.evaluate(
                            """() => {
                                const text = (document.body && document.body.innerText) || '';
                                const hasAuth = text.includes('Authenticator')
                                    || text.includes('Microsoft Authenticator')
                                    || text.includes('authenticator app')
                                    || text.includes('인증자 앱')
                                    || text.includes('인증 앱')
                                    || text.includes('Approve sign in')
                                    || text.includes('Approve the request')
                                    || text.includes('Open your Authenticator')
                                    || text.includes('Enter the code displayed');
                                // 이미 이메일 인증 화면/MFA 선택 화면인 경우는 제외
                                const isEmailScreen = text.includes('Send verification code')
                                    || text.includes('Send new verification code')
                                    || text.includes('인증 코드 보내기')
                                    || text.includes('Multi Factor Authentication email Verification');
                                const isSelectionScreen = text.includes('Multi Factor Authentication Method Selection')
                                    || text.includes('다중 요소 인증 방법 선택');
                                return hasAuth && !isEmailScreen && !isSelectionScreen;
                            }"""
                        )
                        if is_authenticator_screen:
                            await _dump_screenshot("authenticator_detected")
                            switched = await page.evaluate(
                                """() => {
                                    const wanted = [
                                        'use a different method', 'try another way',
                                        'sign in another way', 'use another verification',
                                        "i can't use my microsoft authenticator",
                                        "can't access", 'other ways to sign in',
                                        'choose another way', 'verify your identity another way',
                                        '다른 방법으로 로그인', '다른 방법 시도',
                                        '다른 인증 방법', '다른 방법으로 인증',
                                        '다른 방법 선택', '다른 인증 사용'
                                    ];
                                    const els = document.querySelectorAll('a, button, span, div, input[type="button"], input[type="submit"]');
                                    for (const el of els) {
                                        if (el.offsetParent === null) continue;
                                        const txt = ((el.textContent || el.value || '') + '').trim().toLowerCase();
                                        if (!txt || txt.length > 80) continue;
                                        for (const w of wanted) {
                                            if (txt === w || txt.includes(w)) {
                                                el.click();
                                                return el.textContent || el.value || w;
                                            }
                                        }
                                    }
                                    return null;
                                }"""
                            )
                            if switched:
                                status(f"Authenticator 화면 감지 → '{switched}' 클릭, 이메일 인증으로 전환 시도 (3초 대기)...")
                                authenticator_switched = True
                                await page.wait_for_timeout(3000)
                                continue
                            else:
                                status("Authenticator 화면 감지 — '다른 방법' 링크를 찾지 못함. 스크린샷 _login_authenticator_detected.png 확인 요망.")
                    except Exception as e:
                        status(f"Authenticator 화면 처리 중 오류: {e}")

                # ── MFA 방법 선택 화면 (Authenticator / Email) ──
                if not mfa_method_selected:
                    try:
                        has_mfa_selection = await page.evaluate(
                            """() => {
                                const text = document.body.innerText || '';
                                return text.includes('Multi Factor Authentication Method Selection')
                                    || text.includes('다중 요소 인증 방법 선택')
                                    || text.includes('MFA 방법');
                            }"""
                        )
                        if has_mfa_selection:
                            # "이메일" 또는 "Email" 라디오 클릭
                            _email_radio = page.get_by_text("이메일").or_(page.get_by_text("Email", exact=True))
                            await _email_radio.first.click()
                            await page.wait_for_timeout(500)
                            # "계속" 또는 "Continue" 버튼 클릭
                            _continue_btn = page.get_by_role("button", name="계속").or_(page.get_by_role("button", name="Continue"))
                            await _continue_btn.first.click()
                            status("MFA 방법 선택: 이메일 → 계속 클릭, 인증 화면 대기 (3초)...")
                            mfa_method_selected = True
                            await page.wait_for_timeout(3000)
                            continue
                    except Exception:
                        pass

                # ── Email MFA: "Send new verification code" 클릭 → Outlook에서 코드 읽기 ──
                if not auth_code_used:
                    try:
                        has_email_mfa = await page.evaluate(
                            """() => {
                                const text = document.body.innerText || '';
                                return text.includes('Multi Factor Authentication email Verification')
                                    || text.includes('Send verification code')
                                    || text.includes('Send new verification code')
                                    || text.includes('다중 요소 인증 이메일 인증')
                                    || text.includes('인증 코드 보내기');
                            }"""
                        )
                        if has_email_mfa:
                            # ── 1단계: Send 버튼 클릭 전에 먼저 Outlook을 열어서 현재 최신 코드를 기록 ──
                            status("Outlook을 먼저 열어 현재 최신 MFA 코드 기록 중...")
                            outlook_page = await ctx.new_page()
                            for _ol_retry in range(3):
                                await outlook_page.goto("https://outlook.office.com/mail/")
                                await outlook_page.wait_for_load_state("domcontentloaded")
                                await outlook_page.wait_for_timeout(3000)
                                _page_text = await outlook_page.evaluate("() => document.body.innerText || ''")
                                if 'Bad Gateway' in _page_text or '502' in (await outlook_page.title() or '') or '503' in _page_text:
                                    status(f"Outlook 접속 오류 (502/503). 재시도 {_ol_retry+2}/3...")
                                    await outlook_page.wait_for_timeout(5000)
                                    await outlook_page.reload()
                                    continue
                                break

                            # Outlook 로그인 처리 (Microsoft → Hyosung 리다이렉트 플로우)
                            email_addr, email_pw = _load_credentials()
                            status(f"Outlook URL: {outlook_page.url}")

                            # 사용자가 휴대폰으로 MFA 승인할 시간을 위해 최대 5분(300회) 대기.
                                # outlook.office/mail 도달하면 break로 즉시 빠져나감.
                            for _login_step in range(300):
                                await outlook_page.wait_for_timeout(1000)
                                current_url = outlook_page.url
                                status(f"Outlook 로그인 단계 {_login_step+1}: {current_url[:80]}")

                                # 각 단계별 스크린샷 + 본문 텍스트 저장 (디버그용)
                                try:
                                    _ol_png = os.path.join(_debug_dir, f"_outlook_step{_login_step:02d}.png")
                                    await outlook_page.screenshot(path=_ol_png, full_page=True)
                                    _ol_txt = os.path.join(_debug_dir, f"_outlook_step{_login_step:02d}.txt")
                                    _ol_body = await outlook_page.evaluate("() => document.body ? document.body.innerText : ''")
                                    with open(_ol_txt, "w", encoding="utf-8") as _fh:
                                        _fh.write((_ol_body or "")[:5000])
                                except Exception:
                                    pass

                                # 이미 Outlook 메일함에 도달했으면 종료
                                # (Microsoft가 outlook.office.com → outlook.cloud.microsoft로 리다이렉트하는 경우 모두 커버)
                                _is_inbox_url = (
                                    'outlook.office' in current_url
                                    or 'outlook.cloud.microsoft' in current_url
                                    or 'outlook.live.com' in current_url
                                ) and '/mail' in current_url
                                if _is_inbox_url:
                                    status("Outlook 로그인 완료!")
                                    break
                                # URL이 mail이 아니어도 받은편지함 UI가 보이면 로그인 완료로 간주
                                try:
                                    _inbox_loaded = await outlook_page.evaluate(
                                        """() => {
                                            const text = (document.body && document.body.innerText) || '';
                                            return text.includes('받은 편지함')
                                                || text.includes('Inbox')
                                                || text.includes('받은편지함');
                                        }"""
                                    )
                                    if _inbox_loaded:
                                        status(f"Outlook 받은편지함 감지 (URL: {current_url[:60]}) → 로그인 완료")
                                        break
                                except Exception:
                                    pass

                                # Microsoft 로그인 페이지
                                if 'login.microsoftonline.com' in current_url or 'login.microsoft.com' in current_url:
                                    # 계정 선택 화면 (prompt=select_account): 계정 클릭
                                    try:
                                        account_tile = outlook_page.locator(f'[data-test-id*="{email_addr}"], div[data-test-id*="yongbin"]')
                                        if await account_tile.count() > 0 and await account_tile.first.is_visible():
                                            status("Outlook: 계정 선택 화면 → 계정 클릭")
                                            await account_tile.first.click()
                                            await outlook_page.wait_for_timeout(3000)
                                            continue
                                    except Exception:
                                        pass

                                    # "다른 계정 사용" / "Use another account" 처리
                                    try:
                                        other_acct = outlook_page.locator('#otherTile, [data-test-id="otherTile"]')
                                        if await other_acct.count() > 0 and await other_acct.is_visible():
                                            status("Outlook: '다른 계정 사용' 클릭")
                                            await other_acct.click()
                                            await outlook_page.wait_for_timeout(2000)
                                            continue
                                    except Exception:
                                        pass

                                    # "Use password instead" / "비밀번호 사용" 처리
                                    # (Microsoft가 passwordless/authenticator 승인을 요구할 때 표시되는 fallback 링크)
                                    try:
                                        use_pw_clicked = await outlook_page.evaluate(
                                            """() => {
                                                // 1) Microsoft Azure AD가 사용하는 알려진 ID 우선 시도
                                                const knownIds = [
                                                    'idA_PWD_SwitchToPassword',
                                                    'signInAnotherWay',
                                                    'idA_PWD_SwitchToCredPicker'
                                                ];
                                                for (const id of knownIds) {
                                                    const el = document.getElementById(id);
                                                    if (el && el.offsetParent !== null) {
                                                        el.click();
                                                        return 'id:' + id;
                                                    }
                                                }

                                                // 2) 실제 클릭 가능한 요소(a, button, role=button/link)만 대상
                                                const wanted = [
                                                    'use password instead',
                                                    'use your password instead',
                                                    'use your password',
                                                    'sign in with password',
                                                    'sign in another way',
                                                    '비밀번호 사용', '비밀번호 입력',
                                                    '대신 비밀번호 사용', '비밀번호로 로그인'
                                                ];
                                                const els = document.querySelectorAll(
                                                    'a, button, [role="button"], [role="link"], input[type="button"], input[type="submit"]'
                                                );
                                                for (const el of els) {
                                                    if (el.offsetParent === null) continue;
                                                    // 공백/개행을 단일 스페이스로 정규화 후 lowercase
                                                    const raw = (el.innerText || el.textContent || el.value || '');
                                                    const txt = raw.replace(/\\s+/g, ' ').trim().toLowerCase();
                                                    if (!txt || txt.length > 60) continue;
                                                    // 정확 일치만 (substring 매칭은 컨테이너를 잡아서 클릭이 무효)
                                                    if (wanted.includes(txt)) {
                                                        el.click();
                                                        return (el.tagName || '') + ':' + txt.substring(0, 50);
                                                    }
                                                }
                                                return null;
                                            }"""
                                        )
                                        if use_pw_clicked:
                                            status(f"Outlook: 'Use password instead' 매칭({use_pw_clicked}) → 클릭, 비밀번호 화면 대기 (3초)...")
                                            await outlook_page.wait_for_timeout(3000)
                                            continue
                                    except Exception:
                                        pass

                                    # 1단계: 비밀번호 입력 화면 (이메일보다 먼저 체크!)
                                    _pw_filled = await outlook_page.evaluate(
                                        """(pwd) => {
                                            const pw = document.querySelector('input[type="password"]:not([style*="display:none"]):not([hidden])');
                                            if (pw && pw.offsetParent !== null) {
                                                pw.focus();
                                                pw.value = pwd;
                                                pw.dispatchEvent(new Event('input', {bubbles: true}));
                                                pw.dispatchEvent(new Event('change', {bubbles: true}));
                                                return true;
                                            }
                                            return false;
                                        }""", email_pw or ""
                                    )
                                    if _pw_filled:
                                        status("Outlook: 비밀번호 입력 완료, 로그인 클릭...")
                                        await outlook_page.wait_for_timeout(500)
                                        await outlook_page.evaluate(
                                            """() => {
                                                const btns = document.querySelectorAll('button, input[type="submit"]');
                                                for (const b of btns) {
                                                    const txt = (b.textContent || b.value || '').trim().toLowerCase();
                                                    if (txt.includes('sign in') || txt.includes('log in')
                                                        || txt === '로그인' || txt === '다음') {
                                                        b.click();
                                                        return;
                                                    }
                                                }
                                                const sub = document.querySelector('input[type="submit"], button[type="submit"]');
                                                if (sub) sub.click();
                                            }"""
                                        )
                                        await outlook_page.wait_for_timeout(3000)
                                        continue

                                    # 2단계: 이메일 입력 화면
                                    _email_filled = await outlook_page.evaluate(
                                        """(addr) => {
                                            const el = document.querySelector('input[type="email"], input[name="loginfmt"]');
                                            if (el && el.offsetParent !== null) {
                                                el.focus();
                                                el.value = addr;
                                                el.dispatchEvent(new Event('input', {bubbles: true}));
                                                el.dispatchEvent(new Event('change', {bubbles: true}));
                                                return true;
                                            }
                                            return false;
                                        }""", email_addr or ""
                                    )
                                    if _email_filled:
                                        status("Outlook: 이메일 입력 완료, 다음 클릭...")
                                        await outlook_page.wait_for_timeout(300)
                                        await outlook_page.evaluate(
                                            """() => {
                                                const sub = document.querySelector('input[type="submit"]');
                                                if (sub) { sub.click(); return; }
                                                const btns = document.querySelectorAll('button');
                                                for (const b of btns) {
                                                    const txt = (b.textContent || '').trim().toLowerCase();
                                                    if (txt === 'next' || txt === '다음') { b.click(); return; }
                                                }
                                            }"""
                                        )
                                        try:
                                            await outlook_page.wait_for_selector(
                                                'input[type="password"]', timeout=10000
                                            )
                                        except Exception:
                                            await outlook_page.wait_for_timeout(2000)
                                        continue

                                    # 3단계: "Stay signed in?" / "로그인 상태 유지?" 화면
                                    yes_btn = outlook_page.locator('input[value="Yes"], input[value="예"], button:has-text("Yes"), button:has-text("예")')
                                    if await yes_btn.count() > 0 and await yes_btn.first.is_visible():
                                        status("Outlook: 로그인 상태 유지 → 예")
                                        await yes_btn.first.click()
                                        await outlook_page.wait_for_timeout(3000)
                                        continue

                                # 알 수 없는 페이지 → 대기
                                await outlook_page.wait_for_timeout(2000)

                            await outlook_page.wait_for_timeout(3000)
                            status("Outlook Web 로드 완료. 인증 이메일 검색 중...")

                            # 현재 최신(첫 번째) MFA 이메일의 코드를 기록
                            _get_first_mfa_code_js = """() => {
                                // 방법1: aria-label에서 첫 번째 Daimler MFA 이메일 찾기
                                const items = document.querySelectorAll('[aria-label*="Daimler"], [aria-label*="MFA"], [aria-label*="인증"]');
                                for (const el of items) {
                                    const text = el.getAttribute('aria-label') || el.textContent || '';
                                    const m = text.match(/(\\d{6})/);
                                    if (m) return m[1];
                                }
                                // 방법2: innerText에서 첫 번째 매칭 (최신순 정렬이므로 첫 번째 = 최신)
                                const allText = document.body.innerText || '';
                                let m = allText.match(/(\\d{6})\\s*-\\s*(?:Your )?Daimler Truck/);
                                if (m) return m[1];
                                m = allText.match(/(\\d{6})\\s*[-–]\\s*[^\\n]*(?:MFA|인증|verification)/i);
                                if (m) return m[1];
                                return null;
                            }"""

                            old_first_code = None
                            try:
                                old_first_code = await outlook_page.evaluate(_get_first_mfa_code_js)
                                status(f"현재 최신 MFA 코드: {old_first_code}")
                            except Exception:
                                pass

                            # ── 2단계: 이제 WINGS 페이지로 돌아가서 이메일 입력 → Send verification code 클릭 ──
                            status("WINGS 페이지에서 인증 코드 요청 중...")
                            await page.bring_to_front()
                            await page.wait_for_timeout(1000)

                            # 2-a) 이메일 필드를 실제 주소로 채우기 (마스킹된 placeholder만으로는 검증 실패함)
                            email_fill_result = await page.evaluate(
                                """(email) => {
                                    const inputs = document.querySelectorAll('input');
                                    for (const inp of inputs) {
                                        if (inp.offsetParent === null) continue;
                                        const type = (inp.type || '').toLowerCase();
                                        const ph = (inp.placeholder || '').toLowerCase();
                                        const id = (inp.id || '').toLowerCase();
                                        const name = (inp.name || '').toLowerCase();
                                        const isEmail = type === 'email'
                                            || ph.includes('email') || ph.includes('이메일')
                                            || id.includes('email') || name.includes('email');
                                        if (!isEmail) continue;
                                        // value 셋팅 + React/Angular용 네이티브 setter 사용
                                        const proto = Object.getPrototypeOf(inp);
                                        const setter = Object.getOwnPropertyDescriptor(proto, 'value') &&
                                                       Object.getOwnPropertyDescriptor(proto, 'value').set;
                                        inp.focus();
                                        if (setter) setter.call(inp, email); else inp.value = email;
                                        inp.dispatchEvent(new Event('input', {bubbles: true}));
                                        inp.dispatchEvent(new Event('change', {bubbles: true}));
                                        inp.dispatchEvent(new Event('blur', {bubbles: true}));
                                        return {ok: true, id: inp.id || '', name: inp.name || ''};
                                    }
                                    return {ok: false};
                                }""",
                                email_addr,
                            )
                            status(f"이메일 입력 결과: {email_fill_result}")
                            await page.wait_for_timeout(800)

                            # 2-b) Send verification code 클릭
                            clicked = await page.evaluate(
                                """() => {
                                    const btns = document.querySelectorAll('button');
                                    for (const b of btns) {
                                        const txt = (b.textContent || '').trim();
                                        if (txt === 'Send verification code' || txt === 'Send new verification code'
                                            || txt === '인증 코드 보내기' || txt === '새 인증 코드 보내기') {
                                            b.removeAttribute('aria-hidden');
                                            b.style.display = '';
                                            b.click();
                                            return txt;
                                        }
                                    }
                                    return null;
                                }"""
                            )
                            status(f"Send 버튼 클릭: {clicked} → Outlook에서 새 메일 대기 중...")
                            await page.wait_for_timeout(2000)

                            # 2-c) 검증 에러 메시지가 떴으면 재시도 (이메일 다시 입력 + Send 다시 클릭)
                            try:
                                err_present = await page.evaluate(
                                    """() => {
                                        const text = document.body.innerText || '';
                                        return text.includes('trouble verifying your email')
                                            || text.includes('valid email address')
                                            || text.includes('이메일 주소를 확인');
                                    }"""
                                )
                                if err_present:
                                    status("이메일 검증 에러 감지 → 이메일 재입력 후 Send 재시도...")
                                    await page.evaluate(
                                        """(email) => {
                                            const inputs = document.querySelectorAll('input');
                                            for (const inp of inputs) {
                                                if (inp.offsetParent === null) continue;
                                                const type = (inp.type || '').toLowerCase();
                                                const ph = (inp.placeholder || '').toLowerCase();
                                                const isEmail = type === 'email'
                                                    || ph.includes('email') || ph.includes('이메일');
                                                if (!isEmail) continue;
                                                const proto = Object.getPrototypeOf(inp);
                                                const setter = Object.getOwnPropertyDescriptor(proto, 'value') &&
                                                               Object.getOwnPropertyDescriptor(proto, 'value').set;
                                                inp.focus();
                                                if (setter) setter.call(inp, ''); else inp.value = '';
                                                inp.dispatchEvent(new Event('input', {bubbles: true}));
                                                if (setter) setter.call(inp, email); else inp.value = email;
                                                inp.dispatchEvent(new Event('input', {bubbles: true}));
                                                inp.dispatchEvent(new Event('change', {bubbles: true}));
                                                inp.dispatchEvent(new Event('blur', {bubbles: true}));
                                                return true;
                                            }
                                        }""",
                                        email_addr,
                                    )
                                    await page.wait_for_timeout(800)
                                    await page.evaluate(
                                        """() => {
                                            const btns = document.querySelectorAll('button');
                                            for (const b of btns) {
                                                const txt = (b.textContent || '').trim();
                                                if (txt === 'Send verification code' || txt === 'Send new verification code'
                                                    || txt === '인증 코드 보내기' || txt === '새 인증 코드 보내기') {
                                                    b.click();
                                                    return;
                                                }
                                            }
                                        }"""
                                    )
                                    await page.wait_for_timeout(2000)
                            except Exception:
                                pass

                            # Outlook 탭으로 전환하여 새 코드 감지
                            await outlook_page.bring_to_front()
                            await outlook_page.wait_for_timeout(2000)

                            # 새 MFA 이메일 도착 감지 (최대 3분 대기)
                            email_code = None
                            for attempt in range(36):  # 36 * 5초 = 3분
                                try:
                                    current_first = await outlook_page.evaluate(_get_first_mfa_code_js)
                                    if current_first and current_first != old_first_code:
                                        # 최신 이메일이 바뀜 → 새 인증 코드!
                                        email_code = current_first
                                        status(f"새 인증 코드 감지! {email_code} (이전: {old_first_code})")
                                        break
                                    elif current_first:
                                        status(f"새 메일 대기 중... 현재 최신: {current_first} (attempt {attempt+1}/36)")
                                    else:
                                        status(f"MFA 이메일 대기 중... (attempt {attempt+1}/36)")
                                except Exception as e:
                                    status(f"이메일 검색 오류: {e}")

                                # 주기적으로 새로고침
                                if attempt % 3 == 2:
                                    status(f"Outlook 새로고침 중... (attempt {attempt+1})")
                                    await outlook_page.reload()
                                    await outlook_page.wait_for_timeout(5000)
                                else:
                                    await outlook_page.wait_for_timeout(5000)

                            # Outlook 탭 닫기
                            await outlook_page.close()

                            if email_code and len(email_code) == 6:
                                # WINGS 페이지로 돌아가서 코드 입력
                                status(f"인증 코드 {email_code} 입력 중...")
                                await page.evaluate(
                                    """(code) => {
                                        // 인증 코드 전용 입력 필드 찾기 (이메일 필드 제외)
                                        const allInputs = document.querySelectorAll('input');
                                        for (const inp of allInputs) {
                                            if (inp.offsetParent === null) continue;  // 숨겨진 필드 제외
                                            const ph = (inp.placeholder || '').toLowerCase();
                                            const id = (inp.id || '').toLowerCase();
                                            const name = (inp.name || '').toLowerCase();
                                            const type = (inp.type || '').toLowerCase();
                                            // 이메일 필드 제외
                                            if (type === 'email') continue;
                                            if (ph.includes('email') || ph.includes('이메일')) continue;
                                            if (id.includes('email') || name.includes('email')) continue;
                                            // 인증 코드 필드 매칭: placeholder/id/name에 verification, code, 인증, 코드 등
                                            if (ph.includes('verif') || ph.includes('code') || ph.includes('인증') || ph.includes('코드')
                                                || id.includes('verif') || id.includes('code') || id.includes('otp')
                                                || name.includes('verif') || name.includes('code') || name.includes('otp')) {
                                                inp.focus();
                                                inp.value = code;
                                                inp.dispatchEvent(new Event('input', {bubbles: true}));
                                                inp.dispatchEvent(new Event('change', {bubbles: true}));
                                                return 'found_by_attr';
                                            }
                                        }
                                        // 속성으로 못 찾으면: 이메일 필드가 아닌 마지막 visible text input 사용
                                        const textInputs = [];
                                        for (const inp of allInputs) {
                                            if (inp.offsetParent === null) continue;
                                            const type = (inp.type || '').toLowerCase();
                                            if (type === 'email' || type === 'password' || type === 'hidden' || type === 'submit') continue;
                                            const ph = (inp.placeholder || '').toLowerCase();
                                            if (ph.includes('email') || ph.includes('이메일')) continue;
                                            textInputs.push(inp);
                                        }
                                        // 마지막 text input이 인증코드 필드일 가능성 높음
                                        if (textInputs.length > 1) {
                                            const inp = textInputs[textInputs.length - 1];
                                            inp.focus();
                                            inp.value = code;
                                            inp.dispatchEvent(new Event('input', {bubbles: true}));
                                            inp.dispatchEvent(new Event('change', {bubbles: true}));
                                            return 'found_last_input';
                                        }
                                        return 'not_found';
                                    }""", email_code
                                )
                                await page.wait_for_timeout(500)
                                # "Verify code" / "코드 확인" 버튼 클릭
                                await page.evaluate(
                                    """() => {
                                        const btns = document.querySelectorAll('button, input[type="submit"]');
                                        for (const b of btns) {
                                            const txt = (b.textContent || b.value || '').trim();
                                            if (txt.includes('Verify') || txt.includes('확인')
                                                || txt.includes('verify') || txt.includes('코드 확인')) {
                                                b.click();
                                                return;
                                            }
                                        }
                                    }"""
                                )
                                status("이메일 인증 코드 제출 완료. 로그인 진행 중...")
                                auth_code_used = True
                                await page.wait_for_timeout(5000)
                                continue
                            else:
                                status("이메일에서 인증 코드를 찾지 못했습니다.")
                    except Exception as e:
                        status(f"이메일 MFA 처리 중 오류: {e}")
                        pass

                # "로그인 상태 유지" 화면 처리
                try:
                    stay_signed = page.locator("input#idSIButton9, input[value='Yes'], button:has-text('Yes')")
                    if await stay_signed.count() > 0:
                        await stay_signed.first.click()
                        await page.wait_for_timeout(2000)
                        continue
                except Exception:
                    pass

                await page.wait_for_timeout(500)

            try:
                await page.wait_for_selector("text=Extended search", timeout=60000)
            except Exception:
                await _dump_screenshot("FINAL_TIMEOUT")
                raise RuntimeError(
                    "로그인 완료 대기 시간 초과. _wings_dl/_login_*.png 스크린샷에서 막힌 화면을 확인하세요."
                )
            status("로그인 완료")
            await page.bring_to_front()

        # ── 2. Extended Search 진입 ────────────────────────────────────────────
        status("Extended Search 클릭 중...")
        await page.wait_for_load_state("networkidle", timeout=30000)
        await page.wait_for_timeout(1000)
        await page.click("text=Extended search", timeout=60000)
        await page.wait_for_load_state("networkidle", timeout=15000)

        # ── 3. 기존 필터 조건 제거 ────────────────────────────────────────────
        try:
            remove_btn = page.locator("text=Remove all filter criteria")
            if await remove_btn.is_visible(timeout=2000):
                await remove_btn.click()
                await page.wait_for_timeout(600)
        except Exception:
            pass

        # ── 4. 필터 조건 설정 ─────────────────────────────────────────────────
        status("필터 조건 설정 중...")
        # 'Remove all' 후 WINGS가 빈 행 1개를 재생성할 시간을 보장
        _rowcount = await _wait_filter_rows(page, min_count=1, timeout_ms=15000)
        if _rowcount < 1:
            raise RuntimeError("Extended Search 필터 행이 초기화되지 않았습니다 (15초 초과).")

        if single:
            # 단일 월: equal = YYYY-MM-01
            await _set_filter_row(page, 0, "Requested delivery date", "equal", start_date)

        elif consecutive:
            # 연속 월 (예: 04,05,06): greater equal start AND less equal end
            await _set_filter_row(page, 0, "Requested delivery date", "greater equal", start_date)
            await _copy_filter_row(page, 0)
            await _wait_filter_rows(page, min_count=2, timeout_ms=10000)
            await _set_filter_row(page, 1, "Requested delivery date", "less equal", end_date)

        else:
            # 비연속 월 (예: 04,06): 각 월마다 equal 행 추가
            for i, month in enumerate(months_sorted):
                date_str = month + "-01"
                if i > 0:
                    await _copy_filter_row(page, i - 1)
                    await _wait_filter_rows(page, min_count=i + 1, timeout_ms=10000)
                await _set_filter_row(page, i, "Requested delivery date", "equal", date_str)
            # 행 사이 커넥터를 and → or 로 변경 (비연속 월은 OR 조건)
            await _set_all_row_connectors(page, "or")

        # ── 5. Execute 클릭 → 결과 페이지 대기 ───────────────────────────────
        status("검색 실행 중...")
        await page.click("text=Execute")

        # 3초 후 입력 오류 팝업 확인
        await asyncio.sleep(3)
        try:
            popup = page.locator("text=The requested action could not be completed")
            if await popup.is_visible(timeout=500):
                debug_info = ""
                try:
                    with open("wings_debug.log", encoding="utf-8") as _f:
                        debug_info = _f.read()
                except Exception:
                    pass
                raise RuntimeError(
                    "WINGS 입력 오류(U0033): 필터 조건이 올바르게 설정되지 않았습니다.\n\n"
                    f"디버그 로그:\n{debug_info}"
                )
        except RuntimeError:
            raise
        except Exception:
            pass

        # 결과 페이지의 Export 버튼이 나타날 때까지 대기 (최대 60초)
        status("결과 로드 대기 중...")
        try:
            await page.wait_for_selector("text=Export", timeout=60000)
        except Exception:
            raise RuntimeError("검색 결과 페이지가 로드되지 않았습니다 (60초 초과).")

        # Export 전 4초 추가 대기 (결과 완전 로드)
        await page.wait_for_timeout(4000)

        # ── 6. Export 클릭 → 다운로드 대기 ───────────────────────────────────
        status("Export 클릭 중... 파일 다운로드를 기다리는 중입니다.")

        download_holder = []
        download_event = asyncio.Event()

        def _on_download(dl):
            download_holder.append(dl)
            download_event.set()

        # 현재 페이지 + 팝업으로 열리는 새 페이지에도 download 리스너 등록
        page.on("download", _on_download)

        def _on_new_page(new_page):
            new_page.on("download", _on_download)

        ctx.on("page", _on_new_page)

        # 파일시스템 감시: Export 클릭 전 기존 파일 스냅샷
        # WINGS Export는 실제로 .csv(예: Report_DEF_*.csv)로 떨어지므로 csv도 포함.
        user_dl_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        _snap_dirs = [download_dir, user_dl_dir]
        _watch_exts = ("xlsx", "xls", "csv")
        _before = set()
        for _d in _snap_dirs:
            for _ext in _watch_exts:
                _before.update(glob.glob(os.path.join(_d, f"*.{_ext}")))

        await page.click("text=Export")

        # 방어선 1: Playwright download 이벤트 (30초 대기)
        fpath = None
        try:
            await asyncio.wait_for(download_event.wait(), timeout=30)
            dl = download_holder[0]
            fname = dl.suggested_filename or f"wings_{start_date}_to_{end_date}.xlsx"
            fpath = os.path.join(download_dir, fname)
            await dl.save_as(fpath)
            status(f"다운로드 완료: {fname}")
        except asyncio.TimeoutError:
            # 방어선 2: 파일시스템에서 새 파일 탐지 (최대 60초)
            status("다운로드 파일 감지 중...")
            for _ in range(60):
                _after = set()
                for _d in _snap_dirs:
                    for _ext in _watch_exts:
                        _after.update(glob.glob(os.path.join(_d, f"*.{_ext}")))
                _new = {f for f in (_after - _before) if not f.endswith(".crdownload")}
                if _new:
                    src = max(_new, key=os.path.getmtime)
                    fname = os.path.basename(src)
                    # 사용자 Downloads에 떨어졌다면 작업용 download_dir로 복사
                    if os.path.dirname(src) != os.path.abspath(download_dir):
                        import shutil as _shutil
                        fpath = os.path.join(download_dir, fname)
                        _shutil.copy2(src, fpath)
                    else:
                        fpath = src
                    status(f"다운로드 완료: {fname}")
                    break
                await asyncio.sleep(1)
            if not fpath:
                raise RuntimeError("다운로드 시간 초과 (90초). Export 후 파일이 생성되지 않았습니다.")

        try:
            await ctx.close()
        except Exception:
            pass
        _release_profile_lock()
        return fpath


async def _set_filter_row(page, row_idx: int, field: str, operator: str, value: str):
    """WINGS Extended Search 필터 행 설정.

    사용자 동작 그대로 재현:
    1) 필드 입력창에 키워드 타이핑 → 2초 대기 → 첫 번째 팝업 항목 클릭
    2) 오퍼레이터 입력창 클릭 → 첫 번째 팝업 항목 클릭
    3) 나타나는 날짜 입력창에 직접 타이핑
    """
    log = []

    # ── 0. FilterCriteriaWidget에서 위젯 ID 수집 ──────────────────────────────
    info = await page.evaluate(
        """idx => {
            const rows = dijit.registry.toArray().filter(w =>
                w.declaredClass === 'com.daimler.wings.view.grid.filter.FilterCriteriaWidget'
            );
            if (!rows[idx]) return null;
            const children = dijit.registry.findWidgets(rows[idx].domNode);
            const fieldW = children.find(c => c.declaredClass.includes('DatafieldDataFilteringSelect'));
            const opW    = children.find(c =>
                c.declaredClass.includes('FilteringSelect') && !c.declaredClass.includes('Datafield')
            );
            // 오퍼레이터의 ▼ 화살표 버튼 노드
            let opBtn = null;
            if (opW) {
                // 1) 위젯 프로퍼티로 찾기
                opBtn = opW._buttonNode || opW.downArrowNode || opW._arrowNode || null;
                // 2) 위젯 domNode 안에서 CSS 클래스로 찾기
                if (!opBtn && opW.domNode) {
                    opBtn = opW.domNode.querySelector(
                        '.dijitArrowButton, .dijitDownArrowButton, [class*="ArrowButton"], [class*="arrowButton"]'
                    );
                }
            }
            return {
                fieldId:    fieldW ? fieldW.id : null,
                fieldNode:  fieldW ? (fieldW.focusNode ? fieldW.focusNode.id : null) : null,
                opId:       opW    ? opW.id    : null,
                opNode:     opW    ? (opW.focusNode ? opW.focusNode.id : null) : null,
                opArrow:    opBtn  ? opBtn.id  : null,
                opArrowClass: opBtn ? (opBtn.className || '') : null,
            };
        }""",
        row_idx,
    )
    log.append(f"info={info}")
    if not info or not info.get("fieldId"):
        log.append("ERROR: widget not found")
        _write_debug(row_idx, log)
        return

    field_id   = info["fieldId"]
    field_node = info.get("fieldNode")  # focusNode ID (input element)
    op_id      = info.get("opId")
    op_node    = info.get("opNode")
    op_arrow   = info.get("opArrow")   # ▼ 화살표 버튼 ID

    # ── 1. 필드 입력창에 "Requested" 타이핑 → 팝업 대기 → 첫 번째 항목 클릭 ──
    keyword = field.split()[0]  # e.g. "Requested"

    if field_node:
        # Playwright 자체 클릭으로 포커스 확보 후 타이핑
        try:
            await page.locator(f"#{field_node}").click()
            await page.keyboard.press("Control+a")
            await page.keyboard.press("Delete")
            await page.wait_for_timeout(200)
            await page.keyboard.type(keyword, delay=80)
            log.append(f"field typed via locator: {keyword}")
        except Exception as e:
            log.append(f"field locator error: {e}")
            # fallback: JS focus + keyboard
            await page.evaluate(
                """id => {
                    const w = dijit.byId(id);
                    if (w && w.focusNode) { w.focusNode.focus(); w.focusNode.click(); }
                }""",
                field_id,
            )
            await page.keyboard.type(keyword, delay=80)
            log.append(f"field typed via JS fallback: {keyword}")
    else:
        # focusNode ID 없음 → JS로 위젯 직접 조작
        await page.evaluate(
            """id => {
                const w = dijit.byId(id);
                if (!w) return;
                if (w.focusNode) { w.focusNode.focus(); w.focusNode.click(); }
            }""",
            field_id,
        )
        await page.keyboard.type(keyword, delay=80)
        log.append(f"field typed via JS (no focusNode id): {keyword}")

    await page.wait_for_timeout(3000)  # 팝업이 나타날 때까지 3초 대기 (사용자 지시)

    # 팝업 첫 번째 항목 Playwright 실제 클릭
    first_item_result = await _click_first_popup_item_playwright(page)
    log.append(f"field popup: {first_item_result}")
    await page.wait_for_timeout(1000)  # 필드 선택 후 오퍼레이터 드롭다운 갱신 대기

    # ── 2. 오퍼레이터: ▼ 화살표 클릭 → 팝업 → 첫 번째 항목 클릭 ─────────────
    # 사용자 지시: 텍스트 입력창은 절대 클릭 안 함. ▼ 오른쪽 끝만 클릭.
    if op_id:
        # 오퍼레이터 위젯 전체 domNode의 bounding box를 구해서
        # 오른쪽 끝(▼ 버튼 위치)을 page.mouse.click으로 직접 클릭
        op_bbox = await page.evaluate(
            """id => {
                const w = dijit.byId(id);
                if (!w || !w.domNode) return null;
                const r = w.domNode.getBoundingClientRect();
                return {x: r.x, y: r.y, w: r.width, h: r.height,
                        scrollX: window.scrollX, scrollY: window.scrollY};
            }""",
            op_id,
        )
        log.append(f"op_bbox={op_bbox}")

        if op_bbox:
            # ▼ 는 위젯 오른쪽 끝에 있음 → 오른쪽에서 10px 안쪽, 수직 중앙
            click_x = op_bbox['x'] + op_bbox['scrollX'] + op_bbox['w'] - 10
            click_y = op_bbox['y'] + op_bbox['scrollY'] + op_bbox['h'] / 2
            await page.mouse.click(click_x, click_y)
            log.append(f"op arrow clicked via mouse ({click_x:.0f}, {click_y:.0f})")
        else:
            # bbox 실패 → JS _openDropDown 시도
            await page.evaluate(
                """id => {
                    const w = dijit.byId(id);
                    if (!w) return;
                    if (typeof w._openDropDown === 'function') w._openDropDown();
                }""",
                op_id,
            )
            log.append("op opened via JS _openDropDown (bbox fallback)")

        await page.wait_for_timeout(1200)

        op_result = await _click_popup_item_by_text_playwright(page, operator)
        log.append(f"op popup: {op_result}")
        await page.wait_for_timeout(1000)  # 날짜 입력창 나타날 때까지 대기

    # ── 3. 날짜 입력창 찾기 (FilterCriteriaWidget 자식 + 전역 검색) ──────────
    date_node_id = await page.evaluate(
        """idx => {
            // 방법 1: FilterCriteriaWidget 자식에서 검색
            const rows = dijit.registry.toArray().filter(w =>
                w.declaredClass === 'com.daimler.wings.view.grid.filter.FilterCriteriaWidget'
            );
            if (rows[idx]) {
                const children = dijit.registry.findWidgets(rows[idx].domNode);
                const dateW = children.find(c =>
                    (c.declaredClass.includes('TextBox') || c.declaredClass.includes('ValidationTextBox')) &&
                    !c.declaredClass.includes('FilteringSelect')
                );
                if (dateW && dateW.focusNode) return {id: dateW.id, nodeId: dateW.focusNode.id, src: 'widget_child'};
            }

            // 방법 2: dijit 레지스트리 전체에서 보이는 TextBox 검색
            const allWidgets = dijit.registry.toArray();
            const textBoxes = allWidgets.filter(w =>
                (w.declaredClass.includes('TextBox') || w.declaredClass.includes('ValidationTextBox')) &&
                !w.declaredClass.includes('FilteringSelect') &&
                w.domNode && w.domNode.offsetParent !== null
            );
            if (textBoxes.length > 0) {
                const w = textBoxes[0];
                return {id: w.id, nodeId: w.focusNode ? w.focusNode.id : null, src: 'global_registry', cls: w.declaredClass};
            }

            // 방법 3: DOM에서 보이는 text input 검색 (필터 영역 내)
            const filterArea = document.querySelector('.wings-filter, .filter-criteria, [class*="FilterCriteria"]');
            const target = filterArea || document.body;
            const inputs = Array.from(target.querySelectorAll('input[type="text"], input:not([type])'))
                .filter(el => el.offsetParent !== null && !el.readOnly && !el.disabled);
            // 마지막 input이 보통 새로 생긴 날짜 입력창
            if (inputs.length > 0) {
                const el = inputs[inputs.length - 1];
                return {id: null, nodeId: el.id || null, src: 'dom_input', cls: el.className};
            }

            return null;
        }""",
        row_idx,
    )
    log.append(f"date_node={date_node_id}")

    if date_node_id:
        widget_id = date_node_id.get("id")
        node_id   = date_node_id.get("nodeId")
        src       = date_node_id.get("src", "")

        if widget_id:
            # Dojo 위젯 API로 값 설정
            await page.evaluate(
                """([id, val]) => {
                    const w = dijit.byId(id);
                    if (!w) return;
                    w.set('value', val);
                    if (w.validate) w.validate(false);
                    // change 이벤트 발생시켜 WINGS가 값 인식하도록
                    if (w.focusNode) {
                        w.focusNode.dispatchEvent(new Event('change', {bubbles: true}));
                    }
                }""",
                [widget_id, value],
            )
            log.append(f"date SET via widget ({src}): {value}")
        elif node_id:
            # DOM input에 직접 타이핑
            try:
                await page.locator(f"#{node_id}").click()
                await page.keyboard.press("Control+a")
                await page.keyboard.type(value, delay=80)
                log.append(f"date TYPED via locator ({src}): {value}")
            except Exception as e:
                log.append(f"date locator error: {e}")
        else:
            log.append("date: no usable id found")
    else:
        # 최후 수단: Tab으로 날짜 필드로 이동 후 타이핑
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(300)
        await page.keyboard.type(value, delay=80)
        log.append(f"date TYPED via Tab fallback: {value}")

    _write_debug(row_idx, log)
    await page.wait_for_timeout(400)


async def _click_first_popup_item_playwright(page) -> str:
    """보이는 Dojo 팝업에서 첫 번째 항목을 Playwright 실제 클릭으로 선택한다.

    JS dispatchEvent 대신 Playwright locator().click()을 사용하여
    실제 마우스 이벤트를 발생시킨다.
    """
    # 시도 1: [item] 속성을 가진 요소 (DataGrid 스타일 팝업)
    for sel in ('[item]', '.dijitComboBoxItem', '.dijitMenuItem'):
        try:
            loc = page.locator(sel).first()
            if await loc.count() > 0:
                txt = (await loc.inner_text()).strip()[:50]
                await loc.click(timeout=3000)
                return f"playwright:{sel} '{txt}'"
        except Exception as e:
            # 이 셀렉터는 실패 → 다음 시도
            continue

    # 시도 2: ArrowDown + Enter 키보드 방식
    await page.keyboard.press("ArrowDown")
    await page.wait_for_timeout(300)
    await page.keyboard.press("Enter")
    return "keyboard:ArrowDown+Enter"


async def _click_popup_item_by_text_playwright(page, target_text: str) -> str:
    """팝업에서 target_text와 일치하는 항목을 선택한다.

    방법 1) JS getBoundingClientRect → page.mouse.click (단일 JS 호출, 빠름)
    방법 2) Playwright locator 텍스트 필터 클릭 (fallback)
    방법 3) 첫 번째 항목 fallback
    """
    target_lower = target_text.lower().strip()

    # ── 방법 1: JS 단일 호출로 위치 탐색 → page.mouse.click ─────────────────
    # 루프 없이 JS 한 번으로 처리 → 빠름
    bbox = await page.evaluate(
        """text => {
            const lower = text.toLowerCase().trim();
            const popupSels = [
                '.dijitComboBoxPopup', '.dijitPopup', '.dijitSelectMenu',
                '[role="listbox"]', '[role="list"]'
            ];
            const popups = popupSels.flatMap(s => Array.from(document.querySelectorAll(s)))
                .filter(el => {
                    const cs = window.getComputedStyle(el);
                    return cs.display !== 'none' && cs.visibility !== 'hidden'
                        && el.offsetParent !== null;
                });
            const containers = popups.length > 0 ? popups : [document.body];
            const itemSels = [
                '[item]', '[role="option"]', '.dijitComboBoxItem',
                '.dijitMenuItem', '.dijitSelectItem'
            ];
            for (const exact of [true, false]) {
                for (const container of containers) {
                    for (const isel of itemSels) {
                        for (const el of container.querySelectorAll(isel)) {
                            if (el.offsetParent === null) continue;
                            const t = el.textContent.trim().toLowerCase();
                            if (exact ? t === lower : t.includes(lower)) {
                                const r = el.getBoundingClientRect();
                                return {x: r.x + r.width/2, y: r.y + r.height/2,
                                        partial: !exact, tag: el.tagName};
                            }
                        }
                    }
                    for (const el of container.querySelectorAll('div, li, span')) {
                        if (el.offsetParent === null || el.children.length > 0) continue;
                        const t = el.textContent.trim().toLowerCase();
                        if (exact ? t === lower : t.includes(lower)) {
                            const r = el.getBoundingClientRect();
                            return {x: r.x + r.width/2, y: r.y + r.height/2,
                                    partial: !exact, tag: el.tagName, leaf: true};
                        }
                    }
                }
            }
            return null;
        }""",
        target_lower,
    )

    if bbox:
        await page.mouse.click(bbox["x"], bbox["y"])
        match_type = "partial" if bbox.get("partial") else "exact"
        return f"js+mouse:{match_type}:{target_text}({bbox.get('tag','')})"

    # ── 방법 2: Playwright 텍스트 필터 locator ──────────────────────────────
    for sel in ('[item]', '[role="option"]', '.dijitComboBoxItem',
                '.dijitMenuItem', '.dijitSelectItem'):
        try:
            loc = page.locator(sel).filter(has_text=target_text)
            if await loc.count() > 0:
                await loc.first().click(timeout=3000)
                return f"pw_filter:{sel} '{target_text}'"
        except Exception:
            continue

    # ── 방법 3: 첫 번째 항목 fallback ───────────────────────────────────────
    result = await _click_first_popup_item_playwright(page)
    return f"fallback→{result}"


async def _click_popup_item(page, text: str) -> bool:
    """열려 있는 Dojo 드롭다운 팝업에서 텍스트가 정확히 일치하는 항목을 클릭한다.
    JavaScript DOM 직접 조작 방식으로 정확한 매칭을 보장한다.
    """
    result = await page.evaluate(
        """text => {
            const lower = text.toLowerCase();
            // 모든 Dojo 팝업 순회
            const popups = document.querySelectorAll(
                '.dijitComboBoxPopup, .dijitPopup, .dijitSelectMenu'
            );
            for (const popup of popups) {
                // 숨겨진 팝업 제외
                const style = window.getComputedStyle(popup);
                if (style.display === 'none' || style.visibility === 'hidden') continue;

                // 후보 항목 탐색: 텍스트 노드를 직접 가진 요소 우선
                const candidates = popup.querySelectorAll(
                    '.dijitComboBoxItem, .dijitMenuItem, [item]'
                );
                for (const el of candidates) {
                    if (el.textContent.trim().toLowerCase() === lower) {
                        el.click();
                        return 'clicked: ' + text;
                    }
                }
                // fallback: leaf 텍스트 노드를 가진 아무 요소
                const leaves = popup.querySelectorAll('div, span, li, td');
                for (const el of leaves) {
                    if (el.children.length === 0 &&
                        el.textContent.trim().toLowerCase() === lower) {
                        el.click();
                        return 'clicked (leaf): ' + text;
                    }
                }
            }
            // 디버그: 보이는 팝업 목록
            const visible = Array.from(document.querySelectorAll(
                '.dijitComboBoxPopup, .dijitPopup'
            )).filter(p => window.getComputedStyle(p).display !== 'none')
              .map(p => p.className + ':' + p.children.length);
            return 'not found; popups=' + visible.join('|');
        }""",
        text,
    )
    # 디버그 로그에 결과 기록
    try:
        with open("wings_debug.log", "a", encoding="utf-8") as _f:
            _f.write(f"  _click_popup_item('{text}'): {result}\n")
    except Exception:
        pass
    return result.startswith("clicked")


def _write_debug(row_idx: int, log: list):
    try:
        mode = "a" if row_idx > 0 else "w"
        with open("wings_debug.log", mode, encoding="utf-8") as _f:
            _f.write(f"[filter_row={row_idx}] {log}\n")
    except Exception:
        pass


def download_wings_excel(months: list, download_dir: str = None, on_status=None, auth_code_callback=None) -> str:
    """WINGS에서 Excel 파일을 동기적으로 다운로드한다."""
    if not download_dir:
        download_dir = tempfile.mkdtemp(prefix="wings_dl_")

    # Streamlit에서 호출 시 → bat 파일로 새 CMD 창 띄워 실행 (_wings_daily.bat 방식)
    _in_streamlit = False
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        _in_streamlit = get_script_run_ctx() is not None
    except Exception:
        pass

    if _in_streamlit:
        return _download_via_bat(months, on_status)

    return _download_in_process(months, download_dir, on_status, auth_code_callback)


def _download_in_process(months, download_dir, on_status, auth_code_callback):
    """스케줄러/직접 호출용 — _wings_daily.py에서 사용."""
    loop = asyncio.ProactorEventLoop()
    try:
        return loop.run_until_complete(
            _wings_download_async(months, download_dir, on_status, auth_code_callback)
        )
    finally:
        loop.close()


def _download_via_bat(months, on_status):
    """_wings_daily.bat과 동일한 방식: CMD 창에서 Python 실행 → Chrome 창 보임."""
    import json, time as _time

    _project_dir = os.path.dirname(os.path.abspath(__file__))
    _venv_py = os.path.join(_project_dir, ".venv", "Scripts", "python.exe")
    if not os.path.exists(_venv_py):
        _venv_py = sys.executable

    _dl_dir = os.path.join(_project_dir, "_wings_dl")
    os.makedirs(_dl_dir, exist_ok=True)
    _status_file = os.path.join(_dl_dir, "_status.txt")
    _result_file = os.path.join(_dl_dir, "_result.json")
    _py_script   = os.path.join(_project_dir, "_wings_run.py")
    _bat_script  = os.path.join(_project_dir, "_wings_run.bat")

    for _f in [_status_file, _result_file]:
        try: os.remove(_f)
        except OSError: pass

    # _wings_daily.py와 동일한 구조의 Python 스크립트 생성
    with open(_py_script, "w", encoding="utf-8") as f:
        f.write(f"""\
import sys, json, os
sys.path.insert(0, r"{_project_dir}")
from wings_scraper import _download_in_process

def on_status(msg):
    print(msg)
    with open(r"{_status_file}", "w", encoding="utf-8") as fh:
        fh.write(msg)

try:
    path = _download_in_process({months!r}, r"{_dl_dir}", on_status, None)
    with open(r"{_result_file}", "w", encoding="utf-8") as fh:
        json.dump({{"ok": True, "path": path}}, fh)
    print("SUCCESS:", path)
except Exception as e:
    import traceback
    with open(r"{_result_file}", "w", encoding="utf-8") as fh:
        json.dump({{"ok": False, "error": traceback.format_exc()}}, fh)
    print("ERROR:", e)
    input("Press Enter to close...")
""")

    # _wings_daily.bat과 동일한 구조의 BAT 파일 생성
    with open(_bat_script, "w", encoding="utf-8") as f:
        f.write(f"""\
@echo on
title WINGS Auto-Fetch
echo Starting WINGS Auto-Fetch...
"{_venv_py}" "{_py_script}"
echo.
echo Done.
timeout /t 5
""")

    # CREATE_NEW_CONSOLE로 cmd /c bat 실행 → CMD 창이 뜨고 그 안에서 Python+Chrome 실행
    subprocess.Popen(
        ["cmd", "/c", _bat_script],
        creationflags=subprocess.CREATE_NEW_CONSOLE,
    )

    if on_status:
        on_status("WINGS 다운로드 창이 열립니다...")

    # 결과 대기 (최대 10분)
    _last = ""
    for _ in range(600):
        _time.sleep(1)
        if on_status and os.path.exists(_status_file):
            try:
                _new = open(_status_file, encoding="utf-8").read().strip()
                if _new and _new != _last:
                    _last = _new
                    on_status(_new)
            except Exception:
                pass
        if os.path.exists(_result_file):
            try:
                result = json.load(open(_result_file, encoding="utf-8"))
            except Exception:
                continue
            for _f in (_status_file, _result_file):
                try: os.remove(_f)
                except Exception: pass
            if result["ok"]:
                return result["path"]
            else:
                raise RuntimeError(result["error"])

    raise RuntimeError("WINGS 다운로드 타임아웃 (10분 초과)")
