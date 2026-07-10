# SAM × AFAB Comparison (HTML / GitHub Pages 버전)

기존 Streamlit 앱을 **GitHub Actions(계산) + GitHub Pages(표시)** 구조로 전환한 버전입니다.

> 핵심 원리: **"계산은 Actions에서, 표시만 GitHub Pages에서"**
> 브라우저(HTML)는 이미 완성된 JSON을 읽어서 보여주기만 합니다.

---

## 왜 이 구조인가 (4가지 문제 해결)

| 문제 | 해결 | 방법 |
|------|------|------|
| `.docx` 파싱 | ✅ | 브라우저가 아닌 **GitHub Actions(Python)** 에서 파싱 → 결과만 JSON 저장 |
| GitHub 토큰 보안 | ✅ | 토큰/자격증명은 **GitHub Secrets** 에만 존재, 클라이언트 HTML 에는 없음 |
| rapidfuzz(퍼지 검색) | ✅ | 비교·분석이 Actions 에서 끝난 JSON 을 **표시만** → 클라이언트 퍼지 검색 불필요 |
| WINGS 스크래핑 | ✅ | 원래부터 Actions 로 분리하는 것이 이 아키텍처의 핵심 |

---

## 디렉터리 구조

```
SAM_AFAB_Github/
├── README.md                      # 이 문서
├── requirements.txt               # 백엔드 의존성 (Streamlit 없음)
├── .gitignore
├── .github/workflows/build.yml    # ① WINGS 스크래핑 ② 파싱+비교 ③ data.json 커밋
│
├── backend/                       # "계산" — 순수 Python, Streamlit 의존성 제거됨
│   ├── option_codes.py            # OPTION_CODE_MAP / MANDATORY_CODES (참조 데이터, 원본에서 추출)
│   ├── wings_parser.py            # WINGS CSV/Excel 파싱 → DataFrame
│   ├── sam_parser.py              # SAM .docx 파싱 → {모델: {PTO여부: {codes, file}}}
│   ├── compare.py                 # SAM ↔ WINGS 비교 로직
│   ├── wings_scraper.py           # WINGS 스크래핑 (Playwright, 원본 재활용)
│   └── build_data.py              # 오케스트레이터 → docs/data.json + codes.json
│
├── docs/                          # "표시" — GitHub Pages 루트
│   ├── index.html
│   ├── style.css
│   ├── app.js                     # data.json / codes.json 읽어서 테이블 렌더링
│   ├── data.json                  # Actions 가 생성·커밋 (비교 결과)
│   └── codes.json                 # Actions 가 생성·커밋 (코드 설명 사전)
│
└── sam_files/                     # SAM 원본 (월별 폴더: 2026_04, 2026_05 …)
    └── 2026_04/Internal quotation_*.docx
```

---

## 데이터 흐름

```
[sam_files/YYYY_MM/*.docx]  ─┐
                             ├─► backend/build_data.py ─► docs/data.json ─► docs/index.html (브라우저)
[WINGS 스크래핑/업로드]      ─┘                          └► docs/codes.json
        ↑
   GitHub Secrets (토큰/계정/TOTP) — Actions 안에서만 사용
```

1. **GitHub Actions** (`build.yml`) 가 정해진 시각(매일 06:00 KST)에 실행
2. 필요 시 `wings_scraper.py` 로 WINGS 다운로드 (`--scrape`)
3. `build_data.py` 가 SAM/WINGS 파싱 → `compare()` → `docs/data.json` 생성
4. 변경분을 자동 커밋 → GitHub Pages 가 즉시 갱신
5. 사용자는 정적 HTML 만 보면 됨 (서버·로그인 불필요)

---

## 로컬 실행

```bash
pip install -r requirements.txt

# (A) 이미 받아둔 WINGS 파일로 빌드
python backend/build_data.py --wings path/to/wings.xlsx

# (B) wings_data/ 폴더의 최신 파일 자동 사용
python backend/build_data.py

# (C) WINGS 스크래핑부터 (자격증명 필요)
python backend/build_data.py --scrape 2026_04 2026_05

# 결과 확인: docs/ 를 정적 서버로 띄우기
python -m http.server -d docs 8000   # → http://localhost:8000
```

---

## GitHub 설정

### 1) Secrets 등록  (Settings → Secrets and variables → Actions)
| 이름 | 설명 |
|------|------|
| `WINGS_USER` | WINGS 로그인 ID |
| `WINGS_PASSWORD` | WINGS 비밀번호 |
| `WINGS_TOTP_SECRET` | 2차 인증 TOTP 시크릿 |

> ⚠️ 자격증명은 **절대 코드에 넣지 말 것.** Secrets 에만 보관됩니다.

### 2) GitHub Pages 활성화  (Settings → Pages)
- Source: **Deploy from a branch**
- Branch: `main` / 폴더: **`/docs`**

### 3) 워크플로우 실행
- 자동: 매일 06:00 KST (`build.yml` 의 cron)
- 수동: Actions 탭 → **Build comparison data** → Run workflow
  - `months` 입력에 `2026_04 2026_05` 처럼 넣으면 스크래핑 후 빌드

---

## 원본(Streamlit)에서 재활용 / 제거된 것

**재활용 (Streamlit 의존성만 제거):**
- `OPTION_CODE_MAP`, `MANDATORY_CODES`, `MANDATORY_GROUPS` → `option_codes.py`
- `parse_wings`, `_normalize_model`, SAM `.docx` 파서, `compare` 로직
- WINGS 스크래퍼 (`wings_scraper.py`) 전체

**제거 (클라이언트에 불필요):**
- `streamlit`, `st.cache_*`, `st.session_state`, `st.dialog` 등 모든 UI 코드
  → 사용자 커스터마이즈 코드 집합은 `compare.py` 의 기본값(default)으로 대체
- `rapidfuzz` (퍼지 검색) — 비교는 Actions 에서 끝나므로 불필요
- GitHub 토큰을 다루던 영속화 로직 — Secrets + 자동 커밋으로 대체

---

## TODO / 다음 단계
- [ ] GitHub Actions 환경에서 `wings_scraper.py` 의 Chrome 프로필 의존성 검증 (headless)
- [ ] 사용자 커스텀 코드(예외/필수 코드 편집) 가 필요하면 별도 JSON 설정 파일로 분리
- [ ] `docs/data.json` 초기값은 샘플 데이터 → 첫 실 빌드 후 교체
