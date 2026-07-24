# SAM × AFAB Comparison (정적 사이트 + SharePoint)

SAM(Internal Quotation `.docx`) 과 WINGS(AFAB) export 를 비교해 사양 불일치를 보여주는
대시보드입니다. **서버가 없습니다.** GitHub Pages 로 배포된 정적 파일과
SharePoint(Microsoft Graph) 만으로 동작합니다.

> 핵심 원리: **"계산도 표시도 브라우저에서, 데이터는 SharePoint 에서"**
> 관리자가 `⟳ 데이터 빌드` 를 누르면 **그 브라우저에서** 파이프라인이 돌고,
> 결과 JSON 이 SharePoint 에 저장됩니다. 다른 사용자는 그 결과를 **읽기만** 합니다
> (재계산 없음).

---

## 데이터 소스 = SharePoint (repo 에 원본 없음)

SAM/WINGS/코드/모델규칙 원본은 **repo 에 커밋하지 않고 SharePoint 에서만** 관리합니다.

| SharePoint 폴더 | 용도 |
|---|---|
| `SAM-AFAB_Data/01. SAM_files` | SAM 원본 — 빌드는 **가장 최근 생산월** 폴더만 사용 (`YYYY-MM …`) |
| `SAM-AFAB_Data/02. WINGS_data` | WINGS export — 빌드는 **최신 파일**만 사용 |
| `SAM-AFAB_Data/03. model_rules` | `model_mapping.xlsx` — **모든 모델 매칭 규칙** |
| `SAM-AFAB_Data/04. code` | 코드 사전 / 필수코드 / 카테고리 / cab xlsx |
| `SAM-AFAB_Data/05. output` | **빌드 결과** `data.json` · `codes.json` (대시보드가 읽는 파일) |

접근은 전부 **로그인한 사용자 본인의 위임 권한(Graph)** 으로 이뤄집니다.
클라이언트 시크릿·GitHub Actions·PAT 는 사용하지 않습니다.
설정은 [`docs/ENTRA_SETUP.md`](docs/ENTRA_SETUP.md) 참고.

---

## 동작 흐름

```
[관리자 브라우저]  ⟳ 데이터 빌드
   │  Graph 로 읽기: 01 SAM(최신 월) · 02 WINGS(최신) · 03 규칙 · 04 코드
   │  파싱/비교: docs/lib/{unzip,samparse,wingsparse,refdata,compare,pipeline}.js
   └► Graph 로 저장: 05. output/data.json · codes.json
                                    │
[모든 사용자 브라우저]  대시보드 로드 ◄┘  (없으면 docs/data.json 사본으로 폴백)
```

- **모델 매칭 / 코드 관리** 메뉴에서 `03. model_rules` · `04. code` 의 xlsx 를 웹에서 직접
  편집·저장할 수 있고, 다음 **데이터 빌드** 때 반영됩니다.
- 빌드는 기본적으로 **최신 생산월** SAM 만 사용합니다.

---

## 디렉터리 구조

```
SAM_AFAB_Github/
├── README.md
├── requirements.txt               # 로컬 Python 실행용
├── config.json                    # 로컬 실행 시 SAM/WINGS/코드 폴더 경로
│
├── docs/                          # GitHub Pages 루트 (배포되는 전부)
│   ├── index.html · style.css
│   ├── auth.js                    # M365(Entra) 로그인 게이트 (MSAL)
│   ├── graph.js                   # SharePoint 읽기/쓰기 (Graph)
│   ├── app.js                     # 대시보드 · 편집 화면 · 빌드 버튼
│   ├── lib/                       # 브라우저 비교 파이프라인
│   │   ├── unzip.js               #   .docx(zip) 리더 (DecompressionStream)
│   │   ├── samparse.js            #   SAM .docx 파서
│   │   ├── wingsparse.js          #   WINGS xlsx/csv 파서
│   │   ├── refdata.js             #   참조 워크북 로더 (코드/필수/카테고리/규칙)
│   │   ├── compare.js             #   비교 로직
│   │   ├── pipeline.js            #   오케스트레이터 (SharePoint → JSON)
│   │   └── optioncodes.json       #   option_codes.py 사본 (파서용 코드 사전)
│   ├── data.json · codes.json     # 폴백 사본 (SharePoint 결과가 1순위)
│   └── vendor/                    # SheetJS · MSAL self-host (사내망 CDN 차단 대비)
│
└── backend/                       # 로컬 전용 Python 참조 구현
    ├── option_codes.py · wings_parser.py · sam_parser.py · compare.py
    ├── mandatory_codes.py · model_category.py · rules.py
    ├── build_data.py              # 로컬 폴더로 data.json/codes.json 생성
    ├── build_option_codes_json.py # option_codes.py → docs/lib/optioncodes.json
    ├── build_*_xlsx.py            # 참조 워크북 생성/갱신 도구
    └── wings_scraper.py           # WINGS 스크래핑 (로컬 전용, Playwright)
```

`docs/lib/*.js` 는 `backend/*.py` 의 포팅이며, 동일 입력에서 **445행 × 전 필드 일치**로
검증되었습니다. 로직을 고칠 때는 양쪽을 함께 수정하고 아래 로컬 빌드로 대조하세요.

---

## 로컬 실행

```bash
pip install -r requirements.txt

# 로컬 폴더(sam_files/ · wings_data/ · code/ · model_rules/)로 빌드
python backend/build_data.py                 # 최신 생산월만
python backend/build_data.py --all-months    # 전체 월
python backend/build_data.py --wings path/to/wings.xlsx

# option_codes.py 를 고쳤다면 브라우저용 사본도 갱신
python backend/build_option_codes_json.py

# 정적 서버로 확인 (SharePoint 연동은 배포 도메인에서만 로그인됨)
python -m http.server -d docs 8000   # → http://localhost:8000
```

---

## GitHub 설정

- **Pages**: Settings → Pages → Deploy from a branch → `main` / `/docs`
- **Secrets / Actions 불필요** — 빌드는 브라우저에서 실행됩니다.
- 커스텀 도메인: `docs/CNAME` (`sam-afab.startruckkorea.com`), 로그인 게이트는 이 도메인에서만 강제됩니다.

---

## 참고

- 원본 Streamlit 앱의 참조 데이터(`OPTION_CODE_MAP` 등)와 파싱/비교 로직을 이식했고,
  Streamlit·rapidfuzz 의존성은 제거했습니다.
- 자격증명(WINGS 계정/TOTP)은 로컬 스크래핑에만 쓰이며 저장소에 커밋하지 않습니다
  (`.gitignore` 의 `.wings_credentials`, `.totp_secret`).
