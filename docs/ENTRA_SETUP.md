# SharePoint 연동 설정 (모델 매칭 · 코드 관리)

`모델 매칭`과 `코드 관리` 메뉴는 브라우저에서 **Microsoft Graph** 를 호출해
SharePoint 문서 라이브러리의 Excel 파일을 **직접 읽고 저장**합니다.
이 기능이 작동하려면 Entra(Azure AD) 앱 등록에 아래 설정이 한 번 필요합니다.

앱(클라이언트) ID: `9b247088-5afb-4622-9c5e-b5f27142761d`
테넌트 ID: `19cab1f5-21f4-44df-8ac6-96d6ca595203`

---

## 1) API 권한 추가 (관리자 작업)

Azure Portal → **Microsoft Entra ID → 앱 등록 → (이 앱) → API 사용 권한**

1. **권한 추가 → Microsoft Graph → 위임된 권한(Delegated)**
2. 다음을 추가:
   - `Sites.ReadWrite.All`  ← SharePoint 사이트 파일 읽기/쓰기 (필수)
   - `User.Read`            ← 이미 있음 (로그인)
3. **`<테넌트>에 대한 관리자 동의 허용`** 클릭 → 상태가 모두 초록색이 되도록.

> `Sites.ReadWrite.All` 은 관리자 동의가 필요한 권한입니다.
> 동의 전에는 저장 시 `Graph 403 — Access denied` 가 납니다.

## 2) Redirect URI 확인 (SPA)

앱 등록 → **인증 → 플랫폼: SPA(단일 페이지 애플리케이션)** 에 아래가 등록돼 있어야 함:

- `https://sam-afab.startruckkorea.com/`  (운영, 끝 슬래시 포함)
- (선택) `http://localhost:8000/` 등 — 로컬에서 SharePoint 편집을 테스트할 때만

## 3) 연동되는 SharePoint 위치

- 사이트: `https://startruckkorea.sharepoint.com/sites/SAM-AFAB`
- 문서 라이브러리(기본 drive) = `Shared Documents`
- 모델 매칭 폴더: `SAM-AFAB_Data/03. model_rules`  → `model_mapping.xlsx`
- 코드 관리 폴더: `SAM-AFAB_Data/04. code`         → 폴더 내 모든 `.xlsx`

경로/사이트가 바뀌면 [`graph.js`](graph.js) 상단의 `HOSTNAME` / `SITE_PATH` / `FOLDERS` 를 수정하세요.

---

## 동작 방식 (요약)

- **불러오기**: 웹 → Graph `GET .../content` 로 xlsx 다운로드 → SheetJS 로 파싱해 표에 표시.
- **저장**: 표 편집분을 SheetJS 로 xlsx 로 다시 만들고 Graph `PUT .../content` 로 **덮어쓰기**.
- 첫 저장 시 브라우저 팝업으로 `Sites.ReadWrite.All` **증분 동의**를 한 번 요청합니다.
- 저장 권한은 **로그인한 사용자 본인의 SharePoint 권한**을 그대로 따릅니다
  (그 폴더에 쓰기 권한이 있는 사람만 저장 가능).

---

## 4) 빌드(GitHub Actions)용 앱 전용 권한 — 데이터 빌드에 필수

`데이터 빌드` 버튼 → GitHub Actions 워크플로(`build.yml`)는 **헤드리스**라 사용자 로그인이
불가합니다. 그래서 **앱 전용(application) 권한 + 클라이언트 시크릿**으로 SharePoint 를 읽습니다.

Azure Portal → 앱 등록 → (이 앱):

1. **API 사용 권한 → 권한 추가 → Microsoft Graph → 애플리케이션 권한(Application)**
   - `Sites.Read.All`  (읽기 전용) — 빌드가 SAM/WINGS/code/model_rules 를 내려받는 데 필요
   - `Sites.ReadWrite.All` — 빌드가 갱신된 `model_mapping.xlsx`(인식모델_대조표) 를 SharePoint 에
     되돌려 저장(writeback)하게 하려면 추가. (원치 않으면 `--no-writeback` 로 실행)
   - **관리자 동의** 클릭.
2. **인증서 및 비밀 → 새 클라이언트 비밀 → 값(Value) 복사** (한 번만 보임).
3. GitHub repo(`startruckkorea-dev/sam-afab`) → **Settings → Secrets and variables → Actions** 에 추가:
   - `GRAPH_TENANT_ID`  = `19cab1f5-21f4-44df-8ac6-96d6ca595203`
   - `GRAPH_CLIENT_ID`  = `9b247088-5afb-4622-9c5e-b5f27142761d`
   - `GRAPH_CLIENT_SECRET` = (2번에서 복사한 비밀 값)

> 위임(3번, `Sites.ReadWrite.All` Delegated)은 **웹 화면에서** 관리자가 편집·저장할 때,
> 애플리케이션(4번)은 **빌드가 헤드리스로** 읽고/되돌려쓸 때 쓰입니다 — 둘 다 필요합니다.

## 5) 데이터 빌드 버튼용 GitHub 토큰 (관리자 브라우저)

상단 **⟳ 데이터 빌드** 버튼은 GitHub `workflow_dispatch` 를 호출합니다. 정적 사이트라
서버가 없으므로, 관리자가 **파인그레인드 PAT** 를 한 번 입력하면 브라우저에 저장됩니다.
- GitHub → Settings → Developer settings → **Fine-grained tokens**
- Repository access: `startruckkorea-dev/sam-afab` 만
- Permissions: **Actions → Read and write**
- 생성 후 버튼 첫 클릭 시 붙여넣기 (localStorage 저장, 재입력 불필요)

## SharePoint 소스 폴더

| 폴더 | 용도 |
|------|------|
| `SAM-AFAB_Data/01. SAM_files` | SAM 원본 (`YYYY-MM ...` 생산월 하위폴더; 빌드는 **최신 월**만 사용) |
| `SAM-AFAB_Data/02. WINGS_data` | WINGS export (빌드는 **최신 파일**만 사용) |
| `SAM-AFAB_Data/03. model_rules` | `model_mapping.xlsx` — 모든 모델 매칭 규칙 |
| `SAM-AFAB_Data/04. code` | 코드 사전 / 필수코드 / 카테고리 등 xlsx |

## 문제 해결

| 증상 | 원인 / 해결 |
|------|-------------|
| `로그인이 필요합니다` | 회사 도메인(`sam-afab.startruckkorea.com`)에서 M365 로그인 후 사용 |
| `Graph 403 — Access denied` | 1) 관리자 동의 미완료 → 위 1)번, 2) 해당 폴더 쓰기 권한 없음 |
| `Graph 404` | 폴더/파일 경로 불일치 → `graph.js` 의 `FOLDERS` 경로 확인 |
| `엑셀 라이브러리를 불러오지 못했습니다` | 사내망이 CDN 차단 + `vendor/xlsx.full.min.js` 누락 → self-host 파일 확인 |
