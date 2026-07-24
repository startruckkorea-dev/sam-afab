# SharePoint 연동 설정 (데이터 빌드 · 모델 매칭 · 코드 관리)

`데이터 빌드`, `모델 매칭`, `코드 관리` 는 모두 브라우저에서 **Microsoft Graph** 를 호출해
SharePoint 문서 라이브러리를 **직접 읽고 저장**합니다. 서버·GitHub Actions·클라이언트
시크릿·PAT 는 쓰지 않고, **로그인한 사용자 본인의 위임 권한**만 사용합니다.

앱(클라이언트) ID: `9b247088-5afb-4622-9c5e-b5f27142761d`
테넌트 ID: `19cab1f5-21f4-44df-8ac6-96d6ca595203`

---

## 1) API 권한 — ✅ 이미 완료 (추가 작업 없음)

이 앱(`9b247088-…`)에는 아래 **위임 권한이 이미 부여되고 관리자 동의까지 완료**돼 있습니다.
같은 앱 등록을 **mbtruck-spec**(https://mbtruck-spec.startruckkorea.com) 이 이미 사용 중입니다.

- `User.Read`
- `Sites.ReadWrite.All`
- `Files.ReadWrite.All`
- `Mail.Send`

따라서 **모델 매칭 / 코드 관리 메뉴의 SharePoint 불러오기·저장은 추가 설정 없이 바로 동작**합니다.
(첫 사용 시 브라우저 팝업으로 사용자 동의만 한 번 뜰 수 있습니다.)

> 저장 권한은 결국 **로그인한 사용자 본인의 SharePoint 권한**을 따릅니다 —
> 해당 폴더에 쓰기 권한이 있는 사람만 저장됩니다.

## 2) Redirect URI 확인 (SPA)

앱 등록 → **인증 → 플랫폼: SPA(단일 페이지 애플리케이션)** 에 아래가 등록돼 있어야 함:

- `https://sam-afab.startruckkorea.com/`  (운영, 끝 슬래시 포함)
- (선택) `http://localhost:8000/` 등 — 로컬에서 SharePoint 편집을 테스트할 때만

## 3) 연동되는 SharePoint 위치

- 사이트: `https://startruckkorea.sharepoint.com/sites/SAM-AFAB`
- 문서 라이브러리(기본 drive) = `Shared Documents`
- SAM 원본: `SAM-AFAB_Data/01. SAM_files` (빌드는 **최신 생산월** 폴더만 읽음)
- WINGS export: `SAM-AFAB_Data/02. WINGS_data` (빌드는 **최신 파일**만 읽음)
- 모델 매칭 폴더: `SAM-AFAB_Data/03. model_rules`  → `model_mapping.xlsx`
- 코드 관리 폴더: `SAM-AFAB_Data/04. code`         → 폴더 내 모든 `.xlsx`
- 빌드 결과: `SAM-AFAB_Data/05. output` → `data.json` / `codes.json` (없으면 빌드가 폴더를 만듦)

경로/사이트가 바뀌면 [`graph.js`](graph.js) 상단의 `HOSTNAME` / `SITE_PATH` / `FOLDERS` 를 수정하세요.

---

## 동작 방식 (요약)

- **불러오기**: 웹 → Graph `GET .../content` 로 xlsx 다운로드 → SheetJS 로 파싱해 표에 표시.
- **저장**: 표 편집분을 SheetJS 로 xlsx 로 다시 만들고 Graph `PUT .../content` 로 **덮어쓰기**.
- 첫 저장 시 브라우저 팝업으로 `Sites.ReadWrite.All` **증분 동의**를 한 번 요청합니다.
- 저장 권한은 **로그인한 사용자 본인의 SharePoint 권한**을 그대로 따릅니다
  (그 폴더에 쓰기 권한이 있는 사람만 저장 가능).

---

## 4) 데이터 빌드 — 추가 설정 없음 ✅

**⟳ 데이터 빌드** 버튼은 관리자의 브라우저에서 파이프라인(`lib/pipeline.js`)을 직접 돌립니다.

1. Graph 로 SharePoint 원본을 읽는다 — 최신 WINGS(02) · 최신 생산월 SAM(01) · 규칙(03) · 코드(04)
2. `.docx`/`.xlsx`/`.csv` 를 브라우저에서 파싱하고 비교한다 (SheetJS + 내장 ZIP 리더)
3. 결과 `data.json` / `codes.json` 을 `05. output` 에 저장한다

그래서 **앱 전용(application) 권한 · 클라이언트 시크릿 · GitHub Actions · PAT 가 모두 불필요**합니다.
1)번의 위임 권한(`Sites.ReadWrite.All`)만으로 동작하며, 실제 읽기/쓰기 가능 범위는
**로그인한 사용자 본인의 SharePoint 권한**을 따릅니다 — 즉 `05. output` 에 쓰기 권한이 있는
사람만 빌드를 완료할 수 있고, 나머지 사용자는 저장된 결과를 **읽기만** 합니다.

> 빌드는 이 브라우저에서 1~3분 정도 걸립니다(파일 수에 비례). 진행 상황은 우측 하단
> 로그 패널에 표시됩니다. 탭을 닫으면 중단되므로 완료 메시지까지 열어두세요.

## 5) 다른 사용자에게 반영되는 방식

- 대시보드는 열릴 때 `05. output/data.json` 을 먼저 읽습니다 → **재계산 없이 최신 결과 열람**.
- 아직 빌드 결과가 없거나 Graph 접근이 실패하면, 사이트에 함께 배포된 `docs/data.json` 사본을
  대신 표시합니다(상단 메타 줄에 출처가 표시됨).

## 문제 해결

| 증상 | 원인 / 해결 |
|------|-------------|
| `로그인이 필요합니다` | 회사 도메인(`sam-afab.startruckkorea.com`)에서 M365 로그인 후 사용 |
| `Graph 403 — Access denied` | 1) 관리자 동의 미완료 → 위 1)번, 2) 해당 폴더 쓰기 권한 없음 |
| `Graph 404` | 폴더/파일 경로 불일치 → `graph.js` 의 `FOLDERS` 경로 확인 |
| 빌드 중 `01. SAM_files 에 YYYY-MM 하위폴더가 없습니다` | SAM 폴더 이름이 `2026-07 생산` 형식(앞이 `YYYY-MM`)인지 확인 |
| `엑셀 라이브러리를 불러오지 못했습니다` | 사내망이 CDN 차단 + `vendor/xlsx.full.min.js` 누락 → self-host 파일 확인 |
| 빌드는 됐는데 다른 사람에게 안 보임 | `05. output` 폴더 열람 권한 확인 (읽기 권한 필요) |
