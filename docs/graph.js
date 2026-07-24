// Microsoft Graph 클라이언트 — SharePoint 문서 라이브러리의 Excel 파일을
// 브라우저에서 직접 읽고/쓰기 위한 얇은 래퍼.
// ------------------------------------------------------------------
// 토큰은 auth.js 가 노출하는 window.MB_AUTH.getToken(scopes) 로 얻는다.
// (MSAL 계정이 없으면 예외를 던지므로, 호출부에서 안내 메시지를 띄운다.)
//
// ⚠ Entra 앱 등록(SPA)에 위임 권한 Sites.ReadWrite.All 이 추가되고
//    관리자 동의가 완료돼야 저장이 작동한다. (docs/ENTRA_SETUP.md 참고)
(function () {
  'use strict';

  var GRAPH = 'https://graph.microsoft.com/v1.0';
  // 이 위임 권한들은 같은 Entra 앱(clientId 9b247088-…)에 이미 관리자 동의가
  // 완료돼 있다(mbtruck-spec 이 동일 앱으로 사용 중). 추가 설정 불필요.
  var SCOPES = ['Sites.ReadWrite.All', 'Files.ReadWrite.All'];

  // SharePoint 사이트 좌표. 공유 폴더 URL 에서 그대로 유도된다.
  //   https://startruckkorea.sharepoint.com/sites/SAM-AFAB
  //   문서 라이브러리(기본 drive) 루트 = "Shared Documents"
  var HOSTNAME = 'startruckkorea.sharepoint.com';
  var SITE_PATH = 'sites/SAM-AFAB';

  // 각 메뉴/빌드가 연동되는 폴더(드라이브 루트 기준 상대 경로) + 사람이 여는 공유 링크.
  // 빌드(pipeline.js)는 sam/wings 를 읽고 결과를 output 에 쓴다.
  var FOLDERS = {
    sam: {
      path: 'SAM-AFAB_Data/01. SAM_files',
    },
    wings: {
      path: 'SAM-AFAB_Data/02. WINGS_data',
    },
    model_rules: {
      path: 'SAM-AFAB_Data/03. model_rules',
      shareUrl: 'https://startruckkorea.sharepoint.com/:f:/r/sites/SAM-AFAB/Shared%20Documents/SAM-AFAB_Data/03.%20model_rules?csf=1&web=1&e=Srdcow',
    },
    code: {
      path: 'SAM-AFAB_Data/04. code',
      shareUrl: 'https://startruckkorea.sharepoint.com/:f:/r/sites/SAM-AFAB/Shared%20Documents/SAM-AFAB_Data/04.%20code?csf=1&web=1&e=XmP3MT',
    },
    // 빌드 결과(data.json / codes.json). 관리자가 빌드하면 여기에 저장되고,
    // 다른 사용자는 재계산 없이 이 파일만 읽는다.
    output: {
      path: 'SAM-AFAB_Data/05. output',
    },
  };

  var _siteIdCache = null;

  function available() {
    return !!(window.MB_AUTH && typeof window.MB_AUTH.getToken === 'function');
  }

  async function token() {
    if (!available()) {
      throw new Error('로그인이 필요합니다. 회사 Microsoft 365 계정으로 로그인 후 사용하세요.');
    }
    return window.MB_AUTH.getToken(SCOPES);
  }

  // Graph 경로에서 각 세그먼트를 인코딩(공백·점 포함 폴더명 대응). 슬래시는 유지.
  function encPath(p) {
    return String(p).split('/').map(encodeURIComponent).join('/');
  }

  // 'code' → 'SAM-AFAB_Data/04. code',  'sam/2026-07 생산' → '.../01. SAM_files/2026-07 생산'.
  // 키가 아니면 드라이브 루트 기준 경로 그대로 쓴다.
  function resolvePath(folderKey) {
    var parts = String(folderKey).split('/');
    var head = FOLDERS[parts[0]];
    if (!head) return String(folderKey);
    parts[0] = head.path;
    return parts.join('/');
  }

  async function api(url, opts) {
    opts = opts || {};
    var t = await token();
    var headers = Object.assign({ Authorization: 'Bearer ' + t }, opts.headers || {});
    var res = await fetch(url.indexOf('http') === 0 ? url : GRAPH + url,
      { method: opts.method || 'GET', headers: headers, body: opts.body });
    if (!res.ok) {
      var detail = '';
      try { var j = await res.json(); detail = (j.error && j.error.message) || ''; } catch (e) {}
      throw new Error('Graph ' + res.status + (detail ? ' — ' + detail : '') + ' (' + url + ')');
    }
    return res;
  }

  async function siteId() {
    if (_siteIdCache) return _siteIdCache;
    var res = await api('/sites/' + HOSTNAME + ':/' + SITE_PATH);
    var j = await res.json();
    _siteIdCache = j.id;
    return _siteIdCache;
  }

  // 폴더 안의 항목 목록 (파일/폴더). key = 'sam' | 'wings' | 'model_rules' | 'code'
  // | 'output' | 'sam/<하위폴더>' 또는 명시 경로. 페이지네이션(@odata.nextLink) 처리.
  async function list(folderKey) {
    var path = resolvePath(folderKey);
    var sid = await siteId();
    var url = '/sites/' + sid + '/drive/root:/' + encPath(path) + ':/children'
      + '?$select=name,size,file,folder,lastModifiedDateTime,webUrl&$top=999';
    var out = [];
    while (url) {
      var res = await api(url);
      var j = await res.json();
      (j.value || []).forEach(function (it) {
        out.push({
          name: it.name,
          size: it.size,
          isFolder: !!it.folder,
          modified: it.lastModifiedDateTime,
          webUrl: it.webUrl,
        });
      });
      url = j['@odata.nextLink'] || null;
    }
    return out;
  }

  // 파일 내용을 ArrayBuffer 로 다운로드. folderKey + 파일명.
  async function download(folderKey, filename) {
    var path = resolvePath(folderKey) + '/' + filename;
    var sid = await siteId();
    var res = await api('/sites/' + sid + '/drive/root:/' + encPath(path) + ':/content');
    return res.arrayBuffer();
  }

  // 폴더가 없으면 만든다(있으면 그대로 둔다). 결과 폴더(05. output) 최초 저장 대비.
  async function ensureFolder(folderKey) {
    var path = resolvePath(folderKey);
    var sid = await siteId();
    try {
      await api('/sites/' + sid + '/drive/root:/' + encPath(path) + '?$select=id');
      return true;
    } catch (e) { /* 없으면 아래에서 생성 */ }
    var parts = path.split('/');
    var name = parts.pop();
    var parent = parts.join('/');
    var url = '/sites/' + sid + '/drive/root'
      + (parent ? ':/' + encPath(parent) + ':' : '') + '/children';
    await api(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: name, folder: {},
        '@microsoft.graph.conflictBehavior': 'fail' }),
    });
    return true;
  }

  // JSON 파일 읽기/쓰기 (빌드 결과 data.json / codes.json 용).
  async function downloadJson(folderKey, filename) {
    var buf = await download(folderKey, filename);
    return JSON.parse(new TextDecoder('utf-8').decode(new Uint8Array(buf)));
  }

  async function uploadJson(folderKey, filename, obj) {
    var bytes = new TextEncoder().encode(JSON.stringify(obj));
    return upload(folderKey, filename, bytes);
  }

  // 파일 저장(덮어쓰기 또는 신규). data = ArrayBuffer | Blob | Uint8Array.
  // 4MB 이하는 단순 PUT. 그 이상이면 업로드 세션(청크) 사용.
  async function upload(folderKey, filename, data) {
    var path = resolvePath(folderKey) + '/' + filename;
    var sid = await siteId();
    var bytes = data instanceof ArrayBuffer ? new Uint8Array(data)
      : (data instanceof Uint8Array ? data : new Uint8Array(await data.arrayBuffer()));

    if (bytes.length <= 4 * 1024 * 1024) {
      var res = await api('/sites/' + sid + '/drive/root:/' + encPath(path) + ':/content',
        { method: 'PUT',
          headers: { 'Content-Type': 'application/octet-stream' },
          body: bytes });
      return res.json();
    }
    return uploadLarge(sid, path, bytes);
  }

  async function uploadLarge(sid, path, bytes) {
    var res = await api('/sites/' + sid + '/drive/root:/' + encPath(path) + ':/createUploadSession',
      { method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }) });
    var session = await res.json();
    var uploadUrl = session.uploadUrl;
    var CHUNK = 5 * 1024 * 1024;  // 배수 320KiB 권장, 여기선 넉넉히
    var total = bytes.length;
    for (var start = 0; start < total; start += CHUNK) {
      var end = Math.min(start + CHUNK, total);
      var slice = bytes.subarray(start, end);
      // 업로드 세션 URL 은 자체 인증 토큰 포함 — Authorization 헤더 없이 PUT.
      var r = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(slice.length),
          'Content-Range': 'bytes ' + start + '-' + (end - 1) + '/' + total,
        },
        body: slice,
      });
      if (!r.ok && r.status !== 202 && r.status !== 201 && r.status !== 200) {
        throw new Error('업로드 세션 실패 ' + r.status);
      }
    }
    return { ok: true };
  }

  window.Graph = {
    available: available,
    scopes: SCOPES,
    folders: FOLDERS,
    siteId: siteId,
    list: list,
    download: download,
    downloadJson: downloadJson,
    upload: upload,
    uploadJson: uploadJson,
    ensureFolder: ensureFolder,
  };
})();
