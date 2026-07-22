"""SharePoint (Microsoft Graph, app-only) source for the build.

The build no longer keeps SAM/WINGS/code/model_rules inside the repo. Instead it
pulls them fresh from the SharePoint document library each run, using **app-only**
(client-credentials) Graph auth so it works headlessly in GitHub Actions.

Required env (GitHub Secrets):
    GRAPH_TENANT_ID       Entra tenant id
    GRAPH_CLIENT_ID       app registration (client) id
    GRAPH_CLIENT_SECRET   client secret value

The app registration needs the **application** Graph permission `Sites.Read.All`
(read) — and `Sites.ReadWrite.All` only if the build should also write back —
with admin consent granted. See docs/ENTRA_SETUP.md.

SharePoint layout (site https://startruckkorea.sharepoint.com/sites/SAM-AFAB,
default document library "Shared Documents"):
    SAM-AFAB_Data/01. SAM_files     <- SAM sources, in YYYY-MM month subfolders
    SAM-AFAB_Data/02. WINGS_data    <- WINGS exports (newest file is used)
    SAM-AFAB_Data/03. model_rules   <- model_mapping.xlsx (all matching rules)
    SAM-AFAB_Data/04. code          <- code dictionaries / mandatory / category xlsx
"""
from __future__ import annotations

import os
import re
import urllib.parse
from pathlib import Path

import requests

GRAPH = 'https://graph.microsoft.com/v1.0'
HOSTNAME = 'startruckkorea.sharepoint.com'
SITE_PATH = 'sites/SAM-AFAB'

FOLDERS = {
    'sam':         'SAM-AFAB_Data/01. SAM_files',
    'wings':       'SAM-AFAB_Data/02. WINGS_data',
    'model_rules': 'SAM-AFAB_Data/03. model_rules',
    'code':        'SAM-AFAB_Data/04. code',
}

_TIMEOUT = 120


class SharePointError(RuntimeError):
    pass


def _enc(path: str) -> str:
    """Encode each path segment (spaces/dots in folder names) but keep slashes."""
    return '/'.join(urllib.parse.quote(seg) for seg in str(path).split('/'))


class Graph:
    def __init__(self, tenant=None, client_id=None, client_secret=None):
        self.tenant = tenant or os.environ.get('GRAPH_TENANT_ID')
        self.client_id = client_id or os.environ.get('GRAPH_CLIENT_ID')
        self.client_secret = client_secret or os.environ.get('GRAPH_CLIENT_SECRET')
        if not (self.tenant and self.client_id and self.client_secret):
            raise SharePointError(
                'Missing Graph credentials: set GRAPH_TENANT_ID / GRAPH_CLIENT_ID / '
                'GRAPH_CLIENT_SECRET (see docs/ENTRA_SETUP.md).')
        self._token = None
        self._site_id = None

    # ---- auth ----
    def token(self) -> str:
        if self._token:
            return self._token
        url = f'https://login.microsoftonline.com/{self.tenant}/oauth2/v2.0/token'
        data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': 'https://graph.microsoft.com/.default',
            'grant_type': 'client_credentials',
        }
        r = requests.post(url, data=data, timeout=_TIMEOUT)
        if not r.ok:
            raise SharePointError(f'token request failed {r.status_code}: {r.text[:300]}')
        self._token = r.json()['access_token']
        return self._token

    def _headers(self):
        return {'Authorization': 'Bearer ' + self.token()}

    def _get(self, url):
        full = url if url.startswith('http') else GRAPH + url
        r = requests.get(full, headers=self._headers(), timeout=_TIMEOUT)
        if not r.ok:
            raise SharePointError(f'GET {url} -> {r.status_code}: {r.text[:300]}')
        return r

    # ---- site / drive ----
    def site_id(self) -> str:
        if self._site_id:
            return self._site_id
        j = self._get(f'/sites/{HOSTNAME}:/{SITE_PATH}').json()
        self._site_id = j['id']
        return self._site_id

    def _root(self):
        return f'/sites/{self.site_id()}/drive/root'

    # ---- listing ----
    def list_children(self, folder_key_or_path: str) -> list[dict]:
        path = FOLDERS.get(folder_key_or_path, folder_key_or_path)
        items, url = [], (self._root() + ':/' + _enc(path) + ':/children'
                          '?$select=name,size,file,folder,lastModifiedDateTime&$top=999')
        while url:
            j = self._get(url).json()
            items.extend(j.get('value', []))
            url = j.get('@odata.nextLink')
        return items

    # ---- download ----
    def download_file(self, folder_key_or_path: str, filename: str, dest: Path) -> Path:
        path = FOLDERS.get(folder_key_or_path, folder_key_or_path) + '/' + filename
        r = self._get(self._root() + ':/' + _enc(path) + ':/content')
        dest = Path(dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(r.content)
        return dest

    def upload_file(self, folder_key_or_path: str, filename: str, data: bytes) -> dict:
        """Overwrite (or create) a file with raw bytes. Needs Sites.ReadWrite.All."""
        path = FOLDERS.get(folder_key_or_path, folder_key_or_path) + '/' + filename
        url = GRAPH + self._root() + ':/' + _enc(path) + ':/content'
        r = requests.put(url, headers={**self._headers(),
                                       'Content-Type': 'application/octet-stream'},
                         data=data, timeout=_TIMEOUT)
        if not r.ok:
            raise SharePointError(f'PUT {path} -> {r.status_code}: {r.text[:300]}')
        return r.json()

    def download_folder(self, folder_key_or_path: str, dest_dir: Path,
                        exts=None, log=None) -> list[Path]:
        """Download every file (optionally filtered by extension) into dest_dir."""
        dest_dir = Path(dest_dir)
        dest_dir.mkdir(parents=True, exist_ok=True)
        out = []
        for it in self.list_children(folder_key_or_path):
            if it.get('folder'):
                continue
            name = it['name']
            if exts and Path(name).suffix.lower() not in exts:
                continue
            self.download_file(folder_key_or_path, name, dest_dir / name)
            out.append(dest_dir / name)
            if log:
                log(f'downloaded {name} ({it.get("size", 0)} B)')
        return out


# ---- higher-level helpers used by build_data ----

def _recency(name: str) -> float:
    """Best timestamp from a WINGS filename: trailing 13-digit epoch-ms or a date."""
    m = re.search(r'(\d{13})', name)
    if m:
        return int(m.group(1)) / 1000.0
    m = re.search(r'(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})', name)
    if m:
        try:
            from datetime import datetime, timezone
            y, mo, d = (int(x) for x in m.groups())
            return datetime(y, mo, d, tzinfo=timezone.utc).timestamp()
        except Exception:
            pass
    return 0.0


def fetch_latest_wings(g: Graph, dest_dir: Path, log=None) -> Path:
    """Download only the newest WINGS export from 02. WINGS_data."""
    files = [it for it in g.list_children('wings')
             if not it.get('folder')
             and Path(it['name']).suffix.lower() in {'.xlsx', '.xls', '.csv'}
             and not it['name'].startswith('~$')]
    if not files:
        raise SharePointError('no WINGS files in 02. WINGS_data')
    chosen = max(files, key=lambda it: (_recency(it['name']), it.get('lastModifiedDateTime', '')))
    if log:
        log(f'{len(files)} WINGS files -> newest: {chosen["name"]}')
    return g.download_file('wings', chosen['name'], Path(dest_dir) / chosen['name'])


_MONTH_RE = re.compile(r'^(\d{4})[_-](\d{2})\b')


def fetch_latest_sam_month(g: Graph, dest_root: Path, log=None) -> tuple[str, Path]:
    """Download the SAM files of the most recent production-month subfolder.

    01. SAM_files contains 'YYYY-MM ...' subfolders. Returns (folder_name, local_dir)
    where local_dir holds a single 'YYYY-MM' subfolder so load_sam_by_month picks it up.
    """
    subs = []
    for it in g.list_children('sam'):
        if not it.get('folder'):
            continue
        m = _MONTH_RE.match(it['name'])
        if m:
            subs.append((int(m.group(1)) * 100 + int(m.group(2)), it['name']))
    if not subs:
        raise SharePointError('no YYYY-MM subfolders in 01. SAM_files')
    yyyymm, folder_name = max(subs)
    if log:
        log(f'{len(subs)} SAM months -> latest: {folder_name} ({yyyymm})')
    local_dir = Path(dest_root) / folder_name
    g.download_folder('sam/' + folder_name, local_dir,
                      exts={'.docx', '.csv', '.txt'}, log=log)
    return folder_name, local_dir


def fetch_folder(g: Graph, folder_key: str, dest_dir: Path, exts=None, log=None) -> Path:
    g.download_folder(folder_key, dest_dir, exts=exts, log=log)
    return Path(dest_dir)
