"""Model-conversion rules — single source of truth, loaded from docs/rules.json.

The rules used to be hardcoded across sam_parser.py / compare.py. They now live
in docs/rules.json so they can be viewed/edited from the web UI (and later from a
SharePoint-hosted file). If the JSON is missing or a key is absent, the built-in
DEFAULTS below are used, so the build never breaks.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
RULES_PATH = ROOT / 'docs' / 'rules.json'
XLSX_PATH = ROOT / 'model_rules' / 'model_mapping.xlsx'

# All rules live on ONE sheet, '매칭규칙':  종류 | 기준 | 값 | 메모   (docs/lib/refdata.js too)
#   대체번호  WINGS number -> SAM number tried only after the exact number found nothing
#             in the production month and the 6 months before        -> normalize_historic
#   SAM별칭   SAM document number -> WINGS numbers it is also known as -> reverse_aliases
#   표시치환  WINGS spelling -> unified display                        -> wings_display_replace
#   차종      vehicle -> numbers                                       -> vehicle_keywords
#   옵션      normalize_28xx_to_26xx -> true/false
# Older workbooks (one sheet per rule kind) are still read when '매칭규칙' is absent.
# 이전모델/현재모델 (generation neighbours) were never used for matching and were dropped.
RULE_SHEET = '매칭규칙'
RULE_HEADER = ['종류', '기준', '값', '메모']
# (kind, rules key, value is a comma list)
RULE_KINDS = [
    ('대체번호', 'normalize_historic', False),
    ('SAM별칭', 'reverse_aliases', True),
    ('표시치환', 'wings_display_replace', False),
    ('차종', 'vehicle_keywords', True),
    ('옵션', None, False),
]
OPTION_KEYS = ['normalize_28xx_to_26xx']
_XLSX_MAP_SHEETS = {
    '정규화_과거번호': 'normalize_historic',
    'WINGS표시치환': 'wings_display_replace',
}
_XLSX_LIST_SHEETS = {
    '차종키워드': 'vehicle_keywords',
    '모델별칭': 'reverse_aliases',
}
_TRUE = ('true', '1', 'yes', 'y', 'on')


def _kind(name):
    n = re.sub(r'\s+', '', str(name or '')).lower()
    for kind, key, is_list in RULE_KINDS:
        if kind.lower() == n or (key and key.lower() == n):
            return kind, key, is_list
    return None

# Built-in defaults == the original hardcoded values (fallback only).
DEFAULTS = {
    'normalize_historic': {
        '3253': '4153', '4140': '4440', '2643': '3343',
        '2851': '2651', '2135': '1835', '2863': '2663', '2853': '2653',
    },
    'normalize_28xx_to_26xx': True,
    'reverse_aliases': {'3253': ['4153']},
    'wings_display_replace': {
        '4140': '4440', '2651 LS': '2851 LS', '2653 LS': '2853 LS',
        '2663 LS': '2863 LS', '2643 A': '3343 A',
    },
    'vehicle_keywords': {
        'Actros-L': ['2651', '2851', '2653', '2853', '2663', '2863', '2143'],
        'Actros': ['3363'],
        'Arocs': ['2643', '3343', '4153', '4453', '3253', '2135', '4440', '4140'],
    },
}


def load_rules_from_xlsx(path: Path | None = None) -> dict:
    """Parse model_rules/model_mapping.xlsx rule sheets into a rules dict.

    Returns only the keys present in the workbook (partial). openpyxl is imported
    lazily and every failure is swallowed so the build never breaks on a bad file.
    """
    p = path or XLSX_PATH
    out: dict = {}
    try:
        from openpyxl import load_workbook  # lazy: optional at build time
        wb = load_workbook(Path(p), read_only=True, data_only=True)
    except Exception:
        return out

    def _cell(row, i):
        return '' if len(row) <= i or row[i] is None else str(row[i]).strip()

    def _rows(ws):
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:  # header
                continue
            if not row:
                continue
            k = _cell(row, 0)
            if k:
                yield k, _cell(row, 1)

    if RULE_SHEET in wb.sheetnames:
        found: dict = {}
        for i, row in enumerate(wb[RULE_SHEET].iter_rows(values_only=True)):
            if i == 0 or not row:
                continue
            kd = _kind(_cell(row, 0))
            k, v = _cell(row, 1), _cell(row, 2)
            if not kd or not k:
                continue
            _kind_name, key, is_list = kd
            if key is None:
                if k in OPTION_KEYS:
                    out[k] = v.lower() in _TRUE
                continue
            table = found.setdefault(key, {})
            if is_list:
                lst = table.setdefault(k, [])
                for x in (p.strip() for p in v.split(',')):
                    if x and x not in lst:
                        lst.append(x)
            else:
                table[k] = v
        out.update(found)
        wb.close()
        return out

    for sheet, key in _XLSX_MAP_SHEETS.items():
        if sheet in wb.sheetnames:
            out[key] = {k: v for k, v in _rows(wb[sheet])}
    for sheet, key in _XLSX_LIST_SHEETS.items():
        if sheet in wb.sheetnames:
            out[key] = {k: [s.strip() for s in v.split(',') if s.strip()]
                        for k, v in _rows(wb[sheet])}
    if '옵션' in wb.sheetnames:
        for k, v in _rows(wb['옵션']):
            if k in OPTION_KEYS:
                out[k] = str(v).strip().lower() in _TRUE

    wb.close()
    return out


def load_rules(path: Path | None = None, xlsx_path: Path | None = None) -> dict:
    """Load rules, merging over DEFAULTS so any missing key falls back.

    Priority (highest wins): model_mapping.xlsx > rules.json > built-in DEFAULTS.
    The xlsx is the primary human-editable source; rules.json remains a fallback.
    """
    p = path or RULES_PATH
    merged = {k: (v.copy() if hasattr(v, 'copy') else v) for k, v in DEFAULTS.items()}
    try:
        data = json.loads(Path(p).read_text(encoding='utf-8'))
        for k in DEFAULTS:
            if k in data and data[k] is not None:
                merged[k] = data[k]
    except Exception:
        pass
    # xlsx overrides json/DEFAULTS for whichever keys it defines
    for k, v in load_rules_from_xlsx(xlsx_path).items():
        if v:  # non-empty sheet only, so an accidentally blank sheet can't wipe a rule
            merged[k] = v
    return merged


def rules_to_rows(rules: dict) -> list:
    """rules dict -> '매칭규칙' sheet rows (header first). Same order as refdata.rulesToAoa."""
    rows = [RULE_HEADER[:]]
    for kind, key, _is_list in RULE_KINDS:
        if key is None:
            continue
        for k, v in (rules.get(key) or {}).items():
            rows.append([kind, k, ', '.join(v) if isinstance(v, list) else str(v), ''])
    for k in OPTION_KEYS:
        rows.append(['옵션', k, 'true' if rules.get(k) else 'false', ''])
    return rows


def apply_map(s: str, mapping: dict) -> str:
    """Replace occurrences of any mapping key in s (longest key first)."""
    if not mapping:
        return s
    keys = sorted(mapping.keys(), key=len, reverse=True)
    pat = '|'.join(re.escape(k) for k in keys)
    return re.sub(pat, lambda m: mapping[m.group()], s)


RULES = load_rules()
