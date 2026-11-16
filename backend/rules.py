"""Model-name rules — single source of truth, model_rules/model_mapping.xlsx.

There is one kind of rule left, 모델명 변경 (model rename), on the single sheet
'모델명변경':  From | To | 메모   (docs/lib/refdata.js reads the same sheet).

  A WINGS model name (axle and DNA removed, e.g. '2651 LS') equal to From becomes To —
  one way, chained when a To is itself another From. The renamed name is both the Type
  shown on the dashboard and the number used to find the SAM document.

The older rule kinds (대체번호 / SAM별칭 / 옵션 normalize_28xx_to_26xx / 차종 /
이전모델·현재모델) were dropped. From an older workbook only the display-replace rows
are carried over (the '표시치환' rows of a '매칭규칙' sheet, or the 'WINGS표시치환' sheet).
docs/rules.json, when present, is still a lower-priority fallback.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
RULES_PATH = ROOT / 'docs' / 'rules.json'
XLSX_PATH = ROOT / 'model_rules' / 'model_mapping.xlsx'

RULE_SHEET = '모델명변경'
RULE_HEADER = ['From (WINGS 모델명)', 'To (바꿀 모델명)', '메모']
_LEGACY_KINDS = ('표시치환', 'wings_display_replace', '모델명변경')

# Built-in defaults (used only when the workbook has no rename sheet at all).
DEFAULTS = {
    'model_rename': {
        '4140': '4440', '2651 LS': '2851 LS', '2653 LS': '2853 LS',
        '2663 LS': '2863 LS', '2643 A': '3343 A',
    },
}


def load_rules_from_xlsx(path: Path | None = None) -> dict:
    """Parse the rename rows of model_rules/model_mapping.xlsx.

    Returns {'model_rename': {...}} when the workbook has a rename sheet (even an empty
    one — an emptied sheet means "no renames"), else {}. openpyxl is imported lazily and
    every failure is swallowed so the build never breaks on a bad file.
    """
    p = path or XLSX_PATH
    try:
        from openpyxl import load_workbook  # lazy: optional at build time
        wb = load_workbook(Path(p), read_only=True, data_only=True)
    except Exception:
        return {}

    def _cell(row, i):
        return '' if len(row) <= i or row[i] is None else str(row[i]).strip()

    def _rows(name):
        for i, row in enumerate(wb[name].iter_rows(values_only=True)):
            if i == 0 or not row:
                continue
            yield [_cell(row, j) for j in range(max(3, len(row)))]

    found = None
    if RULE_SHEET in wb.sheetnames:
        found = {r[0]: r[1] for r in _rows(RULE_SHEET) if r[0] and r[1]}
    elif '매칭규칙' in wb.sheetnames:
        found = {}
        for kind, frm, to, *_ in _rows('매칭규칙'):
            if re.sub(r'\s+', '', kind).lower() in _LEGACY_KINDS and frm and to:
                found[frm] = to
    elif 'WINGS표시치환' in wb.sheetnames:
        found = {r[0]: r[1] for r in _rows('WINGS표시치환') if r[0] and r[1]}
    wb.close()
    return {} if found is None else {'model_rename': found}


def load_rules(path: Path | None = None, xlsx_path: Path | None = None) -> dict:
    """Priority (highest wins): model_mapping.xlsx > docs/rules.json > built-in DEFAULTS."""
    p = path or RULES_PATH
    merged = {k: dict(v) for k, v in DEFAULTS.items()}
    try:
        data = json.loads(Path(p).read_text(encoding='utf-8'))
        if isinstance(data.get('model_rename'), dict):
            merged['model_rename'] = data['model_rename']
    except Exception:
        pass
    merged.update(load_rules_from_xlsx(xlsx_path))
    return merged


def rules_to_rows(rules: dict, notes: dict | None = None) -> list:
    """rules dict -> '모델명변경' sheet rows (header first). Same as refdata.rulesToAoa."""
    notes = notes or {}
    rows = [RULE_HEADER[:]]
    for k, v in (rules.get('model_rename') or {}).items():
        rows.append([k, str(v), notes.get(k, '')])
    return rows


def apply_map(s: str, mapping: dict) -> str:
    """Replace occurrences of any mapping key in s (longest key first)."""
    if not mapping:
        return s
    keys = sorted(mapping.keys(), key=len, reverse=True)
    pat = '|'.join(re.escape(k) for k in keys)
    return re.sub(pat, lambda m: mapping[m.group()], s)


RULES = load_rules()
