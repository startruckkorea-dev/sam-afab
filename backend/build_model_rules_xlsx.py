"""Generate model_rules/model_mapping.xlsx — the WINGS↔SAM model-recognition table.

Since matching is now data-driven (each SAM word file carries both its numbers —
filename = 'SAM now', body 'Vehicle type' = 'SAM Baumuster'), this workbook is mainly
a VERIFICATION view of what got recognized, plus one small hand-editable override sheet:

  * 인식모델_대조표  — auto-generated from docs/data.json: which WINGS model matched
    which SAM file, showing both SAM numbers and the match status. (read-only view)
  * 매칭_별칭(수동)  — optional manual overrides: force a WINGS/SAM number to also
    match an extra number (for edge cases where a file uses an off-generation number).
    Read back at build time as rules 'reverse_aliases'.

Run:  python backend/build_model_rules_xlsx.py   (or it is refreshed by build_data.py)
"""
from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from rules import load_rules

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
DATA_PATH = ROOT / 'docs' / 'data.json'
OUT_DIR = ROOT / 'model_rules'
OUT_PATH = OUT_DIR / 'model_mapping.xlsx'

HEADER_FILL = PatternFill('solid', fgColor='1F4E79')
HEADER_FONT = Font(color='FFFFFF', bold=True)
ALIAS_FILL = PatternFill('solid', fgColor='7A5195')


def _style_header(ws, ncols, fill=HEADER_FILL):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = fill
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.freeze_panes = 'A2'


def _autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _ref_rows():
    """Distinct recognized (WINGS, SAM Baumuster, SAM now, status) from docs/data.json."""
    if not DATA_PATH.exists():
        return []
    data = json.loads(DATA_PATH.read_text(encoding='utf-8'))
    seen = {}
    for r in data.get('rows', []):
        wings = str(r.get('Model(WINGS)', '') or '').strip()
        if not wings:
            continue
        row = (
            str(r.get('Vehicle', '') or '').strip(),
            wings,
            str(r.get('SAM Baumuster', '') or '').strip(),
            str(r.get('SAM now', '') or '').strip(),
            str(r.get('Baumuster', '') or '').strip(),
            str(r.get('Subcategory (ID)', '') or '').strip(),
            str(r.get('SAM Status', '') or '').strip(),
            str(r.get('PTO', '') or '').strip(),
            Path(str(r.get('Compared SAM file name', '') or '')).name,
        )
        seen.setdefault(row, row)
    return sorted(seen.values(), key=lambda k: (k[0], k[1], k[8]))


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    rules = load_rules()
    wb = Workbook()

    # --- Sheet 1: recognition table (auto) -----------------------------------
    ws = wb.active
    ws.title = '인식모델_대조표'
    cols = ['차종(Vehicle)', 'WINGS 모델', 'SAM Baumuster(원본)', 'SAM now(수정)',
            'Baumuster', 'Subcategory', '매칭상태', 'PTO', 'SAM 파일']
    ws.append(cols)
    for row in _ref_rows():
        ws.append(list(row))
    _style_header(ws, len(cols))
    _autofit(ws, [16, 16, 18, 16, 12, 11, 12, 8, 52])

    # --- Sheet 2: manual overrides (editable, wins over auto-recognition) -----
    sm = wb.create_sheet('수동매핑')
    sm.append(['WINGS 모델', 'SAM Baumuster(원본)', 'SAM now(수정)', 'SAM 파일 지정(선택)'])
    for wings, v in (rules.get('manual_map') or {}).items():
        v = v or {}
        sm.append([wings, v.get('baumuster', ''), v.get('now', ''), v.get('file', '')])
    _style_header(sm, 4, fill=ALIAS_FILL)
    _autofit(sm, [16, 18, 16, 52])
    sm.cell(row=1, column=6, value='※ 자동인식이 틀릴 때만 사용. 값을 적으면 그 WINGS 모델의 표시를 강제. '
            'SAM 파일 지정(제목 일부)까지 적으면 그 파일과 강제 비교 → 잘못된 파일 매칭 교정.').font = Font(italic=True, color='777777')

    # --- Sheet 3: manual match aliases (editable) ----------------------------
    sa = wb.create_sheet('매칭_별칭(수동)')
    sa.append(['SAM 번호', '추가로 매칭할 번호(쉼표 구분)'])
    for k, v in (rules.get('reverse_aliases') or {}).items():
        sa.append([k, ', '.join(v) if isinstance(v, list) else str(v)])
    _style_header(sa, 2, fill=ALIAS_FILL)
    _autofit(sa, [20, 34])
    sa.cell(row=1, column=4, value='※ 대부분 비워두면 됩니다. 파일 제목이 다른 세대 번호를 쓸 때만 여기에 별칭 추가.').font = Font(italic=True, color='777777')

    wb.save(OUT_PATH)
    print(f'Wrote {OUT_PATH}  (rows: {ws.max_row - 1})')


if __name__ == '__main__':
    build()
