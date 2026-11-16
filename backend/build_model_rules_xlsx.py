"""Keep model_rules/model_mapping.xlsx as ONE sheet, '모델명변경' (From | To | 메모).

모델명 변경 is the only rule kind left (see rules.py). An older workbook is rewritten as
the single sheet carrying over its display-replace rows. The browser build (docs/app.js)
additionally adds the renames that the dropped fallback rules used to provide, found in
the current data; this local tool has no such data and only carries rows over.
A workbook that already has only '모델명변경' is left untouched (its 메모 column is kept).

The auto-generated views that used to live here (인식모델_대조표 · 미매칭_진단) are now
shown on the web page (모델 매칭 > 문제 / 매칭 결과) instead of in the workbook.

Run:  python backend/build_model_rules_xlsx.py   (or it is refreshed by build_data.py)
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from rules import RULE_SHEET, load_rules, rules_to_rows

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
OUT_DIR = ROOT / 'model_rules'
OUT_PATH = OUT_DIR / 'model_mapping.xlsx'

HEADER_FILL = PatternFill('solid', fgColor='1F4E79')
HEADER_FONT = Font(color='FFFFFF', bold=True)


def _already_unified(path: Path) -> bool:
    if not path.exists():
        return False
    try:
        wb = load_workbook(path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
    except Exception:
        return False
    return names == [RULE_SHEET]


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if _already_unified(OUT_PATH):
        print(f'{OUT_PATH} already has the single {RULE_SHEET} sheet - unchanged')
        return
    rows = rules_to_rows(load_rules(xlsx_path=OUT_PATH))
    wb = Workbook()
    ws = wb.active
    ws.title = RULE_SHEET
    for row in rows:
        ws.append(row)
    for c in range(1, len(rows[0]) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.freeze_panes = 'A2'
    for i, w in enumerate([24, 24, 50], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(OUT_PATH)
    print(f'Wrote {OUT_PATH}  ({RULE_SHEET}: {len(rows) - 1} rows)')


if __name__ == '__main__':
    build()
