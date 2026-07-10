"""Generate code/mandatory-codes.xlsx — an editable list of MANDATORY_CODES.

Layout (so the list can be added to / edited by hand):
    A = 카테고리(Category)   all / tractor / rigid / tipper
    B = 코드(Code)           e.g. D2J
    C = 코드명(Description)   e.g. Seat version, Korea
    D = 비고(Note)           optional remark

Source: backend/option_codes.py -> MANDATORY_CODES.
Run:  python backend/build_mandatory_codes_xlsx.py
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from option_codes import MANDATORY_CODES, MANDATORY_GROUPS  # noqa: E402

OUT = ROOT / 'code' / 'mandatory-codes.xlsx'

NAVY = '1F4E79'
GREY = 'F2F2F2'
CAT_FILL = {
    'all': 'E2EFDA',      # green-ish
    'tractor': 'DDEBF7',  # blue-ish
    'rigid': 'FCE4D6',    # orange-ish
    'tipper': 'FFF2CC',   # yellow-ish
}

HEADER_FONT = Font(color='FFFFFF', bold=True, size=11)
BOLD = Font(bold=True)
THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center')

# Display order for categories.
CAT_ORDER = {'all': 0, 'tractor': 1, 'rigid': 2, 'tipper': 3}


def build():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = 'Mandatory'

    headers = ['카테고리(Category)', '코드(Code)', '코드명(Description)', '비고(Note)']
    ws.append(headers)
    for i in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=i)
        c.fill = PatternFill('solid', fgColor=NAVY)
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER

    # Sort: by category order, then by code.
    items = sorted(
        MANDATORY_CODES.items(),
        key=lambda kv: (CAT_ORDER.get(kv[1][2], 99), kv[0]),
    )

    r = 2
    for code, (desc, note, cat) in items:
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=code).font = BOLD
        ws.cell(row=r, column=3, value=desc)
        ws.cell(row=r, column=4, value=note)
        fill = CAT_FILL.get(cat)
        for col in range(1, 5):
            cell = ws.cell(row=r, column=col)
            cell.alignment = WRAP_TOP
            cell.border = BORDER
            if fill and col == 1:
                cell.fill = PatternFill('solid', fgColor=fill)
        r += 1

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 56
    ws.column_dimensions['D'].width = 44
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:D{r - 1}'

    # --- Sheet 2: category / group legend ------------------------------------
    lg = wb.create_sheet('설명(Legend)')
    lg.column_dimensions['A'].width = 16
    lg.column_dimensions['B'].width = 70
    lg.append(['항목', '설명'])
    for i in (1, 2):
        c = lg.cell(row=1, column=i)
        c.fill = PatternFill('solid', fgColor=NAVY)
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER
    legend = [
        ('all', '모든 차량 필수 (All vehicle mandatory)'),
        ('tractor', '트랙터 전용 필수 (BM 963425, 964416, 963403, 964424)'),
        ('rigid', '리지드 전용 필수 (BM 964XXX)'),
        ('tipper', '티퍼 전용 필수 (BM 964230, 964214)'),
        ('', ''),
        ('그룹(Group)', '그룹 내 코드 중 하나만 있으면 충족 (아래):'),
    ]
    for k, v in legend:
        lg.append([k, v])
    for code_set_name, members in MANDATORY_GROUPS.items():
        lg.append([code_set_name, ', '.join(sorted(members))])
    for row in lg.iter_rows(min_row=2, max_row=lg.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.alignment = WRAP_TOP
            cell.border = BORDER
    lg.cell(row=2, column=1).font = BOLD

    wb.save(OUT)
    print(f'Wrote {OUT}  ({len(items)} codes, {len(MANDATORY_GROUPS)} groups)')


if __name__ == '__main__':
    build()
