"""SAM .docx parsing — extracted from streamlit_app.py, Streamlit-free.

Parses SAM internal-quotation .docx files into a per-model mapping:
    { normalized_model: { is_pto(bool): {'codes': set, 'file': str} } }
"""
from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from option_codes import OPTION_CODE_MAP
from mandatory_codes import load_mandatory
from rules import RULES

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# Universe of real SA codes: the option-code database plus hand-maintained
# mandatory codes (a couple, e.g. S1P/S1W, live only in the mandatory list).
# Used to reject free-text tokens that merely LOOK like codes.
KNOWN_CODES = set(OPTION_CODE_MAP) | set(load_mandatory()['set'])


def normalize_model(model: str) -> str:
    """Normalize a model string to a bare match key: strip axle (8x4), DNA, and
    any non-alphanumerics; uppercase. e.g. '2851 LS' / '2851LS 6x2' -> '2851LS'.

    Deliberately does NO historic/generation remapping. Under the current design
    each SAM file carries both its numbers (filename = 'now', body 'Vehicle type'
    = 'Baumuster'), so matching is done against the real numbers, not a guess map.
    """
    if model is None:
        return ''
    if not isinstance(model, str):
        # WINGS 'Model' with no letter suffix (e.g. 3351) is read as a number by Excel.
        try:
            f = float(model)
            if f != f:            # NaN
                return ''
            model = str(int(f)) if f.is_integer() else str(model)
        except (TypeError, ValueError):
            model = str(model)
    tmp = re.sub(r'\d[Xx]\d', '', model)  # remove axle info like 8x4, 6x2
    return re.sub(r'[^A-Z0-9]', '', tmp.upper().replace('DNA', '').strip())


def parse_single_sam_file(file_obj, name: str, mapping: dict, log_fn=None):
    """Parse one SAM file (file-like object) and update mapping in place."""
    model_raw = None
    codes = set()
    paint = set()
    tyre = set()
    full_text = ''

    try:
        if name.lower().endswith('.docx'):
            with zipfile.ZipFile(file_obj) as z:
                xml_content = z.read('word/document.xml')
            root = ET.fromstring(xml_content)

            full_text = "".join(t.text for t in root.iter(f'{W}t') if t.text)

            # Parse equipment codes from the equipment table cell-by-cell.
            codes = set()
            eq_overview_table = None
            fallback_table = None
            for table in root.iter(f'{W}tbl'):
                tbl_text = "".join(t.text or '' for t in table.iter(f'{W}t'))
                if 'Equipment overview' in tbl_text and eq_overview_table is None:
                    eq_overview_table = table
                if 'Standard equipment' in tbl_text and fallback_table is None:
                    fallback_table = table
            target_table = eq_overview_table or fallback_table

            if target_table is not None:
                section = None
                for para in target_table.iter(f'{W}p'):
                    para_text = "".join(t.text or '' for t in para.iter(f'{W}t')).strip()
                    para_upper = para_text.upper()
                    if para_upper in ('STANDARD EQUIPMENT', 'SPECIAL EQUIPMENT',
                                      'ADDITIONAL EQUIPMENT', 'EQUIPMENT OVERVIEW'):
                        section = para_upper
                        continue
                    if not para_text or section is None:
                        continue
                    if section in ('STANDARD EQUIPMENT', 'SPECIAL EQUIPMENT'):
                        codes |= set(re.findall(r'\b([A-Z][A-Z0-9]{2,3})\b', para_upper))
                    elif section == 'ADDITIONAL EQUIPMENT':
                        m = re.match(r'^([A-Z][A-Z0-9]{2,3})\b', para_upper)
                        if m:
                            codes.add(m.group(1))

            # Paint + Tyre (CTT) codes, from the document's 'Paint' / 'Tyres' sections
            # (separate from the equipment table). Compared against the WINGS 'Paint
            # zone' / 'Tyre key' columns:
            #   Paint  'MB 9676Ice silver...'          -> '9676'  (4-digit color code)
            #   Tyres  '1st Axle:2 x 315/80 R 22,5F18L96 81...' -> 'F18L96' (6-char key)
            _EQ_HEADERS = {'STANDARD EQUIPMENT', 'SPECIAL EQUIPMENT',
                           'ADDITIONAL EQUIPMENT', 'EQUIPMENT OVERVIEW'}
            _sec = None
            for para in root.iter(f'{W}p'):
                ptxt = "".join(t.text or '' for t in para.iter(f'{W}t')).strip()
                if not ptxt:
                    continue
                if ptxt in ('Paint', 'Paints'):
                    _sec = 'PAINT'
                    continue
                if ptxt in ('Tyres', 'Tyre'):
                    _sec = 'TYRES'
                    continue
                if ptxt.upper() in _EQ_HEADERS:
                    _sec = None
                    continue
                if _sec == 'PAINT':
                    for _m in re.finditer(r'MB\s*(\d{4})', ptxt):
                        paint.add(_m.group(1))
                elif _sec == 'TYRES':
                    # '...R 22,5F18L96 81738.00' -> code 'F18L96' + 2-digit mfr/load
                    # index '81' -> 'F18L96 81'. The index is glued to the price, so
                    # take exactly two digits after the space (price follows).
                    _m = re.search(r'R\s*\d{2,3},\d\s*([A-Z0-9]{6})\s+(\d{2})(?=[\d.,])', ptxt)
                    if _m:
                        tyre.add(_m.group(1) + ' ' + _m.group(2))
                    else:
                        _m2 = re.search(r'R\s*\d{2,3},\d\s*([A-Z0-9]{6})', ptxt)
                        if _m2:
                            tyre.add(_m2.group(1))

            # Match the Vehicle type as written, KEEPING a trailing 'DNA' suffix so
            # the displayed SAM Baumuster reads e.g. '2663 LSDNA' (not '2663 LS').
            # Matching keys still drop DNA later via normalize_model(), so keeping it
            # here only affects display. The optional '(?:DNA)?' lets the {1,3}-letter
            # group match the base suffix and then swallow the DNA tail.
            full_text_upper = full_text.upper()
            for pattern in [
                r'VEHICLE\s*TYPE[:\s]+([0-9]{4}\s*[A-Z]{1,3}(?:DNA)?)(?=DRIVETRAIN|SUBCATEGORY|BAUMUSTER|\s|[0-9]|$)',
                r'\bTYPE[:\s]+([0-9]{4}\s*[A-Z]{1,3}(?:DNA)?)(?=DRIVETRAIN|SUBCATEGORY|BAUMUSTER|\s|[0-9]|$)',
                r'\bMODEL[:\s]+([0-9]{4}\s*[A-Z]{1,3}(?:DNA)?)(?=DRIVETRAIN|SUBCATEGORY|BAUMUSTER|\s|[0-9]|$)',
            ]:
                m = re.search(pattern, full_text_upper)
                if m:
                    model_raw = m.group(1).strip()
                    break
        else:
            try:
                raw = file_obj.read() if hasattr(file_obj, 'read') else file_obj.getvalue()
                text = raw.decode('utf-8') if isinstance(raw, bytes) else str(raw)
            except Exception:
                text = ''
            raw_codes = re.findall(r'\b[A-Z0-9]{3,4}\b', text.upper())
            codes = set(c for c in raw_codes if any(ch.isdigit() for ch in c))

        # Reject prose words captured by the loose equipment-code regex (e.g. the
        # section label 'CTT' = Customer Tailored Truck). Codes with a digit are
        # kept as-is; an all-letter token is kept only when it's a real SA code.
        codes = {c for c in codes
                 if any(ch.isdigit() for ch in c) or c in KNOWN_CODES}
    except Exception as e:
        if log_fn:
            log_fn(f'SAM file read error ({name}): {str(e)[:80]}')
        return

    # Two designations per SAM file, both kept:
    #   * body 'Vehicle type'  = SAM Baumuster (original/구형 번호)  -> already in model_raw
    #   * filename model        = SAM now (corrected/현행 번호)
    body_model = (model_raw or '').strip()

    fname_upper = name.upper()
    fname_model = None
    m_fname = re.search(r'(\d{4}\s*[A-Z]{0,3})(?=\s+[A-Z]\d[A-Z]|\s+\d[Xx]\d|\s+HUB|\s+CLASSIC|\s+EURO|\s|$)', fname_upper)
    if not m_fname:
        m_fname = re.search(r'(\d{4}\s*[A-Z]{1,3})', fname_upper)
    if m_fname:
        fname_model = m_fname.group(1).strip()

    model_now = fname_model or body_model            # 수정(현행) — 파일제목 우선
    model_baumuster = body_model or fname_model      # 원본(구형) — 본문 Vehicle type

    # Baumuster (factory code) + Subcategory — from the word body only (not the title).
    body_bm = body_sub = ''
    if full_text:
        m = re.search(r'Baumuster[:\s]*([0-9]{5,})', full_text, re.IGNORECASE)
        if m:
            body_bm = m.group(1)
        m = re.search(r'Subcategory[:\s]*([0-9]{1,3}[A-Za-z])', full_text, re.IGNORECASE)
        if m:
            body_sub = m.group(1)

    if (model_now or model_baumuster) and codes:
        is_pto = any('PTO' in OPTION_CODE_MAP.get(c, '').upper() for c in codes)
        if not is_pto:
            _doc_text = full_text if name.lower().endswith('.docx') else ''
            if _doc_text and re.search(r'\bPTO\b', _doc_text, re.IGNORECASE):
                is_pto = True
        if not is_pto and re.search(r'\bPTO\b', name, re.IGNORECASE):
            is_pto = True

        data = {'codes': codes, 'paint': paint, 'tyre': tyre, 'file': name,
                'model_now': model_now, 'model_baumuster': model_baumuster,
                'bm': body_bm, 'sub': body_sub}
        # Index the same entry under BOTH numbers so a WINGS row matches on either.
        # Value is a LIST of candidates: two files can share a model key (e.g. one via
        # its 'now' number, another via its 'baumuster' number) — Baumuster/Subcategory
        # then disambiguate in compare.
        keys = {k for k in (normalize_model(model_now), normalize_model(model_baumuster)) if k}
        for key in keys:
            mapping.setdefault(key, {}).setdefault(is_pto, []).append(data)
        if log_fn:
            log_fn(f"OK '{name}' -> now={model_now!r} baumuster={model_baumuster!r} "
                   f"bm={body_bm} sub={body_sub} keys={sorted(keys)} "
                   f"({'PTO' if is_pto else 'non-PTO'}, {len(codes)} codes)")


def load_sam_from_folder(folder: Path, log_fn=None) -> dict:
    """Load all SAM .docx/.csv/.txt files from a folder; return per-model mapping."""
    mapping = {}
    folder = Path(folder)
    valid_exts = {'.docx', '.csv', '.txt'}
    sam_files = [f for f in sorted(folder.glob('*'))
                 if f.suffix.lower() in valid_exts and not f.name.startswith('.')]
    for fpath in sam_files:
        with open(fpath, 'rb') as fobj:
            parse_single_sam_file(fobj, fpath.name, mapping, log_fn=log_fn)

    # Auto-generate aliases so WINGS (newer numbering) can find SAM (older numbering).
    _reverse_prefixes = RULES.get('reverse_aliases', {})
    for key in list(mapping.keys()):
        m = re.match(r'^(\d+)([A-Z]*)$', key)
        if not m:
            continue
        num, suffix = m.group(1), m.group(2)
        for src_prefix, alias_prefixes in _reverse_prefixes.items():
            if num == src_prefix:
                for ap in alias_prefixes:
                    alias_key = ap + suffix
                    if alias_key not in mapping:
                        mapping[alias_key] = mapping[key]
    return mapping


def load_sam_by_month(root: Path, log_fn=None) -> dict:
    """Scan root/<YYYY_MM>/ subfolders; return { yyyymm(int): mapping }."""
    root = Path(root)
    out = {}
    for sub in sorted(root.glob('*')):
        if not sub.is_dir():
            continue
        # Accept 'YYYY_MM', 'YYYY-MM', and trailing text like 'YYYY-MM 생산'.
        m = re.match(r'^(\d{4})[_-](\d{2})\b', sub.name)
        if not m:
            continue
        yyyymm = int(m.group(1)) * 100 + int(m.group(2))
        mapping = load_sam_from_folder(sub, log_fn=log_fn)
        if mapping:
            out[yyyymm] = mapping
    return out
