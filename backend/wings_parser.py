"""WINGS export (CSV/Excel) parsing — extracted from streamlit_app.py, Streamlit-free.

Produces a normalized DataFrame with one row per Commission no., including the
set of equipment codes (WINGS_codes) and a PTO flag (WINGS_has_pto).
"""
from __future__ import annotations

import re
import pandas as pd


def _extract_codes(text):
    if pd.isna(text):
        return set()
    text = str(text)
    text = re.sub(r'\bnan\b', '', text, flags=re.IGNORECASE)
    # 3-5 chars (word-bounded) so a 4-5 digit CTT code in 'Additional equipment'
    # survives intact — the old unbounded {3,4} silently truncated a 5-char run to
    # its first 4 chars. Current option codes are 3-4 chars, so this is a no-op for
    # them and only rescues longer CTT codes when they appear.
    return set(re.findall(r"\b[A-Z0-9]{3,5}\b", text.upper()))


def _norm_paint(v):
    """WINGS 'Paint zone N color code' cell -> canonical color code string.

    Values arrive as 4-digit MB color numbers, sometimes as Excel/CSV numerics
    (e.g. 9676.0). Returns '9676' to match the SAM 'MB 9676' paint code.
    """
    if pd.isna(v):
        return ''
    s = str(v).strip()
    if not s or s.lower() == 'nan':
        return ''
    s = re.sub(r'\.0+$', '', s)
    m = re.search(r'\d{3,5}', s)
    return m.group(0) if m else ''


def _norm_tyre(v):
    """WINGS 'Tyre key N. axle' cell -> canonical 6-char tyre code (e.g. 'F18L96')."""
    if pd.isna(v):
        return ''
    s = re.sub(r'[^A-Za-z0-9]', '', str(v)).upper()
    return s if (4 <= len(s) <= 8 and s != 'NAN') else ''


def _norm_mfr(v):
    """WINGS 'Tyre manufacturer key N. axle' cell -> 2-digit key (e.g. 80.0 -> '80')."""
    if pd.isna(v):
        return ''
    s = re.sub(r'\.0+$', '', str(v).strip())
    return s if re.fullmatch(r'\d{1,3}', s) else ''


def _axle_no(colname):
    """Axle number from a column name, e.g. 'Tyre key 2. axle' -> '2'."""
    m = re.search(r'(\d+)\s*\.?\s*axle', colname.lower())
    return m.group(1) if m else colname


def _collect(row, cols, fn):
    out = set()
    for c in cols:
        v = fn(row[c])
        if v:
            out.add(v)
    return out


def parse_wings(file) -> pd.DataFrame:
    """Parse a WINGS CSV/Excel file (path or file-like) into a normalized DataFrame."""
    try:
        df = pd.read_csv(file, encoding='utf-8')
    except Exception:
        if hasattr(file, 'seek'):
            file.seek(0)
        df = pd.read_excel(file)

    df.rename(columns={c: c.strip() for c in df.columns}, inplace=True)

    if 'Commission no.' not in df.columns:
        raise ValueError('Cannot find `Commission no.` column in the WINGS file.')

    # Model name column: prefer 'Type (brief)', then 'Type', then 'Baumuster'.
    model_col = None
    for col_name in df.columns:
        if 'type' in col_name.lower() and 'brief' in col_name.lower():
            model_col = col_name
            break
    if not model_col:
        for col_name in df.columns:
            if col_name.lower() == 'type':
                model_col = col_name
                break
    if not model_col:
        model_col = 'Baumuster' if 'Baumuster' in df.columns else None
    if model_col is None:
        model_col = df.columns[1] if len(df.columns) > 1 else 'Commission no.'

    # Option code columns.
    wings_opt_col1 = wings_opt_col2 = None
    for col_name in df.columns:
        low = col_name.lower()
        if 'standard' in low and 'equipment' in low:
            wings_opt_col1 = col_name
        elif 'additional' in low and 'equipment' in low:
            wings_opt_col2 = col_name

    if not wings_opt_col1 and not wings_opt_col2:
        try:
            if df.shape[1] >= 11:
                wings_opt_col1 = df.columns[8]
                wings_opt_col2 = df.columns[10]
        except Exception:
            pass

    if not wings_opt_col1 or not wings_opt_col2:
        for name in df.columns:
            low = name.lower()
            if 'equipment' in low or 'offer code' in low or 'enumeration' in low:
                if wings_opt_col1 is None:
                    wings_opt_col1 = name
                elif wings_opt_col2 is None and name != wings_opt_col1:
                    wings_opt_col2 = name
                    break

    code_cols = [c for c in (wings_opt_col1, wings_opt_col2) if c]
    if code_cols:
        text_parts = []
        for col in dict.fromkeys(code_cols):
            col_data = df[col]
            if isinstance(col_data, pd.DataFrame):
                col_data = col_data.iloc[:, 0]
            text_parts.append(col_data.astype(str))
        combined = text_parts[0]
        for part in text_parts[1:]:
            combined = combined + ' ' + part
        df['WINGS_codes'] = combined.apply(_extract_codes)
        df['WINGS_has_pto'] = combined.str.contains(r'\bPTO\b', case=False, na=False)
    else:
        _all_text = df.astype(str).agg(' '.join, axis=1)
        df['WINGS_codes'] = _all_text.apply(_extract_codes)
        df['WINGS_has_pto'] = _all_text.str.contains(r'\bPTO\b', case=False, na=False)

    # Paint / Tyre CTT codes live in their own columns (new WINGS report format):
    #   'Paint zone 1..4 color code'  and  'Tyre key 1..4. axle'.
    paint_cols = [c for c in df.columns
                  if 'paint' in c.lower() and 'zone' in c.lower()]
    if paint_cols:
        df['WINGS_paint'] = df.apply(lambda r: _collect(r, paint_cols, _norm_paint), axis=1)
    else:
        df['WINGS_paint'] = [set() for _ in range(len(df))]

    # Tyre code lives in 'Tyre key N. axle'; its 2-digit manufacturer/load index in
    # 'Tyre manufacturer key N. axle'. Pair them PER AXLE -> 'F18L96 81' so the key
    # matches the SAM Tyres line ('...R 22,5F18L96 81...') in full.
    tyre_key_cols = {_axle_no(c): c for c in df.columns
                     if 'tyre key' in c.lower() and 'axle' in c.lower()}
    mfr_cols = {_axle_no(c): c for c in df.columns
                if 'manufacturer key' in c.lower() and 'axle' in c.lower()}

    def _row_tyre(row):
        out = set()
        for ax, kc in tyre_key_cols.items():
            key = _norm_tyre(row[kc])
            if not key:
                continue
            mfr = _norm_mfr(row[mfr_cols[ax]]) if ax in mfr_cols else ''
            out.add(key + ' ' + mfr if mfr else key)
        return out

    if tyre_key_cols:
        df['WINGS_tyre'] = df.apply(_row_tyre, axis=1)
    else:
        df['WINGS_tyre'] = [set() for _ in range(len(df))]

    result_cols = ['Commission no.', model_col, 'WINGS_codes', 'WINGS_has_pto',
                   'WINGS_paint', 'WINGS_tyre']
    if 'Baumuster' in df.columns and model_col != 'Baumuster':
        result_cols.insert(2, 'Baumuster')

    for col in ['Order status financial', 'Order status logistical',
                'Additional equipment (enumeration)', 'FIN', 'Subcategory (ID)',
                'Vehicle alterable until', 'Requested delivery date']:
        if col in df.columns:
            result_cols.append(col)

    result = df[result_cols].copy()
    result.rename(columns={model_col: 'Model'}, inplace=True)
    return result
