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
    return set(re.findall(r"[A-Z0-9]{3,4}", text.upper()))


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

    result_cols = ['Commission no.', model_col, 'WINGS_codes', 'WINGS_has_pto']
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
