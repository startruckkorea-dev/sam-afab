"""Orchestrator: parse SAM + WINGS, compare, and write docs/data.json.

This is the "compute" half of the architecture. It runs in GitHub Actions
(or locally), and the only thing GitHub Pages needs is the JSON it produces.

Pipeline:
    1. (optional) scrape WINGS via wings_scraper.download_wings_excel
    2. parse the WINGS export (CSV/Excel)
    3. parse SAM .docx files grouped by month folder (sam_files/YYYY_MM/)
    4. compare() -> result rows
    5. write docs/data.json  ->  { generated_at, rows: [...], summary: {...} }

Usage:
    python build_data.py --wings path/to/wings.xlsx
    python build_data.py --scrape 2026_04 2026_05      # fetch WINGS first
"""
from __future__ import annotations

import argparse
import json
import math
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from wings_parser import parse_wings          # noqa: E402
from sam_parser import load_sam_by_month       # noqa: E402
from compare import compare                    # noqa: E402

CONFIG_PATH = ROOT / 'config.json'
OUT_JSON = ROOT / 'docs' / 'data.json'


def _load_config() -> dict:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding='utf-8'))
    except Exception:
        return {}


def _resolve_dir(cli_val, env_var, cfg_key, default) -> Path:
    """Pick a directory by priority: CLI > env > config.json > default.

    Relative paths are resolved against the repo root, so the same config works
    whether SAM/WINGS live inside the repo or in an external (e.g. SharePoint
    OneDrive-synced) folder given as an absolute path.
    """
    cfg = _load_config()
    val = cli_val or os.environ.get(env_var) or cfg.get(cfg_key) or default
    p = Path(val)
    return p if p.is_absolute() else (ROOT / p)


SAM_ROOT = ROOT / 'sam_files'
WINGS_DIR = ROOT / 'wings_data'

# Columns surfaced to the front-end, in display order.
DISPLAY_COLS = [
    'Commission no.', 'Baumuster', 'Model(WINGS)', 'Vehicle', 'Category', 'Type', 'Cab', 'PTO',
    'SAM Baumuster', 'SAM now', 'Changeability Date', 'Until Dealine',
    'Production date', 'Only_in_SAM', 'Only_in_WINGS', 'Factory Control Codes',
    'Mandatory Codes', 'Order status financial', 'Order status logistical',
    'FIN', 'Subcategory (ID)', 'Compared SAM file name', 'SAM Status',
    '_all_wings_codes', '_all_sam_codes',
]


def _clean(v):
    """JSON-safe scalar."""
    if v is None:
        return ''
    if isinstance(v, float) and math.isnan(v):
        return ''
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime('%Y-%m-%d')
    return v


def _build_code_dict() -> dict:
    """Sales-code dictionary from the spec workbook: code_dict sheet, cols B & C.

    B = code, C = name_en (English description). Path/sheet come from config.json
    (code_dict_file / code_dict_sheet). Falls back to option_codes.OPTION_CODE_MAP
    if the workbook is unavailable, so the build never breaks.
    """
    cfg = _load_config()
    rel = os.environ.get('CODE_DICT_FILE') or cfg.get('code_dict_file') or 'code/mbtruck-spec-data.xlsx'
    sheet = os.environ.get('CODE_DICT_SHEET') or cfg.get('code_dict_sheet') or 'code_dict'
    path = Path(rel)
    if not path.is_absolute():
        path = ROOT / path
    try:
        df = pd.read_excel(path, sheet_name=sheet)
        code_col, desc_col = df.columns[1], df.columns[2]  # B, C
        out = {}
        for _, r in df.iterrows():
            code = str(r[code_col]).strip()
            if not code or code.lower() == 'nan':
                continue
            desc = r[desc_col]
            out[code] = '' if (desc is None or (isinstance(desc, float) and math.isnan(desc))) else str(desc).strip()
        if out:
            print(f'[codes] loaded {len(out)} from {path} [{sheet}] B,C')
            return out
        print(f'[codes] WARN: {path} [{sheet}] empty; falling back to option_codes')
    except Exception as e:
        print(f'[codes] WARN: cannot read {path} [{sheet}] ({str(e)[:80]}); falling back to option_codes')
    from option_codes import OPTION_CODE_MAP
    return dict(OPTION_CODE_MAP)


def _wings_recency(f: Path) -> tuple:
    """Sort key for auto-picking the newest WINGS export among several.

    Prefers a timestamp embedded in the filename (more reliable than mtime, which
    changes on copy/download): a trailing 13-digit epoch-ms (e.g. ..._1783469227544)
    or a leading date (YYYY-MM-DD / YYYYMMDD). Falls back to file mtime.
    Returns (best_epoch_seconds, mtime) so ties break on mtime.
    """
    import re
    name = f.name
    scores = []
    m = re.search(r'(\d{13})', name)           # epoch milliseconds
    if m:
        scores.append(int(m.group(1)) / 1000.0)
    m = re.search(r'(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})', name)  # YYYY-MM-DD
    if m:
        try:
            from datetime import datetime, timezone
            y, mo, d = (int(x) for x in m.groups())
            scores.append(datetime(y, mo, d, tzinfo=timezone.utc).timestamp())
        except Exception:
            pass
    try:
        mtime = f.stat().st_mtime
    except OSError:
        mtime = 0
    return (max(scores) if scores else mtime, mtime)


def _find_latest_wings(wings_dir: Path) -> Path | None:
    if not wings_dir.exists():
        return None
    files = [f for f in wings_dir.glob('*')
             if f.suffix.lower() in {'.xlsx', '.xls', '.csv'}
             and not f.name.startswith('.') and not f.name.startswith('~$')]
    if not files:
        return None
    chosen = max(files, key=_wings_recency)
    if len(files) > 1:
        print(f'[wings] {len(files)} files in {wings_dir.name}/ -> auto-picked '
              f'newest: {chosen.name}')
    return chosen


def build(wings_path: Path, sam_root: Path) -> dict:
    print(f'[build] WINGS file : {wings_path}')
    print(f'[build] SAM folder : {sam_root}')
    df_wings = parse_wings(str(wings_path))
    print(f'[build] WINGS rows : {len(df_wings)}')

    sam_maps = load_sam_by_month(sam_root, log_fn=lambda m: print('  [sam]', m))
    print(f'[build] SAM months : {sorted(sam_maps.keys())}')

    result = compare(df_wings, sam_maps)
    print(f'[build] result rows: {len(result)}')

    cols = [c for c in DISPLAY_COLS if c in result.columns]
    rows = [{c: _clean(r[c]) for c in cols} for _, r in result.iterrows()]

    matched = sum(1 for r in rows if r.get('SAM Status') == 'Match')
    mismatched = sum(1 for r in rows if r.get('SAM Status') == 'Mismatch')
    no_sam = sum(1 for r in rows if r.get('SAM Status') == 'No SAM')
    summary = {
        'total': len(rows),
        'matched': matched,
        'mismatched': mismatched,
        'no_sam': no_sam,
        'sam_months': sorted(sam_maps.keys()),
    }
    return {
        'generated_at': datetime.now(timezone.utc).isoformat(timespec='seconds'),
        'wings_file': wings_path.name,
        'columns': cols,
        'summary': summary,
        'rows': rows,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--wings', help='Path to a WINGS CSV/Excel export.')
    ap.add_argument('--sam-dir', help='SAM source folder (overrides config.json/SAM_DIR).')
    ap.add_argument('--wings-dir', help='WINGS folder to auto-pick latest from (overrides config.json/WINGS_DIR).')
    ap.add_argument('--scrape', nargs='*', metavar='YYYY_MM',
                    help='Scrape WINGS for the given months before building.')
    ap.add_argument('--out', default=str(OUT_JSON), help='Output JSON path.')
    args = ap.parse_args()

    sam_root = _resolve_dir(args.sam_dir, 'SAM_DIR', 'sam_dir', 'sam_files')
    wings_dir = _resolve_dir(args.wings_dir, 'WINGS_DIR', 'wings_dir', 'wings_data')

    wings_path = None
    if args.scrape:
        from wings_scraper import download_wings_excel
        months = [m.replace('_', '-') for m in args.scrape]
        wings_dir.mkdir(parents=True, exist_ok=True)
        print(f'[scrape] months: {months}')
        wings_path = Path(download_wings_excel(months, download_dir=str(wings_dir),
                                               on_status=lambda m: print('  [wings]', m)))
    elif args.wings:
        wings_path = Path(args.wings)
    else:
        wings_path = _find_latest_wings(wings_dir)

    if not wings_path or not Path(wings_path).exists():
        print('[error] No WINGS file found. Use --wings <path> or --scrape <months>.',
              file=sys.stderr)
        sys.exit(1)

    data = build(Path(wings_path), sam_root)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'[done] wrote {out}  ({data["summary"]})')

    # Emit the code dictionary so the front-end can show descriptions on click.
    # Source: code_dict sheet of the spec workbook, columns B (code) + C (name_en).
    options = _build_code_dict()
    from mandatory_codes import load_mandatory
    _mand = load_mandatory()
    codes_out = out.parent / 'codes.json'
    codes_out.write_text(json.dumps({
        'options': options,
        'mandatory': _mand['desc'],
        'mandatory_groups': {g: sorted(m) for g, m in _mand['groups'].items()},
    }, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"[done] wrote {codes_out}  ({len(options)} codes, "
          f"{len(_mand['set'])} mandatory from {_mand['source']})")

    # Refresh the model-recognition workbook (model_rules/model_mapping.xlsx) so it
    # always reflects the latest build. Never let this break the build.
    try:
        import build_model_rules_xlsx
        build_model_rules_xlsx.build()
    except Exception as e:
        print(f'[warn] model_mapping.xlsx refresh skipped: {str(e)[:100]}')

    # Refresh the Baumuster->category workbook so newly-seen BM prefixes appear
    # (existing category edits are preserved). Never let this break the build.
    try:
        import build_model_category_xlsx
        build_model_category_xlsx.build()
    except Exception as e:
        print(f'[warn] model-category.xlsx refresh skipped: {str(e)[:100]}')


if __name__ == '__main__':
    main()
