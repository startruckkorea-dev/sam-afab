"""Export OPTION_CODE_MAP to docs/lib/optioncodes.json for the browser pipeline.

The browser-side build (docs/lib/pipeline.js) must use exactly the same code
dictionary the Python parser used (option_codes.OPTION_CODE_MAP) — it drives
PTO detection and the "known code" filter in the SAM parser. Run this whenever
option_codes.py changes:

    python backend/build_option_codes_json.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from option_codes import OPTION_CODE_MAP  # noqa: E402

OUT = ROOT / 'docs' / 'lib' / 'optioncodes.json'


def build() -> Path:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(OPTION_CODE_MAP, ensure_ascii=False, indent=0),
                   encoding='utf-8')
    print(f'[done] wrote {OUT} ({len(OPTION_CODE_MAP)} codes)')
    return OUT


if __name__ == '__main__':
    build()
