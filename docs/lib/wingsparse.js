// backend/wings_parser.py 의 JS 포팅 — WINGS export(xlsx/csv)를 정규화된 행 배열로.
// xlsx 파싱은 SheetJS(XLSX) 를 사용한다(브라우저: vendor/xlsx.full.min.js).
(function (root) {
  'use strict';

  function getXLSX() {
    if (root.XLSX) return root.XLSX;                       // 브라우저(vendor) 또는 테스트 주입
    if (typeof module !== 'undefined' && module.exports) return require('xlsx');
    throw new Error('SheetJS(XLSX) 가 로드되지 않았습니다.');
  }

  // pandas astype(str) 근사: null/undefined -> '' (뒤에서 \bnan\b 제거와 동일 효과)
  function s(v) {
    if (v === null || v === undefined) return '';
    return String(v);
  }

  function extractCodes(text) {
    let t = s(text);
    if (!t) return new Set();
    t = t.toUpperCase().replace(/\bNAN\b/g, '');
    const out = new Set();
    const re = /\b[A-Z0-9]{3,5}\b/g;
    let m;
    while ((m = re.exec(t)) !== null) out.add(m[0]);
    return out;
  }

  function normPaint(v) {
    let t = s(v).trim();
    if (!t || t.toLowerCase() === 'nan') return '';
    t = t.replace(/\.0+$/, '');
    const m = /\d{3,5}/.exec(t);
    return m ? m[0] : '';
  }

  function normTyre(v) {
    const t = s(v).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    return (t.length >= 4 && t.length <= 8 && t !== 'NAN') ? t : '';
  }

  function normMfr(v) {
    const t = s(v).trim().replace(/\.0+$/, '');
    return /^\d{1,3}$/.test(t) ? t : '';
  }

  function axleNo(colname) {
    const m = /(\d+)\s*\.?\s*axle/.exec(String(colname).toLowerCase());
    return m ? m[1] : colname;
  }

  /**
   * WINGS 파일 파싱.
   * @param {ArrayBuffer|Uint8Array} buf
   * @param {string} filename  (확장자로 csv/xlsx 구분)
   * @returns {Array<object>} 정규화 행들 (Commission no., Model, WINGS_codes:Set, ...)
   */
  function parseWings(buf, filename) {
    const XLSX = getXLSX();
    const isCsv = /\.csv$/i.test(filename || '');
    const wb = isCsv
      ? XLSX.read(new TextDecoder('utf-8').decode(buf instanceof Uint8Array ? buf : new Uint8Array(buf)), { type: 'string' })
      : XLSX.read(buf instanceof Uint8Array ? buf : new Uint8Array(buf), { type: 'array', cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: null });
    if (!aoa.length) return [];

    const cols = (aoa[0] || []).map(function (c) { return s(c).trim(); });
    const rows = [];
    for (let i = 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      const o = {};
      for (let c = 0; c < cols.length; c++) o[cols[c]] = r[c] === undefined ? null : r[c];
      rows.push(o);
    }

    if (cols.indexOf('Commission no.') === -1) {
      throw new Error('WINGS 파일에서 `Commission no.` 열을 찾을 수 없습니다.');
    }

    // 모델명 열: 'Type (brief)' > 'Type' > 'Baumuster' > 두번째 열
    let modelCol = null;
    for (const c of cols) {
      const l = c.toLowerCase();
      if (l.indexOf('type') !== -1 && l.indexOf('brief') !== -1) { modelCol = c; break; }
    }
    if (!modelCol) for (const c of cols) { if (c.toLowerCase() === 'type') { modelCol = c; break; } }
    if (!modelCol) modelCol = cols.indexOf('Baumuster') !== -1 ? 'Baumuster' : null;
    if (modelCol === null) modelCol = cols.length > 1 ? cols[1] : 'Commission no.';

    // 옵션 코드 열
    let c1 = null, c2 = null;
    for (const c of cols) {
      const l = c.toLowerCase();
      if (l.indexOf('standard') !== -1 && l.indexOf('equipment') !== -1) c1 = c;
      else if (l.indexOf('additional') !== -1 && l.indexOf('equipment') !== -1) c2 = c;
    }
    if (!c1 && !c2 && cols.length >= 11) { c1 = cols[8]; c2 = cols[10]; }
    if (!c1 || !c2) {
      for (const c of cols) {
        const l = c.toLowerCase();
        if (l.indexOf('equipment') !== -1 || l.indexOf('offer code') !== -1 || l.indexOf('enumeration') !== -1) {
          if (c1 === null) c1 = c;
          else if (c2 === null && c !== c1) { c2 = c; break; }
        }
      }
    }
    const codeCols = [c1, c2].filter(Boolean).filter(function (v, i, a) { return a.indexOf(v) === i; });

    const paintCols = cols.filter(function (c) {
      const l = c.toLowerCase();
      return l.indexOf('paint') !== -1 && l.indexOf('zone') !== -1;
    });
    const tyreKeyCols = {};
    const mfrCols = {};
    for (const c of cols) {
      const l = c.toLowerCase();
      if (l.indexOf('tyre key') !== -1 && l.indexOf('axle') !== -1) tyreKeyCols[axleNo(c)] = c;
      if (l.indexOf('manufacturer key') !== -1 && l.indexOf('axle') !== -1) mfrCols[axleNo(c)] = c;
    }

    const extraCols = ['Order status financial', 'Order status logistical',
      'Additional equipment (enumeration)', 'FIN', 'Subcategory (ID)',
      'Vehicle alterable until', 'Requested delivery date'].filter(function (c) {
      return cols.indexOf(c) !== -1;
    });

    return rows.map(function (r) {
      let combined;
      if (codeCols.length) combined = codeCols.map(function (c) { return s(r[c]); }).join(' ');
      else combined = cols.map(function (c) { return s(r[c]); }).join(' ');

      const out = {
        'Commission no.': r['Commission no.'],
        'Model': r[modelCol],
        'WINGS_codes': extractCodes(combined),
        'WINGS_has_pto': /\bPTO\b/i.test(combined),
        'WINGS_paint': new Set(paintCols.map(function (c) { return normPaint(r[c]); }).filter(Boolean)),
        'WINGS_tyre': new Set(),
      };
      if (cols.indexOf('Baumuster') !== -1 && modelCol !== 'Baumuster') out['Baumuster'] = r['Baumuster'];
      for (const ax of Object.keys(tyreKeyCols)) {
        const key = normTyre(r[tyreKeyCols[ax]]);
        if (!key) continue;
        const mfr = mfrCols[ax] ? normMfr(r[mfrCols[ax]]) : '';
        out['WINGS_tyre'].add(mfr ? key + ' ' + mfr : key);
      }
      for (const c of extraCols) out[c] = r[c];
      return out;
    });
  }

  const api = { parseWings: parseWings, _extractCodes: extractCodes };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.WingsParse = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
