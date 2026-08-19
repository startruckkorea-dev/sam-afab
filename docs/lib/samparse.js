// backend/sam_parser.py 의 JS 포팅 — SAM Internal-Quotation .docx 에서
// 코드/모델번호/Baumuster/Subcategory/Paint/Tyre 를 추출한다.
// DOM 의존 없이 문자열/정규식만 사용하므로 브라우저·Node 양쪽에서 동일하게 돈다.
(function (root) {
  'use strict';

  const Unzip = (typeof module !== 'undefined' && module.exports)
    ? require('./unzip.js') : root.Unzip;

  // ---------- XML 헬퍼 (ElementTree.iter 와 같은 문서 순서 보장) ----------
  function decodeEntities(s) {
    return s.replace(/&(#x?[0-9A-Fa-f]+|amp|lt|gt|quot|apos);/g, function (_, e) {
      if (e === 'amp') return '&';
      if (e === 'lt') return '<';
      if (e === 'gt') return '>';
      if (e === 'quot') return '"';
      if (e === 'apos') return "'";
      if (e.charAt(0) === '#') {
        const n = e.charAt(1) === 'x' || e.charAt(1) === 'X'
          ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return isNaN(n) ? _ : String.fromCodePoint(n);
      }
      return _;
    });
  }

  const RE_WT = /<w:t(?![\w:-])[^>]*>([\s\S]*?)<\/w:t>/g;

  /** chunk 안의 모든 <w:t> 텍스트를 이어붙인다 (Python: "".join(t.text ...)). */
  function wtText(chunk) {
    let out = '';
    RE_WT.lastIndex = 0;
    let m;
    while ((m = RE_WT.exec(chunk)) !== null) out += decodeEntities(m[1]);
    return out;
  }

  /** tag 의 모든 블록을 문서 순서(시작 위치 오름차순, 부모→자식)로 돌려준다. */
  function findBlocks(xml, tag) {
    const re = new RegExp('<' + tag + '(?![\\w:-])[^>]*>|</' + tag + '\\s*>', 'g');
    const stack = [];
    const out = [];
    let m;
    while ((m = re.exec(xml)) !== null) {
      const tok = m[0];
      if (tok.charAt(1) === '/') {                 // close
        const start = stack.pop();
        if (start !== undefined) out.push([start, re.lastIndex]);
      } else if (tok.slice(-2) === '/>') {         // self-closing
        out.push([m.index, re.lastIndex]);
      } else {                                     // open
        stack.push(m.index);
      }
    }
    out.sort(function (a, b) { return a[0] - b[0]; });
    return out.map(function (p) { return xml.slice(p[0], p[1]); });
  }

  // ---------- normalize_model ----------
  function normalizeModel(model) {
    if (model === null || model === undefined) return '';
    let s;
    if (typeof model === 'number') {
      s = Number.isNaN(model) ? '' : (Number.isInteger(model) ? String(model) : String(model));
    } else {
      s = String(model);
    }
    const tmp = s.replace(/\d[Xx]\d/g, '');            // 축 정보(8x4) 제거
    return tmp.toUpperCase().replace(/DNA/g, '').trim().replace(/[^A-Z0-9]/g, '');
  }

  // ---------- 본체 ----------
  const EQ_HEADERS = new Set(['STANDARD EQUIPMENT', 'SPECIAL EQUIPMENT',
    'ADDITIONAL EQUIPMENT', 'EQUIPMENT OVERVIEW']);

  function hasDigit(s) { return /\d/.test(s); }

  /**
   * SAM .docx 한 개 파싱.
   * @param {ArrayBuffer|Uint8Array} buf  .docx 바이트
   * @param {string} name                파일명
   * @param {object} ctx  { knownCodes: Set, ptoCodes: {codes:Set, excepts:Set} }
   * @returns {object|null} 파싱 결과(코드가 없으면 null)
   */
  async function parseSamDocx(buf, name, ctx) {
    ctx = ctx || {};
    const knownCodes = ctx.knownCodes || new Set();

    let modelRaw = null;
    let codes = new Set();
    const paint = new Set();
    const tyre = new Set();
    let fullText = '';

    const xml = await Unzip.readTextFile(buf, 'word/document.xml');
    fullText = wtText(xml);

    // --- 장비 코드: Equipment overview 표 (없으면 Standard equipment 표) ---
    const tables = findBlocks(xml, 'w:tbl');
    let eqOverview = null, fallback = null;
    for (const tbl of tables) {
      const tblText = wtText(tbl);
      if (eqOverview === null && tblText.indexOf('Equipment overview') !== -1) eqOverview = tbl;
      if (fallback === null && tblText.indexOf('Standard equipment') !== -1) fallback = tbl;
    }
    const target = eqOverview || fallback;

    if (target) {
      let section = null;
      for (const para of findBlocks(target, 'w:p')) {
        const paraText = wtText(para).trim();
        const up = paraText.toUpperCase();
        if (EQ_HEADERS.has(up)) { section = up; continue; }
        if (!paraText || section === null) continue;
        if (section === 'STANDARD EQUIPMENT' || section === 'SPECIAL EQUIPMENT') {
          const re = /\b([A-Z][A-Z0-9]{2,3})\b/g;
          let m;
          while ((m = re.exec(up)) !== null) codes.add(m[1]);
        } else if (section === 'ADDITIONAL EQUIPMENT') {
          const m = /^([A-Z][A-Z0-9]{2,3})\b/.exec(up);
          if (m) codes.add(m[1]);
        }
      }
    }

    // --- Paint / Tyre (CTT) — 문서 전체 문단에서 섹션 추적 ---
    let sec = null;
    for (const para of findBlocks(xml, 'w:p')) {
      const ptxt = wtText(para).trim();
      if (!ptxt) continue;
      if (ptxt === 'Paint' || ptxt === 'Paints') { sec = 'PAINT'; continue; }
      if (ptxt === 'Tyres' || ptxt === 'Tyre') { sec = 'TYRES'; continue; }
      if (EQ_HEADERS.has(ptxt.toUpperCase())) { sec = null; continue; }
      if (sec === 'PAINT') {
        const re = /MB\s*(\d{4})/g;
        let m;
        while ((m = re.exec(ptxt)) !== null) paint.add(m[1]);
      } else if (sec === 'TYRES') {
        const m = /R\s*\d{2,3},\d\s*([A-Z0-9]{6})\s+(\d{2})(?=[\d.,])/.exec(ptxt);
        if (m) tyre.add(m[1] + ' ' + m[2]);
        else {
          const m2 = /R\s*\d{2,3},\d\s*([A-Z0-9]{6})/.exec(ptxt);
          if (m2) tyre.add(m2[1]);
        }
      }
    }

    // --- 본문 Vehicle type (= SAM Baumuster). DNA 접미사는 표시용으로 보존 ---
    const upAll = fullText.toUpperCase();
    const PATTERNS = [
      /VEHICLE\s*TYPE[:\s]+([0-9]{4}\s*[A-Z]{1,3}(?:DNA)?)(?=DRIVETRAIN|SUBCATEGORY|BAUMUSTER|\s|[0-9]|$)/,
      /\bTYPE[:\s]+([0-9]{4}\s*[A-Z]{1,3}(?:DNA)?)(?=DRIVETRAIN|SUBCATEGORY|BAUMUSTER|\s|[0-9]|$)/,
      /\bMODEL[:\s]+([0-9]{4}\s*[A-Z]{1,3}(?:DNA)?)(?=DRIVETRAIN|SUBCATEGORY|BAUMUSTER|\s|[0-9]|$)/,
    ];
    for (const p of PATTERNS) {
      const m = p.exec(upAll);
      if (m) { modelRaw = m[1].trim(); break; }
    }

    // 프로즈 단어 걸러내기: 숫자를 포함하거나 실제 SA 코드일 때만 유지
    codes = new Set([...codes].filter(function (c) {
      return hasDigit(c) || knownCodes.has(c);
    }));

    // --- 파일명 모델 (= SAM now) ---
    const bodyModel = (modelRaw || '').trim();
    const fnameUpper = name.toUpperCase();
    let fnameModel = null;
    let mf = /(\d{4}\s*[A-Z]{0,3})(?=\s+[A-Z]\d[A-Z]|\s+\d[Xx]\d|\s+HUB|\s+CLASSIC|\s+EURO|\s|$)/.exec(fnameUpper);
    if (!mf) mf = /(\d{4}\s*[A-Z]{1,3})/.exec(fnameUpper);
    if (mf) fnameModel = mf[1].trim();

    const modelNow = fnameModel || bodyModel;          // 파일제목 우선
    const modelBaumuster = bodyModel || fnameModel;    // 본문 Vehicle type 우선

    // --- Baumuster / Subcategory (본문에서만) ---
    let bodyBm = '', bodySub = '';
    if (fullText) {
      const m1 = /Baumuster[:\s]*([0-9]{5,})/i.exec(fullText);
      if (m1) bodyBm = m1[1];
      const m2 = /Subcategory[:\s]*([0-9]{1,3}[A-Za-z])/i.exec(fullText);
      if (m2) bodySub = m2[1];
    }

    if (!((modelNow || modelBaumuster) && codes.size)) return null;

    // --- PTO 판정 ---
    // 문서가 실제로 갖고 있는 코드만 본다. 기준은 pto-codes.xlsx (ctx.ptoCodes) 한 곳 —
    // 본문에 'PTO' 라는 낱말이 스쳐 나오거나 파일명에 PTO 가 붙었다는 이유로 넘기지 않는다.
    // 그런 폴백은 PTO 가 달리지 않은 견적서까지 PTO 로 만들어, 그 달의 비PTO 후보를
    // 통째로 비워 버렸다. WINGS 주문도 같은 표로 판정하므로 양쪽이 대칭이다.
    const ptoTable = ctx.ptoCodes || { codes: new Set(), excepts: new Set() };
    let isPto = false;
    for (const c of codes) {
      if (ptoTable.codes.has(c) && !ptoTable.excepts.has(c)) { isPto = true; break; }
    }

    return {
      codes: codes, paint: paint, tyre: tyre, file: name,
      model_now: modelNow, model_baumuster: modelBaumuster,
      bm: bodyBm, sub: bodySub, is_pto: isPto,
    };
  }

  /**
   * load_sam_from_folder 포팅: 파싱 결과들을 모델키별 매핑으로 모으고
   * reverse_aliases 로 별칭 키를 확장한다.
   * @param {Array} parsed  parseSamDocx 결과 배열
   * @param {object} rules  { reverse_aliases: {num: [alias,...]} }
   * @returns {Map} key -> { true/false(PTO) -> [data, ...] }
   */
  function buildSamMapping(parsed, rules) {
    const mapping = new Map();
    for (const d of parsed) {
      if (!d) continue;
      const keys = new Set();
      const a = normalizeModel(d.model_now); if (a) keys.add(a);
      const b = normalizeModel(d.model_baumuster); if (b) keys.add(b);
      for (const k of keys) {
        if (!mapping.has(k)) mapping.set(k, {});
        const byPto = mapping.get(k);
        const flag = d.is_pto ? 'true' : 'false';
        if (!byPto[flag]) byPto[flag] = [];
        byPto[flag].push(d);
      }
    }
    // 별칭 확장 (WINGS 신형 번호 → SAM 구형 번호)
    const rev = (rules && rules.reverse_aliases) || {};
    for (const key of Array.from(mapping.keys())) {
      const m = /^(\d+)([A-Z]*)$/.exec(key);
      if (!m) continue;
      const num = m[1], suffix = m[2];
      for (const srcPrefix of Object.keys(rev)) {
        if (num !== srcPrefix) continue;
        const aliases = rev[srcPrefix] || [];
        for (const ap of aliases) {
          const aliasKey = ap + suffix;
          if (!mapping.has(aliasKey)) mapping.set(aliasKey, mapping.get(key));
        }
      }
    }
    return mapping;
  }

  const api = {
    parseSamDocx: parseSamDocx,
    buildSamMapping: buildSamMapping,
    normalizeModel: normalizeModel,
    _wtText: wtText,
    _findBlocks: findBlocks,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SamParse = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
