// SharePoint 참조 워크북 로더 — mandatory_codes.py / model_category.py /
// compare._load_cab_map / rules.py 의 xlsx 읽기 부분을 JS 로 옮긴 것.
// 입력은 xlsx 바이트(ArrayBuffer|Uint8Array), 출력은 compare.js 가 쓰는 구조체.
(function (root) {
  'use strict';

  function getXLSX() {
    if (root.XLSX) return root.XLSX;
    if (typeof module !== 'undefined' && module.exports) return require('xlsx');
    throw new Error('SheetJS(XLSX) 가 로드되지 않았습니다.');
  }

  // 줄바꿈은 \n 으로 통일한다 — SheetJS 는 셀의 CRLF 를 그대로 주지만 openpyxl(파이썬
  // 빌드)은 \n 으로 돌려주므로, 통일해야 두 산출물이 완전히 같아진다.
  function s(v) {
    return (v === null || v === undefined) ? '' : String(v).replace(/\r\n/g, '\n').trim();
  }

  function sheetAoa(wb, name) {
    const XLSX = getXLSX();
    const ws = (name && wb.Sheets[name]) || wb.Sheets[wb.SheetNames[0]];
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: null });
  }

  function readWb(buf) {
    const XLSX = getXLSX();
    return XLSX.read(buf instanceof Uint8Array ? buf : new Uint8Array(buf), { type: 'array' });
  }

  // ---- mandatory-codes.xlsx : A=cat B=group C=code D=desc E=note ----
  function loadMandatory(buf) {
    const wb = readWb(buf);
    const aoa = sheetAoa(wb, wb.SheetNames.indexOf('Mandatory') !== -1 ? 'Mandatory' : null);
    const desc = {}, set = new Set(), groups = {}, cats = {};
    for (let i = 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const code = s(row[2]);
      if (!code) continue;
      const cat = (s(row[0]) || 'all').toLowerCase() || 'all';
      const group = s(row[1]);
      const d = s(row[3]);
      set.add(code);
      if (!(code in desc) || (!desc[code] && d)) desc[code] = d;
      if (group) { (groups[group] = groups[group] || new Set()).add(code); }
      (cats[code] = cats[code] || new Set()).add(cat);
    }
    return { desc: desc, set: set, groups: groups, cats: cats };
  }

  // ---- model-category.xlsx : A=6자리 BM prefix, B=category ----
  const VALID_CATS = new Set(['tractor', 'rigid', 'tipper']);
  function loadModelCategory(buf) {
    const wb = readWb(buf);
    const aoa = sheetAoa(wb, null);
    const out = {};
    for (let i = 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      if (row[0] === null || row[0] === undefined) continue;
      const m = /^(\d{6})/.exec(s(row[0]));
      if (!m) continue;
      const cat = s(row[1]).toLowerCase();
      if (VALID_CATS.has(cat)) out[m[1]] = cat;
    }
    return out;
  }

  const TRACTOR_PREFIXES = new Set(['963425', '964416', '963403', '964424', '983403']);
  const TIPPER_PREFIXES = new Set(['964230', '964214', '964231']);
  function classifyPrefix(prefix) {
    if (!prefix) return '';
    if (TRACTOR_PREFIXES.has(prefix)) return 'tractor';
    if (TIPPER_PREFIXES.has(prefix)) return 'tipper';
    if (prefix.indexOf('964') === 0) return 'rigid';
    return '';
  }
  function categoryForBaumuster(bm, table) {
    const m = /^(\d{6})/.exec(s(bm));
    if (!m) return '';
    return (table && table[m[1]]) || classifyPrefix(m[1]);
  }

  // ---- cab.xlsx : A=WINGS cab code, C=SAM 파일명에 나타나는 cab variant ----
  function loadCabMap(buf) {
    const wb = readWb(buf);
    const aoa = sheetAoa(wb, null);
    const out = {};
    for (let i = 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      if (!row[0]) continue;
      const code = s(row[0]).toUpperCase();
      const variant = row.length > 2 && row[2] ? s(row[2]).toUpperCase() : '';
      if (code && variant) out[code] = variant;
    }
    return out;
  }

  // ---- mbtruck-spec-data.xlsx [code_dict] : B=code, C=name_en ----
  function loadCodeDict(buf, sheet) {
    const wb = readWb(buf);
    const name = (sheet && wb.SheetNames.indexOf(sheet) !== -1) ? sheet
      : (wb.SheetNames.indexOf('code_dict') !== -1 ? 'code_dict' : null);
    const aoa = sheetAoa(wb, name);
    const out = {};
    for (let i = 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const code = s(row[1]);
      if (!code || code.toLowerCase() === 'nan') continue;
      out[code] = s(row[2]);
    }
    return out;
  }

  // ---- factory-control-codes.xlsx [FactoryControl] : A=유형 B=코드/접두어 C=설명 ----
  // 유형이 '접두어(prefix)' 면 그 머리글자로 시작하는 모든 코드가 Factory Control,
  // '개별코드(code)' 면 그 코드 하나만, '예외(except)' 면 접두어에 걸려도 일반 코드로 본다
  // (예: Z5M·Z5U 는 Z 로 시작하지만 PTO 옵션이다). 워크북이 없으면 아래 기본값을 쓴다.
  const FC_DEFAULTS = {
    prefixes: ['I', 'O', 'Z', 'U'],
    codes: ['DUP0', 'A0B', 'E0D', 'E0Q', 'J7G'],
    excepts: ['Z5M', 'Z5U'],
  };
  function loadFactoryControl(buf) {
    const fallback = {
      prefixes: new Set(FC_DEFAULTS.prefixes),
      codes: new Set(FC_DEFAULTS.codes),
      excepts: new Set(FC_DEFAULTS.excepts),
    };
    if (!buf) return fallback;
    let wb;
    try { wb = readWb(buf); } catch (e) { return fallback; }
    const aoa = sheetAoa(wb, wb.SheetNames.indexOf('FactoryControl') !== -1 ? 'FactoryControl' : null);
    const prefixes = new Set(), codes = new Set(), excepts = new Set();
    for (let i = 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const type = s(row[0]);
      const val = s(row[1]).toUpperCase();
      if (!val || val === '-' || val === '—') continue;          // 안내 문구 행
      if (/except|예외/i.test(type)) excepts.add(val);
      else if (/prefix|접두어/i.test(type)) { if (val.length === 1) prefixes.add(val); }
      else codes.add(val);
    }
    if (!prefixes.size && !codes.size) return fallback;
    // '예외' 행이 하나도 없는(옛 양식) 워크북이면 기본 예외를 그대로 쓴다.
    // 한 줄이라도 적어 두면 워크북 쪽이 전부 이긴다.
    return {
      prefixes: prefixes,
      codes: codes,
      excepts: excepts.size ? excepts : new Set(FC_DEFAULTS.excepts),
    };
  }

  // ---- pto-codes.xlsx [PTO] : A=유형 B=코드 C=설명 ----
  // PTO 여부는 '코드 설명에 PTO 라는 글자가 있나' 가 아니라 이 표로 정한다. 설명 문구에
  // 기대면 준비(provision)·제어 코드까지 PTO 로 넘어가고, 그 차가 왜 PTO 로 잡혔는지
  // 사람이 볼 수도 없다. WINGS 주문과 SAM 견적서 양쪽에 같은 표를 쓴다.
  // 유형이 '제외(except)' 면 PTO 관련이어도 판정에 쓰지 않는다.
  // backend/pto_codes.py 와 같은 값 — 한쪽만 고치지 말 것.
  const PTO_DEFAULTS = {
    codes: ['N0A', 'N0B', 'N0C', 'N0D', 'N0E', 'N0F', 'N0G', 'N0H',
      'N1A', 'N1B', 'N1C', 'N1D', 'N1E', 'N1F', 'N1G', 'N1H', 'N1I', 'N1J', 'N1K', 'N1L', 'N1M',
      'N2E', 'Z5M', 'Z5S', 'Z5U'],
    excepts: ['N6P', 'U2K'],
  };
  function loadPtoCodes(buf) {
    const fallback = {
      codes: new Set(PTO_DEFAULTS.codes),
      excepts: new Set(PTO_DEFAULTS.excepts),
    };
    if (!buf) return fallback;
    let wb;
    try { wb = readWb(buf); } catch (e) { return fallback; }
    const aoa = sheetAoa(wb, wb.SheetNames.indexOf('PTO') !== -1 ? 'PTO' : null);
    const codes = new Set(), excepts = new Set();
    for (let i = 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const type = s(row[0]);
      const val = s(row[1]).toUpperCase();
      if (!val || val === '-' || val === '—') continue;          // 안내 문구 행
      if (/except|제외/i.test(type)) excepts.add(val);
      else codes.add(val);
    }
    if (!codes.size) return fallback;
    return { codes: codes, excepts: excepts };
  }

  // ---- model_mapping.xlsx : 모델명 변경 (rules.py 와 같은 규칙) ----
  // 규칙은 '모델명변경' 시트 한 장:  From | To | 메모
  //   WINGS 모델명(축·DNA 를 뗀 이름, 예 '2651 LS')이 From 과 같으면 To 로 바꾼다 — 한 방향.
  //   바뀐 이름이 대시보드 Type 이자 SAM 을 찾는 번호다(표시와 매칭이 같은 이름을 쓴다).
  // 예전 규칙(대체번호·SAM별칭·옵션·차종·이전/현재모델)은 폐기했다. 예전 워크북에서는
  // 표시치환만 옮겨 온다 — '매칭규칙' 시트의 표시치환 행, 또는 'WINGS표시치환' 시트.
  const RULE_DEFAULTS = {
    model_rename: { '4140': '4440', '2651 LS': '2851 LS', '2653 LS': '2853 LS',
      '2663 LS': '2863 LS', '2643 A': '3343 A' },
  };
  const RULE_SHEET = '모델명변경';
  const RULE_HEADER = ['From (WINGS 모델명)', 'To (바꿀 모델명)', '메모'];

  // 시트 이름 → AOA 를 돌려주는 함수(getAoa)와 시트 이름 목록으로 규칙을 만든다.
  // 워크북(xlsx 바이트)에서도, 편집 중인 화면의 시트에서도 같은 규칙이 나와야 해서 분리했다.
  function rulesFromSheets(names, getAoa) {
    const merged = JSON.parse(JSON.stringify(RULE_DEFAULTS));
    const has = (n) => names.indexOf(n) !== -1;
    const rowsOf = (sheet) => (getAoa(sheet) || []).slice(1).map((row) => (row || []).map(s));
    let found = null;
    if (has(RULE_SHEET)) {
      found = {};
      for (const [from, to] of rowsOf(RULE_SHEET)) if (from && to) found[from] = to;
    } else if (has('매칭규칙')) {
      found = {};
      for (const [kind, from, to] of rowsOf('매칭규칙')) {
        const k = String(kind || '').replace(/\s+/g, '').toLowerCase();
        if ((k === '표시치환' || k === 'wings_display_replace' || k === '모델명변경') && from && to) found[from] = to;
      }
    } else if (has('WINGS표시치환')) {
      found = {};
      for (const [from, to] of rowsOf('WINGS표시치환')) if (from && to) found[from] = to;
    }
    // 시트가 있는데 비었으면 '규칙 없음'으로 본다(예전과 달리 기본값으로 되돌리지 않는다 —
    // 한 장짜리라 일부러 다 지운 것과 실수를 가릴 수 없고, 지운 대로 동작하는 편이 예측 가능하다).
    if (found) merged.model_rename = found;
    return merged;
  }

  function loadRules(buf) {
    if (!buf) return JSON.parse(JSON.stringify(RULE_DEFAULTS));
    let wb;
    try { wb = readWb(buf); } catch (e) { return JSON.parse(JSON.stringify(RULE_DEFAULTS)); }
    return rulesFromSheets(wb.SheetNames, function (n) { return sheetAoa(wb, n); });
  }

  // 규칙 → '모델명변경' 시트 AOA. notes = { From: 메모 }.
  function rulesToAoa(rules, notes) {
    notes = notes || {};
    const aoa = [RULE_HEADER.slice()];
    const m = rules.model_rename || {};
    for (const k of Object.keys(m)) aoa.push([k, String(m[k]), notes[k] || '']);
    return aoa;
  }

  const api = {
    loadMandatory: loadMandatory,
    loadModelCategory: loadModelCategory,
    loadCabMap: loadCabMap,
    loadCodeDict: loadCodeDict,
    loadFactoryControl: loadFactoryControl,
    loadPtoCodes: loadPtoCodes,
    loadRules: loadRules,
    rulesFromSheets: rulesFromSheets,
    rulesToAoa: rulesToAoa,
    RULE_SHEET: RULE_SHEET,
    RULE_HEADER: RULE_HEADER,
    categoryForBaumuster: categoryForBaumuster,
    classifyPrefix: classifyPrefix,
    RULE_DEFAULTS: RULE_DEFAULTS,
    FC_DEFAULTS: FC_DEFAULTS,
    PTO_DEFAULTS: PTO_DEFAULTS,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.RefData = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
