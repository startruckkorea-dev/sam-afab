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

  // ---- model_mapping.xlsx : 모든 매칭 규칙 (rules.py 와 같은 규칙) ----
  // 규칙은 '매칭규칙' 시트 한 장에 모은다:  종류 | 기준 | 값 | 메모
  //   대체번호  기준=WINGS 번호, 값=SAM 번호        → normalize_historic
  //             정확한 번호로 생산월·이전월을 다 찾은 뒤에야 쓰는 '대신 찾을 번호'(양방향).
  //   SAM별칭   기준=SAM 문서 번호, 값=WINGS 번호들  → reverse_aliases
  //             그 SAM 문서를 처음부터 WINGS 번호로도 인식한다(같은 달이면 제 번호처럼 붙는다).
  //   표시치환  기준=WINGS 표기, 값=통일할 표기      → wings_display_replace (목록 표시용)
  //   차종      기준=차종(Vehicle), 값=번호들         → vehicle_keywords
  //   옵션      기준=normalize_28xx_to_26xx, 값=true/false
  // 예전 워크북(종류별 시트 9장)도 그대로 읽는다 — '매칭규칙' 시트가 있으면 그것만 본다.
  // 이전모델·현재모델(세대 이웃)은 매칭에 쓰인 적이 없어 폐기했다.
  const RULE_DEFAULTS = {
    normalize_historic: { '3253': '4153', '4140': '4440', '2643': '3343',
      '2851': '2651', '2135': '1835', '2863': '2663', '2853': '2653' },
    normalize_28xx_to_26xx: true,
    reverse_aliases: { '3253': ['4153'] },
    wings_display_replace: { '4140': '4440', '2651 LS': '2851 LS', '2653 LS': '2853 LS',
      '2663 LS': '2863 LS', '2643 A': '3343 A' },
    vehicle_keywords: {
      'Actros-L': ['2651', '2851', '2653', '2853', '2663', '2863', '2143'],
      'Actros': ['3363'],
      'Arocs': ['2643', '3343', '4153', '4453', '3253', '2135', '4440', '4140'],
    },
  };
  const RULE_SHEET = '매칭규칙';
  const RULE_HEADER = ['종류', '기준', '값', '메모'];
  // 종류(시트에 적는 이름) → 규칙 키. list = 값이 쉼표 목록.
  const RULE_KINDS = [
    { kind: '대체번호', key: 'normalize_historic', list: false },
    { kind: 'SAM별칭', key: 'reverse_aliases', list: true },
    { kind: '표시치환', key: 'wings_display_replace', list: false },
    { kind: '차종', key: 'vehicle_keywords', list: true },
    { kind: '옵션', key: null, list: false },
  ];
  const OPTION_KEYS = ['normalize_28xx_to_26xx'];
  // 예전 시트 이름 → 규칙 키 (마이그레이션·호환 읽기용)
  const LEGACY_MAP_SHEETS = {
    '정규화_과거번호': 'normalize_historic',
    'WINGS표시치환': 'wings_display_replace',
  };
  const LEGACY_LIST_SHEETS = {
    '차종키워드': 'vehicle_keywords',
    '모델별칭': 'reverse_aliases',
  };

  const truthy = (v) => ['true', '1', 'yes', 'y', 'on'].indexOf(String(v).toLowerCase()) !== -1;
  const splitList = (v) => String(v).split(',').map(function (x) { return x.trim(); }).filter(Boolean);
  function kindOf(name) {
    const n = s(name).replace(/\s+/g, '').toLowerCase();
    for (const k of RULE_KINDS) {
      if (k.kind.toLowerCase() === n || (k.key && k.key.toLowerCase() === n)) return k;
    }
    return null;
  }

  // 시트 이름 → AOA 를 돌려주는 함수(getAoa)와 시트 이름 목록으로 규칙을 만든다.
  // 워크북(xlsx 바이트)에서도, 편집 중인 화면의 시트에서도 같은 규칙이 나와야 해서 분리했다.
  function rulesFromSheets(names, getAoa) {
    const merged = JSON.parse(JSON.stringify(RULE_DEFAULTS));
    const has = (n) => names.indexOf(n) !== -1;
    const pairs = function (sheet) {
      const aoa = getAoa(sheet) || [];
      const out = [];
      for (let i = 1; i < aoa.length; i++) {
        const row = aoa[i] || [];
        const k = s(row[0]);
        if (k) out.push([k, s(row[1]), s(row[2])]);
      }
      return out;
    };
    if (has(RULE_SHEET)) {
      const found = {};
      for (const [kindName, k, v] of pairs(RULE_SHEET)) {
        const kd = kindOf(kindName);
        if (!kd || !s(k)) continue;
        if (!kd.key) {
          if (OPTION_KEYS.indexOf(k) !== -1) merged[k] = truthy(v);
          continue;
        }
        if (!found[kd.key]) found[kd.key] = {};
        if (kd.list) {
          const prev = found[kd.key][k] || [];
          for (const x of splitList(v)) if (prev.indexOf(x) === -1) prev.push(x);
          found[kd.key][k] = prev;
        } else {
          found[kd.key][k] = v;
        }
      }
      // 종류가 하나도 없으면 기본값 유지 — 실수로 비운 시트가 규칙을 통째로 지우지 않도록.
      for (const key of Object.keys(found)) if (Object.keys(found[key]).length) merged[key] = found[key];
      return merged;
    }
    for (const sheet of Object.keys(LEGACY_MAP_SHEETS)) {
      if (!has(sheet)) continue;
      const obj = {};
      for (const [k, v] of pairs(sheet)) obj[k] = v;
      if (Object.keys(obj).length) merged[LEGACY_MAP_SHEETS[sheet]] = obj;
    }
    for (const sheet of Object.keys(LEGACY_LIST_SHEETS)) {
      if (!has(sheet)) continue;
      const obj = {};
      for (const [k, v] of pairs(sheet)) obj[k] = splitList(v);
      if (Object.keys(obj).length) merged[LEGACY_LIST_SHEETS[sheet]] = obj;
    }
    if (has('옵션')) {
      for (const [k, v] of pairs('옵션')) {
        if (OPTION_KEYS.indexOf(k) !== -1) merged[k] = truthy(v);
      }
    }
    return merged;
  }

  function loadRules(buf) {
    if (!buf) return JSON.parse(JSON.stringify(RULE_DEFAULTS));
    let wb;
    try { wb = readWb(buf); } catch (e) { return JSON.parse(JSON.stringify(RULE_DEFAULTS)); }
    return rulesFromSheets(wb.SheetNames, function (n) { return sheetAoa(wb, n); });
  }

  // 규칙 → '매칭규칙' 시트 AOA. 예전 워크북을 통합 시트로 옮길 때 쓴다 — 지금 실제로
  // 적용되던 규칙(기본값 포함)을 그대로 옮기므로 매칭 결과가 바뀌지 않는다.
  function rulesToAoa(rules, notes) {
    notes = notes || {};
    const aoa = [RULE_HEADER.slice()];
    for (const kd of RULE_KINDS) {
      if (!kd.key) continue;
      const obj = rules[kd.key] || {};
      for (const k of Object.keys(obj)) {
        const v = obj[k];
        aoa.push([kd.kind, k, Array.isArray(v) ? v.join(', ') : String(v), notes[kd.kind] || '']);
      }
    }
    for (const k of OPTION_KEYS) {
      aoa.push(['옵션', k, rules[k] ? 'true' : 'false', notes['옵션'] || '']);
    }
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
    RULE_KINDS: RULE_KINDS,
    categoryForBaumuster: categoryForBaumuster,
    classifyPrefix: classifyPrefix,
    RULE_DEFAULTS: RULE_DEFAULTS,
    FC_DEFAULTS: FC_DEFAULTS,
    PTO_DEFAULTS: PTO_DEFAULTS,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.RefData = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
