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

  // ---- model_mapping.xlsx : 모든 매칭 규칙 (rules.py 와 동일 시트명) ----
  const RULE_DEFAULTS = {
    normalize_historic: { '3253': '4153', '4140': '4440', '2643': '3343',
      '2851': '2651', '2135': '1835', '2863': '2663', '2853': '2653' },
    normalize_28xx_to_26xx: true,
    reverse_aliases: { '3253': ['4153'] },
    previous_model: { '4453': '4153', '4153': '3253', '3343': '2643', '2853': '2663', '2851': '2661' },
    current_model: { '4453': '4463', '4153': '4163', '3343': '3363', '2853': '2863', '2851': '2861' },
    wings_display_replace: { '4140': '4440', '2651 LS': '2851 LS', '2653 LS': '2853 LS',
      '2663 LS': '2863 LS', '2643 A': '3343 A' },
    vehicle_keywords: {
      'Actros-L': ['2651', '2851', '2653', '2853', '2663', '2863', '2143'],
      'Actros': ['3363'],
      'Arocs': ['2643', '3343', '4153', '4453', '3253', '2135', '4440', '4140'],
    },
  };
  const MAP_SHEETS = {
    '정규화_과거번호': 'normalize_historic',
    '이전모델': 'previous_model',
    '현재모델': 'current_model',
    'WINGS표시치환': 'wings_display_replace',
  };
  // 수동 교정 시트(수동매핑·매칭_별칭(수동))는 폐기했다 — SAM 문서의 번호·코드로만 매칭한다.
  const LIST_SHEETS = {
    '차종키워드': 'vehicle_keywords',
    // SAM 문서 번호 → 같은 차를 가리키는 WINGS 번호들. samparse.buildSamMapping 이
    // 이 표로 모델키를 늘려 준다(예: 3253 → 4153).
    '모델별칭': 'reverse_aliases',
  };

  function loadRules(buf) {
    const merged = JSON.parse(JSON.stringify(RULE_DEFAULTS));
    if (!buf) return merged;
    let wb;
    try { wb = readWb(buf); } catch (e) { return merged; }
    const pairs = function (sheet) {
      const aoa = sheetAoa(wb, sheet);
      const out = [];
      for (let i = 1; i < aoa.length; i++) {
        const row = aoa[i] || [];
        const k = s(row[0]);
        if (k) out.push([k, s(row[1])]);
      }
      return out;
    };
    for (const sheet of Object.keys(MAP_SHEETS)) {
      if (wb.SheetNames.indexOf(sheet) === -1) continue;
      const obj = {};
      for (const [k, v] of pairs(sheet)) obj[k] = v;
      if (Object.keys(obj).length) merged[MAP_SHEETS[sheet]] = obj;
    }
    for (const sheet of Object.keys(LIST_SHEETS)) {
      if (wb.SheetNames.indexOf(sheet) === -1) continue;
      const obj = {};
      for (const [k, v] of pairs(sheet)) {
        obj[k] = v.split(',').map(function (x) { return x.trim(); }).filter(Boolean);
      }
      if (Object.keys(obj).length) merged[LIST_SHEETS[sheet]] = obj;
    }
    if (wb.SheetNames.indexOf('옵션') !== -1) {
      for (const [k, v] of pairs('옵션')) {
        if (k === 'normalize_28xx_to_26xx') {
          merged.normalize_28xx_to_26xx = ['true', '1', 'yes', 'y', 'on'].indexOf(v.toLowerCase()) !== -1;
        }
      }
    }
    return merged;
  }

  const api = {
    loadMandatory: loadMandatory,
    loadModelCategory: loadModelCategory,
    loadCabMap: loadCabMap,
    loadCodeDict: loadCodeDict,
    loadFactoryControl: loadFactoryControl,
    loadPtoCodes: loadPtoCodes,
    loadRules: loadRules,
    categoryForBaumuster: categoryForBaumuster,
    classifyPrefix: classifyPrefix,
    RULE_DEFAULTS: RULE_DEFAULTS,
    FC_DEFAULTS: FC_DEFAULTS,
    PTO_DEFAULTS: PTO_DEFAULTS,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.RefData = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
