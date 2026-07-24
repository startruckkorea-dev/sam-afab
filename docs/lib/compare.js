// backend/compare.py 의 JS 포팅 — WINGS 행 ↔ 월별 SAM 맵 비교.
// 결과 행의 키/값은 Python 판과 동일해야 한다(대시보드가 그대로 읽는다).
(function (root) {
  'use strict';

  const isNode = (typeof module !== 'undefined' && module.exports);
  const SamParse = isNode ? require('./samparse.js') : root.SamParse;
  const RefData = isNode ? require('./refdata.js') : root.RefData;

  const normalizeModel = SamParse.normalizeModel;

  // _is_fc 의 IOZU 규칙이 OPTION_CODE_MAP 유래분을 이미 포함하므로, 남는 건 이 5개뿐.
  const EXTRA_EXCEPT = new Set(['DUP0', 'A0B', 'E0D', 'E0Q', 'J7G']);

  function s(v) { return (v === null || v === undefined) ? '' : String(v); }
  function strip(v) { return s(v).trim(); }

  // ---- 날짜 유틸 (pandas to_datetime / strftime 대응) ----
  function toDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number') {                    // Excel serial
      const ms = Math.round((v - 25569) * 86400 * 1000);
      const d = new Date(ms);
      return isNaN(d.getTime()) ? null : d;
    }
    const t = strip(v);
    if (!t) return null;
    const d = new Date(t.replace(' ', 'T'));
    return isNaN(d.getTime()) ? null : d;
  }
  function fmtDate(d) {
    if (!d) return '';
    const p = function (n) { return String(n).padStart(2, '0'); };
    return d.getUTCFullYear() + '-' + p(d.getUTCMonth() + 1) + '-' + p(d.getUTCDate());
  }
  function dateOnlyUTC(d) {
    return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  // ---- 후보 선택 헬퍼 (_candidates / _match_score / _pick / _cab_ok) ----
  function candidates(entry, preferPto) {
    if (!entry || typeof entry !== 'object') return [];
    const a = entry[preferPto ? 'true' : 'false'];
    const b = entry[preferPto ? 'false' : 'true'];
    const lst = (a && a.length) ? a : ((b && b.length) ? b : []);
    return Array.isArray(lst) ? lst : [lst];
  }
  function matchScore(data, wingsBm, wingsSub, expectedCabs) {
    let score = 0;
    if (expectedCabs && expectedCabs.size) {
      const f = s(data.file).toUpperCase();
      for (const cv of expectedCabs) { if (cv && f.indexOf(cv) !== -1) { score += 3; break; } }
    }
    if (wingsBm && strip(data.bm) === wingsBm) score += 2;
    if (wingsSub && strip(data.sub).toLowerCase() === wingsSub) score += 1;
    return score;
  }
  function pick(entry, preferPto, wingsBm, wingsSub, expectedCabs) {
    const cands = candidates(entry, preferPto);
    if (!cands.length) return null;
    // Python max() 는 동점일 때 first 를 유지한다.
    let best = cands[0], bestScore = matchScore(cands[0], wingsBm, wingsSub, expectedCabs);
    for (let i = 1; i < cands.length; i++) {
      const sc = matchScore(cands[i], wingsBm, wingsSub, expectedCabs);
      if (sc > bestScore) { best = cands[i]; bestScore = sc; }
    }
    return best;
  }
  function cabOk(data, expectedCabs) {
    if (!data || !expectedCabs || !expectedCabs.size) return false;
    const f = s(data.file).toUpperCase();
    for (const cv of expectedCabs) if (cv && f.indexOf(cv) !== -1) return true;
    return false;
  }
  function findSamDataByFile(samMapsList, preferPto, fileSub) {
    const fs = strip(fileSub).toLowerCase();
    if (!fs) return null;
    for (const prefer of [preferPto, !preferPto]) {
      const flag = prefer ? 'true' : 'false';
      for (const map of samMapsList) {
        for (const entry of map.values()) {
          if (!entry || typeof entry !== 'object') continue;
          for (const d of (entry[flag] || [])) {
            if (s(d.file).toLowerCase().indexOf(fs) !== -1) return d;
          }
        }
      }
    }
    return null;
  }

  function splitModel(str) {
    const m = /^(\d+)([A-Z]*)$/.exec(str);
    return m ? [m[1], m[2]] : [str, ''];
  }

  const setDiff = (a, b) => new Set([...a].filter((x) => !b.has(x)));
  const setInter = (a, b) => new Set([...a].filter((x) => b.has(x)));
  const setXor = (a, b) => new Set([...setDiff(a, b), ...setDiff(b, a)]);
  const sortedArr = (st) => [...st].filter(Boolean).sort();

  /**
   * @param {Array<object>} wingsRows  wingsparse.parseWings 결과
   * @param {Object} samMapsByMonth    { yyyymm(number): Map(key -> {true/false:[data]}) }
   * @param {Object} ref  { rules, mandatory:{set,groups,cats}, modelCategory, cabMap, codeDesc, today? }
   * @returns {Array<object>} 결과 행 배열
   */
  function compare(wingsRows, samMapsByMonth, ref) {
    const rules = ref.rules || {};
    const mand = ref.mandatory || { set: new Set(), groups: {}, cats: {} };
    const mandSet = mand.set || new Set();
    const mandGroups = mand.groups || {};
    const mandCats = mand.cats || {};
    const modelCategory = ref.modelCategory || {};
    const cabMap = ref.cabMap || {};
    const codeDesc = ref.codeDesc || {};
    const today = ref.today ? new Date(ref.today) : new Date();
    const todayUTC = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());

    let groupedCodes = new Set();
    for (const g of Object.keys(mandGroups)) for (const c of mandGroups[g]) groupedCodes.add(c);

    function mandApplies(code, rowCat) {
      const cats = mandCats[code] || new Set(['all']);
      if (cats.has('all')) return true;
      return !!rowCat && cats.has(rowCat);
    }
    function isFc(c) {
      return EXTRA_EXCEPT.has(c) || (!!c && 'IOZU'.indexOf(c[0]) !== -1);
    }

    const manualNorm = {};
    for (const k of Object.keys(rules.manual_map || {})) manualNorm[normalizeModel(k)] = rules.manual_map[k];

    const sortedYyyymm = Object.keys(samMapsByMonth).map(Number).sort((a, b) => a - b);

    function samMapsForProdDate(prodRaw) {
      if (!sortedYyyymm.length) return [];
      if (prodRaw) {
        const d = toDate(prodRaw);
        if (d) {
          const ym = d.getUTCFullYear() * 100 + (d.getUTCMonth() + 1);
          const byDist = sortedYyyymm.slice().sort((a, b) => Math.abs(a - ym) - Math.abs(b - ym));
          return byDist.map((m) => samMapsByMonth[m]);
        }
      }
      return sortedYyyymm.slice().reverse().map((m) => samMapsByMonth[m]);
    }

    const out = [];
    for (const r of wingsRows) {
      const prodRaw = ('Requested delivery date' in r) ? r['Requested delivery date'] : '';
      const samMapsList = samMapsForProdDate(prodRaw);

      const com = r['Commission no.'];
      const modelRaw = (r['Model'] !== null && r['Model'] !== undefined && r['Model'] !== '')
        ? r['Model'] : (r['Baumuster'] !== undefined ? r['Baumuster'] : '');
      const wingsCodes = new Set(r['WINGS_codes'] || []);
      const wingsPaint = new Set(r['WINGS_paint'] || []);
      const wingsTyre = new Set(r['WINGS_tyre'] || []);
      const modelNorm = normalizeModel(modelRaw);

      const wingsBm = strip(r['Baumuster']);
      const wingsSub = strip(r['Subcategory (ID)']).toLowerCase();

      const expectedCabs = new Set();
      for (const c of wingsCodes) if (cabMap[c]) expectedCabs.add(cabMap[c]);

      let isPto = false;
      for (const c of wingsCodes) {
        if (s(codeDesc[c]).toUpperCase().indexOf('PTO') !== -1) { isPto = true; break; }
      }
      if (!isPto && r['WINGS_has_pto']) isPto = true;

      function entryForModel(map) {
        const cand = map.get(modelNorm);
        if (candidates(cand, isPto).length) return cand;
        const [numNorm, sufNorm] = splitModel(modelNorm);
        for (const [k, v] of map) {
          const kNorm = normalizeModel(String(k));
          const [numK, sufK] = splitModel(kNorm);
          if (kNorm === modelNorm || (numK === numNorm && sufK === sufNorm)) {
            if (candidates(v, isPto).length) return v;
          }
        }
        return null;
      }

      let samEntry = null, samMap = samMapsList[0] || new Map();
      let fallbackEntry = null, fallbackMap = null, matched = false;
      for (const tryMap of samMapsList) {
        const e = entryForModel(tryMap);
        if (!candidates(e, isPto).length) continue;
        if (fallbackEntry === null) { fallbackEntry = e; fallbackMap = tryMap; }
        if (!expectedCabs.size) { samEntry = e; samMap = tryMap; matched = true; break; }
        if (cabOk(pick(e, isPto, wingsBm, wingsSub, expectedCabs), expectedCabs)) {
          samEntry = e; samMap = tryMap; matched = true; break;
        }
      }
      if (!matched && fallbackEntry !== null) { samEntry = fallbackEntry; samMap = fallbackMap; }

      // PTO 보정: WINGS 코드에 PTO 변형에만 있는 코드가 있으면 PTO 로 본다.
      if (!isPto && samEntry && samEntry['true'] && samEntry['false']) {
        const ptoData = pick(samEntry, true, wingsBm, wingsSub);
        const nptoData = pick(samEntry, false, wingsBm, wingsSub);
        if (ptoData && nptoData) {
          const ptoUnique = setDiff(ptoData.codes, nptoData.codes);
          if (setInter(wingsCodes, ptoUnique).size) isPto = true;
        }
      }

      let samData = pick(samEntry, isPto, wingsBm, wingsSub, expectedCabs);

      const mo = manualNorm[modelNorm];
      if (mo && mo.file) {
        const pinned = findSamDataByFile(samMapsList, isPto, mo.file);
        if (pinned) samData = pinned;
      }

      const samCodes = samData ? samData.codes : new Set();
      const samFile = samData ? samData.file : '';
      const samPaint = samData ? (samData.paint || new Set()) : new Set();
      const samTyre = samData ? (samData.tyre || new Set()) : new Set();

      const rowCat = RefData.categoryForBaumuster(
        wingsBm || (samData ? s(samData.bm) : ''), modelCategory);
      const effMand = new Set([...mandSet].filter((c) => mandApplies(c, rowCat)));
      const effGrouped = setInter(groupedCodes, effMand);

      const onlyW = samCodes.size
        ? sortedArr(new Set([...setDiff(wingsCodes, samCodes)].filter((c) => c && !isFc(c) && !effMand.has(c))))
        : [];
      const onlyS = sortedArr(new Set([...setDiff(samCodes, wingsCodes)].filter((c) => c && !isFc(c) && !effMand.has(c))));
      const exceptRow = samCodes.size
        ? sortedArr(new Set([...setXor(wingsCodes, samCodes)].filter((c) => c && isFc(c))))
        : [];

      const onlyOneSide = setXor(samCodes, wingsCodes);
      const ungrouped = new Set([...onlyOneSide].filter((c) => c && effMand.has(c) && !effGrouped.has(c)));
      const groupFlag = new Set();
      for (const g of Object.keys(mandGroups)) {
        const em = setInter(mandGroups[g], effMand);
        if (!em.size) continue;
        const sHas = setInter(samCodes, em).size > 0;
        const wHas = setInter(wingsCodes, em).size > 0;
        if (sHas !== wHas) for (const c of setInter(onlyOneSide, em)) groupFlag.add(c);
      }
      const mandRow = sortedArr(new Set([...ungrouped, ...groupFlag]));

      const paintMismatch = !!(samData && wingsPaint.size && samPaint.size && setXor(wingsPaint, samPaint).size);
      const tyreMismatch = !!(samData && wingsTyre.size && samTyre.size && setXor(wingsTyre, samTyre).size);

      // SAM 파일명에서 차종/축/캐빈/PTO
      let vehicle = '', axleType = '', cabCode = '', ptoFlag = '';
      if (samFile) {
        const vm = /\b(Actros-L|Actros|Arocs|Atego|eActros|Econic|Unimog)\b/i.exec(samFile);
        if (vm) vehicle = vm[1];
        const am = /\b(\d+x\d+)\b/i.exec(samFile);
        if (am) axleType = am[1];
        const cm = /\b([A-Z]\d[A-Z])\b/.exec(samFile);
        if (cm) cabCode = cm[1];
        if (/\bPTO\b/i.test(samFile)) ptoFlag = 'PTO';
      }
      if (!cabCode && expectedCabs.size) cabCode = [...expectedCabs].sort()[0];
      if (!ptoFlag && isPto) ptoFlag = 'PTO';
      if (!vehicle) {
        const mu = s(modelRaw).toUpperCase();
        const vk = rules.vehicle_keywords || {};
        for (const veh of Object.keys(vk)) {
          if ((vk[veh] || []).some((k) => mu.indexOf(k) !== -1)) { vehicle = veh; break; }
        }
      }

      let samStatus;
      if (!samFile) samStatus = 'No SAM';
      else if (onlyS.length || onlyW.length || paintMismatch || tyreMismatch) samStatus = 'Mismatch';
      else samStatus = 'Match';

      let samBaumuster = samData ? s(samData.model_baumuster) : '';
      let samNow = samData ? s(samData.model_now) : '';
      if (mo) {
        if (mo.baumuster) samBaumuster = mo.baumuster;
        if (mo.now) samNow = mo.now;
      }

      const prodOut = (function () {
        if (!('Requested delivery date' in r)) return '';
        const v = r['Requested delivery date'];
        const d = toDate(v);
        return d ? fmtDate(d) : s(v);
      })();

      const rowDict = {
        'Commission no.': com,
        'Baumuster': ('Baumuster' in r) ? r['Baumuster'] : '',
        'Model(WINGS)': strip(('Model' in r) ? r['Model'] : modelRaw).replace(/DNA$/, ''),
        'Vehicle': vehicle,
        'Category': rowCat,
        'Type': axleType,
        'Cab': cabCode,
        'PTO': ptoFlag,
        'SAM Baumuster': samBaumuster,
        'SAM now': samNow,
        'Changeability Date': '',
        'Until Dealine': '',
        'Production date': prodOut,
        'Only_in_SAM': onlyS.join(','),
        'Only_in_WINGS': samCodes.size ? onlyW.join(',') : '',
        'Factory Control Codes': exceptRow.join(','),
        'Mandatory Codes': mandRow.join(','),
        '_all_wings_codes': sortedArr(wingsCodes).join(','),
        '_all_sam_codes': sortedArr(samCodes).join(','),
        '_paint_wings': sortedArr(wingsPaint).join(','),
        '_paint_sam': sortedArr(samPaint).join(','),
        '_tyre_wings': sortedArr(wingsTyre).join(','),
        '_tyre_sam': sortedArr(samTyre).join(','),
        'Compared SAM file name': samFile,
        'SAM Status': samStatus,
      };

      // Changeability + D-day
      const changeRaw = r['Vehicle alterable until'];
      let changeDisplay = '', daysLeft = '';
      if (changeRaw !== null && changeRaw !== undefined && changeRaw !== '') {
        const cdt = toDate(changeRaw);
        if (cdt) {
          daysLeft = Math.round((dateOnlyUTC(cdt) - todayUTC) / 86400000);
          changeDisplay = fmtDate(cdt);
        } else {
          const t = strip(changeRaw);
          if (t) {
            if (t.toLowerCase() === 'done' || t.toLowerCase() === 'passed') {
              changeDisplay = daysLeft = 'Passed';
            } else changeDisplay = t;
          }
        }
      }
      rowDict['Changeability Date'] = changeDisplay;
      rowDict['Until Dealine'] = daysLeft;

      for (const col of ['Order status financial', 'Order status logistical',
        'Additional equipment (enumeration)', 'FIN', 'Subcategory (ID)',
        'Requested delivery date']) {
        if (col in r) rowDict[col] = r[col];
      }
      out.push(rowDict);
    }
    return out;
  }

  const api = { compare: compare, _toDate: toDate, _fmtDate: fmtDate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Compare = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
