// backend/compare.py 의 JS 포팅 — WINGS 행 ↔ 월별 SAM 맵 비교.
// 결과 행의 키/값은 Python 판과 동일해야 한다(대시보드가 그대로 읽는다).
(function (root) {
  'use strict';

  const isNode = (typeof module !== 'undefined' && module.exports);
  const SamParse = isNode ? require('./samparse.js') : root.SamParse;
  const RefData = isNode ? require('./refdata.js') : root.RefData;

  const normalizeModel = SamParse.normalizeModel;

  // factory-control-codes.xlsx 를 못 읽었을 때 쓰는 개별코드 기본값(워크북의 초기 5행).
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
    // Factory Control 판정은 factory-control-codes.xlsx 를 그대로 따른다(코드 관리에서
    // 추가/삭제한 개별코드가 바로 반영되도록). 워크북이 없으면 아래 기본값.
    const fc = ref.factoryControl || {};
    const fcPrefixes = (fc.prefixes && fc.prefixes.size) ? fc.prefixes : new Set(['I', 'O', 'Z', 'U']);
    const fcCodes = fc.codes || EXTRA_EXCEPT;
    function isFc(c) {
      return !!c && (fcCodes.has(c) || fcPrefixes.has(c[0]));
    }

    const manualNorm = {};
    for (const k of Object.keys(rules.manual_map || {})) manualNorm[normalizeModel(k)] = rules.manual_map[k];

    const sortedYyyymm = Object.keys(samMapsByMonth).map(Number).sort((a, b) => a - b);

    // yyyymm(=년*100+월)에서 k개월 이전의 yyyymm.
    function ymBack(ym, k) {
      const idx = Math.floor(ym / 100) * 12 + (ym % 100 - 1) - k;
      return Math.floor(idx / 12) * 100 + (idx % 12 + 1);
    }

    const out = [];
    for (const r of wingsRows) {
      const prodRaw = ('Requested delivery date' in r) ? r['Requested delivery date'] : '';

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

      function entryForModel(map, pto) {
        const cand = map.get(modelNorm);
        if (candidates(cand, pto).length) return cand;
        const [numNorm, sufNorm] = splitModel(modelNorm);
        for (const [k, v] of map) {
          const kNorm = normalizeModel(String(k));
          const [numK, sufK] = splitModel(kNorm);
          if (kNorm === modelNorm || (numK === numNorm && sufK === sufNorm)) {
            if (candidates(v, pto).length) return v;
          }
        }
        return null;
      }

      const mo = manualNorm[modelNorm];

      // 우선순위 월 맵 목록에서 모델을 해석한다. { data, pto } 반환(없으면 data=null).
      function resolveIn(mapList) {
        let localPto = isPto;
        let entry = null, fbEntry = null, ok = false;
        for (const tryMap of mapList) {
          const e = entryForModel(tryMap, localPto);
          if (!candidates(e, localPto).length) continue;
          if (fbEntry === null) fbEntry = e;
          if (!expectedCabs.size) { entry = e; ok = true; break; }
          if (cabOk(pick(e, localPto, wingsBm, wingsSub, expectedCabs), expectedCabs)) { entry = e; ok = true; break; }
        }
        if (!ok && fbEntry !== null) entry = fbEntry;
        // PTO 보정: WINGS 코드에 PTO 변형에만 있는 코드가 있으면 PTO 로 본다.
        if (!localPto && entry && entry['true'] && entry['false']) {
          const ptoData = pick(entry, true, wingsBm, wingsSub);
          const nptoData = pick(entry, false, wingsBm, wingsSub);
          if (ptoData && nptoData) {
            const ptoUnique = setDiff(ptoData.codes, nptoData.codes);
            if (setInter(wingsCodes, ptoUnique).size) localPto = true;
          }
        }
        let data = pick(entry, localPto, wingsBm, wingsSub, expectedCabs);
        if (mo && mo.file) {
          const pinned = findSamDataByFile(mapList, localPto, mo.file);
          if (pinned) data = pinned;
        }
        return { data: data, pto: localPto };
      }

      // 생산월(prodYm) 폴더를 먼저 보고, 없으면 이전 최대 6개월 폴더로 fallback.
      const prodYm = (function () {
        const d = toDate(prodRaw);
        return d ? d.getUTCFullYear() * 100 + (d.getUTCMonth() + 1) : 0;
      })();
      let currentMaps, prevMaps;
      if (prodYm) {
        currentMaps = samMapsByMonth[prodYm] ? [samMapsByMonth[prodYm]] : [];
        prevMaps = [];
        for (let k = 1; k <= 6; k++) {
          const pm = ymBack(prodYm, k);
          if (samMapsByMonth[pm]) prevMaps.push(samMapsByMonth[pm]);
        }
      } else {
        // 생산월을 알 수 없으면 예전처럼 최신월 우선으로만 비교(‘SAM update요청’ 판정 없음).
        currentMaps = sortedYyyymm.slice().reverse().map((m) => samMapsByMonth[m]);
        prevMaps = [];
      }

      let res = resolveIn(currentMaps);
      let matchSource = 'current';       // 'current' | 'prev' | 'none'
      if (!res.data || !res.data.file) {
        const resPrev = resolveIn(prevMaps);
        if (resPrev.data && resPrev.data.file) { res = resPrev; matchSource = 'prev'; }
        else matchSource = 'none';
      }
      isPto = res.pto;
      let samData = res.data;

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

      // 상태는 항상 Match / Mismatch / No SAM 중 하나 → 전체 = 매치 + 미스매치 + No SAM.
      // 이전(≤6개월) 대체 SAM 으로 매칭된 경우도 그 대체본 기준으로 Match/Mismatch 를 계산하고,
      // 별도 플래그 'SAM Update' 로만 표시한다(총계와 무관하게 겹쳐서 카운트).
      let samStatus;
      if (matchSource === 'none' || !samFile) samStatus = 'No SAM';
      else if (onlyS.length || onlyW.length || paintMismatch || tyreMismatch) samStatus = 'Mismatch';
      else samStatus = 'Match';
      // SAM update 필요 여부: 매칭된 SAM 이 담긴 '폴더(=생산월)'가 이 행의 생산월과 다르면
      // (= 생산월 폴더엔 없고 이전 월 대체본으로 매칭됨) 요청. 상태(Match/Mismatch)와 무관하게 병기.
      const samUpdate = (matchSource === 'prev' && samFile) ? 'Y' : '';

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

      // No SAM: 비교 대상이 없으므로 차이/Mandatory 계산 결과를 남기지 않는다.
      const noSam = (matchSource === 'none' || !samFile);

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
        'Only_in_SAM': noSam ? '' : onlyS.join(','),
        'Only_in_WINGS': (noSam || !samCodes.size) ? '' : onlyW.join(','),
        'Factory Control Codes': noSam ? '' : exceptRow.join(','),
        'Mandatory Codes': noSam ? '' : mandRow.join(','),
        '_all_wings_codes': sortedArr(wingsCodes).join(','),
        '_all_sam_codes': sortedArr(samCodes).join(','),
        '_paint_wings': sortedArr(wingsPaint).join(','),
        '_paint_sam': sortedArr(samPaint).join(','),
        '_tyre_wings': sortedArr(wingsTyre).join(','),
        '_tyre_sam': sortedArr(samTyre).join(','),
        'Compared SAM file name': samFile,
        'SAM Status': samStatus,
        'SAM Update': samUpdate,
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
