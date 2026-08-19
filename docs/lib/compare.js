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
  // 접두어 규칙에서 빼는 코드의 기본값(워크북의 '예외' 행이 없을 때).
  const FC_EXCEPTS = new Set(['Z5M', 'Z5U']);

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

  // ---- 캡·PTO 는 파일명이 아니라 '실제 코드'로 판정한다 ----
  // 캡 : cab.xlsx 의 { WINGS 캡 코드 → SAM 캡 variant }  (F1J → G5F)
  // PTO: 코드 설명에 PTO 가 들어간 코드
  // 둘 다 compare() 시작에서 ref 로 채운다(후보 선택 헬퍼가 모듈 스코프라 여기 둔다).
  let CAB_MAP = {};
  let CODE_DESC = {};
  function cabsIn(codes) {
    const out = new Set();
    for (const c of (codes || [])) if (CAB_MAP[c]) out.add(CAB_MAP[c]);
    return out;
  }
  // 코드 설명이 PTO 옵션을 가리키는가.
  //   'PTO MB…' · 'EnginePTO…'  → 대문자 PTO
  //   'Pto parameterised…'      → 낱말 PTO(대소문자 무시)  — 'adaptor' 는 걸리지 않는다
  //   'Power take-off…'         → 풀어 쓴 표기 (Z5U 처럼 PTO 라는 글자가 없다)
  function isPtoDesc(desc) {
    const d = s(desc);
    return d.indexOf('PTO') !== -1
      || /(?<![A-Za-z])pto(?![A-Za-z])/i.test(d)
      || /power[ -]*take[ -]*off/i.test(d);
  }
  function ptoCodesIn(codes) {
    const out = new Set();
    for (const c of (codes || [])) if (isPtoDesc(CODE_DESC[c])) out.add(c);
    return out;
  }
  // SAM 후보의 캡: 문서의 실제 코드가 1순위, 없으면 파일명 토큰.
  function dataCabMatch(data, expectedCabs) {
    if (!data || !expectedCabs || !expectedCabs.size) return false;
    const cabs = cabsIn(data.codes);
    for (const cv of expectedCabs) if (cv && cabs.has(cv)) return true;
    if (cabs.size) return false;                 // 코드가 있는데 안 맞으면 파일명은 안 본다
    const f = s(data.file).toUpperCase();
    for (const cv of expectedCabs) if (cv && f.indexOf(cv) !== -1) return true;
    return false;
  }

  // ---- 후보 선택 헬퍼 (_candidates / _match_score / _pick / _cab_ok) ----
  // PTO 는 엄격 필터다 — 요청한 PTO 변형이 없으면 후보 없음(= No SAM). 예전처럼 반대
  // 변형으로 넘어가면 PTO 차량이 비PTO SAM 과 조용히 비교돼 차이가 통째로 묻힌다.
  function candidates(entry, preferPto) {
    if (!entry || typeof entry !== 'object') return [];
    const lst = entry[preferPto ? 'true' : 'false'];
    if (!lst) return [];
    return Array.isArray(lst) ? lst : [lst];
  }
  function matchScore(data, wingsBm, wingsSub, expectedCabs) {
    let score = 0;
    if (expectedCabs && expectedCabs.size && dataCabMatch(data, expectedCabs)) score += 3;
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
  function cabOk(data, expectedCabs) { return dataCabMatch(data, expectedCabs); }
  function splitModel(str) {
    const m = /^(\d+)([A-Z]*)$/.exec(str);
    return m ? [m[1], m[2]] : [str, ''];
  }

  /**
   * WINGS 모델키 하나로 SAM 쪽에서 찾아볼 키 후보들을 만든다.
   * 첫 번째는 언제나 원래 키라, 규칙은 '원래 키로 못 찾았을 때만' 개입한다(기존 결과 보존).
   * model_mapping.xlsx 의 정규화_과거번호 · 이전모델 · 현재모델 · 모델별칭 · 옵션 시트가
   * 여기서 처음으로 실제 매칭에 쓰인다 — 시트를 고치면 매칭이 바뀐다.
   */
  function modelKeyCandidates(modelNorm, rules) {
    const [num, suf] = splitModel(modelNorm);
    const out = [modelNorm];
    if (!/^\d+$/.test(num)) return out;
    const add = (n) => {
      const k = String(n || '').toUpperCase().replace(/[^A-Z0-9]/g, '') + suf;
      if (n && k !== suf && out.indexOf(k) === -1) out.push(k);
    };
    // 양방향으로 본다 — 규칙은 'A→B' 한 줄만 적어도 B→A 로도 찾아야 하기 때문.
    const both = (obj) => {
      const o = obj || {};
      add(o[num]);
      for (const k of Object.keys(o)) if (String(o[k]) === num) add(k);
    };
    // 같은 차를 가리키는 표만 쓴다. 이전모델/현재모델 은 '세대 이웃'이라
    // (2853 과 2863 은 다른 차) 키로 쓰면 엉뚱한 모델과 조용히 비교된다.
    both(rules.normalize_historic);
    const rev = rules.reverse_aliases || {};
    for (const k of Object.keys(rev)) {
      const list = rev[k] || [];
      if (k === num) list.forEach(add);
      else if (list.indexOf(num) !== -1) add(k);
    }
    if (rules.normalize_28xx_to_26xx && num.length === 4) {
      if (num.charAt(1) === '8') add(num.charAt(0) + '6' + num.slice(2));
      if (num.charAt(1) === '6') add(num.charAt(0) + '8' + num.slice(2));
    }
    return out;
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
    CAB_MAP = cabMap;
    CODE_DESC = codeDesc;
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
    // '예외' 로 적어 둔 코드는 접두어에 걸려도 일반 코드로 본다(예: Z5M·Z5U 는 PTO 옵션).
    const fcExcepts = fc.excepts || FC_EXCEPTS;
    function isFc(c) {
      if (!c || fcExcepts.has(c)) return false;
      return fcCodes.has(c) || fcPrefixes.has(c[0]);
    }

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
      const keyCands = modelKeyCandidates(modelNorm, rules);

      const wingsBm = strip(r['Baumuster']);
      const wingsSub = strip(r['Subcategory (ID)']).toLowerCase();

      // WINGS 쪽 캡·PTO — 주문에 실제로 걸린 코드에서 찾는다.
      const expectedCabs = cabsIn(wingsCodes);
      const wingsPtoCodes = ptoCodesIn(wingsCodes);
      let isPto = wingsPtoCodes.size > 0 || !!r['WINGS_has_pto'];

      function entryForKey(map, key, pto) {
        const cand = map.get(key);
        if (candidates(cand, pto).length) return cand;
        const [numNorm, sufNorm] = splitModel(key);
        for (const [k, v] of map) {
          const kNorm = normalizeModel(String(k));
          const [numK, sufK] = splitModel(kNorm);
          if (kNorm === key || (numK === numNorm && sufK === sufNorm)) {
            if (candidates(v, pto).length) return v;
          }
        }
        return null;
      }

      // 우선순위 월 맵 목록에서 모델키 하나를 해석한다. { data, pto } 반환(없으면 data=null).
      function resolveIn(mapList, key) {
        let localPto = isPto;
        let entry = null, fbEntry = null, ok = false;
        for (const tryMap of mapList) {
          const e = entryForKey(tryMap, key, localPto);
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
        const data = pick(entry, localPto, wingsBm, wingsSub, expectedCabs);
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

      // 모델키가 바깥 루프다 — 정확한 번호를 생산월·이전월까지 다 찾아본 뒤에야
      // 별칭 번호로 넘어간다. 반대로 하면 가까운 달의 별칭 파일이 제 모델을 이긴다.
      let res = { data: null, pto: isPto };
      let matchSource = 'none';          // 'current' | 'prev' | 'none'
      for (const key of keyCands) {
        const rc = resolveIn(currentMaps, key);
        if (rc.data && rc.data.file) { res = rc; matchSource = 'current'; break; }
        const rp = resolveIn(prevMaps, key);
        if (rp.data && rp.data.file) { res = rp; matchSource = 'prev'; break; }
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

      // 차종·축은 문서에 따로 없어 SAM 파일명에서 뽑는다. 캡·PTO 는 양쪽 '코드'로 판정하고
      // (파일명은 SAM 쪽 코드가 없을 때만 폴백) 두 값을 다 남겨 화면에서 불일치를 표시한다.
      let vehicle = '', axleType = '', fileCab = '', filePto = false;
      if (samFile) {
        // \b 는 밑줄도 단어문자로 보므로 'quotation_Actros'·'4x2_M&M' 에서는 경계가 생기지 않아
        // 차종·축이 통째로 비었다. 파일명은 _ 로 토막나 있으니 앞뒤를 직접 본다.
        const vm = /(?<![A-Za-z0-9])(Actros-L|Actros|Arocs|Atego|eActros|Econic|Unimog)(?![A-Za-z])/i.exec(samFile);
        if (vm) vehicle = vm[1];
        const am = /(?<![0-9])(\d+x\d+)(?![0-9])/i.exec(samFile);
        if (am) axleType = am[1];
        const cm = /(?<![A-Za-z0-9])([A-Z]\d[A-Z])(?![A-Za-z0-9])/.exec(samFile);
        if (cm) fileCab = cm[1];
        if (/(?<![A-Za-z])PTO(?![A-Za-z])/i.test(samFile)) filePto = true;
      }
      const samCabs = samCodes.size ? cabsIn(samCodes) : new Set();
      if (!samCabs.size && fileCab) samCabs.add(fileCab);
      const samPtoCodes = ptoCodesIn(samCodes);
      const samPto = samPtoCodes.size > 0 || filePto;
      // 목록에 한 값만 보여야 하므로 WINGS 기준을 쓰고, WINGS 에 신호가 없으면 SAM 값을 쓴다.
      const cabCode = sortedArr(expectedCabs)[0] || sortedArr(samCabs)[0] || '';
      const ptoFlag = (isPto || samPto) ? 'PTO' : '';
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

      // SAM 번호는 매칭된 파일에서 그대로 온다(수동 덮어쓰기 없음).
      const samBaumuster = samData ? s(samData.model_baumuster) : '';
      const samNow = samData ? s(samData.model_now) : '';

      const prodOut = (function () {
        if (!('Requested delivery date' in r)) return '';
        const v = r['Requested delivery date'];
        const d = toDate(v);
        return d ? fmtDate(d) : s(v);
      })();

      // No SAM: 비교 대상이 없으므로 차이/Mandatory 계산 결과를 남기지 않는다.
      const noSam = (matchSource === 'none' || !samFile);

      // WINGS 표기에 축이 섞여 들어오기도 한다('4153 K' 와 '4153 K 8x4' 가 같은 차).
      // 축은 축 칸으로 옮기고 표기는 통일해야 같은 모델이 두 줄로 갈라지지 않는다.
      let modelDisplay = strip(('Model' in r) ? r['Model'] : modelRaw).replace(/DNA$/, '');
      const inlineAxle = /(?<![0-9])(\d+x\d+)(?![0-9])/i.exec(modelDisplay);
      if (inlineAxle) {
        modelDisplay = strip(modelDisplay.replace(inlineAxle[0], '').replace(/\s{2,}/g, ' '));
        if (!axleType) axleType = inlineAxle[1];
      }
      // WINGS표시치환 시트 — 같은 차가 WINGS 표기 차이로 두 모델처럼 갈라지는 걸 막는다.
      const disp = rules.wings_display_replace || {};
      if (Object.prototype.hasOwnProperty.call(disp, modelDisplay) && disp[modelDisplay]) {
        modelDisplay = String(disp[modelDisplay]);
      }

      const rowDict = {
        'Commission no.': com,
        'Baumuster': ('Baumuster' in r) ? r['Baumuster'] : '',
        'Model(WINGS)': modelDisplay,
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
        // 캡은 Paint/Tyre 처럼 양쪽 값을 남긴다 — 화면에서 불일치를 표시한다.
        // PTO 는 따로 남기지 않는다: PTO 코드(N1G·Z5M…) 차이는 일반 코드 비교에 그대로 나온다.
        '_cab_wings': sortedArr(expectedCabs).join(','),
        '_cab_sam': sortedArr(samCabs).join(','),
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
    // 축·차종은 SAM 파일명(또는 차종키워드)에서만 나오므로 매칭이 없는 행에서는 비어 버린다.
    // 같은 Baumuster(공장코드)는 같은 사양이니, 그 값을 아는 다른 행에서 옮겨 채운다.
    fillFromBaumuster(out, 'Type');
    fillFromBaumuster(out, 'Vehicle');
    return out;
  }

  // 같은 Baumuster 안에서 값이 하나로 모일 때만 채운다(갈리면 그대로 비워 둔다).
  function fillFromBaumuster(rows, col) {
    const known = new Map();
    for (const r of rows) {
      const bm = strip(r['Baumuster']), v = strip(r[col]);
      if (!bm || !v) continue;
      if (!known.has(bm)) known.set(bm, v);
      else if (known.get(bm) !== v) known.set(bm, '');
    }
    for (const r of rows) {
      if (strip(r[col])) continue;
      const v = known.get(strip(r['Baumuster']));
      if (v) r[col] = v;
    }
  }

  const api = { compare: compare, _toDate: toDate, _fmtDate: fmtDate,
                _modelKeyCandidates: modelKeyCandidates };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Compare = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
