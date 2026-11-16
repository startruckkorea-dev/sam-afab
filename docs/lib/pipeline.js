// backend/build_data.py 의 JS 포팅 — 오케스트레이터.
// SharePoint(또는 로컬 폴더)에서 참조 워크북 · 최신 WINGS · 최신 생산월 SAM 을 읽어
// 파싱 → 비교까지 돌리고, 대시보드가 그대로 읽는 data.json / codes.json 객체를 만든다.
//
// 소스는 어댑터로 주입한다(브라우저=Graph, Node 테스트=로컬 파일):
//   src.list(folderKey)              -> [{name, isFolder, size, modified}]
//   src.download(folderKey, name)    -> ArrayBuffer | Uint8Array
//   src.optionCodes()                -> { code: description }   (option_codes.py 사본)
// folderKey 는 'sam' | 'wings' | 'model_rules' | 'code' 와 'sam/<하위폴더>' 형태.
(function (root) {
  'use strict';

  const isNode = (typeof module !== 'undefined' && module.exports);
  const SamParse = isNode ? require('./samparse.js') : root.SamParse;
  const WingsParse = isNode ? require('./wingsparse.js') : root.WingsParse;
  const RefData = isNode ? require('./refdata.js') : root.RefData;
  const Compare = isNode ? require('./compare.js') : root.Compare;

  // build_data.DISPLAY_COLS 와 동일한 순서 — 프런트가 이 순서를 그대로 쓴다.
  const DISPLAY_COLS = [
    'Commission no.', 'Baumuster', 'Model(WINGS)', 'Vehicle', 'Category', 'Type', 'Cab', 'PTO',
    'SAM Baumuster', 'SAM now', 'Changeability Date', 'Until Dealine',
    'Production date', 'Only_in_SAM', 'Only_in_WINGS', 'Factory Control Codes',
    'Mandatory Codes', 'Order status financial', 'Order status logistical',
    'FIN', 'Subcategory (ID)', 'Compared SAM file name', 'SAM Status', 'SAM Update',
    '_all_wings_codes', '_all_sam_codes',
    '_cab_wings', '_cab_sam',
    '_paint_wings', '_paint_sam', '_tyre_wings', '_tyre_sam',
  ];

  // 참조 워크북 파일명(04. code / 03. model_rules). 대소문자 무시로 찾는다.
  const REF_FILES = {
    codeDict: 'mbtruck-spec-data.xlsx',
    mandatory: 'mandatory-codes.xlsx',
    modelCategory: 'model-category.xlsx',
    cab: 'cab.xlsx',
    factoryControl: 'factory-control-codes.xlsx',
    ptoCodes: 'pto-codes.xlsx',
    rules: 'model_mapping.xlsx',
  };

  const WINGS_EXTS = ['.xlsx', '.xls', '.csv'];
  const MONTH_RE = /^(\d{4})[_-](\d{2})\b/;

  function ext(name) {
    const i = String(name).lastIndexOf('.');
    return i < 0 ? '' : String(name).slice(i).toLowerCase();
  }

  function findFile(items, filename) {
    const want = String(filename).toLowerCase();
    for (const it of items) {
      if (!it.isFolder && String(it.name).toLowerCase() === want) return it.name;
    }
    return null;
  }

  // sharepoint._recency 포팅: 파일명의 13자리 epoch-ms 또는 YYYY-MM-DD 를 우선 사용.
  function recency(name) {
    let m = /(\d{13})/.exec(name);
    if (m) return Number(m[1]) / 1000;
    m = /(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})/.exec(name);
    if (m) {
      const t = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      if (!isNaN(t)) return t / 1000;
    }
    return 0;
  }

  function pickLatestWings(items) {
    const files = items.filter((it) => !it.isFolder
      && WINGS_EXTS.indexOf(ext(it.name)) !== -1
      && it.name.indexOf('~$') !== 0);
    if (!files.length) throw new Error('02. WINGS_data 에 WINGS 파일이 없습니다.');
    let best = files[0];
    for (const it of files.slice(1)) {
      const a = [recency(it.name), it.modified || ''];
      const b = [recency(best.name), best.modified || ''];
      if (a[0] > b[0] || (a[0] === b[0] && a[1] > b[1])) best = it;
    }
    return { chosen: best, count: files.length };
  }

  // 01. SAM_files 의 'YYYY-MM …' 하위폴더들 → [{yyyymm, name}] (오름차순)
  function samMonthFolders(items) {
    const out = [];
    for (const it of items) {
      if (!it.isFolder) continue;
      const m = MONTH_RE.exec(it.name);
      if (m) out.push({ yyyymm: Number(m[1]) * 100 + Number(m[2]), name: it.name });
    }
    out.sort((a, b) => a.yyyymm - b.yyyymm);
    return out;
  }

  /**
   * 참조 데이터(04. code + 03. model_rules + option_codes)를 모두 읽는다.
   * @returns {{rules, mandatory, modelCategory, cabMap, codeDesc, codeDict}}
   */
  async function loadRefData(src, log) {
    const optionCodes = await src.optionCodes();
    const codeItems = await src.list('code');

    async function tryLoad(key, loader, fallback) {
      const name = findFile(codeItems, REF_FILES[key]);
      if (!name) {
        log(`[ref] ${REF_FILES[key]} 없음 — 기본값 사용`);
        return fallback;
      }
      try {
        return loader(await src.download('code', name));
      } catch (e) {
        log(`[ref] ${name} 읽기 실패(${String(e.message).slice(0, 80)}) — 기본값 사용`);
        return fallback;
      }
    }

    // mandatory: 워크북이 없으면 빈 집합(파이썬은 option_codes 폴백이지만,
    // 워크북이 소스 오브 트루스이므로 없으면 경고 후 빈 값으로 진행).
    // 워크북들은 서로 독립이므로 동시에 내려받는다.
    async function loadRules() {
      try {
        const items = await src.list('model_rules');
        const name = findFile(items, REF_FILES.rules);
        const r = RefData.loadRules(name ? await src.download('model_rules', name) : null);
        log(`[ref] rules: ${name || '기본값(model_mapping.xlsx 없음)'}`);
        return { rules: r, file: name || '' };
      } catch (e) {
        log(`[ref] model_mapping.xlsx 실패(${String(e.message).slice(0, 80)}) — 기본 규칙 사용`);
        return { rules: RefData.loadRules(null), file: '' };
      }
    }
    const loaded = await Promise.all([
      tryLoad('mandatory', RefData.loadMandatory,
        { desc: {}, set: new Set(), groups: {}, cats: {} }),
      tryLoad('modelCategory', RefData.loadModelCategory, {}),
      tryLoad('cab', RefData.loadCabMap, {}),
      tryLoad('factoryControl', RefData.loadFactoryControl, RefData.loadFactoryControl(null)),
      tryLoad('ptoCodes', RefData.loadPtoCodes, RefData.loadPtoCodes(null)),
      tryLoad('codeDict', (buf) => RefData.loadCodeDict(buf, 'code_dict'), null),
      loadRules(),
    ]);
    const mandatory = loaded[0], modelCategory = loaded[1], cabMap = loaded[2];
    const factoryControl = loaded[3], ptoCodes = loaded[4], codeDict = loaded[5];
    const rules = loaded[6].rules, rulesFile = loaded[6].file;

    log(`[ref] option codes ${Object.keys(optionCodes).length} · mandatory ${mandatory.set.size}`
      + ` · category ${Object.keys(modelCategory).length} · cab ${Object.keys(cabMap).length}`
      + ` · factory control ${factoryControl.prefixes.size}접두어/${factoryControl.codes.size}코드`
      + ` · PTO 코드 ${ptoCodes.codes.size}개(제외 ${ptoCodes.excepts.size})`);

    return {
      rules: rules,
      rulesFile: rulesFile,
      mandatory: mandatory,
      modelCategory: modelCategory,
      cabMap: cabMap,
      factoryControl: factoryControl,
      ptoCodes: ptoCodes,
      codeDesc: optionCodes,                       // 파서/비교는 option_codes 기준
      codeDict: (codeDict && Object.keys(codeDict).length) ? codeDict : optionCodes,
    };
  }

  /** 최신 WINGS 파일 하나를 읽어 정규화된 행 배열로. */
  async function loadWings(src, log) {
    const items = await src.list('wings');
    const { chosen, count } = pickLatestWings(items);
    if (count > 1) log(`[wings] ${count}개 중 최신 선택: ${chosen.name}`);
    const buf = await src.download('wings', chosen.name);
    const rows = WingsParse.parseWings(buf, chosen.name);
    log(`[wings] ${chosen.name} — ${rows.length} 행`);
    return { rows: rows, file: chosen.name };
  }

  // 01. SAM_files 아래의 생산월 폴더를 모은다.
  // 구조: 'MYxx/YYYY-MM …'(연식 폴더 하위) 또는 'YYYY-MM …'(직속, 구 구조) 모두 지원.
  // 반환: [{ yyyymm, name(표시용), path(folderKey 접미사) }] (yyyymm 오름차순)
  async function collectSamMonths(src) {
    const top = await src.list('sam');
    const out = [];
    // (구) 01. SAM_files 바로 아래의 YYYY-MM 폴더
    for (const m of samMonthFolders(top)) {
      out.push({ yyyymm: m.yyyymm, name: m.name, path: m.name });
    }
    // (신) MYxx 연식 폴더 → 그 하위의 YYYY-MM 폴더 (연식 폴더 목록은 동시에 조회)
    const myFolders = top.filter((it) => it.isFolder && /^MY\d{2}\b/i.test(it.name));
    const subLists = await Promise.all(myFolders.map((my) => src.list('sam/' + my.name)));
    myFolders.forEach(function (my, i) {
      for (const m of samMonthFolders(subLists[i])) {
        out.push({ yyyymm: m.yyyymm, name: my.name + '/' + m.name, path: my.name + '/' + m.name });
      }
    });
    out.sort((a, b) => a.yyyymm - b.yyyymm);
    return out;
  }

  // 동시 실행 개수를 limit 으로 묶어 fn 을 돌린다. 결과는 입력 순서대로.
  // SharePoint 는 동시 요청이 너무 많으면 429 로 제한하므로 6~8 정도가 적당하다.
  async function pool(items, limit, fn) {
    const out = new Array(items.length);
    let next = 0;
    async function worker() {
      while (next < items.length) {
        const i = next++;
        out[i] = await fn(items[i], i);
      }
    }
    const n = Math.max(1, Math.min(limit, items.length));
    await Promise.all(Array.from({ length: n }, worker));
    return out;
  }

  const SAM_CONCURRENCY = 8;
  const SAM_CACHE_FILE = 'sam-cache.json';

  // 파일이 바뀌지 않았는지 판정하는 지문 — 크기 + 수정일시 + 파서 버전.
  function samFingerprint(f) {
    return [f.size || 0, f.modified || '', SamParse.PARSER_VERSION].join('|');
  }

  /** 생산월별 SAM 매핑. 존재하는 모든 생산월 폴더를 읽는다. */
  async function loadSam(src, ref, allMonths, log, diag) {
    const months = await collectSamMonths(src);
    if (!months.length) {
      throw new Error('01. SAM_files 에 YYYY-MM 하위폴더가 없습니다. '
        + '(MYxx 연식 폴더 하위의 YYYY-MM 또는 직속 YYYY-MM 폴더가 필요)');
    }
    // 각 행의 생산월 폴더 + 이전 6개월로 'SAM update요청'을 판정하므로, 미래 폴더가
    // 있어도 이전 달이 누락되지 않도록 존재하는 모든 생산월 폴더를 읽는다.
    const targets = months;
    // 진단용 수집 — 어떤 파일이 왜 빠졌는지는 data.json(diagnostics)에 담아 모델 매칭 > 문제 탭이 보여 준다.
    const skipped = (diag && diag.samSkipped) || [];
    const monthStats = (diag && diag.samMonths) || [];
    log(`[sam] 생산월 폴더 ${months.length}개 사용: ${targets.map((m) => m.name).join(', ')}`);

    // 숫자가 없는 코드(JKTO 등)는 '아는 코드' 일 때만 살아남으므로, option_codes 뿐 아니라
    // mbtruck-spec-data.xlsx 의 코드 사전까지 합쳐야 SAM 장비표에서 누락되지 않는다.
    const knownCodes = new Set(Object.keys(ref.codeDesc));
    for (const c of Object.keys(ref.codeDict || {})) knownCodes.add(c);
    for (const c of ref.mandatory.set) knownCodes.add(c);
    const ctx = { codeDesc: ref.codeDesc, knownCodes: knownCodes, ptoCodes: ref.ptoCodes };

    const now = () => ((typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now());
    const tList = now();

    // 1) 생산월 폴더의 파일 목록을 동시에 조회
    const fileLists = await pool(targets, SAM_CONCURRENCY, async function (month) {
      return (await src.list('sam/' + month.path))
        .filter((it) => !it.isFolder && ext(it.name) === '.docx' && it.name.indexOf('.') !== 0)
        .sort((a, b) => (a.name < b.name ? -1 : (a.name > b.name ? 1 : 0)));
    });

    // 2) 이전 빌드의 추출 결과 캐시. 크기·수정일시·파서 버전이 같은 파일은 다시 받지 않는다.
    //    캐시에는 참조 데이터와 무관한 추출 결과(extractSamDocx)만 있으므로, 코드 DB 나
    //    PTO 표가 바뀌어도 finalizeSam 을 매번 다시 적용해 결과가 정확하다.
    let cache = {};
    if (src.readCache) {
      try {
        const c = await src.readCache(SAM_CACHE_FILE);
        if (c && c.files && typeof c.files === 'object') cache = c.files;
      } catch (e) {
        log(`[sam] 캐시 없음/읽기 실패 — 전체 다운로드 (${String(e.message).slice(0, 60)})`);
      }
    }

    const jobs = [];
    targets.forEach(function (month, mi) {
      for (const f of fileLists[mi]) jobs.push({ month: month, mi: mi, f: f, key: month.path + '/' + f.name });
    });
    const tDl = now();
    let hits = 0, fetched = 0;
    // 3) 캐시에 없는 파일만 동시에 내려받아 추출
    const results = await pool(jobs, SAM_CONCURRENCY, async function (job) {
      const fp = samFingerprint(job.f);
      const hit = cache[job.key];
      if (hit && hit.fp === fp && hit.raw && hit.raw.v === SamParse.PARSER_VERSION) {
        hits += 1;
        return { raw: hit.raw, fp: fp };
      }
      try {
        const raw = await SamParse.extractSamDocx(
          await src.download('sam/' + job.month.path, job.f.name), job.f.name);
        fetched += 1;
        return { raw: raw, fp: fp };
      } catch (e) {
        return { error: e };
      }
    });
    const tDone = now();
    log(`[sam] 파일 ${jobs.length}개 — 캐시 ${hits} · 새로 받음 ${fetched}`
      + ` (목록 ${((tDl - tList) / 1000).toFixed(1)}s · 다운로드/추출 ${((tDone - tDl) / 1000).toFixed(1)}s)`);

    // 4) 캐시 갱신 — 지금 폴더에 있는 파일만 남긴다(지워진 파일은 정리).
    const newCache = {};
    let changed = fetched > 0;
    jobs.forEach(function (job, i) {
      const r = results[i];
      if (r && r.raw) newCache[job.key] = { fp: r.fp, raw: r.raw };
    });
    if (!changed && Object.keys(newCache).length !== Object.keys(cache).length) changed = true;
    if (changed && src.writeCache) {
      try {
        await src.writeCache(SAM_CACHE_FILE,
          { parser_version: SamParse.PARSER_VERSION, files: newCache });
      } catch (e) {
        log(`[sam] 캐시 저장 실패(다음 빌드는 전체 다운로드) — ${String(e.message).slice(0, 80)}`);
      }
    }

    // 5) 생산월별로 판정·매핑 (순서 고정 → 로그/결과가 매번 같다)
    const maps = {};
    for (let mi = 0; mi < targets.length; mi++) {
      const month = targets[mi];
      const files = fileLists[mi];
      const parsed = [];
      jobs.forEach(function (job, i) {
        if (job.mi !== mi) return;
        const f = job.f;
        const r = results[i];
        if (r.error) {
          log(`  [sam] 파싱 실패: ${f.name} — ${String(r.error.message).slice(0, 80)}`);
          skipped.push({ month: month.name, file: f.name,
            reason: '파싱 실패 — ' + String(r.error.message).slice(0, 120) });
          return;
        }
        const d = SamParse.finalizeSam(r.raw, ctx);
        if (d) parsed.push(d);
        else {
          log(`  [sam] 건너뜀(모델/코드 인식 실패): ${f.name}`);
          skipped.push({ month: month.name, file: f.name,
            reason: '모델번호 또는 장비코드를 못 찾음 (본문 Vehicle type / Equipment overview 확인)' });
        }
      });
      const mapping = SamParse.buildSamMapping(parsed, ref.rules);
      if (mapping.size) maps[month.yyyymm] = mapping;
      // 'SAM 파일 비교' 메뉴용 카탈로그 — 주문에 붙지 않은 문서까지 읽은 그대로 남긴다.
      // 코드는 비교 결과의 _all_sam_codes 와 같은 원본(finalizeSam 의 codes)이다.
      if (diag && diag.samDocs) {
        const myM = /^(MY\d{2})\b.*\//i.exec(month.path);
        const ym = String(month.yyyymm);
        for (const d of parsed) {
          const sorted = (s) => Array.from(s || []).sort().join(',');
          diag.samDocs.push({
            folder: month.name,
            my: myM ? myM[1].toUpperCase() : '',
            month: ym.slice(0, 4) + '-' + ym.slice(4, 6),
            file: d.file,
            model_now: d.model_now || '',
            model_baumuster: d.model_baumuster || '',
            pto: !!d.is_pto,
            codes: sorted(d.codes),
            paint: sorted(d.paint),
            tyre: sorted(d.tyre),
          });
        }
      }
      monthStats.push({ name: month.name, yyyymm: month.yyyymm, files: files.length,
        parsed: parsed.length, keys: Array.from(mapping.keys()).sort(),
        // 붙지 않은 문서를 설명하려면 '무엇으로 읽혔는지'가 필요하다 — 모델키와 PTO 판정.
        docs: parsed.map((d) => ({
          file: d.file,
          pto: !!d.is_pto,
          keys: [SamParse.normalizeModel(d.model_now), SamParse.normalizeModel(d.model_baumuster)]
            .filter((k, i, a) => k && a.indexOf(k) === i),
        })) });
      log(`[sam] ${month.name}: ${parsed.length}/${files.length} 파일, 모델키 ${mapping.size}개`);
    }
    if (!Object.keys(maps).length) throw new Error('SAM 파일을 하나도 읽지 못했습니다.');
    return maps;
  }

  /**
   * 전체 빌드. build_data.build() + codes.json 생성까지.
   * @param {object} src   소스 어댑터
   * @param {object} opts  { allMonths?:boolean, log?:fn, now?:Date }
   * @returns {{data:object, codes:object}}
   */
  async function build(src, opts) {
    opts = opts || {};
    const log = opts.log || function () {};
    const t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : 0;

    const diag = { samMonths: [], samSkipped: [], samDocs: [], rulesFile: '' };
    // 참조 데이터와 WINGS 는 서로 독립 — 동시에 읽는다. SAM 은 참조 데이터(코드 DB)가 필요.
    const both = await Promise.all([loadRefData(src, log), loadWings(src, log)]);
    const ref = both[0], wings = both[1];
    diag.rulesFile = ref.rulesFile || '';
    const samMaps = await loadSam(src, ref, !!opts.allMonths, log, diag);

    const rowsRaw = Compare.compare(wings.rows, samMaps, {
      rules: ref.rules,
      mandatory: ref.mandatory,
      modelCategory: ref.modelCategory,
      cabMap: ref.cabMap,
      factoryControl: ref.factoryControl,
      ptoCodes: ref.ptoCodes,
      codeDesc: ref.codeDesc,
      today: opts.now,
    });
    log(`[build] 결과 ${rowsRaw.length} 행`);

    // 읽기는 했는데 어떤 주문에도 붙지 않은 SAM 문서 — '파일은 넣었는데 반영이 안 된다'의
    // 나머지 절반이다(모델·생산월 주문이 없거나, PTO·캡 판정이 갈렸거나).
    const used = new Set();
    for (const r of rowsRaw) {
      const f = String(r['Compared SAM file name'] || '').split(/[\\/]/).pop();
      if (f) used.add(f);
    }
    // WINGS 주문을 생산월별로 세어 둔다. SAM 문서가 안 붙는 가장 흔한 이유가
    // '그 생산월 주문이 아직 없다' 인데, 그건 이 수를 봐야만 알 수 있다.
    // (생산일이 비어 있는 주문은 생산월 폴더를 못 정해 최신월부터 훑는다 — 따로 센다.)
    const ordersByMonth = {};
    let ordersNoDate = 0;
    for (const r of rowsRaw) {
      const p = String(r['Production date'] || '').slice(0, 7);
      if (/^\d{4}-\d{2}$/.test(p)) {
        const ym = Number(p.slice(0, 4)) * 100 + Number(p.slice(5, 7));
        ordersByMonth[ym] = (ordersByMonth[ym] || 0) + 1;
      } else ordersNoDate += 1;
    }
    diag.ordersByMonth = ordersByMonth;
    diag.ordersNoDate = ordersNoDate;

    diag.samUnused = [];
    for (const m of diag.samMonths) {
      for (const d of (m.docs || [])) {
        if (used.has(d.file)) continue;
        diag.samUnused.push({ month: m.name, file: d.file, pto: d.pto, keys: d.keys,
          orders: ordersByMonth[m.yyyymm] || 0 });
      }
    }
    if (diag.samUnused.length) log(`[build] 어떤 주문에도 안 붙은 SAM 문서 ${diag.samUnused.length}개`);

    const cols = DISPLAY_COLS.filter((c) => rowsRaw.some((r) => c in r));
    const rows = rowsRaw.map(function (r) {
      const o = {};
      for (const c of cols) o[c] = clean(r[c]);
      return o;
    });

    const count = (st) => rows.filter((r) => r['SAM Status'] === st).length;
    const generatedAt = (opts.now ? new Date(opts.now) : new Date())
      .toISOString().replace(/\.\d{3}Z$/, '+00:00');

    const data = {
      generated_at: generatedAt,
      wings_file: wings.file,
      columns: cols,
      summary: {
        total: rows.length,
        matched: count('Match'),
        mismatched: count('Mismatch'),
        no_sam: count('No SAM'),
        sam_update: rows.filter((r) => r['SAM Update']).length,   // 총계와 무관(겹침)
        sam_months: Object.keys(samMaps).map(Number).sort((a, b) => a - b),
      },
      rows: rows,
      // 읽은 SAM 문서 전체(생산월 폴더별) — 'SAM 파일 비교' 메뉴가 두 문서를 골라 맞대어 본다.
      sam_files: diag.samDocs,
      // 진단을 결과에 같이 실어 둔다 — 안 붙은 SAM 문서는 rows 에 흔적이 없어서,
      // 여기 없으면 빌드한 사람의 브라우저를 벗어나는 순간 사라진다.
      diagnostics: {
        sam_months: diag.samMonths.map((m) => ({
          name: m.name, yyyymm: m.yyyymm, files: m.files, parsed: m.parsed, keys: m.keys })),
        sam_skipped: diag.samSkipped,
        sam_unused: diag.samUnused,
        orders_by_month: diag.ordersByMonth,
        orders_no_date: diag.ordersNoDate,
      },
    };

    const groups = {};
    for (const g of Object.keys(ref.mandatory.groups || {})) {
      groups[g] = Array.from(ref.mandatory.groups[g]).sort();
    }
    const codes = {
      options: ref.codeDict,
      mandatory: ref.mandatory.desc || {},
      mandatory_groups: groups,
    };

    if (t0) log(`[build] 완료 (${Math.round((performance.now() - t0) / 100) / 10}s)`);
    return { data: data, codes: codes, diag: diag };
  }

  // build_data._clean: JSON 에 안전한 스칼라로.
  function clean(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number' && isNaN(v)) return '';
    if (v instanceof Date) {
      const p = (n) => String(n).padStart(2, '0');
      return v.getUTCFullYear() + '-' + p(v.getUTCMonth() + 1) + '-' + p(v.getUTCDate());
    }
    if (v instanceof Set) return Array.from(v).join(',');
    return v;
  }

  const api = {
    build: build,
    loadRefData: loadRefData,
    loadWings: loadWings,
    loadSam: loadSam,
    DISPLAY_COLS: DISPLAY_COLS,
    _recency: recency,
    _samMonthFolders: samMonthFolders,
  };
  if (isNode) module.exports = api;
  else root.Pipeline = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
