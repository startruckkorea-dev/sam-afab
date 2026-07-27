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
    '_paint_wings', '_paint_sam', '_tyre_wings', '_tyre_sam',
  ];

  // 참조 워크북 파일명(04. code / 03. model_rules). 대소문자 무시로 찾는다.
  const REF_FILES = {
    codeDict: 'mbtruck-spec-data.xlsx',
    mandatory: 'mandatory-codes.xlsx',
    modelCategory: 'model-category.xlsx',
    cab: 'cab.xlsx',
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
    const mandatory = await tryLoad('mandatory', RefData.loadMandatory,
      { desc: {}, set: new Set(), groups: {}, cats: {} });
    const modelCategory = await tryLoad('modelCategory', RefData.loadModelCategory, {});
    const cabMap = await tryLoad('cab', RefData.loadCabMap, {});
    const codeDict = await tryLoad('codeDict',
      (buf) => RefData.loadCodeDict(buf, 'code_dict'), null);

    let rules = RefData.RULE_DEFAULTS;
    try {
      const items = await src.list('model_rules');
      const name = findFile(items, REF_FILES.rules);
      rules = RefData.loadRules(name ? await src.download('model_rules', name) : null);
      log(`[ref] rules: ${name || '기본값(model_mapping.xlsx 없음)'}`);
    } catch (e) {
      log(`[ref] model_mapping.xlsx 실패(${String(e.message).slice(0, 80)}) — 기본 규칙 사용`);
      rules = RefData.loadRules(null);
    }

    log(`[ref] option codes ${Object.keys(optionCodes).length} · mandatory ${mandatory.set.size}`
      + ` · category ${Object.keys(modelCategory).length} · cab ${Object.keys(cabMap).length}`);

    return {
      rules: rules,
      mandatory: mandatory,
      modelCategory: modelCategory,
      cabMap: cabMap,
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
    // (신) MYxx 연식 폴더 → 그 하위의 YYYY-MM 폴더
    const myFolders = top.filter((it) => it.isFolder && /^MY\d{2}\b/i.test(it.name));
    for (const my of myFolders) {
      for (const m of samMonthFolders(await src.list('sam/' + my.name))) {
        out.push({ yyyymm: m.yyyymm, name: my.name + '/' + m.name, path: my.name + '/' + m.name });
      }
    }
    out.sort((a, b) => a.yyyymm - b.yyyymm);
    return out;
  }

  /** 생산월별 SAM 매핑. 존재하는 모든 생산월 폴더를 읽는다. */
  async function loadSam(src, ref, allMonths, log) {
    const months = await collectSamMonths(src);
    if (!months.length) {
      throw new Error('01. SAM_files 에 YYYY-MM 하위폴더가 없습니다. '
        + '(MYxx 연식 폴더 하위의 YYYY-MM 또는 직속 YYYY-MM 폴더가 필요)');
    }
    // 각 행의 생산월 폴더 + 이전 6개월로 'SAM update요청'을 판정하므로, 미래 폴더가
    // 있어도 이전 달이 누락되지 않도록 존재하는 모든 생산월 폴더를 읽는다.
    const targets = months;
    log(`[sam] 생산월 폴더 ${months.length}개 사용: ${targets.map((m) => m.name).join(', ')}`);

    const knownCodes = new Set(Object.keys(ref.codeDesc));
    for (const c of ref.mandatory.set) knownCodes.add(c);
    const ctx = { codeDesc: ref.codeDesc, knownCodes: knownCodes };

    const maps = {};
    for (const month of targets) {
      const files = (await src.list('sam/' + month.path))
        .filter((it) => !it.isFolder && ext(it.name) === '.docx' && it.name.indexOf('.') !== 0)
        .sort((a, b) => (a.name < b.name ? -1 : (a.name > b.name ? 1 : 0)));
      const parsed = [];
      for (const f of files) {
        try {
          const d = await SamParse.parseSamDocx(await src.download('sam/' + month.path, f.name),
            f.name, ctx);
          if (d) parsed.push(d);
          else log(`  [sam] 건너뜀(모델/코드 인식 실패): ${f.name}`);
        } catch (e) {
          log(`  [sam] 파싱 실패: ${f.name} — ${String(e.message).slice(0, 80)}`);
        }
      }
      const mapping = SamParse.buildSamMapping(parsed, ref.rules);
      if (mapping.size) maps[month.yyyymm] = mapping;
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

    const ref = await loadRefData(src, log);
    const wings = await loadWings(src, log);
    const samMaps = await loadSam(src, ref, !!opts.allMonths, log);

    const rowsRaw = Compare.compare(wings.rows, samMaps, {
      rules: ref.rules,
      mandatory: ref.mandatory,
      modelCategory: ref.modelCategory,
      cabMap: ref.cabMap,
      codeDesc: ref.codeDesc,
      today: opts.now,
    });
    log(`[build] 결과 ${rowsRaw.length} 행`);

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
    return { data: data, codes: codes };
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
