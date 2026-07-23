/* Mapa Acertos — doc. acerto sem valor (redistribuição EAN neto zero + sobras SAP→0044) */
(function (global) {
  'use strict';
  const MAPA_ACERTOS_JS_VERSION = '1.0.12';
  /** @deprecated test filter removed — full SAP-stock scope. Kept empty for old callers. */
  const MA_TEST_MATERIALS = [];
  /** @deprecated */
  const MAPA_ACERTOS_TEST_MATERIAL = '';
  const MAPA_ACERTOS_DEPOTS = ['8', '9', '11', '22', '44'];
  const MAPA_ACERTOS_SHEETS = { 8: '0008', 9: '0009', 11: '0011', 22: '0022', 44: '0044' };
  /** Depósito destino de sobras/falhas de inventário (não deixar excesso nos deps de venda). */
  const MAPA_ACERTOS_DEPOT_SOBRA = '44';
  /**
   * Template sheet1 (XML): headers rows 7–8, cols A–O (B = Descrição; NOT B–F merge).
   * First data row = 9 (overwrite thin spacer). SUM formulas updated to K9/N9 at export.
   * Never reuse green sample styles from template row 10 (xf 98–104).
   */
  const DATA_START_ROW = 9;
  const DATA_END_ROW = 128;
  const DATA_ROW_HEIGHT = '19.15';
  const DEPOT_CELL = 'M3';
  const TEMPLATE_URL = 'assets/mapa_acertos_template.xlsx?v=' + MAPA_ACERTOS_JS_VERSION;
  /**
   * Clean (no-fill) style ids from template data row 12 — never from green sample row 10.
   * A Código · B Descrição · C Lote · D Armazém · E Doc.Pendentes · F Preço · G IVA · H UMB ·
   * I Existência Física · J Inventário Contabilístico · K Existência Valorizadas ·
   * L Dif (+) · M Dif (−) · N Dif Valorizadas · O Justificações
   * L/M use xf 27/28 (numFmt 164/165). Template used to paint positives/negatives [White]
   * (hiding Dif (−) qty); export patches styles.xml. O xf 52 must be left+dark (not right+theme1).
   */
  const MA_CLEAN_STYLES = {
    A: '24', B: '20', C: '22', D: '20', E: '26', F: '23', G: '20', H: '20',
    I: '20', J: '20', K: '25', L: '27', M: '28', N: '29', O: '52',
  };
  const MA_DATA_COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];
  /** Green sample style ids on template row 10 — remap → clean when stripping. */
  const MA_GREEN_STYLE_REMAP = {
    '98': '24', '99': '20', '100': '22', '101': '23', '102': '27', '103': '28', '104': '29',
    '53': '20', '54': '22', '55': '23', '56': '25', '57': '29',
  };
  /** Float epsilon for qty / valued-diff balance checks. */
  const MA_BALANCE_EPS = 1e-6;
  /** Round money/qty to 3 dp (template numFmt #,##0.000). */
  function maRound3(n) {
    const v = Number(n);
    if (!Number.isFinite(v)) return 0;
    return Math.round(v * 1000) / 1000;
  }
  function maTestMaterialsLabel() {
    return 'todos os artigos com stock SAP';
  }
  /** SAP unit price from buildSapPrecoMap (EAN), with material fallback. */
  function maGetSapPrice(ean, material) {
    let best = 0;
    if (typeof buildSapPrecoMap === 'function' && ean) {
      const p = Number(buildSapPrecoMap()[ean]) || 0;
      if (p > 0) best = p;
    }
    if (best > 0) return best;
    const sap = typeof sapData !== 'undefined' ? sapData : [];
    const mat = maNormCode(material);
    for (var i = 0; i < sap.length; i++) {
      const r = sap[i];
      const p = Number(r.preco) || 0;
      if (p <= 0) continue;
      if (ean && r.EAN_norm === ean && p > best) best = p;
      else if (mat && maNormCode(r.material) === mat && p > best) best = p;
    }
    return best;
  }
  const FFLATE_CDN_URLS = [
    'https://cdn.jsdelivr.net/npm/fflate@0.8.2/umd/index.min.js',
    'https://unpkg.com/fflate@0.8.2/umd/index.min.js',
  ];
  let mapaAcertosRows = [];
  let mapaAcertosTemplateBuf = null;
  let _maFflatePromise = null;

  function maNormCode(v) {
    return String(v ?? '').trim().replace(/^0+(?=\d)/, '');
  }
  function maDepotSheet(dk) {
    return MAPA_ACERTOS_SHEETS[dk] || String(dk).padStart(4, '0');
  }
  function maResolveMaterial(reconRow, depotKey) {
    if (reconRow && reconRow.material) return String(reconRow.material).trim();
    if (typeof buildEanSapCodeMap === 'function' && typeof sapData !== 'undefined' && sapData) buildEanSapCodeMap();
    const ean = reconRow && reconRow.EAN_norm;
    if (!ean) return '';
    const sap = typeof sapData !== 'undefined' ? sapData : [];
    const depotHit = sap.find(function (x) { return x.deposito === depotKey && x.EAN_norm === ean && x.material; });
    if (depotHit) return String(depotHit.material).trim();
    if (typeof _eanSapCodeMap !== 'undefined' && _eanSapCodeMap && _eanSapCodeMap[ean]) return [..._eanSapCodeMap[ean]][0];
    const anyEan = sap.find(function (x) { return x.EAN_norm === ean && x.material; });
    if (anyEan) return String(anyEan.material).trim();
    return '';
  }
  function maResolveCat(ean, existingCat, existingSub) {
    if (existingCat || existingSub) return { cat: existingCat || '', subcat: existingSub || '' };
    const params = typeof occParams !== 'undefined' ? occParams : [];
    const p = params.find(function (x) { return x.ean === ean; });
    return { cat: (p && p.cat) || '', subcat: (p && p.subcat) || '' };
  }
  function maGetUmb(ean) {
    const row = (typeof sapData !== 'undefined' ? sapData : []).find(function (x) { return x.EAN_norm === ean; });
    const umb = row && (row.umb || row.UMB);
    return umb ? String(umb).trim() : 'UN';
  }
  function maMatchesTestMaterial(r, depotKey) {
    const targets = MA_TEST_MATERIALS;
    const mat = maNormCode(maResolveMaterial(r, depotKey) || r.material);
    if (mat && targets.indexOf(mat) >= 0) return true;
    if (targets.indexOf(maNormCode(r.EAN_norm)) >= 0) return true;
    const codes = typeof getSapCode === 'function'
      ? String(getSapCode(r.EAN_norm) || '').split('/').map(function (c) { return maNormCode(c); })
      : [];
    for (var i = 0; i < codes.length; i++) {
      if (codes[i] && targets.indexOf(codes[i]) >= 0) return true;
    }
    return false;
  }

  /**
   * Neto qty / valued diff must be 0 globally and per EAN.
   * Only depot/lote moves within the same EAN — never inventário that creates value gain/loss.
   * Price is per EAN ⇒ neto qty 0 within EAN ⇒ neto valor 0.
   */
  function maValidateNetZero(rows) {
    const byEan = {};
    var netQty = 0;
    var netVal = 0;
    const issues = [];
    for (var i = 0; i < rows.length; i++) {
      const r = rows[i];
      const ean = r.EAN_norm || '';
      const adj = Number(r.adj) || 0;
      const preco = Number(r.preco) > 0 ? Number(r.preco) : maGetSapPrice(ean, r.material);
      const val = adj * (Number(preco) || 0);
      if (!byEan[ean]) byEan[ean] = { qty: 0, val: 0 };
      byEan[ean].qty += adj;
      byEan[ean].val += val;
      netQty += adj;
      netVal += val;
    }
    if (Math.abs(netQty) > MA_BALANCE_EPS) {
      issues.push('neto qty global = ' + maRound3(netQty));
    }
    if (Math.abs(netVal) > MA_BALANCE_EPS) {
      issues.push('neto valor (Σ Dif. Valorizadas) = ' + maRound3(netVal));
    }
    Object.keys(byEan).forEach(function (ean) {
      if (Math.abs(byEan[ean].qty) > MA_BALANCE_EPS) {
        issues.push('EAN ' + ean + ' neto qty = ' + maRound3(byEan[ean].qty));
      }
      if (Math.abs(byEan[ean].val) > MA_BALANCE_EPS) {
        issues.push('EAN ' + ean + ' neto valor = ' + maRound3(byEan[ean].val));
      }
    });
    return { ok: issues.length === 0, netQty: netQty, netVal: netVal, byEan: byEan, issues: issues };
  }

  /**
   * If adjMap for one EAN is not neto-zero (float noise or bug), absorb residual into
   * the largest absolute leg so DESCER+SUBIR stay paired. Returns true if balanced after fix.
   */
  function maFixAdjMapNetZero(adjMap) {
    var sum = 0;
    var keys = Object.keys(adjMap);
    for (var i = 0; i < keys.length; i++) sum += adjMap[keys[i]] || 0;
    if (Math.abs(sum) <= MA_BALANCE_EPS) return true;
    if (Math.abs(sum) > 0.01) {
      // Structural imbalance — refuse to silently invent qty
      return false;
    }
    // Tiny float residual: absorb into largest |adj|
    var bestKey = null;
    var bestAbs = 0;
    for (var j = 0; j < keys.length; j++) {
      var a = Math.abs(adjMap[keys[j]] || 0);
      if (a > bestAbs) { bestAbs = a; bestKey = keys[j]; }
    }
    if (!bestKey) return false;
    adjMap[bestKey] = (adjMap[bestKey] || 0) - sum;
    return Math.abs(Object.keys(adjMap).reduce(function (s, k) { return s + (adjMap[k] || 0); }, 0)) <= MA_BALANCE_EPS;
  }

  /**
   * Balance gaps within one EAN across any (depot, lote) cells.
   * gap = Unilog - SAP. (+) subir SAP; (-) descer SAP. Neto of adjustments = 0.
   * entries: [{ key, gap, ... }] — key uniquely identifies depot+lote.
   */
  function maBalanceGaps(entries) {
    const adj = {};
    entries.forEach(function (d) { adj[d.key] = 0; });
    const sinks = entries.filter(function (d) { return d.gap > 0; })
      .map(function (d) { return { key: d.key, remain: d.gap }; })
      .sort(function (a, b) { return b.remain - a.remain; });
    const sources = entries.filter(function (d) { return d.gap < 0; })
      .map(function (d) { return { key: d.key, remain: -d.gap }; })
      .sort(function (a, b) { return b.remain - a.remain; });
    var si = 0, so = 0;
    while (si < sinks.length && so < sources.length) {
      if (sinks[si].remain <= 0) { si++; continue; }
      if (sources[so].remain <= 0) { so++; continue; }
      var take = Math.min(sinks[si].remain, sources[so].remain);
      if (take <= 0) break;
      adj[sinks[si].key] += take;
      adj[sources[so].key] -= take;
      sinks[si].remain -= take;
      sources[so].remain -= take;
    }
    return adj;
  }

  /**
   * After EAN-level balance: any remaining SAP excess vs Unilog on non-0044 depots
   * (sap_after = qt_sap + adj > qt_uni) must DESCER on the source and SUBIR on 0044
   * (same lote, same EAN, same qty) so neto qty and neto valor stay 0.
   * Excess already on 0044 is left as-is. Never emits unpaired residual.
   * Mutates g.cells (may create 0044|lote) and adjMap; sets sobraOut/sobraIn on cells.
   */
  function maMoveExcessTo0044(g, adjMap) {
    const sourceKeys = Object.keys(g.cells);
    for (var i = 0; i < sourceKeys.length; i++) {
      const e = g.cells[sourceKeys[i]];
      if (!e || e.depot === MAPA_ACERTOS_DEPOT_SOBRA) continue;
      const adj = adjMap[e.key] || 0;
      const sapAfter = (Number(e.qt_sap) || 0) + adj;
      const excess = sapAfter - (Number(e.qt_uni) || 0);
      if (!(excess > 0)) continue;
      adjMap[e.key] = adj - excess;
      e.sobraOut = (e.sobraOut || 0) + excess;
      const tKey = MAPA_ACERTOS_DEPOT_SOBRA + '|' + (e.lote || '');
      if (!g.cells[tKey]) {
        g.cells[tKey] = {
          key: tKey,
          depot: MAPA_ACERTOS_DEPOT_SOBRA,
          depotSheet: maDepotSheet(MAPA_ACERTOS_DEPOT_SOBRA),
          lote: e.lote || '',
          material: e.material || g.material || '',
          desc: e.desc || g.desc || '',
          qt_sap: 0,
          qt_uni: 0,
          gap: 0,
        };
      } else if (!g.cells[tKey].material && (e.material || g.material)) {
        g.cells[tKey].material = e.material || g.material;
      }
      const t = g.cells[tKey];
      adjMap[tKey] = (adjMap[tKey] || 0) + excess;
      t.sobraIn = (t.sobraIn || 0) + excess;
    }
  }

  function buildMapaAcertosRows(options) {
    const opts = options || {};
    /** Full scope by default. Pass testOnly:true only for legacy debug. */
    const testOnly = opts.testOnly === true;
    const filterMaterial = opts.material || '';
    const recon = typeof reconData !== 'undefined' ? reconData : {};
    // Group by EAN only — balance any lote within that product
    const groups = new Map();

    for (const dk of MAPA_ACERTOS_DEPOTS) {
      const rows = recon[dk] || [];
      for (const r of rows) {
        if (testOnly && !maMatchesTestMaterial(r, dk)) continue;
        if (filterMaterial) {
          const mat = maNormCode(maResolveMaterial(r, dk) || r.material);
          if (mat !== maNormCode(filterMaterial) && maNormCode(r.EAN_norm) !== maNormCode(filterMaterial)) continue;
        }
        const material = maResolveMaterial(r, dk) || filterMaterial || '';
        const ean = r.EAN_norm || '';
        const lote = r.Lote_norm || '';
        if (!ean) continue;
        // Scope: articles with SAP stock (qt_sap > 0) in at least this cell, or Unilog gap vs SAP
        const qtSapCell = Number(r.qt_sap) || 0;
        const qtUniCell = Number(r.qt_uni) || 0;
        if (qtSapCell <= 0 && qtUniCell <= 0) continue;
        if (!groups.has(ean)) {
          const cs0 = maResolveCat(ean, r.cat, r.subcat);
          groups.set(ean, {
            EAN_norm: ean,
            material: material,
            desc: r.desc || '',
            cat: cs0.cat,
            subcat: cs0.subcat,
            umb: maGetUmb(ean),
            cells: {},
          });
        }
        const g = groups.get(ean);
        if (!g.material && material) g.material = material;
        if (!g.desc && r.desc) g.desc = r.desc;
        if ((!g.cat && !g.subcat) && (r.cat || r.subcat)) {
          const cs1 = maResolveCat(ean, r.cat, r.subcat);
          g.cat = cs1.cat;
          g.subcat = cs1.subcat;
        }
        const cellKey = dk + '|' + lote;
        const qt_sap = Number(r.qt_sap) || 0;
        const qt_uni = Number(r.qt_uni) || 0;
        const prev = g.cells[cellKey];
        if (prev) {
          prev.qt_sap += qt_sap;
          prev.qt_uni += qt_uni;
          prev.gap = prev.qt_uni - prev.qt_sap;
          if (!prev.material && material) prev.material = material;
          if (!prev.desc && r.desc) prev.desc = r.desc;
        } else {
          g.cells[cellKey] = {
            key: cellKey,
            depot: dk,
            depotSheet: maDepotSheet(dk),
            lote: lote,
            material: material,
            desc: r.desc || '',
            qt_sap: qt_sap,
            qt_uni: qt_uni,
            gap: qt_uni - qt_sap,
          };
        }
      }
    }

    const out = [];
    for (const g of groups.values()) {
      const entries = Object.keys(g.cells).map(function (k) { return g.cells[k]; });
      var sapTotal = 0;
      for (var ei = 0; ei < entries.length; ei++) sapTotal += Number(entries[ei].qt_sap) || 0;
      // Only EANs that have SAP stock somewhere in scope depots
      if (sapTotal <= 0) continue;
      // 1) EAN-level balanced redistrib (match Unilog gaps, neto 0)
      const adjMap = maBalanceGaps(entries);
      const balanceAdjByKey = {};
      Object.keys(adjMap).forEach(function (k) { balanceAdjByKey[k] = adjMap[k]; });
      // 2) Residual SAP excess (after balance) → 0044 (paired DESCER source + SUBIR 0044, same EAN/qty)
      maMoveExcessTo0044(g, adjMap);
      // 3) Harden: per-EAN neto qty must be 0 (never emit unbalanced inventário a valor)
      if (!maFixAdjMapNetZero(adjMap)) {
        try {
          console.error('[mapa-acertos] EAN ' + g.EAN_norm + ' adj desequilibrado — linhas omitidas', adjMap);
        } catch (e) {}
        continue;
      }
      const finalEntries = Object.keys(g.cells).map(function (k) { return g.cells[k]; });
      const preco = maGetSapPrice(g.EAN_norm, g.material || filterMaterial);
      for (const e of finalEntries) {
        const adj = adjMap[e.key] || 0;
        if (!adj) continue;
        const sobra0044 = !!(e.sobraOut || e.sobraIn);
        const hadBalance = !!(balanceAdjByKey[e.key]);
        out.push({
          depot: e.depot,
          depotSheet: e.depotSheet,
          material: e.material || g.material || '',
          EAN_norm: g.EAN_norm,
          desc: e.desc || g.desc || '',
          lote: e.lote || '',
          cat: g.cat || '',
          subcat: g.subcat || '',
          umb: g.umb || 'UN',
          preco: preco,
          qt_uni: e.qt_uni,
          qt_sap: e.qt_sap,
          gap: e.gap,
          adj: adj,
          subir: adj > 0 ? adj : 0,
          descer: adj < 0 ? -adj : 0,
          sobra0044: sobra0044,
          hadBalance: hadBalance,
        });
      }
    }

    out.sort(function (a, b) {
      const d = String(a.depotSheet).localeCompare(String(b.depotSheet));
      if (d) return d;
      const m = String(a.material).localeCompare(String(b.material));
      if (m) return m;
      return String(a.lote || '').localeCompare(String(b.lote || ''));
    });

    mapaAcertosRows = out;
    try {
      const bal = maValidateNetZero(out);
      console.log('[mapa-acertos] scope=full linhas=' + out.length + ' eans≈' + Object.keys(bal.byEan || {}).length);
      console.log('[mapa-acertos] net adj qty (deve ser 0):', bal.netQty, '| net valor (deve ser 0):', bal.netVal, bal.ok ? 'OK' : bal.issues);
    } catch (e) {}
    return out;
  }

  /** Rows after cat/sub filters — used by UI and Excel export. */
  function getMapaAcertosFilteredRows(options) {
    let rows = buildMapaAcertosRows(options || { testOnly: false });

    const cats = [...new Set(rows.map(function (r) { return r.cat; }).filter(Boolean))].sort();
    if (typeof msUpdate === 'function') msUpdate('maCat', cats, 'Categorias', 'renderMapaAcertosPage');
    const selCats = typeof msGet === 'function' ? msGet('maCat') : new Set();
    const subSrc = selCats.size ? rows.filter(function (r) { return selCats.has(r.cat); }) : rows;
    const subs = [...new Set(subSrc.map(function (r) { return r.subcat; }).filter(Boolean))].sort();
    if (typeof msUpdate === 'function') msUpdate('maSub', subs, 'Subcategorias', 'renderMapaAcertosPage');
    const selSubs = typeof msGet === 'function' ? msGet('maSub') : new Set();

    if (selCats.size) rows = rows.filter(function (r) { return selCats.has(r.cat); });
    if (selSubs.size) rows = rows.filter(function (r) { return selSubs.has(r.subcat); });
    return rows;
  }

  function maXmlEsc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function maGetFflateRaw() {
    return typeof fflate !== 'undefined' ? fflate : global.fflate;
  }
  function maFflateApi(f) {
    if (!f) return null;
    // Sync only — async zip/unzip return Promises and corrupt the download.
    const unzip = f.unzipSync;
    const zip = f.zipSync;
    if (!unzip || !zip || !f.strFromU8 || !f.strToU8) return null;
    return { unzip: unzip, zip: zip, strFromU8: f.strFromU8, strToU8: f.strToU8 };
  }
  function maFflate() {
    const api = maFflateApi(maGetFflateRaw());
    if (!api) throw new Error('fflate não carregado — recarrega a página');
    return api;
  }
  function maLoadScript(src) {
    return new Promise(function (resolve, reject) {
      const existing = document.querySelector('script[src="' + src + '"]');
      if (existing && maFflateApi(maGetFflateRaw())) {
        resolve();
        return;
      }
      const s = document.createElement('script');
      s.src = src;
      s.async = true;
      s.onload = function () { resolve(); };
      s.onerror = function () { reject(new Error('Falha ao carregar ' + src)); };
      document.head.appendChild(s);
    });
  }
  async function maEnsureFflate() {
    var api = maFflateApi(maGetFflateRaw());
    if (api) return api;
    if (_maFflatePromise) return _maFflatePromise;
    _maFflatePromise = (async function () {
      for (var i = 0; i < FFLATE_CDN_URLS.length; i++) {
        try {
          await maLoadScript(FFLATE_CDN_URLS[i]);
          api = maFflateApi(maGetFflateRaw());
          if (api) return api;
        } catch (e) { /* try next CDN */ }
      }
      return null;
    })();
    try {
      return await _maFflatePromise;
    } finally {
      if (!maFflateApi(maGetFflateRaw())) _maFflatePromise = null;
    }
  }
  /**
   * Windows ZipFile/Compress-Archive may store entry names with backslashes.
   * fflate keeps them as-is, so lookups for 'xl/workbook.xml' miss and sheets
   * are never patched (export looks like empty template: headers OK, data blank).
   */
  function maNormalizeZipFiles(files) {
    const out = {};
    Object.keys(files).forEach(function (k) {
      const nk = String(k).replace(/\\/g, '/');
      if (!out[nk] || k === nk) out[nk] = files[k];
    });
    return out;
  }
  function maZipGet(files, path) {
    if (files[path]) return files[path];
    const alt = path.replace(/\//g, '\\');
    return files[alt] || null;
  }
  function maBuildSheetPathMap(files) {
    const strFromU8 = maFflate().strFromU8;
    const wbU8 = maZipGet(files, 'xl/workbook.xml');
    const relsU8 = maZipGet(files, 'xl/_rels/workbook.xml.rels');
    if (!wbU8 || !relsU8) {
      throw new Error('Template ZIP inválido — workbook.xml / rels em falta (paths?)');
    }
    const wb = strFromU8(wbU8);
    const rels = strFromU8(relsU8);
    const relMap = {};
    const relRe = /<Relationship[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?>/g;
    var m;
    while ((m = relRe.exec(rels))) {
      var target = m[2].replace(/^\//, '').replace(/\\/g, '/');
      relMap[m[1]] = target.indexOf('xl/') === 0 ? target : 'xl/' + target;
    }
    // also Target-before-Id order
    const relRe2 = /<Relationship[^>]*\bTarget="([^"]+)"[^>]*\bId="([^"]+)"[^>]*\/?>/g;
    while ((m = relRe2.exec(rels))) {
      if (relMap[m[2]]) continue;
      var t2 = m[1].replace(/^\//, '').replace(/\\/g, '/');
      relMap[m[2]] = t2.indexOf('xl/') === 0 ? t2 : 'xl/' + t2;
    }
    const sheetMap = {};
    // Accept name/r:id in either attribute order
    const sheetTagRe = /<sheet\b[^>]*\/?>/g;
    while ((m = sheetTagRe.exec(wb))) {
      const tag = m[0];
      const nm = tag.match(/\bname="([^"]+)"/);
      const rid = tag.match(/\br:id="([^"]+)"/);
      if (!nm || !rid || !relMap[rid[1]]) continue;
      sheetMap[nm[1]] = relMap[rid[1]];
    }
    return sheetMap;
  }
  function maReplaceCell(sheetXml, addr, replacement) {
    const cellRe = new RegExp('<c r="' + addr + '"(?:[^>/]*/>|[^>]*>(?:[\\s\\S]*?)</c>)');
    if (cellRe.test(sheetXml)) return sheetXml.replace(cellRe, replacement);
    // Cell missing — insert into its row ( foreward-slash-safe upsert )
    const rowNum = addr.replace(/^[A-Z]+/, '');
    const rowRe = new RegExp('(<row r="' + rowNum + '"[^>]*>)([\\s\\S]*?)(</row>)');
    if (!rowRe.test(sheetXml)) return sheetXml;
    return sheetXml.replace(rowRe, function (_m, open, inner, close) {
      return open + inner + replacement + close;
    });
  }
  function maCellText(addr, style, text) {
    return '<c r="' + addr + '"' + (style ? ' s="' + style + '"' : '') + ' t="inlineStr"><is><t>' + maXmlEsc(text) + '</t></is></c>';
  }
  function maCellNum(addr, style, num) {
    return '<c r="' + addr + '"' + (style ? ' s="' + style + '"' : '') + '><v>' + num + '</v></c>';
  }
  function maCellEmpty(addr, style) {
    return '<c r="' + addr + '"' + (style ? ' s="' + style + '"' : '') + '/>';
  }
  function maExtractStyle(sheetXml, addr, fallback) {
    const m = sheetXml.match(new RegExp('<c r="' + addr + '"([^>/]*)'));
    if (!m) return fallback;
    const sm = m[1].match(/\bs="(\d+)"/);
    return sm ? sm[1] : fallback;
  }
  function maJustification(row) {
    if (row && row.justificacao) return String(row.justificacao);
    return 'acerto entre depositos';
  }
  /** Dif (+)/(−) from acerto adj: SUBIR → L; DESCER (SAP > física) → M absolute qty. */
  function maDifParts(row) {
    const adj = Number(row.adj) || 0;
    var subir = adj > 0 ? maRound3(adj) : 0;
    var descer = adj < 0 ? maRound3(-adj) : 0;
    if (!subir && Number(row.subir) > 0) subir = maRound3(row.subir);
    if (!descer && Number(row.descer) > 0) descer = maRound3(row.descer);
    return { subir: subir, descer: descer };
  }
  function maSetRowHeight(sheetXml, rowNum, ht) {
    return sheetXml.replace(new RegExp('<row r="' + rowNum + '"([^>]*)>'), function (_m, attrs) {
      var a = String(attrs || '').replace(/\sht="[^"]*"/g, '').replace(/\scustomHeight="[^"]*"/g, '');
      return '<row r="' + rowNum + '"' + a + ' ht="' + ht + '" customHeight="1">';
    });
  }
  function maStripGreenStyles(sheetXml) {
    // Remap known green/sample xf ids on any cell in the data band.
    return sheetXml.replace(/<c r="([A-O])(\d+)" s="(\d+)"/g, function (m, col, rowStr, styleId) {
      const row = parseInt(rowStr, 10);
      if (row < DATA_START_ROW || row > DATA_END_ROW) return m;
      const clean = MA_GREEN_STYLE_REMAP[styleId];
      if (!clean) return m;
      return '<c r="' + col + rowStr + '" s="' + clean + '"';
    });
  }
  /**
   * Template historically hid Dif signs with [White] numFmts and right-aligned Justificações
   * (clips "acerto…" → "o entre depositos"). Force visible black #,##0.000 and left+wrap on O.
   */
  function maFixStylesXml(stylesXml) {
    stylesXml = stylesXml
      .replace(/numFmtId="164" formatCode="[^"]*"/g, 'numFmtId="164" formatCode="#,##0.000;-#,##0.000"')
      .replace(/numFmtId="165" formatCode="[^"]*"/g, 'numFmtId="165" formatCode="#,##0.000;-#,##0.000"');
    // Justificações xf (style 52 only — fontId 43 theme lt1 + right). Do NOT touch xf 24 (Código, font 21).
    stylesXml = stylesXml.replace(
      /<xf numFmtId="49" fontId="43" fillId="0" borderId="10" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"\/><\/xf>/g,
      '<xf numFmtId="49" fontId="21" fillId="0" borderId="10" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" wrapText="1"/></xf>'
    );
    return stylesXml;
  }
  /** Ensure L/M/O columns wide enough for qty and full justificativa. */
  function maEnsureColWidths(sheetXml) {
    sheetXml = sheetXml.replace(
      /<col min="12" max="12"[^/]*\/>/,
      '<col min="12" max="12" width="14" style="3" customWidth="1"/>'
    );
    sheetXml = sheetXml.replace(
      /<col min="13" max="13"[^/]*\/>/,
      '<col min="13" max="13" width="14" style="3" customWidth="1"/>'
    );
    sheetXml = sheetXml.replace(
      /<col min="15" max="15"[^/]*\/>/,
      '<col min="15" max="15" width="32" style="1" customWidth="1"/>'
    );
    return sheetXml;
  }
  function maFixSumFormulas(sheetXml) {
    return sheetXml
      .replace(/SUM\(K10:K128\)/g, 'SUM(K' + DATA_START_ROW + ':K' + DATA_END_ROW + ')')
      .replace(/SUM\(N10:N128\)/g, 'SUM(N' + DATA_START_ROW + ':N' + DATA_END_ROW + ')')
      .replace(/SUM\(K9:K128\)/g, 'SUM(K' + DATA_START_ROW + ':K' + DATA_END_ROW + ')')
      .replace(/SUM\(N9:N128\)/g, 'SUM(N' + DATA_START_ROW + ':N' + DATA_END_ROW + ')');
  }
  /** Dev check: ensure first data row + L/M/O were actually written into sheet XML. */
  function maVerifySheetDiffCells(sheetXml, sheetName, dataRows) {
    var l = 0, m = 0, o = 0;
    var n = Math.min(dataRows.length, DATA_END_ROW - DATA_START_ROW + 1);
    if (n > 0) {
      var a9 = sheetXml.match(new RegExp('<c r="A' + DATA_START_ROW + '"[^>]*>[\\s\\S]*?</c>|<c r="A' + DATA_START_ROW + '"[^>]*/>'));
      var hasMat = a9 && /<(?:v|t)>/.test(a9[0]);
      if (!hasMat) {
        throw new Error('Export sheet ' + sheetName + ': A' + DATA_START_ROW + ' sem dados após patch — abortado');
      }
    }
    for (var i = 0; i < n; i++) {
      var rowNum = DATA_START_ROW + i;
      var dif = maDifParts(dataRows[i]);
      if (dif.subir > 0) {
        if (new RegExp('<c r="L' + rowNum + '"[^>]*>\\s*<v>' + String(dif.subir).replace('.', '\\.') + '</v>').test(sheetXml)) l++;
      }
      if (dif.descer > 0) {
        if (new RegExp('<c r="M' + rowNum + '"[^>]*>\\s*<v>' + String(dif.descer).replace('.', '\\.') + '</v>').test(sheetXml)) m++;
      }
      if (sheetXml.indexOf('acerto entre depositos') >= 0 || sheetXml.indexOf('acerto invent') >= 0) o = 1;
    }
    try {
      console.log('[mapa-acertos] sheet ' + sheetName + ' verify L<' + l + '> M<' + m + '> Otext=' + !!o + ' rows=' + n);
    } catch (e) { /* ignore */ }
  }
  /**
   * Fill one data row — always MA_CLEAN_STYLES (no green fill).
   * L Dif (+) when SUBIR / physical>SAP adj; M Dif (−) when DESCER / SAP>physical.
   * N Dif Valorizadas = (L−M)×preço. O = "acerto entre depositos" only.
   */
  function maFillDataRow(sheetXml, rowNum, row) {
    const styles = MA_CLEAN_STYLES;
    const preco = maRound3(row.preco > 0 ? row.preco : maGetSapPrice(row.EAN_norm, row.material));
    const qtUni = maRound3(row.qt_uni);
    const qtSap = maRound3(row.qt_sap);
    const dif = maDifParts(row);
    const subir = dif.subir;
    const descer = dif.descer;
    const existVal = preco > 0 ? maRound3(qtUni * preco) : null;
    const difVal = preco > 0 ? maRound3((subir - descer) * preco) : null;
    const note = maJustification(row);

    sheetXml = maSetRowHeight(sheetXml, rowNum, DATA_ROW_HEIGHT);
    sheetXml = maReplaceCell(sheetXml, 'A' + rowNum, maCellText('A' + rowNum, styles.A, row.material));
    sheetXml = maReplaceCell(sheetXml, 'B' + rowNum, maCellText('B' + rowNum, styles.B, row.desc || ''));
    sheetXml = maReplaceCell(sheetXml, 'C' + rowNum, row.lote ? maCellText('C' + rowNum, styles.C, row.lote) : maCellEmpty('C' + rowNum, styles.C));
    sheetXml = maReplaceCell(sheetXml, 'D' + rowNum, maCellText('D' + rowNum, styles.D, row.depotSheet));
    sheetXml = maReplaceCell(sheetXml, 'E' + rowNum, maCellEmpty('E' + rowNum, styles.E));
    sheetXml = maReplaceCell(sheetXml, 'F' + rowNum, preco > 0 ? maCellNum('F' + rowNum, styles.F, preco) : maCellEmpty('F' + rowNum, styles.F));
    sheetXml = maReplaceCell(sheetXml, 'G' + rowNum, maCellEmpty('G' + rowNum, styles.G));
    sheetXml = maReplaceCell(sheetXml, 'H' + rowNum, maCellText('H' + rowNum, styles.H, row.umb || 'UN'));
    sheetXml = maReplaceCell(sheetXml, 'I' + rowNum, maCellNum('I' + rowNum, styles.I, qtUni));
    sheetXml = maReplaceCell(sheetXml, 'J' + rowNum, maCellNum('J' + rowNum, styles.J, qtSap));
    sheetXml = maReplaceCell(sheetXml, 'K' + rowNum, existVal != null ? maCellNum('K' + rowNum, styles.K, existVal) : maCellEmpty('K' + rowNum, styles.K));
    sheetXml = maReplaceCell(sheetXml, 'L' + rowNum, subir > 0 ? maCellNum('L' + rowNum, styles.L, subir) : maCellEmpty('L' + rowNum, styles.L));
    sheetXml = maReplaceCell(sheetXml, 'M' + rowNum, descer > 0 ? maCellNum('M' + rowNum, styles.M, descer) : maCellEmpty('M' + rowNum, styles.M));
    sheetXml = maReplaceCell(sheetXml, 'N' + rowNum, difVal != null ? maCellNum('N' + rowNum, styles.N, difVal) : maCellEmpty('N' + rowNum, styles.N));
    sheetXml = maReplaceCell(sheetXml, 'O' + rowNum, maCellText('O' + rowNum, styles.O, note));
    return sheetXml;
  }
  function maClearDataRow(sheetXml, rowNum) {
    sheetXml = maSetRowHeight(sheetXml, rowNum, DATA_ROW_HEIGHT);
    MA_DATA_COLS.forEach(function (col) {
      const addr = col + rowNum;
      sheetXml = maReplaceCell(sheetXml, addr, maCellEmpty(addr, MA_CLEAN_STYLES[col]));
    });
    return sheetXml;
  }
  const _maLastFilled = {};
  function maPatchSheetXml(sheetXml, sheetName, dataRows) {
    sheetXml = maReplaceCell(sheetXml, DEPOT_CELL, maCellText(DEPOT_CELL, maExtractStyle(sheetXml, DEPOT_CELL, '92'), sheetName));
    sheetXml = maStripGreenStyles(sheetXml);
    sheetXml = maEnsureColWidths(sheetXml);
    sheetXml = maFixSumFormulas(sheetXml);
    // First data = row 9 (no blank spacer). Clear prior fill + green sample styles.
    const prev = _maLastFilled[sheetName] || 0;
    const n = Math.min(dataRows.length, DATA_END_ROW - DATA_START_ROW + 1);
    const clearUntil = Math.max(prev, n, 2); // at least clear spacer+sample rows 9–10
    for (var r = DATA_START_ROW; r < DATA_START_ROW + clearUntil; r++) {
      if (r > DATA_END_ROW) break;
      sheetXml = maClearDataRow(sheetXml, r);
    }
    dataRows.forEach(function (row, i) {
      var rowNum = DATA_START_ROW + i;
      if (rowNum > DATA_END_ROW) return;
      sheetXml = maFillDataRow(sheetXml, rowNum, row);
    });
    maVerifySheetDiffCells(sheetXml, sheetName, dataRows);
    _maLastFilled[sheetName] = n;
    return sheetXml;
  }
  /** Drop calcChain after rewriting data cells — stale formula refs trigger Excel "LIVRO REPARADO". */
  function maStripCalcChain(files, strFromU8, strToU8) {
    Object.keys(files).forEach(function (k) {
      if (/calcChain\.xml$/i.test(String(k).replace(/\\/g, '/'))) delete files[k];
    });
    const relsPath = 'xl/_rels/workbook.xml.rels';
    const relsU8 = maZipGet(files, relsPath);
    if (relsU8) {
      var rels = strFromU8(relsU8);
      rels = rels.replace(/<Relationship[^>]*calcChain[^>]*\/?>/gi, '');
      files[relsPath] = strToU8(rels);
    }
    const ctPath = '[Content_Types].xml';
    const ctU8 = maZipGet(files, ctPath);
    if (ctU8) {
      var ct = strFromU8(ctU8);
      ct = ct.replace(/<Override[^>]*calcChain[^>]*\/>/gi, '');
      files[ctPath] = strToU8(ct);
    }
  }

  function maPatchTemplateXlsx(templateBuf, depotRowsMap) {
    const ff = maFflate();
    const files = maNormalizeZipFiles(ff.unzip(new Uint8Array(templateBuf)));
    const stylesU8 = maZipGet(files, 'xl/styles.xml');
    if (stylesU8) {
      files['xl/styles.xml'] = ff.strToU8(maFixStylesXml(ff.strFromU8(stylesU8)));
    }
    const sheetPaths = maBuildSheetPathMap(files);
    if (!Object.keys(sheetPaths).length) {
      throw new Error('Template sem folhas mapeadas (workbook sheet/rId)');
    }
    var patched = 0;
    for (const dk of MAPA_ACERTOS_DEPOTS) {
      const sheetName = MAPA_ACERTOS_SHEETS[dk];
      const path = sheetPaths[sheetName];
      const sheetU8 = path ? maZipGet(files, path) : null;
      if (!path || !sheetU8) {
        if ((depotRowsMap[dk] || []).length) {
          throw new Error('Folha ' + sheetName + ' em falta no template ZIP');
        }
        continue;
      }
      files[path] = ff.strToU8(maPatchSheetXml(ff.strFromU8(sheetU8), sheetName, depotRowsMap[dk] || []));
      patched++;
    }
    if (!patched) throw new Error('Nenhuma folha do Mapa Acertos foi escrita');
    maStripCalcChain(files, ff.strFromU8, ff.strToU8);
    return ff.zip(files);
  }

  /** SheetJS fallback when fflate is unavailable — preserves sheet structure, may lose some template styling. */
  function maSetCell(ws, addr, cell) {
    ws[addr] = cell;
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    if (!m || !ws['!ref']) return;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const c = XLSX.utils.decode_col(m[1]);
    const r = parseInt(m[2], 10) - 1;
    if (c < range.s.c) range.s.c = c;
    if (c > range.e.c) range.e.c = c;
    if (r < range.s.r) range.s.r = r;
    if (r > range.e.r) range.e.r = r;
    ws['!ref'] = XLSX.utils.encode_range(range);
  }
  function maClearSheetDataRow(ws, rowNum) {
    MA_DATA_COLS.forEach(function (col) {
      const addr = col + rowNum;
      if (ws[addr]) delete ws[addr];
    });
  }
  function maFillSheetDataRow(ws, rowNum, row) {
    const preco = maRound3(row.preco > 0 ? row.preco : maGetSapPrice(row.EAN_norm, row.material));
    const qtUni = maRound3(row.qt_uni);
    const qtSap = maRound3(row.qt_sap);
    const dif = maDifParts(row);
    const subir = dif.subir;
    const descer = dif.descer;
    const note = maJustification(row);
    // Plain cells — no style/fill copied from green sample rows.
    maSetCell(ws, 'A' + rowNum, { t: 's', v: String(row.material) });
    maSetCell(ws, 'B' + rowNum, { t: 's', v: String(row.desc || '') });
    if (row.lote) maSetCell(ws, 'C' + rowNum, { t: 's', v: String(row.lote) });
    else delete ws['C' + rowNum];
    maSetCell(ws, 'D' + rowNum, { t: 's', v: String(row.depotSheet) });
    delete ws['E' + rowNum];
    delete ws['G' + rowNum];
    if (preco > 0) maSetCell(ws, 'F' + rowNum, { t: 'n', v: preco });
    else delete ws['F' + rowNum];
    maSetCell(ws, 'H' + rowNum, { t: 's', v: String(row.umb || 'UN') });
    maSetCell(ws, 'I' + rowNum, { t: 'n', v: qtUni });
    maSetCell(ws, 'J' + rowNum, { t: 'n', v: qtSap });
    if (preco > 0) maSetCell(ws, 'K' + rowNum, { t: 'n', v: maRound3(qtUni * preco) });
    else delete ws['K' + rowNum];
    if (subir > 0) maSetCell(ws, 'L' + rowNum, { t: 'n', v: subir });
    else delete ws['L' + rowNum];
    if (descer > 0) maSetCell(ws, 'M' + rowNum, { t: 'n', v: descer });
    else delete ws['M' + rowNum];
    if (preco > 0) maSetCell(ws, 'N' + rowNum, { t: 'n', v: maRound3((subir - descer) * preco) });
    else delete ws['N' + rowNum];
    maSetCell(ws, 'O' + rowNum, { t: 's', v: note });
  }
  function maPatchTemplateXlsxSheetJS(templateBuf, depotRowsMap) {
    if (typeof XLSX === 'undefined' || !XLSX.read || !XLSX.write) {
      throw new Error('XLSX não carregado — recarrega a página');
    }
    const wb = XLSX.read(new Uint8Array(templateBuf), { type: 'array', cellStyles: true });
    for (const dk of MAPA_ACERTOS_DEPOTS) {
      const sheetName = MAPA_ACERTOS_SHEETS[dk];
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;
      maSetCell(ws, DEPOT_CELL, { t: 's', v: sheetName });
      // Fix SUM start row for SheetJS path
      ['K129', 'N129'].forEach(function (addr) {
        const cell = ws[addr];
        if (cell && cell.f) {
          cell.f = String(cell.f)
            .replace(/K10:K128/g, 'K' + DATA_START_ROW + ':K' + DATA_END_ROW)
            .replace(/N10:N128/g, 'N' + DATA_START_ROW + ':N' + DATA_END_ROW);
        }
      });
      const dataRows = depotRowsMap[dk] || [];
      const prev = _maLastFilled[sheetName] || 0;
      const n = Math.min(dataRows.length, DATA_END_ROW - DATA_START_ROW + 1);
      const clearUntil = Math.max(prev, n, 2);
      for (var r = DATA_START_ROW; r < DATA_START_ROW + clearUntil; r++) {
        if (r > DATA_END_ROW) break;
        maClearSheetDataRow(ws, r);
      }
      dataRows.forEach(function (row, i) {
        var rowNum = DATA_START_ROW + i;
        if (rowNum > DATA_END_ROW) return;
        maFillSheetDataRow(ws, rowNum, row);
      });
      _maLastFilled[sheetName] = n;
    }
    const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    return out instanceof Uint8Array ? out : new Uint8Array(out);
  }

  function maDownloadXlsx(buf, filename) {
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }
  async function loadMapaAcertosTemplate() {
    if (mapaAcertosTemplateBuf) return mapaAcertosTemplateBuf;
    const resp = await fetch(TEMPLATE_URL);
    if (!resp.ok) throw new Error('Template não encontrado (' + TEMPLATE_URL + ')');
    mapaAcertosTemplateBuf = await resp.arrayBuffer();
    return mapaAcertosTemplateBuf;
  }
  function mapaAcertosExportFilename() {
    return 'acerto_sem_valor_' + new Date().toISOString().slice(0, 10) + '.xlsx';
  }

  /**
   * Shared Excel emit pipeline (Mapa Acertos / Acerto Inventário).
   * Same template + ZIP patch — callers only choose which rows go into each depot sheet.
   * @returns {{ outBuf: Uint8Array, written: number, truncated: boolean, maxPerSheet: number, depotRowsMap: Object }}
   */
  async function maExportPatchedXlsx(rows, filename) {
    const depotRowsMap = {};
    var truncated = false;
    const maxPerSheet = DATA_END_ROW - DATA_START_ROW + 1;
    for (const dk of MAPA_ACERTOS_DEPOTS) depotRowsMap[dk] = [];
    for (const r of rows) {
      const list = depotRowsMap[r.depot];
      if (!list) continue;
      if (list.length >= maxPerSheet) { truncated = true; continue; }
      list.push(r);
    }
    // Always re-fetch template (avoid stale ArrayBuffer)
    mapaAcertosTemplateBuf = null;
    const buf = await loadMapaAcertosTemplate();
    const ff = await maEnsureFflate();
    var outBuf;
    if (ff) {
      outBuf = maPatchTemplateXlsx(buf, depotRowsMap);
    } else if (typeof XLSX !== 'undefined' && XLSX.read && XLSX.write) {
      console.warn('[mapa-acertos] fflate indisponível — export via SheetJS');
      outBuf = maPatchTemplateXlsxSheetJS(buf, depotRowsMap);
    } else {
      throw new Error('fflate não carregado — recarrega a página');
    }
    const written = MAPA_ACERTOS_DEPOTS.reduce(function (s, dk) {
      return s + (depotRowsMap[dk] || []).length;
    }, 0);
    if (filename) maDownloadXlsx(outBuf, filename);
    return { outBuf: outBuf, written: written, truncated: truncated, maxPerSheet: maxPerSheet, depotRowsMap: depotRowsMap };
  }

  async function exportMapaAcertosXlsx() {
    if (typeof company !== 'undefined' && company !== 'DFB') {
      toast('Acerto sem valor disponível apenas para DFB', 'error');
      return;
    }
    if (!Object.keys(typeof reconData !== 'undefined' ? reconData : {}).length) {
      toast('Sem dados de conciliação — carrega SAP e Unilog e processa', 'error');
      return;
    }
    const rows = getMapaAcertosFilteredRows({ testOnly: false });
    if (!rows.length) {
      const selCats = typeof msGet === 'function' ? msGet('maCat') : new Set();
      const selSubs = typeof msGet === 'function' ? msGet('maSub') : new Set();
      toast(
        selCats.size || selSubs.size
          ? 'Sem acertos nos filtros actuais (categoria/subcategoria)'
          : 'Sem acertos (sem gaps balanceáveis nem sobra SAP fora do 0044)',
        'info'
      );
      renderMapaAcertosPage();
      return;
    }
    const bal = maValidateNetZero(rows);
    if (!bal.ok) {
      try { console.error('[mapa-acertos] export bloqueado — balanceamento inválido', bal); } catch (e) {}
      toast('Export bloqueado: Σ Dif. Valorizadas / neto qty ≠ 0 (deve ser 0). ' + bal.issues.join('; '), 'error');
      renderMapaAcertosPage();
      return;
    }
    try {
      const result = await maExportPatchedXlsx(rows, mapaAcertosExportFilename());
      const subir = rows.reduce(function (s, r) { return s + r.subir; }, 0);
      const descer = rows.reduce(function (s, r) { return s + r.descer; }, 0);
      var msg = 'Exportado: ' + result.written + ' linhas · subir ' + subir + ' / descer ' + descer + ' · neto qty/valor 0';
      const selCats = typeof msGet === 'function' ? msGet('maCat') : new Set();
      const selSubs = typeof msGet === 'function' ? msGet('maSub') : new Set();
      if (selCats.size || selSubs.size) msg += ' (filtrado)';
      if (result.truncated) msg += ' · aviso: limite ' + result.maxPerSheet + ' linhas/folha (template)';
      toast(msg, result.truncated ? 'info' : 'success');
      renderMapaAcertosPage();
    } catch (e) {
      toast('Erro ao exportar: ' + (e.message || e), 'error');
    }
  }

  function renderMapaAcertosKpis(rows) {
    const el = document.getElementById('maKpis');
    if (!el) return;
    if (!rows.length) { el.innerHTML = ''; return; }
    const subir = rows.reduce(function (s, r) { return s + r.subir; }, 0);
    const descer = rows.reduce(function (s, r) { return s + r.descer; }, 0);
    const bal = maValidateNetZero(rows);
    const deps = new Set(rows.map(function (r) { return r.depotSheet; })).size;
    const fmt = typeof fmtKvNum === 'function' ? fmtKvNum : function (n) { return n; };
    const netCls = bal.ok ? '' : ' r';
    const eans = new Set(rows.map(function (r) { return r.EAN_norm; })).size;
    el.innerHTML =
      '<div class="kpi"><div class="kl">Linhas acerto</div><div class="kv">' + fmt(rows.length) + '</div><div class="ks">' + fmt(eans) + ' EAN · todos c/ stock SAP</div></div>' +
      '<div class="kpi"><div class="kl">Subir (Σ +)</div><div class="kv y">' + fmt(subir) + '</div><div class="ks">aumentar SAP</div></div>' +
      '<div class="kpi"><div class="kl">Descer (Σ −)</div><div class="kv r">' + fmt(descer) + '</div><div class="ks">diminuir SAP</div></div>' +
      '<div class="kpi"><div class="kl">Neto qty / valor</div><div class="kv' + netCls + '">' + fmt(maRound3(bal.netQty)) + ' / ' + fmt(maRound3(bal.netVal)) + '</div><div class="ks">ambos devem ser 0 · ' + fmt(deps) + ' deps</div></div>';
  }

  function renderMapaAcertosTable(rows) {
    const body = document.getElementById('maBody');
    const empty = document.getElementById('maEmpty');
    const wrap = document.getElementById('maTableWrap');
    const msgEl = document.getElementById('maEmptyMsg');
    if (!body) return;
    if (!rows.length) {
      body.innerHTML = '';
      if (empty) empty.style.display = 'block';
      if (wrap) wrap.style.display = 'none';
      const selCats = typeof msGet === 'function' ? msGet('maCat') : new Set();
      const selSubs = typeof msGet === 'function' ? msGet('maSub') : new Set();
      if (msgEl && (selCats.size || selSubs.size)) {
        msgEl.textContent = 'Nenhuma linha com os filtros activos — limpa categoria/subcategoria';
      }
      return;
    }
    if (empty) empty.style.display = 'none';
    if (wrap) wrap.style.display = '';
    const sorted = typeof sorts !== 'undefined' && sorts.ma ? rows.slice() : rows;
    if (typeof sorts !== 'undefined' && sorts.ma && typeof applySort === 'function') applySort('ma', sorted);
    const num = typeof fmtPtNum === 'function' ? fmtPtNum : function (n) { return n; };
    const signed = typeof fmtKvSigned === 'function' ? fmtKvSigned : function (n) { return n; };
    body.innerHTML = sorted.map(function (r) {
      const acao = r.adj > 0
        ? '<span class="badge binfo">Subir ' + num(r.adj) + '</span>'
        : '<span class="badge bwarn">Descer ' + num(-r.adj) + '</span>';
      return '<tr>' +
        '<td><span class="badge bgray">' + r.depotSheet + '</span></td>' +
        '<td style="font-family:var(--mono);font-size:11px">' + r.material + '</td>' +
        '<td style="font-family:var(--mono);font-size:11px">' + (r.EAN_norm || '—') + '</td>' +
        '<td class="col-desc">' + (r.desc || '—') + '</td>' +
        '<td style="font-family:var(--mono);font-size:11px">' + (r.lote || '—') + '</td>' +
        '<td class="num">' + num(r.qt_sap) + '</td>' +
        '<td class="num">' + num(r.qt_uni) + '</td>' +
        '<td class="num">' + signed(r.gap) + '</td>' +
        '<td class="num">' + signed(r.adj) + '</td>' +
        '<td>' + acao + '</td></tr>';
    }).join('');
    const table = document.getElementById('maTable');
    if (typeof updateSortHeaders === 'function') updateSortHeaders('ma');
    if (table && typeof enableTableSort === 'function') enableTableSort(table);
    if (typeof scheduleTableSort === 'function') scheduleTableSort();
    if (typeof scheduleColResize === 'function') scheduleColResize();
  }

  function renderMapaAcertosPage() {
    const qbNote = document.getElementById('maQbNote');
    const main = document.getElementById('maMain');
    if (typeof company !== 'undefined' && company !== 'DFB') {
      if (qbNote) qbNote.style.display = 'block';
      if (main) main.style.display = 'none';
      return;
    }
    if (qbNote) qbNote.style.display = 'none';
    if (main) main.style.display = '';
    if (!Object.keys(typeof reconData !== 'undefined' ? reconData : {}).length) {
      renderMapaAcertosKpis([]);
      renderMapaAcertosTable([]);
      const empty = document.getElementById('maEmpty');
      const msg = document.getElementById('maEmptyMsg');
      if (empty) empty.style.display = 'block';
      if (msg) msg.textContent = 'Carrega SAP + Unilog e clica Processar na página Dados';
      return;
    }
    const rows = getMapaAcertosFilteredRows({ testOnly: false });
    renderMapaAcertosKpis(rows);
    renderMapaAcertosTable(rows);
    const empty = document.getElementById('maEmpty');
    const msg = document.getElementById('maEmptyMsg');
    if (!rows.length && empty) {
      empty.style.display = 'block';
      const selCats = typeof msGet === 'function' ? msGet('maCat') : new Set();
      const selSubs = typeof msGet === 'function' ? msGet('maSub') : new Set();
      if (msg) {
        msg.textContent = selCats.size || selSubs.size
          ? 'Nenhuma linha com os filtros activos — limpa categoria/subcategoria'
          : 'Sem acertos (sem gaps SAP↔Unilog balanceáveis no EAN, nem sobra SAP fora do 0044 para mover).';
      }
    }
  }

  function maDoSort(col) {
    if (typeof doSort === 'function') doSort('ma', col);
  }

  global.MAPA_ACERTOS_JS_VERSION = MAPA_ACERTOS_JS_VERSION;
  global.MA_TEST_MATERIALS = MA_TEST_MATERIALS;
  global.MAPA_ACERTOS_TEST_MATERIALS = MA_TEST_MATERIALS;
  global.MAPA_ACERTOS_TEST_MATERIAL = MAPA_ACERTOS_TEST_MATERIAL;
  global.MAPA_ACERTOS_DEPOTS = MAPA_ACERTOS_DEPOTS;
  global.buildMapaAcertosRows = buildMapaAcertosRows;
  global.getMapaAcertosFilteredRows = getMapaAcertosFilteredRows;
  global.renderMapaAcertosPage = renderMapaAcertosPage;
  global.exportMapaAcertosXlsx = exportMapaAcertosXlsx;
  global.mapaAcertosExportFilename = mapaAcertosExportFilename;
  global.maExportPatchedXlsx = maExportPatchedXlsx;
  global.maPatchTemplateXlsx = maPatchTemplateXlsx;
  global.maPatchTemplateXlsxSheetJS = maPatchTemplateXlsxSheetJS;
  global.loadMapaAcertosTemplate = loadMapaAcertosTemplate;
  global.maDownloadXlsx = maDownloadXlsx;
  global.maDoSort = maDoSort;
  global.maBalanceGaps = maBalanceGaps;
  global.maMoveExcessTo0044 = maMoveExcessTo0044;
  global.maValidateNetZero = maValidateNetZero;
  global.maEnsureFflate = maEnsureFflate;
  global.maGetSapPrice = maGetSapPrice;
  global.maTestMaterialsLabel = maTestMaterialsLabel;
  global.MAPA_ACERTOS_DEPOT_SOBRA = MAPA_ACERTOS_DEPOT_SOBRA;
  global.MAPA_ACERTOS_DATA_START_ROW = DATA_START_ROW;
  global.MAPA_ACERTOS_DATA_END_ROW = DATA_END_ROW;
  global.MAPA_ACERTOS_TEMPLATE_URL = TEMPLATE_URL;
})(typeof window !== 'undefined' ? window : globalThis);
