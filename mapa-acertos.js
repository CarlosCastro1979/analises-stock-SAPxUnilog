/* Mapa Acertos — doc. acerto sem valor (redistribuição EAN neto zero; cx→0044 só se ainda fora) */
(function (global) {
  'use strict';
  const MAPA_ACERTOS_JS_VERSION = '1.0.16';
  /** @deprecated test filter removed — full SAP-stock scope. Kept empty for old callers. */
  const MA_TEST_MATERIALS = [];
  /** @deprecated */
  const MAPA_ACERTOS_TEST_MATERIAL = '';
  const MAPA_ACERTOS_DEPOTS = ['8', '9', '11', '22', '44'];
  const MAPA_ACERTOS_SHEETS = { 8: '0008', 9: '0009', 11: '0011', 22: '0022', 44: '0044' };
  /** Destino desejado para caixas de papelão (sem WMS): consolidar SAP fora deste dep. */
  const MAPA_ACERTOS_DEPOT_SOBRA = '44';
  const MA_CAIXAS_CAT = (typeof CAIXAS_PAPELAO_CAT !== 'undefined') ? CAIXAS_PAPELAO_CAT : 'Caixas de Papelão';
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
   * Acerto sem valor (semValorExcel):
   *   I Física = alvo pós-acerto (≈ Unilog/WMS; WMS não aparece) · J Contabilístico = SAP actual ·
   *   K Existência Valorizadas = Física × preço SAP · L Subir · M Descer · N Dif Valorizadas · O Justif.
   *   Contabilístico + Subir − Descer = Física (acerto pelo físico).
   * Inventário acerto (classic): I Física = Unilog · J Contabilístico = SAP · K = Física × preço.
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
   * Conservation rule: never create/destroy SAP qty — only move between depots/lotes of the same EAN.
   * Checks: Σ adj = 0, Σ subir = Σ descer, Σ SAP after = Σ SAP before (per EAN and global),
   * and Σ Dif. Valorizadas = 0 (preço SAP).
   */
  function maValidateNetZero(rows) {
    const byEan = {};
    var netQty = 0;
    var netVal = 0;
    var sumSubir = 0;
    var sumDescer = 0;
    var sapBefore = 0;
    var sapAfter = 0;
    const issues = [];
    for (var i = 0; i < rows.length; i++) {
      const r = rows[i];
      const ean = r.EAN_norm || '';
      const adj = Number(r.adj) || 0;
      const subir = Number(r.subir) > 0 ? Number(r.subir) : (adj > 0 ? adj : 0);
      const descer = Number(r.descer) > 0 ? Number(r.descer) : (adj < 0 ? -adj : 0);
      const before = Number(r.qt_sap) || 0;
      const after = r.qt_sap_after != null ? Number(r.qt_sap_after) : before + subir - descer;
      const preco = Number(r.preco) > 0 ? Number(r.preco) : maGetSapPrice(ean, r.material);
      const val = adj * (Number(preco) || 0);
      if (!byEan[ean]) byEan[ean] = { qty: 0, val: 0, sapBefore: 0, sapAfter: 0, subir: 0, descer: 0 };
      byEan[ean].qty += adj;
      byEan[ean].val += val;
      byEan[ean].sapBefore += before;
      byEan[ean].sapAfter += after;
      byEan[ean].subir += subir;
      byEan[ean].descer += descer;
      netQty += adj;
      netVal += val;
      sumSubir += subir;
      sumDescer += descer;
      sapBefore += before;
      sapAfter += after;
    }
    if (Math.abs(netQty) > MA_BALANCE_EPS) {
      issues.push('neto qty global = ' + maRound3(netQty) + ' (Σ subir deve = Σ descer)');
    }
    if (Math.abs(sumSubir - sumDescer) > MA_BALANCE_EPS) {
      issues.push('Σ subir (' + maRound3(sumSubir) + ') ≠ Σ descer (' + maRound3(sumDescer) + ')');
    }
    if (Math.abs(sapAfter - sapBefore) > MA_BALANCE_EPS) {
      issues.push('Σ SAP depois (' + maRound3(sapAfter) + ') ≠ Σ SAP antes (' + maRound3(sapBefore) + ')');
    }
    if (Math.abs(netVal) > MA_BALANCE_EPS) {
      issues.push('neto valor (Σ Dif. Valorizadas) = ' + maRound3(netVal));
    }
    Object.keys(byEan).forEach(function (ean) {
      const b = byEan[ean];
      if (Math.abs(b.qty) > MA_BALANCE_EPS) {
        issues.push('EAN ' + ean + ' neto qty = ' + maRound3(b.qty));
      }
      if (Math.abs(b.sapAfter - b.sapBefore) > MA_BALANCE_EPS) {
        issues.push('EAN ' + ean + ' Σ SAP depois (' + maRound3(b.sapAfter) + ') ≠ antes (' + maRound3(b.sapBefore) + ')');
      }
      if (Math.abs(b.subir - b.descer) > MA_BALANCE_EPS) {
        issues.push('EAN ' + ean + ' subir (' + maRound3(b.subir) + ') ≠ descer (' + maRound3(b.descer) + ')');
      }
      if (Math.abs(b.val) > MA_BALANCE_EPS) {
        issues.push('EAN ' + ean + ' neto valor = ' + maRound3(b.val));
      }
    });
    return {
      ok: issues.length === 0,
      netQty: netQty,
      netVal: netVal,
      sumSubir: sumSubir,
      sumDescer: sumDescer,
      sapBefore: sapBefore,
      sapAfter: sapAfter,
      byEan: byEan,
      issues: issues,
    };
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
   * Balance gaps within one EAN. Primary key = depot totals (or any entries with .key/.gap).
   * gap = Unilog - SAP. (+) subir SAP; (-) descer SAP. Neto of adjustments = 0.
   * Only transfers when excess in some entry can cover shortage in another.
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

  /** Caixas de papelão: sem WMS — físico desejado = SAP no 0044. */
  function maIsCaixasPapelao(ean, desc, cat) {
    if (cat && String(cat) === MA_CAIXAS_CAT) return true;
    if (typeof isCaixasPapelao === 'function') return !!isCaixasPapelao(ean, desc);
    if (desc && /\bPAPEL[ÃA]O\b|\bCAIXA[S]?\s+DE\s+PAPEL|\bCX\.?\s+PAPEL/i.test(String(desc))) return true;
    return false;
  }
  /**
   * Só para caixas (físico desejado = 0044): SAP ainda fora do 0044 → DESCER source + SUBIR 0044.
   * Excess já no 0044 fica. NÃO usar para artigos com WMS — sobra residual fica para inventário.
   * adjDepot keyed by depot id ('8','9',...). Mutates adjDepot.
   */
  function maMoveExcessTo0044Depot(depotAgg, adjDepot) {
    const deps = Object.keys(depotAgg);
    for (var i = 0; i < deps.length; i++) {
      const depot = deps[i];
      if (depot === MAPA_ACERTOS_DEPOT_SOBRA) continue;
      const d = depotAgg[depot];
      if (!d) continue;
      const adj = adjDepot[depot] || 0;
      // Cx: físico desejado é 0044 — move SAP residual fora do 0044 (vs uni local, tipicamente 0).
      const sapAfter = (Number(d.qt_sap) || 0) + adj;
      const excess = sapAfter - (Number(d.qt_uni) || 0);
      if (!(excess > 0)) continue;
      adjDepot[depot] = adj - excess;
      adjDepot[MAPA_ACERTOS_DEPOT_SOBRA] = (adjDepot[MAPA_ACERTOS_DEPOT_SOBRA] || 0) + excess;
      if (!depotAgg[MAPA_ACERTOS_DEPOT_SOBRA]) {
        depotAgg[MAPA_ACERTOS_DEPOT_SOBRA] = { key: MAPA_ACERTOS_DEPOT_SOBRA, depot: MAPA_ACERTOS_DEPOT_SOBRA, qt_sap: 0, qt_uni: 0, gap: 0 };
      }
    }
  }

  /**
   * Legacy lote-level excess→0044 (kept for callers/tests). Prefer maMoveExcessTo0044Depot.
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

  function maEnsureCell(g, depot, lote, seed) {
    const tKey = depot + '|' + (lote || '');
    if (!g.cells[tKey]) {
      g.cells[tKey] = {
        key: tKey,
        depot: depot,
        depotSheet: maDepotSheet(depot),
        lote: lote || '',
        material: (seed && seed.material) || g.material || '',
        desc: (seed && seed.desc) || g.desc || '',
        qt_sap: 0,
        qt_uni: 0,
        gap: 0,
      };
    }
    return g.cells[tKey];
  }

  /**
   * Split depot-level adj onto lote lines for Excel emission.
   * DESCER: take from lotes with SAP (prefer local excess sap>uni).
   * SUBIR: fill lotes with shortage (uni>sap), then dump remainder on best lote / empty.
   */
  function maAllocateDepotAdjToLotes(g, adjDepot) {
    const loteAdj = {};
    Object.keys(g.cells).forEach(function (k) { loteAdj[k] = 0; });
    const depots = Object.keys(adjDepot);
    for (var di = 0; di < depots.length; di++) {
      const depot = depots[di];
      var need = adjDepot[depot] || 0;
      if (!need) continue;
      var cells = Object.keys(g.cells).map(function (k) { return g.cells[k]; })
        .filter(function (c) { return c.depot === depot; });

      if (need < 0) {
        var remainDown = -need;
        cells = cells.slice().sort(function (a, b) {
          var ea = (Number(a.qt_sap) || 0) - (Number(a.qt_uni) || 0);
          var eb = (Number(b.qt_sap) || 0) - (Number(b.qt_uni) || 0);
          if (eb !== ea) return eb - ea;
          return (Number(b.qt_sap) || 0) - (Number(a.qt_sap) || 0);
        });
        for (var i = 0; i < cells.length && remainDown > MA_BALANCE_EPS; i++) {
          var avail = Math.max(0, Number(cells[i].qt_sap) || 0);
          var take = Math.min(remainDown, avail);
          if (take <= 0) continue;
          loteAdj[cells[i].key] = (loteAdj[cells[i].key] || 0) - take;
          remainDown -= take;
        }
        if (remainDown > MA_BALANCE_EPS && cells.length) {
          loteAdj[cells[0].key] = (loteAdj[cells[0].key] || 0) - remainDown;
          remainDown = 0;
        }
      } else {
        var remainUp = need;
        cells = cells.slice().sort(function (a, b) {
          var ga = (Number(a.qt_uni) || 0) - (Number(a.qt_sap) || 0);
          var gb = (Number(b.qt_uni) || 0) - (Number(b.qt_sap) || 0);
          return gb - ga;
        });
        for (var j = 0; j < cells.length && remainUp > MA_BALANCE_EPS; j++) {
          var gap = Math.max(0, (Number(cells[j].qt_uni) || 0) - (Number(cells[j].qt_sap) || 0));
          var put = Math.min(remainUp, gap);
          if (put <= 0) continue;
          loteAdj[cells[j].key] = (loteAdj[cells[j].key] || 0) + put;
          remainUp -= put;
        }
        if (remainUp > MA_BALANCE_EPS) {
          var dump = cells.find(function (c) { return (Number(c.qt_uni) || 0) > 0; }) || cells[0];
          if (!dump) {
            dump = maEnsureCell(g, depot, '', null);
            loteAdj[dump.key] = 0;
          }
          loteAdj[dump.key] = (loteAdj[dump.key] || 0) + remainUp;
          remainUp = 0;
        }
      }
    }
    return loteAdj;
  }

  /**
   * Recon exclui caixas do dep. 0011 (fora da dif. SAP×WMS). Para o mapa, o físico
   * desejado das caixas é 0044 — precisamos ver SAP em TODOS os deps (incl. 0011).
   * Soma SAP por dep|lote e aplica como fonte de verdade (evita double-count com recon).
   */
  function maAugmentCaixasSapIntoGroups(groups) {
    const sap = typeof sapData !== 'undefined' && sapData ? sapData : [];
    const sapSum = {}; // ean -> cellKey -> {dk,lote,qt,material,desc}
    for (var i = 0; i < sap.length; i++) {
      const r = sap[i];
      const dk = String(r.deposito || '');
      if (MAPA_ACERTOS_DEPOTS.indexOf(dk) < 0) continue;
      const ean = r.EAN_norm || '';
      if (!ean) continue;
      const qt = Number(r.qt) || 0;
      if (qt <= 0) continue;
      if (!maIsCaixasPapelao(ean, r.desc, '')) continue;
      const lote = r.Lote_norm || '';
      const cellKey = dk + '|' + lote;
      if (!sapSum[ean]) sapSum[ean] = {};
      if (!sapSum[ean][cellKey]) {
        sapSum[ean][cellKey] = {
          depot: dk,
          lote: lote,
          qt: 0,
          material: r.material ? String(r.material).trim() : '',
          desc: r.desc || '',
        };
      }
      sapSum[ean][cellKey].qt += qt;
      if (!sapSum[ean][cellKey].material && r.material) sapSum[ean][cellKey].material = String(r.material).trim();
      if (!sapSum[ean][cellKey].desc && r.desc) sapSum[ean][cellKey].desc = r.desc;
    }
    Object.keys(sapSum).forEach(function (ean) {
      const cells = sapSum[ean];
      if (!groups.has(ean)) {
        const first = cells[Object.keys(cells)[0]];
        const cs0 = maResolveCat(ean, '', '');
        groups.set(ean, {
          EAN_norm: ean,
          material: (first && first.material) || '',
          desc: (first && first.desc) || '',
          cat: cs0.cat || MA_CAIXAS_CAT,
          subcat: cs0.subcat || '',
          umb: maGetUmb(ean),
          cells: {},
          _caixas: true,
        });
      }
      const g = groups.get(ean);
      g._caixas = true;
      if (!g.cat) g.cat = MA_CAIXAS_CAT;
      Object.keys(cells).forEach(function (cellKey) {
        const s = cells[cellKey];
        if (!g.material && s.material) g.material = s.material;
        if (!g.desc && s.desc) g.desc = s.desc;
        const prev = g.cells[cellKey];
        if (prev) {
          prev.qt_sap = s.qt; // SAP raw = source of truth for caixas
          prev.gap = prev.qt_uni - prev.qt_sap;
          if (!prev.material && s.material) prev.material = s.material;
          if (!prev.desc && s.desc) prev.desc = s.desc;
        } else {
          g.cells[cellKey] = {
            key: cellKey,
            depot: s.depot,
            depotSheet: maDepotSheet(s.depot),
            lote: s.lote,
            material: s.material,
            desc: s.desc,
            qt_sap: s.qt,
            qt_uni: 0,
            gap: -s.qt,
          };
        }
      });
    });
  }

  function buildMapaAcertosRows(options) {
    const opts = options || {};
    /** Full scope by default. Pass testOnly:true only for legacy debug. */
    const testOnly = opts.testOnly === true;
    const filterMaterial = opts.material || '';
    const recon = typeof reconData !== 'undefined' ? reconData : {};
    // Group by EAN — balance at depot totals, emit by lote
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
            _caixas: maIsCaixasPapelao(ean, r.desc || '', cs0.cat || r.cat),
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
        if (!g._caixas) g._caixas = maIsCaixasPapelao(ean, g.desc || r.desc, g.cat || r.cat);
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

    // Caixas no 0011 estão fora do recon — injectar SAP directo para consolidar → 0044.
    maAugmentCaixasSapIntoGroups(groups);

    const out = [];
    for (const g of groups.values()) {
      const entries = Object.keys(g.cells).map(function (k) { return g.cells[k]; });
      var sapTotal = 0;
      for (var ei = 0; ei < entries.length; ei++) {
        sapTotal += Number(entries[ei].qt_sap) || 0;
      }
      if (sapTotal <= 0) continue;

      const isCaixas = !!(g._caixas || maIsCaixasPapelao(g.EAN_norm, g.desc, g.cat));

      // 1) Aggregate by depot — match Unilog depot totals (físico desejado = WMS)
      const depotAgg = {};
      for (var ci = 0; ci < entries.length; ci++) {
        const e = entries[ci];
        if (!depotAgg[e.depot]) {
          depotAgg[e.depot] = { key: e.depot, depot: e.depot, qt_sap: 0, qt_uni: 0, gap: 0 };
        }
        depotAgg[e.depot].qt_sap += Number(e.qt_sap) || 0;
        depotAgg[e.depot].qt_uni += Number(e.qt_uni) || 0;
      }
      Object.keys(depotAgg).forEach(function (dk) {
        const d = depotAgg[dk];
        d.gap = d.qt_uni - d.qt_sap;
      });

      // Já alinhado ao físico? (SAP == Unilog por dep.; cx já só no 0044)
      var needsAdj = false;
      if (isCaixas) {
        Object.keys(depotAgg).forEach(function (dk) {
          if (dk === MAPA_ACERTOS_DEPOT_SOBRA) return;
          if ((Number(depotAgg[dk].qt_sap) || 0) > MA_BALANCE_EPS) needsAdj = true;
        });
      } else {
        Object.keys(depotAgg).forEach(function (dk) {
          if (Math.abs(Number(depotAgg[dk].gap) || 0) > MA_BALANCE_EPS) needsAdj = true;
        });
      }
      if (!needsAdj) continue;

      const depotEntries = Object.keys(depotAgg).map(function (dk) { return depotAgg[dk]; });
      // Artigos com WMS: só transferências que aproximam SAP do Unilog (sem inventar destino 0044).
      const adjDepot = isCaixas ? {} : maBalanceGaps(depotEntries);
      if (!isCaixas) {
        Object.keys(depotAgg).forEach(function (dk) {
          if (adjDepot[dk] == null) adjDepot[dk] = 0;
        });
      } else {
        Object.keys(depotAgg).forEach(function (dk) { adjDepot[dk] = 0; });
      }
      const balanceDepot = {};
      Object.keys(adjDepot).forEach(function (k) { balanceDepot[k] = adjDepot[k]; });
      // 2) Caixas sem WMS: físico desejado = 0044 — consolidar SAP ainda fora do 0044.
      //    NÃO despejar sobra residual de artigos com Unilog (isso inventava "Subir 0044"
      //    mesmo quando 0044/WMS já estavam correctos).
      if (isCaixas) {
        maMoveExcessTo0044Depot(depotAgg, adjDepot);
      }
      // Sem movimento real → omitir (SAP já = físico desejado)
      var hasMove = Object.keys(adjDepot).some(function (k) {
        return Math.abs(Number(adjDepot[k]) || 0) > MA_BALANCE_EPS;
      });
      if (!hasMove) continue;
      if (!maFixAdjMapNetZero(adjDepot)) {
        try {
          console.error('[mapa-acertos] EAN ' + g.EAN_norm + ' adj depot desequilibrado — omitido', adjDepot);
        } catch (e) {}
        continue;
      }
      // 3) Allocate depot adj onto lotes for emission
      const adjMap = maAllocateDepotAdjToLotes(g, adjDepot);
      if (!maFixAdjMapNetZero(adjMap)) {
        try {
          console.error('[mapa-acertos] EAN ' + g.EAN_norm + ' adj lote desequilibrado — omitido', adjMap);
        } catch (e) {}
        continue;
      }
      const finalEntries = Object.keys(g.cells).map(function (k) { return g.cells[k]; });
      const preco = maGetSapPrice(g.EAN_norm, g.material || filterMaterial);
      for (const e of finalEntries) {
        const adj = adjMap[e.key] || 0;
        if (!adj) continue;
        const subir = adj > 0 ? adj : 0;
        const descer = adj < 0 ? -adj : 0;
        const qt_sap_after = (Number(e.qt_sap) || 0) + subir - descer;
        const sobra0044 = isCaixas && e.depot === MAPA_ACERTOS_DEPOT_SOBRA
          ? (adjDepot[MAPA_ACERTOS_DEPOT_SOBRA] || 0) > 0
          : false;
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
          qt_sap_after: qt_sap_after,
          gap: e.gap,
          adj: adj,
          subir: subir,
          descer: descer,
          sobra0044: !!sobra0044,
          hadBalance: !!(balanceDepot[e.depot]),
          /** Excel: Contabilístico=SAP actual, Física=alvo (WMS só nos bastidores). */
          semValorExcel: true,
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
  /** Dif (+)/(−): Física > Contabilístico → Subir (L); Física < Contabilístico → Descer (M). */
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
   * Excel qty mapping.
   * Acerto sem valor: Contabilístico = SAP actual; Física = alvo pós-acerto (≈ Unilog/WMS;
   *   WMS qty never written to the sheet). Contabilístico + Subir − Descer = Física.
   * Inventário: Física = Unilog; Contabilístico = SAP (classic count vs book).
   */
  function maRowExcelQtys(row) {
    const dif = maDifParts(row);
    const subir = dif.subir;
    const descer = dif.descer;
    const qtSap = maRound3(row.qt_sap);
    const qtUni = maRound3(row.qt_uni);
    const sapAfter = row.qt_sap_after != null
      ? maRound3(row.qt_sap_after)
      : maRound3(qtSap + subir - descer);
    if (row.semValorExcel) {
      return { fisica: sapAfter, contabilistico: qtSap, subir: subir, descer: descer };
    }
    return { fisica: qtUni, contabilistico: qtSap, subir: subir, descer: descer };
  }
  /** Headers for acerto sem valor — WMS not named; Contabilístico = SAP; Física = alvo. */
  function maPatchAcertoHeadersSemValor(sheetXml) {
    function hdr(addr, fallbackStyle, text) {
      return maCellText(addr, maExtractStyle(sheetXml, addr, fallbackStyle), text);
    }
    sheetXml = maReplaceCell(sheetXml, 'I7', hdr('I7', '20', 'Existência'));
    sheetXml = maReplaceCell(sheetXml, 'I8', hdr('I8', '20', 'Física'));
    sheetXml = maReplaceCell(sheetXml, 'J7', hdr('J7', '20', 'Inventário'));
    sheetXml = maReplaceCell(sheetXml, 'J8', hdr('J8', '20', 'Contabilístico'));
    sheetXml = maReplaceCell(sheetXml, 'K7', hdr('K7', '20', 'Existência'));
    sheetXml = maReplaceCell(sheetXml, 'K8', hdr('K8', '20', 'Valorizadas'));
    sheetXml = maReplaceCell(sheetXml, 'L7', hdr('L7', '20', 'Diferenças'));
    sheetXml = maReplaceCell(sheetXml, 'L8', hdr('L8', '20', 'Subir (+)'));
    sheetXml = maReplaceCell(sheetXml, 'M8', hdr('M8', '20', 'Descer (−)'));
    sheetXml = maReplaceCell(sheetXml, 'N7', hdr('N7', '20', 'Diferenças'));
    sheetXml = maReplaceCell(sheetXml, 'N8', hdr('N8', '20', 'Valorizadas'));
    return sheetXml;
  }
  /**
   * Fill one data row — always MA_CLEAN_STYLES (no green fill).
   * See maRowExcelQtys for Física / Contabilístico semantics.
   */
  function maFillDataRow(sheetXml, rowNum, row) {
    const styles = MA_CLEAN_STYLES;
    const preco = maRound3(row.preco > 0 ? row.preco : maGetSapPrice(row.EAN_norm, row.material));
    const q = maRowExcelQtys(row);
    const existVal = preco > 0 ? maRound3(q.fisica * preco) : null;
    const difVal = preco > 0 ? maRound3((q.subir - q.descer) * preco) : null;
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
    sheetXml = maReplaceCell(sheetXml, 'I' + rowNum, maCellNum('I' + rowNum, styles.I, q.fisica));
    sheetXml = maReplaceCell(sheetXml, 'J' + rowNum, maCellNum('J' + rowNum, styles.J, q.contabilistico));
    sheetXml = maReplaceCell(sheetXml, 'K' + rowNum, existVal != null ? maCellNum('K' + rowNum, styles.K, existVal) : maCellEmpty('K' + rowNum, styles.K));
    sheetXml = maReplaceCell(sheetXml, 'L' + rowNum, q.subir > 0 ? maCellNum('L' + rowNum, styles.L, q.subir) : maCellEmpty('L' + rowNum, styles.L));
    sheetXml = maReplaceCell(sheetXml, 'M' + rowNum, q.descer > 0 ? maCellNum('M' + rowNum, styles.M, q.descer) : maCellEmpty('M' + rowNum, styles.M));
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
    if (dataRows.some(function (r) { return r.semValorExcel; })) {
      sheetXml = maPatchAcertoHeadersSemValor(sheetXml);
    }
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
    const q = maRowExcelQtys(row);
    const note = maJustification(row);
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
    maSetCell(ws, 'I' + rowNum, { t: 'n', v: q.fisica });
    maSetCell(ws, 'J' + rowNum, { t: 'n', v: q.contabilistico });
    if (preco > 0) maSetCell(ws, 'K' + rowNum, { t: 'n', v: maRound3(q.fisica * preco) });
    else delete ws['K' + rowNum];
    if (q.subir > 0) maSetCell(ws, 'L' + rowNum, { t: 'n', v: q.subir });
    else delete ws['L' + rowNum];
    if (q.descer > 0) maSetCell(ws, 'M' + rowNum, { t: 'n', v: q.descer });
    else delete ws['M' + rowNum];
    if (preco > 0) maSetCell(ws, 'N' + rowNum, { t: 'n', v: maRound3((q.subir - q.descer) * preco) });
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
      const dataRows = depotRowsMap[dk] || [];
      if (dataRows.some(function (r) { return r.semValorExcel; })) {
        maSetCell(ws, 'I7', { t: 's', v: 'Existência' });
        maSetCell(ws, 'I8', { t: 's', v: 'Física' });
        maSetCell(ws, 'J7', { t: 's', v: 'Inventário' });
        maSetCell(ws, 'J8', { t: 's', v: 'Contabilístico' });
        maSetCell(ws, 'K7', { t: 's', v: 'Existência' });
        maSetCell(ws, 'K8', { t: 's', v: 'Valorizadas' });
        maSetCell(ws, 'L7', { t: 's', v: 'Diferenças' });
        maSetCell(ws, 'L8', { t: 's', v: 'Subir (+)' });
        maSetCell(ws, 'M8', { t: 's', v: 'Descer (−)' });
        maSetCell(ws, 'N7', { t: 's', v: 'Diferenças' });
        maSetCell(ws, 'N8', { t: 's', v: 'Valorizadas' });
      }
      ['K129', 'N129'].forEach(function (addr) {
        const cell = ws[addr];
        if (cell && cell.f) {
          cell.f = String(cell.f)
            .replace(/K10:K128/g, 'K' + DATA_START_ROW + ':K' + DATA_END_ROW)
            .replace(/N10:N128/g, 'N' + DATA_START_ROW + ':N' + DATA_END_ROW);
        }
      });
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
          : 'Sem acertos (sem gaps balanceáveis SAP↔Unilog; caixas já no 0044 omitidas)',
        'info'
      );
      renderMapaAcertosPage();
      return;
    }
    const bal = maValidateNetZero(rows);
    if (!bal.ok) {
      try { console.error('[mapa-acertos] export bloqueado — balanceamento inválido', bal); } catch (e) {}
      toast('Export bloqueado: total SAP por EAN deve conservar-se (só transferências entre depósitos/lotes). ' + bal.issues.slice(0, 3).join('; '), 'error');
      renderMapaAcertosPage();
      return;
    }
    try {
      const result = await maExportPatchedXlsx(rows, mapaAcertosExportFilename());
      const subir = rows.reduce(function (s, r) { return s + r.subir; }, 0);
      const descer = rows.reduce(function (s, r) { return s + r.descer; }, 0);
      var msg = 'Exportado: ' + result.written + ' linhas · subir ' + subir + ' = descer ' + descer + ' · Σ SAP conservado';
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
      '<div class="kpi"><div class="kl">Subir (Σ +)</div><div class="kv y">' + fmt(subir) + '</div><div class="ks">entrada noutro dep/lote</div></div>' +
      '<div class="kpi"><div class="kl">Descer (Σ −)</div><div class="kv r">' + fmt(descer) + '</div><div class="ks">saída deste dep/lote</div></div>' +
      '<div class="kpi"><div class="kl">Σ SAP / neto</div><div class="kv' + netCls + '">' + fmt(maRound3(bal.sapBefore)) + ' → ' + fmt(maRound3(bal.sapAfter)) + '</div><div class="ks">neto ' + fmt(maRound3(bal.netQty)) + ' (deve 0) · ' + fmt(deps) + ' deps</div></div>';
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
    body.innerHTML = sorted.map(function (r) {
      const subir = Number(r.subir) || 0;
      const descer = Number(r.descer) || 0;
      const alvo = r.qt_sap_after != null ? r.qt_sap_after : (Number(r.qt_sap) || 0) + subir - descer;
      const acao = r.adj > 0
        ? '<span class="badge binfo">Subir ' + num(r.adj) + '</span>'
        : '<span class="badge bwarn">Descer ' + num(-r.adj) + '</span>';
      return '<tr>' +
        '<td><span class="badge bgray">' + r.depotSheet + '</span></td>' +
        '<td style="font-family:var(--mono);font-size:11px">' + r.material + '</td>' +
        '<td style="font-family:var(--mono);font-size:11px">' + (r.EAN_norm || '—') + '</td>' +
        '<td class="col-desc">' + (r.desc || '—') + '</td>' +
        '<td style="font-family:var(--mono);font-size:11px">' + (r.lote || '—') + '</td>' +
        '<td class="num">' + num(alvo) + '</td>' +
        '<td class="num">' + num(r.qt_sap) + '</td>' +
        '<td class="num">' + num(subir) + '</td>' +
        '<td class="num">' + num(descer) + '</td>' +
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
          : 'Sem acertos (SAP já = físico desejado por depósito; caixas já no 0044 omitidas; sobras sem destino WMS ficam para Acerto Inventário).';
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
  global.maMoveExcessTo0044Depot = maMoveExcessTo0044Depot;
  global.maIsCaixasPapelao = maIsCaixasPapelao;
  global.maAllocateDepotAdjToLotes = maAllocateDepotAdjToLotes;
  global.maValidateNetZero = maValidateNetZero;
  global.maEnsureFflate = maEnsureFflate;
  global.maGetSapPrice = maGetSapPrice;
  global.maTestMaterialsLabel = maTestMaterialsLabel;
  global.MAPA_ACERTOS_DEPOT_SOBRA = MAPA_ACERTOS_DEPOT_SOBRA;
  global.MAPA_ACERTOS_DATA_START_ROW = DATA_START_ROW;
  global.MAPA_ACERTOS_DATA_END_ROW = DATA_END_ROW;
  global.MAPA_ACERTOS_TEMPLATE_URL = TEMPLATE_URL;
})(typeof window !== 'undefined' ? window : globalThis);
