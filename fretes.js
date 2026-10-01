// fretes.js v1.8.91 — restore: FORCE AP + scrub shared ghost faces + persist; multi-CT-e pago once
const FRETES_JS_VERSION = '1.8.91';
/** Unilog ops started ~10 Nov 2025 — coverage table ignores months before this (YYYY-MM). */
const FRETES_OPS_START_MES = '2025-11';
/**
 * Bump when ZFACT valorNF/dtEmissao column semantics change.
 * v9: on restore, FORCE AP from stored docVal + scrub shared ghost faces (docVal=0, same face on ≥3 NFs);
 *     auto-persist corrected map (no manual Processar). Multi-CT-e: identical pago rows → take once.
 * v8: FORCE valorNF = Val.total incl.imp. once/NF (never lose to modest lineSum gap, e.g. 449796 vs 442314).
 * v7: cloud maps may still hold SUM(face)×lines (NF 99977≈11.2M) — sanitize on load; user must Processar ZFACT.
 * (Do NOT auto-reparse multi-MB Excel on open — Chrome OOM.)
 */
const SAP_NF_MAP_PARSER_VERSION = 9;

/** Max JSON bytes before base64 (~6 MB raw → ~8 MB b64 in Supabase text column). */
const QZ_PERSIST_MAX_JSON_BYTES = 6 * 1024 * 1024;

const sb = db;

const $ = id => document.getElementById('fte-' + id);

let fteInited = false;
let sapNfMap = {};
/** Provenance of merged ZFACT loads (persisted with sap map pack). */
let sapMapSourceFiles = [];
let lastCtePack = null;

let currentNFs = [];
let currentSummary = null;
let activeSubPanel = null;
let currentUploadId = null;
let isSaving = false;
let tableSort = { col: 'diff', dir: -1 };
let subTableSort = { col: 'diff', dir: -1 };
let monthSort = { col: 'mesRef', dir: 1 };
let selectedMonth = '';
let monthlyRows = [];
let _fteSkipAutosave = false;
let _fteLoadedCompany = null;
let activeCteSub = 'resumo';
let fteCteFileName = '';
let fteCteBuffer = null;
let fteSapFileName = '';
let fteSapBuffer = null;
let fteSapMapNeedsReprocess = false;
let fteQzPendingFiles = [];
let quinzenalPack = null;
let _fteLoadSavedPromise = null;
let qzExpandedMonths = new Set();
/** 'conciliacao' | 'quinzenal' | null — drives Análise CT-e source + CT-e vs QZ confronto behaviour */
let cteAnalysisSource = null;

function hasConciliacaoCte() {
  return cteAnalysisSource === 'conciliacao' && currentNFs.length > 0;
}

function hasCteAnalysisData() {
  return currentNFs.length > 0;
}

/** Yield to the browser so menus stay responsive during heavy Fretes restore. */
function fteYield(label) {
  if (label) {
    try { fteSetProcessing(true, label); } catch (_) {}
  }
  return new Promise(resolve => {
    try {
      requestAnimationFrame(() => setTimeout(resolve, 0));
    } catch (_) {
      setTimeout(resolve, 0);
    }
  });
}

function fteMarkCompanyLoaded() {
  _fteLoadedCompany = fteCompany();
}

/** QZ pack and/or Análise CT-e rows — Resumo/Análise need this; SAP map alone is not enough. */
function fteHasAnalysisData() {
  return !!(quinzenalPack?.files?.length || currentNFs.length);
}

function fteCompany() {
  return typeof company !== 'undefined' ? company : 'DFB';
}

/** Wait until Supabase probe finished (avoid silent empty restore while still "conectando…"). */
async function fteWaitForDb(maxMs = 8000) {
  const start = Date.now();
  while (Date.now() - start < maxMs) {
    const lbl = document.getElementById('dbLbl')?.textContent || '';
    if (/Supabase OK/i.test(lbl)) return true;
    if (/offline/i.test(lbl)) return false;
    await fteYield();
  }
  const lbl = document.getElementById('dbLbl')?.textContent || '';
  return /Supabase OK/i.test(lbl);
}

/** Show loading inside the visible analysis empty panel (Carregamento spinner is hidden when tab≠carregamento). */
function fteShowAnalysisLoading(tab, label) {
  const msg = label || 'A carregar Fretes da cloud…';
  const html = `<div class="empty-i"><span class="spin" style="width:22px;height:22px;border-width:3px"></span></div><div><strong>${msg}</strong></div><div style="font-size:11px;margin-top:8px">Quinzenais + mapa SAP — não é preciso voltar a carregar os Excels.</div>`;
  const map = {
    'analise-cte': 'cteEmpty',
    'analise-b2c': 'b2cEmpty',
    'cte-vs-qz': 'b2bEmpty',
    'resumo-total': 'resumoEmpty'
  };
  const emptyId = map[tab];
  if (!emptyId) return;
  const empty = $(emptyId);
  const contentMap = {
    cteEmpty: 'cteContent',
    b2cEmpty: 'b2cContent',
    b2bEmpty: 'b2bContent',
    resumoEmpty: 'resumoContent'
  };
  const content = $(contentMap[emptyId]);
  if (empty) {
    empty.style.display = 'block';
    empty.innerHTML = html;
  }
  if (content) content.style.display = 'none';
}

function fteCteSlot() {
  return typeof EXCEL_SLOTS !== 'undefined' ? EXCEL_SLOTS.FRETES_CTE : 'fretes_cte';
}

function fteSapSlot() {
  return typeof EXCEL_SLOTS !== 'undefined' ? EXCEL_SLOTS.FRETES_SAP_NF : 'fretes_sap_nf';
}

/** Accumulated SAP NF map (JSON) — history across partial ZFACT uploads. */
function fteSapMapSlot() {
  return typeof EXCEL_SLOTS !== 'undefined' ? EXCEL_SLOTS.FRETES_SAP_MAP : 'fretes_sap_map';
}

function fteQuinzenalSlot() {
  return typeof EXCEL_SLOTS !== 'undefined' ? EXCEL_SLOTS.FRETES_QUINZENAL : 'fretes_quinzenal';
}

function setCteZoneLoaded(name) {
  const fn = $('cteFn');
  const zone = $('cteZone');
  if (fn) fn.textContent = name ? '✓ ' + name : '';
  zone?.classList.toggle('loaded', !!name);
  checkFteBtn();
}

function setSapZoneLoaded(name) {
  const fn = $('sapFn');
  const zone = $('sapZone');
  if (name) fteSapFileName = name;
  if (fn) fn.textContent = name ? '✓ ' + name : '';
  zone?.classList.toggle('loaded', !!name);
  invalidateZfactMonthsCache();
  checkFteBtn();
}

function fteSetProcessing(active, label) {
  const spin = $('procSpin');
  const btn = $('procBtn');
  const loadBtn = $('loadLastBtn');
  const note = $('procNote');
  if (spin) {
    spin.hidden = !active;
    spin.style.display = active ? 'inline-flex' : 'none';
  }
  if (note) {
    if (active && label) {
      note.textContent = label;
      note.hidden = false;
      note.style.display = 'inline';
    } else {
      note.textContent = '';
      note.hidden = true;
      note.style.display = 'none';
    }
  }
  if (btn) btn.disabled = active || !(fteCteBuffer || fteSapBuffer || fteQzPendingFiles.length || quinzenalPack?.files?.length);
  if (loadBtn) loadBtn.disabled = !!active;
}

function checkFteBtn() {
  const btn = $('procBtn');
  if (btn) btn.disabled = !(fteCteBuffer || fteSapBuffer || fteQzPendingFiles.length || quinzenalPack?.files?.length);
}

const FTE_TAB_IDS = ['carregamento', 'analise-cte', 'analise-b2c', 'cte-vs-qz', 'resumo-total'];

const FTE_ANALYSIS_TABS = new Set(['analise-cte', 'analise-b2c', 'cte-vs-qz', 'resumo-total']);

function switchFteTab(tab) {
  document.querySelectorAll('.fte-tab').forEach(x => {
    x.classList.toggle('active', x.dataset.tab === tab);
  });
  FTE_TAB_IDS.forEach(id => {
    const el = $('tab-' + id);
    if (el) el.style.display = id === tab ? 'block' : 'none';
  });
  // Analysis tabs need cloud data — lazy restore (never on Carregamento alone).
  if (FTE_ANALYSIS_TABS.has(tab) && typeof fteNeedsCloudReload === 'function' && fteNeedsCloudReload()) {
    fteSetProcessing(true, 'A carregar Fretes…');
    fteShowAnalysisLoading(tab, 'A carregar Fretes da cloud…');
    const paintThenLoad = () => {
      loadSavedFretesFiles(true).then(() => {
        if (tab === 'analise-cte') renderCteSubPanels();
        else if (tab === 'analise-b2c') renderB2cAnalysisTab();
        else if (tab === 'cte-vs-qz') renderB2bCompareTab();
        else if (tab === 'resumo-total') renderResumoTotal();
        // SAP map alone is not enough for Resumo/Análise — surface failure even if ok===true.
        if (!fteHasAnalysisData()) {
          fteToastError('Quinzenais não restaurados da cloud — usa «Carregar último guardado» em Carregamento.');
        }
      }).catch(e => {
        console.warn('[fretes] lazy tab load', e);
        fteToastError('Erro ao carregar Fretes: ' + (e?.message || e));
      })
        .finally(() => { try { fteSetProcessing(false); } catch (_) {} });
    };
    try { requestAnimationFrame(() => setTimeout(paintThenLoad, 0)); }
    catch (_) { setTimeout(paintThenLoad, 0); }
    return;
  }
  if (tab === 'analise-cte') renderCteSubPanels();
  if (tab === 'analise-b2c') renderB2cAnalysisTab();
  if (tab === 'cte-vs-qz') renderB2bCompareTab();
  if (tab === 'resumo-total') renderResumoTotal();
}

function switchCteSub(sub) {
  activeCteSub = sub || 'resumo';
  document.querySelectorAll('.fte-cte-subtabs .qz-subtab').forEach(x => {
    x.classList.toggle('active', x.dataset.cteSub === activeCteSub);
  });
  renderCteSubPanels();
  if (activeCteSub === 'mensal') renderMonthlyTable();
}

function renderCteSubPanels() {
  const hasData = currentNFs.length > 0;
  const empty = $('cteEmpty');
  const content = $('cteContent');
  if (empty) {
    empty.style.display = hasData ? 'none' : 'block';
    if (!hasData) {
      empty.innerHTML = '<div class="empty-i">📊</div><div>Carrega <strong>quinzenais B2B</strong> + <strong>ZFACT</strong> na aba <strong>Carregamento de dados</strong></div>';
    }
  }
  if (content) content.style.display = hasData ? 'block' : 'none';
  if (!hasData) return;
  ['resumo', 'detalhe', 'anomalias', 'mensal'].forEach(sub => {
    const el = $('cte-sub-' + sub);
    if (el) el.style.display = activeCteSub === sub ? 'block' : 'none';
  });
  if (activeCteSub === 'mensal') renderMonthlyTable();
}

async function applyFretesFileLabelsFromMeta() {
  // Metadata only — never select file_data (QZ packs are multi-MB and freeze the UI).
  const fetchMeta = typeof fetchExcelFileMeta === 'function' ? fetchExcelFileMeta : fetchExcelFiles;
  if (typeof fetchMeta !== 'function') return;
  try {
    if (typeof syncCompanyDepotMap === 'function') syncCompanyDepotMap();
    const m = await fetchMeta([fteCteSlot(), fteSapSlot(), fteQuinzenalSlot()]);
    const cte = m[fteCteSlot()];
    const sap = m[fteSapSlot()];
    if (fteCteBuffer || cte?.file_name) setCteZoneLoaded(fteCteFileName || cte?.file_name || '');
    else setCteZoneLoaded('');
    if (fteSapBuffer || sap?.file_name) {
      if (!fteSapFileName && sap?.file_name) fteSapFileName = sap.file_name;
      setSapZoneLoaded(fteSapFileName || sap?.file_name || '');
    } else setSapZoneLoaded('');
    syncQzUploadZone();
    updateQzFileNote();
    updateFretesFileStatus(m);
    renderLoadCoverageSummary();
  } catch (e) { /* offline */ }
}

function quinzenalPackCounts(pack) {
  if (!pack?.files?.length) return null;
  return {
    total: pack.files.length,
    b2c: pack.files.filter(f => f.canal === 'B2C').length,
    b2b: pack.files.filter(f => f.canal === 'B2B').length
  };
}

function quinzenalExcelStatusPart(rec) {
  // NEVER parseQuinzenalPackFromRec here — that JSON.parse of multi-MB packs freezes menus.
  const counts = quinzenalPackCounts(quinzenalPack);
  const dt = typeof fmtExcelFileDate === 'function' ? fmtExcelFileDate(rec?.uploaded_at) : '';
  if (counts?.total) {
    const n = counts.total;
    return `Quinzenais: ${n} ficheiro${n !== 1 ? 's' : ''}${dt ? ' (' + dt + ')' : ''}`;
  }
  if (rec?.file_name) return `Quinzenais: ${rec.file_name}${dt ? ' (' + dt + ')' : ''}`;
  return '';
}

function parseQuinzenalPackFromRec(rec) {
  if (!rec?.file_data || typeof base64ToArrayBuffer !== 'function') return null;
  try {
    const b64 = rec.file_data;
    rec.file_data = null; // drop base64 ASAP — pack JSON stays in quinzenalPack
    const pack = JSON.parse(new TextDecoder().decode(base64ToArrayBuffer(b64)));
    if (pack?.fileBinaries) delete pack.fileBinaries; // never keep 38 xlsx binaries in RAM
    return pack;
  } catch (err) {
    console.error('[fretes] parse quinzenal pack', err);
    return null;
  }
}

function persistDateValue(v) {
  if (v == null || v === '') return null;
  const d = parseSapBrDate(v);
  if (d) return dateToYmd(d);
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v.trim())) return v.trim().slice(0, 10);
  return v;
}

function slimB2bRowForPersist(r) {
  if (!r) return r;
  return {
    nf: r.nf, nfKey: r.nfKey, valorNF: r.valorNF, pago: r.pago, nCte: r.nCte,
    transportador: r.transportador, destinatario: r.destinatario,
    modalidade: r.modalidade || '',
    mesKey: r.mesKey, mesLabel: r.mesLabel,
    quinzenaKey: r.quinzenaKey, quinzenaLabel: r.quinzenaLabel, fileName: r.fileName,
    dtNF: persistDateValue(r.dtNF),
    dtNFQz: persistDateValue(r.dtNFQz || r.dtNF),
    temDevolucao: !!r.temDevolucao,
    ctes: Array.isArray(r.ctes) ? r.ctes.map(c => ({
      numCte: c.numCte, dtCte: persistDateValue(c.dtCte), pago: num(c.pago),
      devolucao: !!c.devolucao, tipoOp: c.tipoOp || '', peso: num(c.peso)
    })) : undefined
  };
}

function slimB2cRowForPersist(r) {
  if (!r) return r;
  return {
    nf: r.nf, pedido: r.pedido, dtColeta: persistDateValue(r.dtColeta) || r.dtColeta,
    dtNF: persistDateValue(r.dtNF) || r.dtNF,
    numCte: r.numCte, transportador: r.transportador, destinatario: r.destinatario,
    valorProdutos: r.valorProdutos, valorNF: r.valorNF, pago: r.pago, zona: r.zona,
    quinzenaKey: r.quinzenaKey, quinzenaLabel: r.quinzenaLabel, fileName: r.fileName,
    mesKey: r.mesKey, mesLabel: r.mesLabel
  };
}

function slimQuinzenalPackForPersist(pack) {
  if (!pack) return null;
  return {
    version: pack.version || 2,
    updatedAt: new Date().toISOString(),
    files: pack.files || [],
    failedFiles: pack.failedFiles || [],
    b2bRows: (pack.b2bRows || []).map(slimB2bRowForPersist),
    b2cRows: (pack.b2cRows || []).map(slimB2cRowForPersist)
  };
}

function fmtByteSize(n) {
  if (!n || n < 1024) return (n || 0) + ' B';
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
  return (n / (1024 * 1024)).toFixed(2) + ' MB';
}

/** Merge key for quinzenal rows: canal + NF (pedido fallback for B2C without NF). */
function qzRowMergeKey(r, canal) {
  const ch = canal || (/B2B/i.test(r?.fileName || '') ? 'B2B' : 'B2C');
  const nf = r?.nfKey || normNFKey(r?.nf);
  if (nf) return `${ch}|nf:${nf}`;
  const ped = r?.pedido != null && String(r.pedido).trim() !== '' ? String(r.pedido).trim() : '';
  if (ped) return `${ch}|ped:${ped}`;
  return '';
}

/**
 * Merge quinzenal packs without treating the new Excel as full truth.
 * - Same fileName → replace that file's rows (re-upload of same report).
 * - Else NF-level (canal|nf): update overlapping NFs; keep NFs absent from the new file.
 */
function mergeQuinzenalPacks(prev, next) {
  if (!prev?.files?.length) return next;
  if (!next?.files?.length) return prev;
  const replaceNames = new Set((next.files || []).map(f => f.fileName));
  const nextB2bKeys = new Set(
    (next.b2bRows || []).map(r => qzRowMergeKey(r, 'B2B')).filter(Boolean)
  );
  const nextB2cKeys = new Set(
    (next.b2cRows || []).map(r => qzRowMergeKey(r, 'B2C')).filter(Boolean)
  );
  const keepPrevRow = (r, canal, nextKeys) => {
    if (replaceNames.has(r.fileName)) return false;
    const k = qzRowMergeKey(r, canal);
    if (k && nextKeys.has(k)) return false;
    return true;
  };
  const b2bRows = [
    ...(prev.b2bRows || []).filter(r => keepPrevRow(r, 'B2B', nextB2bKeys)),
    ...(next.b2bRows || [])
  ];
  const b2cRows = [
    ...(prev.b2cRows || []).filter(r => keepPrevRow(r, 'B2C', nextB2cKeys)),
    ...(next.b2cRows || [])
  ];
  const usedNames = new Set(
    [...b2bRows, ...b2cRows].map(r => r.fileName).filter(Boolean)
  );
  const files = [
    ...(prev.files || []).filter(f => !replaceNames.has(f.fileName) && usedNames.has(f.fileName)),
    ...(next.files || [])
  ];
  const failKeys = new Set((next.failedFiles || []).map(f => f.fileName));
  const failedFiles = [
    ...(prev.failedFiles || []).filter(f => !failKeys.has(f.fileName)),
    ...(next.failedFiles || [])
  ];
  const pack = {
    version: 2,
    updatedAt: new Date().toISOString(),
    files, failedFiles, b2bRows, b2cRows
  };
  pack.b2bCompare = buildB2BCompare(b2bRows);
  pack.b2bMonthTotals = buildB2BMonthTotals(b2bRows);
  pack.b2bQuinzenaTotals = buildB2BQuinzenaTotals(b2bRows);
  pack.b2cMonthTotals = buildB2CMonthTotals(b2cRows);
  pack.b2cQuinzenaTotals = buildB2CQuinzenaTotals(b2cRows);
  pack.b2cRegionTotals = buildB2CRegionTotals(b2cRows);
  return pack;
}

function updateFretesFileStatus(meta) {
  const el = $('fileStatus');
  if (!el) return;
  const apply = (m) => {
    const cte = m?.[fteCteSlot()];
    const sap = m?.[fteSapSlot()];
    const parts = [];
    if (typeof excelStatusPart === 'function') {
      const c = excelStatusPart(cte, 'CT-e:');
      const s = excelStatusPart(sap, 'SAP:');
      if (c) parts.push(c);
      if (s) parts.push(s);
    } else {
      if (cte?.file_name) parts.push('CT-e: ' + cte.file_name);
      if (sap?.file_name) parts.push('SAP: ' + sap.file_name);
    }
    const qzPart = quinzenalExcelStatusPart(m?.[fteQuinzenalSlot()]);
    if (qzPart) parts.push(qzPart);
    if (parts.length) {
      el.textContent = parts.join(' · ');
      el.style.display = 'block';
    } else {
      el.textContent = '';
      el.style.display = 'none';
    }
  };
  if (meta) {
    apply(meta);
    return;
  }
  const fetchMeta = typeof fetchExcelFileMeta === 'function' ? fetchExcelFileMeta : fetchExcelFiles;
  if (typeof fetchMeta !== 'function') return;
  fetchMeta([fteCteSlot(), fteSapSlot(), fteQuinzenalSlot()]).then(apply).catch(() => {
    el.style.display = 'none';
  });
}

/** Cached ZFACT month keys derived from sapNfMap dtEmissao — invalidate on map change. */
let _zfactMonthsCache = null;

function invalidateZfactMonthsCache() {
  _zfactMonthsCache = null;
}

function mesKeyFromDtEmissao(dt) {
  if (!dt) return null;
  let d = dt;
  if (!(d instanceof Date)) {
    d = (typeof parseSapBrDate === 'function' ? parseSapBrDate(dt) : null)
      || (typeof dt === 'string' || typeof dt === 'number' ? new Date(dt) : null);
  }
  if (!d || isNaN(d.getTime())) return null;
  const k = monthKey(d);
  return k && k !== 'sem-data' ? k : null;
}

/** Short + full month tokens for ZFACT filenames ("jan a Agosto", "nov a dez"). */
const ZFACT_FILE_MES_TOKENS = {
  janeiro: 1, jan: 1, fevereiro: 2, fev: 2, marco: 3, março: 3, mar: 3,
  abril: 4, abr: 4, maio: 5, mai: 5, junho: 6, jun: 6, julho: 7, jul: 7,
  agosto: 8, ago: 8, setembro: 9, set: 9, outubro: 10, out: 10,
  novembro: 11, nov: 11, dezembro: 12, dez: 12
};

/**
 * Derive YYYY-MM keys from ZFACT filename ranges when map dates are incomplete.
 * Examples: "vendas zfact jan a Agosto.XLSX", "vendas zfact nov a dez.XLSX".
 */
function parseZfactFileNameMonths(name) {
  const set = new Set();
  const raw = String(name || '');
  if (!raw) return set;
  const norm = raw.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const base = norm.replace(/\.xlsx?$/i, '');
  let ano = null;
  const ym = base.match(/(20\d{2})/);
  if (ym) ano = parseInt(ym[1], 10);
  const found = [];
  // Longer tokens first so "janeiro" wins over "jan"
  const tokens = Object.keys(ZFACT_FILE_MES_TOKENS).sort((a, b) => b.length - a.length);
  const usedSpans = [];
  tokens.forEach(tok => {
    const re = new RegExp(`(?:^|[^a-z])(${tok})(?:[^a-z]|$)`, 'i');
    const m = base.match(re);
    if (!m) return;
    const idx = m.index + m[0].indexOf(m[1]);
    if (usedSpans.some(([a, b]) => idx < b && idx + tok.length > a)) return;
    usedSpans.push([idx, idx + tok.length]);
    found.push({ mes: ZFACT_FILE_MES_TOKENS[tok], idx });
  });
  found.sort((a, b) => a.idx - b.idx);
  const months = [...new Set(found.map(f => f.mes))].sort((a, b) => a - b);
  if (!months.length) return set;
  const isRange = /\ba\b/.test(base) && months.length >= 1;
  if (!ano) {
    // nov/dez alone → ops start year 2025; jan–ago style → 2026
    ano = (months[0] >= 11 && months[months.length - 1] >= 11) ? 2025 : 2026;
  }
  if (isRange && months.length >= 1) {
    const lo = months[0];
    const hi = months[months.length - 1];
    if (hi >= lo) {
      for (let m = lo; m <= hi; m++) set.add(`${ano}-${String(m).padStart(2, '0')}`);
    } else {
      // e.g. nov a fev crossing year — uncommon for these filenames
      for (let m = lo; m <= 12; m++) set.add(`${ano}-${String(m).padStart(2, '0')}`);
      for (let m = 1; m <= hi; m++) set.add(`${ano + 1}-${String(m).padStart(2, '0')}`);
    }
  } else {
    months.forEach(m => set.add(`${ano}-${String(m).padStart(2, '0')}`));
  }
  return set;
}

function collectZfactFilenameMonthHints() {
  const set = new Set();
  const names = new Set();
  if (fteSapFileName) names.add(fteSapFileName);
  (sapMapSourceFiles || []).forEach(f => { if (f?.fileName) names.add(f.fileName); });
  names.forEach(n => parseZfactFileNameMonths(n).forEach(k => set.add(k)));
  return set;
}

function getZfactMonthsSet() {
  if (_zfactMonthsCache) return _zfactMonthsCache;
  const set = new Set();
  const map = sapNfMap || {};
  const keys = Object.keys(map);
  for (let i = 0; i < keys.length; i++) {
    const k = mesKeyFromDtEmissao(map[keys[i]]?.dtEmissao);
    if (k) set.add(k);
  }
  // Fallback: currentNFs with SAP / analysis dates when map empty
  if (!set.size && currentNFs?.length) {
    currentNFs.forEach(nf => {
      const k = mesKeyFromDtEmissao(nf.dtNF || nf.dtEmissao)
        || (nf.mesRef && nf.mesRef !== 'sem-data' ? nf.mesRef : null);
      if (k) set.add(k);
    });
  }
  // Filename range hint (e.g. "jan a Agosto") fills gaps when map dates were not parsed
  collectZfactFilenameMonthHints().forEach(k => set.add(k));
  _zfactMonthsCache = set;
  return set;
}

function enumerateMesKeysInclusive(minKey, maxKey) {
  const a = String(minKey || '').match(/^(\d{4})-(\d{2})$/);
  const b = String(maxKey || '').match(/^(\d{4})-(\d{2})$/);
  if (!a || !b) return [];
  let y = +a[1], m = +a[2];
  const y2 = +b[1], m2 = +b[2];
  const out = [];
  // Safety: never synthesize more than 36 months
  for (let i = 0; i < 36; i++) {
    out.push(`${y}-${String(m).padStart(2, '0')}`);
    if (y === y2 && m === m2) break;
    m++;
    if (m > 12) { m = 1; y++; }
  }
  return out;
}

function resolveFretesOpsStartMes(allKeys) {
  // Prefer earliest Nov in loaded data at/after default ops start; else constant.
  const novs = [...(allKeys || [])].filter(k => /^\d{4}-11$/.test(k)).sort();
  const fromData = novs.find(k => k >= FRETES_OPS_START_MES);
  return fromData || FRETES_OPS_START_MES;
}

function covQzMark(qs, opts) {
  if (!qs || !qs.size) return { text: '—', cls: 'cov-miss', title: 'Sem quinzenal' };
  const has1 = qs.has(1);
  const has2 = qs.has(2);
  const hasUnk = [...qs].some(q => q !== 1 && q !== 2);
  if (has1 && has2) return { text: '✓', cls: 'cov-ok', title: '1ªQ + 2ªQ' };
  if (has1 && !has2) return { text: '1ªQ', cls: 'cov-partial', title: 'Só 1ª quinzena' };
  if (has2 && !has1) {
    if (opts?.allowSecondQuinzenaOnly) {
      return { text: '2ªQ', cls: 'cov-ok', title: 'Só 2ªQ — esperado (ops desde 10 Nov)' };
    }
    return { text: '2ªQ', cls: 'cov-partial', title: 'Só 2ª quinzena' };
  }
  if (hasUnk) return { text: '✓', cls: 'cov-ok', title: 'Ficheiro(s) no mês' };
  return { text: '—', cls: 'cov-miss', title: 'Sem quinzenal' };
}

function quinzenaFromFileMeta(f) {
  if (!f) return 0;
  const q = Number(f.quinzena);
  if (q === 1 || q === 2) return q;
  const qk = String(f.quinzenaKey || '');
  const m = qk.match(/-Q([12])$/i);
  if (m) return Number(m[1]);
  const label = String(f.quinzenaLabel || '');
  if (/1\s*ª?\s*Q/i.test(label)) return 1;
  if (/2\s*ª?\s*Q/i.test(label)) return 2;
  return 0;
}

function buildLoadCoverageRows() {
  const b2b = new Map(); // mesKey -> Set(quinzena)
  const b2c = new Map();
  const addQz = (mesKey, canal, quinzena) => {
    if (!mesKey || mesKey === 'sem-mes') return;
    const map = canal === 'B2B' ? b2b : (canal === 'B2C' ? b2c : null);
    if (!map) return;
    if (!map.has(mesKey)) map.set(mesKey, new Set());
    const q = Number(quinzena);
    map.get(mesKey).add(Number.isFinite(q) && q > 0 ? q : 0);
  };
  (quinzenalPack?.files || []).forEach(f => {
    addQz(f.mesKey, f.canal, quinzenaFromFileMeta(f));
  });
  (fteQzPendingFiles || []).forEach(file => {
    const meta = parseQuinzenalFileName(file.name || file.fileName || '');
    addQz(meta.mesKey, meta.canal, meta.quinzena);
  });
  const zfact = getZfactMonthsSet();
  const allKeys = new Set([...b2b.keys(), ...b2c.keys(), ...zfact]);
  if (!allKeys.size) return [];
  const opsStart = resolveFretesOpsStartMes(allKeys);
  const sorted = [...allKeys]
    .filter(k => /^\d{4}-\d{2}$/.test(k) && k >= opsStart)
    .sort();
  if (!sorted.length) return [];
  const span = enumerateMesKeysInclusive(sorted[0], sorted[sorted.length - 1]);
  const keys = (span.length && span.length <= 36 ? span : sorted).filter(k => k >= opsStart);
  return keys.map(mesKey => {
    // Ops-start Nov only: Unilog 2ªQ-only is expected (not later Novembers)
    const allow2qOnly = mesKey === FRETES_OPS_START_MES;
    const b2bMark = covQzMark(b2b.get(mesKey), { allowSecondQuinzenaOnly: allow2qOnly });
    const b2cMark = covQzMark(b2c.get(mesKey), { allowSecondQuinzenaOnly: allow2qOnly });
    const hasZ = zfact.has(mesKey);
    const zMark = hasZ
      ? { text: '✓', cls: 'cov-ok', title: 'NFs ZFACT neste mês' }
      : { text: '—', cls: 'cov-miss', title: 'Sem ZFACT neste mês' };
    const hasQz = b2bMark.text !== '—' || b2cMark.text !== '—';
    const partialQz = b2bMark.cls === 'cov-partial' || b2cMark.cls === 'cov-partial';
    const canalGap = (b2bMark.text !== '—') !== (b2cMark.text !== '—');
    const crossGap = hasQz !== hasZ;
    return {
      mesKey,
      mesLabel: fmtMesLabel(mesKey),
      year: mesKey.slice(0, 4),
      b2b: b2bMark,
      b2c: b2cMark,
      zfact: zMark,
      gap: canalGap || crossGap || partialQz
    };
  });
}

function renderLoadCoverageSummary() {
  const el = $('coverageSummary');
  if (!el) return;
  const rows = buildLoadCoverageRows();
  if (!rows.length) {
    el.style.display = 'none';
    el.innerHTML = '';
    return;
  }
  const gaps = rows.filter(r => r.gap).length;
  let lastYear = '';
  const body = rows.map(r => {
    let yearRow = '';
    if (r.year !== lastYear) {
      lastYear = r.year;
      yearRow = `<tr class="cov-year"><td colspan="4">${r.year}</td></tr>`;
    }
    const cell = (m) => `<td class="cov-cell ${m.cls}" title="${m.title || ''}">${m.text}</td>`;
    return `${yearRow}<tr class="${r.gap ? 'cov-gap' : ''}">
      <td class="cov-mes">${r.mesLabel}</td>
      ${cell(r.b2b)}${cell(r.b2c)}${cell(r.zfact)}
    </tr>`;
  }).join('');
  el.style.display = 'block';
  el.innerHTML = `
    <div class="fte-cov-head">
      <div class="fte-cov-title">Cobertura por mês</div>
      <div class="fte-cov-hint">${rows.length} mês(es)${gaps ? ` · ${gaps} com lacuna` : ' · completo'}</div>
    </div>
    <div class="fte-cov-tbl-wrap">
      <table class="fte-cov-tbl">
        <thead><tr>
          <th>Mês</th><th style="text-align:center">B2B</th><th style="text-align:center">B2C</th><th style="text-align:center">ZFACT</th>
        </tr></thead>
        <tbody>${body}</tbody>
      </table>
    </div>
    <div class="fte-cov-foot">Cobertura desde ${fmtMesLabel(FRETES_OPS_START_MES)} (início da operação ~10 Nov). Linhas a amarelo: falta B2B ou B2C, só uma quinzena (excepto Nov inicial: só 2ªQ é esperado), ou Unilog sem ZFACT (e vice-versa).</div>`;
}

async function persistFretesFile(slot, fileName, arrayBuffer) {
  if (typeof upsertExcelBinary !== 'function') {
    fteToastError('Persistência Excel indisponível — recarrega a página.');
    return false;
  }
  const byteLen = arrayBuffer?.byteLength || 0;
  console.log('[fretes] persist', fteCompany(), slot, fileName, 'bytes', byteLen);
  try {
    await upsertExcelBinary(slot, fileName, arrayBuffer);
    updateFretesFileStatus();
    console.log('[fretes] persist ok', fteCompany(), slot, fileName);
    return true;
  } catch (err) {
    console.error('[fretes] persist', fteCompany(), slot, err);
    if (typeof isExcelFilesTableMissing === 'function' && isExcelFilesTableMissing(err)) {
      fteToastError('Tabela logistica_excel_files em falta no Supabase.');
    } else {
      fteToastError('Erro ao guardar Excel: ' + (err.message || err));
    }
    return false;
  }
}

const NF_COLUMNS = [
  { key: 'nf', label: 'Nota Fiscal', type: 'string' },
  { key: 'dtNF', label: 'Data NF', type: 'date', title: 'Data de emissão (SAP, se carregado)' },
  { key: 'cliente', label: 'Cliente', type: 'string', title: 'Nome do cliente (SAP, se carregado)' },
  { key: 'transportador', label: 'Transportador', type: 'string' },
  { key: 'modalidade', label: 'Modal.', type: 'string' },
  { key: 'valorNF', label: 'Valor NF (Unilog)', type: 'number', right: true, title: 'Valor da NF no Excel Unilog' },
  { key: 'valorSAP', label: 'Valor SAP', type: 'number', right: true, title: 'Valor da NF no Excel SAP' },
  { key: 'valorDiff', label: 'Δ SAP', type: 'number', right: true, title: 'Diferença valor SAP − Unilog' },
  { key: 'nCte', label: 'Qtd CT-e', type: 'number', right: true, title: 'Quantidade de CT-e para esta NF/cobrança (col. E do Excel Unilog)' },
  { key: 'pago', label: 'Pago', type: 'number', right: true },
  { key: 'esperado', label: 'Esperado (6%)', type: 'number', right: true },
  { key: 'diff', label: 'Δ frete', type: 'number', right: true, title: 'Pago − esperado 6% (não é Δ SAP vs Unilog)' },
  { key: 'pct', label: '% pago', type: 'number', right: true },
  { key: 'status', label: 'Estado', type: 'status' }
];

const SUB_COLUMNS = [
  { key: 'nf', label: 'NF', type: 'string' },
  { key: 'dtNF', label: 'Data NF', type: 'date' },
  { key: 'cliente', label: 'Cliente', type: 'string' },
  { key: 'transportador', label: 'Transportador', type: 'string' },
  { key: 'nCte', label: 'Qtd CT-e', type: 'number' },
  { key: 'valorNF', label: 'Valor Unilog', type: 'number', right: true },
  { key: 'valorSAP', label: 'Valor SAP', type: 'number', right: true },
  { key: 'valorDiff', label: 'Δ SAP', type: 'number', right: true },
  { key: 'pago', label: 'Pago', type: 'number', right: true },
  { key: 'esperado', label: 'Esperado', type: 'number', right: true },
  { key: 'diff', label: 'Δ frete', type: 'number', right: true, title: 'Pago − esperado 6% (não é Δ SAP vs Unilog)' },
  { key: 'pct', label: '% pago', type: 'number', right: true }
];

const SAP_ANOMALY_TYPE = 'NF não encontrada no SAP';

const ANOMALY_COLUMNS = [
  { key: 'nf', label: 'Nota Fiscal', type: 'string' },
  { key: 'cliente', label: 'Cliente (Unilog)', type: 'string' },
  { key: 'dtNF', label: 'Data', type: 'date', title: 'Data NF no Unilog (SAP indisponível)' },
  { key: 'valorNF', label: 'Valor NF Unilog', type: 'number', right: true },
  { key: 'transportador', label: 'Transportador', type: 'string' },
  { key: 'nCte', label: 'Qtd CT-e', type: 'number', right: true },
  { key: 'pago', label: 'Valor pago', type: 'number', right: true },
  { key: 'status', label: 'Estado frete', type: 'status' },
  { key: 'sapAnomalyType', label: 'Tipo anomalia', type: 'string' },
  { key: 'sapAnomalyReason', label: 'Explicação', type: 'string' }
];

const MONTH_COLUMNS = [
  { key: 'mesRef', label: 'Mês', type: 'mesRef' },
  { key: 'totalNF', label: 'NFs', type: 'number', right: true },
  { key: 'totalValorNF', label: 'Faturação', type: 'number', right: true },
  { key: 'totalPago', label: 'Pago', type: 'number', right: true },
  { key: 'totalEsperado', label: 'Esperado (6%)', type: 'number', right: true },
  { key: 'excesso', label: 'Excesso', type: 'number', right: true },
  { key: 'pctPagoSobreNF', label: '% pago/NF', type: 'number', right: true },
  { key: 'qtd_fix165', label: 'Sobr. R$165', type: 'number', right: true },
  { key: 'qtd_min', label: 'Tarifa mín.', type: 'number', right: true },
  { key: 'qtd_flag', label: 'Investigar', type: 'number', right: true }
];

const MESES_PT = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

const FIELD_ALIASES = {
  nf: ['nota fiscal', 'nf', 'n nota fiscal', 'num nota fiscal', 'num. nota fiscal', 'nota fiscal numero', 'notafiscal'],
  valorNF: ['valor nf', 'valor da nf', 'valor nota fiscal', 'vl nf', 'valor nf r$'],
  transportador: ['transportador', 'transportadora', 'nome transportador', 'transp'],
  modalidade: ['modalidade', 'modal', 'mod'],
  dtNF: [
    'dt nf', 'data nf', 'datanf', 'data da nf', 'data nota fiscal', 'dt nota fiscal',
    'data emissao', 'data de emissao', 'data emissão', 'data de emissão',
    'dt emissao', 'dt emissão', 'emissao nf', 'emissão nf',
    'data emissao nf', 'data emissão nf', 'datas de emissao nf', 'datas de emissão nf'
  ],
  numCte: ['num. cte', 'num cte', 'numero cte', 'n cte-e', 'num cte-e', 'numero do cte', 'numero do cte-e', 'chave cte', 'chave cte-e', 'numero cte-e', 'n cte e', 'num cte e', 'no cte-e', 'n. cte-e'],
  qtdCte: ['qtd cte', 'qtd. cte', 'qtd ct-e', 'qtd. ct-e', 'quantidade cte', 'quantidade ct-e', 'quantidade de cte', 'quantidade de ct-e', 'n ctes', 'numero ctes', 'total cte', 'total ct-e', 'total de cte', 'qtd cobranca', 'qtd cobrança', 'quantidade cobranca', 'quantidade cobrança', 'nº cte', 'no cte', 'n. cte', 'n cte'],
  dtCte: ['dt cte', 'dt ct-e', 'dt. cte', 'dt. ct-e', 'data cte', 'data ct-e', 'data cte-e', 'dt cte-e', 'data do cte', 'data do ct-e', 'data cte e'],
  pago: ['total fatura rev.', 'total fatura rev', 'total fatura', 'valor fatura', 'valor pago', 'total pago', 'vl pago', 'frete pago', 'valor cte', 'valor do frete'],
  devolucao: ['devolucao', 'devolução', 'e devolucao', 'retorno'],
  tipoOp: ['tipo operacao', 'tipo operação', 'tipo op', 'operacao'],
  peso: ['peso', 'peso kg', 'peso (kg)']
};

/** Line-level billing (sum per NF). Doc-level totals repeat on every row — never sum those. */
const SAP_LINE_VALOR_ALIASES = [
  'valor bruto', 'vl bruto', 'val bruto', 'vlr bruto', 'vlr.bruto', 'val.bruto',
  'valor linha', 'vl linha', 'valor item',
  'val faturamento', 'valor faturamento', 'valor mercadoria', 'netwr'
];
/**
 * ZFACT NFe export col AP — "Val.total incl.imp." — matches Unilog Valor NF.
 * Prefer over Valor líquido (AI) / Valor Bruto when present (e.g. NF 100934: AP 6673.50 vs AI 6951.56).
 */
const SAP_NF_TOTAL_INCL_IMP_ALIASES = [
  'val.total incl.imp.', 'val.total incl.imp', 'val total incl imp', 'val.total incl imp',
  'valor total incl.imp', 'valor total incl imp', 'valor total incluindo', 'valor total com imposto',
  'total incl.imp', 'total incl.imp.', 'total incluindo imposto', 'total com imposto',
  'vnf', 'valor total da nf-e', 'valor total nfe'
];
const SAP_DOC_VALOR_ALIASES = [
  ...SAP_NF_TOTAL_INCL_IMP_ALIASES,
  'doc total', 'doc. total', 'doctotal', 'valor documento', 'valor total', 'total nf', 'total documento',
  'total do documento', 'total geral', 'valor nf', 'valor da nf', 'valor nota fiscal', 'vl nf',
  'vl documento', 'vlr total', 'vlr documento', 'montante documento', 'montante doc',
  'amount', 'valor moeda', 'valor em brl', 'total brl', 'total moeda',
  'valor da fatura', 'valor fatura', 'invoice total', 'billing value'
];
/** Líquido / net — never treat as NF face doc total (Unilog Valor NF is gross/doc face). */
const SAP_NET_VALOR_ALIASES = [
  'valor liquido', 'vl liquido', 'valor liquido item', 'net value', 'netwr', 'valor mercadoria'
];
/** Last resort only — never preferred over Valor Bruto / doc total. */
const SAP_VALOR_FALLBACK_ALIASES = ['vl total', 'kwert', 'montante'];
/**
 * Never use as NF billing base — ZFACT Valor Líquido (col AI / classic J) ≠ Unilog Valor NF;
 * also impostos, taxa 5,5%, frete, unit prices.
 */
const SAP_VALOR_EXCLUDE_ALIASES = [
  'valor liquido', 'vl liquido', 'vlr liquido', 'liquido', 'valor liquido item',
  'imposto', 'impostos', 'icms', 'pis', 'cofins', 'ipi', 'iss',
  'taxa', 'frete', 'fee', 'armazenagem', 'comissao', 'comissão',
  '5,5', '5.5', 'percentual', '%', 'valor unitario', 'preco unitario', 'preço unitário',
  'preco liq', 'preço líq', 'preco liquido'
];
/** ZFACT export: D=Nota Fiscal (3), N=valor bruto (13) — only a fallback when headers missing. */
const SAP_ZFACT_BRUTO_COL = 13;
/** Typical ZFACT Valor Líquido column (0-based) — never treat as NF billing. */
const SAP_ZFACT_LIQUIDO_COL = 9;
/** ZFACT NFe export: AP = Val.total incl.imp. (Excel col 42 → 0-based 41) — matches Unilog. */
const SAP_ZFACT_TOTAL_INCL_IMP_COL = 41;
/** ISO currency suffixes seen in SAP GUI Excel exports (e.g. "4 992,26 BRL"). */
const SAP_CURRENCY_SUFFIX_RE = /(?:BRL|USD|EUR|GBP|JPY|CHF|CAD|AUD|CNY|MXN|ARS|CLP|COP|PEN)\s*$/i;

const SAP_NF_PREFERRED_ALIASES = [
  'nota fiscal', 'n nota fiscal', 'num nota fiscal', 'num. nota fiscal', 'notafiscal',
  'nº nf', 'numero nf', 'nº da nf', 'numero da nf', 'nf-e', 'nfe', 'nº nf-e', 'numero nf-e',
  'referencia', 'referência', 'reference', 'xblnr', 'assignment'
];
const SAP_NF_WEAK_ALIASES = [
  'docnum', 'nº doc', 'num doc', 'n documento', 'no documento', 'numero documento', 'número documento',
  'billing document', 'documento de faturamento'
];

const SAP_ALIASES = {
  // Prefer Unilog-facing NF / referência over SAP DocNum (VBELN) — docnum last.
  nf: [...SAP_NF_PREFERRED_ALIASES, 'nf', ...SAP_NF_WEAK_ALIASES],
  dtEmissao: [
    'dt emissao', 'data emissao', 'dt emissão', 'data emissão', 'data doc', 'dt nf', 'data nf', 'data nota fiscal',
    // Wide ZFACT NFe export (col AQ/AR): "Data de criação" — values often sit one col to the right
    'data de criacao', 'data de criação', 'data criacao', 'data criação', 'dt criacao', 'dt criação', 'criado em'
  ],
  cliente: [
    'cliente', 'nome cliente', 'razao social', 'razão social', 'destinatario', 'destinatário', 'cardname', 'nome do cliente',
    'nome', 'id parceiro', 'parceiro'
  ],
  material: ['material', 'cod material', 'cod. material', 'codigo material', 'item code', 'cod item', 'cod. item', 'nº item', 'num item', 'item no', 'produto'],
  valorNF: [...SAP_LINE_VALOR_ALIASES, ...SAP_DOC_VALOR_ALIASES, ...SAP_VALOR_FALLBACK_ALIASES]
};

/** SAP NF export layout: col B=cliente, C=data emissão, D=NF (0-based indices). */
const SAP_COL_IDX_DEFAULT = { cliente: 1, dtEmissao: 2, nf: 3, valorNF: 4 };
let activeSapColIdx = { ...SAP_COL_IDX_DEFAULT };

function fteToast(msg) {
  if (typeof toast === 'function') toast(msg, 'success');
}

function fteToastError(msg) {
  if (typeof toast === 'function') toast(msg, 'error');
}

/** SheetJS CE (0.18.x) cannot decrypt ECMA-376 password-protected xlsx. */
function isEncryptedXlsxError(err) {
  const m = String(err?.message || err || '');
  return /encryptioninfo|encrypted file|password|agile encryption|office crypto|file is password-protected/i.test(m);
}

function formatXlsxReadError(err, context) {
  if (isEncryptedXlsxError(err)) {
    return 'Ficheiro SAP protegido por password — não é possível ler ficheiros Excel encriptados. '
      + 'Exporta do SAP sem proteção, ou abre no Excel e guarda como .xlsx sem password '
      + '(Ficheiro → Guardar como). Alternativa: exportar como CSV.';
  }
  const prefix = context === 'sap' ? 'Erro a ler SAP' : 'Erro a ler o ficheiro';
  return prefix + ': ' + String(err?.message || err || 'ficheiro inválido');
}

function setSapLoadStatus(msg, ok) {
  const el = $('sapLoadStatus') || $('sapFn');
  if (!el) return;
  el.textContent = msg || '';
  if (el.id === 'fte-sapFn' && !msg) return;
  el.style.color = ok === false ? 'var(--red)' : (ok === true ? 'var(--green)' : 'var(--muted)');
}

function setLoadbar(msg, isError) {
  const el = $('loadbar');
  if (!el) return;
  el.style.display = 'block';
  el.textContent = msg;
  el.style.color = isError ? 'var(--red)' : 'var(--muted)';
}

function readWorkbookFromArrayBuffer(data, opts) {
  return XLSX.read(data, { type: 'array', cellDates: true, ...opts });
}

/** SAP NF: raw serials/strings — avoid cellDates (SheetJS UTC Date → 1969/1970 display bugs). */
function readSapWorkbookFromArrayBuffer(data) {
  return XLSX.read(data, { type: 'array', cellDates: false });
}

const SAP_DATE_MIN_YEAR = 1990;
const SAP_DATE_MAX_YEAR = 2100;

function isPlausibleSapDate(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return false;
  const y = d.getFullYear();
  return y >= SAP_DATE_MIN_YEAR && y <= SAP_DATE_MAX_YEAR;
}

function sapDateFromParts(yr, mo, dd) {
  if (yr < SAP_DATE_MIN_YEAR || yr > SAP_DATE_MAX_YEAR || mo < 1 || mo > 12 || dd < 1 || dd > 31) return null;
  const d = new Date(yr, mo - 1, dd);
  if (isNaN(d.getTime()) || d.getFullYear() !== yr || d.getMonth() !== mo - 1 || d.getDate() !== dd) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function sapDateFromExcelSerial(serial) {
  if (typeof XLSX === 'undefined' || !XLSX.SSF) return null;
  if (serial === null || serial === undefined || serial === '' || serial <= 0 || serial >= 1000000) return null;
  const dc = XLSX.SSF.parse_date_code(serial);
  if (!dc || dc.y < SAP_DATE_MIN_YEAR || dc.y > SAP_DATE_MAX_YEAR) return null;
  return sapDateFromParts(dc.y, dc.m, dc.d);
}

function updateSaveStatus(msg, ok) {
  const el = $('saveStatus');
  if (!el) return;
  el.textContent = msg || '';
  el.style.color = ok === false ? 'var(--redflag)' : 'var(--gray)';
}

function showResultsView(opts = {}) {
  const lb = $('loadbar');
  if (lb) lb.style.display = 'none';
  renderCteSubPanels();
  const shouldSwitch = opts.switchTab !== false && !_fteSkipAutosave && currentNFs.length;
  if (shouldSwitch) switchFteTab('analise-cte');
}

function fmtMoney(v) { return 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function fmtPct(v) { return (v * 100).toFixed(1) + '%'; }

function normCol(s) {
  return String(s || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

/**
 * Canonical NF key for Unilog ↔ SAP matching.
 * Strips optional série suffix (-1, -2, -A…) and leading zeros from the numeric part.
 * Same normalization on both sides.
 *
 * Unit samples (all → "97723"):
 *   "97723", "00097723", "97723-1", "00097723-2", "0000097723-1", 97723, 97723.0
 */
function normNFKey(nf) {
  if (nf === null || nf === undefined || nf === '') return '';
  if (typeof nf === 'number' && !isNaN(nf)) {
    nf = Math.abs(nf - Math.round(nf)) < 1e-6 ? String(Math.round(nf)) : String(nf);
  }
  let s = String(nf).trim();
  const floatMatch = s.match(/^(\d+)\.0+$/);
  if (floatMatch) s = floatMatch[1];
  const dashIdx = s.indexOf('-');
  const base = dashIdx >= 0 ? s.slice(0, dashIdx) : s;
  const digits = base.replace(/\D/g, '');
  if (!digits) return '';
  return digits.replace(/^0+/, '') || '0';
}

/** Console self-check for normNFKey — run once at module load. */
function _debugNormNFKeySamples() {
  const samples = [
    ['97723', '97723'],
    ['00097723', '97723'],
    ['97723-1', '97723'],
    ['97723-2', '97723'],
    ['00097723-2', '97723'],
    ['0000097723-1', '97723'],
    ['97723-A', '97723'],
    [97723, '97723'],
    [97723.0, '97723'],
    ['97723.0', '97723']
  ];
  const mismatches = samples.filter(([raw, exp]) => normNFKey(raw) !== exp);
  if (mismatches.length) {
    console.warn('[normNFKey] sample mismatches:', mismatches.map(([raw, exp]) => ({
      raw, got: normNFKey(raw), expected: exp
    })));
  } else {
    console.debug('[normNFKey] samples OK:', samples.map(([raw]) => ({ raw, key: normNFKey(raw) })));
  }
}
_debugNormNFKeySamples();

function lookupSapEntry(nf) {
  const key = normNFKey(nf);
  if (!key) return null;
  return sapNfMap[key] || null;
}

function countSapMatches(nfList) {
  if (!nfList?.length || !Object.keys(sapNfMap).length) return 0;
  return nfList.filter(nf => lookupSapEntry(nf.nf)).length;
}

function logSapKeyDebug(nfList) {
  const uniSample = (nfList || []).slice(0, 5).map(n => ({
    raw: n.nf,
    key: normNFKey(n.nf)
  }));
  const sapSample = [...new Set(Object.keys(sapNfMap))].slice(0, 5);
  console.debug('[SAP match] sample Unilog NF keys:', uniSample, 'sample SAP map keys:', sapSample);
}

function findField(row, field) {
  const aliases = FIELD_ALIASES[field];
  const entries = Object.entries(row);
  for (const alias of aliases) {
    const hit = entries.find(([k]) => normCol(k) === alias);
    if (hit && hit[1] !== null && hit[1] !== undefined && hit[1] !== '') return hit[1];
  }
  for (const alias of aliases) {
    const hit = entries.find(([k]) => {
      const nk = normCol(k);
      return nk.includes(alias) || alias.includes(nk);
    });
    if (hit && hit[1] !== null && hit[1] !== undefined && hit[1] !== '') return hit[1];
  }
  return null;
}

function sapHeaderMatchesField(hdr, aliases) {
  const h = normCol(hdr);
  if (!h) return false;
  return aliases.some(alias => h === alias || h.includes(alias) || alias.includes(h));
}

function sapHeaderIsExcluded(hdr) {
  const h = normCol(hdr);
  if (!h) return false;
  // AP / Val.total incl.imp. must never be excluded just because the label contains "imposto(s)".
  if (h.includes('incl.imp') || h.includes('incl imp') || h.includes('incluindo imposto')
      || h.includes('incluindo imp') || h.includes('total incl') || h.includes('com imposto')) {
    return false;
  }
  // Only header-contains-alias (not alias-contains-header) — avoids excluding bare "Valor"
  // because exclude list has "valor liquido".
  return SAP_VALOR_EXCLUDE_ALIASES.some(alias => h === alias || h.includes(alias));
}

function findSapFieldKeyByAliases(row, aliases, opts = {}) {
  const entries = Object.entries(row);
  const skipExcluded = opts.skipExcluded !== false;
  const keyOk = (k) => !skipExcluded || !sapHeaderIsExcluded(k);
  for (const alias of aliases) {
    const hit = entries.find(([k]) => keyOk(k) && normCol(k) === alias);
    if (hit && hit[1] !== null && hit[1] !== undefined && hit[1] !== '') return hit[0];
  }
  for (const alias of aliases) {
    const hit = entries.find(([k]) => {
      if (!keyOk(k)) return false;
      const nk = normCol(k);
      return nk.includes(alias) || alias.includes(nk);
    });
    if (hit && hit[1] !== null && hit[1] !== undefined && hit[1] !== '') return hit[0];
  }
  return null;
}

function findSapFieldByAliases(row, aliases) {
  const key = findSapFieldKeyByAliases(row, aliases);
  if (!key) return null;
  const v = row[key];
  return v !== null && v !== undefined && v !== '' ? v : null;
}

function findSapField(row, field) {
  return findSapFieldByAliases(row, SAP_ALIASES[field]);
}

/** Prefer exact / strong "Val.total incl.imp." headers over fuzzy alias.includes(short name). */
function findSapInclImpFieldKey(row) {
  const entries = Object.entries(row || {});
  for (const [k, v] of entries) {
    if (v === null || v === undefined || v === '') continue;
    if (isSapTotalInclImpHeader(k)) return k;
  }
  // Exact alias match first (avoid "Total" / short keys matching via alias.includes).
  for (const alias of SAP_NF_TOTAL_INCL_IMP_ALIASES) {
    const hit = entries.find(([k, v]) => {
      if (v === null || v === undefined || v === '') return false;
      return normCol(k) === alias;
    });
    if (hit) return hit[0];
  }
  for (const alias of SAP_NF_TOTAL_INCL_IMP_ALIASES) {
    if (alias.length < 10) continue; // skip short aliases like 'vnf' for fuzzy
    const hit = entries.find(([k, v]) => {
      if (v === null || v === undefined || v === '') return false;
      if (sapHeaderIsExcluded(k) && !isSapTotalInclImpHeader(k)) return false;
      const nk = normCol(k);
      if (nk.includes(alias)) return true;
      // Only allow alias⊇header when header itself looks like an incl.imp. / NF total label.
      if (alias.includes(nk) && (nk.includes('incl') || nk.includes('imp') || nk.includes('val.total')
          || nk.includes('valor total') || nk.length >= 14)) return true;
      return false;
    });
    if (hit) return hit[0];
  }
  return null;
}

function findSapValorFields(row) {
  // Prefer Unilog-matching total (AP / Val.total incl.imp.) over generic doc aliases.
  const preferredDocKey = findSapInclImpFieldKey(row);
  const docKey = preferredDocKey || findSapFieldKeyByAliases(row, SAP_DOC_VALOR_ALIASES);
  const doc = docKey ? row[docKey] : null;
  let lineKey = findSapFieldKeyByAliases(row, SAP_LINE_VALOR_ALIASES);
  if (lineKey && docKey && lineKey === docKey) lineKey = null;
  const line = lineKey ? row[lineKey] : null;
  const fallback = !line && !doc ? findSapFieldByAliases(row, SAP_VALOR_FALLBACK_ALIASES) : null;
  // Doc total (AP) is the NF billing base when present — primary must not be líquido/bruto alone.
  return {
    line,
    doc,
    fallback,
    primary: doc || line || fallback,
    fromInclImp: !!(preferredDocKey && doc != null && doc !== '')
  };
}

function isSapTotalInclImpHeader(hdr) {
  const h = normCol(hdr);
  if (!h || sapHeaderIsExcluded(hdr)) return false;
  if (h.includes('incl.imp') || h.includes('incl imp') || h.includes('incluindo imposto')) return true;
  return sapHeaderMatchesField(hdr, SAP_NF_TOTAL_INCL_IMP_ALIASES);
}

function resolveSapValorColFromHeader(headerRow) {
  let lineCol = null;
  let docCol = null;
  let preferredDocCol = null;
  const isBrutoHeader = (hdr) => {
    const h = normCol(hdr);
    if (!h || sapHeaderIsExcluded(hdr)) return false;
    if (h === 'valor bruto' || h.startsWith('valor bruto')) return true;
    if (h.includes('valor bruto') || h.includes('vl bruto') || h.includes('vlr bruto') || h.includes('val bruto')) return true;
    return SAP_LINE_VALOR_ALIASES.some(a => h === a || h.startsWith(a + ' '));
  };
  (headerRow || []).forEach((hdr, j) => {
    if (isBrutoHeader(hdr)) lineCol = j;
  });
  if (lineCol == null) {
    (headerRow || []).forEach((hdr, j) => {
      if (sapHeaderIsExcluded(hdr)) return;
      if (sapHeaderMatchesField(hdr, SAP_LINE_VALOR_ALIASES)) lineCol = j;
    });
  }
  (headerRow || []).forEach((hdr, j) => {
    if (sapHeaderIsExcluded(hdr)) return;
    // Col AP "Val.total incl.imp." — Unilog-matching NF total (beats generic "valor total").
    if (isSapTotalInclImpHeader(hdr)) preferredDocCol = j;
    else if (sapHeaderMatchesField(hdr, SAP_NET_VALOR_ALIASES)) return; // never líquido as doc
    else if (sapHeaderMatchesField(hdr, SAP_DOC_VALOR_ALIASES)) docCol = j;
  });
  if (preferredDocCol != null) docCol = preferredDocCol;
  if (lineCol != null && lineCol === docCol) lineCol = null;
  return { lineCol, docCol };
}

/** Prefer Nota Fiscal / referência over DocNum when both headers exist. */
function resolveSapNfColFromHeader(headerRow) {
  if (!headerRow || !headerRow.length) return null;
  let prefer = null;
  let weak = null;
  headerRow.forEach((hdr, j) => {
    const h = normCol(hdr);
    if (!h) return;
    if (prefer == null && SAP_NF_PREFERRED_ALIASES.some(a => h === a || h.includes(a) || a.includes(h))) {
      prefer = j;
    }
    if (weak == null && SAP_NF_WEAK_ALIASES.some(a => h === a || h.includes(a) || a.includes(h))) {
      weak = j;
    }
  });
  return prefer != null ? prefer : weak;
}

function normalizeRow(row) {
  return {
    nf: findField(row, 'nf'),
    valorNF: findField(row, 'valorNF'),
    transportador: findField(row, 'transportador'),
    modalidade: findField(row, 'modalidade'),
    dtNF: findField(row, 'dtNF'),
    numCte: findField(row, 'numCte'),
    qtdCte: findField(row, 'qtdCte'),
    dtCte: findField(row, 'dtCte'),
    pago: findField(row, 'pago'),
    devolucao: findField(row, 'devolucao'),
    tipoOp: findField(row, 'tipoOp'),
    peso: findField(row, 'peso')
  };
}

/** Normalize 2-digit years (Excel BR often uses 26 → 2026). */
function sapNormalizeYear(yr) {
  if (yr >= 0 && yr < 100) return yr >= 70 ? 1900 + yr : 2000 + yr;
  return yr;
}

/** Persist/restore-safe date string YYYY-MM-DD (local calendar day). */
function dateToYmd(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function parseSapBrDate(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    const d = new Date(v.getFullYear(), v.getMonth(), v.getDate());
    d.setHours(0, 0, 0, 0);
    return isPlausibleSapDate(d) ? d : null;
  }
  if (typeof v === 'number' && !isNaN(v)) {
    if (v <= 0) return null;
    if (v >= 1e12) {
      const d = new Date(v);
      if (isPlausibleSapDate(d)) {
        d.setHours(0, 0, 0, 0);
        return d;
      }
      return null;
    }
    return sapDateFromExcelSerial(v);
  }
  const s = String(v).trim();
  if (!s) return null;
  // ISO datetime from JSON.stringify(Date) — must run before parseFloat("2026-…") serial trap
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);
  if (iso) {
    const d = sapDateFromParts(parseInt(iso[1], 10), parseInt(iso[2], 10), parseInt(iso[3], 10));
    if (d) return d;
  }
  const dotParts = s.split('.');
  if (dotParts.length === 3 && dotParts.every(p => /^\d+$/.test(p))) {
    const dd = parseInt(dotParts[0], 10), mo = parseInt(dotParts[1], 10), yr = sapNormalizeYear(parseInt(dotParts[2], 10));
    const d = sapDateFromParts(yr, mo, dd);
    if (d) return d;
  }
  const slashParts = s.split('/');
  if (slashParts.length >= 3) {
    const yr = sapNormalizeYear(parseInt(String(slashParts[2]).trim(), 10));
    const mo = parseInt(slashParts[1], 10);
    const dd = parseInt(slashParts[0], 10);
    const d = sapDateFromParts(yr, mo, dd);
    if (d) return d;
  }
  const m = s.match(/^(\d{2})[.\-/](\d{2})[.\-/](\d{2,4})$/) || s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const yr = m[1].length === 4 ? parseInt(m[1], 10) : sapNormalizeYear(parseInt(m[3], 10));
    const mo = parseInt(m[2], 10);
    const dd = m[1].length === 4 ? parseInt(m[3], 10) : parseInt(m[1], 10);
    const d = sapDateFromParts(yr, mo, dd);
    if (d) return d;
  }
  // Excel serial as string — only if the whole token is numeric (avoid "2026-07-15…" → 2026)
  if (/^\d+([.,]\d+)?$/.test(s)) {
    const serial = parseFloat(s.replace(',', '.'));
    if (!isNaN(serial) && serial >= 1 && serial < 1000000) {
      const d = sapDateFromExcelSerial(serial);
      if (d) return d;
    }
  }
  return null;
}

function isSapHeaderAtCol(headerRow, colIdx, field) {
  const h = normCol(headerRow?.[colIdx]);
  if (!h) return false;
  return SAP_ALIASES[field].some(alias => h === alias || h.includes(alias) || alias.includes(h));
}

function looksLikeSapNfCell(v) {
  if (v === null || v === undefined || v === '') return false;
  const key = normNFKey(v);
  return key.length >= 4 && key.length <= 12;
}

/** Max plausible NF billing amount in ZFACT warehouse exports (R$). */
const SAP_VALOR_MAX = 5_000_000;
/** Near-equal Unilog vs SAP (cent rounding) — do not show / flag as Δ. */
const SAP_VALOR_NEAR_EPS = 0.05;

function hasSapValorFormatting(v) {
  const raw = String(v).trim();
  if (!raw) return false;
  if (/R\$/i.test(raw) || SAP_CURRENCY_SUFFIX_RE.test(raw)) return true;
  const s = raw.replace(/[\s\u00a0\u202f]/g, '');
  return /^R\$/i.test(s) || s.includes(',') || SAP_CURRENCY_SUFFIX_RE.test(s);
}

/** Reject SAP doc entry / NF identifiers mistaken for monetary valor (no R$/comma/currency). */
function looksLikeSapDocOrNfNumber(v) {
  if (v === null || v === undefined || v === '') return false;
  if (hasSapValorFormatting(v)) return false;
  const key = normNFKey(v);
  if (key.length >= 8) return true;
  const s = String(v).trim().replace(/[\s\u00a0\u202f]/g, '').replace(/^R\$/i, '');
  if (/^\d{7,}$/.test(s)) return true;
  if (typeof v === 'number' && Number.isInteger(v) && Math.abs(v) >= 1_000_000) return true;
  return false;
}

/** Parse NF monetary values from SAP/ZFACT exports (BR format, R$, "N BRL", thousands). */
function parseSapNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  let s = String(v).trim().replace(/[\s\u00a0\u202f]/g, '');
  if (!s || s === '-') return 0;
  s = s.replace(/^R\$/i, '');
  s = s.replace(SAP_CURRENCY_SUFFIX_RE, '');
  if (!s || s === '-') return 0;
  if (s.includes(',')) {
    const n = parseFloat(s.replace(/\./g, '').replace(',', '.'));
    return Number.isFinite(n) ? n : 0;
  }
  if (s.includes('.')) {
    const parts = s.split('.');
    if (parts.length > 2) {
      const n = parseFloat(s.replace(/\./g, ''));
      return Number.isFinite(n) ? n : 0;
    }
    const intPart = parts[0];
    const fracPart = parts[1] || '';
    if (!fracPart || /^0+$/.test(fracPart)) {
      const n = parseFloat(intPart);
      return Number.isFinite(n) ? n : 0;
    }
    if (intPart !== '0' && fracPart.length === 3 && intPart.length >= 1 && intPart.length <= 3) {
      const n = parseFloat(intPart + fracPart);
      if (Number.isFinite(n)) return n;
    }
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : 0;
  }
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}

function looksLikeSapValorCell(v, opts = {}) {
  if (v === null || v === undefined || v === '') return false;
  const n = parseSapNum(v);
  if (n === 0 || Math.abs(n) > SAP_VALOR_MAX) return false;
  if (looksLikeSapDocOrNfNumber(v)) return false;
  const s = String(v).trim().replace(/[\s\u00a0\u202f]/g, '');
  if (/^R\$/i.test(s) || s.includes(',') || SAP_CURRENCY_SUFFIX_RE.test(s) || SAP_CURRENCY_SUFFIX_RE.test(String(v))) {
    return true;
  }
  // Valor bruto column (header / ZFACT col N): accept whole reais stored as integers
  // (e.g. NF 99642 TIMOR R$786 — Excel serial 786 was dropped, sum became 1977 not 2763).
  if (opts.trustColumn) {
    if (typeof v === 'number' && Number.isInteger(v) && v >= 1e9) return false;
    return true;
  }
  // Heuristic column scan only: reject bare 1–3 digit integers (likely qty, not R$).
  if (/^\d{1,3}$/.test(s) && n < 1000) return false;
  return n >= 10 || s.includes('.') || (typeof v === 'number' && !Number.isInteger(v));
}

/** True when cell looks like SAP currency export "4 992,26 BRL". */
function looksLikeSapCurrencyAmount(v) {
  if (v === null || v === undefined || v === '') return false;
  const raw = String(v).trim();
  if (!SAP_CURRENCY_SUFFIX_RE.test(raw) && !/^R\$/i.test(raw)) return false;
  return looksLikeSapValorCell(v, { trustColumn: true });
}

/** Console self-check for SAP valor heuristics — run once at module load. */
function _debugSapValorSamples() {
  const accept = [
    [43200, true],
    ['43.200,00', true],
    ['43200', true],
    [1234.56, true],
    ['1.234,56', true],
    [999999, true],
    [786, true],
    ['786,00', true],
    [-6000, true],
    ['-6.000,00', true],
    ['-6000,00', true],
    ['4 992,26 BRL', true],
    ['4\u00a0992,26 BRL', true],
    ['4992,26BRL', true]
  ];
  const reject = [
    [9540067078, false],
    ['9540067078', false],
    ['9.540.067.078,00', false],
    [9540067312, false],
    [9540067546, false],
    [0, false],
    ['', false],
    ['0,00', false],
    ['0,00 BRL', false]
  ];
  const all = [...accept, ...reject];
  const mismatches = all.filter(([v, exp]) => looksLikeSapValorCell(v, { trustColumn: true }) !== exp);
  if (mismatches.length) {
    console.warn('[SAP valor] sample mismatches:', mismatches.map(([v, exp]) => ({
      v, expected: exp, got: looksLikeSapValorCell(v, { trustColumn: true }), parsed: parseSapNum(v)
    })));
  } else {
    console.debug('[SAP valor] samples OK:', accept.length + reject.length, 'cases');
  }
  const brlOk = Math.abs(parseSapNum('4 992,26 BRL') - 4992.26) < 0.001
    && Math.abs(parseSapNum('4\u00a0992,26 BRL') - 4992.26) < 0.001;
  if (!brlOk) console.warn('[SAP valor] BRL parse failed', parseSapNum('4 992,26 BRL'));
}
_debugSapValorSamples();

/** ZFACT: prefer col AP Val.total incl.imp.; never Valor Líquido (AI/J); bruto N only as line fallback. */
function pickSapValorFromLine(line, colIdx, headerRow) {
  const idx = colIdx || activeSapColIdx;
  if (!idx || !Array.isArray(line)) return { line: null, doc: null };

  const headerValor = resolveSapValorColFromHeader(headerRow);
  const out = { line: null, doc: null, fromInclImp: false };
  const nfKey = normNFKey(line[idx.nf]);
  const isSameNf = (v) => nfKey && normNFKey(v) === nfKey;
  const colExcluded = (j) => {
    if (headerRow && sapHeaderIsExcluded(headerRow[j])) return true;
    // Classic ZFACT layout without usable headers: col J = líquido (not NF total).
    if ((!headerRow || !headerRow.length) && j === SAP_ZFACT_LIQUIDO_COL) return true;
    return false;
  };

  const tryLine = (v) => {
    if (v != null && v !== '' && !isSameNf(v) && looksLikeSapValorCell(v, { trustColumn: true })) {
      out.line = v;
    }
  };
  const tryDoc = (v) => {
    if (v != null && v !== '' && !isSameNf(v) && looksLikeSapValorCell(v, { trustColumn: true })) {
      out.doc = v;
    }
  };

  // Prefer Unilog-matching doc total (AP / Val.total incl.imp.) before line bruto.
  if (headerValor.docCol != null && !colExcluded(headerValor.docCol)) {
    tryDoc(line[headerValor.docCol]);
    if (out.doc != null && headerRow && isSapTotalInclImpHeader(headerRow[headerValor.docCol])) {
      out.fromInclImp = true;
    }
  }
  if (!out.doc && line.length > SAP_ZFACT_TOTAL_INCL_IMP_COL && !colExcluded(SAP_ZFACT_TOTAL_INCL_IMP_COL)) {
    const apHdr = headerRow?.[SAP_ZFACT_TOTAL_INCL_IMP_COL];
    if (!headerRow?.length || isSapTotalInclImpHeader(apHdr)) {
      tryDoc(line[SAP_ZFACT_TOTAL_INCL_IMP_COL]);
      if (out.doc != null) out.fromInclImp = true;
    }
  }

  if (headerValor.lineCol != null && !colExcluded(headerValor.lineCol)) tryLine(line[headerValor.lineCol]);

  // Hardcoded col N only when it actually holds a non-zero amount (wide exports shift columns).
  if (!out.line && line.length > SAP_ZFACT_BRUTO_COL && !colExcluded(SAP_ZFACT_BRUTO_COL)) {
    tryLine(line[SAP_ZFACT_BRUTO_COL]);
  }

  if (!out.line && idx.valorNF != null && idx.valorNF !== headerValor.docCol && !colExcluded(idx.valorNF)) {
    tryLine(line[idx.valorNF]);
  }

  // SAP GUI exports often put the NF total as "4 992,26 BRL" without a matching header alias.
  // Prefer the largest currency cell (doc face), not a smaller net/partial BRL earlier in the row.
  if (!out.doc) {
    let best = null;
    let bestAbs = 0;
    for (let j = 0; j < line.length; j++) {
      if (colExcluded(j)) continue;
      const cell = line[j];
      if (!looksLikeSapCurrencyAmount(cell) || isSameNf(cell)) continue;
      const abs = Math.abs(parseSapNum(cell));
      if (abs > bestAbs) {
        bestAbs = abs;
        best = cell;
      }
    }
    if (best != null) tryDoc(best);
  }

  // When line bruto is missing/zero but a repeating currency doc total exists, keep doc only.
  if (out.line && out.doc && sapValoresClose(parseSapNum(out.line), parseSapNum(out.doc))) out.line = null;
  return out;
}

(function _debugPickSapValorWideRow() {
  const wideLine = new Array(20).fill(null);
  wideLine[1] = 'CENCOSUD BRASIL COMERCIAL S.A';
  wideLine[2] = '07/07/2026';
  wideLine[3] = '100796';
  wideLine[9] = '354,45'; // Valor Líquido (col J) — must NOT become NF total
  wideLine[13] = '0,00';
  wideLine[14] = '0,00';
  wideLine[15] = '354,45';
  wideLine[17] = '4 992,26 BRL';
  const picked = pickSapValorFromLine(wideLine, { cliente: 1, dtEmissao: 2, nf: 3, valorNF: 13 }, []);
  const pickedDoc = parseSapNum(picked.doc);
  if (Math.abs(pickedDoc - 4992.26) > 0.01) {
    console.warn('[SAP valor] wide-row BRL doc pick failed', picked, pickedDoc);
  } else {
    console.debug('[SAP valor] wide-row BRL doc pick OK', pickedDoc);
  }
  // Headered ZFACT: Valor Líquido + Valor Bruto — always prefer bruto / BRL doc, never líquido.
  const hdr = new Array(20).fill(null);
  hdr[1] = 'Cliente'; hdr[2] = 'Dt.emissão'; hdr[3] = 'Nota Fiscal';
  hdr[9] = 'Valor Líquido'; hdr[13] = 'Valor Bruto'; hdr[17] = 'Valor em BRL';
  const lineBruto = new Array(20).fill(null);
  lineBruto[1] = 'CLIENTE X'; lineBruto[2] = '01/05/2026'; lineBruto[3] = '99635';
  lineBruto[9] = '40.000,00'; lineBruto[13] = '49.937,37'; lineBruto[17] = '49 937,37 BRL';
  const pickedHdr = pickSapValorFromLine(lineBruto, { cliente: 1, dtEmissao: 2, nf: 3, valorNF: 13 }, hdr);
  const lineN = parseSapNum(pickedHdr.line);
  const docN = parseSapNum(pickedHdr.doc);
  const liquidoUsed = Math.abs(lineN - 40000) < 0.01 && Math.abs(docN - 49937.37) > 0.01;
  if (liquidoUsed || (Math.abs(lineN - 49937.37) > 0.01 && Math.abs(docN - 49937.37) > 0.01)) {
    console.warn('[SAP valor] líquido trap — expected bruto/doc 49937.37, got', pickedHdr, lineN, docN);
  } else {
    console.debug('[SAP valor] líquido trap OK', { line: lineN, doc: docN });
  }
  const resolved = resolveSapValorColFromHeader(hdr);
  if (resolved.lineCol !== 13) {
    console.warn('[SAP valor] bruto header col expected 13, got', resolved);
  }
  // NF 100934 / ZFACT NFe: AI Valor líquido 6951.56 ≠ Unilog; AP Val.total incl.imp. 6673.50 matches.
  const hdrAp = new Array(45).fill(null);
  hdrAp[0] = 'Nº da nota fiscal eletrônica';
  hdrAp[14] = 'Nome';
  hdrAp[24] = 'Material';
  hdrAp[34] = 'Valor líquido';
  hdrAp[41] = 'Val.total incl.imp.';
  const line100934 = new Array(45).fill(null);
  line100934[0] = '100934';
  line100934[14] = 'CLIENTE Y';
  line100934[24] = '7801547';
  line100934[34] = '6 951,56';
  line100934[41] = '6 673,50';
  const pickedAp = pickSapValorFromLine(line100934, { cliente: 14, dtEmissao: 43, nf: 0, valorNF: 41 }, hdrAp);
  const apDoc = parseSapNum(pickedAp.doc);
  const apLine = parseSapNum(pickedAp.line);
  const resolvedAp = resolveSapValorColFromHeader(hdrAp);
  if (resolvedAp.docCol !== 41 || Math.abs(apDoc - 6673.5) > 0.01 || Math.abs(apLine - 6951.56) < 0.01) {
    console.warn('[SAP valor] AP/Val.total trap — expected doc 6673.50 not líquido 6951.56, got',
      { resolvedAp, pickedAp, apDoc, apLine });
  } else {
    console.debug('[SAP valor] AP/Val.total incl.imp. OK', apDoc);
  }
})();

function looksLikeSapDataRow(line, colIdx) {
  const idx = colIdx || activeSapColIdx;
  if (!Array.isArray(line) || line.length <= idx.nf) return false;
  if (!looksLikeSapNfCell(line[idx.nf])) return false;
  const cliente = line[idx.cliente];
  const dtRaw = line[idx.dtEmissao];
  const hasCliente = cliente !== null && cliente !== undefined && String(cliente).trim() !== '';
  const hasDate = !!parseSapBrDate(dtRaw);
  return hasCliente || hasDate;
}

function hasSapPositionalCols(line, colIdx) {
  const idx = colIdx || activeSapColIdx;
  if (!Array.isArray(line) || line.length <= idx.nf) return false;
  return looksLikeSapNfCell(line[idx.nf]);
}

function applySapColPositionalFallback(row, line, headerRow, standardLayout, colIdx) {
  const idx = colIdx || activeSapColIdx;
  if (!line || !hasSapPositionalCols(line, idx)) return row;

  const nf = line[idx.nf];
  const cliente = line[idx.cliente];
  const dtRaw = line[idx.dtEmissao];
  const picked = pickSapValorFromLine(line, idx, headerRow);

  if (nf !== null && nf !== undefined && String(nf).trim() !== '') row.nf = nf;
  if (cliente !== null && cliente !== undefined && String(cliente).trim() !== '') {
    row.cliente = String(cliente).trim();
  }
  const parsed = parseSapBrDate(dtRaw);
  if (parsed) {
    row.dtEmissao = parsed;
    row._sapDateSource = 'colC';
  }
  if (picked.line != null && picked.line !== '') {
    const dupDoc = row.valorDoc != null && row.valorDoc !== ''
      && sapValoresClose(sapParsedValor(picked.line), sapParsedValor(row.valorDoc));
    if (!dupDoc && (!row.valorLine || !looksLikeSapValorCell(row.valorLine, { trustColumn: true }))) {
      row.valorLine = picked.line;
    }
  }
  if (picked.doc != null && picked.doc !== '') {
    if (!row.valorDoc || !looksLikeSapValorCell(row.valorDoc, { trustColumn: true })) row.valorDoc = picked.doc;
    if (picked.fromInclImp) row.docFromInclImp = true;
  }
  // Doc total (AP) is the Unilog-matching NF value — prefer over line bruto.
  row.valorNF = row.valorDoc ?? row.valorLine ?? row.valorNF;
  row._sapPositionalLayout = true;
  return row;
}

function isStandardSapNfLayout(headerRow, sampleRows, colIdx) {
  const idx = colIdx || activeSapColIdx;
  if (headerRow &&
      isSapHeaderAtCol(headerRow, idx.cliente, 'cliente') &&
      isSapHeaderAtCol(headerRow, idx.dtEmissao, 'dtEmissao') &&
      isSapHeaderAtCol(headerRow, idx.nf, 'nf')) return true;
  let hits = 0;
  for (const line of (sampleRows || []).slice(0, 5)) {
    if (looksLikeSapDataRow(line, idx)) hits++;
  }
  return hits >= 2;
}

function countParseableSapDatesInCol(dataLines, col, sample = 20) {
  if (col == null || col < 0) return 0;
  let n = 0;
  for (const line of (dataLines || []).slice(0, sample)) {
    if (parseSapBrDate(line?.[col])) n++;
  }
  return n;
}

/**
 * Wide ZFACT NFe: header "Data de criação" may sit on AQ while Excel serials are on AR.
 * Prefer the column (header / neighbour / scan) with the most parseable dates.
 */
function resolveSapDtEmissaoCol(headerRow, dataLines) {
  let hdrCol = null;
  (headerRow || []).forEach((h, j) => {
    if (hdrCol == null && isSapHeaderAtCol(headerRow, j, 'dtEmissao')) hdrCol = j;
  });
  const candidates = new Set();
  if (hdrCol != null) {
    candidates.add(hdrCol);
    if (hdrCol + 1 < 80) candidates.add(hdrCol + 1);
    if (hdrCol - 1 >= 0) candidates.add(hdrCol - 1);
  }
  candidates.add(2);  // classic col C
  candidates.add(43); // wide ZFACT AR (Data/criação serials)
  let best = hdrCol != null ? hdrCol : 2;
  let bestN = countParseableSapDatesInCol(dataLines, best);
  candidates.forEach(c => {
    const n = countParseableSapDatesInCol(dataLines, c);
    if (n > bestN) { bestN = n; best = c; }
  });
  if (bestN < 3) {
    const width = Math.min(60, Math.max(0, ...(dataLines || []).slice(0, 8).map(l => (l && l.length) || 0)));
    for (let c = 0; c < width; c++) {
      const n = countParseableSapDatesInCol(dataLines, c);
      if (n > bestN) { bestN = n; best = c; }
    }
  }
  return best;
}

function resolveSapClienteColFromHeader(headerRow) {
  if (!headerRow || !headerRow.length) return null;
  let prefer = null;
  let weak = null;
  headerRow.forEach((hdr, j) => {
    const h = normCol(hdr);
    if (!h) return;
    if (prefer == null && (h === 'nome' || h === 'nome cliente' || h === 'cliente' || h === 'cardname')) prefer = j;
    if (weak == null && SAP_ALIASES.cliente.some(a => h === a || h.includes(a) || a.includes(h))) weak = j;
  });
  return prefer != null ? prefer : weak;
}

function scoreSapLayout(dataLines, colIdx, headerRow) {
  let hits = 0;
  let dateHits = 0;
  let valorHits = 0;
  for (const line of (dataLines || []).slice(0, 15)) {
    if (looksLikeSapDataRow(line, colIdx)) hits++;
    if (parseSapBrDate(line?.[colIdx.dtEmissao])) dateHits++;
    const v = pickSapValorFromLine(line, colIdx, headerRow);
    if (looksLikeSapValorCell(v.line || v.doc)) valorHits++;
  }
  // Date hits weigh heavily — B/C/D without dates must not beat wide AR layouts.
  return { hits, dateHits, valorHits, score: hits * 10 + dateHits * 25 + valorHits };
}

function finalizeSapColIdx(candidate, headerRow, dataLines, headerValor, headerNfCol, headerClienteCol) {
  const out = { ...candidate };
  if (headerNfCol != null) out.nf = headerNfCol;
  if (headerClienteCol != null) out.cliente = headerClienteCol;
  if (headerValor.docCol != null) out.valorNF = headerValor.docCol;
  else if (headerValor.lineCol != null && out.valorNF == null) out.valorNF = headerValor.lineCol;
  out.dtEmissao = resolveSapDtEmissaoCol(headerRow, dataLines);
  return out;
}

function resolveSapColIdx(headerRow, dataLines) {
  const headerValor = resolveSapValorColFromHeader(headerRow);
  const headerNfCol = resolveSapNfColFromHeader(headerRow);
  const headerClienteCol = resolveSapClienteColFromHeader(headerRow);
  const valorCol = headerValor.docCol ?? headerValor.lineCol ?? SAP_ZFACT_BRUTO_COL;
  const dtCol = resolveSapDtEmissaoCol(headerRow, dataLines);
  const layouts = [
    { cliente: headerClienteCol ?? 13, dtEmissao: dtCol, nf: headerNfCol ?? 21, valorNF: valorCol, label: 'wide ZFACT Nome/Data/Doc' },
    { cliente: 1, dtEmissao: 2, nf: 3, valorNF: valorCol, label: 'B/C/D + valor (ZFACT)' },
    { cliente: 1, dtEmissao: 2, nf: 3, valorNF: headerValor.lineCol ?? 4, label: 'B/C/D/E' },
    { cliente: 1, dtEmissao: 2, nf: 3, valorNF: headerValor.lineCol ?? 5, label: 'B/C/D/F' },
    { cliente: 1, dtEmissao: 2, nf: 3, valorNF: headerValor.lineCol ?? 6, label: 'B/C/D/G' },
    { cliente: 0, dtEmissao: 1, nf: 2, valorNF: headerValor.lineCol ?? 3, label: 'A/B/C/D' },
    { cliente: 2, dtEmissao: 3, nf: 4, valorNF: headerValor.lineCol ?? 5, label: 'C/D/E/F' },
    { cliente: 2, dtEmissao: 3, nf: 4, valorNF: headerValor.lineCol ?? 6, label: 'C/D/E/G' },
    { cliente: 13, dtEmissao: 43, nf: 21, valorNF: valorCol, label: 'N/AR/V (ZFACT NFe)' }
  ];
  for (const layout of layouts) {
    const nfCol = headerNfCol != null ? headerNfCol : layout.nf;
    const candidate = finalizeSapColIdx(
      { ...layout, nf: nfCol },
      headerRow, dataLines, headerValor, headerNfCol, headerClienteCol
    );
    if (headerRow &&
        isSapHeaderAtCol(headerRow, candidate.cliente, 'cliente') &&
        (isSapHeaderAtCol(headerRow, candidate.dtEmissao, 'dtEmissao')
          || isSapHeaderAtCol(headerRow, (candidate.dtEmissao || 1) - 1, 'dtEmissao')
          || countParseableSapDatesInCol(dataLines, candidate.dtEmissao) >= 3) &&
        (headerNfCol != null || isSapHeaderAtCol(headerRow, candidate.nf, 'nf'))) {
      console.debug('[SAP NF] Layout por cabeçalho:', candidate.label,
        `dtCol=${candidate.dtEmissao}`,
        headerNfCol != null ? `(NF col ${headerNfCol})` : '',
        headerValor.docCol != null ? `(Val.total/AP col ${headerValor.docCol})` : '',
        headerValor.lineCol != null ? `(valor bruto col ${headerValor.lineCol})` : '');
      return candidate;
    }
  }
  let best = finalizeSapColIdx(layouts[0], headerRow, dataLines, headerValor, headerNfCol, headerClienteCol);
  let bestResult = { hits: 0, dateHits: 0, valorHits: 0, score: 0 };
  for (const layout of layouts) {
    const candidate = finalizeSapColIdx(
      headerNfCol != null ? { ...layout, nf: headerNfCol } : layout,
      headerRow, dataLines, headerValor, headerNfCol, headerClienteCol
    );
    const r = scoreSapLayout(dataLines, candidate, headerRow);
    if (r.score > bestResult.score
      || (r.score === bestResult.score && r.dateHits > bestResult.dateHits)
      || (r.score === bestResult.score && r.dateHits === bestResult.dateHits && r.hits > bestResult.hits)) {
      bestResult = r;
      best = candidate;
    }
  }
  if (bestResult.hits >= 2 || (bestResult.hits >= 1 && bestResult.valorHits >= 1) || bestResult.dateHits >= 3) {
    console.debug('[SAP NF] Layout detectado por dados:', best.label,
      `(${bestResult.hits} linhas, ${bestResult.dateHits} datas, ${bestResult.valorHits} valores, dtCol=${best.dtEmissao})`);
    return best;
  }
  console.debug('[SAP NF] Layout padrão B/C/D (fallback) + dt scan');
  return finalizeSapColIdx({ ...SAP_COL_IDX_DEFAULT }, headerRow, dataLines, headerValor, headerNfCol, headerClienteCol);
}

function logSapRowDebug(line, row, idx, colIdx) {
  const c = colIdx || activeSapColIdx;
  const b = line?.[c.cliente];
  const dt = line?.[c.dtEmissao];
  const d = line?.[c.nf];
  console.debug(`[SAP NF] row ${idx + 1} raw B/C/D:`, b, dt, d,
    '→ parsed:', row?.cliente || '(vazio)', row?.dtEmissao || '(vazio)', row?.nf || '(vazio)');
}

function summarizeSapDateSources(rows) {
  const counts = {};
  rows.forEach(r => {
    const src = r._sapDateSource || (r._sapPositionalLayout ? 'colC' : 'alias');
    counts[src] = (counts[src] || 0) + 1;
  });
  return counts;
}

function parseSapSheetRows(sheet) {
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  if (!raw.length) return { rows: [], headers: [] };

  let headerIdx = raw.findIndex(row => Array.isArray(row) && isSapHeaderRow(row));
  let dataStart = headerIdx >= 0 ? headerIdx + 1 : 0;
  let headerRow = headerIdx >= 0 ? (raw[headerIdx] || []) : [];

  if (headerIdx < 0 && (looksLikeSapDataRow(raw[0]) || hasSapPositionalCols(raw[0]))) {
    dataStart = 0;
    headerRow = [];
  } else if (headerIdx < 0) {
    headerIdx = 0;
    headerRow = raw[0] || [];
    dataStart = 1;
  }

  const headers = headerRow.map(h => String(h || '').trim()).filter(Boolean);
  const dataLines = [];
  for (let i = dataStart; i < raw.length; i++) {
    const line = raw[i];
    if (!Array.isArray(line) || line.every(c => c === null || c === undefined || c === '')) continue;
    dataLines.push(line);
  }

  activeSapColIdx = resolveSapColIdx(headerRow, dataLines);
  const standardLayout = isStandardSapNfLayout(headerRow, dataLines, activeSapColIdx);
  const hasMaterialCol = (headerRow || []).some(h => sapHeaderMatchesField(h, SAP_ALIASES.material));
  if (standardLayout) {
    console.debug('[SAP NF] Layout padrão: col. B=cliente, C=data emissão, D=nota fiscal');
  }

  const rows = [];
  for (let i = 0; i < dataLines.length; i++) {
    const line = dataLines[i];
    const obj = {};
    headerRow.forEach((h, j) => { if (h) obj[h] = line[j] ?? null; });
    let row = normalizeSapRow(obj);
    row._hasMaterialCol = hasMaterialCol;
    row = applySapColPositionalFallback(row, line, headerRow, standardLayout, activeSapColIdx);
    row = sanitizeSapRowValors(row);
    if (i < 3) logSapRowDebug(line, row, i, activeSapColIdx);
    rows.push(row);
  }
  return { rows, headers };
}

function normalizeSapRow(row) {
  const dtRaw = findSapField(row, 'dtEmissao');
  const valores = findSapValorFields(row);
  const materialRaw = findSapField(row, 'material');
  return {
    nf: findSapField(row, 'nf'),
    dtEmissao: parseSapBrDate(dtRaw),
    cliente: findSapField(row, 'cliente'),
    material: materialRaw != null && materialRaw !== '' ? String(materialRaw).trim() : '',
    valorLine: valores.line,
    valorDoc: valores.doc,
    valorNF: valores.primary,
    docFromInclImp: !!valores.fromInclImp
  };
}

function isHeaderRow(cells) {
  const normalized = cells.map(c => normCol(c));
  return FIELD_ALIASES.nf.some(alias => normalized.some(c => c === alias || c.includes(alias) || alias.includes(c)));
}

function isSapHeaderRow(cells) {
  const normalized = cells.map(c => normCol(c));
  return SAP_ALIASES.nf.some(alias => normalized.some(c => c === alias || c.includes(alias) || alias.includes(c)));
}

function isQtdCteHeader(hdr) {
  const h = normCol(hdr);
  if (!h) return true;
  if (h.includes('qtd') || h.includes('quant')) return true;
  return FIELD_ALIASES.qtdCte.some(a => h === a || h.includes(a) || a.includes(h));
}

function applyColEQtdFallback(row, line, headerRow) {
  const eHdr = headerRow?.[4];
  const eVal = line?.[4];
  if (eVal === null || eVal === undefined || eVal === '') return row;
  if (row.qtdCte !== null && row.qtdCte !== undefined && row.qtdCte !== '') return row;
  if (!isQtdCteHeader(eHdr)) return row;
  row.qtdCte = eVal;
  return row;
}

function parseSheetRows(sheet, normalizeFn, headerFn, rowHook) {
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  if (!raw.length) return { rows: [], headers: [] };

  let headerIdx = raw.findIndex(row => Array.isArray(row) && headerFn(row));
  if (headerIdx < 0) headerIdx = 0;

  const headerRow = raw[headerIdx] || [];
  const headers = headerRow.map(h => String(h || '').trim()).filter(Boolean);
  const rows = [];
  for (let i = headerIdx + 1; i < raw.length; i++) {
    const line = raw[i];
    if (!Array.isArray(line) || line.every(c => c === null || c === undefined || c === '')) continue;
    const obj = {};
    headerRow.forEach((h, j) => { if (h) obj[h] = line[j] ?? null; });
    let row = normalizeFn(obj);
    if (rowHook) row = rowHook(row, line, headerRow);
    row = applyColEQtdFallback(row, line, headerRow);
    rows.push(row);
  }
  return { rows, headers };
}

function countValidRows(rows) {
  return rows.filter(r => r.nf !== null && r.nf !== undefined && String(r.nf).trim() !== '').length;
}

function loadRowsFromWorkbook(wb) {
  const preferred = ['Export', 'Dados', 'Data', 'Conciliacao', 'Conciliação', 'Board'];
  const names = [...new Set([...preferred.filter(n => wb.SheetNames.includes(n)), ...wb.SheetNames])];
  let best = { rows: [], sheetName: names[0], headers: [], valid: 0 };

  for (const name of names) {
    const parsed = parseSheetRows(wb.Sheets[name], normalizeRow, isHeaderRow);
    const valid = countValidRows(parsed.rows);
    if (valid > best.valid) best = { rows: parsed.rows, sheetName: name, headers: parsed.headers, valid };
  }
  return best;
}

function loadSapRowsFromWorkbook(wb) {
  const preferred = ['SAP', 'Notas', 'NFs', 'Export', 'Dados', 'Data'];
  const names = [...new Set([...preferred.filter(n => wb.SheetNames.includes(n)), ...wb.SheetNames])];
  let best = { rows: [], sheetName: names[0], headers: [], valid: 0 };

  for (const name of names) {
    const parsed = parseSapSheetRows(wb.Sheets[name]);
    const valid = countValidRows(parsed.rows);
    if (valid > best.valid) best = { rows: parsed.rows, sheetName: name, headers: parsed.headers, valid };
  }
  return best;
}

/** Skip header/subtotal rows when ZFACT has a material/item column. */
function sapRowIsBillableLine(r) {
  if (!r._hasMaterialCol) return true;
  return !!(r.material && String(r.material).trim());
}

function sapParsedValor(v) {
  if (!looksLikeSapValorCell(v, { trustColumn: true })) return 0;
  return parseSapNum(v);
}

/** Line-level valor bruto from ZFACT row — validated before aggregation. */
function sapRowLineValor(r) {
  return sapParsedValor(r.valorLine ?? (!r.valorDoc ? r.valorNF : null));
}

/** Doc-level NF total — same on every line; must not be summed across rows. */
function sapRowDocValor(r) {
  return sapParsedValor(r.valorDoc);
}

function sapValoresClose(a, b, eps = 0.01) {
  return Math.abs(a - b) < eps;
}

/** Drop valorLine when it duplicates valorDoc (doc total mistaken for line bruto). */
function sanitizeSapRowValors(r) {
  const doc = sapParsedValor(r.valorDoc);
  const line = sapParsedValor(r.valorLine);
  if (doc > 0 && line > 0 && sapValoresClose(line, doc)) r.valorLine = null;
  return r;
}

function finalizeSapNfValor(acc) {
  const vals = acc.lineVals || [];
  const lineSum = acc.lineSum;
  const doc = acc.docVal > 0 ? acc.docVal : 0;

  // FORCE: programmed ZFACT col "Val.total incl.imp." (AP) = Unilog face — once per NF.
  // Never lose to SUM(valor bruto) even when the gap is modest (NF 99977: 449796 vs 442314.49).
  if (acc.docFromInclImp && doc > 0) return doc;

  // Unilog "Valor NF" = invoice face (gross). ZFACT may expose:
  // - lineSum = SUM valor bruto (face) OR partial líquido lines
  // - docVal  = AP incl.imp. / repeating "N BRL" (often face, sometimes net without tax)
  // Rule: partial lines (sum << doc) → doc once (never SUM AP); modest bruto > net → lineSum.
  // NEVER sum identical repeating doc/face amounts × line count (NF 99977: 442k × 25 ≈ 11M).
  if (vals.length > 0) {
    const allIdentical = vals.every(v => sapValoresClose(v, vals[0]));
    const nearDocCount = doc > 0
      ? vals.filter(v => sapValoresClose(v, doc, Math.max(0.05, Math.abs(doc) * 0.002))).length
      : 0;

    // Doc total repeated as line bruto on every row — take once.
    if (allIdentical && doc > 0 && sapValoresClose(vals[0], doc)) return doc;
    if (nearDocCount >= 2 && nearDocCount === vals.length) return doc;

    if (doc > 0) {
      if (sapValoresClose(lineSum, doc, 0.05)) return doc;
      const ref = Math.max(Math.abs(doc), Math.abs(lineSum), 1);
      const gap = Math.abs(lineSum - doc) / ref;
      if (gap > 0.02) {
        // Partial líquido lines (e.g. 354+620 vs 4 992 BRL) → doc is the NF face.
        if (lineSum < doc) return doc;
        // Inflated: identical (or near-doc) "lines" summed ×N — take doc once.
        // Real bruto>net premium is small (NF 101267 ~6%); ×1.5+ is classic doc×lines bug.
        if (lineSum > doc * 1.5) {
          if (allIdentical || nearDocCount >= 2) return doc;
          if (lineSum > doc * 3) return doc;
        }
        // Bruto/line face > smaller net/BRL doc (NF 101267: 2259,34 vs 2121,44) → lineSum.
        return lineSum;
      }
    }

    // Identical values on every row without a usable doc — repeating face, take once.
    // Keep summing short material lists (2 equal billable lines / same-price SKUs).
    if (allIdentical && vals.length > 1) {
      if (!(acc.hasMaterialCol && vals.length <= 2)) return vals[0];
    }

    // ZFACT with material/item lines: SUM every billable line (even equal bruto / same SKU).
    if (acc.hasMaterialCol) return lineSum;

    return lineSum;
  }

  // Doc only (e.g. AP Val.total incl.imp. with líquido excluded from lines) — once, never sum.
  if (doc > 0) return doc;

  if (acc.fallbackVals.length) {
    const fb = acc.fallbackVals;
    // Repeating document-level fallback on every line — once, never sum.
    if (fb.every(v => sapValoresClose(v, fb[0]))) return fb[0];
    return Math.max(...fb);
  }
  return 0;
}

/** Among SAP candidates, prefer the amount closest to Unilog NF face (not líquido). */
function pickSapValorNearUnilog(sap, unilogVal) {
  if (!sap) return null;
  const u = num(unilogVal);
  const docN = num(sap.docVal);

  // Programmed AP / Val.total incl.imp. — always wins (Unilog face).
  if (sap.docFromInclImp && docN > 0) return docN;
  // Doc matches Unilog within cent tolerance → Δ=0 candidate.
  if (u > 0 && docN > 0 && Math.abs(docN - u) <= SAP_VALOR_NEAR_EPS) return docN;

  const candidates = [];
  const add = (v) => {
    const n = num(v);
    if (!(n > 0)) return;
    if (candidates.some(c => sapValoresClose(c, n))) return;
    candidates.push(n);
  };
  add(sap.valorNF);
  add(sap.lineSum);
  add(sap.docVal);
  if (!candidates.length) return sap.valorNF ?? null;
  if (!(u > 0)) return sap.valorNF ?? candidates[0];
  let best = candidates[0];
  let bestDiff = Math.abs(best - u);
  for (let i = 1; i < candidates.length; i++) {
    const c = candidates[i];
    const d = Math.abs(c - u);
    if (d < bestDiff - 1e-9) {
      best = c;
      bestDiff = d;
    }
  }
  // Inflated sum (doc×lines): if |sap| > 3×|unilog| and a once-doc candidate is near Unilog, use it.
  if (Math.abs(best) > 3 * Math.abs(u)) {
    let nearBest = null;
    let nearDiff = Infinity;
    for (const c of candidates) {
      if (Math.abs(c) > 3 * Math.abs(u)) continue;
      const d = Math.abs(c - u);
      if (d <= Math.abs(u) * 0.5 && d < nearDiff) {
        nearBest = c;
        nearDiff = d;
      }
    }
    if (nearBest != null) return nearBest;
  }
  return best;
}

const SAP_NF_DEBUG_KEYS = new Set(['99635', '99636', '99641', '99642', '99633', '18596', '2223', '18591', '99977']);

/** GROUP BY NF key → doc total (AP) once per NF; else SUM valor bruto of billable lines. */
function buildSapNfMap(rows) {
  const map = {};

  function ensureAcc(r, key) {
    if (!map[key]) {
      map[key] = {
        nf: r.nf,
        cliente: r.cliente ? String(r.cliente).trim() : '',
        dtEmissao: parseSapBrDate(r.dtEmissao),
        lineSum: 0,
        lineCount: 0,
        lineVals: [],
        docVal: 0,
        docFromInclImp: false,
        fallbackVals: [],
        hasMaterialCol: false,
        materialKeys: new Set()
      };
    }
    return map[key];
  }

  // Pass 1: doc total from every row (incl. NF header rows without material).
  rows.forEach(r => {
    r = sanitizeSapRowValors(r);
    const nf = r.nf;
    if (nf === null || nf === undefined || String(nf).trim() === '') return;
    const key = normNFKey(nf);
    if (!key) return;
    const docValor = sapRowDocValor(r);
    if (docValor > 0) {
      const acc = ensureAcc(r, key);
      acc.docVal = Math.max(acc.docVal, docValor);
      if (r.docFromInclImp) acc.docFromInclImp = true;
    }
  });

  rows.forEach(r => {
    r = sanitizeSapRowValors(r);
    const nf = r.nf;
    if (nf === null || nf === undefined || String(nf).trim() === '') return;
    if (!sapRowIsBillableLine(r)) return;
    const key = normNFKey(nf);
    if (!key) return;

    const lineValor = sapRowLineValor(r);
    const docValor = sapRowDocValor(r);
    const cliente = r.cliente ? String(r.cliente).trim() : '';
    const dtEmissao = parseSapBrDate(r.dtEmissao);

    const acc = ensureAcc(r, key);

    if (r._hasMaterialCol) acc.hasMaterialCol = true;
    if (r._hasMaterialCol && r.material) {
      const matKey = String(r.material).trim();
      if (matKey) acc.materialKeys.add(matKey);
    }

    if (docValor > 0) {
      acc.docVal = Math.max(acc.docVal, docValor);
      if (r.docFromInclImp) acc.docFromInclImp = true;
    }
    if (lineValor !== 0) {
      // Never accumulate a line that duplicates the NF doc total (repeats on every item row).
      const docRef = docValor > 0 ? docValor : acc.docVal;
      if (!(docRef > 0 && sapValoresClose(lineValor, docRef))) {
        acc.lineSum += lineValor;
        acc.lineCount++;
        acc.lineVals.push(lineValor);
      }
    } else if (!r.valorLine && !r.valorDoc && r.valorNF) {
      const fb = sapParsedValor(r.valorNF);
      if (fb > 0) acc.fallbackVals.push(fb);
    }

    if (!acc.cliente && cliente) acc.cliente = cliente;
    if (!acc.dtEmissao && dtEmissao) acc.dtEmissao = dtEmissao;
  });

  const multiLine = Object.values(map).filter(e => e.lineCount > 1).length;
  if (multiLine) {
    console.debug('[SAP NF] ZFACT agregação: %d NFs com várias linhas (valor bruto somado) de %d NFs únicas',
      multiLine, Object.keys(map).length);
  }

  Object.keys(map).forEach(key => {
    const acc = map[key];
    const valorNF = finalizeSapNfValor(acc);
    if (SAP_NF_DEBUG_KEYS.has(key)) {
      console.debug('[SAP NF] buildSapNfMap debug', key, {
        nf: acc.nf,
        lineCount: acc.lineCount,
        lineVals: acc.lineVals,
        lineSum: acc.lineSum,
        docVal: acc.docVal,
        fallbackVals: acc.fallbackVals,
        valorNF
      });
    }
    map[key] = {
      nf: acc.nf,
      cliente: acc.cliente,
      dtEmissao: acc.dtEmissao,
      valorNF,
      lineSum: acc.lineSum || 0,
      docVal: acc.docVal || 0,
      docFromInclImp: !!acc.docFromInclImp
    };
  });
  return map;
}

/** Console self-check — ZFACT aggregation must match Unilog NF total semantics. */
function _debugBuildSapNfMapAggregation() {
  const rows = [
    { nf: '97723', cliente: 'A', dtEmissao: '01/01/2025', valorLine: '1.000,00' },
    { nf: '97723-1', cliente: 'A', dtEmissao: '01/01/2025', valorLine: '2.500,50' },
    { nf: '88888', cliente: 'B', dtEmissao: '02/01/2025', valorLine: 9540067078 },
    { nf: '88888', cliente: 'B', dtEmissao: '02/01/2025', valorLine: '500,00' },
    { nf: '99999', cliente: 'C', dtEmissao: '03/01/2025', valorLine: '' },
    { nf: '99635', cliente: 'RALLU', dtEmissao: '01/05/2026', valorDoc: '49.937,37' },
    { nf: '99635', cliente: 'RALLU', dtEmissao: '01/05/2026', valorDoc: '49.937,37' },
    { nf: '99635', cliente: 'RALLU', dtEmissao: '01/05/2026', valorDoc: '49.937,37' },
    { nf: '99635', cliente: 'RALLU', dtEmissao: '01/05/2026', valorLine: '10.000,00' },
    { nf: '99635', cliente: 'RALLU', dtEmissao: '01/05/2026', valorLine: '20.000,00' },
    { nf: '99635', cliente: 'RALLU', dtEmissao: '01/05/2026', valorLine: '19.937,37' }
  ];
  const map = buildSapNfMap(rows);
  const mapDocOnly = buildSapNfMap([
    { nf: '99635', valorDoc: '49.937,37' },
    { nf: '99635', valorDoc: '49.937,37' },
    { nf: '99635', valorDoc: '49.937,37' }
  ]);
  const mapBrutoRepeat = buildSapNfMap([
    { nf: '99635', valorLine: '49.937,37' },
    { nf: '99635', valorLine: '49.937,37' },
    { nf: '99635', valorLine: '49.937,37' },
    { nf: '99635', valorLine: '49.937,37' }
  ]);
  const mapDupPositional = buildSapNfMap([
    { nf: '99635', valorLine: '49.937,37', valorDoc: '49.937,37' },
    { nf: '99635', valorLine: '49.937,37', valorDoc: '49.937,37' },
    { nf: '99635', valorLine: '49.937,37', valorDoc: '49.937,37' },
    { nf: '99635', valorLine: '49.937,37', valorDoc: '49.937,37' }
  ]);
  const mapLineDocSameRow = buildSapNfMap([
    { nf: '99635', valorLine: '10.000,00', valorDoc: '49.937,37' },
    { nf: '99635', valorLine: '20.000,00', valorDoc: '49.937,37' },
    { nf: '99635', valorLine: '19.937,37', valorDoc: '49.937,37' }
  ]);
  const mapInflatedRepeat = buildSapNfMap([
    { nf: '99635', valorLine: '49.937,37' },
    { nf: '99635', valorLine: '49.937,37' },
    { nf: '99635', valorLine: '49.937,37' },
    { nf: '99635', valorLine: '25.000,00' }
  ]);
  const map99635brutoZfact = buildSapNfMap([
    { nf: '000099635-2', _hasMaterialCol: true, material: 'SKU1', valorLine: '641,50' },
    { nf: '000099635-2', _hasMaterialCol: true, material: 'SKU2', valorLine: '1.453,33' },
    { nf: '000099635-2', _hasMaterialCol: true, material: 'SKU3', valorLine: '4.542,08' },
    { nf: '000099635-2', _hasMaterialCol: true, material: 'SKU4', valorLine: '43.300,46' }
  ]);
  const map99635feeLines = buildSapNfMap([
    { nf: '99635', _hasMaterialCol: true, material: 'SKU1', valorLine: '641,50' },
    { nf: '99635', _hasMaterialCol: true, material: 'SKU2', valorLine: '10.000,00' },
    { nf: '99635', _hasMaterialCol: true, material: 'SKU3', valorLine: '39.295,87' }
  ]);
  const map99636bruto = buildSapNfMap([
    { nf: '99636', _hasMaterialCol: true, material: 'SKU1', valorLine: '30.000,00' },
    { nf: '99636', _hasMaterialCol: true, material: 'SKU2', valorLine: '28.266,08' }
  ]);
  const map99636lineDoc = buildSapNfMap([
    { nf: '99636', _hasMaterialCol: true, material: 'SKU1', valorLine: '30.000,00', valorDoc: '58.266,08' },
    { nf: '99636', _hasMaterialCol: true, material: 'SKU2', valorLine: '28.266,08', valorDoc: '58.266,08' }
  ]);
  const map2223 = buildSapNfMap([
    { nf: '2223', material: 'X', valorLine: '12.345,67' }
  ]);
  const map18591 = buildSapNfMap([
    { nf: '000018591', material: 'Y', valorLine: '8.500,00' },
    { nf: '000018591', material: 'Z', valorLine: '1.200,50' }
  ]);
  const map99641 = buildSapNfMap([
    { nf: '000099641-2', _hasMaterialCol: true, material: '', valorLine: '0,00' },
    { nf: '000099641-2', _hasMaterialCol: true, material: 'SKU-A', valorLine: '331,90' },
    { nf: '000099641-2', _hasMaterialCol: true, material: 'SKU-B', valorLine: '331,90' }
  ]);
  const map99641sameMat = buildSapNfMap([
    { nf: '000099641-2', _hasMaterialCol: true, material: 'SKU-A', valorLine: '331,90' },
    { nf: '000099641-2', _hasMaterialCol: true, material: 'SKU-A', valorLine: '331,90' }
  ]);
  const map99642 = buildSapNfMap([
    { nf: '000099642-2', _hasMaterialCol: true, material: '7806814 CUBA', valorLine: '0,00' },
    { nf: '000099642-2', _hasMaterialCol: true, material: '7806814 CUBA', valorLine: 308.91 },
    { nf: '000099642-2', _hasMaterialCol: true, material: '7806814 CUBA', valorLine: 679.59 },
    { nf: '000099642-2', _hasMaterialCol: true, material: '7801734 COLÔMBIA', valorLine: '0,00' },
    { nf: '000099642-2', _hasMaterialCol: true, material: '7801734 COLÔMBIA', valorLine: 988.5 },
    { nf: '000099642-2', _hasMaterialCol: true, material: '7802415 TIMOR', valorLine: '0,00' },
    { nf: '000099642-2', _hasMaterialCol: true, material: '7802415 TIMOR', valorLine: 786 }
  ]);
  // NF 18596: doc 7079 (+11540) cancelled by doc 7310 (-11540); doc 7312 (+24346) is the NF total.
  const map18596 = buildSapNfMap([
    { nf: '000018596-004', _hasMaterialCol: true, material: 'SKU-A', valorLine: '11.540,00' },
    { nf: '000018596-004', _hasMaterialCol: true, material: 'SKU-A', valorLine: '-11.540,00' },
    { nf: '000018596-004', _hasMaterialCol: true, material: 'SKU-B', valorLine: '24.346,00' }
  ]);
  const map18596multi = buildSapNfMap([
    { nf: '000018596-004', _hasMaterialCol: true, material: 'M1', valorLine: '6.000,00' },
    { nf: '000018596-004', _hasMaterialCol: true, material: 'M2', valorLine: '6.000,00' },
    { nf: '000018596-004', _hasMaterialCol: true, material: 'M1', valorLine: '-6.000,00' },
    { nf: '000018596-004', _hasMaterialCol: true, material: 'M2', valorLine: '-6.000,00' },
    { nf: '000018596-004', _hasMaterialCol: true, material: 'M3', valorLine: '24.346,00' }
  ]);
  // NF 100796 / ZFACT: valor bruto cols are 0,00; NF total repeats as "4 992,26 BRL" (use once).
  const map100796brl = buildSapNfMap([
    { nf: '100796', _hasMaterialCol: true, material: 'DELTA Q', valorLine: '0,00', valorDoc: '4 992,26 BRL' },
    { nf: '100796', _hasMaterialCol: true, material: 'DELTA QH', valorLine: '0,00', valorDoc: '4 992,26 BRL' },
    { nf: '100796', _hasMaterialCol: true, material: 'COLOMBIA', valorLine: '0,00', valorDoc: '4\u00a0992,26 BRL' }
  ]);
  // Same NF: line nets present + repeating BRL doc total → sum lines (or doc if lines empty).
  const map100796lines = buildSapNfMap([
    { nf: '100796', _hasMaterialCol: true, material: 'A', valorLine: '354,45', valorDoc: '4 992,26 BRL' },
    { nf: '100796', _hasMaterialCol: true, material: 'B', valorLine: '620,29', valorDoc: '4 992,26 BRL' },
    { nf: '100796', _hasMaterialCol: true, material: 'C', valorLine: '4017,52', valorDoc: '4 992,26 BRL' }
  ]);
  // Partial líquido lines (≠ NF total) + BRL doc → must use doc once, not 354+620.
  const map100796partial = buildSapNfMap([
    { nf: '100796', _hasMaterialCol: true, material: 'DELTA Q', valorLine: '354,45', valorDoc: '4 992,26 BRL' },
    { nf: '100796', _hasMaterialCol: true, material: 'DELTA QH', valorLine: '620,29', valorDoc: '4 992,26 BRL' }
  ]);
  // NF 101267: bruto line face 2259,34 vs smaller net/BRL doc 2121,44 → prefer Unilog face (lineSum).
  const map101267 = buildSapNfMap([
    { nf: '101267', _hasMaterialCol: true, material: 'A', valorLine: '1.200,00', valorDoc: '2.121,44 BRL' },
    { nf: '101267', _hasMaterialCol: true, material: 'B', valorLine: '1.059,34', valorDoc: '2.121,44 BRL' }
  ]);
  const map101267netOnly = buildSapNfMap([
    { nf: '101267', _hasMaterialCol: true, material: 'A', valorLine: '1.060,72', valorDoc: '2.121,44 BRL' },
    { nf: '101267', _hasMaterialCol: true, material: 'B', valorLine: '1.060,72', valorDoc: '2.121,44 BRL' }
  ]);
  const ok97723 = Math.abs((map['97723']?.valorNF || 0) - 3500.5) < 0.01;
  const ok88888 = Math.abs((map['88888']?.valorNF || 0) - 500) < 0.01;
  const ok99999 = map['99999']?.valorNF === 0;
  const ok99635lines = Math.abs((map['99635']?.valorNF || 0) - 49937.37) < 0.01;
  const ok99635doc = Math.abs((mapDocOnly['99635']?.valorNF || 0) - 49937.37) < 0.01;
  const ok99635repeat = Math.abs((mapBrutoRepeat['99635']?.valorNF || 0) - 49937.37) < 0.01;
  const ok99635dupPos = Math.abs((mapDupPositional['99635']?.valorNF || 0) - 49937.37) < 0.01;
  const ok99635lineDoc = Math.abs((mapLineDocSameRow['99635']?.valorNF || 0) - 49937.37) < 0.01;
  const ok99635inflate = Math.abs((mapInflatedRepeat['99635']?.valorNF || 0) - 174812.11) < 0.01;
  const ok99635brutoZfact = Math.abs((map99635brutoZfact['99635']?.valorNF || 0) - 49937.37) < 0.01;
  const ok99635feeLines = Math.abs((map99635feeLines['99635']?.valorNF || 0) - 49937.37) < 0.01;
  const ok99636bruto = Math.abs((map99636bruto['99636']?.valorNF || 0) - 58266.08) < 0.01;
  const ok99636lineDoc = Math.abs((map99636lineDoc['99636']?.valorNF || 0) - 58266.08) < 0.01;
  const ok2223 = Math.abs((map2223['2223']?.valorNF || 0) - 12345.67) < 0.01;
  const ok18591 = Math.abs((map18591['18591']?.valorNF || 0) - 9700.5) < 0.01;
  const ok99641 = Math.abs((map99641['99641']?.valorNF || 0) - 663.8) < 0.01;
  const ok99641sameMat = Math.abs((map99641sameMat['99641']?.valorNF || 0) - 663.8) < 0.01;
  const ok99642 = Math.abs((map99642['99642']?.valorNF || 0) - 2763) < 0.01;
  const ok18596 = Math.abs((map18596['18596']?.valorNF || 0) - 24346) < 0.01;
  const ok18596multi = Math.abs((map18596multi['18596']?.valorNF || 0) - 24346) < 0.01;
  const ok100796brl = Math.abs((map100796brl['100796']?.valorNF || 0) - 4992.26) < 0.01;
  const ok100796lines = Math.abs((map100796lines['100796']?.valorNF || 0) - 4992.26) < 0.01;
  const ok100796partial = Math.abs((map100796partial['100796']?.valorNF || 0) - 4992.26) < 0.01;
  const map100934 = buildSapNfMap([normalizeSapRow({
    'Nº da nota fiscal eletrônica': '100934',
    Nome: 'CLIENTE Y',
    Material: '7801547',
    'Valor líquido': '6 951,56',
    'Val.total incl.imp.': '6 673,50'
  })]);
  const ok100934ap = Math.abs((map100934['100934']?.valorNF || 0) - 6673.5) < 0.01;
  // NF 99977: AP Val.total incl.imp. repeats on every item line — take once, never ×N.
  const rows99977 = [];
  for (let i = 0; i < 25; i++) {
    rows99977.push({
      nf: '99977',
      _hasMaterialCol: true,
      material: 'SKU-' + i,
      valorDoc: '442.314,49',
      docFromInclImp: true,
      valorLine: String(1000 + i * 10) + ',00'
    });
  }
  const map99977 = buildSapNfMap(rows99977);
  const ok99977apOnce = Math.abs((map99977['99977']?.valorNF || 0) - 442314.49) < 0.01;
  // Screenshot bug: Valor Bruto = full face on every material line (×25 → ~11M). Take once.
  const map99977brutoFaceRepeat = buildSapNfMap(Array.from({ length: 25 }, (_, i) => ({
    nf: '99977',
    _hasMaterialCol: true,
    material: 'SKU-' + i,
    valorLine: '442.314,49',
    valorDoc: '442.314,49',
    docFromInclImp: true
  })));
  const ok99977brutoFace = Math.abs((map99977brutoFaceRepeat['99977']?.valorNF || 0) - 442314.49) < 0.01;
  // Same pattern with slightly different repeating line face vs Unilog/doc (11244900/25 = 449796).
  const map99977lineInflate = buildSapNfMap(Array.from({ length: 25 }, (_, i) => ({
    nf: '99977',
    _hasMaterialCol: true,
    material: 'SKU-' + i,
    valorLine: '449.796,00',
    valorDoc: '442.314,49',
    docFromInclImp: true
  })));
  const ok99977lineInflate = Math.abs((map99977lineInflate['99977']?.valorNF || 0) - 442314.49) < 0.01;
  // CRITICAL: SUM(bruto)=449796 vs AP 442314.49 (gap ~1.7% < 2%) must still FORCE AP — not lineSum.
  const map99977modestGap = buildSapNfMap([
    { nf: '99977', _hasMaterialCol: true, material: 'A', valorLine: '200.000,00', valorDoc: '442 314,49', docFromInclImp: true },
    { nf: '99977', _hasMaterialCol: true, material: 'B', valorLine: '249.796,00', valorDoc: '442 314,49', docFromInclImp: true }
  ]);
  const ok99977modestGap = Math.abs((map99977modestGap['99977']?.valorNF || 0) - 442314.49) < 0.01
    && Math.abs((map99977modestGap['99977']?.lineSum || 0) - 449796) < 0.01;
  const near99977 = pickSapValorNearUnilog(
    { valorNF: 449796, lineSum: 449796, docVal: 442314.49, docFromInclImp: true },
    442314.49
  );
  const ok99977near = Math.abs((near99977 || 0) - 442314.49) < 0.01;
  // Space thousands in Val.total incl.imp. (Excel "442 314,49").
  const map99977space = buildSapNfMap(Array.from({ length: 3 }, (_, i) => normalizeSapRow({
    'Nº da nota fiscal eletrônica': '99977',
    Nome: 'ARC',
    Material: 'M' + i,
    'Valor Bruto': String(100000 + i * 1000) + ',00',
    'Val.total incl.imp.': '442 314,49'
  })));
  const ok99977space = Math.abs((map99977space['99977']?.valorNF || 0) - 442314.49) < 0.01
    && !!map99977space['99977']?.docFromInclImp;
  // Repeating face without doc column — still take once when ≥3 identical material lines.
  const map99977noDoc = buildSapNfMap(Array.from({ length: 25 }, (_, i) => ({
    nf: '99977',
    _hasMaterialCol: true,
    material: 'SKU-' + i,
    valorLine: '442.314,49'
  })));
  const ok99977noDoc = Math.abs((map99977noDoc['99977']?.valorNF || 0) - 442314.49) < 0.01;
  const map99977normalize = buildSapNfMap(Array.from({ length: 25 }, (_, i) => normalizeSapRow({
    'Nº da nota fiscal eletrônica': '99977',
    Nome: 'ARC TRANSPORTE',
    Material: 'MAT-' + i,
    'Valor Bruto': '1.000,00',
    'Val.total incl.imp.': '442.314,49'
  })));
  const ok99977norm = Math.abs((map99977normalize['99977']?.valorNF || 0) - 442314.49) < 0.01
    && !!map99977normalize['99977']?.docFromInclImp;
  // Header with "impostos" must still count as AP doc total (not excluded).
  const map99977imp = buildSapNfMap([normalizeSapRow({
    'Nº da nota fiscal eletrônica': '99977',
    Nome: 'ARC',
    Material: 'X',
    'Val.total incluindo impostos': '442.314,49'
  })]);
  const ok99977impHdr = Math.abs((map99977imp['99977']?.valorNF || 0) - 442314.49) < 0.01;
  const ok101267 = Math.abs((map101267['101267']?.valorNF || 0) - 2259.34) < 0.01;
  const near101267 = pickSapValorNearUnilog(
    { valorNF: 2121.44, lineSum: 2259.34, docVal: 2121.44 },
    2259.34
  );
  const ok101267near = Math.abs((near101267 || 0) - 2259.34) < 0.01;
  // When lineSum equals net doc, keep doc (no better face candidate in map alone).
  const ok101267netOnly = Math.abs((map101267netOnly['101267']?.valorNF || 0) - 2121.44) < 0.01;
  if (!ok97723 || !ok88888 || !ok99999 || !ok99635lines || !ok99635doc || !ok99635repeat || !ok99635dupPos
      || !ok99635lineDoc || !ok99635inflate || !ok99635brutoZfact || !ok99635feeLines || !ok99636bruto
      || !ok99636lineDoc || !ok2223 || !ok18591 || !ok99641 || !ok99641sameMat || !ok99642
      || !ok18596 || !ok18596multi || !ok100796brl || !ok100796lines || !ok100796partial || !ok100934ap
      || !ok99977apOnce || !ok99977brutoFace || !ok99977lineInflate || !ok99977modestGap || !ok99977near
      || !ok99977space || !ok99977noDoc || !ok99977norm || !ok99977impHdr
      || !ok101267 || !ok101267near || !ok101267netOnly) {
    console.warn('[SAP NF] buildSapNfMap aggregation mismatches:', {
      ok97723, got97723: map['97723']?.valorNF,
      ok88888, got88888: map['88888']?.valorNF,
      ok99999, got99999: map['99999']?.valorNF,
      ok99635lines, got99635lines: map['99635']?.valorNF,
      ok99635doc, got99635doc: mapDocOnly['99635']?.valorNF,
      ok99635repeat, got99635repeat: mapBrutoRepeat['99635']?.valorNF,
      ok99635dupPos, got99635dupPos: mapDupPositional['99635']?.valorNF,
      ok99635lineDoc, got99635lineDoc: mapLineDocSameRow['99635']?.valorNF,
      ok99635inflate, got99635inflate: mapInflatedRepeat['99635']?.valorNF,
      ok99635brutoZfact, got99635brutoZfact: map99635brutoZfact['99635']?.valorNF,
      ok99635feeLines, got99635feeLines: map99635feeLines['99635']?.valorNF,
      ok99636bruto, got99636bruto: map99636bruto['99636']?.valorNF,
      ok99636lineDoc, got99636lineDoc: map99636lineDoc['99636']?.valorNF,
      ok2223, got2223: map2223['2223']?.valorNF,
      ok18591, got18591: map18591['18591']?.valorNF,
      ok99641, got99641: map99641['99641']?.valorNF,
      ok99641sameMat, got99641sameMat: map99641sameMat['99641']?.valorNF,
      ok99642, got99642: map99642['99642']?.valorNF,
      ok18596, got18596: map18596['18596']?.valorNF,
      ok18596multi, got18596multi: map18596multi['18596']?.valorNF,
      ok100796brl, got100796brl: map100796brl['100796']?.valorNF,
      ok100796lines, got100796lines: map100796lines['100796']?.valorNF,
      ok100796partial, got100796partial: map100796partial['100796']?.valorNF,
      ok100934ap, got100934ap: map100934['100934']?.valorNF,
      ok99977apOnce, got99977: map99977['99977']?.valorNF,
      ok99977brutoFace, got99977brutoFace: map99977brutoFaceRepeat['99977']?.valorNF,
      ok99977lineInflate, got99977lineInflate: map99977lineInflate['99977']?.valorNF,
      ok99977modestGap, got99977modestGap: map99977modestGap['99977']?.valorNF,
      ok99977near, got99977near: near99977,
      ok99977space, got99977space: map99977space['99977']?.valorNF,
      ok99977noDoc, got99977noDoc: map99977noDoc['99977']?.valorNF,
      ok99977norm, got99977norm: map99977normalize['99977']?.valorNF,
      ok99977impHdr, got99977impHdr: map99977imp['99977']?.valorNF,
      ok101267, got101267: map101267['101267']?.valorNF,
      ok101267near, got101267near: near101267,
      ok101267netOnly, got101267netOnly: map101267netOnly['101267']?.valorNF
    });
  } else {
    console.debug('[SAP NF] buildSapNfMap aggregation OK');
  }

  // Ghost faces: same bruto-once on ≥3 NFs without AP must be scrubbed (screenshot 12 Δ).
  const ghostMap = {
    '99977': { valorNF: 449796, lineSum: 11244900, docVal: 0, docFromInclImp: false },
    '100070': { valorNF: 449796, lineSum: 9895512, docVal: 0, docFromInclImp: false },
    '100476': { valorNF: 449796, lineSum: 6297144, docVal: 0, docFromInclImp: false },
    '97767': { valorNF: 194644.25, lineSum: 4671462, docVal: 0, docFromInclImp: false },
    '97830': { valorNF: 194644.25, lineSum: 4671462, docVal: 0, docFromInclImp: false },
    '101466': { valorNF: 194644.25, lineSum: 8564347, docVal: 0, docFromInclImp: false },
    '100934': { valorNF: 6673.5, lineSum: 6673.5, docVal: 6673.5, docFromInclImp: true }
  };
  const nGhost = scrubSharedGhostSapFaces(ghostMap);
  const okGhostScrub = nGhost >= 6
    && parseSapNum(ghostMap['99977'].valorNF) === 0
    && parseSapNum(ghostMap['97767'].valorNF) === 0
    && parseSapNum(ghostMap['100934'].valorNF) === 6673.5;
  if (!okGhostScrub) console.warn('[SAP NF] ghost scrub failed', nGhost, ghostMap['99977'], ghostMap['100934']);

  // FORCE refinalize from stored docVal when AP present but valorNF still bruto.
  const refMap = {
    '99977': { valorNF: 449796, lineSum: 449796, docVal: 442314.49, docFromInclImp: true }
  };
  const nRef = refinalizeSapMapFromStoredFields(refMap);
  const okRefinalize = nRef === 1 && Math.abs(parseSapNum(refMap['99977'].valorNF) - 442314.49) < 0.01;
  if (!okRefinalize) console.warn('[SAP NF] refinalize stored AP failed', refMap['99977']);
}
_debugBuildSapNfMapAggregation();

/** Relevant SAP vs Unilog NF value gap: abs diff > R$1 AND > 0.5% of the larger value. */
const SAP_VALOR_DIFF_MIN_ABS = 1.0;
const SAP_VALOR_DIFF_MIN_PCT = 0.005;

function isRelevantValorDiff(unilogVal, sapVal) {
  const u = num(unilogVal);
  const s = num(sapVal);
  const diff = Math.abs(s - u);
  if (diff <= SAP_VALOR_NEAR_EPS) return false;
  if (diff <= SAP_VALOR_DIFF_MIN_ABS) return false;
  const ref = Math.max(Math.abs(s), Math.abs(u), 1);
  return (diff / ref) > SAP_VALOR_DIFF_MIN_PCT;
}

function isSapLoaded() {
  return Object.keys(sapNfMap).length > 0;
}

function buildSapMissingReason(nf) {
  const parts = [
    'NF presente no Unilog mas não encontrada no export SAP (chave normalizada com série/zeros).'
  ];
  if (nf.nCte > 0 && nf.pago > 0) {
    parts.push(
      `Unilog emitiu ${nf.nCte} CT-e com ${fmtMoney(nf.pago)} pago — verificar se a NF existe no SAP ou se o CT-e está associado a NF errada/inexistente (ex.: transporte não realizado).`
    );
  } else if (nf.nCte > 0) {
    parts.push(
      `Unilog regista ${nf.nCte} CT-e sem valor pago — possível CT-e emitido sem transporte ou NF inexistente no SAP.`
    );
  } else {
    parts.push('Sem CT-e associado no Unilog — confirmar se a NF deveria existir no SAP.');
  }
  return parts.join(' ');
}

function applySapToNf(nf) {
  if (nf.valorUnilog === undefined || nf.valorUnilog === null) {
    nf.valorUnilog = num(nf.valorNF);
  }

  const sapLoaded = isSapLoaded();
  const sap = sapLoaded ? lookupSapEntry(nf.nf) : null;
  nf.sapFound = sapLoaded && !!sap;
  nf.sapMissing = sapLoaded && !sap;
  nf.sapAnomalyType = nf.sapMissing ? SAP_ANOMALY_TYPE : '';
  nf.sapAnomalyReason = nf.sapMissing ? buildSapMissingReason(nf) : '';

  if (!sap) {
    nf.valorSAP = null;
    nf.valorDiff = null;
    nf.sapValorMismatch = false;
    return nf;
  }

  if (sap.cliente) nf.cliente = sap.cliente;

  // Keep quinzenal Data NF before SAP overwrite (month bucketing fallback)
  if (nf.dtNF && !nf.dtNFQz && !nf.dtSAP) {
    nf.dtNFQz = persistDateValue(nf.dtNF) || nf.dtNF;
  }
  const sapDt = parseSapBrDate(sap.dtEmissao);
  if (sapDt) {
    nf.dtSAP = sapDt;
    nf.dtNF = sapDt;
    delete nf.mesRef;
    delete nf.dtRef;
  }

  const unilogVal = nf.valorUnilog;
  nf.valorSAP = pickSapValorNearUnilog(sap, unilogVal);
  let sapVal = num(nf.valorSAP);

  // No usable SAP face (ghost scrubbed / empty AP) — NF exists but valor unknown; not a Δ.
  if (!(sapVal > 0)) {
    nf.valorSAP = null;
    nf.valorDiff = null;
    nf.sapValorMismatch = false;
    return nf;
  }

  // Stale cloud map safety: valorNF was SUM of repeating face × N lines (e.g. 442k×25≈11M).
  // If SAP ≈ Unilog×N for integer N∈[3,100], recover once-per-NF face without Excel re-parse.
  if (unilogVal > 0 && sapVal > unilogVal * 2.5) {
    const ratio = sapVal / unilogVal;
    const nRep = Math.round(ratio);
    if (nRep >= 3 && nRep <= 100 && Math.abs(ratio - nRep) / nRep < 0.05) {
      const once = sapVal / nRep;
      console.warn('[SAP NF] deflated summed valor', nf.nf, sapVal, '÷', nRep, '→', once);
      sapVal = once;
      nf.valorSAP = once;
    }
  }

  if ((!nf.valorNF || nf.valorNF === 0) && sapVal) {
    nf.valorNF = sapVal;
    nf.esperado = nf.valorNF * CTE_PCT_TARGET;
    nf.diff = nf.pago - nf.esperado;
    nf.pct = nf.valorNF > 0 ? nf.pago / nf.valorNF : 0;
  }

  const rawDiff = sapVal - unilogVal;
  // Cent-level rounding (4.992,25 vs 4.992,26) is not a SAP↔Unilog difference.
  if (Math.abs(rawDiff) <= SAP_VALOR_NEAR_EPS) {
    nf.valorDiff = 0;
    nf.sapValorMismatch = false;
  } else {
    nf.valorDiff = rawDiff;
    nf.sapValorMismatch = isRelevantValorDiff(unilogVal, sapVal);
  }

  return nf;
}

function applySapToList(list) {
  return list.map(nf => {
    applySapToNf(nf);
    return enrichNF(nf);
  });
}

function processArrayBufferCte(arrayBuffer, fileName, opts = {}) {
  try {
    const wb = readWorkbookFromArrayBuffer(arrayBuffer);
    const { rows, sheetName, headers } = loadRowsFromWorkbook(wb);
    if (!rows.length) {
      const msg = 'Nenhuma linha de dados encontrada. Folhas: ' + wb.SheetNames.join(', ');
      if (!opts.silent) {
        setLoadbar(msg, false);
        fteToastError(msg);
      }
      return false;
    }
    if (!opts.silent) setLoadbar('A ler ' + fileName + ' ...', false);
    processRows(rows, fileName, sheetName, headers, { source: 'conciliacao' });
    return true;
  } catch (err) {
    console.error(err);
    const msg = formatXlsxReadError(err, 'cte');
    if (!opts.silent) {
      setLoadbar(msg, true);
      fteToastError(msg);
    }
    return false;
  }
}

/** Async CT-e parse — yields during processRowsAsync so Processar e Guardar stays responsive. */
async function processArrayBufferCteAsync(arrayBuffer, fileName, opts = {}) {
  try {
    await fteYield('A ler Excel CT-e…');
    const wb = readWorkbookFromArrayBuffer(arrayBuffer);
    const { rows, sheetName, headers } = loadRowsFromWorkbook(wb);
    if (!rows.length) {
      const msg = 'Nenhuma linha de dados encontrada. Folhas: ' + wb.SheetNames.join(', ');
      if (!opts.silent) {
        setLoadbar(msg, false);
        fteToastError(msg);
      }
      return false;
    }
    if (!opts.silent) setLoadbar('A ler ' + fileName + ' ...', false);
    await processRowsAsync(rows, fileName, sheetName, headers, { source: 'conciliacao', autosave: opts.autosave });
    return true;
  } catch (err) {
    console.error(err);
    const msg = formatXlsxReadError(err, 'cte');
    if (!opts.silent) {
      setLoadbar(msg, true);
      fteToastError(msg);
    }
    return false;
  }
}

/** Merge SAP NF maps by normalized NF key — next overwrites; prev keys absent from next are kept. */
function mergeSapNfMaps(prev, next) {
  if (!next || !Object.keys(next).length) return prev && typeof prev === 'object' ? prev : {};
  if (!prev || !Object.keys(prev).length) return { ...next };
  const out = { ...prev };
  Object.keys(next).forEach(k => {
    const inc = next[k];
    const old = out[k];
    if (!old) {
      out[k] = inc;
      return;
    }
    const incValor = parseSapNum(inc?.valorNF);
    const oldValor = parseSapNum(old?.valorNF);
    const scrubbed = inc && Object.prototype.hasOwnProperty.call(inc, '_ghostFaceScrubbed');
    out[k] = {
      nf: inc?.nf != null && String(inc.nf).trim() !== '' ? inc.nf : old.nf,
      cliente: (inc?.cliente && String(inc.cliente).trim()) || old.cliente || '',
      // Never wipe a known emissão date with a blank incoming parse (wide ZFACT layout bugs).
      dtEmissao: inc?.dtEmissao || old.dtEmissao || null,
      // Ghost scrub sets valorNF=0 on purpose — do not resurrect old bogus face.
      valorNF: scrubbed ? 0 : (incValor > 0 ? inc.valorNF : (oldValor > 0 ? old.valorNF : (inc?.valorNF ?? old.valorNF))),
      lineSum: num(inc?.lineSum) > 0 ? inc.lineSum : (num(old?.lineSum) || 0),
      docVal: num(inc?.docVal) > 0 ? inc.docVal : (num(old?.docVal) || 0),
      docFromInclImp: !!(inc?.docFromInclImp || old?.docFromInclImp)
    };
  });
  return out;
}

function trackSapMapSourceFile(fileName, nfCount) {
  const name = fileName || 'sap.xlsx';
  sapMapSourceFiles = [
    ...(sapMapSourceFiles || []).filter(f => f.fileName !== name),
    { fileName: name, nfCount: nfCount || 0, mergedAt: new Date().toISOString() }
  ];
}

function reviveSapNfMapEntry(e) {
  if (!e || typeof e !== 'object') return null;
  let dt = e.dtEmissao || null;
  if (dt && !(dt instanceof Date)) {
    dt = (typeof parseSapBrDate === 'function' ? parseSapBrDate(dt) : null)
      || (typeof dt === 'string' || typeof dt === 'number' ? new Date(dt) : null);
    if (dt && isNaN(dt.getTime())) dt = null;
  }
  return {
    nf: e.nf,
    cliente: e.cliente || '',
    dtEmissao: dt,
    valorNF: e.valorNF,
    lineSum: e.lineSum || 0,
    docVal: e.docVal || 0,
    docFromInclImp: !!e.docFromInclImp
  };
}

function reviveSapNfMap(raw) {
  const map = {};
  if (!raw || typeof raw !== 'object') return map;
  Object.keys(raw).forEach(k => {
    const entry = reviveSapNfMapEntry(raw[k]);
    if (entry) map[k] = entry;
  });
  return map;
}

/** Chunked revive — large accumulated SAP maps must not block the UI for seconds. */
async function reviveSapNfMapAsync(raw) {
  const map = {};
  if (!raw || typeof raw !== 'object') return map;
  const keys = Object.keys(raw);
  const CHUNK = 2500;
  for (let i = 0; i < keys.length; i += CHUNK) {
    const end = Math.min(i + CHUNK, keys.length);
    for (let j = i; j < end; j++) {
      const entry = reviveSapNfMapEntry(raw[keys[j]]);
      if (entry) map[keys[j]] = entry;
    }
    if (end < keys.length) await fteYield(`A indexar mapa SAP… ${end}/${keys.length}`);
  }
  return map;
}

function slimSapNfMapForPersist(map) {
  const out = {};
  Object.keys(map || {}).forEach(k => {
    const e = map[k];
    if (!e) return;
    let dt = e.dtEmissao || null;
    if (dt instanceof Date && !isNaN(dt.getTime())) dt = dateToYmd(dt) || dt.toISOString().slice(0, 10);
    else if (typeof dt === 'string' && /^\d{4}-\d{2}-\d{2}/.test(dt.trim())) dt = dt.trim().slice(0, 10);
    else if (dt != null && typeof dt !== 'string') {
      const parsed = parseSapBrDate(dt);
      dt = parsed ? (dateToYmd(parsed) || String(dt)) : String(dt);
    }
    out[k] = {
      nf: e.nf,
      cliente: e.cliente || '',
      dtEmissao: dt,
      valorNF: e.valorNF,
      lineSum: e.lineSum || 0,
      docVal: e.docVal || 0,
      docFromInclImp: !!e.docFromInclImp
    };
  });
  return out;
}

async function persistSapNfMap(map, opts = {}) {
  if (typeof upsertExcelBinary !== 'function') {
    if (!opts.silent) fteToastError('Persistência Excel indisponível — recarrega a página.');
    return false;
  }
  const m = map || sapNfMap;
  const n = Object.keys(m || {}).length;
  if (!n) {
    console.warn('[fretes] persist sap map skip — empty');
    return false;
  }
  const pack = {
    version: SAP_NF_MAP_PARSER_VERSION,
    updatedAt: new Date().toISOString(),
    files: sapMapSourceFiles || [],
    map: slimSapNfMapForPersist(m)
  };
  const json = JSON.stringify(pack);
  const buf = new TextEncoder().encode(json).buffer;
  const jsonBytes = buf.byteLength;
  if (jsonBytes > QZ_PERSIST_MAX_JSON_BYTES) {
    const msg = `Mapa SAP demasiado grande (${fmtByteSize(jsonBytes)}).`;
    console.error('[fretes] persist sap map too large', fteCompany(), jsonBytes);
    if (!opts.silent) fteToastError(msg);
    return false;
  }
  const slot = fteSapMapSlot();
  const fileName = `sap_nf_map_${n}.json`;
  console.log('[fretes] persist sap map', fteCompany(), slot, n, 'NFs', 'jsonBytes', jsonBytes);
  try {
    await upsertExcelBinary(slot, fileName, buf);
    console.log('[fretes] persist sap map ok', fteCompany(), slot, fileName);
    if (!opts.silent) fteToast(`Mapa SAP acumulado guardado (${n} NFs).`);
    return true;
  } catch (err) {
    console.error('[fretes] persist sap map', fteCompany(), slot, err);
    if (!opts.silent) {
      if (typeof isExcelFilesTableMissing === 'function' && isExcelFilesTableMissing(err)) {
        fteToastError('Tabela logistica_excel_files em falta no Supabase.');
      } else {
        fteToastError('Erro ao guardar mapa SAP: ' + (err.message || err));
      }
    }
    return false;
  }
}

function parseSapNfMapFromRec(rec) {
  if (!rec?.file_data || typeof base64ToArrayBuffer !== 'function') return null;
  try {
    const b64 = rec.file_data;
    rec.file_data = null;
    const buf = base64ToArrayBuffer(b64);
    const text = new TextDecoder().decode(buf);
    const parsed = JSON.parse(text);
    if (!parsed) return null;
    if (parsed.map && typeof parsed.map === 'object') {
      return {
        raw: parsed.map,
        files: Array.isArray(parsed.files) ? parsed.files : [],
        version: Number(parsed.version) || 1
      };
    }
    // Legacy: raw map object
    return { raw: parsed, files: [], version: 1 };
  } catch (err) {
    console.error('[fretes] parse sap map', err);
    return null;
  }
}

/**
 * Cloud JSON maps may still store SUM(repeating face)×lines from older parsers.
 * Deflate to once-per-NF when valor looks like unit×N (N=3..80). No Excel re-parse (OOM-safe).
 * Returns count of entries corrected.
 */
function sanitizeInflatedSapMapValors(map) {
  if (!map || typeof map !== 'object') return 0;
  let fixed = 0;
  const n99977 = parseSapNum(map['99977']?.valorNF);
  const fingerprint99977 = n99977 > 5_000_000 && Math.abs(n99977 - 11_244_900) < 1;
  Object.keys(map).forEach(k => {
    const e = map[k];
    if (!e) return;
    const v = parseSapNum(e.valorNF);
    const doc = parseSapNum(e.docVal);
    // Prefer stored AP / Val.total incl.imp. over a wrong inflated/deflated face.
    if (e.docFromInclImp && doc > 0 && Math.abs(v - doc) > SAP_VALOR_NEAR_EPS) {
      e.valorNF = doc;
      fixed++;
      return;
    }
    if (!(v > 500_000)) return;
    let unitOnce = 0;
    for (let n = 3; n <= 80; n++) {
      const unit = v / n;
      if (unit < 1000) break;
      if (unit > 2_000_000) continue;
      if (Math.abs(unit * n - v) < 0.05) {
        // Prefer whole-reais or cent-stable units (typical SAP exports).
        if (Math.abs(unit - Math.round(unit)) < 0.01 || Math.abs(unit * 100 - Math.round(unit * 100)) < 0.01) {
          unitOnce = unit;
          break;
        }
      }
    }
    if (unitOnce > 0) {
      // If we also have AP doc, prefer doc (449796 = 11244900/25 is a false face for NF 99977).
      if (doc > 0 && Math.abs(unitOnce - doc) > 1) {
        e.valorNF = doc;
        if (e.docFromInclImp == null) e.docFromInclImp = true;
      } else {
        e.valorNF = unitOnce;
      }
      fixed++;
    } else if (fingerprint99977 && k === '99977') {
      e.valorNF = doc > 0 ? doc : 0;
      fixed++;
    }
  });
  if (fixed) {
    console.warn('[fretes] sanitized', fixed, 'inflated SAP map valor(s) (SUM×lines → once)');
  }
  return fixed;
}

/**
 * FORCE re-finalize from stored map fields (no Excel): valorNF = docVal when AP/doc is reliable.
 * Covers stale packs where lineSum/bruto won before v8 FORCE, but docVal was persisted.
 */
function refinalizeSapMapFromStoredFields(map) {
  if (!map || typeof map !== 'object') return 0;
  let fixed = 0;
  Object.keys(map).forEach(k => {
    const e = map[k];
    if (!e) return;
    const doc = parseSapNum(e.docVal);
    const v = parseSapNum(e.valorNF);
    const ls = parseSapNum(e.lineSum);
    if (!(doc > 0)) return;
    const forceAp = !!e.docFromInclImp
      || (ls > 0 && ls > doc * 1.5)
      || (v > 0 && ls > 0 && sapValoresClose(v, ls) && Math.abs(v - doc) > SAP_VALOR_NEAR_EPS);
    if (!forceAp) return;
    if (Math.abs(v - doc) <= SAP_VALOR_NEAR_EPS) {
      if (!e.docFromInclImp) e.docFromInclImp = true;
      return;
    }
    e.valorNF = doc;
    e.docFromInclImp = true;
    fixed++;
  });
  if (fixed) console.warn('[fretes] refinalized', fixed, 'SAP map valor(s) from stored docVal/AP');
  return fixed;
}

/**
 * ZFACT rows without Val.total often leave the same bogus "face" (repeating bruto) on many NFs
 * (e.g. 449796 on 99977/100070/…, 194644.25 on 97767/97830/…). Those are not real NF totals —
 * clear valor so we do not invent SAP≠Unilog Δ. Keep cliente/dates (NF still "found").
 */
function scrubSharedGhostSapFaces(map, minShare = 3) {
  if (!map || typeof map !== 'object') return 0;
  const byFace = new Map();
  Object.keys(map).forEach(k => {
    const e = map[k];
    if (!e) return;
    const doc = parseSapNum(e.docVal);
    if (doc > 0 && e.docFromInclImp) return; // trusted AP
    const v = parseSapNum(e.valorNF);
    if (!(v > 1000)) return;
    const faceKey = String(Math.round(v * 100) / 100);
    if (!byFace.has(faceKey)) byFace.set(faceKey, []);
    byFace.get(faceKey).push(k);
  });
  let fixed = 0;
  byFace.forEach((keys, faceKey) => {
    if (keys.length < minShare) return;
    keys.forEach(k => {
      const e = map[k];
      if (!e) return;
      const doc = parseSapNum(e.docVal);
      if (doc > 0 && e.docFromInclImp) return;
      e.valorNF = 0;
      // Keep lineSum for debug; do not use as face without AP.
      e._ghostFaceScrubbed = parseSapNum(faceKey);
      fixed++;
    });
  });
  if (fixed) {
    console.warn('[fretes] scrubbed', fixed, 'shared ghost SAP face(s) (same valor on ≥' + minShare + ' NFs, no AP)');
  }
  return fixed;
}

/** Run all OOM-safe map repairs. Returns total entries changed. */
function repairSapNfMapInPlace(map) {
  if (!map || typeof map !== 'object') return 0;
  const n1 = sanitizeInflatedSapMapValors(map);
  const n2 = refinalizeSapMapFromStoredFields(map);
  const n3 = scrubSharedGhostSapFaces(map);
  return n1 + n2 + n3;
}

/** Restore accumulated SAP map from cloud JSON (does not clear keys absent from Excel). */
async function loadSavedSapNfMap(meta, opts = {}) {
  if (typeof fetchExcelFiles !== 'function') return false;
  try {
    const m = meta || await fetchExcelFiles([fteSapMapSlot()]);
    const rec = m[fteSapMapSlot()];
    if (!rec?.file_data) return false;
    await fteYield('A descodificar mapa SAP…');
    const parsed = parseSapNfMapFromRec(rec);
    if (!parsed?.raw || typeof parsed.raw !== 'object') {
      console.warn('[fretes] sap map parse empty', fteCompany(), rec.file_name);
      return false;
    }
    // Stale parser: still LOAD the map (dates/cliente). Do NOT auto-reparse multi-MB ZFACT
    // Excel on open (Chrome OOM). Repair from stored fields + scrub ghosts; persist silently.
    const stale = (parsed.version || 1) < SAP_NF_MAP_PARSER_VERSION;
    if (stale) {
      console.warn('[fretes] sap map stale parser v' + (parsed.version || 1)
        + ' < ' + SAP_NF_MAP_PARSER_VERSION
        + ' — repair from stored docVal/ghost scrub; persist mapa v' + SAP_NF_MAP_PARSER_VERSION);
    }
    const revived = await reviveSapNfMapAsync(parsed.raw);
    if (!Object.keys(revived).length) {
      console.warn('[fretes] sap map parse empty', fteCompany(), rec.file_name);
      return false;
    }
    const nFixed = repairSapNfMapInPlace(revived);
    fteSapMapNeedsReprocess = !!(stale && nFixed === 0);
    sapNfMap = mergeSapNfMaps(sapNfMap, revived);
    // Re-run repair on merged map (incoming may revive ghosts wiped in revived-only pass).
    const nMergedFix = repairSapNfMapInPlace(sapNfMap);
    const totalFixed = nFixed + nMergedFix;
    invalidateZfactMonthsCache();
    if (parsed.files?.length) {
      const byName = new Map((sapMapSourceFiles || []).map(f => [f.fileName, f]));
      parsed.files.forEach(f => {
        if (f?.fileName) byName.set(f.fileName, f);
      });
      sapMapSourceFiles = [...byName.values()];
    }
    if (totalFixed > 0 || stale) {
      // Auto-persist corrected map — user must not remember to Processar ZFACT.
      try {
        await persistSapNfMap(sapNfMap, { silent: true });
        fteSapMapNeedsReprocess = false;
        if (!opts.silent) {
          fteToast(totalFixed > 0
            ? `Mapa SAP corrigido na cloud (${totalFixed} valor(es)).`
            : 'Mapa SAP actualizado na cloud (parser v' + SAP_NF_MAP_PARSER_VERSION + ').');
        }
      } catch (persistErr) {
        console.warn('[fretes] persist repaired sap map', persistErr);
        fteSapMapNeedsReprocess = true;
        if (!opts.silent) {
          fteToast('Mapa SAP corrigido em memória — Processar ZFACT se o restore falhar a gravar.');
        }
      }
    } else if (fteSapMapNeedsReprocess && !opts.silent) {
      fteToast('Mapa SAP desatualizado — em Carregamento, clica Processar e Guardar (ZFACT) uma vez.');
    }
    console.log('[fretes] load sap map', fteCompany(), rec.file_name, 'mapSize', Object.keys(sapNfMap).length,
      'stale', stale, 'repaired', totalFixed);
    return true;
  } catch (err) {
    console.error('[fretes] load sap map', err);
    if (!opts.silent) fteToastError('Erro ao carregar mapa SAP acumulado.');
    return false;
  }
}

function applySapProcessResult(incoming, fileName, rowsLen, opts = {}) {
  repairSapNfMapInPlace(incoming);
  const nIncoming = Object.keys(incoming).length;
  const nBefore = Object.keys(sapNfMap).length;
  sapNfMap = mergeSapNfMaps(sapNfMap, incoming);
  repairSapNfMapInPlace(sapNfMap);
  invalidateZfactMonthsCache();
  fteSapMapNeedsReprocess = false;
  trackSapMapSourceFile(fileName, nIncoming);
  const nMapped = Object.keys(sapNfMap).length;
  const nWithValor = Object.values(sapNfMap).filter(e => parseSapNum(e.valorNF) > 0).length;
  if (!opts.silent) {
    console.log('[SAP] Map merge:', nBefore, '+', nIncoming, '→', nMapped, 'NFs (', rowsLen, 'linhas ZFACT), with valor:', nWithValor);
  }

  if (!opts.skipEnrich) {
    reEnrichAfterSapLoad();
    if (typeof window.refreshArmazemSapValidation === 'function') {
      try {
        window.refreshArmazemSapValidation();
      } catch (armErr) {
        console.warn('[SAP] refreshArmazemSapValidation', armErr);
      }
    }
  }

  const nMatched = currentSummary?.nSapMatched ?? countSapMatches(currentNFs);
  if (!opts.silent) {
    logSapKeyDebug(currentNFs);
    setSapLoadStatus(
      nMapped + ' NFs SAP acumuladas (' + nIncoming + ' neste ficheiro) · ' + nMatched + ' matched com Unilog',
      nMatched > 0 || !currentNFs.length
    );
    if (nMatched > 0) {
      fteToast('Dados SAP carregados — ' + nMatched + ' NF(s) cruzadas com Unilog.');
      const nMismatch = currentSummary?.nValorMismatch || 0;
      if (nMismatch) fteToast(nMismatch + ' NF(s) com diferença relevante de valor SAP vs Unilog.');
      const nMissing = currentSummary?.nSapMissing || 0;
      if (nMissing) fteToast(nMissing + ' NF(s) no Unilog sem correspondência no SAP — ver Anomalias.');
    } else if (currentNFs.length) {
      setSapLoadStatus(nMapped + ' NFs SAP acumuladas · 0 matched com Unilog', false);
      fteToastError('SAP carregado mas nenhuma NF cruzou com Unilog — verifica formato/chaves.');
    } else {
      fteToast('Dados SAP mesclados (' + nMapped + ' NFs acumuladas, ' + nIncoming + ' neste ficheiro).');
    }
  }
  renderLoadCoverageSummary();
  return true;
}

function processArrayBufferSap(arrayBuffer, fileName, opts = {}) {
  try {
    const wb = readSapWorkbookFromArrayBuffer(arrayBuffer);
    const { rows, sheetName } = loadSapRowsFromWorkbook(wb);
    if (!rows.length) {
      const msg = 'Nenhuma linha SAP encontrada na folha "' + sheetName + '".';
      if (!opts.silent) setSapLoadStatus(msg, false);
      return false;
    }
    return applySapProcessResult(buildSapNfMap(rows), fileName, rows.length, opts);
  } catch (err) {
    console.error(err);
    const msg = formatXlsxReadError(err, 'sap');
    if (!opts.silent) {
      setSapLoadStatus(msg, false);
      fteToastError(msg);
    }
    return false;
  }
}

/** Async ZFACT — yield before/after XLSX.read so the spinner paints; avoid sync restore on boot. */
async function processArrayBufferSapAsync(arrayBuffer, fileName, opts = {}) {
  try {
    await fteYield(opts.yieldLabel || 'A ler Excel ZFACT…');
    // Double yield so Chrome paints "A processar ZFACT…" before the heavy sync read.
    await fteYield(opts.yieldLabel || 'A parsear Excel ZFACT…');
    const wb = readSapWorkbookFromArrayBuffer(arrayBuffer);
    await fteYield('A mapear NFs SAP…');
    const { rows, sheetName } = loadSapRowsFromWorkbook(wb);
    if (!rows.length) {
      const msg = 'Nenhuma linha SAP encontrada na folha "' + sheetName + '".';
      if (!opts.silent) setSapLoadStatus(msg, false);
      return false;
    }
    await fteYield(`A agregar ${rows.length} linhas ZFACT…`);
    const incoming = buildSapNfMap(rows);
    await fteYield('A cruzar SAP…');
    return applySapProcessResult(incoming, fileName, rows.length, opts);
  } catch (err) {
    console.error(err);
    const msg = formatXlsxReadError(err, 'sap');
    if (!opts.silent) {
      setSapLoadStatus(msg, false);
      fteToastError(msg);
    }
    return false;
  }
}

function selectCteFile(file) {
  if (!file) return;
  fteCteFileName = file.name;
  const fn = $('cteFn');
  if (fn) fn.textContent = '⏳ A ler ' + file.name + '…';
  const reader = new FileReader();
  reader.onload = (e) => {
    fteCteBuffer = e.target.result;
    setCteZoneLoaded(file.name);
    fteToast('CT-e seleccionado — clica Processar e Guardar');
  };
  reader.onerror = () => {
    if (fn) fn.textContent = '';
    fteToastError('Erro ao ler ficheiro CT-e.');
  };
  reader.readAsArrayBuffer(file);
}

function selectSapFile(file) {
  if (!file) return;
  fteSapFileName = file.name;
  const fn = $('sapFn');
  if (fn) fn.textContent = '⏳ A ler ' + file.name + '…';
  const reader = new FileReader();
  reader.onload = (e) => {
    fteSapBuffer = e.target.result;
    setSapZoneLoaded(file.name);
    fteToast('SAP NF seleccionado');
  };
  reader.onerror = () => {
    if (fn) fn.textContent = '';
    fteToastError('Erro ao ler ficheiro SAP.');
  };
  reader.readAsArrayBuffer(file);
}

async function processQuinzenalPending() {
  if (!fteQzPendingFiles.length) return true;
  const results = await Promise.all(fteQzPendingFiles.map(f => processQuinzenalFile(f)));
  const ok = results.filter(r => r.ok);
  const fail = results.filter(r => !r.ok);
  if (!ok.length) {
    fteToastError('Nenhum ficheiro quinzenal processado.');
    return false;
  }
  const fileBinaries = { ...(quinzenalPack?.fileBinaries || {}) };
  ok.forEach(r => {
    if (r.arrayBuffer && typeof arrayBufferToBase64 === 'function') {
      fileBinaries[r.meta.fileName] = arrayBufferToBase64(r.arrayBuffer);
    }
  });
  const prevPack = quinzenalPack?.files?.length ? quinzenalPack : null;
  const built = buildQuinzenalPack(results, fileBinaries);
  quinzenalPack = (prevPack && built?.files?.length) ? mergeQuinzenalPacks(prevPack, built) : built;
  refreshQuinzenalCompare();
  fteQzPendingFiles = [];
  const qzInput = $('qzFileInput');
  if (qzInput) qzInput.value = '';
  syncQzUploadZone();
  if (fail.length) fteToast(`${ok.length} quinzenais OK, ${fail.length} com erro`);
  return true;
}

async function processAndSaveFretes() {
  const hasCte = !!(fteCteBuffer && fteCteFileName);
  const hasSap = !!(fteSapBuffer && fteSapFileName);
  const hasQzPending = fteQzPendingFiles.length > 0;
  const hasQzInMem = !!(quinzenalPack?.files?.length);
  if (!hasCte && !hasSap && !hasQzPending && !hasQzInMem) {
    fteToastError('Selecciona pelo menos um ficheiro (quinzenais ou SAP NF/ZFACT).');
    return;
  }
  fteSetProcessing(true, 'A processar…');

  const errors = [];
  let cteProcessed = false;
  let sapProcessed = false;
  let qzProcessed = false;

  try {
    // Process quinzenais in memory first (not blocked by CT-e cloud save)
    if (hasQzPending) {
      fteSetProcessing(true, 'A processar quinzenais…');
      const ok = await processQuinzenalPending();
      if (!ok) return;
      qzProcessed = true;
    }

    if (hasCte) {
      fteSetProcessing(true, 'A processar CT-e / SAP…');
      // Do NOT clear sapNfMap — CT-e reload must not wipe accumulated ZFACT history
      const ok = await processArrayBufferCteAsync(fteCteBuffer, fteCteFileName, { autosave: false });
      if (!ok) return;
      if (hasSap) {
        await fteYield('A processar SAP NF…');
        sapProcessed = !!(await processArrayBufferSapAsync(fteSapBuffer, fteSapFileName));
      }
      cteProcessed = true;
    } else if (hasSap) {
      // ZFACT / SAP NF alone — required for Fretes anomalies + Armazém NF 5,5% validation
      fteSetProcessing(true, 'A processar SAP NF (ZFACT)…');
      sapProcessed = !!(await processArrayBufferSapAsync(fteSapBuffer, fteSapFileName, {
        yieldLabel: 'A processar SAP NF (ZFACT)…'
      }));
      if (!sapProcessed) return;
    }

    // Sem Conciliacao CT-e×NF: montar Análise CT-e a partir dos quinzenais B2B
    let qzCteBuilt = false;
    if (!cteProcessed && !hasConciliacaoCte() && quinzenalPack?.b2bRows?.length) {
      fteSetProcessing(true, 'A montar Análise CT-e a partir dos quinzenais…');
      qzCteBuilt = await processCteAnalysisFromQuinzenalAsync({ switchTab: false });
      if (qzCteBuilt && isSapLoaded()) reEnrichAfterSapLoad();
    }

    fteSetProcessing(true, 'A guardar na cloud…');

    // Persist independently — quinzenais / SAP must not be skipped when CT-e save fails
    let qzSaved = true;
    let cteSaved = true;
    let sapSaved = true;

    if (quinzenalPack?.files?.length) {
      qzSaved = await persistQuinzenalPack(quinzenalPack, { silent: true });
      if (!qzSaved) errors.push('quinzenais');
    }

    if (cteProcessed) {
      cteSaved = await persistFretesFile(fteCteSlot(), fteCteFileName, fteCteBuffer);
      if (!cteSaved) errors.push('CT-e');
      if (cteSaved) _fteLoadedCompany = fteCompany();
    }

    if (hasSap && (sapProcessed || cteProcessed || qzCteBuilt)) {
      sapSaved = await persistFretesFile(fteSapSlot(), fteSapFileName, fteSapBuffer);
      if (!sapSaved) errors.push('SAP');
    }

    // Always persist accumulated SAP map when we have NFs (merge history ≠ last Excel alone)
    if (sapProcessed && isSapLoaded()) {
      const mapSaved = await persistSapNfMap(sapNfMap, { silent: true });
      if (!mapSaved) errors.push('mapa SAP');
    }

    syncQzUploadZone();
    updateQzFileNote();
    updateFretesFileStatus();
    renderLoadCoverageSummary();

    if (errors.length) {
      fteToastError('Processado em memória, mas falhou guardar na cloud: ' + errors.join(', ') + ' — verifica a consola.');
      return;
    }

    const parts = [];
    if (cteProcessed) parts.push('CT-e' + (sapProcessed ? ' + SAP' : ''));
    else if (sapProcessed) parts.push('SAP NF (ZFACT)');
    if (qzProcessed || quinzenalPack?.files?.length) {
      const c = qzFileCounts();
      parts.push(`quinzenais (${c.b2c} B2C · ${c.b2b} B2B)`);
    }
    if (qzCteBuilt) parts.push('Análise CT-e via QZ');
    fteToast('Processado e guardado na cloud: ' + parts.join(', ') + '.');
    const hasQz = !!(quinzenalPack?.files?.length);
    if (cteProcessed || qzCteBuilt) switchFteTab('analise-cte');
    else if (hasQz && !cteProcessed) switchFteTab(quinzenalPack?.b2cRows?.length ? 'analise-b2c' : 'cte-vs-qz');
  } catch (err) {
    console.error('[fretes] processAndSave', err);
    fteToastError('Erro: ' + (err.message || err));
  } finally {
    fteSetProcessing(false);
  }
}

function reEnrichAfterSapLoad() {
  if (lastCtePack) {
    processRows(lastCtePack.rows, lastCtePack.fileName, lastCtePack.sheetName, lastCtePack.headers, {
      source: lastCtePack.source || 'conciliacao',
      switchTab: false,
      autosave: false
    });
    return;
  }
  if (!currentNFs.length) return;
  currentNFs = applySapToList(currentNFs.map(nf => {
    delete nf.mesRef;
    delete nf.dtRef;
    return nf;
  }));
  currentSummary = computeSummary(currentNFs, currentSummary?.fileName || 'SAP enrich');
  renderAll();
}

function num(v) { return (v === null || v === undefined || v === '') ? 0 : Number(v); }

/** Sim/Não (and common variants) → boolean — used by Conciliacao + quinzenal B2B. */
function isDevolucaoFlag(v) {
  if (v === true || v === 1) return true;
  if (v === false || v === 0 || v == null || v === '') return false;
  const s = String(v).trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (!s) return false;
  if (s === 'nao' || s === 'n' || s === 'no' || s === 'false' || s === '0') return false;
  return s === 'sim' || s === 's' || s === 'yes' || s === 'true' || s === '1' || s === 'x';
}

const CTE_PCT_TARGET = 0.06;
const CTE_PCT_LOW = 0.059;
const CTE_PCT_HIGH = 0.061;
/** Recurring flat overcharge seen on high-value NFs (still ~6% because NF is large). */
const CTE_FIX_SURCHARGE = 165;
const CTE_FIX_SURCHARGE_LO = 160;
const CTE_FIX_SURCHARGE_HI = 170;

function fmtPp(v) { return (v * 100).toFixed(2) + ' p.p.'; }

function isFixed165Surcharge(diff) {
  const d = num(diff);
  return d >= CTE_FIX_SURCHARGE_LO && d <= CTE_FIX_SURCHARGE_HI;
}

function normCteKey(cte, idx) {
  const n = cte?.numCte;
  if (n !== null && n !== undefined && String(n).trim() !== '') return String(n).trim();
  return `_row_${idx}_${cte?.dtCte || ''}_${num(cte?.pago)}`;
}

function looksLikeCteDocNumber(n) {
  const s = String(n).trim();
  if (!s) return false;
  const digits = s.replace(/\D/g, '');
  if (digits.length >= 9) return true;
  if (/[./-]/.test(s) && digits.length >= 6) return true;
  return false;
}

/** Distinct CT-e document numbers per NF — ignores short integers (often col. E qtd mis-mapped as numCte). */
function countDistinctCtes(ctes) {
  if (!ctes?.length) return 0;
  const withNum = ctes.map(c => c?.numCte).filter(n => n !== null && n !== undefined && String(n).trim() !== '');
  const docNums = withNum.filter(looksLikeCteDocNumber);
  if (docNums.length > 0) {
    return new Set(docNums.map(n => String(n).trim())).size;
  }
  if (withNum.length === ctes.length) {
    const distinct = new Set(withNum.map(n => String(n).trim())).size;
    if (distinct === 1 && ctes.length > 1) return ctes.length;
    return distinct;
  }
  return ctes.length;
}

/** Best count: Excel col. E (qtd), distinct doc numbers, or row count. */
function resolveNfCteCount(ctes, qtdFromSource) {
  const fromRows = countDistinctCtes(ctes);
  const explicit = num(qtdFromSource);
  const rowCount = ctes?.length || 0;
  return Math.max(fromRows, explicit, rowCount);
}

function analyzeSingleCteMotivo(pago, esperado, pct) {
  const diff = pago - esperado;
  const diffPct = pct - CTE_PCT_TARGET;
  if (isFixed165Surcharge(diff)) {
    return {
      status: 'fix165',
      motivo: `Sobretaxa fixa ~R$ ${CTE_FIX_SURCHARGE}: pagou ${fmtMoney(pago)} (${fmtPct(pct)} da NF) vs esperado ${fmtMoney(esperado)} — excedente +${fmtMoney(diff)}. Padrão recorrente de cobrança adicional; questionar transportador/Unilog.`
    };
  }
  if (pct >= CTE_PCT_LOW && pct <= CTE_PCT_HIGH) {
    return { status: 'ok', motivo: `Conforme: ${fmtMoney(pago)} (${fmtPct(pct)} da NF), esperado ${fmtMoney(esperado)}.` };
  }
  if (pct < CTE_PCT_LOW) {
    return {
      status: 'low',
      motivo: `Abaixo de 6%: pagou ${fmtMoney(pago)} (${fmtPct(pct)}) vs esperado ${fmtMoney(esperado)} — desvio ${fmtMoney(diff)} (${fmtPp(diffPct)}). Questionar Unilog.`
    };
  }
  return {
    status: 'min',
    motivo: `Acima de 6,1%: pagou ${fmtMoney(pago)} (${fmtPct(pct)}) vs esperado ${fmtMoney(esperado)} — excedente +${fmtMoney(diff)} (+${fmtPp(diffPct)}). Provável tarifa mínima em NF de baixo valor; distinto da sobretaxa fixa R$ ${CTE_FIX_SURCHARGE}.`
  };
}

function analyzeCteEntry(cte, idx, valorNF, nCteTotal) {
  const pago = num(cte.pago);
  const esperado = valorNF * CTE_PCT_TARGET;
  const pct = valorNF > 0 ? pago / valorNF : 0;
  const diff = pago - esperado;
  const diffPct = pct - CTE_PCT_TARGET;
  let explicacao;

  if (nCteTotal === 1) {
    explicacao = analyzeSingleCteMotivo(pago, esperado, pct).motivo;
  } else if (cte.devolucao) {
    explicacao = `Devolução — CT-e ${cte.numCte || ('#' + (idx + 1))}: ${fmtMoney(pago)} (${fmtPct(pct)} da NF). Validar se frete de retorno está correcto.`;
  } else {
    const pctNote = pct > CTE_PCT_HIGH ? ' — possível cobrança sobre valor total da NF' : '';
    explicacao = `CT-e ${cte.numCte || ('#' + (idx + 1))}: ${fmtMoney(pago)} = ${fmtPct(pct)} da NF${pctNote}. Desvio vs 6%: ${fmtMoney(diff)} (${fmtPp(diffPct)}). Validar individualmente.`;
  }

  return {
    ...cte,
    pago,
    pct,
    esperado,
    diff,
    diffPct,
    explicacao,
    cteKey: cte.cteKey || normCteKey(cte, idx),
    validacao: cte.validacao || 'pendente'
  };
}

/**
 * Quinzenal B2B often repeats the NF-level "Total fatura" on every CT-e row.
 * Summing identical pagos × N (e.g. 91×R$26.703) explodes % pago to 500%+.
 * Take once when all positive pagos match; or when every row is ~6% of the full NF (≥3 CT-e).
 */
function aggregateNfPago(ctes, valorNF) {
  const pagos = (ctes || []).map(c => num(c.pago));
  const positive = pagos.filter(p => p > 0);
  if (!positive.length) return { pago: 0, coalescedOnce: false };
  if (positive.length === 1) return { pago: positive[0], coalescedOnce: false };

  const first = positive[0];
  const eps = Math.max(0.05, Math.abs(first) * 0.002);
  const allClose = positive.every(p => Math.abs(p - first) <= eps);
  if (allClose) return { pago: first, coalescedOnce: true };

  const v = num(valorNF);
  if (v > 0 && positive.length >= 3) {
    const allNearFullPct = positive.every(p => {
      const pct = p / v;
      return pct >= (CTE_PCT_LOW - 0.005) && pct <= (CTE_PCT_HIGH + 0.02);
    });
    if (allNearFullPct) return { pago: Math.max(...positive), coalescedOnce: true };
  }

  return { pago: positive.reduce((s, p) => s + p, 0), coalescedOnce: false };
}

(function _debugAggregateNfPago() {
  const pagoOnce = aggregateNfPago(
    Array.from({ length: 91 }, () => ({ pago: 26703.87 })),
    442314.49
  );
  const okPagoOnce = pagoOnce.coalescedOnce && Math.abs(pagoOnce.pago - 26703.87) < 0.01;
  const pagoSum = aggregateNfPago(
    [{ pago: 100 }, { pago: 200 }, { pago: 50 }],
    10000
  );
  const okSum = !pagoSum.coalescedOnce && Math.abs(pagoSum.pago - 350) < 0.01;
  if (!okPagoOnce || !okSum) console.warn('[SAP NF] aggregateNfPago failed', { pagoOnce, pagoSum });
  else console.debug('[SAP NF] aggregateNfPago OK');
})();

function buildNfRecord(g) {
  const ctesRaw = g.ctes || [];
  const nCte = resolveNfCteCount(ctesRaw, g.qtdCteFromSource);
  const ctes = ctesRaw.map((c, i) => analyzeCteEntry(c, i, g.valorNF, nCte));
  const { pago, coalescedOnce } = aggregateNfPago(ctes, g.valorNF);
  const esperado = g.valorNF * CTE_PCT_TARGET;
  const diff = pago - esperado;
  const pct = g.valorNF > 0 ? pago / g.valorNF : 0;

  let status, motivo;
  // Identical NF-level frete repeated on every CT-e row → classify like a single logical charge.
  if (nCte === 1 || (coalescedOnce && !g.temDevolucao)) {
    ({ status, motivo } = analyzeSingleCteMotivo(pago, esperado, pct));
    if (coalescedOnce && nCte > 1) {
      motivo = `${nCte} CT-e com total fatura repetido nas linhas quinzenais — pago ${fmtMoney(pago)} (1×, não ×${nCte}). `
        + motivo;
    }
  } else {
    const nDev = ctes.filter(c => c.devolucao).length;
    const pendente = ctes.filter(c => c.validacao === 'pendente').length;
    if (g.temDevolucao) {
      status = 'dev';
      motivo = `${nCte} CT-e (${nDev} devolução). Total ${fmtMoney(pago)} (${fmtPct(pct)} da NF) vs esperado ${fmtMoney(esperado)} — desvio ${fmtMoney(diff)}. Validar cada CT-e abaixo${pendente ? ` (${pendente} pendentes)` : ''}.`;
    } else {
      status = 'flag';
      motivo = `${nCte} CT-e sem devolução. Total ${fmtMoney(pago)} (${fmtPct(pct)} da NF) — desvio ${fmtMoney(diff)}. Cada CT-e precisa de validação individual${pendente ? ` (${pendente} pendentes)` : ''}.`;
    }
  }

  return {
    nf: g.nf, cliente: g.cliente || '', transportador: g.transportador, modalidade: g.modalidade,
    valorNF: g.valorNF, valorUnilog: g.valorNF, valorSAP: null, valorDiff: null,
    sapFound: false, sapMissing: false, sapAnomalyType: '', sapAnomalyReason: '',
    sapValorMismatch: false,
    nCte, pago, esperado, diff, pct, status, motivo, ctes,
    temDevolucao: g.temDevolucao, dtNF: g.dtNF,
    dtNFQz: g.dtNFQz || g.dtNF || null,
    qtdCteFromSource: g.qtdCteFromSource || 0,
    pagoCoalescedOnce: !!coalescedOnce,
    mesKey: g.mesKey || g.qzMesKey || null,
    qzMesKey: g.qzMesKey || g.mesKey || null
  };
}

function setCteValidacao(nfStr, cteKey, value) {
  const nf = currentNFs.find(x => String(x.nf) === String(nfStr));
  if (!nf) return;
  const idx = nf.ctes.findIndex(c => c.cteKey === cteKey);
  if (idx < 0) return;
  nf.ctes[idx].validacao = value;
  const rebuilt = buildNfRecord({
    nf: nf.nf, cliente: nf.cliente, transportador: nf.transportador, modalidade: nf.modalidade,
    valorNF: nf.valorNF, ctes: nf.ctes, temDevolucao: nf.temDevolucao, dtNF: nf.dtNF,
    dtNFQz: nf.dtNFQz, qtdCteFromSource: nf.qtdCteFromSource,
    mesKey: nf.mesKey, qzMesKey: nf.qzMesKey
  });
  Object.assign(nf, rebuilt);
  delete nf.mesRef;
  delete nf.dtRef;
  enrichNF(nf);
  renderTable();
}

function reopenNfDetail(nfStr) {
  const nfRec = currentNFs.find(n => String(n.nf) === String(nfStr));
  if (!nfRec) return;
  document.querySelectorAll('.detail-row').forEach(d => d.remove());
  const rows = $('nfTableBody').querySelectorAll('tr.row-clickable');
  for (const row of rows) {
    if (String(row.cells[0]?.textContent || '').trim() === String(nfStr)) {
      toggleDetail(row, nfRec);
      break;
    }
  }
}

function processRows(rows, fileName, sheetName, headers, opts = {}) {
  lastCtePack = { rows, fileName, sheetName, headers, source: opts.source || 'conciliacao' };
  cteAnalysisSource = opts.source || 'conciliacao';

  setLoadbar('A processar ' + rows.length + ' linhas (folha "' + sheetName + '")...');

  const byNF = {};
  rows.forEach(r => {
    const nf = r.nf;
    if (nf === null || nf === undefined || String(nf).trim() === '') return;
    const nfKey = normNFKey(nf);
    if (!byNF[nfKey]) byNF[nfKey] = {
      nf: String(nf).trim(), valorNF: num(r.valorNF), transportador: r.transportador || '-',
      modalidade: r.modalidade || '-', ctes: [], temDevolucao: false,
      dtNF: r.dtNF || r.dtNFQz || null,
      dtNFQz: r.dtNFQz || r.dtNF || null,
      cliente: '', qtdCteFromSource: 0,
      mesKey: r.mesKey || r.qzMesKey || null,
      qzMesKey: r.qzMesKey || r.mesKey || null
    };
    const g = byNF[nfKey];
    const qtdRow = num(r.qtdCte);
    if (qtdRow > g.qtdCteFromSource) g.qtdCteFromSource = qtdRow;
    if (num(r.valorNF) > g.valorNF) g.valorNF = num(r.valorNF);
    if (r.transportador && g.transportador === '-') g.transportador = r.transportador;
    if (r.modalidade && g.modalidade === '-') g.modalidade = r.modalidade;
    const rowDtNF = r.dtNFQz || r.dtNF;
    if (rowDtNF) {
      g.dtNFQz = pickEarliestDateValue(g.dtNFQz, rowDtNF);
      g.dtNF = pickEarliestDateValue(g.dtNF, rowDtNF);
    }
    if (!g.mesKey && (r.mesKey || r.qzMesKey)) {
      g.mesKey = r.mesKey || r.qzMesKey;
      g.qzMesKey = r.qzMesKey || r.mesKey;
    }
    const isDev = isDevolucaoFlag(r.devolucao);
    g.ctes.push({
      numCte: r.numCte, dtCte: r.dtCte, pago: num(r.pago),
      devolucao: isDev,
      tipoOp: r.tipoOp, peso: num(r.peso)
    });
    if (isDev) g.temDevolucao = true;
  });

  let nfList = Object.values(byNF).map(buildNfRecord);

  nfList = applySapToList(nfList);
  nfList.sort((a, b) => b.diff - a.diff);
  currentNFs = nfList;
  currentSummary = computeSummary(nfList, fileName);
  tableSort = { col: 'diff', dir: -1 };
  selectedMonth = '';

  if (!nfList.length) {
    const cols = (headers && headers.length) ? headers.join(', ') : '(não detetadas)';
    setLoadbar('0 notas fiscais encontradas na folha "' + sheetName + '". Colunas: ' + cols, true);
    switchFteTab('carregamento');
    fteToast('Não foi possível ler notas fiscais — verifica se o ficheiro tem a coluna "Nota Fiscal".');
    return;
  }

  renderAll();
  showResultsView(opts);
  if (!_fteSkipAutosave && opts.autosave !== false) autosaveToCloud();
}

/** Chunked processRows — yields every ~800 lines so the main thread stays responsive. */
async function processRowsAsync(rows, fileName, sheetName, headers, opts = {}) {
  lastCtePack = { rows, fileName, sheetName, headers, source: opts.source || 'conciliacao' };
  cteAnalysisSource = opts.source || 'conciliacao';

  setLoadbar('A processar ' + rows.length + ' linhas (folha "' + sheetName + '")...');
  fteSetProcessing(true, 'A processar ' + rows.length + ' linhas…');

  const byNF = {};
  const CHUNK = 800;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const end = Math.min(i + CHUNK, rows.length);
    for (let j = i; j < end; j++) {
      const r = rows[j];
      const nf = r.nf;
      if (nf === null || nf === undefined || String(nf).trim() === '') continue;
      const nfKey = normNFKey(nf);
      if (!byNF[nfKey]) byNF[nfKey] = {
        nf: String(nf).trim(), valorNF: num(r.valorNF), transportador: r.transportador || '-',
        modalidade: r.modalidade || '-', ctes: [], temDevolucao: false,
        dtNF: r.dtNF || r.dtNFQz || null,
        dtNFQz: r.dtNFQz || r.dtNF || null,
        cliente: '', qtdCteFromSource: 0,
        mesKey: r.mesKey || r.qzMesKey || null,
        qzMesKey: r.qzMesKey || r.mesKey || null
      };
      const g = byNF[nfKey];
      const qtdRow = num(r.qtdCte);
      if (qtdRow > g.qtdCteFromSource) g.qtdCteFromSource = qtdRow;
      if (num(r.valorNF) > g.valorNF) g.valorNF = num(r.valorNF);
      if (r.transportador && g.transportador === '-') g.transportador = r.transportador;
      if (r.modalidade && g.modalidade === '-') g.modalidade = r.modalidade;
      const rowDtNF = r.dtNFQz || r.dtNF;
      if (rowDtNF) {
        g.dtNFQz = pickEarliestDateValue(g.dtNFQz, rowDtNF);
        g.dtNF = pickEarliestDateValue(g.dtNF, rowDtNF);
      }
      if (!g.mesKey && (r.mesKey || r.qzMesKey)) {
        g.mesKey = r.mesKey || r.qzMesKey;
        g.qzMesKey = r.qzMesKey || r.mesKey;
      }
      const isDev = isDevolucaoFlag(r.devolucao);
      g.ctes.push({
        numCte: r.numCte, dtCte: r.dtCte, pago: num(r.pago),
        devolucao: isDev,
        tipoOp: r.tipoOp, peso: num(r.peso)
      });
      if (isDev) g.temDevolucao = true;
    }
    if (end < rows.length) {
      await fteYield(`A processar ${end}/${rows.length} linhas…`);
    }
  }

  await fteYield('A classificar NFs…');
  let nfList = Object.values(byNF).map(buildNfRecord);
  nfList = applySapToList(nfList);
  nfList.sort((a, b) => b.diff - a.diff);
  currentNFs = nfList;
  currentSummary = computeSummary(nfList, fileName);
  tableSort = { col: 'diff', dir: -1 };
  selectedMonth = '';

  if (!nfList.length) {
    const cols = (headers && headers.length) ? headers.join(', ') : '(não detetadas)';
    setLoadbar('0 notas fiscais encontradas na folha "' + sheetName + '". Colunas: ' + cols, true);
    switchFteTab('carregamento');
    fteToast('Não foi possível ler notas fiscais — verifica se o ficheiro tem a coluna "Nota Fiscal".');
    return;
  }

  await fteYield('A desenhar Análise CT-e…');
  renderAll();
  showResultsView(opts);
  if (!_fteSkipAutosave && opts.autosave !== false) autosaveToCloud();
}

function computeSummary(list, fileName) {
  const totalPago = list.reduce((s, x) => s + x.pago, 0);
  const totalValorNF = list.reduce((s, x) => s + x.valorNF, 0);
  const totalEsperado = list.reduce((s, x) => s + x.esperado, 0);
  const excesso = totalPago - totalEsperado;
  const dates = list.flatMap(x => x.ctes.map(c => parseCteDateValue(c.dtCte))).filter(Boolean);
  const periodoInicio = dates.length ? new Date(Math.min(...dates)) : null;
  const periodoFim = dates.length ? new Date(Math.max(...dates)) : null;

  const byStatus = { ok: 0, fix165: 0, min: 0, dev: 0, flag: 0, low: 0 };
  const excessoByStatus = { ok: 0, fix165: 0, min: 0, dev: 0, flag: 0, low: 0 };
  const pagoByStatus = { ok: 0, fix165: 0, min: 0, dev: 0, flag: 0, low: 0 };
  let nUnicoCte = 0, nMultiplosCte = 0;
  let excessoPositivo = 0, deficit = 0;

  let nValorMismatch = 0, nSapMatched = 0, nSapMissing = 0;
  list.forEach(x => {
    byStatus[x.status]++;
    excessoByStatus[x.status] += x.diff;
    pagoByStatus[x.status] += x.pago;
    if (x.nCte === 1) nUnicoCte++; else nMultiplosCte++;
    if (x.diff > 0) excessoPositivo += x.diff;
    else if (x.diff < 0) deficit += x.diff;
    if (x.sapFound) nSapMatched++;
    if (x.sapMissing) nSapMissing++;
    if (x.sapValorMismatch) nValorMismatch++;
  });

  return {
    fileName, totalNF: list.length, totalPago, totalValorNF, totalEsperado, excesso,
    pctPagoSobreNF: totalValorNF ? totalPago / totalValorNF : 0,
    excessoPositivo, deficit, excessoPct: totalEsperado ? excesso / totalEsperado : 0,
    periodoInicio, periodoFim, byStatus, excessoByStatus, pagoByStatus,
    nUnicoCte, nMultiplosCte, nValorMismatch, nSapMatched, nSapMissing,
    sapLoaded: isSapLoaded()
  };
}

function getSapAnomalies(list) {
  if (!isSapLoaded()) return [];
  return list.filter(x => x.sapMissing);
}

function anomalyExportRow(x) {
  enrichNF(x);
  const meta = statusMeta[x.status] || {};
  return {
    'Nota Fiscal': x.nf,
    'Cliente (Unilog)': x.cliente || '',
    'Data NF (Unilog)': x.dtNF ? new Date(x.dtNF) : '',
    'Mês': fmtMesLabel(x.mesRef),
    'Valor NF (Unilog)': x.valorUnilog ?? x.valorNF,
    'Transportador': x.transportador,
    'Qtd CT-e': x.nCte,
    'Valor pago': x.pago,
    'Estado frete': meta.label || x.status,
    'Tipo anomalia': x.sapAnomalyType || SAP_ANOMALY_TYPE,
    'Explicação': x.sapAnomalyReason || buildSapMissingReason(x)
  };
}

function fmtDate(d) {
  if (!d) return '-';
  const dt = d instanceof Date ? d : (parseSapBrDate(d) || new Date(d));
  if (isNaN(dt) || !isPlausibleSapDate(dt)) return '-';
  return dt.toLocaleDateString('pt-BR');
}

function fmtValorDiff(v, mismatch) {
  if (v === null || v === undefined) return '-';
  if (Math.abs(num(v)) <= SAP_VALOR_NEAR_EPS) return fmtMoney(0);
  const color = mismatch ? '#b3261e' : 'inherit';
  const sign = v > 0 ? '+' : '';
  return `<span style="color:${color}">${sign}${fmtMoney(v)}</span>`;
}

function nfColCount() { return NF_COLUMNS.length; }

function monthKey(d) {
  const dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt)) return 'sem-data';
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}`;
}

function fmtMesLabel(key) {
  if (key === 'sem-data') return 'Sem data';
  const [y, m] = key.split('-');
  return `${MESES_PT[Number(m) - 1]}/${y}`;
}

function mesRefSortKey(mesRef) {
  if (mesRef === 'sem-data' || mesRef === 'sem-mes') return '9999-99';
  return mesRef || '9999-99';
}

function compareMesRef(a, b) {
  return mesRefSortKey(a).localeCompare(mesRefSortKey(b));
}

function sortMonthlyRows(rows, sortState) {
  const colDef = MONTH_COLUMNS.find(c => c.key === sortState.col) || MONTH_COLUMNS[0];
  if (colDef.key === 'mesRef') {
    const dir = sortState.dir;
    return [...rows].sort((a, b) => compareMesRef(a.mesRef, b.mesRef) * dir);
  }
  return sortRows(rows, MONTH_COLUMNS, sortState);
}

function aggregateMonthTotals(rows) {
  const t = {
    totalNF: 0, totalValorNF: 0, totalPago: 0, totalEsperado: 0, excesso: 0,
    qtd_ok: 0, qtd_fix165: 0, qtd_min: 0, qtd_dev: 0, qtd_flag: 0, qtd_low: 0
  };
  rows.forEach(m => {
    t.totalNF += m.totalNF;
    t.totalValorNF += m.totalValorNF;
    t.totalPago += m.totalPago;
    t.totalEsperado += m.totalEsperado;
    t.excesso += m.excesso;
    t.qtd_ok += m.qtd_ok;
    t.qtd_fix165 += m.qtd_fix165;
    t.qtd_min += m.qtd_min;
    t.qtd_dev += m.qtd_dev;
    t.qtd_flag += m.qtd_flag;
    t.qtd_low += m.qtd_low;
  });
  t.pctPagoSobreNF = t.totalValorNF ? t.totalPago / t.totalValorNF : 0;
  t.excessoPct = t.totalEsperado ? t.excesso / t.totalEsperado : 0;
  return t;
}

function buildMonthlyDisplayRows(sorted) {
  const out = [];
  const dated = sorted.filter(m => m.mesRef !== 'sem-data');
  const noData = sorted.filter(m => m.mesRef === 'sem-data');
  let lastYear = null;
  let yearGroup = [];

  dated.forEach(m => {
    const year = m.mesRef.slice(0, 4);
    if (lastYear && year !== lastYear && yearGroup.length) {
      out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...aggregateMonthTotals(yearGroup) });
      yearGroup = [];
    }
    if (year !== lastYear) {
      out.push({ type: 'year', label: year, year });
      lastYear = year;
    }
    out.push({ type: 'month', ...m });
    yearGroup.push(m);
  });
  if (yearGroup.length && lastYear) {
    out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...aggregateMonthTotals(yearGroup) });
  }
  noData.forEach(m => out.push({ type: 'month', ...m }));
  if (sorted.length) {
    out.push({ type: 'total', label: 'Total geral', ...aggregateMonthTotals(sorted) });
  }
  return out;
}


function monthExportRowFull(m, label) {
  return {
    'Mês': label || m.mesLabel,
    'NFs': m.totalNF,
    'Faturação': m.totalValorNF,
    'Pago': m.totalPago,
    'Esperado (6%)': m.totalEsperado,
    'Excesso': m.excesso,
    '% pago/NF': m.pctPagoSobreNF,
    'Excesso % vs esperado': m.excessoPct,
    'Conformes': m.qtd_ok,
    'Sobretaxa R$165': m.qtd_fix165,
    'Tarifa mínima': m.qtd_min,
    'Com devolução': m.qtd_dev,
    'A investigar': m.qtd_flag,
    'Abaixo': m.qtd_low
  };
}

function buildMonthlyExportRowsFull(rows) {
  const sorted = [...rows].sort((a, b) => compareMesRef(a.mesRef, b.mesRef));
  const out = [];
  const dated = sorted.filter(m => m.mesRef !== 'sem-data');
  const noData = sorted.filter(m => m.mesRef === 'sem-data');
  let lastYear = null;
  let yearGroup = [];

  dated.forEach(m => {
    const year = m.mesRef.slice(0, 4);
    if (lastYear && year !== lastYear && yearGroup.length) {
      out.push(monthExportRowFull(aggregateMonthTotals(yearGroup), `Total ${lastYear}`));
      yearGroup = [];
    }
    out.push(monthExportRowFull(m));
    yearGroup.push(m);
    lastYear = year;
  });
  if (yearGroup.length && lastYear) {
    out.push(monthExportRowFull(aggregateMonthTotals(yearGroup), `Total ${lastYear}`));
  }
  noData.forEach(m => out.push(monthExportRowFull(m)));
  if (sorted.length) out.push(monthExportRowFull(aggregateMonthTotals(sorted), 'Total geral'));
  return out;
}

function buildMonthlyExportRows(rows) {
  return buildMonthlyExportRowsFull(rows).map(r => ({
    'Mês': r['Mês'],
    'NFs': r['NFs'],
    'Faturação': r['Faturação'],
    'Pago': r['Pago'],
    'Esperado (6%)': r['Esperado (6%)'],
    'Excesso': r['Excesso'],
    '% pago/NF': r['% pago/NF'],
    'Sobretaxa R$165': r['Sobretaxa R$165'],
    'Tarifa mínima': r['Tarifa mínima'],
    'A investigar': r['A investigar'],
    'Conformes': r['Conformes']
  }));
}

function parseCteDateValue(v) {
  if (v == null || v === '') return null;
  const d = parseSapBrDate(v);
  if (d && !isNaN(d) && isPlausibleSapDate(d)) return d;
  if (v instanceof Date && !isNaN(v) && isPlausibleSapDate(v)) {
    const local = new Date(v.getFullYear(), v.getMonth(), v.getDate());
    local.setHours(0, 0, 0, 0);
    return local;
  }
  return null;
}

/** Earliest plausible date among candidates (Data NF / Dt CTE aggregation). */
function pickEarliestDateValue(...vals) {
  let bestRaw = null;
  let bestT = Infinity;
  for (const v of vals) {
    if (v == null || v === '') continue;
    const d = parseCteDateValue(v);
    if (!d) {
      if (bestRaw == null) bestRaw = persistDateValue(v) || v;
      continue;
    }
    const t = d.getTime();
    if (t < bestT) {
      bestT = t;
      bestRaw = persistDateValue(v) || v;
    }
  }
  return bestRaw;
}

/**
 * Month bucket for Por mês:
 * 1) SAP emissão (dtSAP / ZFACT) when present and valid
 * 2) else Data NF from quinzenais (invoice emission)
 * 3) else min Dt CTE
 * 4) else mesKey from the quinzenal file name / pack
 * Never stick on stale 'sem-data'.
 */
function enrichNF(nf) {
  if (nf.mesRef && nf.mesRef !== 'sem-data' && nf.dtRef) return nf;
  // 1. SAP emissão
  let d = parseCteDateValue(nf.dtSAP);
  // 2. Data NF (quinzenal emissão da fatura) — preferred QZ date for bucketing
  if (!d) d = parseCteDateValue(nf.dtNFQz);
  // 3. dtNF may still hold QZ Data NF (no SAP overwrite) or legacy SAP date
  if (!d) d = parseCteDateValue(nf.dtNF);
  // 4. min Dt CTE across CT-e lines
  if (!d && nf.ctes?.length) {
    const dates = nf.ctes.map(c => parseCteDateValue(c.dtCte)).filter(Boolean);
    if (dates.length) d = new Date(Math.min(...dates.map(x => x.getTime())));
  }
  if (d && !isNaN(d) && isPlausibleSapDate(d)) {
    nf.dtRef = d;
    nf.mesRef = monthKey(d);
    return nf;
  }
  // Fallback: quinzenal file month (e.g. "1ªQ Julho 2026") when dates blank
  const qzMes = nf.mesKey || nf.qzMesKey || null;
  if (qzMes && qzMes !== 'sem-mes' && qzMes !== 'sem-data' && /^\d{4}-\d{2}$/.test(String(qzMes))) {
    nf.dtRef = null;
    nf.mesRef = qzMes;
    return nf;
  }
  nf.dtRef = null;
  nf.mesRef = 'sem-data';
  return nf;
}

function computeMonthly(list) {
  const byMes = {};
  list.forEach(x => {
    enrichNF(x);
    const m = x.mesRef || 'sem-data';
    if (!byMes[m]) byMes[m] = {
      mesRef: m, mesLabel: fmtMesLabel(m), items: [],
      totalNF: 0, totalValorNF: 0, totalPago: 0, totalEsperado: 0, excesso: 0,
      qtd_ok: 0, qtd_fix165: 0, qtd_min: 0, qtd_dev: 0, qtd_flag: 0, qtd_low: 0
    };
    const b = byMes[m];
    b.items.push(x);
    b.totalNF++;
    b.totalValorNF += x.valorNF;
    b.totalPago += x.pago;
    b.totalEsperado += x.esperado;
    b.excesso += x.diff;
    b['qtd_' + x.status]++;
  });
  return Object.values(byMes).map(b => ({
    ...b,
    pctPagoSobreNF: b.totalValorNF ? b.totalPago / b.totalValorNF : 0,
    excessoPct: b.totalEsperado ? b.excesso / b.totalEsperado : 0
  })).sort((a, b) => compareMesRef(a.mesRef, b.mesRef));
}

function sortValue(row, col, type) {
  if (type === 'status') return (statusMeta[row.status] || {}).label || row.status || '';
  if (type === 'mesRef') return mesRefSortKey(row.mesRef);
  if (type === 'number') return num(row[col]);
  if (type === 'date') {
    const d = row[col];
    if (!d) return 0;
    const dt = d instanceof Date ? d : new Date(d);
    return isNaN(dt) ? 0 : dt.getTime();
  }
  return String(row[col] ?? '').toLowerCase();
}

function sortRows(list, columns, sortState) {
  const colDef = columns.find(c => c.key === sortState.col) || columns[0];
  const dir = sortState.dir;
  return [...list].sort((a, b) => {
    const va = sortValue(a, colDef.key, colDef.type);
    const vb = sortValue(b, colDef.key, colDef.type);
    if (va < vb) return -1 * dir;
    if (va > vb) return 1 * dir;
    return 0;
  });
}

function fteScheduleTableSort() {
  if (typeof scheduleTableSort === 'function') scheduleTableSort();
  else if (typeof enableTableSort === 'function') {
    document.querySelectorAll('#page-fretes .tbl-wrap table').forEach(t => {
      if (!t.dataset.managedSort) enableTableSort(t);
    });
  }
}

function fteEnableDomSort(bodyEl) {
  const table = bodyEl?.closest?.('table');
  if (table && typeof enableTableSort === 'function') enableTableSort(table);
}

function renderSortableHead(containerId, columns, sortState, onSort) {
  const tr = $(containerId);
  if (window.TableSort) {
    window.TableSort.renderDataHead(tr, columns, sortState, onSort);
    return;
  }
  if (!tr) return;
  tr.innerHTML = columns.map(c => {
    const sorted = sortState.col === c.key;
    const cls = ['sortable', c.right ? 'right' : '', sorted ? (sortState.dir === 1 ? 'sorted-asc' : 'sorted-desc') : ''].filter(Boolean).join(' ');
    const ind = sorted ? (sortState.dir === 1 ? '▲' : '▼') : '↕';
    const titleAttr = c.title ? ` title="${c.title.replace(/"/g, '&quot;')}"` : '';
    return `<th class="${cls}" data-col="${c.key}"${titleAttr}>${c.label}<span class="sort-ind">${ind}</span></th>`;
  }).join('');
  tr.querySelectorAll('th.sortable').forEach(th => {
    th.addEventListener('click', e => {
      e.stopPropagation();
      const col = th.dataset.col;
      if (sortState.col === col) sortState.dir *= -1;
      else { sortState.col = col; sortState.dir = 1; }
      onSort();
    });
  });
}

function updateSortHeadIndicators(containerId, sortState) {
  if (window.TableSort) {
    window.TableSort.updateDataHeadIndicators($(containerId), sortState);
    return;
  }
  const tr = $(containerId);
  if (!tr) return;
  tr.querySelectorAll('th.sortable').forEach(th => {
    const sorted = sortState.col === th.dataset.col;
    th.classList.remove('sorted-asc', 'sorted-desc');
    if (sorted) th.classList.add(sortState.dir === 1 ? 'sorted-asc' : 'sorted-desc');
    const ind = th.querySelector('.sort-ind');
    if (ind) ind.textContent = sorted ? (sortState.dir === 1 ? '▲' : '▼') : '↕';
  });
}

function initMainTableHead() {
  renderSortableHead('nfTableHead', NF_COLUMNS, tableSort, renderTable);
}

function nfExportRow(x) {
  enrichNF(x);
  return {
    'Nota Fiscal': x.nf,
    'Data NF': x.dtNF ? new Date(x.dtNF) : '',
    'Cliente': x.cliente || '',
    'Mês': fmtMesLabel(x.mesRef),
    'Transportador': x.transportador,
    'Modalidade': x.modalidade,
    'Valor NF (Unilog)': x.valorUnilog ?? x.valorNF,
    'Valor SAP': x.valorSAP ?? '',
    'Δ SAP (SAP − Unilog)': x.valorDiff ?? '',
    'Δ valor relevante': x.sapValorMismatch ? 'Sim' : 'Não',
    'Qtd CT-e': x.nCte,
    'Pago': x.pago,
    'Esperado (6%)': x.esperado,
    'Diferença': x.diff,
    '% Pago': x.pct,
    'Estado': (statusMeta[x.status] || {}).label || x.status,
    'Motivo': x.motivo,
    'NF no SAP': x.sapMissing ? 'Não' : (x.sapFound ? 'Sim' : ''),
    'Tipo anomalia SAP': x.sapAnomalyType || '',
    'Explicação anomalia SAP': x.sapAnomalyReason || ''
  };
}

function cteExportRows() {
  const rows = [];
  currentNFs.forEach(x => {
    (x.ctes || []).forEach(c => {
      rows.push({
        'Nota Fiscal': x.nf,
        'Cliente': x.cliente || '',
        'Transportador': x.transportador,
        'Nº CT-e': c.numCte,
        'Data CT-e': c.dtCte ? new Date(c.dtCte) : '',
        'Tipo Operação': c.tipoOp || '',
        'Devolução': c.devolucao ? 'Sim' : 'Não',
        'Peso (kg)': c.peso,
        'Pago': c.pago,
        '% da NF': c.pct,
        'Desvio vs 6%': c.diff,
        'Explicação': c.explicacao || '',
        'Validação': c.validacao || 'pendente',
        'Estado NF': (statusMeta[x.status] || {}).label || x.status
      });
    });
  });
  return rows;
}

function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename);
}

function exportFullWorkbook() {
  if (!currentNFs.length || !currentSummary) {
    fteToast('Não há dados para exportar.');
    return;
  }
  const s = currentSummary;
  const wb = XLSX.utils.book_new();

  const resumo = [
    ['Conciliação CT-e × Nota Fiscal — Delta Foods Brasil'],
    ['Ficheiro', s.fileName],
    ['Período CT-e', `${fmtDate(s.periodoInicio)} a ${fmtDate(s.periodoFim)}`],
    ['Total NFs', s.totalNF],
    ['CT-e único / múltiplos', `${s.nUnicoCte} / ${s.nMultiplosCte}`],
    [],
    ['Indicador', 'Valor'],
    ['Faturação (valor total NFs)', s.totalValorNF],
    ['Total pago (frete)', s.totalPago],
    ['% pago sobre NFs', s.pctPagoSobreNF],
    ['Esperado (6%)', s.totalEsperado],
    ['Excesso líquido', s.excesso],
    ['Excesso positivo', s.excessoPositivo],
    ['Déficit', s.deficit],
    [],
    ['Categoria', 'Nº NFs', 'Diferença', 'Pago'],
    ['Conforme (~6%)', s.byStatus.ok, s.excessoByStatus.ok, s.pagoByStatus.ok],
    ['Sobretaxa fixa (~R$ 165)', s.byStatus.fix165, s.excessoByStatus.fix165, s.pagoByStatus.fix165],
    ['Tarifa mínima', s.byStatus.min, s.excessoByStatus.min, s.pagoByStatus.min],
    ['Com devolução', s.byStatus.dev, s.excessoByStatus.dev, s.pagoByStatus.dev],
    ['Sem devolução — investigar', s.byStatus.flag, s.excessoByStatus.flag, s.pagoByStatus.flag],
    ['Abaixo do esperado', s.byStatus.low, s.excessoByStatus.low, s.pagoByStatus.low],
  ];
  if (s.sapLoaded) {
    resumo.push([]);
    resumo.push(['SAP', 'Valor']);
    resumo.push(['NFs cruzadas com SAP', s.nSapMatched]);
    resumo.push(['Δ valor SAP ≠ Unilog', s.nValorMismatch]);
    resumo.push(['NF sem correspondência SAP (anomalias)', s.nSapMissing]);
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumo), 'Resumo');

  const byCat = [];
  Object.keys(CATEGORY_META).forEach(key => {
    currentNFs.filter(x => x.status === key).forEach(x => byCat.push({ ...nfExportRow(x), '_Categoria': CATEGORY_META[key].title }));
  });
  if (byCat.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(byCat), 'Por categoria');
  }

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(currentNFs.map(nfExportRow)), 'Notas Fiscais');

  const anomalies = getSapAnomalies(currentNFs);
  if (anomalies.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(anomalies.map(anomalyExportRow)), 'Anomalias SAP');
  }

  const ctes = cteExportRows();
  if (ctes.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ctes), 'CT-e');

  const mensal = buildMonthlyExportRows(monthlyRows);
  if (mensal.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mensal), 'Por mês');

  const stamp = new Date().toISOString().slice(0, 10);
  downloadWorkbook(wb, `Conciliacao_CTE_NF_${stamp}.xlsx`);
  fteToast('Excel exportado com todas as folhas.');
}

function exportFilteredView() {
  const list = getFilteredNFs();
  if (!list.length) { fteToast('Nenhuma linha na vista actual.'); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(list.map(nfExportRow)), 'Vista filtrada');
  const ctes = [];
  list.forEach(x => (x.ctes || []).forEach(c => {
    ctes.push({
      'Nota Fiscal': x.nf, 'Cliente': x.cliente || '', 'Nº CT-e': c.numCte,
      'Data CT-e': c.dtCte ? new Date(c.dtCte) : '',
      'Devolução': c.devolucao ? 'Sim' : 'Não', 'Pago': c.pago
    });
  }));
  if (ctes.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ctes), 'CT-e filtrados');
  downloadWorkbook(wb, `Conciliacao_filtrada_${new Date().toISOString().slice(0, 10)}.xlsx`);
  fteToast('Vista filtrada exportada.');
}

function exportMonthlyWorkbook() {
  if (!monthlyRows.length) { fteToast('Sem dados mensais.'); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildMonthlyExportRowsFull(monthlyRows)), 'Resumo mensal');
  const det = [];
  monthlyRows.forEach(m => {
    m.items.forEach(x => det.push({ ...nfExportRow(x), 'Mês': m.mesLabel }));
  });
  if (det.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(det), 'Detalhe por mês');
  downloadWorkbook(wb, `Conciliacao_mensal_${new Date().toISOString().slice(0, 10)}.xlsx`);
  fteToast('Exportação mensal concluída.');
}

function getFilteredNFs() {
  const statusF = $('filterStatus').value;
  const transpF = $('filterTransp').value;
  const mesF = $('filterMes').value;
  const searchF = $('searchNF').value.trim().toLowerCase();
  const valorDiffF = $('filterValorDiff')?.checked;
  const sapMissingF = $('filterSapMissing')?.checked;
  let list = currentNFs;
  if (statusF) list = list.filter(x => x.status === statusF);
  if (transpF) list = list.filter(x => x.transportador === transpF);
  if (mesF) list = list.filter(x => x.mesRef === mesF);
  if (valorDiffF) list = list.filter(x => x.sapValorMismatch);
  if (sapMissingF) list = list.filter(x => x.sapMissing);
  if (searchF) list = list.filter(x =>
    String(x.nf).includes(searchF) ||
    String(x.cliente || '').toLowerCase().includes(searchF)
  );
  return sortRows(list, NF_COLUMNS, tableSort);
}

function renderMonthlyTable() {
  if (!currentNFs.length) {
    $('monthTableBody').innerHTML = '<tr><td colspan="10" class="empty">Carrega uma análise primeiro.</td></tr>';
    return;
  }
  monthlyRows = computeMonthly(currentNFs);
  renderSortableHead('monthTableHead', MONTH_COLUMNS, monthSort, renderMonthlyTable);
  const sorted = sortMonthlyRows(monthlyRows, monthSort);
  const groupByYear = monthSort.col === 'mesRef';
  const displayRows = groupByYear
    ? buildMonthlyDisplayRows(sorted)
    : [
      ...sorted.map(m => ({ type: 'month', ...m })),
      ...(sorted.length ? [{ type: 'total', label: 'Total geral', ...aggregateMonthTotals(sorted) }] : [])
    ];
  const maxExcesso = Math.max(...sorted.map(m => m.excesso), 0);

  const body = $('monthTableBody');
  body.innerHTML = displayRows.map(row => {
    if (row.type === 'year') {
      return `<tr class="month-year-row"><td colspan="10"><strong>${row.label}</strong></td></tr>`;
    }
    const isTotal = row.type === 'subtotal' || row.type === 'total';
    const m = row;
    const high = !isTotal && m.excesso > 0 && m.excesso >= maxExcesso * 0.85;
    const active = !isTotal && selectedMonth === m.mesRef;
    const label = row.type === 'month' ? m.mesLabel : row.label;
    const trCls = isTotal
      ? (row.type === 'total' ? 'month-total-row' : 'month-subtotal-row')
      : `month-row-clickable${high ? ' month-row-high' : ''}${active ? ' active-month' : ''}`;
    const dataMes = row.type === 'month' ? ` data-mes="${m.mesRef}"` : '';
    return `<tr class="${trCls}"${dataMes}>
      <td><strong>${label}</strong></td>
      <td class="right">${m.totalNF}</td>
      <td class="right">${fmtMoney(m.totalValorNF)}</td>
      <td class="right">${fmtMoney(m.totalPago)}</td>
      <td class="right">${fmtMoney(m.totalEsperado)}</td>
      <td class="right" style="color:${m.excesso > 0 ? '#b3261e' : 'inherit'}">${fmtMoney(m.excesso)}</td>
      <td class="right">${fmtPct(m.pctPagoSobreNF)}</td>
      <td class="right" style="color:${m.qtd_fix165 ? '#c2410c' : 'inherit'}">${m.qtd_fix165}</td>
      <td class="right">${m.qtd_min}</td>
      <td class="right" style="color:${m.qtd_flag ? '#b3261e' : 'inherit'}">${m.qtd_flag}</td>
    </tr>`;
  }).join('');

  body.querySelectorAll('tr.month-row-clickable').forEach(tr => {
    tr.addEventListener('click', () => {
      selectedMonth = tr.dataset.mes;
      $('filterMes').value = selectedMonth;
      switchFteTab('analise-cte');
      switchCteSub('detalhe');
      renderTable();
      $('nfTable')?.scrollIntoView({ behavior: 'smooth' });
      fteToast('Filtrado por ' + fmtMesLabel(selectedMonth));
    });
  });
  fteScheduleTableSort();
}

function renderAll() {
  const s = currentSummary;
  if (!s) return;
  const sapNote = s.sapLoaded ? ` · SAP: ${s.nSapMatched}/${s.totalNF} NFs cruzadas` : '';
  const mismatchNote = s.nValorMismatch ? ` · <span style="color:#b3261e;font-weight:600">${s.nValorMismatch} Δ valor SAP≠Unilog</span>` : '';
  const sapMissingNote = s.nSapMissing ? ` · <span style="color:#b3261e;font-weight:600">${s.nSapMissing} NF sem SAP</span>` : '';
  const srcNote = cteAnalysisSource === 'quinzenal'
    ? ' · <span style="color:var(--muted)">fonte: quinzenais B2B</span>'
    : (cteAnalysisSource === 'conciliacao' ? ' · <span style="color:var(--muted)">fonte: CT-e×NF</span>' : '');
  const periodoLabel = $('periodoLabel');
  if (periodoLabel) {
    periodoLabel.innerHTML =
      `Ficheiro: ${s.fileName} · Período CT-e: ${fmtDate(s.periodoInicio)} a ${fmtDate(s.periodoFim)} · ${s.totalNF} notas fiscais · ${s.nUnicoCte} com CT-e único · ${s.nMultiplosCte} com múltiplos CT-e${sapNote}${mismatchNote}${sapMissingNote}${srcNote}`;
  }

  const warnEl = $('dataWarning');
  if (warnEl) {
    if (s.loadWarning) {
      warnEl.style.display = 'block';
      warnEl.className = 'data-warn';
      warnEl.innerHTML = s.loadWarning;
    } else if (cteAnalysisSource === 'quinzenal') {
      warnEl.style.display = 'block';
      warnEl.className = 'data-warn';
      warnEl.innerHTML = '<strong>Análise a partir dos quinzenais B2B:</strong> 6% / R$165 / multi-CT-e usam linhas do relatório quinzenal (pago, Devolução, Num. CTE, Dt CTE). '
        + 'ZFACT continua a alimentar anomalias SAP.';
    } else {
      warnEl.style.display = 'none';
      warnEl.innerHTML = '';
    }
  }

  const sapKpi = s.sapLoaded ? `
    <div class="kpi${s.nValorMismatch ? ' flag' : ''}"><div class="label">Δ valor SAP ≠ Unilog</div><div class="value">${s.nValorMismatch}</div><div class="sub">${s.nSapMatched} NFs cruzadas com SAP</div></div>
    <div class="kpi${s.nSapMissing ? ' flag' : ''}"><div class="label">NF sem SAP</div><div class="value">${s.nSapMissing}</div><div class="sub">anomalias para investigar</div></div>` : '';

  const kpisEl = $('kpis');
  if (kpisEl) kpisEl.innerHTML = `
    <div class="kpi"><div class="label">Faturação (valor NFs)</div><div class="value">${fmtMoney(s.totalValorNF)}</div><div class="sub">${s.totalNF} notas fiscais</div></div>
    <div class="kpi"><div class="label">Total pago (frete)</div><div class="value">${fmtMoney(s.totalPago)}</div><div class="sub">${fmtPct(s.pctPagoSobreNF)} sobre faturação (meta: 6%)</div></div>
    <div class="kpi"><div class="label">Esperado (6%)</div><div class="value">${fmtMoney(s.totalEsperado)}</div></div>
    <div class="kpi flag"><div class="label">Excesso líquido</div><div class="value">${fmtMoney(s.excesso)}</div><div class="sub">+${fmtPct(s.excessoPct)} vs esperado</div></div>
    <div class="kpi"><div class="label">NFs conformes</div><div class="value">${s.byStatus.ok}</div><div class="sub">${s.totalNF ? fmtPct(s.byStatus.ok / s.totalNF) : '0%'} do total · sem sobretaxa R$165</div></div>
    <div class="kpi${s.byStatus.fix165 ? ' flag' : ''}"><div class="label">Sobretaxa ~R$165</div><div class="value">${s.byStatus.fix165}</div><div class="sub">excedente R$160–170 · padrão recorrente</div></div>
    <div class="kpi flag"><div class="label">NFs a investigar ⚠️</div><div class="value">${s.byStatus.flag}</div><div class="sub">múltiplos CT-e sem devolução</div></div>
    ${sapKpi}
  `;

  const reconcileBox = $('reconcileBox');
  if (reconcileBox) reconcileBox.innerHTML = `
    <div class="reconcile">
      <strong>Reconciliação:</strong> Faturação ${fmtMoney(s.totalValorNF)} · Pago ${fmtMoney(s.totalPago)} (${fmtPct(s.pctPagoSobreNF)}) = Esperado ${fmtMoney(s.totalEsperado)} + Excesso ${fmtMoney(s.excessoPositivo)} + Déficit ${fmtMoney(s.deficit)}
      <div class="eq">
        <span class="chip ok"><strong>${fmtMoney(s.totalPago)}</strong> pago</span>
        <span>=</span>
        <span class="chip">${fmtMoney(s.totalEsperado)} esperado</span>
        <span>+</span>
        <span class="chip pos">${fmtMoney(s.excessoPositivo)} excesso</span>
        <span>+</span>
        <span class="chip neg">${fmtMoney(s.deficit)} déficit</span>
        <span style="color:var(--gray);font-size:12px;">(${Math.abs(s.totalPago - (s.totalEsperado + s.excessoPositivo + s.deficit)) < 0.02 ? '✓ bate certo' : 'verificar arredondamentos'})</span>
      </div>
    </div>`;

  const bd = [
    { key: 'ok', cls: 'bd-ok', title: 'Conforme (~6%)', desc: '1 CT-e, pago entre 5,9% e 6,1% do valor da NF, sem sobretaxa fixa R$165.' },
    { key: 'fix165', cls: 'bd-fix165', title: 'Sobretaxa fixa (~R$ 165)', desc: '1 CT-e com excedente entre R$160 e R$170 sobre o esperado (6%) — cobrança adicional recorrente, mesmo quando o % da NF fica ~6%.' },
    { key: 'min', cls: 'bd-min', title: 'Tarifa mínima (provável)', desc: '1 CT-e acima de 6,1% — típico em NFs de baixo valor (frete mínimo do transportador). Distinto da sobretaxa fixa R$165.' },
    { key: 'dev', cls: 'bd-dev', title: 'Com devolução', desc: 'Múltiplos CT-e com retorno assinalado (ida + volta).' },
    { key: 'flag', cls: 'bd-flag', title: '⚠️ Sem devolução — investigar', desc: 'Múltiplos CT-e sem devolução — cada um pode cobrar % sobre a NF inteira.' },
    { key: 'low', cls: 'bd-low', title: 'Abaixo do esperado', desc: '1 CT-e abaixo de 5,9% — pagou menos que os 6%.' },
  ];
  const breakdownEl = $('breakdown');
  if (breakdownEl) {
    breakdownEl.innerHTML = bd.filter(b => (b.key !== 'low' && b.key !== 'fix165') || s.byStatus[b.key]).map(b => `
    <div class="bd-card ${b.cls}${activeSubPanel === b.key ? ' active' : ''}" data-status="${b.key}" role="button" tabindex="0">
      <div class="n">${s.byStatus[b.key]} NF · dif. ${fmtMoney(s.excessoByStatus[b.key])}</div>
      <div class="t"><strong>${b.title}</strong><br>${b.desc}</div>
      <div class="t muted" style="margin-top:6px;">Pago: ${fmtMoney(s.pagoByStatus[b.key])}</div>
    </div>
  `).join('');

    document.querySelectorAll('.bd-card').forEach(card => {
      card.addEventListener('click', () => toggleSubPanel(card.dataset.status));
      card.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleSubPanel(card.dataset.status); } });
    });
  }

  renderSubPanels();

  const transpSet = [...new Set(currentNFs.map(x => x.transportador))].sort();
  const tSel = $('filterTransp');
  if (tSel) {
    tSel.innerHTML = '<option value="">Todos os transportadores</option>' + transpSet.map(t => `<option value="${t}">${t}</option>`).join('');
  }

  monthlyRows = computeMonthly(currentNFs);
  const mSel = $('filterMes');
  if (mSel) {
    mSel.innerHTML = '<option value="">Todos os meses</option>' + monthlyRows.map(m =>
      `<option value="${m.mesRef}"${selectedMonth === m.mesRef ? ' selected' : ''}>${m.mesLabel} (${m.totalNF} NF)</option>`
    ).join('');
  }

  initMainTableHead();
  renderAnomaliesPanel();
  renderTable();
  renderMonthlyTable();
  renderCteSubPanels();
  if (quinzenalPack?.files?.length) {
    refreshQuinzenalCompare();
    if ($('tab-analise-b2c')?.style.display === 'block') renderB2cAnalysisTab();
    if ($('tab-cte-vs-qz')?.style.display === 'block') renderB2bCompareTab();
    if ($('tab-resumo-total')?.style.display === 'block') renderResumoTotal();
  }
  refreshCustoUnilogIfVisible();
}

function renderAnomaliesPanel() {
  const section = $('anomaliesSection');
  const hint = $('anomaliesHint');
  if (!section || !hint) return;

  const s = currentSummary;
  if (!s?.sapLoaded) {
    hint.style.display = currentNFs.length ? 'block' : 'none';
    const body = $('anomaliesTableBody');
    if (body && currentNFs.length) {
      body.innerHTML = '<tr><td colspan="10" class="empty">Carrega o Excel SAP para detectar anomalias.</td></tr>';
    }
    const countEl = $('anomaliesCount');
    if (countEl) countEl.textContent = '';
    return;
  }

  hint.style.display = 'none';

  const anomalies = getSapAnomalies(currentNFs);
  const countEl = $('anomaliesCount');
  if (countEl) countEl.textContent = anomalies.length ? `${anomalies.length} anomalia(s)` : 'Nenhuma';

  const body = $('anomaliesTableBody');
  if (!body) return;

  if (!anomalies.length) {
    body.innerHTML = '<tr><td colspan="10" class="empty">Nenhuma NF do Unilog em falta no SAP — cruzamento OK.</td></tr>';
    return;
  }

  body.innerHTML = anomalies.map(x => {
    const meta = statusMeta[x.status] || {};
    enrichNF(x);
    return `<tr class="row-clickable nf-sap-missing" data-nf="${x.nf}">
      <td>${x.nf}</td>
      <td>${x.cliente || '-'}</td>
      <td>${fmtDate(x.dtNF)}</td>
      <td class="right">${fmtMoney(x.valorUnilog ?? x.valorNF)}</td>
      <td>${x.transportador}</td>
      <td class="right">${x.nCte > 1 ? `<strong>${x.nCte}</strong>` : x.nCte}</td>
      <td class="right">${fmtMoney(x.pago)}</td>
      <td><span class="badge ${meta.cls}">${meta.label}</span></td>
      <td>${x.sapAnomalyType || SAP_ANOMALY_TYPE}</td>
      <td class="anomaly-reason">${x.sapAnomalyReason || buildSapMissingReason(x)}</td>
    </tr>`;
  }).join('');

  body.querySelectorAll('tr.row-clickable').forEach(tr => {
    tr.addEventListener('click', () => {
      $('searchNF').value = tr.dataset.nf;
      if ($('filterSapMissing')) $('filterSapMissing').checked = true;
      renderTable();
      $('nfTable').scrollIntoView({ behavior: 'smooth' });
    });
  });

  const exportBtn = $('exportAnomaliesBtn');
  if (exportBtn && !exportBtn._bound) {
    exportBtn._bound = true;
    exportBtn.addEventListener('click', () => {
      const list = getSapAnomalies(currentNFs);
      if (!list.length) { fteToast('Sem anomalias SAP para exportar.'); return; }
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(list.map(anomalyExportRow)), 'Anomalias SAP');
      downloadWorkbook(wb, `Anomalias_SAP_${new Date().toISOString().slice(0, 10)}.xlsx`);
      fteToast('Anomalias SAP exportadas.');
    });
  }
  fteEnableDomSort(body);
}

const CATEGORY_META = {
  ok: { title: 'Conforme (~6%)', cls: 'b-ok' },
  fix165: { title: 'Sobretaxa fixa (~R$ 165)', cls: 'b-fix165' },
  min: { title: 'Tarifa mínima (provável)', cls: 'b-min' },
  dev: { title: 'Com devolução', cls: 'b-dev' },
  flag: { title: '⚠️ Sem devolução — investigar', cls: 'b-flag' },
  low: { title: 'Abaixo do esperado', cls: 'b-low' }
};

function toggleSubPanel(status) {
  activeSubPanel = activeSubPanel === status ? null : status;
  if (activeSubPanel) subTableSort = { col: 'diff', dir: -1 };
  $('filterStatus').value = activeSubPanel || '';
  document.querySelectorAll('.bd-card').forEach(c => {
    c.classList.toggle('active', c.dataset.status === activeSubPanel);
  });
  renderSubPanels();
  renderTable();
  if (activeSubPanel) {
    const el = $('sub-' + activeSubPanel);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

function renderSubPanels() {
  const host = $('subPanels');
  if (!activeSubPanel) { host.innerHTML = ''; return; }

  const meta = CATEGORY_META[activeSubPanel];
  let list = currentNFs.filter(x => x.status === activeSubPanel);
  list = sortRows(list, SUB_COLUMNS, subTableSort);
  const s = currentSummary;
  const totalDif = s.excessoByStatus[activeSubPanel];
  const totalPago = s.pagoByStatus[activeSubPanel];

  const rows = list.map(x => {
    const cteInfo = x.nCte > 1
      ? `${x.nCte} CT-e${x.temDevolucao ? ' (com devolução)' : ''}`
      : '1 CT-e';
    const mismatchCls = x.sapValorMismatch ? ' nf-val-mismatch' : '';
    return `<tr class="row-clickable${mismatchCls}" data-nf="${x.nf}">
      <td>${x.nf}</td><td>${fmtDate(x.dtNF)}</td><td>${x.cliente || '-'}</td><td>${x.transportador}</td><td>${cteInfo}</td>
      <td class="right">${fmtMoney(x.valorUnilog ?? x.valorNF)}</td>
      <td class="right">${x.valorSAP != null ? fmtMoney(x.valorSAP) : '-'}</td>
      <td class="right">${fmtValorDiff(x.valorDiff, x.sapValorMismatch)}</td>
      <td class="right">${fmtMoney(x.pago)}</td>
      <td class="right">${fmtMoney(x.esperado)}</td>
      <td class="right" style="color:${x.diff > 0.5 ? '#b3261e' : (x.diff < -0.5 ? '#555' : 'inherit')}">${fmtMoney(x.diff)}</td>
      <td class="right">${fmtPct(x.pct)}</td>
    </tr>`;
  }).join('');

  host.innerHTML = `
    <div class="sub-panel open" id="sub-${activeSubPanel}">
      <div class="card">
        <div class="sub-head">
          <h4>${meta.title} — ${list.length} NF · dif. ${fmtMoney(totalDif)} · pago ${fmtMoney(totalPago)}</h4>
          <div style="display:flex;gap:8px;">
            <button class="btn secondary sub-export" style="padding:6px 12px;font-size:12px;" type="button">Exportar categoria</button>
            <button class="btn secondary sub-close" style="padding:6px 12px;font-size:12px;" type="button">Fechar</button>
          </div>
        </div>
        <div class="tbl-wrap" style="max-height:360px;border:none;border-radius:0;">
          <table>
            <thead><tr id="subTableHead"></tr></thead>
            <tbody>${rows || '<tr><td colspan="12" class="empty">Sem NFs nesta categoria.</td></tr>'}</tbody>
          </table>
        </div>
      </div>
    </div>`;

  renderSortableHead('subTableHead', SUB_COLUMNS, subTableSort, () => renderSubPanels());

  host.querySelectorAll('tr.row-clickable').forEach(tr => {
    tr.addEventListener('click', () => {
      $('searchNF').value = tr.dataset.nf;
      renderTable();
      $('nfTable').scrollIntoView({ behavior: 'smooth' });
    });
  });
  const closeBtn = host.querySelector('.sub-close');
  if (closeBtn) closeBtn.addEventListener('click', () => toggleSubPanel(activeSubPanel));
  const exportBtn = host.querySelector('.sub-export');
  if (exportBtn) exportBtn.addEventListener('click', () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(list.map(nfExportRow)), meta.title.slice(0, 31));
    downloadWorkbook(wb, `Conciliacao_${activeSubPanel}_${new Date().toISOString().slice(0, 10)}.xlsx`);
    fteToast('Categoria exportada.');
  });
}

const statusMeta = {
  ok: { cls: 'b-ok', label: 'Conforme' },
  fix165: { cls: 'b-fix165', label: 'Sobretaxa R$165' },
  min: { cls: 'b-min', label: 'Tarifa mínima' },
  dev: { cls: 'b-dev', label: 'Com devolução' },
  flag: { cls: 'b-flag', label: '⚠️ Investigar' },
  low: { cls: 'b-low', label: 'Abaixo' }
};

const FTE_NF_ROW_PAGE = 400;
let fteNfRowLimit = FTE_NF_ROW_PAGE;
let _fteNfKeepLimit = false;
function showMoreNfRows() {
  fteNfRowLimit += FTE_NF_ROW_PAGE;
  _fteNfKeepLimit = true;
  renderTable();
}

function renderTable() {
  const list = getFilteredNFs();
  updateSortHeadIndicators('nfTableHead', tableSort);
  if (!_fteNfKeepLimit) fteNfRowLimit = FTE_NF_ROW_PAGE;
  _fteNfKeepLimit = false;

  const body = $('nfTableBody');
  const display = list.slice(0, fteNfRowLimit);
  const parts = [];
  for (let i = 0; i < display.length; i++) {
    const x = display[i];
    enrichNF(x);
    const meta = statusMeta[x.status] || statusMeta.flag;
    const cls = 'row-clickable'
      + (x.sapValorMismatch ? ' nf-val-mismatch' : '')
      + (x.sapMissing ? ' nf-sap-missing' : '');
    parts.push(`<tr class="${cls}" data-nf-i="${i}">
      <td>${x.nf}</td>
      <td>${fmtDate(x.dtNF)}</td>
      <td>${x.cliente || '-'}</td>
      <td>${x.transportador}</td>
      <td>${x.modalidade}</td>
      <td class="right">${fmtMoney(x.valorUnilog ?? x.valorNF)}</td>
      <td class="right">${x.valorSAP != null ? fmtMoney(x.valorSAP) : '-'}</td>
      <td class="right">${fmtValorDiff(x.valorDiff, x.sapValorMismatch)}</td>
      <td class="right">${x.nCte > 1 ? `<strong>${x.nCte}</strong>` : x.nCte}</td>
      <td class="right">${fmtMoney(x.pago)}</td>
      <td class="right">${fmtMoney(x.esperado)}</td>
      <td class="right" style="color:${x.diff > 0.5 ? '#b3261e' : (x.diff < -0.5 ? '#555' : 'inherit')}">${fmtMoney(x.diff)}</td>
      <td class="right">${fmtPct(x.pct)}</td>
      <td><span class="badge ${meta.cls}">${meta.label}</span></td>
    </tr>`);
  }
  body.innerHTML = parts.length
    ? parts.join('')
    : `<tr><td colspan="${nfColCount()}" class="empty">Sem resultados para este filtro.</td></tr>`;
  body.querySelectorAll('tr.row-clickable').forEach(tr => {
    const i = parseInt(tr.getAttribute('data-nf-i'), 10);
    const x = display[i];
    if (x) tr.addEventListener('click', () => toggleDetail(tr, x));
  });
  const moreWrap = $('nfMoreWrap');
  const moreHint = $('nfMoreHint');
  if (moreWrap) {
    const hasMore = list.length > display.length;
    moreWrap.style.display = hasMore ? 'block' : 'none';
    if (moreHint) moreHint.textContent = hasMore ? `A mostrar ${display.length} de ${list.length}` : '';
  }
}

function toggleDetail(tr, x) {
  const next = tr.nextSibling;
  if (next && next.classList && next.classList.contains('detail-row')) { next.remove(); return; }
  document.querySelectorAll('.detail-row').forEach(d => d.remove());
  const dr = document.createElement('tr');
  dr.className = 'detail-row';
  const multi = x.nCte > 1;
  const ctesRows = x.ctes.map(c => {
    const valCell = multi ? `<td class="cte-val-cell">
      <div class="cte-val-btns">
        <button type="button" class="cte-val-btn ok${c.validacao === 'validado' ? ' active' : ''}" data-nf="${x.nf}" data-key="${c.cteKey}" data-val="validado">✓ Validar</button>
        <button type="button" class="cte-val-btn no${c.validacao === 'rejeitado' ? ' active' : ''}" data-nf="${x.nf}" data-key="${c.cteKey}" data-val="rejeitado">✗ Rejeitar</button>
      </div>
      <span class="cte-val-lbl ${c.validacao}">${c.validacao === 'validado' ? 'Validado' : (c.validacao === 'rejeitado' ? 'Rejeitado' : 'Pendente')}</span>
    </td>` : '';
    const diffColor = c.diff > 0.5 ? '#b3261e' : (c.diff < -0.5 ? '#555' : 'inherit');
    return `<tr class="cte-row${c.validacao !== 'pendente' ? ' cte-row-' + c.validacao : ''}">
      <td>${c.numCte || '-'}</td>
      <td>${c.dtCte ? new Date(c.dtCte).toLocaleDateString('pt-BR') : '-'}</td>
      <td>${c.tipoOp || '-'}</td>
      <td>${c.devolucao ? 'Sim' : 'Não'}</td>
      <td class="right">${fmtMoney(c.peso)} kg</td>
      <td class="right">${fmtMoney(c.pago)}</td>
      <td class="right">${fmtPct(c.pct)}</td>
      <td class="right" style="color:${diffColor}">${fmtMoney(c.diff)}</td>
      <td class="cte-explicacao">${c.explicacao || '-'}</td>
      ${valCell}
    </tr>`;
  }).join('');
  const valHead = multi ? '<th>Validação</th>' : '';
  const summaryExtra = x.nCte === 1
    ? `<p style="margin:0 0 8px;font-size:11px;color:var(--muted);">Esperado 6%: ${fmtMoney(x.esperado)} · Desvio: <span style="color:${x.diff > 0.5 ? '#b3261e' : (x.diff < -0.5 ? '#555' : 'inherit')}">${fmtMoney(x.diff)} (${fmtPp(x.pct - CTE_PCT_TARGET)})</span></p>`
    : `<p style="margin:0 0 8px;font-size:11px;color:var(--muted);">${x.nCte} CT-e · Total pago ${fmtMoney(x.pago)} (${fmtPct(x.pct)} da NF) · Validar ou rejeitar cada CT-e individualmente.</p>`;
  const sapBlock = x.sapMissing
    ? `<p style="margin:0 0 8px;font-size:11px;color:#b3261e;">
      <strong>Anomalia SAP:</strong> ${x.sapAnomalyType || SAP_ANOMALY_TYPE} — ${x.sapAnomalyReason || buildSapMissingReason(x)}
    </p>`
    : (x.sapFound ? `<p style="margin:0 0 8px;font-size:11px;color:var(--muted);">
    <strong>SAP:</strong> ${fmtDate(x.dtNF)} · ${x.cliente || '-'} · Unilog ${fmtMoney(x.valorUnilog ?? x.valorNF)} vs SAP ${fmtMoney(x.valorSAP)}${x.sapValorMismatch ? ` — <span style="color:#b3261e;font-weight:600">Δ relevante ${fmtMoney(x.valorDiff)}</span>` : (x.valorDiff != null && Math.abs(num(x.valorDiff)) > SAP_VALOR_NEAR_EPS ? ` (Δ ${fmtMoney(x.valorDiff)})` : ' — valores alinhados')}
  </p>` : '');
  dr.innerHTML = `<td colspan="${nfColCount()}"><div class="detail-inner">
    <p style="margin:0 0 6px;"><strong>Análise:</strong> ${x.motivo}</p>
    ${sapBlock}
    ${summaryExtra}
    <table class="cte-detail-table"><thead><tr>
      <th>Nº CT-e</th><th>Data CT-e</th><th>Tipo Op.</th><th>Devolução</th><th>Peso</th><th>Pago</th><th>% NF</th><th>Desvio 6%</th><th>Explicação</th>${valHead}
    </tr></thead><tbody>${ctesRows}</tbody></table>
  </div></td>`;
  tr.parentNode.insertBefore(dr, tr.nextSibling);
  if (multi) {
    dr.querySelectorAll('.cte-val-btn').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        setCteValidacao(btn.dataset.nf, btn.dataset.key, btn.dataset.val);
        reopenNfDetail(btn.dataset.nf);
      });
    });
  }
}

function nfsFromDetalhe(rows) {
  return rows.map(d => {
    let ctes = [];
    let cliente = '';
    let dtNF = null;
    if (d.ctes_json) {
      try {
        const parsed = JSON.parse(d.ctes_json);
        if (Array.isArray(parsed)) {
          ctes = parsed;
        } else if (parsed && typeof parsed === 'object') {
          ctes = parsed.ctes || [];
          cliente = parsed.cliente || '';
          dtNF = parsed.dtSAP || null;
        }
      } catch (e) { ctes = []; }
    }
    return enrichNF(buildNfRecord({
      nf: d.nota_fiscal,
      cliente,
      transportador: d.transportador,
      modalidade: d.modalidade,
      valorNF: d.valor_nf,
      ctes,
      temDevolucao: ctes.some(c => c.devolucao) || d.status === 'dev',
      dtNF
    }));
  }).sort((a, b) => b.diff - a.diff);
}

async function insertUploadAndDetails() {
  const s = currentSummary;
  const { data: upload, error: e1 } = await sb.from('cte_nf_uploads').insert({
    ficheiro_nome: s.fileName,
    periodo_inicio: s.periodoInicio ? s.periodoInicio.toISOString().slice(0, 10) : null,
    periodo_fim: s.periodoFim ? s.periodoFim.toISOString().slice(0, 10) : null,
    total_nf: s.totalNF,
    total_pago: s.totalPago,
    total_esperado: s.totalEsperado,
    excesso: s.excesso,
    excesso_pct: s.excessoPct,
    qtd_ok: s.byStatus.ok, qtd_min: s.byStatus.min, qtd_dev: s.byStatus.dev,
    qtd_flag: s.byStatus.flag, qtd_low: s.byStatus.low,
    uploader_nome: null
  }).select().single();
  if (e1) throw e1;

  const detalhes = currentNFs.map(x => ({
    upload_id: upload.id, nota_fiscal: String(x.nf), transportador: x.transportador,
    modalidade: x.modalidade, valor_nf: x.valorNF, n_cte: x.nCte, pago: x.pago,
    esperado: x.esperado, diferenca: x.diff, pct_pago: x.pct, status: x.status, motivo: x.motivo,
    ctes_json: JSON.stringify({ ctes: x.ctes || [], cliente: x.cliente || '', dtSAP: x.dtNF || null })
  }));

  const chunkSize = 500;
  for (let i = 0; i < detalhes.length; i += chunkSize) {
    const chunk = detalhes.slice(i, i + chunkSize);
    let { error: e2 } = await sb.from('cte_nf_detalhe').insert(chunk);
    if (e2 && /ctes_json|column/i.test(e2.message || '')) {
      const slim = chunk.map(({ ctes_json, ...rest }) => rest);
      ({ error: e2 } = await sb.from('cte_nf_detalhe').insert(slim));
    }
    if (e2) throw e2;
  }
  return upload;
}

async function autosaveToCloud() {
  if (!currentNFs.length || !currentSummary || isSaving) return;
  isSaving = true;
  updateSaveStatus('A guardar na cloud...');
  try {
    const upload = await insertUploadAndDetails();
    currentUploadId = upload.id;
    const when = new Date(upload.criado_em || Date.now()).toLocaleString('pt-BR');
    updateSaveStatus('Guardado · visível para todos · ' + when, true);
    fteToast('Análise guardada automaticamente — visível para toda a equipa.');
  } catch (err) {
    console.error(err);
    updateSaveStatus('Erro ao guardar na cloud', false);
    fteToast('Erro ao guardar: ' + (err.message || err));
  } finally {
    isSaving = false;
  }
}

async function fetchAllDetalhes(uploadId) {
  const all = [];
  const pageSize = 1000;
  let from = 0;
  while (true) {
    const { data, error } = await sb.from('cte_nf_detalhe')
      .select('*')
      .eq('upload_id', uploadId)
      .order('diferenca', { ascending: false })
      .range(from, from + pageSize - 1);
    if (error) throw error;
    if (!data?.length) break;
    all.push(...data);
    if (data.length < pageSize) break;
    from += pageSize;
  }
  return all;
}

async function loadUploadById(id, uploadMeta) {
  let upload = uploadMeta;
  if (!upload) {
    const { data, error } = await sb.from('cte_nf_uploads').select('*').eq('id', id).single();
    if (error || !data) return false;
    upload = data;
  }

  updateSaveStatus('A carregar todas as NFs...');
  let detalhes;
  try {
    detalhes = await fetchAllDetalhes(id);
  } catch (e2) {
    updateSaveStatus('Erro ao carregar detalhe', false);
    return false;
  }
  if (!detalhes.length) return false;

  currentNFs = nfsFromDetalhe(detalhes);
  if (Object.keys(sapNfMap).length) {
    currentNFs = applySapToList(currentNFs.map(nf => {
      delete nf.mesRef;
      delete nf.dtRef;
      return nf;
    }));
  }
  currentSummary = computeSummary(currentNFs, upload.ficheiro_nome);

  const savedNf = upload.total_nf || 0;
  const savedPago = num(upload.total_pago);
  const loadedNf = currentNFs.length;
  const loadedPago = currentSummary.totalPago;
  let loadWarning = '';

  if (savedNf > loadedNf) {
    loadWarning = `<strong>Dados incompletos na cloud:</strong> guardadas ${savedNf} NFs, carregadas ${loadedNf}. ` +
      `Total pago guardado: ${fmtMoney(savedPago)} · calculado agora: ${fmtMoney(loadedPago)}. ` +
      `Carrega o Excel de novo para actualizar.`;
  } else if (Math.abs(savedPago - loadedPago) > 1) {
    loadWarning = `<strong>Totais diferentes:</strong> total pago guardado ${fmtMoney(savedPago)} vs recalculado ${fmtMoney(loadedPago)}. ` +
      `Pode haver linhas em falta — volta a carregar o ficheiro Excel.`;
  }
  currentSummary.loadWarning = loadWarning;

  currentUploadId = upload.id;
  cteAnalysisSource = 'conciliacao';
  activeSubPanel = null;
  selectedMonth = '';
  renderAll();
  showResultsView();
  const when = new Date(upload.criado_em).toLocaleString('pt-BR');
  updateSaveStatus(`Análise partilhada · ${loadedNf} NFs · ${when}`, true);
  if (loadWarning) fteToast('Atenção: dados possivelmente incompletos — ver aviso amarelo.');
  return true;
}

async function loadLatestFromCloud() {
  updateSaveStatus('A carregar análise partilhada...');
  const { data, error } = await sb.from('cte_nf_uploads')
    .select('*').order('criado_em', { ascending: false }).limit(1);
  if (error) {
    updateSaveStatus('Erro ao carregar', false);
    return false;
  }
  if (!data?.length) {
    updateSaveStatus('');
    return false;
  }
  const ok = await loadUploadById(data[0].id, data[0]);
  if (ok) fteToast('Análise partilhada carregada.');
  return ok;
}

async function restoreSapFromRec(sapRec, silent = false) {
  if (!sapRec?.file_data || typeof base64ToArrayBuffer !== 'function') return false;
  try {
    await fteYield('A descodificar ZFACT…');
    const b64 = sapRec.file_data;
    sapRec.file_data = null;
    const buf = base64ToArrayBuffer(b64);
    const ok = await processArrayBufferSapAsync(buf, sapRec.file_name || 'sap.xlsx', {
      silent: true,
      yieldLabel: 'A processar ZFACT…'
    });
    if (!ok) return false;
    fteSapFileName = sapRec.file_name || '';
    // Do not retain ArrayBuffer after map is built — re-upload from disk if needed.
    fteSapBuffer = null;
    setSapZoneLoaded(sapRec.file_name);
    return true;
  } catch (err) {
    console.error('[fretes] restore sap', err);
    if (!silent) fteToastError('Erro ao carregar SAP NF guardado.');
    return false;
  }
}

async function restoreCteFromRec(cteRec, sapRec, silent = false) {
  if (!cteRec?.file_data) return false;
  const co = fteCompany();
  if (_fteLoadedCompany === co && currentNFs.length && lastCtePack?.fileName === cteRec.file_name) {
    if (sapRec?.file_data && !isSapLoaded()) {
      await restoreSapFromRec(sapRec, silent);
    }
    return true;
  }
  _fteSkipAutosave = true;
  try {
    // Keep accumulated sapNfMap — Excel merge updates keys; does not wipe history
    await fteYield('A descodificar CT-e…');
    const b64 = cteRec.file_data;
    cteRec.file_data = null;
    const cteBuf = base64ToArrayBuffer(b64);
    const ok = await processArrayBufferCteAsync(cteBuf, cteRec.file_name || 'cte.xlsx', { silent: true, autosave: false });
    if (!ok) return false;
    if (sapRec?.file_data && !isSapLoaded()) {
      await restoreSapFromRec(sapRec, silent);
    }
    _fteLoadedCompany = co;
    fteCteFileName = cteRec.file_name || '';
    fteCteBuffer = null; // parsed currentNFs are enough
    setCteZoneLoaded(cteRec.file_name);
    if (sapRec?.file_name && !fteSapFileName) {
      fteSapFileName = sapRec.file_name || '';
      setSapZoneLoaded(sapRec.file_name);
    }
    return true;
  } catch (err) {
    console.error('[fretes] restore cte', err);
    if (!silent) fteToastError('Erro ao carregar ficheiros guardados.');
    return false;
  } finally {
    _fteSkipAutosave = false;
  }
}

async function ensureCteLoadedForQz(silent = true) {
  if (currentNFs.length) return true;
  if (typeof fetchExcelFiles !== 'function' || typeof base64ToArrayBuffer !== 'function') return false;
  try {
    const meta = await fetchExcelFiles([fteCteSlot(), fteSapMapSlot(), fteSapSlot()]);
    await loadSavedSapNfMap(meta, { silent: true });
    const ok = await restoreCteFromRec(meta[fteCteSlot()], meta[fteSapSlot()], silent);
    if (ok) refreshQuinzenalCompare();
    return ok;
  } catch (err) {
    console.error('[fretes] ensure cte for qz', err);
    return false;
  }
}

async function loadSavedFretesFiles(silent = false) {
  if (_fteLoadSavedPromise) return _fteLoadSavedPromise;
  // Cap concurrent restores with stock — wait if stock is mid-load.
  if (typeof window._stockLoadPromise !== 'undefined' && window._stockLoadPromise) {
    try { await window._stockLoadPromise; } catch (_) {}
  }
  // Always show spinner — silent only suppresses success toasts (menus must not look "dead").
  fteSetProcessing(true, 'A carregar Fretes da cloud…');
  _fteLoadSavedPromise = _loadSavedFretesFilesImpl(silent).finally(() => {
    fteSetProcessing(false);
    _fteLoadSavedPromise = null;
    try { window._fteLoadSavedPromise = null; } catch (_) {}
  });
  try { window._fteLoadSavedPromise = _fteLoadSavedPromise; } catch (_) {}
  return _fteLoadSavedPromise;
}

async function _loadSavedFretesFilesImpl(silent = false) {
  if (typeof fetchExcelFiles !== 'function' || typeof base64ToArrayBuffer !== 'function') {
    if (!silent) fteToastError('Persistência Excel indisponível — recarrega a página.');
    return false;
  }
  if (typeof syncCompanyDepotMap === 'function') syncCompanyDepotMap();
  const co = fteCompany();
  console.log('[fretes] load saved', co);

  // Instant path: already have QZ/CTE for this company — paint from RAM.
  // SAP map alone (Armazém / partial) must NOT skip quinzenal restore.
  if (_fteLoadedCompany === co && fteHasAnalysisData()) {
    syncQzUploadZone();
    updateQzFileNote();
    updateFretesFileStatus();
    renderLoadCoverageSummary();
    const activeTab = document.querySelector('.fte-tab.active')?.dataset?.tab;
    if (activeTab === 'analise-cte') renderCteSubPanels();
    else if (activeTab === 'analise-b2c') renderB2cAnalysisTab();
    else if (activeTab === 'cte-vs-qz') renderB2bCompareTab();
    else if (activeTab === 'resumo-total') renderResumoTotal();
    refreshCustoUnilogIfVisible();
    return true;
  }

  await fteYield('A aguardar ligação…');
  const dbOk = await fteWaitForDb(8000);
  if (!dbOk) {
    console.warn('[fretes] load saved aborted — Supabase not ready', co);
    if (!silent) fteToastError('Cloud offline / a ligar — espera «Supabase OK» e tenta de novo.');
    return false;
  }

  let meta = {};
  try {
    await fteYield('A descarregar Fretes (QZ + mapa)…');
    // Slim first: quinzenal pack JSON + SAP NF map only — NOT ZFACT/CTE Excel binaries.
    meta = await fetchExcelFiles([fteQuinzenalSlot(), fteSapMapSlot()]);
  } catch (err) {
    console.error('[fretes] load saved', co, err);
    if (!silent) {
      if (typeof isExcelFilesTableMissing === 'function' && isExcelFilesTableMissing(err)) {
        fteToastError('Tabela logistica_excel_files em falta no Supabase.');
      } else {
        fteToastError('Erro ao carregar ficheiros: ' + (err.message || err));
      }
    }
    return false;
  }

  let cteRec = null;
  let sapRec = null;
  const hadCteInMem = _fteLoadedCompany === co && currentNFs.length;
  const hadQzInMem = !!quinzenalPack?.files?.length;

  // 1) Restore quinzenais (structured JSON) — preferred source of truth
  await fteYield('A restaurar quinzenais…');
  const qzLoaded = await loadSavedQuinzenalPack(silent, meta, { deferCompare: true, skipRender: true });
  if (qzLoaded) console.log('[fretes] restore quinzenal', co, quinzenalPack?.files?.length || 0, 'files');
  // Free QZ base64 slot from meta after parse
  if (meta[fteQuinzenalSlot()]) meta[fteQuinzenalSlot()].file_data = null;

  // 2) Build Análise CT-e from QZ B2B (Conciliacao removed) — yield before heavy sync
  let qzCteBuilt = false;
  if (!hadCteInMem && !hasConciliacaoCte() && quinzenalPack?.b2bRows?.length) {
    await fteYield('A montar Análise CT-e…');
    qzCteBuilt = await processCteAnalysisFromQuinzenalAsync({ switchTab: false });
    console.log('[fretes] restore cte-from-qz', co, 'ok', qzCteBuilt, 'nfs', currentNFs.length);
  }

  // 3) Legacy CTE Excel only if QZ path did not produce analysis
  let cteOk = !!(qzCteBuilt || currentNFs.length);
  if (!cteOk) {
    try {
      await fteYield('A descarregar CT-e legado…');
      const legacy = await fetchExcelFiles([fteCteSlot()]);
      cteRec = legacy[fteCteSlot()];
      meta[fteCteSlot()] = cteRec;
    } catch (e) { console.warn('[fretes] cte fetch', e); }
  }
  if (!cteOk && cteRec?.file_data) {
    await fteYield('A processar CT-e Excel…');
    const cteDataLen = cteRec.file_data.length;
    cteOk = await restoreCteFromRec(cteRec, null, silent);
    cteRec.file_data = null;
    console.log('[fretes] restore cte', co, cteRec.file_name, 'dataLen', cteDataLen, 'ok', cteOk);
  } else if (!cteOk && hadCteInMem) {
    cteOk = true;
    console.log('[fretes] restore cte from memory', co, currentNFs.length, 'NFs');
  } else if (!cteOk && cteRec?.file_name && !cteRec?.file_data) {
    console.warn('[fretes] cte metadata only', co, cteRec.file_name);
    if (!silent) {
      fteToastError('CT-e guardado só com nome — clica Processar e Guardar para persistir o ficheiro.');
    }
  }

  // 4) ZFACT / SAP NF map — prefer accumulated JSON (fast). Re-parse Excel when map
  //    missing/stale OR filename range (ex. "jan a Agosto") claims months absent from map dates.
  await fteYield('A restaurar mapa SAP…');
  let mapHad = await loadSavedSapNfMap(meta, { silent: true });
  if (meta[fteSapMapSlot()]) meta[fteSapMapSlot()].file_data = null;

  // Resolve ZFACT Excel filename early (for coverage hints + gap detection)
  try {
    const fetchMeta = typeof fetchExcelFileMeta === 'function' ? fetchExcelFileMeta : null;
    if (fetchMeta) {
      const labels = await fetchMeta([fteSapSlot()]);
      sapRec = labels[fteSapSlot()];
      if (sapRec?.file_name) {
        fteSapFileName = sapRec.file_name || fteSapFileName;
        setSapZoneLoaded(sapRec.file_name);
      }
    }
  } catch (_) {}

  const mapMonths = new Set();
  Object.keys(sapNfMap || {}).forEach(k => {
    const mk = mesKeyFromDtEmissao(sapNfMap[k]?.dtEmissao);
    if (mk) mapMonths.add(mk);
  });
  const fileMonths = collectZfactFilenameMonthHints();
  const mapMissesFileMonths = [...fileMonths].some(m => !mapMonths.has(m));
  // Never auto-force multi-MB ZFACT Excel re-parse on restore (Chrome OOM → white tab).
  // Coverage gaps: keep map + warn; user re-processes ZFACT from Carregamento when needed.
  if (mapHad && mapMissesFileMonths) {
    console.warn('[fretes] sap map months incomplete vs filename (keeping map, no auto re-parse)', {
      mapMonths: [...mapMonths].sort(),
      fileMonths: [...fileMonths].sort(),
      file: fteSapFileName
    });
    if (!silent) {
      fteToast('Cobertura ZFACT incompleta vs nome do ficheiro — em Carregamento, processa de novo o ZFACT.');
    }
  }

  if (mapHad && Object.keys(sapNfMap).length) {
    // Ensure RAM map is repaired even if loadSavedSapNfMap ran earlier in session.
    const nFix = repairSapNfMapInPlace(sapNfMap);
    if (nFix > 0) {
      try { await persistSapNfMap(sapNfMap, { silent: true }); fteSapMapNeedsReprocess = false; }
      catch (e) { console.warn('[fretes] persist repaired map on restore', e); }
    }
    if (qzCteBuilt || cteAnalysisSource === 'quinzenal' || currentNFs.length) {
      await fteYield('A cruzar SAP…');
      reEnrichAfterSapLoad();
      if (typeof window.refreshArmazemSapValidation === 'function') {
        try { window.refreshArmazemSapValidation(); } catch (armErr) {
          console.warn('[SAP] refreshArmazemSapValidation', armErr);
        }
      }
    }
    if (fteSapMapNeedsReprocess && !silent) {
      fteToast('Mapa SAP ainda incompleto em algumas NFs (sem Val.total) — Processar ZFACT completo se quiseres regravar AP.');
    }
    console.log('[fretes] restore sap from map JSON', co, 'mapSize', Object.keys(sapNfMap).length,
      'skippedZfactExcel', true, 'mapMissesFileMonths', mapMissesFileMonths,
      'needsReprocess', fteSapMapNeedsReprocess, 'repaired', nFix);
  } else {
    // Map missing only — fetch ZFACT Excel (guarded; failures must not blank the app)
    try {
      await fteYield('A descarregar ZFACT…');
      const sapMeta = await fetchExcelFiles([fteSapSlot()]);
      sapRec = sapMeta[fteSapSlot()];
    } catch (e) { console.warn('[fretes] sap excel fetch', e); }
    if (sapRec?.file_data) {
      try {
        await fteYield('A processar ZFACT…');
        const sapOnlyOk = await restoreSapFromRec(sapRec, silent);
        sapRec.file_data = null;
        console.log('[fretes] restore sap', co, sapRec.file_name, 'ok', sapOnlyOk,
          'mapHad', mapHad, 'mapSize', Object.keys(sapNfMap).length);
        const nAfter = Object.keys(sapNfMap).length;
        if (nAfter > 0) {
          await persistSapNfMap(sapNfMap, { silent: true });
        }
      } catch (e) {
        console.error('[fretes] restore sap excel failed (kept UI alive)', e);
        try { sapRec.file_data = null; } catch (_) {}
        if (!silent) fteToastError('Falha ao processar ZFACT guardado — tenta Processar de novo em Carregamento.');
      }
    } else if (Object.keys(sapNfMap).length && (qzCteBuilt || cteAnalysisSource === 'quinzenal' || currentNFs.length)) {
      await fteYield('A cruzar SAP…');
      reEnrichAfterSapLoad();
      if (typeof window.refreshArmazemSapValidation === 'function') {
        try { window.refreshArmazemSapValidation(); } catch (armErr) {
          console.warn('[SAP] refreshArmazemSapValidation', armErr);
        }
      }
    }
  }
  // Free any leftover Fretes Excel buffers after restore
  fteCteBuffer = null;
  fteSapBuffer = null;
  if (quinzenalPack?.fileBinaries) delete quinzenalPack.fileBinaries;

  await fteYield('A calcular confrontos…');
  refreshQuinzenalCompare();
  syncQzUploadZone();
  updateQzFileNote();
  updateFretesFileStatus(meta);
  renderLoadCoverageSummary();

  const freshCte = !!(cteOk && !hadCteInMem && !qzCteBuilt && cteRec);
  const freshQz = !!(qzLoaded && !hadQzInMem);

  if (!cteOk && !qzCteBuilt && !silent && !qzLoaded && !cteRec) {
    fteToastError('Sem ficheiros CT-e guardados para ' + co + '.');
  } else if ((freshCte || freshQz || qzCteBuilt) && !silent) {
    const parts = [];
    if (freshCte) parts.push('CT-e');
    if (freshQz) parts.push('quinzenais');
    if (qzCteBuilt && !freshCte) parts.push('Análise CT-e via QZ');
    fteToast('Ficheiros fretes restaurados: ' + parts.join(', ') + ' (' + co + ').');
  }

  // Only mark company loaded when analysis data exists — SAP map alone must not block retry.
  if (qzLoaded || cteOk || qzCteBuilt || fteHasAnalysisData()) {
    fteMarkCompanyLoaded();
  }

  await fteYield('A desenhar tabelas…');
  const activeTab = document.querySelector('.fte-tab.active')?.dataset?.tab;
  if (activeTab === 'analise-cte') renderCteSubPanels();
  else if (activeTab === 'analise-b2c') renderB2cAnalysisTab();
  else if (activeTab === 'cte-vs-qz') renderB2bCompareTab();
  else if (activeTab === 'resumo-total') renderResumoTotal();
  refreshCustoUnilogIfVisible();
  return cteOk || qzLoaded || qzCteBuilt || fteHasAnalysisData() || isSapLoaded();
}

// ── QUINZENAL UNILOG (B2B diff vs CT-e · B2C vendas/fretes) ──

const QZ_MESES = {
  janeiro: 1, fevereiro: 2, marco: 3, março: 3, abril: 4, maio: 5, junho: 6,
  julho: 7, agosto: 8, setembro: 9, outubro: 10, novembro: 11, dezembro: 12
};

const QZ_B2B_ALIASES = {
  nf: ['nota fiscal', 'nf', 'n nota fiscal'],
  valorNF: ['valor nf', 'valor da nf', 'vl nf'],
  numCte: ['num. cte', 'num cte', 'numero cte', 'numeros cte', 'n cte', 'numero do cte', 'nº cte'],
  pago: ['total fatura rev.', 'total fatura rev', 'total fatura', 'valor fatura'],
  transportador: ['transportador', 'transportadora'],
  dtNF: [
    'dt nf', 'data nf', 'datanf', 'data da nf', 'data nota fiscal', 'dt nota fiscal',
    'data emissao', 'data de emissao', 'data emissão', 'data de emissão',
    'dt emissao', 'dt emissão', 'emissao nf', 'emissão nf',
    'data emissao nf', 'data emissão nf'
  ],
  dtCte: ['dt cte', 'dt ct-e', 'dt. cte', 'dt. ct-e', 'data cte', 'data ct-e', 'data cte-e', 'dt cte-e', 'data do cte', 'data do ct-e', 'data cte e'],
  devolucao: ['devolucao', 'devolução', 'e devolucao', 'retorno'],
  modalidade: ['modalidade', 'modal', 'mod'],
  tipoOp: ['tipo operacao', 'tipo operação', 'tipo op', 'operacao'],
  peso: ['peso', 'peso kg', 'peso (kg)'],
  destinatario: ['destinatario', 'destinatário']
};

const QZ_B2C_ALIASES = {
  pedido: ['numero pedido', 'número pedido', 'n pedido', 'pedido'],
  nf: ['numeros nf', 'números nf', 'nota fiscal', 'n nota fiscal'],
  chaveNF: ['chaves acesso nf', 'chave acesso nf', 'chave nf', 'chave da nota fiscal', 'numeros nf', 'números nf'],
  dtColeta: ['data da coleta', 'data coleta'],
  dtNF: ['datas de emissao nf', 'datas de emissão nf', 'data emissao nf', 'data ct-e', 'data cte'],
  numCte: ['numeros ct-e', 'numeros cte', 'números ct-e', 'numero ct-e', 'numero cte'],
  transportador: ['nome transportadora', 'transportadora', 'transportador'],
  destinatario: ['nome destinatario', 'nome destinatário', 'destinatario', 'destinatário'],
  valorProdutos: ['valor produtos', 'valor produto', 'vl produtos'],
  valorNF: ['total nf', 'valor total nf'],
  pago: ['valor fatura', 'total fatura', 'frete'],
  zona: ['zona entrega', 'zona de entrega', 'regiao entrega', 'região entrega']
};

function findQzField(row, field, aliasesMap) {
  const aliases = aliasesMap[field];
  const entries = Object.entries(row);
  for (const alias of aliases) {
    const hit = entries.find(([k]) => normCol(k) === alias);
    if (hit && hit[1] !== null && hit[1] !== undefined && hit[1] !== '') return hit[1];
  }
  for (const alias of aliases) {
    const hit = entries.find(([k]) => {
      const nk = normCol(k);
      return nk.includes(alias) || alias.includes(nk);
    });
    if (hit && hit[1] !== null && hit[1] !== undefined && hit[1] !== '') return hit[1];
  }
  return null;
}

function nfFromNfeKey(key) {
  const s = String(key || '').replace(/\D/g, '');
  if (s.length >= 44) return s.slice(25, 34).replace(/^0+/, '') || '0';
  if (s.length >= 9) return s.slice(-9).replace(/^0+/, '') || '0';
  return s.replace(/^0+/, '') || '';
}

function parseQuinzenalFileName(name) {
  const canal = /B2B/i.test(name) ? 'B2B' : (/B2C/i.test(name) ? 'B2C' : null);
  const quinzena = /1\s*ª?\s*Q|1aQ/i.test(name) ? 1 : (/2\s*ª?\s*Q|2aQ/i.test(name) ? 2 : null);
  let mes = null;
  let ano = null;
  const baseName = String(name || '').replace(/\.xlsx?$/i, '');
  const ym = baseName.match(/(\d{4})\s*$/) || String(name || '').match(/(\d{4})/);
  if (ym) ano = parseInt(ym[1], 10);
  const norm = baseName.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  for (const [k, v] of Object.entries(QZ_MESES)) {
    const kk = k.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (norm.includes(kk)) { mes = v; break; }
  }
  const mesKey = ano && mes ? `${ano}-${String(mes).padStart(2, '0')}` : null;
  const mesLabel = mesKey ? fmtMesLabel(mesKey) : null;
  const quinzenaKey = ano && mes && quinzena ? `${mesKey}-Q${quinzena}` : null;
  const qLabel = quinzena && mes && ano ? `${quinzena}ªQ ${MESES_PT[mes - 1]}/${ano}` : name;
  return { canal, quinzena, mes, ano, mesKey, mesLabel, quinzenaKey, quinzenaLabel: qLabel, fileName: name };
}

function mesKeyFromQuinzenaKey(qk) {
  if (!qk || qk === 'sem-quinzena') return 'sem-mes';
  const parts = qk.split('-');
  if (parts.length >= 2 && /^\d{4}$/.test(parts[0]) && /^\d{2}$/.test(parts[1])) return `${parts[0]}-${parts[1]}`;
  return 'sem-mes';
}

function resolveB2cRowMesKey(r, fileMeta) {
  const d = parseSapBrDate(r?.dtColeta) || parseSapBrDate(r?.dtNF);
  if (d) {
    const k = monthKey(d);
    if (k && k !== 'sem-data') return k;
  }
  if (r?.mesKey && r.mesKey !== 'sem-mes') return r.mesKey;
  const meta = fileMeta || r;
  if (meta?.mesKey && meta.mesKey !== 'sem-mes') return meta.mesKey;
  const fromQz = mesKeyFromQuinzenaKey(r?.quinzenaKey || meta?.quinzenaKey);
  if (fromQz && fromQz !== 'sem-mes') return fromQz;
  if (meta?.mes && meta?.ano) return `${meta.ano}-${String(meta.mes).padStart(2, '0')}`;
  return 'sem-mes';
}

function resolveB2cRowMesLabel(r, mesKey, fileMeta) {
  if (mesKey && mesKey !== 'sem-mes') return fmtMesLabel(mesKey);
  const meta = fileMeta || r;
  if (meta?.mesLabel) return meta.mesLabel;
  return 'Sem mês';
}

function enrichB2cRowMesFields(r, fileMeta) {
  const mesKey = resolveB2cRowMesKey(r, fileMeta);
  r.mesKey = mesKey;
  r.mesLabel = resolveB2cRowMesLabel(r, mesKey, fileMeta);
  return r;
}

function quinzenaDateRange(meta) {
  if (!meta?.ano || !meta?.mes || !meta?.quinzena) return null;
  const start = new Date(meta.ano, meta.mes - 1, meta.quinzena === 1 ? 1 : 16);
  start.setHours(0, 0, 0, 0);
  const end = meta.quinzena === 1
    ? new Date(meta.ano, meta.mes - 1, 15, 23, 59, 59, 999)
    : new Date(meta.ano, meta.mes, 0, 23, 59, 59, 999);
  return { start, end };
}

function monthDateRange(ano, mes) {
  if (!ano || !mes) return null;
  const start = new Date(ano, mes - 1, 1);
  start.setHours(0, 0, 0, 0);
  const end = new Date(ano, mes, 0, 23, 59, 59, 999);
  return { start, end };
}

function detectCanalFromWorkbook(wb) {
  let b2bScore = 0;
  let b2cScore = 0;
  for (const name of wb.SheetNames) {
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null });
    for (const row of raw.slice(0, 30)) {
      if (!Array.isArray(row)) continue;
      if (isQzB2BHeader(row)) b2bScore += 3;
      if (isQzB2CHeader(row)) b2cScore += 3;
    }
  }
  if (b2bScore > b2cScore) return 'B2B';
  if (b2cScore > b2bScore) return 'B2C';
  return null;
}

function isQzB2BHeader(cells) {
  const n = cells.map(c => normCol(c));
  return n.some(c => c === 'nota fiscal' || c === 'nf') && n.some(c => c.includes('total fatura'));
}

function isQzB2CHeader(cells) {
  const n = cells.map(c => normCol(c));
  const hasFrete = n.some(c => c.includes('valor fatura') || c.includes('valor produto') || c.includes('valor produtos'));
  const hasId = n.some(c => c.includes('total nf') || c.includes('pedido') || c.includes('chave') || c.includes('nota fiscal'));
  return hasFrete && hasId;
}

function normalizeQzB2BRow(row) {
  const dtCteRaw = findQzField(row, 'dtCte', QZ_B2B_ALIASES);
  const dtNFRaw = findQzField(row, 'dtNF', QZ_B2B_ALIASES);
  // Coerce Excel Date / serial / BR string at ingest so restore + month bucketing stay stable
  const dtCte = persistDateValue(dtCteRaw) || dtCteRaw;
  const dtNF = persistDateValue(dtNFRaw) || dtNFRaw;
  return {
    nf: findQzField(row, 'nf', QZ_B2B_ALIASES),
    valorNF: findQzField(row, 'valorNF', QZ_B2B_ALIASES),
    numCte: findQzField(row, 'numCte', QZ_B2B_ALIASES),
    pago: findQzField(row, 'pago', QZ_B2B_ALIASES),
    transportador: findQzField(row, 'transportador', QZ_B2B_ALIASES),
    dtNF,
    dtCte,
    devolucao: findQzField(row, 'devolucao', QZ_B2B_ALIASES),
    modalidade: findQzField(row, 'modalidade', QZ_B2B_ALIASES),
    tipoOp: findQzField(row, 'tipoOp', QZ_B2B_ALIASES),
    peso: findQzField(row, 'peso', QZ_B2B_ALIASES),
    destinatario: findQzField(row, 'destinatario', QZ_B2B_ALIASES)
  };
}

function normalizeQzB2CRow(row) {
  const chave = findQzField(row, 'chaveNF', QZ_B2C_ALIASES);
  const pedido = findQzField(row, 'pedido', QZ_B2C_ALIASES);
  const nfDirect = findQzField(row, 'nf', QZ_B2C_ALIASES);
  const nfRaw = chave && /^\d{1,12}$/.test(String(chave).trim()) ? String(chave).trim() : null;
  let nf = nfRaw || (nfDirect ? String(nfDirect).trim() : null) || nfFromNfeKey(chave);
  if (!nf && pedido) nf = String(pedido).trim();
  return {
    pedido, nf, chaveNF: chave,
    dtColeta: findQzField(row, 'dtColeta', QZ_B2C_ALIASES),
    dtNF: findQzField(row, 'dtNF', QZ_B2C_ALIASES),
    numCte: findQzField(row, 'numCte', QZ_B2C_ALIASES),
    transportador: findQzField(row, 'transportador', QZ_B2C_ALIASES),
    destinatario: findQzField(row, 'destinatario', QZ_B2C_ALIASES),
    valorProdutos: findQzField(row, 'valorProdutos', QZ_B2C_ALIASES),
    valorNF: findQzField(row, 'valorNF', QZ_B2C_ALIASES),
    pago: findQzField(row, 'pago', QZ_B2C_ALIASES),
    zona: findQzField(row, 'zona', QZ_B2C_ALIASES)
  };
}

function parseQzSheetRows(sheet, canal) {
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  if (!raw.length) return { rows: [], headers: [] };
  const headerFn = canal === 'B2B' ? isQzB2BHeader : isQzB2CHeader;
  const normFn = canal === 'B2B' ? normalizeQzB2BRow : normalizeQzB2CRow;
  let headerIdx = raw.findIndex(row => Array.isArray(row) && headerFn(row));
  if (headerIdx < 0) return { rows: [], headers: [] };
  const headerRow = raw[headerIdx] || [];
  const headers = headerRow.map(h => String(h || '').trim()).filter(Boolean);
  const rows = [];
  for (let i = headerIdx + 1; i < raw.length; i++) {
    const line = raw[i];
    if (!Array.isArray(line) || line.every(c => c === null || c === undefined || c === '')) continue;
    const obj = {};
    headerRow.forEach((h, j) => { if (h) obj[h] = line[j] ?? null; });
    const row = normFn(obj);
    if (canal === 'B2B') {
      if (!row.nf || String(row.nf).trim() === '') continue;
      const nfStr = String(row.nf).trim().toLowerCase();
      if (nfStr.includes('total') || nfStr === 'na') continue;
    } else {
      if (!row.pago && !row.valorNF && !row.pedido) continue;
      if (String(row.pedido || '').toLowerCase().includes('pedido')) continue;
    }
    rows.push(row);
  }
  return { rows, headers };
}

function loadQuinzenalFromWorkbook(wb, fileMeta) {
  const canal = fileMeta.canal || (/B2B/i.test(fileMeta.fileName) ? 'B2B' : 'B2C');
  let best = { rows: [], headers: [], valid: 0, sheetName: wb.SheetNames[0] };
  for (const name of wb.SheetNames) {
    const parsed = parseQzSheetRows(wb.Sheets[name], canal);
    if (parsed.rows.length > best.valid) best = { ...parsed, sheetName: name, valid: parsed.rows.length };
  }
  return { ...best, meta: fileMeta, canal };
}

function aggregateQzB2BRows(rows, fileMeta) {
  const byNf = {};
  rows.forEach(r => {
    const key = normNFKey(r.nf);
    if (!key) return;
    if (!byNf[key]) {
      byNf[key] = {
        nf: String(r.nf).trim(), nfKey: key, valorNF: 0, pago: 0, nCteSet: new Set(),
        transportador: r.transportador || '', destinatario: r.destinatario || '',
        modalidade: r.modalidade || '',
        temDevolucao: false, ctes: [], dtNF: null, dtNFQz: null,
        mesKey: fileMeta.mesKey || mesKeyFromQuinzenaKey(fileMeta.quinzenaKey),
        mesLabel: fileMeta.mesLabel || (fileMeta.mesKey ? fmtMesLabel(fileMeta.mesKey) : ''),
        quinzenaKey: fileMeta.quinzenaKey, quinzenaLabel: fileMeta.quinzenaLabel, fileName: fileMeta.fileName
      };
    }
    const g = byNf[key];
    if (num(r.valorNF) > g.valorNF) g.valorNF = num(r.valorNF);
    const cteKey = r.numCte != null && String(r.numCte).trim() !== '' ? String(r.numCte).trim() : null;
    // Duplicate QZ rows for the same CT-e — do not double-count pago / CT-e list.
    if (cteKey && g.nCteSet.has(cteKey)) return;
    if (cteKey) g.nCteSet.add(cteKey);
    g.pago += num(r.pago);
    if (r.transportador && g.transportador === '') g.transportador = r.transportador;
    if (r.modalidade && !g.modalidade) g.modalidade = r.modalidade;
    if (r.dtNF) {
      const kept = pickEarliestDateValue(g.dtNF, r.dtNF);
      g.dtNF = kept;
      g.dtNFQz = kept;
    }
    const isDev = isDevolucaoFlag(r.devolucao);
    g.ctes.push({
      numCte: cteKey,
      dtCte: persistDateValue(r.dtCte) || r.dtCte || null,
      pago: num(r.pago),
      devolucao: isDev,
      tipoOp: r.tipoOp || '',
      peso: num(r.peso)
    });
    if (isDev) g.temDevolucao = true;
  });
  return Object.values(byNf).map(g => {
    const nCte = g.nCteSet.size || g.ctes.length || 1;
    return { ...g, nCte, nCteSet: undefined };
  });
}

function expandQzB2bToCteLines(b2bRows) {
  const lines = [];
  (b2bRows || []).forEach(nf => {
    const qzMesKey = nf.mesKey || mesKeyFromQuinzenaKey(nf.quinzenaKey) || null;
    const qzMesLabel = nf.mesLabel || (qzMesKey && qzMesKey !== 'sem-mes' ? fmtMesLabel(qzMesKey) : '');
    const ctes = Array.isArray(nf.ctes) ? nf.ctes : [];
    if (ctes.length) {
      ctes.forEach(c => {
        lines.push({
          nf: nf.nf,
          valorNF: nf.valorNF,
          transportador: nf.transportador || '-',
          modalidade: nf.modalidade || '-',
          dtNF: persistDateValue(nf.dtNF) || nf.dtNF || null,
          dtNFQz: persistDateValue(nf.dtNFQz || nf.dtNF) || nf.dtNFQz || nf.dtNF || null,
          numCte: c.numCte,
          qtdCte: nf.nCte || ctes.length,
          dtCte: persistDateValue(c.dtCte) || c.dtCte || null,
          pago: num(c.pago),
          devolucao: c.devolucao ? 'sim' : 'nao',
          tipoOp: c.tipoOp || '',
          peso: num(c.peso),
          mesKey: qzMesKey,
          mesLabel: qzMesLabel,
          qzMesKey
        });
      });
      return;
    }
    // Legacy pack without per-CT-e lines — one synthetic row per NF
    lines.push({
      nf: nf.nf,
      valorNF: nf.valorNF,
      transportador: nf.transportador || '-',
      modalidade: nf.modalidade || '-',
      dtNF: persistDateValue(nf.dtNF) || nf.dtNF || null,
      dtNFQz: persistDateValue(nf.dtNFQz || nf.dtNF) || nf.dtNFQz || nf.dtNF || null,
      numCte: null,
      qtdCte: nf.nCte || 1,
      dtCte: null,
      pago: num(nf.pago),
      devolucao: nf.temDevolucao ? 'sim' : 'nao',
      tipoOp: '',
      peso: 0,
      mesKey: qzMesKey,
      mesLabel: qzMesLabel,
      qzMesKey
    });
  });
  return lines;
}

/**
 * Build Análise CT-e boards from quinzenal B2B when Conciliacao CT-e×NF is absent.
 * One QZ line ≈ one CT-e (pago / devolução / Num. CTE / Dt CTE).
 */
function processCteAnalysisFromQuinzenal(opts = {}) {
  const b2b = quinzenalPack?.b2bRows || [];
  if (!b2b.length) return false;
  const lines = expandQzB2bToCteLines(b2b);
  if (!lines.length) return false;
  const nFiles = quinzenalPack?.files?.filter(f => f.canal === 'B2B').length || 0;
  const label = nFiles
    ? `Quinzenais B2B (${nFiles} ficheiro${nFiles !== 1 ? 's' : ''})`
    : 'Quinzenais B2B';
  const headers = ['Nota Fiscal', 'Valor NF', 'Num. CTE', 'Dt CTE', 'Total Fatura Rev.', 'Devolução', 'Transportador'];
  _fteSkipAutosave = true;
  try {
    processRows(lines, label, 'B2B QZ', headers, {
      source: 'quinzenal',
      switchTab: opts.switchTab !== false,
      autosave: false
    });
  } finally {
    _fteSkipAutosave = false;
  }
  if (currentNFs.length) fteMarkCompanyLoaded();
  return currentNFs.length > 0;
}

/** Async variant — yields so the UI can paint a spinner during large QZ→CT-e rebuilds. */
async function processCteAnalysisFromQuinzenalAsync(opts = {}) {
  const b2b = quinzenalPack?.b2bRows || [];
  if (!b2b.length) return false;
  await fteYield('A expandir linhas QZ…');
  const lines = expandQzB2bToCteLines(b2b);
  if (!lines.length) return false;
  const nFiles = quinzenalPack?.files?.filter(f => f.canal === 'B2B').length || 0;
  const label = nFiles
    ? `Quinzenais B2B (${nFiles} ficheiro${nFiles !== 1 ? 's' : ''})`
    : 'Quinzenais B2B';
  const headers = ['Nota Fiscal', 'Valor NF', 'Num. CTE', 'Dt CTE', 'Total Fatura Rev.', 'Devolução', 'Transportador'];
  _fteSkipAutosave = true;
  try {
    await processRowsAsync(lines, label, 'B2B QZ', headers, {
      source: 'quinzenal',
      switchTab: opts.switchTab !== false,
      autosave: false
    });
  } finally {
    _fteSkipAutosave = false;
  }
  if (currentNFs.length) fteMarkCompanyLoaded();
  return currentNFs.length > 0;
}

function cteNfsInQuinzena(range) {
  if (!range || !hasConciliacaoCte()) return [];
  return currentNFs.filter(nf => {
    enrichNF(nf);
    const d = nf.dtRef || (nf.dtNF ? (parseSapBrDate(nf.dtNF) || (nf.dtNF instanceof Date ? nf.dtNF : null)) : null);
    if (!d || isNaN(d)) return false;
    return d >= range.start && d <= range.end;
  });
}

function cteNfsInMonth(ano, mes) {
  return cteNfsInQuinzena(monthDateRange(ano, mes));
}

function cteMesKeyFromNf(nf) {
  enrichNF(nf);
  return nf.mesRef === 'sem-data' ? 'sem-mes' : nf.mesRef;
}

function buildCteMonthTotalsMap() {
  const byMes = {};
  if (!hasConciliacaoCte()) return byMes;
  currentNFs.forEach(nf => {
    const k = cteMesKeyFromNf(nf);
    if (!byMes[k]) byMes[k] = { nfCount: 0, totalValorNF: 0, totalPago: 0, totalCte: 0 };
    const b = byMes[k];
    b.nfCount++;
    b.totalValorNF += nf.valorNF;
    b.totalPago += nf.pago;
    b.totalCte += nf.nCte || 0;
  });
  return byMes;
}

function cteGrandTotalsFromNFs() {
  if (!hasConciliacaoCte()) {
    return { nfCount: 0, totalValorNF: 0, totalPago: 0, totalCte: 0 };
  }
  if (currentSummary) {
    return {
      nfCount: currentSummary.totalNF,
      totalValorNF: currentSummary.totalValorNF,
      totalPago: currentSummary.totalPago,
      totalCte: currentNFs.reduce((s, x) => s + (x.nCte || 0), 0)
    };
  }
  return {
    nfCount: currentNFs.length,
    totalValorNF: currentNFs.reduce((s, x) => s + x.valorNF, 0),
    totalPago: currentNFs.reduce((s, x) => s + x.pago, 0),
    totalCte: currentNFs.reduce((s, x) => s + (x.nCte || 0), 0)
  };
}

function mergeQzB2BNf(into, r) {
  into.valorNF = Math.max(into.valorNF, r.valorNF);
  into.pago += r.pago;
  into.nCte += r.nCte;
  if (r.dtNF || r.dtNFQz) {
    const picked = pickEarliestDateValue(into.dtNFQz || into.dtNF, r.dtNFQz || r.dtNF);
    into.dtNF = picked;
    into.dtNFQz = picked;
  }
  if (r.temDevolucao) into.temDevolucao = true;
  if (Array.isArray(r.ctes) && r.ctes.length) {
    if (!Array.isArray(into.ctes)) into.ctes = [];
    into.ctes.push(...r.ctes);
  }
}

function quinzenaFromKey(qk) {
  const m = String(qk || '').match(/-Q([12])$/);
  return m ? parseInt(m[1], 10) : null;
}

function parseQuinzenaKeyParts(qk) {
  const m = String(qk || '').match(/^(\d{4})-(\d{2})-Q([12])$/);
  if (!m) return null;
  return { ano: parseInt(m[1], 10), mes: parseInt(m[2], 10), quinzena: parseInt(m[3], 10) };
}

function qzExpandKey(panel, mesKey) {
  return `${panel}:${mesKey}`;
}

function qzByMesKey(qzTotals) {
  const map = {};
  (qzTotals || []).forEach(q => {
    const k = q.mesKey || 'sem-mes';
    if (!map[k]) map[k] = [];
    map[k].push(q);
  });
  Object.values(map).forEach(arr => arr.sort((a, b) => (a.quinzena || 0) - (b.quinzena || 0)));
  return map;
}

function qzQuinzenaCellLabel(q) {
  const base = q.quinzena === 1 ? '1ª quinzena' : (q.quinzena === 2 ? '2ª quinzena' : (q.quinzenaLabel || 'Quinzena'));
  if (q.fileName && q.fileName !== base) {
    const safe = String(q.fileName).replace(/"/g, '&quot;');
    return `${base}<div class="qz-qz-file" title="${safe}">${safe}</div>`;
  }
  return base;
}

function buildB2BQuinzenaTotals(b2bRows) {
  const byQz = {};
  b2bRows.forEach(r => {
    const qk = r.quinzenaKey || 'sem-quinzena';
    if (!byQz[qk]) {
      byQz[qk] = {
        mesKey: r.mesKey || mesKeyFromQuinzenaKey(qk),
        mesLabel: r.mesLabel || (r.mesKey ? fmtMesLabel(r.mesKey) : ''),
        quinzenaKey: qk,
        quinzena: quinzenaFromKey(qk),
        quinzenaLabel: r.quinzenaLabel || (quinzenaFromKey(qk) ? `${quinzenaFromKey(qk)}ªQ` : qk),
        fileName: r.fileName || '',
        byNf: {}
      };
    }
    const nfKey = r.nfKey || normNFKey(r.nf);
    if (!nfKey) return;
    if (!byQz[qk].byNf[nfKey]) byQz[qk].byNf[nfKey] = { valorNF: 0, pago: 0, nCte: 0 };
    mergeQzB2BNf(byQz[qk].byNf[nfKey], r);
    if (r.fileName) byQz[qk].fileName = r.fileName;
  });
  return Object.values(byQz).map(b => {
    const nfs = Object.values(b.byNf);
    const out = {
      mesKey: b.mesKey, mesLabel: b.mesLabel,
      quinzenaKey: b.quinzenaKey, quinzena: b.quinzena,
      quinzenaLabel: b.quinzenaLabel, fileName: b.fileName,
      nfCountQz: nfs.length,
      totalValorNFQz: nfs.reduce((s, x) => s + x.valorNF, 0),
      totalPagoQz: nfs.reduce((s, x) => s + x.pago, 0),
      totalCteQz: nfs.reduce((s, x) => s + x.nCte, 0),
      nfCountCte: 0, totalValorNFCte: 0, totalPagoCte: 0, totalCteCte: 0
    };
    const pk = parseQuinzenaKeyParts(b.quinzenaKey);
    if (pk) {
      const cteList = cteNfsInQuinzena(quinzenaDateRange(pk));
      out.nfCountCte = cteList.length;
      out.totalValorNFCte = cteList.reduce((s, x) => s + x.valorNF, 0);
      out.totalPagoCte = cteList.reduce((s, x) => s + x.pago, 0);
      out.totalCteCte = cteList.reduce((s, x) => s + x.nCte, 0);
    }
    return out;
  }).sort((a, b) => compareMesRef(a.mesKey, b.mesKey) || (a.quinzena || 0) - (b.quinzena || 0));
}

function buildB2CQuinzenaTotals(b2cRows) {
  const byQz = {};
  b2cRows.forEach(r => {
    const mesKey = resolveB2cRowMesKey(r);
    const qk = r.quinzenaKey || 'sem-quinzena';
    const groupKey = `${mesKey}|${qk}`;
    if (!byQz[groupKey]) {
      byQz[groupKey] = {
        mesKey,
        mesLabel: resolveB2cRowMesLabel(r, mesKey),
        quinzenaKey: qk,
        quinzena: quinzenaFromKey(qk),
        quinzenaLabel: r.quinzenaLabel || (quinzenaFromKey(qk) ? `${quinzenaFromKey(qk)}ªQ` : qk),
        fileName: r.fileName || '',
        count: 0, totalVendas: 0, totalFrete: 0
      };
    }
    byQz[groupKey].count++;
    byQz[groupKey].totalVendas += r.valorNF;
    byQz[groupKey].totalFrete += r.pago;
    if (r.fileName) byQz[groupKey].fileName = r.fileName;
  });
  return Object.values(byQz).map(b => ({
    ...b,
    pctFrete: b.totalVendas ? b.totalFrete / b.totalVendas : 0
  })).sort((a, b) => compareMesRef(a.mesKey, b.mesKey) || (a.quinzena || 0) - (b.quinzena || 0));
}

function buildB2BMonthTotals(b2bRows) {
  const qzByMes = {};
  b2bRows.forEach(r => {
    const k = r.mesKey || mesKeyFromQuinzenaKey(r.quinzenaKey);
    if (!qzByMes[k]) {
      qzByMes[k] = {
        mesKey: k, mesLabel: r.mesLabel || (k !== 'sem-mes' ? fmtMesLabel(k) : 'Sem mês'),
        byNf: {}
      };
    }
    const nfKey = r.nfKey || normNFKey(r.nf);
    if (!nfKey) return;
    if (!qzByMes[k].byNf[nfKey]) qzByMes[k].byNf[nfKey] = { valorNF: 0, pago: 0, nCte: 0 };
    mergeQzB2BNf(qzByMes[k].byNf[nfKey], r);
  });
  const cteByMes = buildCteMonthTotalsMap();
  const allKeys = new Set([...Object.keys(qzByMes), ...Object.keys(cteByMes)]);
  return [...allKeys].map(k => {
    const qz = qzByMes[k];
    const cte = cteByMes[k];
    const nfs = qz ? Object.values(qz.byNf) : [];
    return {
      mesKey: k,
      mesLabel: qz?.mesLabel || (k !== 'sem-mes' ? fmtMesLabel(k) : 'Sem mês'),
      nfCountQz: nfs.length,
      totalValorNFQz: nfs.reduce((s, x) => s + x.valorNF, 0),
      totalPagoQz: nfs.reduce((s, x) => s + x.pago, 0),
      totalCteQz: nfs.reduce((s, x) => s + x.nCte, 0),
      nfCountCte: cte?.nfCount || 0,
      totalValorNFCte: cte?.totalValorNF || 0,
      totalPagoCte: cte?.totalPago || 0,
      totalCteCte: cte?.totalCte || 0
    };
  }).sort((a, b) => compareMesRef(a.mesKey, b.mesKey));
}

function buildB2CMonthTotals(b2cRows) {
  const byMes = {};
  b2cRows.forEach(r => {
    const k = resolveB2cRowMesKey(r);
    if (!byMes[k]) byMes[k] = { mesKey: k, mesLabel: resolveB2cRowMesLabel(r, k), count: 0, totalVendas: 0, totalFrete: 0 };
    byMes[k].count++; byMes[k].totalVendas += r.valorNF; byMes[k].totalFrete += r.pago;
  });
  return Object.values(byMes).map(b => ({
    ...b,
    pctFrete: b.totalVendas ? b.totalFrete / b.totalVendas : 0,
    freteMedio: b.count ? b.totalFrete / b.count : 0
  })).sort((a, b) => compareMesRef(a.mesKey, b.mesKey));
}

function aggregateB2cMonthGroup(group) {
  const t = { count: 0, totalVendas: 0, totalFrete: 0 };
  group.forEach(m => { t.count += m.count; t.totalVendas += m.totalVendas; t.totalFrete += m.totalFrete; });
  t.pctFrete = t.totalVendas ? t.totalFrete / t.totalVendas : 0;
  t.freteMedio = t.count ? t.totalFrete / t.count : 0;
  return t;
}

function aggregateB2bMonthGroup(group) {
  const t = {
    nfCountQz: 0, nfCountCte: 0,
    totalValorNFQz: 0, totalValorNFCte: 0,
    totalPagoQz: 0, totalPagoCte: 0,
    totalCteQz: 0, totalCteCte: 0
  };
  group.forEach(m => {
    t.nfCountQz += m.nfCountQz;
    t.nfCountCte += m.nfCountCte;
    t.totalValorNFQz += m.totalValorNFQz;
    t.totalValorNFCte += m.totalValorNFCte;
    t.totalPagoQz += m.totalPagoQz;
    t.totalPagoCte += m.totalPagoCte;
    t.totalCteQz += m.totalCteQz;
    t.totalCteCte += m.totalCteCte;
  });
  return t;
}

function buildB2bMonthDisplayRows(totals, cteGrand) {
  const out = [];
  const dated = totals.filter(m => m.mesKey !== 'sem-mes');
  const noData = totals.filter(m => m.mesKey === 'sem-mes');
  let lastYear = null;
  let yearGroup = [];

  dated.forEach(m => {
    const year = m.mesKey.slice(0, 4);
    if (lastYear && year !== lastYear && yearGroup.length) {
      out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...aggregateB2bMonthGroup(yearGroup) });
      yearGroup = [];
    }
    if (year !== lastYear) {
      out.push({ type: 'year', label: year, year });
      lastYear = year;
    }
    out.push({ type: 'month', ...m });
    yearGroup.push(m);
  });
  if (yearGroup.length && lastYear) {
    out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...aggregateB2bMonthGroup(yearGroup) });
  }
  noData.forEach(m => out.push({ type: 'month', ...m }));
  if (totals.length) {
    const t = aggregateB2bMonthGroup(totals);
    if (cteGrand) {
      t.nfCountCte = cteGrand.nfCount;
      t.totalValorNFCte = cteGrand.totalValorNF;
      t.totalPagoCte = cteGrand.totalPago;
      t.totalCteCte = cteGrand.totalCte;
    }
    out.push({ type: 'total', label: 'Total geral', ...t });
  }
  return out;
}

function buildB2cMonthDisplayRows(totals) {
  const out = [];
  const dated = totals.filter(m => m.mesKey !== 'sem-mes');
  const noData = totals.filter(m => m.mesKey === 'sem-mes');
  let lastYear = null;
  let yearGroup = [];

  dated.forEach(m => {
    const year = m.mesKey.slice(0, 4);
    if (lastYear && year !== lastYear && yearGroup.length) {
      out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...aggregateB2cMonthGroup(yearGroup) });
      yearGroup = [];
    }
    if (year !== lastYear) {
      out.push({ type: 'year', label: year, year });
      lastYear = year;
    }
    out.push({ type: 'month', ...m });
    yearGroup.push(m);
  });
  if (yearGroup.length && lastYear) {
    out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...aggregateB2cMonthGroup(yearGroup) });
  }
  noData.forEach(m => out.push({ type: 'month', ...m }));
  if (totals.length) {
    out.push({ type: 'total', label: 'Total geral', ...aggregateB2cMonthGroup(totals) });
  }
  return out;
}

function buildB2CRegionTotals(b2cRows) {
  const byReg = {};
  b2cRows.forEach(r => {
    const reg = String(r.zona || '').trim() || 'Sem região';
    if (!byReg[reg]) byReg[reg] = { regiao: reg, count: 0, totalVendas: 0, totalFrete: 0 };
    byReg[reg].count++;
    byReg[reg].totalVendas += r.valorNF;
    byReg[reg].totalFrete += r.pago;
  });
  return Object.values(byReg).map(b => ({
    ...b,
    pctFrete: b.totalVendas ? b.totalFrete / b.totalVendas : 0,
    freteMedio: b.count ? b.totalFrete / b.count : 0
  })).sort((a, b) => b.totalVendas - a.totalVendas);
}

function buildB2CYearTotals(monthTotals) {
  const byYear = {};
  monthTotals.forEach(m => {
    const year = m.mesKey === 'sem-mes' ? 'sem-ano' : m.mesKey.slice(0, 4);
    if (!byYear[year]) byYear[year] = { year, label: year === 'sem-ano' ? 'Sem ano' : year, count: 0, totalVendas: 0, totalFrete: 0 };
    byYear[year].count += m.count;
    byYear[year].totalVendas += m.totalVendas;
    byYear[year].totalFrete += m.totalFrete;
  });
  return Object.values(byYear).map(b => ({
    ...b,
    pctFrete: b.totalVendas ? b.totalFrete / b.totalVendas : 0,
    freteMedio: b.count ? b.totalFrete / b.count : 0
  })).sort((a, b) => String(a.year).localeCompare(String(b.year)));
}

function renderB2cKpiBlock(label, stats) {
  return `<div class="qz-year-kpi-block">
    <div class="qz-year-kpi-label">${label}</div>
    <div class="kpis">
      <div class="kpi"><div class="label">Entregas B2C</div><div class="value">${stats.count}</div></div>
      <div class="kpi"><div class="label">Vendas (Total NF)</div><div class="value">${fmtMoney(stats.totalVendas)}</div></div>
      <div class="kpi"><div class="label">Frete pago</div><div class="value">${fmtMoney(stats.totalFrete)}</div></div>
      <div class="kpi"><div class="label">% frete/vendas</div><div class="value">${fmtPct(stats.pctFrete)}</div></div>
    </div>
  </div>`;
}

function renderB2bCompareKpis(monthTotals, cteGrand, compareRows) {
  const t = aggregateB2bMonthGroup(monthTotals || []);
  const hasCteSide = hasConciliacaoCte();
  if (!hasCteSide) {
    return `<div class="qz-year-kpi-block">
      <div class="qz-year-kpi-label">Quinzenal B2B (vista só QZ)</div>
      <div class="kpis">
        <div class="kpi"><div class="label">NFs</div><div class="value">${t.nfCountQz}</div></div>
        <div class="kpi"><div class="label">Faturação QZ</div><div class="value">${fmtMoney(t.totalValorNFQz)}</div></div>
        <div class="kpi"><div class="label">Frete QZ</div><div class="value">${fmtMoney(t.totalPagoQz)}</div></div>
        <div class="kpi"><div class="label">CT-e (distintos)</div><div class="value">${t.totalCteQz}</div></div>
      </div>
    </div>`;
  }
  const fatCte = cteGrand?.totalValorNF ?? t.totalValorNFCte;
  const freteCte = cteGrand?.totalPago ?? t.totalPagoCte;
  const s = qzCompareSummary(compareRows || []);
  return `<div class="qz-year-kpi-block">
    <div class="qz-year-kpi-label">Faturação (valor NF)</div>
    <div class="kpis">
      <div class="kpi"><div class="label">Faturação QZ</div><div class="value">${fmtMoney(t.totalValorNFQz)}</div></div>
      <div class="kpi"><div class="label">Faturação CT-e</div><div class="value">${fmtMoney(fatCte)}</div></div>
      <div class="kpi flag"><div class="label">Δ total</div><div class="value">${fmtQzDiff(fatCte - t.totalValorNFQz, 1)}</div></div>
    </div>
  </div>
  <div class="qz-year-kpi-block">
    <div class="qz-year-kpi-label">Fretes cobrados</div>
    <div class="kpis">
      <div class="kpi"><div class="label">Frete QZ</div><div class="value">${fmtMoney(t.totalPagoQz)}</div></div>
      <div class="kpi"><div class="label">Frete CT-e</div><div class="value">${fmtMoney(freteCte)}</div></div>
      <div class="kpi flag"><div class="label">Δ total</div><div class="value">${fmtQzDiff(freteCte - t.totalPagoQz, 0.5)}</div></div>
    </div>
  </div>
  <div class="qz-year-kpi-block">
    <div class="qz-year-kpi-label">Confronto NF (quinzenal vs CT-e) <span style="font-weight:400;font-size:10px;color:var(--muted)">— clica para filtrar</span></div>
    <div class="kpis">
      <div class="kpi qz-kpi-filter" data-qz-filter="onlyQz" role="button" tabindex="0" title="NFs só no quinzenal B2B">
        <div class="label">Só quinzenal</div><div class="value">${s.onlyQz}</div>
      </div>
      <div class="kpi qz-kpi-filter" data-qz-filter="onlyCte" role="button" tabindex="0" title="NFs só no export CT-e">
        <div class="label">Só CT-e</div><div class="value">${s.onlyCte}</div>
      </div>
      <div class="kpi qz-kpi-filter" data-qz-filter="diff" role="button" tabindex="0" title="NFs em ambos com diferença de valor/frete/CT-e">
        <div class="label">Com diferença</div><div class="value">${s.diff}</div>
      </div>
      <div class="kpi"><div class="label">OK (match)</div><div class="value">${s.match}</div></div>
    </div>
  </div>`;
}

function renderB2bMonthTableRow(label, m, rowClass) {
  return `<tr class="${rowClass || 'qz-month-row'}"><td><strong>${label}</strong></td>
    <td class="right">${m.nfCountQz}</td><td class="right">${m.nfCountCte}</td><td class="right">${m.nfCountCte - m.nfCountQz}</td>
    <td class="right">${fmtMoney(m.totalValorNFQz)}</td><td class="right">${fmtMoney(m.totalValorNFCte)}</td>
    <td class="right">${fmtQzDiff(m.totalValorNFCte - m.totalValorNFQz, 1)}</td>
    <td class="right">${fmtMoney(m.totalPagoQz)}</td><td class="right">${fmtMoney(m.totalPagoCte)}</td>
    <td class="right">${fmtQzDiff(m.totalPagoCte - m.totalPagoQz, 0.5)}</td>
    <td class="right">${m.totalCteQz}</td><td class="right">${m.totalCteCte}</td><td class="right">${m.totalCteCte - m.totalCteQz}</td></tr>`;
}

function renderB2cMonthTableRow(label, m, rowClass) {
  const freteMedio = m.freteMedio ?? (m.count ? m.totalFrete / m.count : 0);
  return `<tr class="${rowClass || 'qz-month-row'}"><td><strong>${label}</strong></td>
    <td class="right">${m.count}</td><td class="right">${fmtMoney(m.totalVendas)}</td>
    <td class="right">${fmtMoney(m.totalFrete)}</td><td class="right">${fmtPct(m.pctFrete)}</td>
    <td class="right">${fmtMoney(freteMedio)}</td></tr>`;
}

function parseB2bQzDtNF(v) {
  if (!v) return null;
  const d = parseSapBrDate(v);
  if (d && !isNaN(d) && isPlausibleSapDate(d)) return d;
  if (v instanceof Date && !isNaN(v) && isPlausibleSapDate(v)) return v;
  return null;
}

function parseB2bCteDtNF(cte) {
  if (!cte) return null;
  enrichNF(cte);
  return cte.dtRef || parseB2bQzDtNF(cte.dtNF);
}

function fmtB2bCompareMesCell(r) {
  const qz = r.mesLabelQz || r.mesLabel || '';
  const cte = r.mesLabelCte || '';
  if (qz && cte && qz !== cte) return `${qz} · ${cte}`;
  return qz || cte || '-';
}

function fmtB2bCompareDateCell(r) {
  const qz = fmtDate(r.dtNFQz);
  const cte = fmtDate(r.dtNFCte);
  if (r.status === 'onlyQz') return qz;
  if (r.status === 'onlyCte') return cte;
  if (qz !== '-' && cte !== '-' && qz !== cte) {
    return `<span style="color:#b3261e" title="Data QZ vs CT-e/SAP">${qz} · ${cte}</span>`;
  }
  if (qz !== '-' && cte !== '-') return qz;
  return qz !== '-' ? qz : cte;
}

function buildB2BCompare(b2bRows) {
  const qzMap = {};
  b2bRows.forEach(r => {
    const k = r.nfKey;
    if (!qzMap[k]) qzMap[k] = { ...r };
    else mergeQzB2BNf(qzMap[k], r);
  });
  const cteMap = {};
  // Only Conciliacao CT-e×NF counts as the CT-e side — QZ-derived currentNFs must not self-compare
  if (hasConciliacaoCte()) {
    currentNFs.forEach(nf => { const k = normNFKey(nf.nf); if (k) cteMap[k] = nf; });
  }
  const keys = new Set([...Object.keys(qzMap), ...Object.keys(cteMap)]);
  const rows = [];
  keys.forEach(k => {
    const qz = qzMap[k];
    const cte = cteMap[k];
    const valorNFQz = qz?.valorNF || 0;
    const valorNFCte = cte?.valorNF || 0;
    const pagoQz = qz?.pago || 0;
    const pagoCte = cte?.pago || 0;
    const nCteQz = qz?.nCte || 0;
    const nCteCte = cte?.nCte || 0;
    const dtNFQz = parseB2bQzDtNF(qz?.dtNF);
    const dtNFCte = parseB2bCteDtNF(cte);
    const mesKeyQz = qz?.mesKey || mesKeyFromQuinzenaKey(qz?.quinzenaKey) || '';
    const mesLabelQz = qz?.mesLabel || (mesKeyQz ? fmtMesLabel(mesKeyQz) : '');
    let mesKeyCte = '';
    let mesLabelCte = '';
    if (cte) {
      enrichNF(cte);
      mesKeyCte = cte.mesRef && cte.mesRef !== 'sem-data' ? cte.mesRef : '';
      mesLabelCte = mesKeyCte ? fmtMesLabel(mesKeyCte) : '';
    }
    let status = 'match';
    if (!cte && Object.keys(cteMap).length === 0) {
      // No Conciliacao loaded — QZ inventory view, not a mismatch
      status = 'qzView';
    } else if (qz && !cte) status = 'onlyQz';
    else if (!qz && cte) status = 'onlyCte';
    else if (Math.abs(valorNFQz - valorNFCte) > 1 || Math.abs(pagoQz - pagoCte) > 0.5 || nCteQz !== nCteCte) status = 'diff';
    rows.push({
      nfKey: k, nf: qz?.nf || cte?.nf || k,
      mesKey: mesKeyQz || mesKeyCte,
      mesKeyQz, mesKeyCte,
      mesLabel: mesLabelQz || mesLabelCte,
      mesLabelQz, mesLabelCte,
      dtNFQz, dtNFCte,
      quinzenaKey: qz?.quinzenaKey || '', quinzenaLabel: qz?.quinzenaLabel || '', fileName: qz?.fileName || '',
      valorNFQz, valorNFCte, diffValorNF: valorNFCte - valorNFQz,
      pagoQz, pagoCte, diffPago: pagoCte - pagoQz,
      nCteQz, nCteCte, diffCte: nCteCte - nCteQz, status,
      transportador: qz?.transportador || cte?.transportador || ''
    });
  });
  const ord = { onlyQz: 0, onlyCte: 1, diff: 2, match: 3 };
  return rows.sort((a, b) => (ord[a.status] ?? 9) - (ord[b.status] ?? 9) || String(a.nf).localeCompare(String(b.nf)));
}

function buildQuinzenalPack(fileResults, fileBinaries) {
  const b2bRows = [];
  const b2cRows = [];
  const files = [];
  const failedFiles = [];
  const binaries = { ...(fileBinaries || {}) };
  fileResults.forEach(fr => {
    if (!fr.ok) {
      failedFiles.push({ fileName: fr.fileName, error: fr.error || 'Erro desconhecido' });
      return;
    }
    const { meta, canal, rows, sheetName } = fr;
    if (canal === 'B2B') {
      const agg = aggregateQzB2BRows(rows, meta);
      b2bRows.push(...agg);
      files.push({
        fileName: meta.fileName, canal, mesKey: meta.mesKey, mesLabel: meta.mesLabel,
        quinzenaKey: meta.quinzenaKey, quinzenaLabel: meta.quinzenaLabel,
        sheetName, rowCount: rows.length, nfCount: agg.length,
        totalValorNF: agg.reduce((s, x) => s + x.valorNF, 0),
        totalPago: agg.reduce((s, x) => s + x.pago, 0),
        totalCte: agg.reduce((s, x) => s + x.nCte, 0)
      });
    } else {
      rows.forEach(r => {
        const row = enrichB2cRowMesFields({
          ...r, valorNF: num(r.valorNF) || num(r.valorProdutos), pago: num(r.pago),
          quinzenaKey: meta.quinzenaKey, quinzenaLabel: meta.quinzenaLabel, fileName: meta.fileName
        }, meta);
        b2cRows.push(row);
      });
      files.push({
        fileName: meta.fileName, canal, mesKey: meta.mesKey, mesLabel: meta.mesLabel,
        quinzenaKey: meta.quinzenaKey, quinzenaLabel: meta.quinzenaLabel,
        sheetName, rowCount: rows.length, nfCount: rows.length,
        totalValorNF: rows.reduce((s, x) => s + (num(x.valorNF) || num(x.valorProdutos)), 0),
        totalPago: rows.reduce((s, x) => s + num(x.pago), 0),
        totalCte: rows.filter(x => x.numCte).length
      });
    }
  });
  const pack = { version: 2, updatedAt: new Date().toISOString(), files, failedFiles, b2bRows, b2cRows, fileBinaries: binaries };
  pack.b2bCompare = buildB2BCompare(b2bRows);
  pack.b2bMonthTotals = buildB2BMonthTotals(b2bRows);
  pack.b2bQuinzenaTotals = buildB2BQuinzenaTotals(b2bRows);
  pack.b2cMonthTotals = buildB2CMonthTotals(b2cRows);
  pack.b2cQuinzenaTotals = buildB2CQuinzenaTotals(b2cRows);
  pack.b2cRegionTotals = buildB2CRegionTotals(b2cRows);
  return pack;
}

function qzFileCounts() {
  const files = quinzenalPack?.files || [];
  const failed = quinzenalPack?.failedFiles || [];
  return {
    b2c: files.filter(f => f.canal === 'B2C').length,
    b2b: files.filter(f => f.canal === 'B2B').length,
    total: files.length,
    failed: failed.length
  };
}

function qzFileCountNoteHtml() {
  const c = qzFileCounts();
  if (!c.total && !c.failed) return '';
  let s = `${c.b2c} ficheiro${c.b2c !== 1 ? 's' : ''} B2C · ${c.b2b} ficheiro${c.b2b !== 1 ? 's' : ''} B2B carregados`;
  if (c.failed) {
    const details = (quinzenalPack.failedFiles || []).map(f => {
      const shortName = String(f.fileName || '').replace(/\.xlsx?$/i, '').replace(/^Gest[aã]o de frete DELTA[_ ]?B2[BC]\s*-\s*/i, '');
      return `${shortName || f.fileName}: ${f.error || 'Erro desconhecido'}`;
    }).join(' · ');
    s += ` · ${c.failed} com erro (${details})`;
  }
  return s;
}

function updateQzFileNote() {
  const el = $('qzFileNote');
  if (!el) return;
  const note = qzFileCountNoteHtml();
  if (note) { el.style.display = 'block'; el.textContent = note; }
  else { el.style.display = 'none'; el.textContent = ''; }
}

function qzCompareSummary(rows) {
  const s = { total: rows.length, onlyQz: 0, onlyCte: 0, diff: 0, match: 0, diffValor: 0 };
  rows.forEach(r => {
    s[r.status] = (s[r.status] || 0) + 1;
    if (Math.abs(r.diffValorNF) > 1) s.diffValor++;
  });
  return s;
}

function refreshQuinzenalCompare() {
  if (!quinzenalPack?.b2bRows) return;
  quinzenalPack.b2bCompare = buildB2BCompare(quinzenalPack.b2bRows);
  quinzenalPack.b2bMonthTotals = buildB2BMonthTotals(quinzenalPack.b2bRows);
  quinzenalPack.b2bQuinzenaTotals = buildB2BQuinzenaTotals(quinzenalPack.b2bRows);
  if (quinzenalPack.b2cRows) {
    quinzenalPack.b2cRows.forEach(r => enrichB2cRowMesFields(r));
    quinzenalPack.b2cMonthTotals = buildB2CMonthTotals(quinzenalPack.b2cRows);
    quinzenalPack.b2cQuinzenaTotals = buildB2CQuinzenaTotals(quinzenalPack.b2cRows);
    quinzenalPack.b2cRegionTotals = buildB2CRegionTotals(quinzenalPack.b2cRows);
  }
}

function updateQzProcessStatus() {
  /* qz status shown in fileStatus */
}

function syncQzUploadZone() {
  const fn = $('qzFn');
  const zone = $('qzZone');
  const list = $('qzFileList');
  const pending = fteQzPendingFiles.length;
  const saved = quinzenalPack?.files?.length || 0;
  if (pending) {
    if (fn) fn.textContent = `✓ ${pending} ficheiro(s) seleccionado(s)`;
    zone?.classList.add('loaded');
    if (list) {
      list.style.display = 'block';
      list.innerHTML = fteQzPendingFiles.map(f => `<div>${f.name}</div>`).join('');
    }
  } else if (saved) {
    const c = qzFileCounts();
    if (fn) fn.textContent = `✓ ${c.b2c} B2C · ${c.b2b} B2B carregados`;
    zone?.classList.add('loaded');
    if (list) {
      list.style.display = 'block';
      list.innerHTML = quinzenalPack.files.map(f => `<div>${f.fileName}</div>`).join('');
    }
  } else {
    if (fn) fn.textContent = '';
    zone?.classList.remove('loaded');
    if (list) { list.style.display = 'none'; list.innerHTML = ''; }
  }
  checkFteBtn();
  renderLoadCoverageSummary();
}

function setQzZoneLoaded(count) {
  syncQzUploadZone();
}

async function persistQuinzenalPack(pack, opts = {}) {
  if (typeof upsertExcelBinary !== 'function') {
    fteToastError('Persistência Excel indisponível — recarrega a página.');
    return false;
  }
  if (!pack?.files?.length) {
    console.warn('[fretes] persist quinzenal skip — empty pack');
    if (!opts.silent) fteToastError('Sem ficheiros quinzenais para guardar.');
    return false;
  }
  const slim = slimQuinzenalPackForPersist(pack);
  const json = JSON.stringify(slim);
  const buf = new TextEncoder().encode(json).buffer;
  const jsonBytes = buf.byteLength;
  if (jsonBytes > QZ_PERSIST_MAX_JSON_BYTES) {
    const msg = `Quinzenais demasiado grandes para guardar (${fmtByteSize(jsonBytes)} > ${fmtByteSize(QZ_PERSIST_MAX_JSON_BYTES)}). Reduz ficheiros ou contacta suporte.`;
    console.error('[fretes] persist quinzenal too large', fteCompany(), jsonBytes);
    fteToastError(msg);
    return false;
  }
  const slot = fteQuinzenalSlot();
  const counts = quinzenalPackCounts(slim) || qzFileCounts();
  const fileName = `quinzenal_${counts.total}_${counts.b2c}B2C_${counts.b2b}B2B.json`;
  console.log('[fretes] persist quinzenal', fteCompany(), slot, counts.total, 'files',
    counts.b2c, 'B2C', counts.b2b, 'B2B', 'jsonBytes', jsonBytes);
  try {
    await upsertExcelBinary(slot, fileName, buf);
    if (typeof fetchExcelFileMeta === 'function') {
      const verify = await fetchExcelFileMeta([slot]);
      const rec = verify[slot];
      if (!rec?.file_name) {
        const msg = `Quinzenais: guardado não confirmado na cloud (${fteCompany()}) — o servidor não devolveu metadados.`;
        console.error('[fretes] persist quinzenal verify failed', fteCompany(), slot, rec);
        fteToastError(msg);
        return false;
      }
    }
    // Never keep raw xlsx binaries after persist — slim JSON is enough for restore.
    if (pack?.fileBinaries) delete pack.fileBinaries;
    if (quinzenalPack?.fileBinaries) delete quinzenalPack.fileBinaries;
    syncQzUploadZone();
    updateQzFileNote();
    updateQzProcessStatus();
    updateFretesFileStatus();
    console.log('[fretes] persist quinzenal ok', fteCompany(), slot, fileName, 'jsonBytes', jsonBytes);
    if (!opts.silent) fteToast(`Quinzenais guardados: ${counts.b2c} B2C · ${counts.b2b} B2B`);
    return true;
  } catch (err) {
    console.error('[fretes] persist quinzenal', fteCompany(), slot, 'jsonBytes', jsonBytes, err);
    if (typeof isExcelFilesTableMissing === 'function' && isExcelFilesTableMissing(err)) {
      fteToastError('Tabela logistica_excel_files em falta no Supabase.');
    } else {
      fteToastError(`Erro ao guardar quinzenais (${counts.b2c} B2C · ${counts.b2b} B2B, ${fmtByteSize(jsonBytes)}): ${err.message || err}`);
    }
    return false;
  }
}

async function loadSavedQuinzenalPack(silent, meta, opts = {}) {
  if (typeof fetchExcelFiles !== 'function') return false;
  try {
    if (typeof syncCompanyDepotMap === 'function') syncCompanyDepotMap();
    const m = meta || await fetchExcelFiles([fteQuinzenalSlot()]);
    const rec = m[fteQuinzenalSlot()];
    if (!rec?.file_data) {
      if (rec?.file_name) {
        console.warn('[fretes] quinzenal metadata only', fteCompany(), rec.file_name);
        if (!silent) fteToastError('Quinzenais guardados só com nome — volta a carregar os Excels e clica Processar e Guardar.');
      }
      return false;
    }
    // Already have this pack in RAM — skip multi-MB base64→JSON parse
    if (quinzenalPack?.files?.length && _fteLoadedCompany === fteCompany()) {
      console.log('[fretes] quinzenal cache hit', fteCompany(), quinzenalPack.files.length, 'files');
      if (!opts?.deferCompare) refreshQuinzenalCompare();
      syncQzUploadZone();
      updateQzFileNote();
      updateQzProcessStatus();
      return true;
    }
    // Capture length BEFORE parse — parseQuinzenalPackFromRec nulls rec.file_data to free RAM.
    const dataLen = rec.file_data?.length || 0;
    await fteYield('A descodificar quinzenais…');
    const parsed = parseQuinzenalPackFromRec(rec);
    if (!parsed?.files?.length) {
      console.warn('[fretes] quinzenal parse empty', fteCompany(), rec.file_name, 'dataLen', dataLen);
      if (!silent) fteToastError('Quinzenais guardados corrompidos — volta a carregar os ficheiros e clica Processar e Guardar.');
      return false;
    }
    quinzenalPack = parsed;
    const c = quinzenalPackCounts(quinzenalPack);
    console.log('[fretes] load quinzenal', fteCompany(), rec.file_name, 'dataLen', dataLen,
      'files', c?.total || 0, c?.b2c || 0, 'B2C', c?.b2b || 0, 'B2B');
    if (!opts?.deferCompare) {
      await fteYield('A calcular confrontos QZ…');
      refreshQuinzenalCompare();
    }
    syncQzUploadZone();
    updateQzFileNote();
    updateQzProcessStatus();
    if (!silent) fteToast(`Quinzenais restaurados: ${c?.b2c || 0} B2C · ${c?.b2b || 0} B2B`);
    if (!opts?.skipRender) {
      const activeTab = document.querySelector('.fte-tab.active')?.dataset?.tab;
      if (activeTab === 'analise-b2c') renderB2cAnalysisTab();
      else if (activeTab === 'cte-vs-qz') renderB2bCompareTab();
      else if (activeTab === 'resumo-total') renderResumoTotal();
    }
    return true;
  } catch (err) {
    console.error('[fretes] load quinzenal', fteCompany(), err);
    if (!silent) fteToastError('Erro ao carregar quinzenais: ' + (err.message || err));
    return false;
  }
}

function diagnoseQzParseFailure(wb, canal) {
  const sheetNames = wb.SheetNames || [];
  if (!sheetNames.length) return 'Workbook sem folhas';
  const headerFn = canal === 'B2B' ? isQzB2BHeader : isQzB2CHeader;
  const headerHint = canal === 'B2B'
    ? 'Nota fiscal + Total fatura'
    : 'Pedido/chave NF + Valor fatura ou Valor produto';
  for (const name of sheetNames) {
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null });
    if (!raw.length) continue;
    const headerIdx = raw.findIndex(row => Array.isArray(row) && headerFn(row));
    if (headerIdx < 0) continue;
    const parsed = parseQzSheetRows(wb.Sheets[name], canal);
    if (parsed.rows.length) return null;
    const cols = (parsed.headers || []).slice(0, 8).join(', ') || 'sem colunas';
    return `Folha "${name}": cabeçalho ${canal} reconhecido mas 0 linhas (${cols})`;
  }
  return `Cabeçalho ${canal} não encontrado (${headerHint}) — folhas: ${sheetNames.join(', ')}`;
}

function processQuinzenalFile(file) {
  return new Promise((resolve) => {
    const meta = parseQuinzenalFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = readWorkbookFromArrayBuffer(e.target.result);
        if (!meta.canal) {
          const detected = detectCanalFromWorkbook(wb);
          if (detected) meta.canal = detected;
        }
        if (!meta.canal) { resolve({ ok: false, fileName: file.name, error: 'Canal B2B/B2C não detectado no nome nem nas folhas' }); return; }
        const loaded = loadQuinzenalFromWorkbook(wb, meta);
        if (!loaded.rows.length) {
          const err = diagnoseQzParseFailure(wb, meta.canal) || 'Sem linhas de detalhe';
          resolve({ ok: false, fileName: file.name, error: err });
          return;
        }
        resolve({ ok: true, ...loaded, meta, canal: meta.canal, arrayBuffer: e.target.result });
      } catch (err) { resolve({ ok: false, fileName: file.name, error: String(err.message || err) }); }
    };
    reader.onerror = () => resolve({ ok: false, fileName: file.name, error: 'Erro de leitura do ficheiro' });
    reader.readAsArrayBuffer(file);
  });
}

function fmtQzDiff(v, threshold) {
  if (Math.abs(v) <= (threshold ?? 0.01)) return '<span style="color:var(--green)">0</span>';
  const sign = v > 0 ? '+' : '';
  return `<span style="color:${v > 0 ? '#b3261e' : '#555'}">${sign}${fmtMoney(v)}</span>`;
}

function renderB2bQuinzenaRow(q) {
  return `<tr class="qz-qz-row"><td style="padding-left:26px">${qzQuinzenaCellLabel(q)}</td>
    <td class="right">${q.nfCountQz}</td><td class="right">${q.nfCountCte}</td><td class="right">${q.nfCountCte - q.nfCountQz}</td>
    <td class="right">${fmtMoney(q.totalValorNFQz)}</td><td class="right">${fmtMoney(q.totalValorNFCte)}</td>
    <td class="right">${fmtQzDiff(q.totalValorNFCte - q.totalValorNFQz, 1)}</td>
    <td class="right">${fmtMoney(q.totalPagoQz)}</td><td class="right">${fmtMoney(q.totalPagoCte)}</td>
    <td class="right">${fmtQzDiff(q.totalPagoCte - q.totalPagoQz, 0.5)}</td>
    <td class="right">${q.totalCteQz}</td><td class="right">${q.totalCteCte}</td><td class="right">${q.totalCteCte - q.totalCteQz}</td></tr>`;
}

function renderB2cQuinzenaRow(q) {
  return `<tr class="qz-qz-row"><td style="padding-left:26px">${qzQuinzenaCellLabel(q)}</td>
    <td class="right">${q.count}</td><td class="right">${fmtMoney(q.totalVendas)}</td>
    <td class="right">${fmtMoney(q.totalFrete)}</td><td class="right">${fmtPct(q.pctFrete)}</td></tr>`;
}

function bindQzExpandClicks(bodyEl, panel) {
  if (!bodyEl || bodyEl._qzExpandBound) return;
  bodyEl._qzExpandBound = true;
  bodyEl.addEventListener('click', e => {
    const tr = e.target.closest('tr.qz-expandable');
    if (!tr || tr.dataset.panel !== panel) return;
    const key = qzExpandKey(panel, tr.dataset.mes);
    if (qzExpandedMonths.has(key)) qzExpandedMonths.delete(key);
    else qzExpandedMonths.add(key);
    if (panel === 'b2b') renderQzB2b();
    else renderQzB2c();
  });
}

function fmtQzStatus(s) {
  return { match: '<span style="color:var(--green)">OK</span>', diff: '<span style="color:#b3261e">Diferença</span>',
    onlyQz: '<span style="color:#b3261e">Só quinzenal</span>', onlyCte: '<span style="color:#b3261e">Só CT-e</span>',
    qzView: '<span style="color:var(--muted)">Quinzenal</span>' }[s] || s;
}

function populateQzMonthFilters() {
  if (!quinzenalPack) return;
  const keys = [...new Set((quinzenalPack.b2bRows || []).map(r => r.mesKey || mesKeyFromQuinzenaKey(r.quinzenaKey)).filter(k => k && k !== 'sem-mes'))].sort(compareMesRef);
  const sel = $('qzFilterQuinzena');
  if (sel) {
    const cur = sel.value;
    sel.innerHTML = '<option value="">Todos os meses</option>' + keys.map(k => {
      const lbl = quinzenalPack.b2bRows.find(r => (r.mesKey || mesKeyFromQuinzenaKey(r.quinzenaKey)) === k)?.mesLabel || fmtMesLabel(k);
      return `<option value="${k}">${lbl}</option>`;
    }).join('');
    sel.value = keys.includes(cur) ? cur : '';
  }
  const keysB2c = [...new Set((quinzenalPack.b2cRows || []).map(r => resolveB2cRowMesKey(r)).filter(k => k && k !== 'sem-mes'))].sort(compareMesRef);
  const selB2c = $('qzB2cFilterQuinzena');
  if (selB2c) {
    const cur = selB2c.value;
    selB2c.innerHTML = '<option value="">Todos os meses</option>' + keysB2c.map(k => {
      const lbl = quinzenalPack.b2cRows.find(r => resolveB2cRowMesKey(r) === k)?.mesLabel || fmtMesLabel(k);
      return `<option value="${k}">${lbl}</option>`;
    }).join('');
    selB2c.value = keysB2c.includes(cur) ? cur : '';
  }
}

async function renderB2bCompareTab() {
  const empty = $('b2bEmpty');
  const content = $('b2bContent');
  const warn = $('qzWarn');
  if (quinzenalPack?.b2bRows?.length && !hasConciliacaoCte()) {
    await ensureCteLoadedForQz(true);
  }
  if (warn) {
    if (quinzenalPack?.b2bRows?.length && !hasConciliacaoCte()) {
      warn.style.display = 'block';
      warn.innerHTML = '<strong>Vista quinzenal:</strong> a <strong>Análise CT-e</strong> e este confronto usam os quinzenais B2B (pago, devolução, Num. CTE). '
        + 'Carrega também o <strong>ZFACT</strong> para anomalias SAP.';
    } else warn.style.display = 'none';
  }
  const hasB2b = !!(quinzenalPack?.b2bRows?.length);
  if (empty) empty.style.display = hasB2b ? 'none' : 'block';
  if (content) content.style.display = hasB2b ? 'block' : 'none';
  if (!hasB2b) return;
  updateQzFileNote();
  populateQzMonthFilters();
  refreshQuinzenalCompare();
  renderQzB2b();
  // Soften mismatch panels when QZ-only (no Conciliacao side)
  const missingPanels = document.querySelector('.qz-missing-panels');
  if (missingPanels) missingPanels.style.display = hasConciliacaoCte() ? '' : 'none';
  const mismatchTitle = missingPanels?.previousElementSibling;
  // section-title for mismatch is before the panels; hide its following description too via parent query
  const mismatchSection = $('exportB2bMismatchBtn')?.closest('.section-title');
  if (mismatchSection) mismatchSection.style.display = hasConciliacaoCte() ? '' : 'none';
  const mismatchHint = mismatchSection?.nextElementSibling;
  if (mismatchHint && mismatchHint.tagName === 'P') mismatchHint.style.display = hasConciliacaoCte() ? '' : 'none';
}

async function renderB2cAnalysisTab() {
  const empty = $('b2cEmpty');
  const content = $('b2cContent');
  const hasB2c = !!(quinzenalPack?.b2cRows?.length);
  if (empty) empty.style.display = hasB2c ? 'none' : 'block';
  if (content) content.style.display = hasB2c ? 'block' : 'none';
  if (!hasB2c) return;
  const note = $('b2cFileNote');
  if (note) {
    const fc = qzFileCounts();
    if (fc.b2c) {
      note.style.display = 'block';
      note.textContent = `${fc.b2c} ficheiro(s) B2C carregado(s)` + (fc.failed ? ` · ${fc.failed} com erro` : '');
    } else note.style.display = 'none';
  }
  populateQzMonthFilters();
  renderQzB2c();
}

function computeResumoTotal() {
  const cte = currentSummary;
  const b2cRows = quinzenalPack?.b2cRows || [];
  const b2bRows = quinzenalPack?.b2bRows || [];
  const fatCte = cte?.totalValorNF || 0;
  const b2cVendas = b2cRows.reduce((s, r) => s + num(r.valorNF), 0);
  const freteCte = cte?.totalPago || 0;
  const freteB2c = b2cRows.reduce((s, r) => s + num(r.pago), 0);
  const freteQzB2b = b2bRows.reduce((s, r) => s + num(r.pago), 0);
  const freteTotal = freteCte + freteB2c;
  const debitoUnilog = freteTotal;
  const deltaQzB2b = freteCte - freteQzB2b;
  return {
    fatCte, b2cVendas, fatTotal: fatCte + b2cVendas,
    freteCte, freteB2c, freteQzB2b, freteTotal, debitoUnilog, deltaQzB2b,
    hasCte: !!cte, hasB2c: b2cRows.length > 0, hasB2bQz: b2bRows.length > 0
  };
}

function resumoMesKey(raw) {
  return raw === 'sem-data' ? 'sem-mes' : raw;
}

function resumoMesLabel(k, cteLabel) {
  if (k === 'sem-mes') return cteLabel || 'Sem data';
  return cteLabel || fmtMesLabel(k);
}

function enrichResumoMonthPct(t) {
  t.pctFreteB2b = t.fatCte > 0 ? t.freteCte / t.fatCte : null;
  t.pctFreteB2c = t.vendasB2c > 0 ? t.freteB2c / t.vendasB2c : null;
  t.pctFreteTotal = t.fatTotal > 0 ? t.freteTotal / t.fatTotal : null;
  return t;
}

function fmtResumoPct(v) {
  return v != null ? fmtPct(v) : '-';
}

function aggregateResumoMonthGroup(group) {
  const t = { fatCte: 0, vendasB2c: 0, freteCte: 0, freteB2c: 0, freteQzB2b: 0 };
  group.forEach(m => {
    t.fatCte += m.fatCte;
    t.vendasB2c += m.vendasB2c;
    t.freteCte += m.freteCte;
    t.freteB2c += m.freteB2c;
    t.freteQzB2b += m.freteQzB2b;
  });
  t.fatTotal = t.fatCte + t.vendasB2c;
  t.freteTotal = t.freteCte + t.freteB2c;
  t.deltaQzB2b = t.freteCte - t.freteQzB2b;
  return enrichResumoMonthPct(t);
}

function buildResumoMonthDisplayRows(totals) {
  const out = [];
  const dated = totals.filter(m => m.mesKey !== 'sem-mes');
  const noData = totals.filter(m => m.mesKey === 'sem-mes');
  let lastYear = null;
  let yearGroup = [];

  dated.forEach(m => {
    const year = m.mesKey.slice(0, 4);
    if (lastYear && year !== lastYear && yearGroup.length) {
      out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...aggregateResumoMonthGroup(yearGroup) });
      yearGroup = [];
    }
    if (year !== lastYear) {
      out.push({ type: 'year', label: year, year });
      lastYear = year;
    }
    out.push({ type: 'month', ...m });
    yearGroup.push(m);
  });
  if (yearGroup.length && lastYear) {
    out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...aggregateResumoMonthGroup(yearGroup) });
  }
  noData.forEach(m => out.push({ type: 'month', ...m }));
  if (totals.length) {
    out.push({ type: 'total', label: 'Total geral', ...aggregateResumoMonthGroup(totals) });
  }
  return out;
}

function renderResumoMonthTableRow(label, m, rowClass) {
  return `<tr class="${rowClass || 'qz-month-row'}"><td><strong>${label}</strong></td>
    <td class="right">${fmtMoney(m.fatCte)}</td><td class="right">${fmtMoney(m.vendasB2c)}</td><td class="right">${fmtMoney(m.fatTotal)}</td>
    <td class="right">${fmtMoney(m.freteCte)}</td><td class="right">${fmtMoney(m.freteB2c)}</td><td class="right">${fmtMoney(m.freteTotal)}</td>
    <td class="right">${fmtMoney(m.freteQzB2b)}</td>
    <td class="right">${fmtQzDiff(m.deltaQzB2b, 0.5)}</td>
    <td class="right">${fmtResumoPct(m.pctFreteB2b)}</td><td class="right">${fmtResumoPct(m.pctFreteB2c)}</td><td class="right">${fmtResumoPct(m.pctFreteTotal)}</td></tr>`;
}

function buildResumoMonthlyRows() {
  const cteByMes = {};
  (monthlyRows.length ? monthlyRows : computeMonthly(currentNFs)).forEach(m => {
    const k = resumoMesKey(m.mesRef);
    cteByMes[k] = { mesLabel: m.mesLabel, fatCte: m.totalValorNF, freteCte: m.totalPago };
  });
  const b2cByMes = {};
  (quinzenalPack?.b2cMonthTotals || []).forEach(m => {
    b2cByMes[m.mesKey] = { vendas: m.totalVendas, frete: m.totalFrete };
  });
  const b2bQzByMes = {};
  (quinzenalPack?.b2bMonthTotals || []).forEach(m => {
    b2bQzByMes[m.mesKey] = { freteQz: m.totalPagoQz };
  });
  const keys = [...new Set([...Object.keys(cteByMes), ...Object.keys(b2cByMes), ...Object.keys(b2bQzByMes)])]
    .filter(k => k).sort(compareMesRef);
  return keys.map(k => {
    const c = cteByMes[k] || {};
    const b2c = b2cByMes[k] || {};
    const b2b = b2bQzByMes[k] || {};
    const fatCte = c.fatCte || 0;
    const vendasB2c = b2c.vendas || 0;
    const freteCte = c.freteCte || 0;
    const freteB2c = b2c.frete || 0;
    const freteQzB2b = b2b.freteQz || 0;
    return enrichResumoMonthPct({
      mesKey: k,
      mesLabel: resumoMesLabel(k, c.mesLabel),
      fatCte, vendasB2c, fatTotal: fatCte + vendasB2c,
      freteCte, freteB2c, freteTotal: freteCte + freteB2c,
      freteQzB2b, deltaQzB2b: freteCte - freteQzB2b
    });
  });
}

function renderResumoTotal() {
  const r = computeResumoTotal();
  const hasData = r.hasCte || r.hasB2c || r.hasB2bQz;
  const empty = $('resumoEmpty');
  const content = $('resumoContent');
  if (empty) {
    empty.style.display = hasData ? 'none' : 'block';
    if (!hasData) {
      empty.innerHTML = '<div class="empty-i">📈</div><div>Sem dados em memória. Abre esta aba para restaurar da cloud, ou clica <strong>Carregar último guardado</strong> em <strong>Carregamento de dados</strong>.</div>';
    }
  }
  if (content) content.style.display = hasData ? 'block' : 'none';
  if (!hasData) return;

  const kpis = $('resumoKpis');
  if (kpis) {
    kpis.innerHTML = `
      <div class="kpi"><div class="label">Faturação total</div><div class="value">${fmtMoney(r.fatTotal)}</div>
        <div class="sub">CT-e ${fmtMoney(r.fatCte)}${r.hasB2c ? ' + B2C ' + fmtMoney(r.b2cVendas) : ''}</div></div>
      <div class="kpi"><div class="label">Custo frete total</div><div class="value">${fmtMoney(r.freteTotal)}</div>
        <div class="sub">CT-e ${fmtMoney(r.freteCte)}${r.hasB2c ? ' + B2C ' + fmtMoney(r.freteB2c) : ''}</div></div>
      <div class="kpi"><div class="label">Débito Unilog</div><div class="value">${fmtMoney(r.debitoUnilog)}</div>
        <div class="sub">Soma fretes cobrados (CT-e + B2C)</div></div>
      <div class="kpi${Math.abs(r.deltaQzB2b) > 1 ? ' flag' : ''}"><div class="label">Δ CT-e vs QZ B2B</div>
        <div class="value">${fmtMoney(r.deltaQzB2b)}</div><div class="sub">Frete CT-e − frete quinzenal B2B</div></div>`;
  }

  const rec = $('resumoReconcile');
  if (rec) {
    rec.innerHTML = `<strong>Consolidação:</strong> Faturação ${fmtMoney(r.fatTotal)} · Frete total ${fmtMoney(r.freteTotal)} = Débito Unilog ${fmtMoney(r.debitoUnilog)}
      ${r.hasB2bQz ? ` · Frete quinzenal B2B ${fmtMoney(r.freteQzB2b)} (Δ ${fmtMoney(r.deltaQzB2b)} vs CT-e)` : ''}`;
  }

  const body = $('resumoMonthBody');
  if (body) {
    const monthTotals = buildResumoMonthlyRows();
    if (!monthTotals.length) {
      body.innerHTML = '<tr><td colspan="12" class="empty">Sem dados mensais — carrega CT-e e quinzenais</td></tr>';
    } else {
      const displayRows = buildResumoMonthDisplayRows(monthTotals);
      body.innerHTML = displayRows.map(row => {
        if (row.type === 'year') {
          return `<tr class="month-year-row"><td colspan="12"><strong>${row.label}</strong></td></tr>`;
        }
        if (row.type === 'subtotal' || row.type === 'total') {
          const cls = row.type === 'total' ? 'month-total-row' : 'month-subtotal-row';
          return renderResumoMonthTableRow(row.label, row, cls);
        }
        return renderResumoMonthTableRow(row.mesLabel, row, 'qz-month-row');
      }).join('');
    }
    fteEnableDomSort(body);
  }
}

function fmtB2bMissingQzQuinzena(r) {
  const qz = r.quinzenaLabel || r.quinzenaKey || '';
  const fn = r.fileName || '';
  if (fn && fn !== qz) {
    const safe = String(fn).replace(/"/g, '&quot;');
    return `${qz || '-'}<div class="qz-qz-file" title="${safe}">${safe}</div>`;
  }
  return qz || '-';
}

function b2bCompareMissingQz() {
  return (quinzenalPack?.b2bCompare || []).filter(r => r.status === 'onlyQz');
}

function b2bCompareMissingCte() {
  return (quinzenalPack?.b2bCompare || []).filter(r => r.status === 'onlyCte');
}

function b2bMissingQzExportRows(rows) {
  return rows.map(r => ({
    NF: r.nf,
    'Data NF': r.dtNFQz ? new Date(r.dtNFQz) : '',
    Mês: r.mesLabelQz || r.mesLabel || '',
    'Valor NF': r.valorNFQz,
    Frete: r.pagoQz,
    Quinzena: r.quinzenaLabel || '',
    Ficheiro: r.fileName || '',
    Transportador: r.transportador || ''
  }));
}

function b2bMissingCteExportRows(rows) {
  return rows.map(r => ({
    NF: r.nf,
    'Data NF': r.dtNFCte ? new Date(r.dtNFCte) : '',
    Mês: r.mesLabelCte || r.mesLabel || '',
    'Valor NF': r.valorNFCte,
    Frete: r.pagoCte,
    Transportador: r.transportador || ''
  }));
}

function b2bDiffExportRows(rows) {
  return rows.map(r => ({
    NF: r.nf,
    'Data QZ': r.dtNFQz ? new Date(r.dtNFQz) : '',
    'Data CT-e': r.dtNFCte ? new Date(r.dtNFCte) : '',
    Mês: fmtB2bCompareMesCell(r),
    'Valor NF QZ': r.valorNFQz, 'Valor NF CT-e': r.valorNFCte, 'Δ valor NF': r.diffValorNF,
    'Frete QZ': r.pagoQz, 'Frete CT-e': r.pagoCte, 'Δ frete': r.diffPago,
    'Qtd CT-e QZ': r.nCteQz, 'Qtd CT-e CT-e': r.nCteCte, 'Δ CT-e': r.diffCte
  }));
}

function exportB2bMissingQz() {
  if (!quinzenalPack?.b2bRows?.length) { fteToast('Sem dados quinzenais B2B.'); return; }
  refreshQuinzenalCompare();
  const rows = b2bCompareMissingQz();
  if (!rows.length) { fteToast('Sem NFs só no quinzenal.'); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2bMissingQzExportRows(rows)), 'Só quinzenal');
  downloadWorkbook(wb, `B2B_so_quinzenal_${new Date().toISOString().slice(0, 10)}.xlsx`);
  fteToast(`${rows.length} NF(s) exportadas (só quinzenal).`);
}

function exportB2bMissingCte() {
  if (!currentNFs.length) { fteToast('Sem dados CT-e.'); return; }
  refreshQuinzenalCompare();
  const rows = b2bCompareMissingCte();
  if (!rows.length) { fteToast('Sem NFs só no CT-e.'); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2bMissingCteExportRows(rows)), 'Só CT-e');
  downloadWorkbook(wb, `B2B_so_CTe_${new Date().toISOString().slice(0, 10)}.xlsx`);
  fteToast(`${rows.length} NF(s) exportadas (só CT-e).`);
}

function exportB2bMismatchWorkbook() {
  if (!quinzenalPack?.b2bRows?.length && !currentNFs.length) { fteToast('Sem dados para confronto B2B.'); return; }
  refreshQuinzenalCompare();
  const cmp = quinzenalPack?.b2bCompare || [];
  const qz = cmp.filter(r => r.status === 'onlyQz');
  const cte = cmp.filter(r => r.status === 'onlyCte');
  const diff = cmp.filter(r => r.status === 'diff');
  if (!qz.length && !cte.length && !diff.length) { fteToast('Sem dados de confronto B2B.'); return; }
  const wb = XLSX.utils.book_new();
  if (qz.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2bMissingQzExportRows(qz)), 'Só quinzenal');
  if (cte.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2bMissingCteExportRows(cte)), 'Só CT-e');
  if (diff.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2bDiffExportRows(diff)), 'Com diferença');
  downloadWorkbook(wb, `B2B_confronto_Unilog_${new Date().toISOString().slice(0, 10)}.xlsx`);
  fteToast('Workbook confronto Unilog exportado.');
}

function renderQzB2b() {
  const rows = quinzenalPack.b2bCompare || [];
  const monthTotals = quinzenalPack.b2bMonthTotals || [];
  const cteGrand = cteGrandTotalsFromNFs();
  const kpis = $('qzB2bKpis');
  if (kpis) kpis.innerHTML = renderB2bCompareKpis(monthTotals, cteGrand, rows);

  const qzBody = $('qzB2bQuinzenaBody');
  if (qzBody) {
    const displayRows = buildB2bMonthDisplayRows(monthTotals, cteGrand);
    qzBody.innerHTML = displayRows.map(row => {
      if (row.type === 'year') {
        return `<tr class="month-year-row"><td colspan="13"><strong>${row.label}</strong></td></tr>`;
      }
      if (row.type === 'subtotal' || row.type === 'total') {
        const cls = row.type === 'total' ? 'month-total-row' : 'month-subtotal-row';
        return renderB2bMonthTableRow(row.label, row, cls);
      }
      return renderB2bMonthTableRow(row.mesLabel, row, 'qz-month-row');
    }).join('');
  }

  const onlyQz = rows.filter(r => r.status === 'onlyQz');
  const onlyCte = rows.filter(r => r.status === 'onlyCte');
  const qzMissingCount = $('qzMissingQzCount');
  if (qzMissingCount) qzMissingCount.textContent = onlyQz.length ? `(${onlyQz.length} NF${onlyQz.length !== 1 ? 's' : ''})` : '';
  const cteMissingCount = $('qzMissingCteCount');
  if (cteMissingCount) cteMissingCount.textContent = onlyCte.length ? `(${onlyCte.length} NF${onlyCte.length !== 1 ? 's' : ''})` : '';

  const qzMissBody = $('qzMissingQzBody');
  if (qzMissBody) {
    qzMissBody.innerHTML = onlyQz.length ? onlyQz.map(r => `<tr class="qz-only-row">
      <td>${r.nf}</td><td>${fmtDate(r.dtNFQz)}</td><td>${r.mesLabelQz || r.mesLabel || '-'}</td>
      <td class="right">${fmtMoney(r.valorNFQz)}</td><td class="right">${fmtMoney(r.pagoQz)}</td>
      <td>${fmtB2bMissingQzQuinzena(r)}</td></tr>`).join('')
      : '<tr><td colspan="6" class="empty">Nenhuma NF só no quinzenal</td></tr>';
  }

  const cteMissBody = $('qzMissingCteBody');
  if (cteMissBody) {
    cteMissBody.innerHTML = onlyCte.length ? onlyCte.map(r => `<tr class="qz-only-row">
      <td>${r.nf}</td><td>${fmtDate(r.dtNFCte)}</td><td>${r.mesLabelCte || r.mesLabel || '-'}</td>
      <td class="right">${fmtMoney(r.valorNFCte)}</td><td class="right">${fmtMoney(r.pagoCte)}</td>
      <td>${r.transportador || '-'}</td></tr>`).join('')
      : '<tr><td colspan="6" class="empty">Nenhuma NF só no CT-e</td></tr>';
  }

  const qFilter = $('qzFilterQuinzena')?.value || '';
  const dFilter = $('qzFilterDiff')?.value || '';
  const search = ($('qzSearchNF')?.value || '').trim();
  let list = rows;
  if (qFilter) {
    list = list.filter(r =>
      (r.mesKeyQz || r.mesKey || mesKeyFromQuinzenaKey(r.quinzenaKey)) === qFilter ||
      r.mesKeyCte === qFilter
    );
  }
  if (dFilter === 'onlyQz') list = list.filter(r => r.status === 'onlyQz');
  else if (dFilter === 'onlyCte') list = list.filter(r => r.status === 'onlyCte');
  else if (dFilter === 'diff') list = list.filter(r => r.status === 'diff');
  else if (dFilter === 'valor') list = list.filter(r => Math.abs(r.diffValorNF) > 1);
  else if (dFilter === 'frete') list = list.filter(r => Math.abs(r.diffPago) > 0.5);
  else if (dFilter === 'cte') list = list.filter(r => r.diffCte !== 0);
  if (search) list = list.filter(r => String(r.nf).includes(search));

  const detBody = $('qzB2bDetailBody');
  if (detBody) {
    detBody.innerHTML = list.slice(0, 2000).map(r => {
      const cls = r.status === 'onlyQz' || r.status === 'onlyCte' ? 'qz-only-row' : (r.status === 'diff' ? 'qz-diff-row' : '');
      const pctNote = r.valorNFQz > 0 && r.pagoQz > 0 ? fmtPct(r.pagoQz / r.valorNFQz) : '-';
      return `<tr class="${cls}"><td>${r.nf}</td>
        <td>${fmtB2bCompareMesCell(r)}</td><td>${fmtB2bCompareDateCell(r)}</td>
        <td class="right">${r.valorNFQz ? fmtMoney(r.valorNFQz) : '-'}</td><td class="right">${r.valorNFCte ? fmtMoney(r.valorNFCte) : '-'}</td>
        <td class="right">${fmtQzDiff(r.diffValorNF, 1)}</td>
        <td class="right">${r.pagoQz ? fmtMoney(r.pagoQz) : '-'}</td><td class="right">${r.pagoCte ? fmtMoney(r.pagoCte) : '-'}</td>
        <td class="right">${fmtQzDiff(r.diffPago, 0.5)}</td>
        <td class="right">${r.nCteQz || '-'}</td><td class="right">${r.nCteCte || '-'}</td><td class="right">${r.diffCte || 0}</td>
        <td>${fmtQzStatus(r.status)}</td><td class="right">${pctNote}</td></tr>`;
    }).join('') || '<tr><td colspan="14" class="empty">Sem dados B2B — carrega quinzenais B2B</td></tr>';
  }
  fteEnableDomSort(qzBody);
  fteEnableDomSort(qzMissBody);
  fteEnableDomSort(cteMissBody);
  fteEnableDomSort(detBody);

  // Soften CT-e columns emphasis when QZ-only
  const b2bContent = $('b2bContent');
  if (b2bContent) b2bContent.classList.toggle('qz-only-view', !hasConciliacaoCte());
}

function renderQzB2c() {
  const rows = quinzenalPack.b2cRows || [];
  const monthTotals = quinzenalPack.b2cMonthTotals || [];
  const yearTotals = buildB2CYearTotals(monthTotals);
  const grandTotal = aggregateB2cMonthGroup(monthTotals);
  const kpis = $('qzB2cKpis');
  if (kpis) {
    let kpiHtml = yearTotals.map(y => renderB2cKpiBlock(y.label, y)).join('');
    if (yearTotals.length > 1) kpiHtml += renderB2cKpiBlock('Total geral', grandTotal);
    kpis.innerHTML = kpiHtml || renderB2cKpiBlock('Total geral', { count: 0, totalVendas: 0, totalFrete: 0, pctFrete: 0 });
  }

  const regBody = $('qzB2cRegionBody');
  if (regBody) {
    const regions = quinzenalPack.b2cRegionTotals || buildB2CRegionTotals(rows);
    regBody.innerHTML = regions.map(r => `<tr>
      <td><strong>${r.regiao}</strong></td>
      <td class="right">${r.count}</td>
      <td class="right">${fmtMoney(r.totalVendas)}</td>
      <td class="right">${fmtMoney(r.totalFrete)}</td>
      <td class="right">${fmtPct(r.pctFrete)}</td>
      <td class="right">${fmtMoney(r.freteMedio)}</td></tr>`).join('')
      || '<tr><td colspan="6" class="empty">Sem dados por região</td></tr>';
  }

  const qBody = $('qzB2cQuinzenaBody');
  if (qBody) {
    const displayRows = buildB2cMonthDisplayRows(monthTotals);
    qBody.innerHTML = displayRows.map(row => {
      if (row.type === 'year') {
        return `<tr class="month-year-row"><td colspan="6"><strong>${row.label}</strong></td></tr>`;
      }
      if (row.type === 'subtotal' || row.type === 'total') {
        const cls = row.type === 'total' ? 'month-total-row' : 'month-subtotal-row';
        return renderB2cMonthTableRow(row.label, row, cls);
      }
      return renderB2cMonthTableRow(row.mesLabel, row, 'qz-month-row');
    }).join('');
  }

  const qFilter = $('qzB2cFilterQuinzena')?.value || '';
  const search = ($('qzB2cSearch')?.value || '').toLowerCase();
  let list = rows;
  if (qFilter) list = list.filter(r => resolveB2cRowMesKey(r) === qFilter);
  if (search) list = list.filter(r => String(r.nf || '').includes(search) ||
    String(r.destinatario || '').toLowerCase().includes(search) ||
    String(r.zona || '').toLowerCase().includes(search));

  const detBody = $('qzB2cDetailBody');
  if (detBody) {
    detBody.innerHTML = list.slice(0, 2000).map(r => `<tr>
      <td>${r.nf || '-'}</td>
      <td>${fmtDate(r.dtColeta || r.dtNF)}</td><td>${r.destinatario || '-'}</td><td>${r.transportador || '-'}</td>
      <td class="right">${fmtMoney(r.valorNF)}</td><td class="right">${fmtMoney(r.pago)}</td>
      <td class="right">${fmtPct(r.valorNF ? r.pago / r.valorNF : 0)}</td><td>${r.zona || '-'}</td></tr>`).join('')
      || '<tr><td colspan="8" class="empty">Sem dados B2C</td></tr>';
  }
  fteEnableDomSort(regBody);
  fteEnableDomSort(qBody);
  fteEnableDomSort(detBody);
}

function exportQuinzenalWorkbook() {
  if (!quinzenalPack?.files?.length) { fteToast('Sem dados quinzenais.'); return; }
  refreshQuinzenalCompare();
  const wb = XLSX.utils.book_new();
  const resumo = (quinzenalPack.b2bMonthTotals || []).map(t => ({
    Mês: t.mesLabel, 'NFs QZ': t.nfCountQz, 'NFs CT-e': t.nfCountCte, 'Δ NFs': t.nfCountCte - t.nfCountQz,
    'Faturação QZ': t.totalValorNFQz, 'Faturação CT-e': t.totalValorNFCte, 'Δ valor NF': t.totalValorNFCte - t.totalValorNFQz,
    'Frete QZ': t.totalPagoQz, 'Frete CT-e': t.totalPagoCte, 'Δ frete': t.totalPagoCte - t.totalPagoQz,
    'CT-e QZ': t.totalCteQz, 'CT-e CT-e': t.totalCteCte, 'Δ CT-e': t.totalCteCte - t.totalCteQz
  }));
  if (resumo.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumo), 'Resumo B2B mensal');
  const b2b = (quinzenalPack.b2bCompare || []).map(r => ({
    NF: r.nf, Mês: fmtB2bCompareMesCell(r),
    'Data QZ': r.dtNFQz ? new Date(r.dtNFQz) : '',
    'Data CT-e': r.dtNFCte ? new Date(r.dtNFCte) : '',
    'Valor NF QZ': r.valorNFQz, 'Valor NF CT-e': r.valorNFCte, 'Δ valor NF': r.diffValorNF,
    'Frete QZ': r.pagoQz, 'Frete CT-e': r.pagoCte, 'Δ frete': r.diffPago,
    'Qtd CT-e QZ': r.nCteQz, 'Qtd CT-e CT-e': r.nCteCte, 'Δ CT-e': r.diffCte, Estado: r.status,
    '% pago QZ (6%)': r.valorNFQz > 0 ? r.pagoQz / r.valorNFQz : null
  }));
  if (b2b.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2b), 'Diff B2B');
  const b2cMes = (quinzenalPack.b2cMonthTotals || []).map(t => ({
    Mês: t.mesLabel, Entregas: t.count, Vendas: t.totalVendas, 'Frete pago': t.totalFrete,
    '% frete/vendas': t.pctFrete, 'Frete médio': t.freteMedio ?? (t.count ? t.totalFrete / t.count : 0)
  }));
  if (b2cMes.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2cMes), 'Resumo B2C mensal');
  const b2cReg = (quinzenalPack.b2cRegionTotals || buildB2CRegionTotals(quinzenalPack.b2cRows || [])).map(r => ({
    Região: r.regiao, Entregas: r.count, Vendas: r.totalVendas, 'Frete pago': r.totalFrete,
    '% frete/vendas': r.pctFrete, 'Frete médio': r.freteMedio ?? (r.count ? r.totalFrete / r.count : 0)
  }));
  if (b2cReg.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2cReg), 'Resumo B2C região');
  const b2c = (quinzenalPack.b2cRows || []).map(r => ({
    NF: r.nf, 'Data coleta': r.dtColeta || r.dtNF,
    Destinatário: r.destinatario, Transportador: r.transportador, Vendas: r.valorNF, Frete: r.pago,
    '% frete/vendas': r.valorNF ? r.pago / r.valorNF : null, Região: r.zona
  }));
  if (b2c.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2c), 'B2C');
  downloadWorkbook(wb, `Unilog_quinzenal_${new Date().toISOString().slice(0, 10)}.xlsx`);
  fteToast('Excel quinzenal exportado.');
}

function reloadFretesForCompany() {
  _fteLoadedCompany = null;
  currentNFs = [];
  currentSummary = null;
  currentUploadId = null;
  sapNfMap = {};
  sapMapSourceFiles = [];
  invalidateZfactMonthsCache();
  lastCtePack = null;
  cteAnalysisSource = null;
  activeSubPanel = null;
  activeCteSub = 'resumo';
  selectedMonth = '';
  fteCteFileName = '';
  fteCteBuffer = null;
  fteSapFileName = '';
  fteSapBuffer = null;
  fteQzPendingFiles = [];
  quinzenalPack = null;
  qzExpandedMonths = new Set();
  syncQzUploadZone();
  setCteZoneLoaded('');
  setSapZoneLoaded('');
  setLoadbar('');
  updateSaveStatus('');
  updateFretesFileStatus();
  switchFteTab('carregamento');
  // Labels only — never auto-parse multi-MB ZFACT/QZ on company switch.
  if (typeof applyFretesFileLabelsFromMeta === 'function') {
    applyFretesFileLabelsFromMeta().catch(() => {});
  }
}

function initFretes() {
  if (fteInited) return;
  fteInited = true;

  document.querySelectorAll('.fte-tab').forEach(t => {
    t.addEventListener('click', () => switchFteTab(t.dataset.tab));
  });

  document.querySelectorAll('.fte-cte-subtabs .qz-subtab').forEach(st => {
    st.addEventListener('click', () => switchCteSub(st.dataset.cteSub));
  });

  const cteZone = $('cteZone');
  const fileInput = $('fileInput');
  if (cteZone && fileInput) {
    cteZone.addEventListener('click', () => fileInput.click());
    cteZone.addEventListener('dragover', e => { e.preventDefault(); cteZone.classList.add('drag'); });
    cteZone.addEventListener('dragleave', () => cteZone.classList.remove('drag'));
    cteZone.addEventListener('drop', e => {
      e.preventDefault(); cteZone.classList.remove('drag');
      if (e.dataTransfer.files.length) selectCteFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', e => { if (e.target.files.length) selectCteFile(e.target.files[0]); });
  }

  const sapZone = $('sapZone');
  const sapFileInput = $('sapFileInput');
  if (sapZone && sapFileInput) {
    sapZone.addEventListener('click', () => sapFileInput.click());
    sapZone.addEventListener('dragover', e => { e.preventDefault(); sapZone.classList.add('drag'); });
    sapZone.addEventListener('dragleave', () => sapZone.classList.remove('drag'));
    sapZone.addEventListener('drop', e => {
      e.preventDefault(); sapZone.classList.remove('drag');
      if (e.dataTransfer.files.length) selectSapFile(e.dataTransfer.files[0]);
    });
    sapFileInput.addEventListener('change', e => { if (e.target.files.length) selectSapFile(e.target.files[0]); });
  }

  $('procBtn')?.addEventListener('click', () => processAndSaveFretes());
  $('nfMoreBtn')?.addEventListener('click', () => showMoreNfRows());
  $('loadLastBtn')?.addEventListener('click', async () => {
    const ok = await loadSavedFretesFiles(false);
    if (ok && currentNFs.length) switchFteTab('analise-cte');
    else if (ok && quinzenalPack?.files?.length) switchFteTab(quinzenalPack.b2cRows?.length ? 'analise-b2c' : 'cte-vs-qz');
  });

  const qzZone = $('qzZone');
  const qzInput = $('qzFileInput');
  if (qzZone && qzInput) {
    qzZone.addEventListener('click', () => qzInput.click());
    qzZone.addEventListener('dragover', e => { e.preventDefault(); qzZone.classList.add('drag'); });
    qzZone.addEventListener('dragleave', () => qzZone.classList.remove('drag'));
    qzZone.addEventListener('drop', e => {
      e.preventDefault(); qzZone.classList.remove('drag');
      if (e.dataTransfer.files.length) {
        fteQzPendingFiles = Array.from(e.dataTransfer.files);
        syncQzUploadZone();
      }
    });
    qzInput.addEventListener('change', e => {
      if (e.target.files.length) {
        fteQzPendingFiles = Array.from(e.target.files);
        syncQzUploadZone();
      }
    });
  }

  ['qzFilterQuinzena', 'qzFilterDiff', 'qzSearchNF'].forEach(id => {
    const el = $(id);
    if (!el) return;
    const fn = () => renderB2bCompareTab();
    el.addEventListener('input', fn);
    el.addEventListener('change', fn);
  });
  const qzKpiRoot = $('qzB2bKpis');
  if (qzKpiRoot && !qzKpiRoot._qzFilterBound) {
    qzKpiRoot._qzFilterBound = true;
    const applyKpiFilter = filter => {
      const sel = $('qzFilterDiff');
      if (!sel) return;
      sel.value = sel.value === filter ? '' : filter;
      renderQzB2b();
    };
    qzKpiRoot.addEventListener('click', e => {
      const kpi = e.target.closest('.qz-kpi-filter');
      if (!kpi?.dataset.qzFilter) return;
      applyKpiFilter(kpi.dataset.qzFilter);
    });
    qzKpiRoot.addEventListener('keydown', e => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      const kpi = e.target.closest('.qz-kpi-filter');
      if (!kpi?.dataset.qzFilter) return;
      e.preventDefault();
      applyKpiFilter(kpi.dataset.qzFilter);
    });
  }
  ['qzB2cFilterQuinzena', 'qzB2cSearch'].forEach(id => {
    const el = $(id);
    if (!el) return;
    const fn = () => renderB2cAnalysisTab();
    el.addEventListener('input', fn);
    el.addEventListener('change', fn);
  });

  $('newFileBtn')?.addEventListener('click', () => switchFteTab('carregamento'));
  $('exportBtn')?.addEventListener('click', () => exportFullWorkbook());
  $('exportFilteredBtn')?.addEventListener('click', () => exportFilteredView());
  $('exportMensalBtn')?.addEventListener('click', () => exportMonthlyWorkbook());
  $('exportQzBtn')?.addEventListener('click', () => exportQuinzenalWorkbook());
  $('exportQzB2cBtn')?.addEventListener('click', () => exportQuinzenalWorkbook());
  $('exportB2bMissingQzBtn')?.addEventListener('click', () => exportB2bMissingQz());
  $('exportB2bMissingCteBtn')?.addEventListener('click', () => exportB2bMissingCte());
  $('exportB2bMismatchBtn')?.addEventListener('click', () => exportB2bMismatchWorkbook());

  ['filterStatus', 'filterTransp', 'filterMes', 'searchNF', 'filterValorDiff', 'filterSapMissing'].forEach(id => {
    const el = $(id);
    if (!el) return;
    el.addEventListener('input', () => {
      if (id === 'filterStatus' && $('filterStatus').value !== activeSubPanel) {
        activeSubPanel = $('filterStatus').value || null;
        document.querySelectorAll('.bd-card').forEach(c => {
          c.classList.toggle('active', c.dataset.status === activeSubPanel);
        });
        renderSubPanels();
      }
      if (id === 'filterMes') {
        selectedMonth = $('filterMes').value || '';
      }
      renderTable();
    });
    el.addEventListener('change', () => {
      if (id === 'filterStatus' && $('filterStatus').value !== activeSubPanel) {
        activeSubPanel = $('filterStatus').value || null;
        document.querySelectorAll('.bd-card').forEach(c => {
          c.classList.toggle('active', c.dataset.status === activeSubPanel);
        });
        renderSubPanels();
      }
      if (id === 'filterMes') {
        selectedMonth = $('filterMes').value || '';
      }
      renderTable();
    });
  });

  // Labels only on init — never auto-start "A processar ZFACT…" (freezes Chrome).
  // Full restore: "Carregar último guardado" or opening an analysis tab.
  applyFretesFileLabelsFromMeta().catch(() => {});
}

/** True when this company has no Fretes analysis data in memory yet (needs cloud load). */
function fteNeedsCloudReload() {
  const co = fteCompany();
  // SAP map alone is shared with Armazém — never treat it as "Fretes already restored".
  if (_fteLoadedCompany === co && fteHasAnalysisData()) {
    return false;
  }
  // Trust in-memory QZ/CTE even if mark was missed (pre-1.8.67 bug)
  if (fteHasAnalysisData()) {
    if (_fteLoadedCompany === co || _fteLoadedCompany == null) {
      fteMarkCompanyLoaded();
      return false;
    }
  }
  return true;
}

/** Shared SAP NF map for Fretes + Armazém modules. */
/**
 * Monthly frete rows for Total Custo Unilog.
 * Source of truth for costs = quinzenais Unilog:
 *   B2B = b2bMonthTotals.totalPagoQz (Gestão de Frete DELTA B2B)
 *   B2C = b2cMonthTotals.totalFrete  (Gestão de Frete DELTA B2C)
 * CT-e totalPago is kept for reference / Δ only — not used as primary cost.
 * Fallback to CT-e freteCte only when quinzenal B2B is empty for that month.
 */
function getFretesMonthlyCustoRows() {
  try {
    return (typeof buildResumoMonthlyRows === 'function' ? buildResumoMonthlyRows() : []).map(m => {
      const freteCte = m.freteCte || 0;
      const freteQzB2b = m.freteQzB2b || 0;
      const freteB2b = freteQzB2b > 0 ? freteQzB2b : freteCte;
      const freteB2c = m.freteB2c || 0;
      return {
        mesKey: m.mesKey,
        mesLabel: m.mesLabel,
        freteB2b,
        freteCte,
        freteQzB2b,
        freteB2c,
        fretePago: freteB2b,
        label: freteQzB2b > 0
          ? 'Fretes B2B (quinzenal) + B2C (quinzenal)'
          : 'Fretes B2B (CT-e fallback) + B2C (quinzenal)'
      };
    });
  } catch (e) {
    console.warn('[fretes] getFretesMonthlyCustoRows', e);
    return [];
  }
}

function refreshCustoUnilogIfVisible() {
  try {
    if (document.getElementById('page-custoUnilog')?.classList.contains('active')
      && typeof window.renderCustoUnilog === 'function') {
      window.renderCustoUnilog();
    }
  } catch (e) {
    console.warn('[fretes] refreshCustoUnilogIfVisible', e);
  }
}

window.getFretesMonthlyCustoRows = getFretesMonthlyCustoRows;
window.initFretes = initFretes;
window.loadSavedFretesFiles = loadSavedFretesFiles;
window.FretesSAP = {
  getMap: () => sapNfMap,
  normNFKey,
  parseSapNum,
  buildSapNfMap,
  mergeSapNfMaps,
  lookupSapEntry,
  isRelevantValorDiff,
  isSapLoaded,
  loadSapRowsFromWorkbook,
  ensureLoaded: (silent = true) => loadSavedFretesFiles(silent).then(() => isSapLoaded()),
  restoreSapFromRec,
  persistSapNfMap,
  loadSavedSapNfMap,
  repairSapNfMapInPlace,
  scrubSharedGhostSapFaces,
  refinalizeSapMapFromStoredFields,
  aggregateNfPago
};
window.fteNeedsCloudReload = fteNeedsCloudReload;
