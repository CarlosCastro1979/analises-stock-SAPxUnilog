// fretes.js v1.5.7
const FRETES_JS_VERSION = '1.5.7';

const sb = db;

const $ = id => document.getElementById('fte-' + id);

let fteInited = false;
let sapNfMap = {};
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
let historyCache = [];
let _fteSkipAutosave = false;
let _fteLoadedCompany = null;
let fteCteFileName = '';
let fteCteBuffer = null;
let fteSapFileName = '';
let fteSapBuffer = null;
let fteQzPendingFiles = [];
let quinzenalPack = null;
let activeQzPanel = 'b2b';
let _fteLoadSavedPromise = null;
let qzExpandedMonths = new Set();

function fteCompany() {
  return typeof company !== 'undefined' ? company : 'DFB';
}

function fteCteSlot() {
  return typeof EXCEL_SLOTS !== 'undefined' ? EXCEL_SLOTS.FRETES_CTE : 'fretes_cte';
}

function fteSapSlot() {
  return typeof EXCEL_SLOTS !== 'undefined' ? EXCEL_SLOTS.FRETES_SAP_NF : 'fretes_sap_nf';
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
  if (fn) fn.textContent = name ? '✓ ' + name : '';
  zone?.classList.toggle('loaded', !!name);
}

function checkFteBtn() {
  const btn = $('procBtn');
  if (btn) btn.disabled = !fteCteBuffer;
}

async function applyFretesFileLabelsFromMeta() {
  if (typeof fetchExcelFiles !== 'function') return;
  try {
    const m = await fetchExcelFiles([fteCteSlot(), fteSapSlot(), fteQuinzenalSlot()]);
    const cte = m[fteCteSlot()];
    const sap = m[fteSapSlot()];
    if (cte?.file_name) setCteZoneLoaded(cte.file_name);
    if (sap?.file_name) setSapZoneLoaded(sap.file_name);
    updateFretesFileStatus(m);
  } catch (e) { /* offline */ }
}

function quinzenalStatusFromPack(pack) {
  if (!pack?.files?.length) return '';
  const b2c = pack.files.filter(f => f.canal === 'B2C').length;
  const b2b = pack.files.filter(f => f.canal === 'B2B').length;
  return `Últimos quinzenais: ${b2c} B2C · ${b2b} B2B`;
}

function parseQuinzenalPackFromRec(rec) {
  if (!rec?.file_data || typeof base64ToArrayBuffer !== 'function') return null;
  try {
    return JSON.parse(new TextDecoder().decode(base64ToArrayBuffer(rec.file_data)));
  } catch (_) {
    return null;
  }
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
    const qzPart = quinzenalStatusFromPack(quinzenalPack) ||
      quinzenalStatusFromPack(parseQuinzenalPackFromRec(m?.[fteQuinzenalSlot()]));
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
  if (typeof fetchExcelFiles !== 'function') return;
  fetchExcelFiles([fteCteSlot(), fteSapSlot(), fteQuinzenalSlot()]).then(apply).catch(() => {
    el.style.display = 'none';
  });
}

async function persistFretesFile(slot, fileName, arrayBuffer) {
  if (typeof upsertExcelBinary !== 'function') {
    fteToastError('Persistência Excel indisponível — recarrega a página.');
    return;
  }
  try {
    await upsertExcelBinary(slot, fileName, arrayBuffer);
    updateFretesFileStatus();
  } catch (err) {
    console.error('[fretes] persist', err);
    if (typeof isExcelFilesTableMissing === 'function' && isExcelFilesTableMissing(err)) {
      fteToastError('Tabela logistica_excel_files em falta no Supabase.');
    } else {
      fteToastError('Erro ao guardar Excel: ' + (err.message || err));
    }
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
  { key: 'diff', label: 'Diferença', type: 'number', right: true },
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
  { key: 'diff', label: 'Diferença', type: 'number', right: true },
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
  { key: 'qtd_min', label: 'Tarifa mín.', type: 'number', right: true },
  { key: 'qtd_flag', label: 'Investigar', type: 'number', right: true }
];

const MESES_PT = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];

const FIELD_ALIASES = {
  nf: ['nota fiscal', 'nf', 'n nota fiscal', 'num nota fiscal', 'num. nota fiscal', 'nota fiscal numero', 'notafiscal'],
  valorNF: ['valor nf', 'valor da nf', 'valor nota fiscal', 'vl nf', 'valor nf r$'],
  transportador: ['transportador', 'transportadora', 'nome transportador', 'transp'],
  modalidade: ['modalidade', 'modal', 'mod'],
  dtNF: ['dt nf', 'data nf', 'data nota fiscal', 'dt nota fiscal'],
  numCte: ['num. cte', 'num cte', 'numero cte', 'n cte-e', 'num cte-e', 'numero do cte', 'numero do cte-e', 'chave cte', 'chave cte-e', 'numero cte-e', 'n cte e', 'num cte e', 'no cte-e', 'n. cte-e'],
  qtdCte: ['qtd cte', 'qtd. cte', 'qtd ct-e', 'qtd. ct-e', 'quantidade cte', 'quantidade ct-e', 'quantidade de cte', 'quantidade de ct-e', 'n ctes', 'numero ctes', 'total cte', 'total ct-e', 'total de cte', 'qtd cobranca', 'qtd cobrança', 'quantidade cobranca', 'quantidade cobrança', 'nº cte', 'no cte', 'n. cte', 'n cte'],
  dtCte: ['dt cte', 'data cte', 'dt cte-e', 'data cte-e', 'data do cte'],
  pago: ['total fatura rev.', 'total fatura rev', 'total fatura', 'valor fatura', 'valor pago', 'total pago', 'vl pago', 'frete pago', 'valor cte', 'valor do frete'],
  devolucao: ['devolucao', 'devolução', 'e devolucao', 'retorno'],
  tipoOp: ['tipo operacao', 'tipo operação', 'tipo op', 'operacao'],
  peso: ['peso', 'peso kg', 'peso (kg)']
};

const SAP_ALIASES = {
  nf: ['nota fiscal', 'nf', 'n nota fiscal', 'num nota fiscal', 'num. nota fiscal', 'notafiscal', 'nº nf', 'numero nf', 'docnum'],
  dtEmissao: ['dt emissao', 'data emissao', 'dt emissão', 'data emissão', 'data doc', 'dt nf', 'data nf', 'data nota fiscal'],
  cliente: ['cliente', 'nome cliente', 'razao social', 'razão social', 'destinatario', 'destinatário', 'cardname', 'nome do cliente'],
  valorNF: ['valor nf', 'valor da nf', 'valor nota fiscal', 'vl nf', 'valor total', 'total nf', 'valor documento', 'montante', 'valor']
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

function showResultsView() {
  $('loadbar').style.display = 'none';
  const us = $('uploadSection');
  if (us) us.style.display = 'none';
  $('results').style.display = 'block';
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

function findSapField(row, field) {
  const aliases = SAP_ALIASES[field];
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
  const dotParts = s.split('.');
  if (dotParts.length === 3 && dotParts.every(p => /^\d+$/.test(p))) {
    const dd = parseInt(dotParts[0], 10), mo = parseInt(dotParts[1], 10), yr = parseInt(dotParts[2], 10);
    const d = sapDateFromParts(yr, mo, dd);
    if (d) return d;
  }
  const slashParts = s.split('/');
  if (slashParts.length >= 3) {
    const yr = parseInt(String(slashParts[2]).trim(), 10);
    const mo = parseInt(slashParts[1], 10);
    const dd = parseInt(slashParts[0], 10);
    const d = sapDateFromParts(yr, mo, dd);
    if (d) return d;
  }
  const m = s.match(/^(\d{2})[.\-/](\d{2})[.\-/](\d{4})$/) || s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const yr = m[1].length === 4 ? parseInt(m[1], 10) : parseInt(m[3], 10);
    const mo = parseInt(m[2], 10);
    const dd = m[1].length === 4 ? parseInt(m[3], 10) : parseInt(m[1], 10);
    const d = sapDateFromParts(yr, mo, dd);
    if (d) return d;
  }
  const serial = parseFloat(s.replace(',', '.'));
  if (!isNaN(serial) && serial >= 1 && serial < 1000000) {
    const d = sapDateFromExcelSerial(serial);
    if (d) return d;
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
  const valorRaw = idx.valorNF != null ? line[idx.valorNF] : null;

  if (nf !== null && nf !== undefined && String(nf).trim() !== '') row.nf = nf;
  if (cliente !== null && cliente !== undefined && String(cliente).trim() !== '') {
    row.cliente = String(cliente).trim();
  }
  const parsed = parseSapBrDate(dtRaw);
  if (parsed) {
    row.dtEmissao = parsed;
    row._sapDateSource = 'colC';
  }
  if (valorRaw !== null && valorRaw !== undefined && valorRaw !== '' && !row.valorNF) {
    row.valorNF = valorRaw;
  }
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

function scoreSapLayout(dataLines, colIdx) {
  let hits = 0;
  for (const line of (dataLines || []).slice(0, 15)) {
    if (looksLikeSapDataRow(line, colIdx)) hits++;
  }
  return hits;
}

function resolveSapColIdx(headerRow, dataLines) {
  const layouts = [
    { cliente: 1, dtEmissao: 2, nf: 3, valorNF: 4, label: 'B/C/D/E' },
    { cliente: 0, dtEmissao: 1, nf: 2, valorNF: 3, label: 'A/B/C/D' },
    { cliente: 2, dtEmissao: 3, nf: 4, valorNF: 5, label: 'C/D/E/F' }
  ];
  for (const layout of layouts) {
    if (headerRow &&
        isSapHeaderAtCol(headerRow, layout.cliente, 'cliente') &&
        isSapHeaderAtCol(headerRow, layout.dtEmissao, 'dtEmissao') &&
        isSapHeaderAtCol(headerRow, layout.nf, 'nf')) {
      console.debug('[SAP NF] Layout por cabeçalho:', layout.label);
      return layout;
    }
  }
  let best = layouts[0];
  let bestScore = 0;
  for (const layout of layouts) {
    const s = scoreSapLayout(dataLines, layout);
    if (s > bestScore) {
      bestScore = s;
      best = layout;
    }
  }
  if (bestScore >= 2) {
    console.debug('[SAP NF] Layout detectado por dados:', best.label, `(${bestScore} linhas)`);
    return best;
  }
  console.debug('[SAP NF] Layout padrão B/C/D (fallback)');
  return { ...SAP_COL_IDX_DEFAULT };
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
  if (standardLayout) {
    console.debug('[SAP NF] Layout padrão: col. B=cliente, C=data emissão, D=nota fiscal');
  }

  const rows = [];
  for (let i = 0; i < dataLines.length; i++) {
    const line = dataLines[i];
    const obj = {};
    headerRow.forEach((h, j) => { if (h) obj[h] = line[j] ?? null; });
    let row = normalizeSapRow(obj);
    row = applySapColPositionalFallback(row, line, headerRow, standardLayout, activeSapColIdx);
    if (i < 3) logSapRowDebug(line, row, i, activeSapColIdx);
    rows.push(row);
  }
  return { rows, headers };
}

function normalizeSapRow(row) {
  const dtRaw = findSapField(row, 'dtEmissao');
  return {
    nf: findSapField(row, 'nf'),
    dtEmissao: parseSapBrDate(dtRaw),
    cliente: findSapField(row, 'cliente'),
    valorNF: findSapField(row, 'valorNF')
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

function buildSapNfMap(rows) {
  const map = {};
  rows.forEach(r => {
    const nf = r.nf;
    if (nf === null || nf === undefined || String(nf).trim() === '') return;
    const key = normNFKey(nf);
    if (!key) return;
    const entry = {
      nf,
      cliente: r.cliente ? String(r.cliente).trim() : '',
      dtEmissao: parseSapBrDate(r.dtEmissao),
      valorNF: num(r.valorNF)
    };
    map[key] = entry;
  });
  return map;
}

/** Relevant SAP vs Unilog NF value gap: abs diff > R$1 AND > 0.5% of the larger value. */
const SAP_VALOR_DIFF_MIN_ABS = 1.0;
const SAP_VALOR_DIFF_MIN_PCT = 0.005;

function isRelevantValorDiff(unilogVal, sapVal) {
  const u = num(unilogVal);
  const s = num(sapVal);
  const diff = Math.abs(s - u);
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

  const sapDt = parseSapBrDate(sap.dtEmissao);
  if (sapDt) {
    nf.dtNF = sapDt;
    delete nf.mesRef;
    delete nf.dtRef;
  }

  nf.valorSAP = sap.valorNF;
  const unilogVal = nf.valorUnilog;

  if ((!nf.valorNF || nf.valorNF === 0) && sap.valorNF) {
    nf.valorNF = sap.valorNF;
    nf.esperado = nf.valorNF * CTE_PCT_TARGET;
    nf.diff = nf.pago - nf.esperado;
    nf.pct = nf.valorNF > 0 ? nf.pago / nf.valorNF : 0;
  }

  nf.valorDiff = sap.valorNF - unilogVal;
  nf.sapValorMismatch = isRelevantValorDiff(unilogVal, sap.valorNF);

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
    processRows(rows, fileName, sheetName, headers);
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

function processArrayBufferSap(arrayBuffer, fileName, opts = {}) {
  try {
    const wb = readSapWorkbookFromArrayBuffer(arrayBuffer);
    const { rows, sheetName } = loadSapRowsFromWorkbook(wb);
    if (!rows.length) {
      const msg = 'Nenhuma linha SAP encontrada na folha "' + sheetName + '".';
      if (!opts.silent) setSapLoadStatus(msg, false);
      return false;
    }
    sapNfMap = buildSapNfMap(rows);
    const nMapped = Object.keys(sapNfMap).length;
    if (!opts.silent) {
      console.log('[SAP] Map size:', nMapped, 'keys (canonical NF)');
    }

    reEnrichAfterSapLoad();

    const nMatched = currentSummary?.nSapMatched ?? countSapMatches(currentNFs);
    if (!opts.silent) {
      logSapKeyDebug(currentNFs);
      setSapLoadStatus(
        nMapped + ' NFs SAP mapeadas · ' + nMatched + ' matched com Unilog',
        nMatched > 0 || !currentNFs.length
      );
      if (nMatched > 0) {
        fteToast('Dados SAP carregados — ' + nMatched + ' NF(s) cruzadas com Unilog.');
        const nMismatch = currentSummary?.nValorMismatch || 0;
        if (nMismatch) fteToast(nMismatch + ' NF(s) com diferença relevante de valor SAP vs Unilog.');
        const nMissing = currentSummary?.nSapMissing || 0;
        if (nMissing) fteToast(nMissing + ' NF(s) no Unilog sem correspondência no SAP — ver Anomalias.');
      } else if (currentNFs.length) {
        setSapLoadStatus(nMapped + ' NFs SAP mapeadas · 0 matched com Unilog', false);
        fteToastError('SAP carregado mas nenhuma NF cruzou com Unilog — verifica formato/chaves.');
      } else {
        fteToast('Dados SAP carregados (' + nMapped + ' NFs). Carrega o Excel Unilog para cruzar.');
      }
    }
    return true;
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
  const reader = new FileReader();
  reader.onload = (e) => {
    fteCteBuffer = e.target.result;
    setCteZoneLoaded(file.name);
    fteToast('CT-e seleccionado — clica Processar e Guardar');
  };
  reader.onerror = () => fteToastError('Erro ao ler ficheiro CT-e.');
  reader.readAsArrayBuffer(file);
}

function selectSapFile(file) {
  if (!file) return;
  fteSapFileName = file.name;
  const reader = new FileReader();
  reader.onload = (e) => {
    fteSapBuffer = e.target.result;
    setSapZoneLoaded(file.name);
    fteToast('SAP NF seleccionado');
  };
  reader.onerror = () => fteToastError('Erro ao ler ficheiro SAP.');
  reader.readAsArrayBuffer(file);
}

async function processAndSaveFretes() {
  if (!fteCteBuffer || !fteCteFileName) {
    fteToastError('Selecciona o Excel CT-e/NF primeiro.');
    return;
  }
  const spin = $('procSpin');
  const procBtn = $('procBtn');
  if (spin) spin.style.display = '';
  if (procBtn) procBtn.disabled = true;

  try {
    sapNfMap = {};
    const ok = processArrayBufferCte(fteCteBuffer, fteCteFileName);
    if (!ok) return;

    if (fteSapBuffer && fteSapFileName) {
      processArrayBufferSap(fteSapBuffer, fteSapFileName);
    }

    await persistFretesFile(fteCteSlot(), fteCteFileName, fteCteBuffer);
    if (fteSapBuffer && fteSapFileName) {
      await persistFretesFile(fteSapSlot(), fteSapFileName, fteSapBuffer);
    }

    _fteLoadedCompany = fteCompany();
    updateFretesFileStatus();
    fteToast('Processado e guardado.');
  } catch (err) {
    console.error('[fretes] processAndSave', err);
    fteToastError('Erro: ' + (err.message || err));
  } finally {
    if (spin) spin.style.display = 'none';
    checkFteBtn();
  }
}

function reEnrichAfterSapLoad() {
  if (lastCtePack) {
    processRows(lastCtePack.rows, lastCtePack.fileName, lastCtePack.sheetName, lastCtePack.headers);
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

const CTE_PCT_TARGET = 0.06;
const CTE_PCT_LOW = 0.059;
const CTE_PCT_HIGH = 0.061;

function fmtPp(v) { return (v * 100).toFixed(2) + ' p.p.'; }

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
    motivo: `Acima de 6%: pagou ${fmtMoney(pago)} (${fmtPct(pct)}) vs esperado ${fmtMoney(esperado)} — excedente +${fmtMoney(diff)} (+${fmtPp(diffPct)}). Provável tarifa mínima; questionar Unilog se inesperado.`
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

function buildNfRecord(g) {
  const ctesRaw = g.ctes || [];
  const nCte = resolveNfCteCount(ctesRaw, g.qtdCteFromSource);
  const ctes = ctesRaw.map((c, i) => analyzeCteEntry(c, i, g.valorNF, nCte));
  const pago = ctes.reduce((s, c) => s + c.pago, 0);
  const esperado = g.valorNF * CTE_PCT_TARGET;
  const diff = pago - esperado;
  const pct = g.valorNF > 0 ? pago / g.valorNF : 0;

  let status, motivo;
  if (nCte === 1) {
    ({ status, motivo } = analyzeSingleCteMotivo(pago, esperado, pct));
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
    temDevolucao: g.temDevolucao, dtNF: g.dtNF, qtdCteFromSource: g.qtdCteFromSource || 0
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
    qtdCteFromSource: nf.qtdCteFromSource
  });
  Object.assign(nf, rebuilt);
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

function processRows(rows, fileName, sheetName, headers) {
  lastCtePack = { rows, fileName, sheetName, headers };

  $('loadbar').textContent =
    'A processar ' + rows.length + ' linhas (folha "' + sheetName + '")...';

  const byNF = {};
  rows.forEach(r => {
    const nf = r.nf;
    if (nf === null || nf === undefined || String(nf).trim() === '') return;
    const nfKey = normNFKey(nf);
    if (!byNF[nfKey]) byNF[nfKey] = {
      nf: String(nf).trim(), valorNF: num(r.valorNF), transportador: r.transportador || '-',
      modalidade: r.modalidade || '-', ctes: [], temDevolucao: false, dtNF: r.dtNF,
      cliente: '', qtdCteFromSource: 0
    };
    const g = byNF[nfKey];
    const qtdRow = num(r.qtdCte);
    if (qtdRow > g.qtdCteFromSource) g.qtdCteFromSource = qtdRow;
    if (num(r.valorNF) > g.valorNF) g.valorNF = num(r.valorNF);
    if (r.transportador && g.transportador === '-') g.transportador = r.transportador;
    if (r.modalidade && g.modalidade === '-') g.modalidade = r.modalidade;
    if (r.dtNF && !g.dtNF) g.dtNF = r.dtNF;
    g.ctes.push({
      numCte: r.numCte, dtCte: r.dtCte, pago: num(r.pago),
      devolucao: (r.devolucao || '').toString().toLowerCase() === 'sim',
      tipoOp: r.tipoOp, peso: num(r.peso)
    });
    if ((r.devolucao || '').toString().toLowerCase() === 'sim') g.temDevolucao = true;
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
    $('loadbar').style.display = 'block';
    $('loadbar').textContent =
      '0 notas fiscais encontradas na folha "' + sheetName + '". Colunas: ' + cols;
    const us = $('uploadSection');
    if (us) us.style.display = 'block';
    $('results').style.display = 'none';
    fteToast('Não foi possível ler notas fiscais — verifica se o ficheiro tem a coluna "Nota Fiscal".');
    return;
  }

  renderAll();
  showResultsView();
  if (!_fteSkipAutosave) autosaveToCloud();
}

function computeSummary(list, fileName) {
  const totalPago = list.reduce((s, x) => s + x.pago, 0);
  const totalValorNF = list.reduce((s, x) => s + x.valorNF, 0);
  const totalEsperado = list.reduce((s, x) => s + x.esperado, 0);
  const excesso = totalPago - totalEsperado;
  const dates = list.flatMap(x => x.ctes.map(c => c.dtCte)).filter(Boolean).map(d => new Date(d));
  const periodoInicio = dates.length ? new Date(Math.min(...dates)) : null;
  const periodoFim = dates.length ? new Date(Math.max(...dates)) : null;

  const byStatus = { ok: 0, min: 0, dev: 0, flag: 0, low: 0 };
  const excessoByStatus = { ok: 0, min: 0, dev: 0, flag: 0, low: 0 };
  const pagoByStatus = { ok: 0, min: 0, dev: 0, flag: 0, low: 0 };
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
    qtd_ok: 0, qtd_min: 0, qtd_dev: 0, qtd_flag: 0, qtd_low: 0
  };
  rows.forEach(m => {
    t.totalNF += m.totalNF;
    t.totalValorNF += m.totalValorNF;
    t.totalPago += m.totalPago;
    t.totalEsperado += m.totalEsperado;
    t.excesso += m.excesso;
    t.qtd_ok += m.qtd_ok;
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
    'Tarifa mínima': r['Tarifa mínima'],
    'A investigar': r['A investigar'],
    'Conformes': r['Conformes']
  }));
}

function enrichNF(nf) {
  if (nf.mesRef) return nf;
  let d = nf.dtNF ? (parseSapBrDate(nf.dtNF) || (nf.dtNF instanceof Date && isPlausibleSapDate(nf.dtNF) ? nf.dtNF : null)) : null;
  if ((!d || isNaN(d)) && nf.ctes?.length) {
    const dates = nf.ctes.map(c => c.dtCte).filter(Boolean).map(x => parseSapBrDate(x) || new Date(x)).filter(x => x && !isNaN(x) && isPlausibleSapDate(x));
    if (dates.length) d = new Date(Math.min(...dates.map(x => x.getTime())));
  }
  nf.dtRef = d && !isNaN(d) ? d : null;
  nf.mesRef = d && !isNaN(d) ? monthKey(d) : 'sem-data';
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
      qtd_ok: 0, qtd_min: 0, qtd_dev: 0, qtd_flag: 0, qtd_low: 0
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

function renderSortableHead(containerId, columns, sortState, onSort) {
  const tr = $(containerId);
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

async function exportHistoryWorkbook() {
  if (!historyCache.length) {
    const { data, error } = await sb.from('cte_nf_uploads').select('*').order('criado_em', { ascending: false }).limit(50);
    if (error) { fteToast('Erro ao carregar histórico: ' + error.message); return; }
    historyCache = data || [];
  }
  if (!historyCache.length) { fteToast('Histórico vazio.'); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(historyCache.map(u => ({
    'Ficheiro': u.ficheiro_nome,
    'Data': new Date(u.criado_em),
    'Período início': u.periodo_inicio ? new Date(u.periodo_inicio) : '',
    'Período fim': u.periodo_fim ? new Date(u.periodo_fim) : '',
    'Total NFs': u.total_nf,
    'Total pago': u.total_pago,
    'Esperado': u.total_esperado,
    'Excesso': u.excesso,
    'Conformes': u.qtd_ok,
    'Tarifa mínima': u.qtd_min,
    'Com devolução': u.qtd_dev,
    'A investigar': u.qtd_flag,
    'Abaixo': u.qtd_low
  }))), 'Histórico');
  downloadWorkbook(wb, `Historico_analises_${new Date().toISOString().slice(0, 10)}.xlsx`);
  fteToast('Histórico exportado.');
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
    $('monthTableBody').innerHTML = '<tr><td colspan="9" class="empty">Carrega uma análise primeiro.</td></tr>';
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
      return `<tr class="month-year-row"><td colspan="9"><strong>${row.label}</strong></td></tr>`;
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
      <td class="right">${m.qtd_min}</td>
      <td class="right" style="color:${m.qtd_flag ? '#b3261e' : 'inherit'}">${m.qtd_flag}</td>
    </tr>`;
  }).join('');

  body.querySelectorAll('tr.month-row-clickable').forEach(tr => {
    tr.addEventListener('click', () => {
      selectedMonth = tr.dataset.mes;
      $('filterMes').value = selectedMonth;
      document.querySelectorAll('.fte-tab').forEach(x => x.classList.remove('active'));
      document.querySelector('.fte-tab[data-tab="analise"]').classList.add('active');
      $('tab-analise').style.display = 'block';
      $('tab-mensal').style.display = 'none';
      $('tab-historico').style.display = 'none';
      renderTable();
      $('nfTable').scrollIntoView({ behavior: 'smooth' });
      fteToast('Filtrado por ' + fmtMesLabel(selectedMonth));
    });
  });
}

function renderAll() {
  const s = currentSummary;
  const sapNote = s.sapLoaded ? ` · SAP: ${s.nSapMatched}/${s.totalNF} NFs cruzadas` : '';
  const mismatchNote = s.nValorMismatch ? ` · <span style="color:#b3261e;font-weight:600">${s.nValorMismatch} Δ valor SAP≠Unilog</span>` : '';
  const sapMissingNote = s.nSapMissing ? ` · <span style="color:#b3261e;font-weight:600">${s.nSapMissing} NF sem SAP</span>` : '';
  $('periodoLabel').innerHTML =
    `Ficheiro: ${s.fileName} · Período CT-e: ${fmtDate(s.periodoInicio)} a ${fmtDate(s.periodoFim)} · ${s.totalNF} notas fiscais · ${s.nUnicoCte} com CT-e único · ${s.nMultiplosCte} com múltiplos CT-e${sapNote}${mismatchNote}${sapMissingNote}`;

  const warnEl = $('dataWarning');
  if (s.loadWarning) {
    warnEl.style.display = 'block';
    warnEl.className = 'data-warn';
    warnEl.innerHTML = s.loadWarning;
  } else {
    warnEl.style.display = 'none';
    warnEl.innerHTML = '';
  }

  const sapKpi = s.sapLoaded ? `
    <div class="kpi${s.nValorMismatch ? ' flag' : ''}"><div class="label">Δ valor SAP ≠ Unilog</div><div class="value">${s.nValorMismatch}</div><div class="sub">${s.nSapMatched} NFs cruzadas com SAP</div></div>
    <div class="kpi${s.nSapMissing ? ' flag' : ''}"><div class="label">NF sem SAP</div><div class="value">${s.nSapMissing}</div><div class="sub">anomalias para investigar</div></div>` : '';

  $('kpis').innerHTML = `
    <div class="kpi"><div class="label">Faturação (valor NFs)</div><div class="value">${fmtMoney(s.totalValorNF)}</div><div class="sub">${s.totalNF} notas fiscais</div></div>
    <div class="kpi"><div class="label">Total pago (frete)</div><div class="value">${fmtMoney(s.totalPago)}</div><div class="sub">${fmtPct(s.pctPagoSobreNF)} sobre faturação (meta: 6%)</div></div>
    <div class="kpi"><div class="label">Esperado (6%)</div><div class="value">${fmtMoney(s.totalEsperado)}</div></div>
    <div class="kpi flag"><div class="label">Excesso líquido</div><div class="value">${fmtMoney(s.excesso)}</div><div class="sub">+${fmtPct(s.excessoPct)} vs esperado</div></div>
    <div class="kpi"><div class="label">NFs conformes</div><div class="value">${s.byStatus.ok}</div><div class="sub">${s.totalNF ? fmtPct(s.byStatus.ok / s.totalNF) : '0%'} do total</div></div>
    <div class="kpi flag"><div class="label">NFs a investigar ⚠️</div><div class="value">${s.byStatus.flag}</div><div class="sub">múltiplos CT-e sem devolução</div></div>
    ${sapKpi}
  `;

  $('reconcileBox').innerHTML = `
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
    { key: 'ok', cls: 'bd-ok', title: 'Conforme (~6%)', desc: '1 CT-e, pago entre 5,9% e 6,1% do valor da NF.' },
    { key: 'min', cls: 'bd-min', title: 'Tarifa mínima (provável)', desc: '1 CT-e acima de 6% — NF de baixo valor, frete mínimo do transportador.' },
    { key: 'dev', cls: 'bd-dev', title: 'Com devolução', desc: 'Múltiplos CT-e com retorno assinalado (ida + volta).' },
    { key: 'flag', cls: 'bd-flag', title: '⚠️ Sem devolução — investigar', desc: 'Múltiplos CT-e sem devolução — cada um pode cobrar % sobre a NF inteira.' },
    { key: 'low', cls: 'bd-low', title: 'Abaixo do esperado', desc: '1 CT-e abaixo de 5,9% — pagou menos que os 6%.' },
  ];
  $('breakdown').innerHTML = bd.filter(b => b.key !== 'low' || s.byStatus.low).map(b => `
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

  renderSubPanels();

  const transpSet = [...new Set(currentNFs.map(x => x.transportador))].sort();
  const tSel = $('filterTransp');
  tSel.innerHTML = '<option value="">Todos os transportadores</option>' + transpSet.map(t => `<option value="${t}">${t}</option>`).join('');

  monthlyRows = computeMonthly(currentNFs);
  const mSel = $('filterMes');
  mSel.innerHTML = '<option value="">Todos os meses</option>' + monthlyRows.map(m =>
    `<option value="${m.mesRef}"${selectedMonth === m.mesRef ? ' selected' : ''}>${m.mesLabel} (${m.totalNF} NF)</option>`
  ).join('');

  initMainTableHead();
  renderAnomaliesPanel();
  renderTable();
  renderMonthlyTable();
  if (quinzenalPack?.b2bRows?.length) {
    refreshQuinzenalCompare();
    if ($('tab-quinzenal')?.style.display === 'block') renderQuinzenalTab();
  }
}

function renderAnomaliesPanel() {
  const section = $('anomaliesSection');
  const hint = $('anomaliesHint');
  if (!section || !hint) return;

  const s = currentSummary;
  if (!s?.sapLoaded) {
    section.style.display = 'none';
    hint.style.display = currentNFs.length ? 'block' : 'none';
    return;
  }

  hint.style.display = 'none';
  section.style.display = 'block';

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
}

const CATEGORY_META = {
  ok: { title: 'Conforme (~6%)', cls: 'b-ok' },
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
  min: { cls: 'b-min', label: 'Tarifa mínima' },
  dev: { cls: 'b-dev', label: 'Com devolução' },
  flag: { cls: 'b-flag', label: '⚠️ Investigar' },
  low: { cls: 'b-low', label: 'Abaixo' }
};

function renderTable() {
  const list = getFilteredNFs();
  updateSortHeadIndicators('nfTableHead', tableSort);

  const body = $('nfTableBody');
  body.innerHTML = '';
  list.forEach((x) => {
    const tr = document.createElement('tr');
    tr.className = 'row-clickable'
      + (x.sapValorMismatch ? ' nf-val-mismatch' : '')
      + (x.sapMissing ? ' nf-sap-missing' : '');
    const meta = statusMeta[x.status];
    enrichNF(x);
    tr.innerHTML = `
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
    `;
    tr.addEventListener('click', () => toggleDetail(tr, x));
    body.appendChild(tr);
  });
  if (list.length === 0) {
    body.innerHTML = `<tr><td colspan="${nfColCount()}" class="empty">Sem resultados para este filtro.</td></tr>`;
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
    <strong>SAP:</strong> ${fmtDate(x.dtNF)} · ${x.cliente || '-'} · Unilog ${fmtMoney(x.valorUnilog ?? x.valorNF)} vs SAP ${fmtMoney(x.valorSAP)}${x.sapValorMismatch ? ` — <span style="color:#b3261e;font-weight:600">Δ relevante ${fmtMoney(x.valorDiff)}</span>` : (x.valorDiff != null ? ` (Δ ${fmtMoney(x.valorDiff)})` : '')}
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

async function loadHistory() {
  const el = $('historyList');
  el.innerHTML = '<p class="muted">A carregar...</p>';
  const { data, error } = await sb.from('cte_nf_uploads').select('*').order('criado_em', { ascending: false }).limit(50);
  if (error) { el.innerHTML = '<p class="muted">Erro a carregar histórico: ' + error.message + '</p>'; return; }
  historyCache = data || [];
  if (!historyCache.length) { el.innerHTML = '<div class="empty">Ainda não há análises guardadas.</div>'; return; }
  el.innerHTML = historyCache.map(u => `
    <div class="hist-item${u.id === currentUploadId ? ' active-hist' : ''}" data-id="${u.id}">
      <div>
        <div><strong>${u.ficheiro_nome || '(sem nome)'}</strong>${u.id === currentUploadId ? ' · <span style="color:var(--green);">activa</span>' : ''}</div>
        <div class="meta">${new Date(u.criado_em).toLocaleString('pt-BR')}
        ${u.periodo_inicio ? ` · período ${new Date(u.periodo_inicio).toLocaleDateString('pt-BR')} a ${new Date(u.periodo_fim).toLocaleDateString('pt-BR')}` : ''}
        · ${u.total_nf} NFs</div>
      </div>
      <div class="nums">
        <div><div class="v">${fmtMoney(u.total_pago)}</div><div class="l">Pago</div></div>
        <div><div class="v" style="color:#b3261e;">${fmtMoney(u.excesso)}</div><div class="l">Excesso</div></div>
        <div><div class="v">${u.qtd_flag || 0}</div><div class="l">A investigar</div></div>
      </div>
    </div>
  `).join('');

  el.querySelectorAll('.hist-item').forEach(item => {
    item.addEventListener('click', async () => {
      const ok = await loadUploadById(item.dataset.id);
      if (ok) {
        document.querySelectorAll('.fte-tab').forEach(x => x.classList.remove('active'));
        document.querySelector('.fte-tab[data-tab="analise"]').classList.add('active');
        $('tab-analise').style.display = 'block';
        $('tab-historico').style.display = 'none';
        fteToast('Análise do histórico carregada.');
      }
    });
  });
}

async function restoreCteFromRec(cteRec, sapRec, silent = false) {
  if (!cteRec?.file_data) return false;
  const co = fteCompany();
  if (_fteLoadedCompany === co && currentNFs.length && lastCtePack?.fileName === cteRec.file_name) {
    return true;
  }
  _fteSkipAutosave = true;
  try {
    sapNfMap = {};
    const cteBuf = base64ToArrayBuffer(cteRec.file_data);
    const ok = processArrayBufferCte(cteBuf, cteRec.file_name || 'cte.xlsx', { silent: true });
    if (!ok) return false;
    if (sapRec?.file_data) {
      processArrayBufferSap(base64ToArrayBuffer(sapRec.file_data), sapRec.file_name || 'sap.xlsx', { silent: true });
    }
    _fteLoadedCompany = co;
    fteCteFileName = cteRec.file_name || '';
    fteCteBuffer = cteBuf;
    setCteZoneLoaded(cteRec.file_name);
    if (sapRec?.file_data) {
      fteSapFileName = sapRec.file_name || '';
      fteSapBuffer = base64ToArrayBuffer(sapRec.file_data);
      setSapZoneLoaded(sapRec.file_name);
    } else {
      fteSapFileName = '';
      fteSapBuffer = null;
      setSapZoneLoaded('');
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
    const meta = await fetchExcelFiles([fteCteSlot(), fteSapSlot()]);
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
  _fteLoadSavedPromise = _loadSavedFretesFilesImpl(silent).finally(() => { _fteLoadSavedPromise = null; });
  return _fteLoadSavedPromise;
}

async function _loadSavedFretesFilesImpl(silent = false) {
  if (typeof fetchExcelFiles !== 'function' || typeof base64ToArrayBuffer !== 'function') {
    if (!silent) fteToastError('Persistência Excel indisponível — recarrega a página.');
    return false;
  }
  const co = fteCompany();
  let meta;
  try {
    meta = await fetchExcelFiles([fteCteSlot(), fteSapSlot(), fteQuinzenalSlot()]);
  } catch (err) {
    console.error('[fretes] load saved', err);
    if (!silent) {
      if (typeof isExcelFilesTableMissing === 'function' && isExcelFilesTableMissing(err)) {
        fteToastError('Tabela logistica_excel_files em falta no Supabase.');
      } else {
        fteToastError('Erro ao carregar ficheiros: ' + (err.message || err));
      }
    }
    return false;
  }

  const cteRec = meta[fteCteSlot()];
  const sapRec = meta[fteSapSlot()];

  let cteOk = false;
  if (cteRec?.file_data) {
    cteOk = await restoreCteFromRec(cteRec, sapRec, silent);
  } else if (_fteLoadedCompany === co && currentNFs.length) {
    cteOk = true;
  } else if (cteRec?.file_name && !silent) {
    fteToastError('CT-e guardado só com nome — clica Processar e Guardar para persistir o ficheiro.');
  }

  const qzLoaded = await loadSavedQuinzenalPack(silent, meta, { deferCompare: true, skipRender: true });
  refreshQuinzenalCompare();
  updateFretesFileStatus(meta);

  if (!cteOk && !silent && !qzLoaded && !cteRec) {
    fteToastError('Sem ficheiros CT-e guardados para ' + co + '.');
  } else if (cteOk && !silent) {
    fteToast('Últimos ficheiros fretes carregados.');
  }

  if ($('tab-quinzenal')?.style.display === 'block') renderQuinzenalTab();
  return cteOk || qzLoaded || currentNFs.length > 0;
}

// ── QUINZENAL UNILOG (B2B diff vs CT-e · B2C vendas/fretes) ──

const QZ_MESES = {
  janeiro: 1, fevereiro: 2, marco: 3, março: 3, abril: 4, maio: 5, junho: 6,
  julho: 7, agosto: 8, setembro: 9, outubro: 10, novembro: 11, dezembro: 12
};

const QZ_B2B_ALIASES = {
  nf: ['nota fiscal', 'nf', 'n nota fiscal'],
  valorNF: ['valor nf', 'valor da nf', 'vl nf'],
  numCte: ['num. cte', 'num cte', 'numero cte', 'numeros cte', 'n cte'],
  pago: ['total fatura rev.', 'total fatura rev', 'total fatura', 'valor fatura'],
  transportador: ['transportador', 'transportadora'],
  dtNF: ['dt nf', 'data nf', 'data nota fiscal'],
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
  return {
    nf: findQzField(row, 'nf', QZ_B2B_ALIASES),
    valorNF: findQzField(row, 'valorNF', QZ_B2B_ALIASES),
    numCte: findQzField(row, 'numCte', QZ_B2B_ALIASES),
    pago: findQzField(row, 'pago', QZ_B2B_ALIASES),
    transportador: findQzField(row, 'transportador', QZ_B2B_ALIASES),
    dtNF: findQzField(row, 'dtNF', QZ_B2B_ALIASES),
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
        mesKey: fileMeta.mesKey || mesKeyFromQuinzenaKey(fileMeta.quinzenaKey),
        mesLabel: fileMeta.mesLabel || (fileMeta.mesKey ? fmtMesLabel(fileMeta.mesKey) : ''),
        quinzenaKey: fileMeta.quinzenaKey, quinzenaLabel: fileMeta.quinzenaLabel, fileName: fileMeta.fileName
      };
    }
    const g = byNf[key];
    if (num(r.valorNF) > g.valorNF) g.valorNF = num(r.valorNF);
    g.pago += num(r.pago);
    if (r.numCte) g.nCteSet.add(String(r.numCte).trim());
    if (r.transportador && g.transportador === '') g.transportador = r.transportador;
  });
  return Object.values(byNf).map(g => ({ ...g, nCte: g.nCteSet.size || 1, nCteSet: undefined }));
}

function cteNfsInQuinzena(range) {
  if (!range || !currentNFs.length) return [];
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

function mergeQzB2BNf(into, r) {
  into.valorNF = Math.max(into.valorNF, r.valorNF);
  into.pago += r.pago;
  into.nCte += r.nCte;
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
  const byMes = {};
  b2bRows.forEach(r => {
    const k = r.mesKey || mesKeyFromQuinzenaKey(r.quinzenaKey);
    if (!byMes[k]) {
      byMes[k] = {
        mesKey: k, mesLabel: r.mesLabel || (k !== 'sem-mes' ? fmtMesLabel(k) : 'Sem mês'),
        byNf: {}
      };
    }
    const nfKey = r.nfKey || normNFKey(r.nf);
    if (!nfKey) return;
    if (!byMes[k].byNf[nfKey]) byMes[k].byNf[nfKey] = { valorNF: 0, pago: 0, nCte: 0 };
    mergeQzB2BNf(byMes[k].byNf[nfKey], r);
  });
  return Object.values(byMes).map(b => {
    const nfs = Object.values(b.byNf);
    const out = {
      mesKey: b.mesKey, mesLabel: b.mesLabel,
      nfCountQz: nfs.length,
      totalValorNFQz: nfs.reduce((s, x) => s + x.valorNF, 0),
      totalPagoQz: nfs.reduce((s, x) => s + x.pago, 0),
      totalCteQz: nfs.reduce((s, x) => s + x.nCte, 0),
      nfCountCte: 0, totalValorNFCte: 0, totalPagoCte: 0, totalCteCte: 0
    };
    if (b.mesKey !== 'sem-mes') {
      const parts = b.mesKey.split('-');
      const ano = parseInt(parts[0], 10);
      const mes = parseInt(parts[1], 10);
      const cteList = cteNfsInMonth(ano, mes);
      out.nfCountCte = cteList.length;
      out.totalValorNFCte = cteList.reduce((s, x) => s + x.valorNF, 0);
      out.totalPagoCte = cteList.reduce((s, x) => s + x.pago, 0);
      out.totalCteCte = cteList.reduce((s, x) => s + x.nCte, 0);
    }
    return out;
  }).sort((a, b) => compareMesRef(a.mesKey, b.mesKey));
}

function buildB2CMonthTotals(b2cRows) {
  const byMes = {};
  b2cRows.forEach(r => {
    const k = resolveB2cRowMesKey(r);
    if (!byMes[k]) byMes[k] = { mesKey: k, mesLabel: resolveB2cRowMesLabel(r, k), count: 0, totalVendas: 0, totalFrete: 0 };
    byMes[k].count++; byMes[k].totalVendas += r.valorNF; byMes[k].totalFrete += r.pago;
  });
  return Object.values(byMes).map(b => ({ ...b, pctFrete: b.totalVendas ? b.totalFrete / b.totalVendas : 0 }))
    .sort((a, b) => compareMesRef(a.mesKey, b.mesKey));
}

function buildB2BCompare(b2bRows) {
  const qzMap = {};
  b2bRows.forEach(r => {
    const k = r.nfKey;
    if (!qzMap[k]) qzMap[k] = { ...r };
    else mergeQzB2BNf(qzMap[k], r);
  });
  const cteMap = {};
  currentNFs.forEach(nf => { const k = normNFKey(nf.nf); if (k) cteMap[k] = nf; });
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
    let status = 'match';
    if (qz && !cte) status = 'onlyQz';
    else if (!qz && cte) status = 'onlyCte';
    else if (Math.abs(valorNFQz - valorNFCte) > 1 || Math.abs(pagoQz - pagoCte) > 0.5 || nCteQz !== nCteCte) status = 'diff';
    rows.push({
      nfKey: k, nf: qz?.nf || cte?.nf || k,
      mesKey: qz?.mesKey || mesKeyFromQuinzenaKey(qz?.quinzenaKey),
      mesLabel: qz?.mesLabel || (qz?.mesKey ? fmtMesLabel(qz.mesKey) : ''),
      quinzenaKey: qz?.quinzenaKey || '', quinzenaLabel: qz?.quinzenaLabel || '',
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
  }
}

function updateQzProcessStatus() {
  const st = $('qzStatus');
  if (!st) return;
  st.textContent = quinzenalPack?.files?.length ? quinzenalStatusFromPack(quinzenalPack) : '';
}

function syncQzUploadZone() {
  const fn = $('qzFn');
  const zone = $('qzZone');
  const list = $('qzFileList');
  const btn = $('qzProcBtn');
  const pending = fteQzPendingFiles.length;
  const saved = quinzenalPack?.files?.length || 0;
  if (pending) {
    if (fn) fn.textContent = `✓ ${pending} ficheiro(s) seleccionado(s)`;
    zone?.classList.add('loaded');
    if (list) {
      list.style.display = 'block';
      list.innerHTML = fteQzPendingFiles.map(f => `<div>${f.name}</div>`).join('');
    }
    if (btn) btn.disabled = false;
  } else if (saved) {
    if (fn) fn.textContent = `✓ ${saved} ficheiro(s) guardado(s)`;
    zone?.classList.add('loaded');
    if (list) {
      list.style.display = 'block';
      list.innerHTML = quinzenalPack.files.map(f => `<div>${f.fileName}</div>`).join('');
    }
    if (btn) btn.disabled = true;
  } else {
    if (fn) fn.textContent = '';
    zone?.classList.remove('loaded');
    if (list) { list.style.display = 'none'; list.innerHTML = ''; }
    if (btn) btn.disabled = true;
  }
}

function setQzZoneLoaded(count) {
  syncQzUploadZone();
}

async function persistQuinzenalPack(pack) {
  if (typeof upsertExcelBinary !== 'function') return;
  const buf = new TextEncoder().encode(JSON.stringify(pack)).buffer;
  try {
    await upsertExcelBinary(fteQuinzenalSlot(), `quinzenal_${pack.files.length}_files.json`, buf);
    updateQzProcessStatus();
    updateFretesFileStatus();
  } catch (err) { console.error('[fretes] persist quinzenal', err); }
}

async function loadSavedQuinzenalPack(silent, meta, opts = {}) {
  if (typeof fetchExcelFiles !== 'function') return false;
  try {
    const m = meta || await fetchExcelFiles([fteQuinzenalSlot()]);
    const rec = m[fteQuinzenalSlot()];
    if (!rec?.file_data) return false;
    quinzenalPack = parseQuinzenalPackFromRec(rec);
    if (!quinzenalPack?.files?.length) return false;
    if (!opts?.deferCompare) refreshQuinzenalCompare();
    syncQzUploadZone();
    updateQzProcessStatus();
    if (!silent) fteToast('Relatórios quinzenais carregados.');
    if (!opts?.skipRender) renderQuinzenalTab();
    return true;
  } catch (err) { console.error('[fretes] load quinzenal', err); return false; }
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

async function processAndSaveQuinzenal() {
  if (!fteQzPendingFiles.length) { fteToastError('Selecciona ficheiros quinzenais primeiro.'); return; }
  const spin = $('qzSpin');
  const btn = $('qzProcBtn');
  if (spin) spin.style.display = '';
  if (btn) btn.disabled = true;
  try {
    const results = await Promise.all(fteQzPendingFiles.map(f => processQuinzenalFile(f)));
    const ok = results.filter(r => r.ok);
    const fail = results.filter(r => !r.ok);
    if (!ok.length) { fteToastError('Nenhum ficheiro processado.'); return; }
    const fileBinaries = {};
    ok.forEach(r => {
      if (!r.arrayBuffer || typeof arrayBufferToBase64 !== 'function') return;
      fileBinaries[r.meta.fileName] = arrayBufferToBase64(r.arrayBuffer);
    });
    quinzenalPack = buildQuinzenalPack(results, fileBinaries);
    await persistQuinzenalPack(quinzenalPack);
    fteQzPendingFiles = [];
    const qzInput = $('qzFileInput');
    if (qzInput) qzInput.value = '';
    syncQzUploadZone();
    const fc = qzFileCounts();
    const st = $('qzStatus');
    if (st) st.textContent = quinzenalStatusFromPack(quinzenalPack) ||
      `${fc.total} ficheiro(s) · ${quinzenalPack.b2bRows.length} NFs B2B · ${quinzenalPack.b2cRows.length} linhas B2C`;
    if (fail.length) fteToast(`${ok.length} OK, ${fail.length} com erro`);
    else fteToast('Quinzenais processados e guardados.');
    renderQuinzenalTab();
  } catch (err) {
    console.error('[fretes] quinzenal', err);
    fteToastError('Erro: ' + (err.message || err));
  } finally {
    if (spin) spin.style.display = 'none';
    if (btn) btn.disabled = !fteQzPendingFiles.length;
  }
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
    onlyQz: '<span style="color:#b3261e">Só quinzenal</span>', onlyCte: '<span style="color:#b3261e">Só CT-e</span>' }[s] || s;
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

async function renderQuinzenalTab() {
  const empty = $('qzEmpty');
  const b2bPanel = $('qzB2b');
  const b2cPanel = $('qzB2c');
  const warn = $('qzWarn');
  if (quinzenalPack?.b2bRows?.length && !currentNFs.length) {
    await ensureCteLoadedForQz(true);
  }
  if (warn) {
    if (quinzenalPack?.b2bRows?.length && !currentNFs.length) {
      warn.style.display = 'block';
      warn.innerHTML = '<strong>Export Unilog indisponível:</strong> processa o Excel CT-e/NF na secção acima (ou Processar e Guardar) para comparar B2B quinzenal vs export original.';
    } else warn.style.display = 'none';
  }
  if (!quinzenalPack?.files?.length) {
    if (empty) empty.style.display = 'block';
    if (b2bPanel) b2bPanel.style.display = 'none';
    if (b2cPanel) b2cPanel.style.display = 'none';
    return;
  }
  if (empty) empty.style.display = 'none';
  updateQzFileNote();
  populateQzMonthFilters();
  refreshQuinzenalCompare();
  if (activeQzPanel === 'b2b') {
    if (b2bPanel) b2bPanel.style.display = 'block';
    if (b2cPanel) b2cPanel.style.display = 'none';
    renderQzB2b();
  } else {
    if (b2bPanel) b2bPanel.style.display = 'none';
    if (b2cPanel) b2cPanel.style.display = 'block';
    renderQzB2c();
  }
}

function renderQzB2b() {
  const rows = quinzenalPack.b2bCompare || [];
  const sum = qzCompareSummary(rows);
  const kpis = $('qzB2bKpis');
  if (kpis) kpis.innerHTML = `
    <div class="kpi"><div class="label">NFs comparadas</div><div class="value">${sum.total}</div></div>
    <div class="kpi flag"><div class="label">Só quinzenal</div><div class="value">${sum.onlyQz}</div></div>
    <div class="kpi flag"><div class="label">Só CT-e</div><div class="value">${sum.onlyCte}</div></div>
    <div class="kpi"><div class="label">Com diferença</div><div class="value">${sum.diff}</div></div>
    <div class="kpi"><div class="label">Δ valor NF</div><div class="value">${sum.diffValor}</div></div>
    <div class="kpi"><div class="label">Conformes</div><div class="value">${sum.match}</div></div>`;

  const qzBody = $('qzB2bQuinzenaBody');
  if (qzBody) {
    const totals = quinzenalPack.b2bMonthTotals || [];
    const qzByMes = qzByMesKey(quinzenalPack.b2bQuinzenaTotals || buildB2BQuinzenaTotals(quinzenalPack.b2bRows || []));
    let tQz = { nf: 0, val: 0, frete: 0, cte: 0, nfC: 0, valC: 0, freteC: 0, cteC: 0 };
    qzBody.innerHTML = totals.map(t => {
      tQz.nf += t.nfCountQz; tQz.val += t.totalValorNFQz; tQz.frete += t.totalPagoQz; tQz.cte += t.totalCteQz;
      tQz.nfC += t.nfCountCte; tQz.valC += t.totalValorNFCte; tQz.freteC += t.totalPagoCte; tQz.cteC += t.totalCteCte;
      const qzRows = qzByMes[t.mesKey] || [];
      const hasQz = qzRows.length > 0;
      const expanded = qzExpandedMonths.has(qzExpandKey('b2b', t.mesKey));
      const chev = hasQz ? `<span class="qz-chevron">${expanded ? '▾' : '▸'}</span> ` : '';
      let html = `<tr class="qz-month-row${hasQz ? ' qz-expandable' : ''}${expanded ? ' expanded' : ''}" data-panel="b2b" data-mes="${t.mesKey}">
        <td>${chev}<strong>${t.mesLabel}</strong></td>
        <td class="right">${t.nfCountQz}</td><td class="right">${t.nfCountCte}</td><td class="right">${t.nfCountCte - t.nfCountQz}</td>
        <td class="right">${fmtMoney(t.totalValorNFQz)}</td><td class="right">${fmtMoney(t.totalValorNFCte)}</td>
        <td class="right">${fmtQzDiff(t.totalValorNFCte - t.totalValorNFQz, 1)}</td>
        <td class="right">${fmtMoney(t.totalPagoQz)}</td><td class="right">${fmtMoney(t.totalPagoCte)}</td>
        <td class="right">${fmtQzDiff(t.totalPagoCte - t.totalPagoQz, 0.5)}</td>
        <td class="right">${t.totalCteQz}</td><td class="right">${t.totalCteCte}</td><td class="right">${t.totalCteCte - t.totalCteQz}</td></tr>`;
      if (expanded && hasQz) html += qzRows.map(renderB2bQuinzenaRow).join('');
      return html;
    }).join('') + `<tr class="month-total-row"><td><strong>Total geral</strong></td>
      <td class="right">${tQz.nf}</td><td class="right">${tQz.nfC}</td><td class="right">${tQz.nfC - tQz.nf}</td>
      <td class="right">${fmtMoney(tQz.val)}</td><td class="right">${fmtMoney(tQz.valC)}</td>
      <td class="right">${fmtQzDiff(tQz.valC - tQz.val, 1)}</td>
      <td class="right">${fmtMoney(tQz.frete)}</td><td class="right">${fmtMoney(tQz.freteC)}</td>
      <td class="right">${fmtQzDiff(tQz.freteC - tQz.frete, 0.5)}</td>
      <td class="right">${tQz.cte}</td><td class="right">${tQz.cteC}</td><td class="right">${tQz.cteC - tQz.cte}</td></tr>`;
    bindQzExpandClicks(qzBody, 'b2b');
  }

  const qFilter = $('qzFilterQuinzena')?.value || '';
  const dFilter = $('qzFilterDiff')?.value || '';
  const search = ($('qzSearchNF')?.value || '').trim();
  let list = rows;
  if (qFilter) list = list.filter(r => (r.mesKey || mesKeyFromQuinzenaKey(r.quinzenaKey)) === qFilter);
  if (dFilter === 'onlyQz') list = list.filter(r => r.status === 'onlyQz');
  else if (dFilter === 'onlyCte') list = list.filter(r => r.status === 'onlyCte');
  else if (dFilter === 'valor') list = list.filter(r => Math.abs(r.diffValorNF) > 1);
  else if (dFilter === 'frete') list = list.filter(r => Math.abs(r.diffPago) > 0.5);
  else if (dFilter === 'cte') list = list.filter(r => r.diffCte !== 0);
  if (search) list = list.filter(r => String(r.nf).includes(search));

  const detBody = $('qzB2bDetailBody');
  if (detBody) {
    detBody.innerHTML = list.slice(0, 2000).map(r => {
      const cls = r.status === 'onlyQz' || r.status === 'onlyCte' ? 'qz-only-row' : (r.status === 'diff' ? 'qz-diff-row' : '');
      const pctNote = r.valorNFQz > 0 && r.pagoQz > 0 ? fmtPct(r.pagoQz / r.valorNFQz) : '-';
      return `<tr class="${cls}"><td>${r.nf}</td><td>${r.quinzenaLabel || '-'}</td>
        <td class="right">${r.valorNFQz ? fmtMoney(r.valorNFQz) : '-'}</td><td class="right">${r.valorNFCte ? fmtMoney(r.valorNFCte) : '-'}</td>
        <td class="right">${fmtQzDiff(r.diffValorNF, 1)}</td>
        <td class="right">${r.pagoQz ? fmtMoney(r.pagoQz) : '-'}</td><td class="right">${r.pagoCte ? fmtMoney(r.pagoCte) : '-'}</td>
        <td class="right">${fmtQzDiff(r.diffPago, 0.5)}</td>
        <td class="right">${r.nCteQz || '-'}</td><td class="right">${r.nCteCte || '-'}</td><td class="right">${r.diffCte || 0}</td>
        <td>${fmtQzStatus(r.status)}</td><td class="right">${pctNote}</td></tr>`;
    }).join('') || '<tr><td colspan="13" class="empty">Sem dados B2B — carrega quinzenais B2B e análise CT-e</td></tr>';
  }
}

function renderQzB2c() {
  const rows = quinzenalPack.b2cRows || [];
  const totalVendas = rows.reduce((s, r) => s + r.valorNF, 0);
  const totalFrete = rows.reduce((s, r) => s + r.pago, 0);
  const kpis = $('qzB2cKpis');
  if (kpis) kpis.innerHTML = `
    <div class="kpi"><div class="label">Entregas B2C</div><div class="value">${rows.length}</div></div>
    <div class="kpi"><div class="label">Vendas (Total NF)</div><div class="value">${fmtMoney(totalVendas)}</div></div>
    <div class="kpi"><div class="label">Frete pago</div><div class="value">${fmtMoney(totalFrete)}</div></div>
    <div class="kpi"><div class="label">% frete/vendas</div><div class="value">${fmtPct(totalVendas ? totalFrete / totalVendas : 0)}</div></div>`;

  const qBody = $('qzB2cQuinzenaBody');
  if (qBody) {
    const totals = quinzenalPack.b2cMonthTotals || [];
    const qzByMes = qzByMesKey(quinzenalPack.b2cQuinzenaTotals || buildB2CQuinzenaTotals(quinzenalPack.b2cRows || []));
    let t = { c: 0, v: 0, f: 0 };
    qBody.innerHTML = totals.map(b => {
      t.c += b.count; t.v += b.totalVendas; t.f += b.totalFrete;
      const qzRows = qzByMes[b.mesKey] || [];
      const hasQz = qzRows.length > 0;
      const expanded = qzExpandedMonths.has(qzExpandKey('b2c', b.mesKey));
      const chev = hasQz ? `<span class="qz-chevron">${expanded ? '▾' : '▸'}</span> ` : '';
      let html = `<tr class="qz-month-row${hasQz ? ' qz-expandable' : ''}${expanded ? ' expanded' : ''}" data-panel="b2c" data-mes="${b.mesKey}">
        <td>${chev}<strong>${b.mesLabel}</strong></td><td class="right">${b.count}</td>
        <td class="right">${fmtMoney(b.totalVendas)}</td><td class="right">${fmtMoney(b.totalFrete)}</td>
        <td class="right">${fmtPct(b.pctFrete)}</td></tr>`;
      if (expanded && hasQz) html += qzRows.map(renderB2cQuinzenaRow).join('');
      return html;
    }).join('') + `<tr class="month-total-row"><td><strong>Total geral</strong></td>
      <td class="right">${t.c}</td><td class="right">${fmtMoney(t.v)}</td><td class="right">${fmtMoney(t.f)}</td>
      <td class="right">${fmtPct(t.v ? t.f / t.v : 0)}</td></tr>`;
    bindQzExpandClicks(qBody, 'b2c');
  }

  const qFilter = $('qzB2cFilterQuinzena')?.value || '';
  const search = ($('qzB2cSearch')?.value || '').toLowerCase();
  let list = rows;
  if (qFilter) list = list.filter(r => resolveB2cRowMesKey(r) === qFilter);
  if (search) list = list.filter(r => String(r.pedido || '').includes(search) || String(r.nf || '').includes(search) ||
    String(r.destinatario || '').toLowerCase().includes(search));

  const detBody = $('qzB2cDetailBody');
  if (detBody) {
    detBody.innerHTML = list.slice(0, 2000).map(r => `<tr>
      <td>${r.pedido || '-'}</td><td>${r.nf || '-'}</td><td>${r.quinzenaLabel || '-'}</td>
      <td>${fmtDate(r.dtColeta || r.dtNF)}</td><td>${r.destinatario || '-'}</td><td>${r.transportador || '-'}</td>
      <td class="right">${fmtMoney(r.valorNF)}</td><td class="right">${fmtMoney(r.pago)}</td>
      <td class="right">${fmtPct(r.valorNF ? r.pago / r.valorNF : 0)}</td><td>${r.zona || '-'}</td></tr>`).join('')
      || '<tr><td colspan="10" class="empty">Sem dados B2C</td></tr>';
  }
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
    NF: r.nf, Quinzena: r.quinzenaLabel, 'Valor NF QZ': r.valorNFQz, 'Valor NF CT-e': r.valorNFCte, 'Δ valor NF': r.diffValorNF,
    'Frete QZ': r.pagoQz, 'Frete CT-e': r.pagoCte, 'Δ frete': r.diffPago,
    'Qtd CT-e QZ': r.nCteQz, 'Qtd CT-e CT-e': r.nCteCte, 'Δ CT-e': r.diffCte, Estado: r.status,
    '% pago QZ (6%)': r.valorNFQz > 0 ? r.pagoQz / r.valorNFQz : null
  }));
  if (b2b.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2b), 'Diff B2B');
  const b2cMes = (quinzenalPack.b2cMonthTotals || []).map(t => ({
    Mês: t.mesLabel, Entregas: t.count, Vendas: t.totalVendas, 'Frete pago': t.totalFrete, '% frete/vendas': t.pctFrete
  }));
  if (b2cMes.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2cMes), 'Resumo B2C mensal');
  const b2c = (quinzenalPack.b2cRows || []).map(r => ({
    Pedido: r.pedido, NF: r.nf, Quinzena: r.quinzenaLabel, 'Data coleta': r.dtColeta || r.dtNF,
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
  lastCtePack = null;
  activeSubPanel = null;
  selectedMonth = '';
  fteCteFileName = '';
  fteCteBuffer = null;
  fteSapFileName = '';
  fteSapBuffer = null;
  fteQzPendingFiles = [];
  quinzenalPack = null;
  qzExpandedMonths = new Set();
  syncQzUploadZone();
  const results = $('results');
  if (results) results.style.display = 'none';
  const us = $('uploadSection');
  if (us) us.style.display = 'block';
  setCteZoneLoaded('');
  setSapZoneLoaded('');
  setLoadbar('');
  updateSaveStatus('');
  updateFretesFileStatus();
  if (typeof loadSavedFretesFiles === 'function') loadSavedFretesFiles(true);
}

function initFretes() {
  if (fteInited) return;
  fteInited = true;

  document.querySelectorAll('.fte-tab').forEach(t => {
    t.addEventListener('click', () => {
      document.querySelectorAll('.fte-tab').forEach(x => x.classList.remove('active'));
      t.classList.add('active');
      $('tab-analise').style.display = t.dataset.tab === 'analise' ? 'block' : 'none';
      $('tab-mensal').style.display = t.dataset.tab === 'mensal' ? 'block' : 'none';
      $('tab-historico').style.display = t.dataset.tab === 'historico' ? 'block' : 'none';
      $('tab-quinzenal').style.display = t.dataset.tab === 'quinzenal' ? 'block' : 'none';
      if (t.dataset.tab === 'historico') loadHistory();
      if (t.dataset.tab === 'mensal') renderMonthlyTable();
      if (t.dataset.tab === 'quinzenal') renderQuinzenalTab();
    });
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
  $('loadLastBtn')?.addEventListener('click', () => loadSavedFretesFiles(false));
  $('qzProcBtn')?.addEventListener('click', () => processAndSaveQuinzenal());

  document.querySelectorAll('.qz-subtab').forEach(st => {
    st.addEventListener('click', () => {
      document.querySelectorAll('.qz-subtab').forEach(x => x.classList.remove('active'));
      st.classList.add('active');
      activeQzPanel = st.dataset.qz || 'b2b';
      renderQuinzenalTab();
    });
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

  ['qzFilterQuinzena', 'qzFilterDiff', 'qzSearchNF', 'qzB2cFilterQuinzena', 'qzB2cSearch'].forEach(id => {
    const el = $(id);
    if (!el) return;
    const fn = () => renderQuinzenalTab();
    el.addEventListener('input', fn);
    el.addEventListener('change', fn);
  });

  $('newFileBtn').addEventListener('click', () => {
    $('results').style.display = 'none';
    const us = $('uploadSection');
    if (us) us.style.display = 'block';
  });
  $('exportBtn').addEventListener('click', () => exportFullWorkbook());
  $('exportFilteredBtn').addEventListener('click', () => exportFilteredView());
  $('exportHistoryBtn').addEventListener('click', () => exportHistoryWorkbook());
  $('exportMensalBtn').addEventListener('click', () => exportMonthlyWorkbook());
  $('exportQzBtn')?.addEventListener('click', () => exportQuinzenalWorkbook());

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

  applyFretesFileLabelsFromMeta();
  loadSavedFretesFiles(true);
}
