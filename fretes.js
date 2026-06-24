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
let monthSort = { col: 'excesso', dir: -1 };
let selectedMonth = '';
let monthlyRows = [];
let historyCache = [];

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

const MONTH_COLUMNS = [
  { key: 'mesLabel', label: 'Mês', type: 'string' },
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
const SAP_COL_IDX = { cliente: 1, dtEmissao: 2, nf: 3 };

function fteToast(msg) {
  if (typeof toast === 'function') toast(msg, 'success');
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
  else $('uploadZone').style.display = 'none';
  $('results').style.display = 'block';
}

function fmtMoney(v) { return 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function fmtPct(v) { return (v * 100).toFixed(1) + '%'; }

function normCol(s) {
  return String(s || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function normNFKey(nf) {
  const s = String(nf || '').trim();
  const digits = s.replace(/\D/g, '');
  return digits || s;
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

function sapHeaderMatchesField(hdr, field) {
  const h = normCol(hdr);
  if (!h) return false;
  return SAP_ALIASES[field].some(alias => h === alias || h.includes(alias) || alias.includes(h));
}

function parseSapBrDate(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    const d = new Date(v.getFullYear(), v.getMonth(), v.getDate());
    d.setHours(0, 0, 0, 0);
    return d;
  }
  if (typeof v === 'number' && !isNaN(v)) {
    if (typeof XLSX !== 'undefined' && XLSX.SSF && v >= 1 && v < 1000000) {
      const dc = XLSX.SSF.parse_date_code(v);
      if (dc && dc.y >= 1990 && dc.y <= 2100) {
        const d = new Date(dc.y, dc.m - 1, dc.d);
        d.setHours(0, 0, 0, 0);
        return d;
      }
    }
    return null;
  }
  const s = String(v).trim();
  const dotParts = s.split('.');
  if (dotParts.length === 3) {
    const dd = parseInt(dotParts[0], 10), mo = parseInt(dotParts[1], 10), yr = parseInt(dotParts[2], 10);
    if (yr >= 1990 && yr <= 2100 && mo >= 1 && mo <= 12 && dd >= 1 && dd <= 31) {
      const d = new Date(yr, mo - 1, dd);
      if (!isNaN(d.getTime()) && d.getFullYear() === yr && d.getMonth() === mo - 1 && d.getDate() === dd) {
        d.setHours(0, 0, 0, 0);
        return d;
      }
    }
  }
  const slashParts = s.split('/');
  if (slashParts.length >= 3) {
    const yr = parseInt(String(slashParts[2]).trim(), 10);
    const mo = parseInt(slashParts[1], 10);
    const dd = parseInt(slashParts[0], 10);
    if (yr >= 1990 && yr <= 2100 && mo >= 1 && mo <= 12 && dd >= 1 && dd <= 31) {
      const d = new Date(yr, mo - 1, dd);
      if (!isNaN(d.getTime())) { d.setHours(0, 0, 0, 0); return d; }
    }
  }
  const m = s.match(/^(\d{2})[.\-/](\d{2})[.\-/](\d{4})$/) || s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const yr = m[1].length === 4 ? parseInt(m[1], 10) : parseInt(m[3], 10);
    const mo = parseInt(m[2], 10);
    const dd = m[1].length === 4 ? parseInt(m[3], 10) : parseInt(m[1], 10);
    if (yr >= 1990 && yr <= 2100 && mo >= 1 && mo <= 12 && dd >= 1 && dd <= 31) {
      const d = new Date(yr, mo - 1, dd);
      if (!isNaN(d.getTime())) { d.setHours(0, 0, 0, 0); return d; }
    }
  }
  if (typeof XLSX !== 'undefined' && XLSX.SSF) {
    const serial = parseFloat(s.replace(',', '.'));
    if (!isNaN(serial) && serial >= 1 && serial < 1000000) {
      const dc = XLSX.SSF.parse_date_code(serial);
      if (dc && dc.y >= 1990 && dc.y <= 2100) {
        const d = new Date(dc.y, dc.m - 1, dc.d);
        d.setHours(0, 0, 0, 0);
        return d;
      }
    }
  }
  const d2 = new Date(s);
  if (!isNaN(d2.getTime())) { d2.setHours(0, 0, 0, 0); return d2; }
  return null;
}

function looksLikeSapDataRow(line) {
  if (!Array.isArray(line) || line.length <= SAP_COL_IDX.nf) return false;
  const cliente = line[SAP_COL_IDX.cliente];
  if (cliente === null || cliente === undefined || String(cliente).trim() === '') return false;
  if (!looksLikeSapNfCell(line[SAP_COL_IDX.nf])) return false;
  return !!parseSapBrDate(line[SAP_COL_IDX.dtEmissao]);
}

function isStandardSapNfLayout(headerRow, sampleRows) {
  if (headerRow &&
      isSapHeaderAtCol(headerRow, SAP_COL_IDX.cliente, 'cliente') &&
      isSapHeaderAtCol(headerRow, SAP_COL_IDX.dtEmissao, 'dtEmissao') &&
      isSapHeaderAtCol(headerRow, SAP_COL_IDX.nf, 'nf')) return true;
  let hits = 0;
  for (const line of (sampleRows || []).slice(0, 5)) {
    if (looksLikeSapDataRow(line)) hits++;
  }
  return hits >= 2;
}

function applySapColPositionalFallback(row, line, headerRow, standardLayout) {
  if (!line) return row;
  const useLayout = standardLayout || isStandardSapNfLayout(headerRow) || looksLikeSapDataRow(line);
  if (!useLayout) return row;

  const nf = line[SAP_COL_IDX.nf];
  const cliente = line[SAP_COL_IDX.cliente];
  const dtRaw = line[SAP_COL_IDX.dtEmissao];

  if (nf !== null && nf !== undefined && String(nf).trim() !== '') row.nf = nf;
  if (cliente !== null && cliente !== undefined && String(cliente).trim() !== '') {
    row.cliente = String(cliente).trim();
  }
  const parsed = parseSapBrDate(dtRaw);
  if (parsed) {
    row.dtEmissao = parsed;
    row._sapDateSource = 'colC';
  } else if (dtRaw !== null && dtRaw !== undefined && dtRaw !== '') {
    row.dtEmissao = dtRaw;
    row._sapDateSource = 'colC-raw';
  }
  row._sapPositionalLayout = true;
  return row;
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
  if (headerIdx < 0) headerIdx = 0;

  const headerRow = raw[headerIdx] || [];
  const headers = headerRow.map(h => String(h || '').trim()).filter(Boolean);
  const dataLines = [];
  for (let i = headerIdx + 1; i < raw.length; i++) {
    const line = raw[i];
    if (!Array.isArray(line) || line.every(c => c === null || c === undefined || c === '')) continue;
    dataLines.push(line);
  }

  const standardLayout = isStandardSapNfLayout(headerRow, dataLines);
  if (standardLayout) {
    console.debug('[SAP NF] Layout padrão: col. B=cliente, C=data emissão, D=nota fiscal');
  }

  const rows = [];
  for (const line of dataLines) {
    const obj = {};
    headerRow.forEach((h, j) => { if (h) obj[h] = line[j] ?? null; });
    let row = normalizeSapRow(obj);
    row = applySapColPositionalFallback(row, line, headerRow, standardLayout);
    rows.push(row);
  }
  return { rows, headers };
}

function normalizeSapRow(row) {
  return {
    nf: findSapField(row, 'nf'),
    dtEmissao: findSapField(row, 'dtEmissao'),
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

function isSapHeaderAtCol(headerRow, colIdx, field) {
  const h = normCol(headerRow?.[colIdx]);
  if (!h) return false;
  return SAP_ALIASES[field].some(alias => h === alias || h.includes(alias) || alias.includes(h));
}

function looksLikeSapNfCell(v) {
  if (v === null || v === undefined || v === '') return false;
  const digits = String(v).trim().replace(/\D/g, '');
  return digits.length >= 4 && digits.length <= 12;
}

function looksLikeSapDateCell(v) {
  if (v === null || v === undefined || v === '') return false;
  if (v instanceof Date) return !isNaN(v);
  if (typeof v === 'number' && v > 1000 && v < 100000) return true;
  const s = String(v).trim();
  return /^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(s) || /^\d{4}-\d{2}-\d{2}/.test(s);
}

function isStandardSapLayout(headerRow) {
  if (!headerRow?.length) return false;
  if (isSapHeaderAtCol(headerRow, SAP_COL_IDX.nf, 'nf')) return true;
  return isSapHeaderAtCol(headerRow, SAP_COL_IDX.dtEmissao, 'dtEmissao') &&
    isSapHeaderAtCol(headerRow, SAP_COL_IDX.cliente, 'cliente');
}

function looksLikeSapDataRow(line) {
  return looksLikeSapNfCell(line?.[SAP_COL_IDX.nf]) && looksLikeSapDateCell(line?.[SAP_COL_IDX.dtEmissao]);
}

function parseSapDate(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) return isNaN(val) ? null : val;
  if (typeof val === 'number') {
    if (val > 1000 && val < 100000) {
      const epoch = Date.UTC(1899, 11, 30);
      const d = new Date(epoch + val * 86400000);
      return isNaN(d) ? null : d;
    }
    return null;
  }
  const s = String(val).trim();
  const br = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (br) {
    const y = br[3].length === 2
      ? (Number(br[3]) > 50 ? 1900 + Number(br[3]) : 2000 + Number(br[3]))
      : Number(br[3]);
    const d = new Date(y, Number(br[2]) - 1, Number(br[1]));
    return isNaN(d) ? null : d;
  }
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    const d = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return isNaN(d) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function applySapColPositionalFallback(row, line, headerRow) {
  if (!line) return row;
  const useLayout = isStandardSapLayout(headerRow) || looksLikeSapDataRow(line);
  if (!useLayout) return row;

  const nf = line[SAP_COL_IDX.nf];
  const cliente = line[SAP_COL_IDX.cliente];
  const dtRaw = line[SAP_COL_IDX.dtEmissao];

  if (nf !== null && nf !== undefined && nf !== '') row.nf = nf;
  if (cliente !== null && cliente !== undefined && String(cliente).trim() !== '') {
    row.cliente = String(cliente).trim();
  }
  const parsed = parseSapDate(dtRaw);
  if (parsed) {
    row.dtEmissao = parsed;
    row._sapDateSource = 'colC';
  } else if (dtRaw !== null && dtRaw !== undefined && dtRaw !== '') {
    row.dtEmissao = dtRaw;
    row._sapDateSource = 'colC-raw';
  }
  row._sapPositionalLayout = true;
  return row;
}

function summarizeSapDateSources(rows) {
  const counts = {};
  rows.forEach(r => {
    const src = r._sapDateSource || (r._sapPositionalLayout ? 'colC' : 'alias');
    counts[src] = (counts[src] || 0) + 1;
  });
  return counts;
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
    const parsed = parseSheetRows(wb.Sheets[name], normalizeSapRow, isSapHeaderRow, applySapColPositionalFallback);
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
    const entry = {
      nf,
      cliente: r.cliente ? String(r.cliente).trim() : '',
      dtEmissao: parseSapDate(r.dtEmissao) || r.dtEmissao,
      valorNF: num(r.valorNF)
    };
    const key = normNFKey(nf);
    map[key] = entry;
    const raw = String(nf).trim();
    if (raw !== key) map[raw] = entry;
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

function applySapToNf(nf) {
  if (nf.valorUnilog === undefined || nf.valorUnilog === null) {
    nf.valorUnilog = num(nf.valorNF);
  }

  const key = normNFKey(nf.nf);
  const sap = sapNfMap[key] || sapNfMap[String(nf.nf).trim()];
  nf.sapFound = !!sap;

  if (!sap) {
    nf.valorSAP = null;
    nf.valorDiff = null;
    nf.sapValorMismatch = false;
    return nf;
  }

  if (sap.cliente) nf.cliente = sap.cliente;

  if (sap.dtEmissao) {
    nf.dtNF = sap.dtEmissao;
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

function handleFile(file) {
  $('loadbar').style.display = 'block';
  $('loadbar').textContent = 'A ler ' + file.name + ' ...';
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const { rows, sheetName, headers } = loadRowsFromWorkbook(wb);
      if (!rows.length) {
        $('loadbar').textContent =
          'Nenhuma linha de dados encontrada. Folhas: ' + wb.SheetNames.join(', ');
        return;
      }
      processRows(rows, file.name, sheetName, headers);
    } catch (err) {
      console.error(err);
      $('loadbar').textContent = 'Erro a ler o ficheiro: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleSapFile(file) {
  const statusEl = $('sapLoadStatus');
  if (statusEl) statusEl.textContent = 'A ler ' + file.name + ' ...';
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const { rows, sheetName } = loadSapRowsFromWorkbook(wb);
      if (!rows.length) {
        if (statusEl) statusEl.textContent = 'Nenhuma linha SAP encontrada na folha "' + sheetName + '".';
        return;
      }
      sapNfMap = buildSapNfMap(rows);
      const nKeys = new Set(Object.values(sapNfMap).map(x => x.nf)).size;
      const dateSources = summarizeSapDateSources(rows);
      console.log('[SAP] Data NF — fontes por coluna:', dateSources,
        dateSources.colC || dateSources['colC-raw'] ? '(col. C)' : '(alias cabeçalho)');
      if (statusEl) statusEl.textContent = rows.length + ' linhas SAP · ' + nKeys + ' NFs mapeadas';
      fteToast('Dados SAP carregados.');

      if (lastCtePack) {
        processRows(lastCtePack.rows, lastCtePack.fileName, lastCtePack.sheetName, lastCtePack.headers);
      } else if (currentNFs.length) {
        currentNFs = applySapToList(currentNFs.map(nf => {
          delete nf.mesRef;
          delete nf.dtRef;
          return nf;
        }));
        currentSummary = computeSummary(currentNFs, currentSummary?.fileName || 'SAP enrich');
        renderAll();
        const nMismatch = currentSummary.nValorMismatch || 0;
        if (nMismatch) fteToast(nMismatch + ' NF(s) com diferença relevante de valor SAP vs Unilog.');
      }
    } catch (err) {
      console.error(err);
      if (statusEl) statusEl.textContent = 'Erro a ler SAP: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
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
    sapFound: false, sapValorMismatch: false,
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
    else $('uploadZone').style.display = 'block';
    $('results').style.display = 'none';
    fteToast('Não foi possível ler notas fiscais — verifica se o ficheiro tem a coluna "Nota Fiscal".');
    return;
  }

  renderAll();
  showResultsView();
  autosaveToCloud();
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

  let nValorMismatch = 0, nSapMatched = 0;
  list.forEach(x => {
    byStatus[x.status]++;
    excessoByStatus[x.status] += x.diff;
    pagoByStatus[x.status] += x.pago;
    if (x.nCte === 1) nUnicoCte++; else nMultiplosCte++;
    if (x.diff > 0) excessoPositivo += x.diff;
    else if (x.diff < 0) deficit += x.diff;
    if (x.sapFound) nSapMatched++;
    if (x.sapValorMismatch) nValorMismatch++;
  });

  return {
    fileName, totalNF: list.length, totalPago, totalValorNF, totalEsperado, excesso,
    pctPagoSobreNF: totalValorNF ? totalPago / totalValorNF : 0,
    excessoPositivo, deficit, excessoPct: totalEsperado ? excesso / totalEsperado : 0,
    periodoInicio, periodoFim, byStatus, excessoByStatus, pagoByStatus,
    nUnicoCte, nMultiplosCte, nValorMismatch, nSapMatched,
    sapLoaded: Object.keys(sapNfMap).length > 0
  };
}

function fmtDate(d) {
  if (!d) return '-';
  const dt = d instanceof Date ? d : new Date(d);
  return isNaN(dt) ? '-' : dt.toLocaleDateString('pt-BR');
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

function enrichNF(nf) {
  if (nf.mesRef) return nf;
  let d = nf.dtNF ? new Date(nf.dtNF) : null;
  if ((!d || isNaN(d)) && nf.ctes?.length) {
    const dates = nf.ctes.map(c => c.dtCte).filter(Boolean).map(x => new Date(x)).filter(x => !isNaN(x));
    if (dates.length) d = new Date(Math.min(...dates));
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
  })).sort((a, b) => a.mesRef.localeCompare(b.mesRef));
}

function sortValue(row, col, type) {
  if (type === 'status') return (statusMeta[row.status] || {}).label || row.status || '';
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
    'Motivo': x.motivo
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
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumo), 'Resumo');

  const byCat = [];
  Object.keys(CATEGORY_META).forEach(key => {
    currentNFs.filter(x => x.status === key).forEach(x => byCat.push({ ...nfExportRow(x), '_Categoria': CATEGORY_META[key].title }));
  });
  if (byCat.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(byCat), 'Por categoria');
  }

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(currentNFs.map(nfExportRow)), 'Notas Fiscais');

  const ctes = cteExportRows();
  if (ctes.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ctes), 'CT-e');

  const mensal = monthlyRows.map(m => ({
    'Mês': m.mesLabel,
    'NFs': m.totalNF,
    'Faturação': m.totalValorNF,
    'Pago': m.totalPago,
    'Esperado (6%)': m.totalEsperado,
    'Excesso': m.excesso,
    '% pago/NF': m.pctPagoSobreNF,
    'Tarifa mínima': m.qtd_min,
    'A investigar': m.qtd_flag,
    'Conformes': m.qtd_ok
  }));
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
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(monthlyRows.map(m => ({
    'Mês': m.mesLabel,
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
  }))), 'Resumo mensal');
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
  let list = currentNFs;
  if (statusF) list = list.filter(x => x.status === statusF);
  if (transpF) list = list.filter(x => x.transportador === transpF);
  if (mesF) list = list.filter(x => x.mesRef === mesF);
  if (valorDiffF) list = list.filter(x => x.sapValorMismatch);
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
  const sorted = sortRows(monthlyRows, MONTH_COLUMNS, monthSort);
  const maxExcesso = Math.max(...sorted.map(m => m.excesso), 0);

  const body = $('monthTableBody');
  body.innerHTML = sorted.map(m => {
    const high = m.excesso > 0 && m.excesso >= maxExcesso * 0.85;
    const active = selectedMonth === m.mesRef;
    return `<tr class="month-row-clickable${high ? ' month-row-high' : ''}${active ? ' active-month' : ''}" data-mes="${m.mesRef}">
      <td><strong>${m.mesLabel}</strong></td>
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
  $('periodoLabel').innerHTML =
    `Ficheiro: ${s.fileName} · Período CT-e: ${fmtDate(s.periodoInicio)} a ${fmtDate(s.periodoFim)} · ${s.totalNF} notas fiscais · ${s.nUnicoCte} com CT-e único · ${s.nMultiplosCte} com múltiplos CT-e${sapNote}${mismatchNote}`;

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
    <div class="kpi${s.nValorMismatch ? ' flag' : ''}"><div class="label">Δ valor SAP ≠ Unilog</div><div class="value">${s.nValorMismatch}</div><div class="sub">${s.nSapMatched} NFs cruzadas com SAP</div></div>` : '';

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
  renderTable();
  renderMonthlyTable();
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
    tr.className = 'row-clickable' + (x.sapValorMismatch ? ' nf-val-mismatch' : '');
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
  const sapBlock = x.sapFound ? `<p style="margin:0 0 8px;font-size:11px;color:var(--muted);">
    <strong>SAP:</strong> ${fmtDate(x.dtNF)} · ${x.cliente || '-'} · Unilog ${fmtMoney(x.valorUnilog ?? x.valorNF)} vs SAP ${fmtMoney(x.valorSAP)}${x.sapValorMismatch ? ` — <span style="color:#b3261e;font-weight:600">Δ relevante ${fmtMoney(x.valorDiff)}</span>` : (x.valorDiff != null ? ` (Δ ${fmtMoney(x.valorDiff)})` : '')}
  </p>` : '';
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
      if (t.dataset.tab === 'historico') loadHistory();
      if (t.dataset.tab === 'mensal') renderMonthlyTable();
    });
  });

  const uploadZone = $('uploadZone');
  const fileInput = $('fileInput');
  uploadZone.addEventListener('click', () => fileInput.click());
  uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.classList.add('drag'); });
  uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag'));
  uploadZone.addEventListener('drop', e => {
    e.preventDefault(); uploadZone.classList.remove('drag');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', e => { if (e.target.files.length) handleFile(e.target.files[0]); });

  const sapUploadZone = $('sapUploadZone');
  const sapFileInput = $('sapFileInput');
  if (sapUploadZone && sapFileInput) {
    sapUploadZone.addEventListener('click', () => sapFileInput.click());
    sapUploadZone.addEventListener('dragover', e => { e.preventDefault(); sapUploadZone.classList.add('drag'); });
    sapUploadZone.addEventListener('dragleave', () => sapUploadZone.classList.remove('drag'));
    sapUploadZone.addEventListener('drop', e => {
      e.preventDefault(); sapUploadZone.classList.remove('drag');
      if (e.dataTransfer.files.length) handleSapFile(e.dataTransfer.files[0]);
    });
    sapFileInput.addEventListener('change', e => { if (e.target.files.length) handleSapFile(e.target.files[0]); });
  }

  $('newFileBtn').addEventListener('click', () => {
    $('results').style.display = 'none';
    const us = $('uploadSection');
    if (us) us.style.display = 'block';
    else $('uploadZone').style.display = 'block';
  });
  $('exportBtn').addEventListener('click', () => exportFullWorkbook());
  $('exportFilteredBtn').addEventListener('click', () => exportFilteredView());
  $('exportHistoryBtn').addEventListener('click', () => exportHistoryWorkbook());
  $('exportMensalBtn').addEventListener('click', () => exportMonthlyWorkbook());

  ['filterStatus', 'filterTransp', 'filterMes', 'searchNF', 'filterValorDiff'].forEach(id => {
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

  loadLatestFromCloud();
}
