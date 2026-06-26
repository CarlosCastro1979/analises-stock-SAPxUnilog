// armazem.js v1.0.6
const ARMAZEM_JS_VERSION = '1.0.6';

const ARM_MINIMO_CONTRATUAL = 120000;
const ARM_NF_RATE = 0.055;
const ARM_PERSIST_MAX_JSON_BYTES = 6 * 1024 * 1024;
const ARM_MAX_FILE_BYTES = 15 * 1024 * 1024;
const ARM_PARSE_TIMEOUT_MS = 90_000;

const $arm = id => document.getElementById('arm-' + id);

let armInited = false;
let armPack = null;
let armPendingFiles = [];
let armActiveTab = 'carregamento';
let armCatalogOverrides = {};
let armSorts = {
  mensal: { col: 'mesKey', dir: 1 },
  acumulado: { col: 'valor', dir: -1 },
  nf: { col: 'nf', dir: 1 },
  adicionais: { col: 'valor', dir: -1 },
  catalogo: { col: 'servico', dir: 1 }
};

const CATALOGO_DESPESAS = [
  { id: 'nf_percentual', label: '5,5% sobre NF expedida' },
  { id: 'armazenagem', label: 'Armazenagem' },
  { id: 'hora_extra', label: 'Hora extra' },
  { id: 'etiquetagem', label: 'Etiquetagem' },
  { id: 'readequacao', label: 'Readequação de produtos' },
  { id: 'impostos', label: 'Impostos / taxas' },
  { id: 'outros', label: 'Outros (validar)' }
];

const ARM_MESES_PT = {
  janeiro: 1, jan: 1,
  fevereiro: 2, fev: 2,
  marco: 3, março: 3, mar: 3,
  abril: 4, abr: 4,
  maio: 5, mai: 5,
  junho: 6, jun: 6,
  julho: 7, jul: 7,
  agosto: 8, ago: 8,
  setembro: 9, set: 9,
  outubro: 10, out: 10,
  novembro: 11, nov: 11,
  dezembro: 12, dez: 12
};

const ARM_MESES_NOME = [
  '', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
];

function armCompany() {
  return typeof company !== 'undefined' ? company : 'DFB';
}

function armSlot() {
  return typeof EXCEL_SLOTS !== 'undefined' ? EXCEL_SLOTS.ARMAZEM : 'armazem_monthly';
}

function armToast(msg, type) {
  if (typeof toast === 'function') toast(msg, type || 'success');
  else console.log('[armazem]', msg);
}

function armFmtBytes(n) {
  const b = Number(n) || 0;
  if (b < 1024 * 1024) return Math.round(b / 1024) + ' KB';
  return (b / (1024 * 1024)).toFixed(1) + ' MB';
}

function armYield() {
  return new Promise(r => setTimeout(r, 0));
}

function armWithTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error(
      `Timeout (${Math.round(ms / 1000)}s) ao processar ${label || 'ficheiro'} — ficheiro demasiado grande ou complexo.`
    )), ms))
  ]);
}

function parseBrNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  let s = String(v).trim().replace(/\s/g, '');
  if (!s || s === '-') return 0;
  s = s.replace(/^R\$\s?/i, '');

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

function armNum(v) {
  return parseBrNum(v);
}

function armFmtMoney(v) {
  if (typeof fmtMoney === 'function') return fmtMoney(v);
  return 'R$ ' + (v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtArmNum(n, maxDec) {
  if (n == null || n === '') return '—';
  const v = Number(n);
  if (!Number.isFinite(v)) return '—';
  const md = maxDec == null ? 2 : maxDec;
  const scale = Math.pow(10, md);
  const rounded = Math.round(v * scale) / scale;
  const eps = 1 / (scale * 10000);
  if (Math.abs(rounded - Math.round(rounded)) < eps) {
    return Math.round(rounded).toLocaleString('pt-BR', { maximumFractionDigits: 0 });
  }
  return rounded.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: md });
}

function fmtArmQty(n) { return fmtArmNum(n, 2); }
function fmtArmUnit(n) { return fmtArmNum(n, 2); }

function fmtArmPct(n, maxDec) {
  if (n == null || n === '') return '—';
  const v = Number(n);
  if (!Number.isFinite(v)) return '—';
  const md = maxDec == null ? 1 : maxDec;
  return v.toLocaleString('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: md }) + '%';
}

function armEsc(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
}

function normalizeServicoName(raw) {
  let s = String(raw || '').trim();
  s = s.replace(/^\[[\d.]+\]\s*/i, '');
  s = s.replace(/\s+/g, ' ').toUpperCase();
  if (/PERCENTUAL\s+SOBRE\s+NF\s+EXPEDID|ARMAZENAGEM\s+POR\s*%|5[,.]5\s*%/.test(s)) return 'PERCENTUAL SOBRE NF EXPEDIDA';
  if (/ARMAZENAGEM\s+EXCEDENTE/.test(s)) return 'ARMAZENAGEM EXCEDENTE';
  if (/ARMAZENAGEM\s+POR\s+POSI/.test(s)) return 'ARMAZENAGEM POR POSICAO PALLET';
  if (/ETIQUETAGEM/.test(s)) return 'ETIQUETAGEM EAN - POR UNIDADE';
  if (/READEQUA/.test(s)) return 'READEQUACAO DE PRODUTOS - POR UNIDADE';
  if (/HORA\s+EXTRA|H\.\s*E|H\.E|\bHE\b/.test(s)) return 'HORA EXTRA';
  if (/DECARGA\s+DE\s+MERCADORIA/.test(s)) return 'DECARGA DE MERCADORIA - POR PALLET';
  if (/IMPOSTO|PIS|COFINS|ISS/.test(s)) return 'IMPOSTOS SERVICOS';
  return s;
}

function armMesNomeLong(mesKey, mesLabel) {
  const m = String(mesKey || '').match(/^(\d{4})-(\d{2})$/);
  if (m) {
    const mi = parseInt(m[2], 10);
    return `${ARM_MESES_NOME[mi] || m[2]} ${m[1]}`;
  }
  if (mesLabel) return mesLabel;
  return mesKey || '—';
}

function armMesSheetName(mesKey) {
  const m = String(mesKey || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return 'Mes';
  const mi = parseInt(m[2], 10);
  const short = (ARM_MESES_NOME[mi] || m[2]).slice(0, 3);
  return `${short} ${m[1]}`.slice(0, 31);
}

function armResumoRows(month) {
  return (month.servicos || []).filter(s => s.normName !== 'IMPOSTOS SERVICOS' && !/^IMPOSTO/i.test(s.normName));
}

function armFmtResumoQty(v, isNf) {
  if (isNf) return armFmtMoney(v);
  return fmtArmQty(v);
}

function armFmtResumoUnit(v, isNf) {
  if (isNf) return fmtArmPct(ARM_NF_RATE * 100);
  return fmtArmUnit(v);
}

function armResumoRowHtml(s, opts) {
  const isNf = !!s.isNf;
  const nome = opts?.useNorm ? s.normName : (s.rawName || s.normName);
  return `<tr>
    <td>${armEsc(nome)}</td>
    <td class="right">${armFmtResumoQty(s.qtde, isNf)}</td>
    <td class="right">${armFmtResumoUnit(s.valorUnit, isNf)}</td>
    <td class="right">${armFmtMoney(s.valor)}</td>
  </tr>`;
}

function aggregateResumoByNorm(months) {
  const map = {};
  months.forEach(m => {
    armResumoRows(m).forEach(s => {
      const norm = normalizeServicoName(s.rawName);
      if (!map[norm]) {
        map[norm] = { normName: norm, qtde: 0, valor: 0, unitSum: 0, isNf: false, rawVariants: new Set() };
      }
      const row = map[norm];
      row.qtde += s.qtde || 0;
      row.valor += s.valor || 0;
      if (s.qtde && s.valorUnit) row.unitSum += s.qtde * s.valorUnit;
      if (s.isNf) row.isNf = true;
      row.rawVariants.add(s.rawName);
    });
  });
  return Object.values(map).map(r => ({
    normName: r.normName,
    rawName: r.normName,
    qtde: r.qtde,
    valor: r.valor,
    valorUnit: r.isNf ? ARM_NF_RATE : (r.qtde ? r.unitSum / r.qtde : 0),
    isNf: r.isNf,
    rawVariants: [...r.rawVariants]
  })).sort((a, b) => a.normName.localeCompare(b.normName, 'pt'));
}

function resumoTableHtml(rows, opts) {
  const qtyHdr = opts?.total ? 'Qtde Serviço / Valor Faturação' : 'Qtde Serviço';
  const valHdr = opts?.total ? 'Valor Calculado / Pago' : 'Valor Calculado';
  const body = rows.map(s => armResumoRowHtml(s, opts)).join('');
  const subtotal = rows.reduce((s, r) => s + (r.valor || 0), 0);
  return `
    <div class="tbl-wrap arm-resumo-tbl">
      <table>
        <thead><tr>
          <th>Tipo de Serviço</th>
          <th class="right">${qtyHdr}</th>
          <th class="right">Valor Unitário</th>
          <th class="right">${valHdr}</th>
        </tr></thead>
        <tbody>${body}
          <tr class="arm-resumo-subtotal"><td><strong>Subtotal</strong></td><td></td><td></td><td class="right"><strong>${armFmtMoney(subtotal)}</strong></td></tr>
        </tbody>
      </table>
    </div>`;
}

function catalogStorageKey() {
  return 'armazem_catalog_' + armCompany();
}

function loadCatalogOverrides() {
  try {
    armCatalogOverrides = JSON.parse(localStorage.getItem(catalogStorageKey()) || '{}') || {};
  } catch (e) {
    armCatalogOverrides = {};
  }
}

function saveCatalogOverrides() {
  localStorage.setItem(catalogStorageKey(), JSON.stringify(armCatalogOverrides));
}

function guessCatalogCategory(normName) {
  const n = normalizeServicoName(normName);
  if (n === 'PERCENTUAL SOBRE NF EXPEDIDA') return { id: 'nf_percentual', sure: true };
  if (/ARMAZENAGEM/.test(n)) return { id: 'armazenagem', sure: true };
  if (/HORA\s+EXTRA/.test(n)) return { id: 'hora_extra', sure: true };
  if (/ETIQUETAGEM/.test(n)) return { id: 'etiquetagem', sure: true };
  if (/READEQUACAO/.test(n)) return { id: 'readequacao', sure: true };
  if (/IMPOSTO/.test(n)) return { id: 'impostos', sure: true };
  return { id: 'outros', sure: false };
}

function resolveCatalogCategory(rawName) {
  const norm = normalizeServicoName(rawName);
  if (armCatalogOverrides[norm]) return { id: armCatalogOverrides[norm], sure: true, norm, raw: rawName };
  const g = guessCatalogCategory(norm);
  return { id: g.id, sure: g.sure, norm, raw: rawName };
}

function catalogLabel(id) {
  return CATALOGO_DESPESAS.find(c => c.id === id)?.label || id;
}

function parseMonthFromFileName(fileName) {
  const base = String(fileName || '').replace(/\.xlsx?$/i, '');
  const m = base.match(/(\d{4})\s*$/);
  const year = m ? parseInt(m[1], 10) : null;
  const low = base.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  let month = null;
  for (const [name, num] of Object.entries(ARM_MESES_PT)) {
    const key = name.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (low.includes(key)) { month = num; break; }
  }
  if (!month || !year) return { mesKey: '', mesLabel: base };
  const mesKey = `${year}-${String(month).padStart(2, '0')}`;
  const mesLabel = `${String(month).padStart(2, '0')}/${year}`;
  return { mesKey, mesLabel };
}

function sheetToMatrix(ws) {
  if (!ws || !ws['!ref']) return [];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
}

function cellText(matrix, r, c) {
  const row = matrix[r];
  if (!row) return '';
  const v = row[c];
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  return String(v).trim();
}

function armCellNum(matrix, r, c) {
  return armNum(cellText(matrix, r, c));
}

function normHdrCell(v) {
  return String(v || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function findResumoHeaderRow(matrix) {
  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (/tipo de servi/i.test(String(row[c] || ''))) return r;
    }
  }
  return -1;
}

function scanResumoCols(matrix, hdrRow, fallback) {
  const cols = { ...(fallback || { nome: 1, qtde: 2, unit: 3, valor: 4 }) };
  const row = matrix[hdrRow] || [];
  for (let c = 0; c < row.length; c++) {
    const h = normHdrCell(row[c]);
    if (!h) continue;
    if (/qtde|quantidade/.test(h)) cols.qtde = c;
    else if (/valor unit|unitario/.test(h)) cols.unit = c;
    else if (/valor calc|calculado/.test(h)) cols.valor = c;
  }
  for (let c = 0; c < row.length; c++) {
    const h = normHdrCell(row[c]);
    if (/codigo servi|cod servi/.test(h)) cols.nome = c;
  }
  if (cols.nome === (fallback?.nome ?? 1)) {
    for (let c = 0; c < row.length; c++) {
      const h = normHdrCell(row[c]);
      if (/tipo de servi/.test(h)) cols.nome = c;
    }
  }
  return cols;
}

function armFinishServico(rawName, qtde, valorUnit, valor, isNf) {
  const q = armNum(qtde);
  if (isNf) {
    const v = q * ARM_NF_RATE;
    return { qtde: q, valorUnit: ARM_NF_RATE, valor: v };
  }
  const u = armNum(valorUnit);
  const calc = q * u;
  return { qtde: q, valorUnit: u, valor: calc > 0 ? calc : armNum(valor) };
}

function findSheetName(wb, pattern) {
  const re = typeof pattern === 'string' ? new RegExp(pattern, 'i') : pattern;
  return wb.SheetNames.find(n => re.test(n)) || null;
}

function parseBrDate(s) {
  const t = String(s || '').trim();
  const m = t.match(/^(\d{2})[./](\d{2})[./](\d{2,4})/);
  if (!m) return t;
  const y = m[3].length === 2 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10);
  return `${y}-${m[2]}-${m[1]}`;
}

function isNfPercentualService(name) {
  return normalizeServicoName(name) === 'PERCENTUAL SOBRE NF EXPEDIDA';
}

function parseArmazemV2(wb, fileName) {
  const infoName = findSheetName(wb, /^Informa/);
  const resumoName = findSheetName(wb, /^Resumo de Servi/);
  const diarioName = findSheetName(wb, /^Detalhamento Di/);
  if (!resumoName) throw new Error('Folha "Resumo de Serviços" não encontrada (formato v2).');

  const meta = parseMonthFromFileName(fileName);
  const info = infoName ? sheetToMatrix(wb.Sheets[infoName]) : [];
  const resumo = sheetToMatrix(wb.Sheets[resumoName]);

  let dataInicial = '';
  let dataFinal = '';
  let totalServicos = 0;
  let valorMinimo = ARM_MINIMO_CONTRATUAL;
  let valorApurado = 0;
  let valorTotal = 0;

  let depositante = '';
  for (let r = 0; r < info.length; r++) {
    const a = cellText(info, r, 0);
    const b = cellText(info, r, 1);
    if (/Depositante/i.test(a)) depositante = b;
    if (/Data Inicial/i.test(a)) dataInicial = parseBrDate(b);
    if (/Data Final/i.test(a)) dataFinal = parseBrDate(b);
    if (/Total dos Servi/i.test(a)) totalServicos = armNum(b);
    if (/Valor M[ií]nimo Contratual/i.test(a)) valorMinimo = armNum(b) || ARM_MINIMO_CONTRATUAL;
    if (/Valor Apurado/i.test(a)) valorApurado = armNum(b);
    if (/VALOR FINAL A FATURAR|Valor Total \(com impostos\)/i.test(a)) valorTotal = armNum(b) || valorTotal;
  }

  if (!meta.mesKey && dataInicial) meta.mesKey = dataInicial.slice(0, 7);
  if (meta.mesKey && !meta.mesLabel) {
    const [y, m] = meta.mesKey.split('-');
    meta.mesLabel = `${m}/${y}`;
  }

  const servicos = [];
  const hdr = findResumoHeaderRow(resumo);
  if (hdr < 0) throw new Error('Cabeçalho do resumo não encontrado.');
  const cols = scanResumoCols(resumo, hdr);

  for (let r = hdr + 1; r < resumo.length; r++) {
    const nome = cellText(resumo, r, cols.nome) || cellText(resumo, r, 0);
    const codigo = cellText(resumo, r, cols.nome) || nome;
    if (!nome) continue;
    if (/^Apura/i.test(nome) || /^Total$/i.test(nome) || /^Impostos/i.test(nome)) continue;
    const isNf = isNfPercentualService(nome);
    const nums = armFinishServico(
      nome,
      armCellNum(resumo, r, cols.qtde),
      armCellNum(resumo, r, cols.unit),
      armCellNum(resumo, r, cols.valor),
      isNf
    );
    if (!nums.valor && !nums.qtde) continue;
    const norm = normalizeServicoName(nome);
    const cat = resolveCatalogCategory(nome);
    servicos.push({
      rawName: nome, normName: norm, codigo,
      qtde: nums.qtde, valorUnit: nums.valorUnit, valor: nums.valor,
      catalogId: cat.id, catalogSure: cat.sure, isNf
    });
  }

  const nfRows = [];
  const nfSheet = wb.SheetNames.find(n => /PERCENTUAL SOBRE NF/i.test(n));
  if (nfSheet) {
    const m = sheetToMatrix(wb.Sheets[nfSheet]);
    let dataHdr = -1;
    for (let r = 0; r < m.length; r++) {
      if (/Data Cobran/i.test(cellText(m, r, 0)) && /Documento/i.test(cellText(m, r, 1))) { dataHdr = r; break; }
    }
    if (dataHdr >= 0) {
      for (let r = dataHdr + 1; r < m.length; r++) {
        const dt = cellText(m, r, 0);
        const nf = cellText(m, r, 1);
        const valorNF = armCellNum(m, r, 2);
        const fee = armCellNum(m, r, 3);
        if (!nf || nf === '-') continue;
        nfRows.push({
          data: parseBrDate(dt), nf, valorNF, taxa: ARM_NF_RATE,
          fee, mesKey: meta.mesKey, fileName
        });
      }
    }
  } else if (diarioName) {
    const m = sheetToMatrix(wb.Sheets[diarioName]);
    let hdrD = -1;
    for (let r = 0; r < m.length; r++) {
      if (/^Data$/i.test(cellText(m, r, 0)) && /Tipo de Servi/i.test(cellText(m, r, 1))) { hdrD = r; break; }
    }
    if (hdrD >= 0) {
      for (let r = hdrD + 1; r < m.length; r++) {
        const dt = cellText(m, r, 0);
        const svc = cellText(m, r, 1);
        if (!isNfPercentualService(svc)) continue;
        const nf = cellText(m, r, 2);
        const valorNF = armCellNum(m, r, 3);
        const fee = armCellNum(m, r, 4);
        if (!nf || nf === '-') continue;
        nfRows.push({
          data: parseBrDate(dt), nf, valorNF, taxa: ARM_NF_RATE, fee,
          mesKey: meta.mesKey, fileName
        });
      }
    }
  }

  const totalCalc = servicos.filter(s => !/IMPOSTO/i.test(s.normName)).reduce((a, s) => a + s.valor, 0);
  return buildMonthRecord({
    fileName, format: 'v2', meta, depositante, dataInicial, dataFinal,
    totalServicos: totalServicos || totalCalc,
    valorMinimo, valorApurado: valorApurado || totalCalc, valorTotal,
    servicos, nfRows
  });
}

function parseArmazemV1(wb, fileName) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const m = sheetToMatrix(sheet);
  const meta = parseMonthFromFileName(fileName);

  let depositante = '';
  for (let r = 0; r < Math.min(20, m.length); r++) {
    for (let c = 0; c < 10; c++) {
      const v = cellText(m, r, c);
      if (/DELTA FOODS|14830817000100/i.test(v)) depositante = v;
    }
  }

  const servicos = [];
  let resumoHdr = -1;
  for (let r = 0; r < m.length; r++) {
    if (/^\s*Resumo\s*$/i.test(cellText(m, r, 1))) { resumoHdr = r; break; }
  }
  if (resumoHdr >= 0) {
    const hdrRow = resumoHdr + 1;
    const cols = scanResumoCols(m, hdrRow, { nome: 1, qtde: 17, unit: 25, valor: 28 });
    for (let r = resumoHdr + 2; r < m.length; r++) {
      const nome = cellText(m, r, cols.nome);
      if (!nome) continue;
      if (/^Apura/i.test(nome) || /IMPOSTO/i.test(nome)) continue;
      const isNf = isNfPercentualService(nome);
      const nums = armFinishServico(
        nome,
        armCellNum(m, r, cols.qtde),
        armCellNum(m, r, cols.unit),
        armCellNum(m, r, cols.valor),
        isNf
      );
      if (!nome || (!nums.valor && !nums.qtde)) continue;
      const norm = normalizeServicoName(nome);
      const cat = resolveCatalogCategory(nome);
      servicos.push({
        rawName: nome, normName: norm, codigo: nome,
        qtde: nums.qtde, valorUnit: nums.valorUnit, valor: nums.valor,
        catalogId: cat.id, catalogSure: cat.sure, isNf
      });
    }
  }

  const nfRows = [];
  for (let r = 0; r < m.length; r++) {
    const svc = cellText(m, r, 2);
    if (!isNfPercentualService(svc)) continue;
    let hdr = r + 3;
    while (hdr < m.length && !/N[oº] Doc/i.test(cellText(m, hdr, 7)) && !/N[oº] Doc/i.test(cellText(m, hdr, 8))) hdr++;
    if (hdr >= m.length) continue;
    for (let rr = hdr + 1; rr < m.length; rr++) {
      if (/^Regra:/i.test(cellText(m, rr, 1)) || /^\s*Resumo\s*$/i.test(cellText(m, rr, 1))) break;
      const nf = cellText(m, rr, 7);
      const dt = cellText(m, rr, 1);
      const valorNF = armCellNum(m, rr, 28);
      const taxa = armCellNum(m, rr, 32) || ARM_NF_RATE;
      const fee = armCellNum(m, rr, 36);
      if (!nf) continue;
      nfRows.push({
        data: parseBrDate(dt), nf, valorNF, taxa, fee,
        mesKey: meta.mesKey, fileName
      });
    }
  }

  const totalCalc = servicos.filter(s => !/IMPOSTO/i.test(s.normName)).reduce((a, s) => a + s.valor, 0);
  let dataInicial = nfRows[0]?.data || '';
  let dataFinal = nfRows[nfRows.length - 1]?.data || '';

  return buildMonthRecord({
    fileName, format: 'v1', meta, depositante, dataInicial, dataFinal,
    totalServicos: totalCalc, valorMinimo: ARM_MINIMO_CONTRATUAL,
    valorApurado: totalCalc, valorTotal: totalCalc,
    servicos, nfRows
  });
}

function buildMonthRecord(opts) {
  const nfRows = (opts.nfRows || []).map(applySapToArmNf);
  const belowMin = opts.totalServicos < (opts.valorMinimo || ARM_MINIMO_CONTRATUAL);
  return {
    fileName: opts.fileName,
    format: opts.format,
    mesKey: opts.meta?.mesKey || '',
    mesLabel: opts.meta?.mesLabel || opts.fileName,
    dataInicial: opts.dataInicial || '',
    dataFinal: opts.dataFinal || '',
    depositante: opts.depositante || '',
    totalServicos: opts.totalServicos || 0,
    valorMinimo: opts.valorMinimo || ARM_MINIMO_CONTRATUAL,
    valorApurado: opts.valorApurado || 0,
    valorTotal: opts.valorTotal || 0,
    belowMin,
    servicos: opts.servicos || [],
    nfRows,
    adicionais: (opts.servicos || []).filter(s => !s.isNf && s.catalogId !== 'impostos')
  };
}

function parseArmazemWorkbook(wb, fileName) {
  const hasV2 = wb.SheetNames.some(n => /Resumo de Servi/i.test(n));
  const rec = hasV2 ? parseArmazemV2(wb, fileName) : parseArmazemV1(wb, fileName);
  if (armCompany() !== 'DFB') {
    rec.companyWarning = 'Módulo configurado para Delta Foods Brasil (DFB).';
  }
  if (rec.depositante && !/DELTA FOODS/i.test(rec.depositante)) {
    rec.depositWarning = 'Depositante não parece ser Delta Foods Brasil.';
  }
  return rec;
}

function sapApi() {
  return window.FretesSAP || null;
}

function applySapToArmNf(row) {
  const api = sapApi();
  const out = { ...row };
  out.valorUnilog = armNum(row.valorNF);
  out.feeUnilog = armNum(row.fee);
  out.feeExpected = out.valorUnilog * ARM_NF_RATE;
  out.feeDelta = out.feeUnilog - out.feeExpected;

  if (!api || !api.isSapLoaded()) {
    out.sapFound = false;
    out.sapMissing = false;
    out.sapValor = null;
    out.valorDiff = false;
    return out;
  }

  const sap = api.lookupSapEntry(row.nf);
  out.sapFound = !!sap;
  out.sapMissing = !sap;
  out.sapValor = sap ? armNum(sap.valorNF) : null;
  out.sapCliente = sap?.cliente || '';
  out.valorDiff = sap && api.isRelevantValorDiff(out.valorUnilog, out.sapValor);
  return out;
}

function refreshAllSapOnNfs() {
  if (!armPack?.months?.length) return;
  armPack.months.forEach(m => {
    m.nfRows = (m.nfRows || []).map(r => applySapToArmNf(r));
  });
}

function detectParserVersion(wb) {
  return wb.SheetNames.some(n => /Resumo de Servi/i.test(n)) ? 'v2' : 'v1';
}

async function readFileToWorkbook(file) {
  const buf = await file.arrayBuffer();
  try {
    return { wb: XLSX.read(buf, { type: 'array' }), buffer: buf };
  } catch (err) {
    const msg = typeof formatXlsxReadError === 'function' ? formatXlsxReadError(err) : String(err.message || err);
    throw new Error(msg);
  }
}

function mergeArmPack(prev, next) {
  if (!prev?.months?.length) return next;
  if (!next?.months?.length) {
    return {
      ...prev,
      failedFiles: [...(prev.failedFiles || []), ...(next.failedFiles || [])]
    };
  }
  const replace = new Set(next.months.map(m => m.mesKey || m.fileName));
  const months = [
    ...prev.months.filter(m => !replace.has(m.mesKey || m.fileName)),
    ...next.months
  ].sort((a, b) => String(a.mesKey).localeCompare(String(b.mesKey)));
  return {
    version: 2,
    updatedAt: new Date().toISOString(),
    months,
    failedFiles: [...(prev.failedFiles || []), ...(next.failedFiles || [])]
  };
}

function buildArmPackFromFiles(fileResults) {
  const months = [];
  const failedFiles = [];
  fileResults.forEach(fr => {
    if (fr.error) {
      failedFiles.push({ fileName: fr.fileName, error: fr.error });
      return;
    }
    months.push(fr.record);
  });
  months.sort((a, b) => String(a.mesKey).localeCompare(String(b.mesKey)));
  return { version: 2, updatedAt: new Date().toISOString(), months, failedFiles };
}

function slimArmPackForPersist(pack) {
  if (!pack) return null;
  return {
    version: pack.version || 2,
    updatedAt: new Date().toISOString(),
    months: (pack.months || []).map(m => ({
      fileName: m.fileName, format: m.format, mesKey: m.mesKey, mesLabel: m.mesLabel,
      dataInicial: m.dataInicial, dataFinal: m.dataFinal,
      totalServicos: m.totalServicos, valorMinimo: m.valorMinimo,
      valorApurado: m.valorApurado, valorTotal: m.valorTotal, belowMin: m.belowMin,
      servicos: m.servicos, nfRows: m.nfRows, adicionais: m.adicionais
    })),
    failedFiles: pack.failedFiles || []
  };
}

async function persistArmPack(pack) {
  if (!pack?.months?.length || typeof upsertExcelBinary !== 'function') return false;
  const slim = slimArmPackForPersist(pack);
  const json = JSON.stringify(slim);
  const bytes = new TextEncoder().encode(json);
  if (bytes.length > ARM_PERSIST_MAX_JSON_BYTES) {
    armToast('Dados demasiado grandes para guardar na cloud — reduz ficheiros ou meses com muitas NFs.', 'error');
    return false;
  }
  const label = `${pack.months.length} mês(es) armazém`;
  try {
    await upsertExcelBinary(armSlot(), label, bytes.buffer);
    return true;
  } catch (e) {
    console.error('[armazem] persist', e);
    armToast('Erro ao guardar na cloud: ' + (e.message || e), 'error');
    return false;
  }
}

function parseArmPackFromRec(rec) {
  if (!rec?.file_data || typeof base64ToArrayBuffer !== 'function') return null;
  try {
    const pack = JSON.parse(new TextDecoder().decode(base64ToArrayBuffer(rec.file_data)));
    if (pack?.months) {
      pack.months.forEach(m => {
        m.nfRows = (m.nfRows || []).map(applySapToArmNf);
      });
    }
    return pack;
  } catch (e) {
    console.error('[armazem] parse pack', e);
    return null;
  }
}

function aggregateByCatalog(months) {
  const map = {};
  months.forEach(m => {
    (m.servicos || []).forEach(s => {
      const cat = resolveCatalogCategory(s.rawName);
      const id = cat.id;
      if (!map[id]) map[id] = { catalogId: id, label: catalogLabel(id), valor: 0, qtde: 0, months: new Set(), servicos: new Set() };
      map[id].valor += s.valor;
      map[id].qtde += s.qtde;
      map[id].months.add(m.mesLabel);
      map[id].servicos.add(s.normName);
    });
  });
  return Object.values(map).map(r => ({
    ...r, months: [...r.months].sort(), servicos: [...r.servicos].sort()
  }));
}

function collectUnknownServicos(months) {
  const seen = new Map();
  months.forEach(m => {
    (m.servicos || []).forEach(s => {
      const cat = resolveCatalogCategory(s.rawName);
      if (!cat.sure && !armCatalogOverrides[s.normName]) {
        if (!seen.has(s.normName)) seen.set(s.normName, { normName: s.normName, rawName: s.rawName, count: 0 });
        seen.get(s.normName).count++;
      }
    });
  });
  return [...seen.values()].sort((a, b) => a.normName.localeCompare(b.normName));
}

function armDoSort(tableId, col) {
  const st = armSorts[tableId];
  if (!st) return;
  if (st.col === col) st.dir *= -1;
  else { st.col = col; st.dir = 1; }
  renderArmActiveTab();
}

function armApplySort(rows, tableId, getters) {
  const st = armSorts[tableId];
  if (!st) return rows;
  const g = getters[st.col];
  if (!g) return rows;
  return [...rows].sort((a, b) => {
    const va = g(a), vb = g(b);
    if (va === vb) return 0;
    if (va === '' || va === null || va === undefined) return 1;
    if (vb === '' || vb === null || vb === undefined) return -1;
    if (typeof va === 'number' && typeof vb === 'number') return st.dir * (va - vb);
    return st.dir * String(va).localeCompare(String(vb), 'pt', { numeric: true });
  });
}

function sortTh(tableId, col, label, cls) {
  const st = armSorts[tableId];
  const active = st?.col === col;
  const ind = active ? (st.dir === 1 ? '▲' : '▼') : '↕';
  const sorted = active ? (st.dir === 1 ? 'sorted-asc' : 'sorted-desc') : '';
  return `<th class="sortable ${cls || ''} ${sorted}" onclick="armDoSort('${tableId}','${col}')">${label}<span class="sort-ind">${ind}</span></th>`;
}

function switchArmTab(tab) {
  armActiveTab = tab;
  document.querySelectorAll('#page-armazem .arm-tab').forEach(t => {
    t.classList.toggle('active', t.dataset.tab === tab);
  });
  ['carregamento', 'mensal', 'resumo', 'nf', 'adicionais', 'catalogo'].forEach(id => {
    const el = $arm('tab-' + id);
    if (el) el.style.display = id === tab ? 'block' : 'none';
  });
  renderArmActiveTab();
}

function renderArmActiveTab() {
  if (armActiveTab === 'mensal') renderArmMensal();
  if (armActiveTab === 'resumo') renderArmAcumulado();
  if (armActiveTab === 'nf') renderArmNf();
  if (armActiveTab === 'adicionais') renderArmAdicionais();
  if (armActiveTab === 'catalogo') renderArmCatalogo();
  if (typeof scheduleColResize === 'function') scheduleColResize();
}

function renderArmDfbGate() {
  const gate = $arm('dfbGate');
  if (!gate) return;
  gate.style.display = armCompany() === 'DFB' ? 'none' : 'block';
}

function armIsExcelFile(file) {
  return /\.xlsx?$/i.test(String(file?.name || ''));
}

function armFilterExcelFiles(fileList) {
  return Array.from(fileList || []).filter(armIsExcelFile);
}

function armOnFileSelected(input) {
  if (!input?.files?.length) return;
  const files = armFilterExcelFiles(input.files);
  if (!files.length) {
    armToast('Seleciona ficheiros .xls ou .xlsx.', 'error');
    input.value = '';
    return;
  }
  armPendingFiles = files;
  updateArmFileZone();
  armToast(`${files.length} ficheiro(s) selecionado(s).`, 'success');
  input.value = '';
}

function updateArmFileZone() {
  const fn = $arm('fileFn');
  const list = $arm('fileList');
  const zone = $arm('fileZone');
  const n = armPack?.months?.length || 0;
  const pend = armPendingFiles.length;
  if (fn) fn.textContent = n ? `✓ ${n} mês(es) carregado(s)` : (pend ? `${pend} ficheiro(s) selecionado(s)` : '');
  zone?.classList.toggle('loaded', n > 0 || pend > 0);
  if (list) {
    const names = [
      ...(armPack?.months || []).map(m => m.fileName),
      ...armPendingFiles.map(f => f.name + ' (pendente)')
    ];
    if (names.length) {
      list.style.display = 'block';
      list.innerHTML = names.map(n => `<div>${armEsc(n)}</div>`).join('');
    } else {
      list.style.display = 'none';
      list.innerHTML = '';
    }
  }
  const sapNote = $arm('sapNote');
  const api = sapApi();
  if (sapNote) {
    const nSap = api?.isSapLoaded() ? Object.keys(api.getMap()).length : 0;
    sapNote.innerHTML = nSap
      ? `✓ SAP NF carregado no módulo <strong>Fretes</strong> (${nSap.toLocaleString('pt-BR')} NFs) — reutilizado para validação 5,5%.`
      : '⚠ Carrega o Excel SAP NF no módulo <strong>Fretes CT-e</strong> para validar valores das NFs (export cobre todos os meses).';
    sapNote.style.display = 'block';
  }
  const btn = $arm('procBtn');
  if (btn) btn.disabled = !armPendingFiles.length;
}

function renderArmMensal() {
  const empty = $arm('mensalEmpty');
  const content = $arm('mensalContent');
  const months = armPack?.months || [];
  if (!months.length) {
    if (empty) empty.style.display = 'block';
    if (content) content.style.display = 'none';
    return;
  }
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'block';

  const below = months.filter(m => m.belowMin).length;
  const total = months.reduce((s, m) => s + m.totalServicos, 0);
  const kpis = $arm('mensalKpis');
  if (kpis) {
    kpis.innerHTML = `
      <div class="kpi"><div class="label">Meses</div><div class="value">${months.length}</div></div>
      <div class="kpi"><div class="label">Total serviços</div><div class="value">${armFmtMoney(total)}</div></div>
      <div class="kpi${below ? ' flag' : ''}"><div class="label">Meses &lt; mínimo</div><div class="value">${below}</div><div class="sub">Mín. ${armFmtMoney(ARM_MINIMO_CONTRATUAL)}</div></div>
      <div class="kpi"><div class="label">NFs 5,5%</div><div class="value">${months.reduce((s, m) => s + (m.nfRows?.length || 0), 0)}</div></div>`;
  }

  const rows = armApplySort(months, 'mensal', {
    mesKey: r => r.mesKey,
    mesLabel: r => r.mesLabel,
    totalServicos: r => r.totalServicos,
    valorMinimo: r => r.valorMinimo,
    belowMin: r => r.belowMin ? 1 : 0,
    nfCount: r => r.nfRows?.length || 0,
    adicionais: r => (r.adicionais || []).reduce((s, a) => s + a.valor, 0)
  });

  const body = $arm('mensalBody');
  if (body) {
    body.innerHTML = rows.map(m => {
      const addVal = (m.adicionais || []).reduce((s, a) => s + a.valor, 0);
      const cls = m.belowMin ? 'month-row-high' : '';
      return `<tr class="${cls}">
        <td>${armEsc(m.mesLabel)}</td>
        <td class="right">${armFmtMoney(m.totalServicos)}</td>
        <td class="right">${armFmtMoney(m.valorMinimo)}</td>
        <td>${m.belowMin ? '<span class="badge b-flag">Abaixo mínimo</span>' : '<span class="badge b-ok">OK</span>'}</td>
        <td class="right">${m.nfRows?.length || 0}</td>
        <td class="right">${armFmtMoney(addVal)}</td>
        <td style="font-size:10px;color:var(--muted)">${armEsc(m.fileName)}</td>
      </tr>`;
    }).join('');
  }

  const thead = $arm('mensalHead');
  if (thead) {
    thead.innerHTML = `
      ${sortTh('mensal', 'mesLabel', 'Mês')}
      ${sortTh('mensal', 'totalServicos', 'Total serviços', 'right')}
      ${sortTh('mensal', 'valorMinimo', 'Mínimo contratual', 'right')}
      ${sortTh('mensal', 'belowMin', 'Alerta')}
      ${sortTh('mensal', 'nfCount', 'NFs 5,5%', 'right')}
      ${sortTh('mensal', 'adicionais', 'Adicionais', 'right')}
      <th>Ficheiro</th>`;
  }
}

function renderArmAcumulado() {
  const empty = $arm('acumuladoEmpty');
  const content = $arm('acumuladoContent');
  const sections = $arm('acumuladoSections');
  const months = armPack?.months || [];
  if (!months.length) {
    if (empty) empty.style.display = 'block';
    if (content) content.style.display = 'none';
    return;
  }
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'block';

  const totalServ = months.reduce((s, m) => s + (m.totalServicos || 0), 0);
  const agg = aggregateResumoByNorm(months);
  const totalAgg = agg.reduce((s, r) => s + r.valor, 0);

  const kpis = $arm('acumuladoKpis');
  if (kpis) {
    kpis.innerHTML = `
      <div class="kpi"><div class="label">Meses</div><div class="value">${months.length}</div></div>
      <div class="kpi"><div class="label">Tipos de serviço (total)</div><div class="value">${agg.length}</div></div>
      <div class="kpi"><div class="label">Soma resumo</div><div class="value">${armFmtMoney(totalAgg)}</div></div>
      <div class="kpi"><div class="label">Total serviços (faturas)</div><div class="value">${armFmtMoney(totalServ)}</div></div>`;
  }

  if (sections) {
    const monthBlocks = months.map(m => {
      const rows = armResumoRows(m);
      const titulo = armMesNomeLong(m.mesKey, m.mesLabel);
      return `<div class="arm-resumo-month">
        <div class="section-title arm-resumo-month-title">${armEsc(titulo)}</div>
        ${resumoTableHtml(rows)}
      </div>`;
    }).join('');

    const totalBlock = `<div class="arm-resumo-month arm-resumo-total-block">
      <div class="section-title arm-resumo-month-title">Total</div>
      ${resumoTableHtml(agg, { useNorm: true, total: true })}
    </div>`;

    sections.innerHTML = monthBlocks + totalBlock;
  }
}

function renderArmNf() {
  const empty = $arm('nfEmpty');
  const content = $arm('nfContent');
  const months = armPack?.months || [];
  const allNf = months.flatMap(m => (m.nfRows || []).map(r => ({ ...r, mesLabel: m.mesLabel })));
  if (!allNf.length) {
    if (empty) empty.style.display = 'block';
    if (content) content.style.display = 'none';
    return;
  }
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'block';

  const filterMes = $arm('nfFilterMes')?.value || '';
  const filterIssue = $arm('nfFilterIssue')?.value || '';
  const search = ($arm('nfSearch')?.value || '').toLowerCase();

  let rows = allNf;
  if (filterMes) rows = rows.filter(r => r.mesKey === filterMes);
  if (filterIssue === 'sap-missing') rows = rows.filter(r => r.sapMissing);
  if (filterIssue === 'valor-diff') rows = rows.filter(r => r.valorDiff);
  if (filterIssue === 'fee-diff') rows = rows.filter(r => Math.abs(r.feeDelta) > 0.05);
  if (search) rows = rows.filter(r => String(r.nf).includes(search) || String(r.sapCliente || '').toLowerCase().includes(search));

  rows = armApplySort(rows, 'nf', {
    mesLabel: r => r.mesLabel,
    nf: r => r.nf,
    valorNF: r => r.valorNF,
    sapValor: r => r.sapValor ?? -1,
    fee: r => r.fee,
    feeDelta: r => r.feeDelta
  });

  const api = sapApi();
  const sapLoaded = api?.isSapLoaded();
  const missing = allNf.filter(r => r.sapMissing).length;
  const valDiff = allNf.filter(r => r.valorDiff).length;

  const kpis = $arm('nfKpis');
  if (kpis) {
    kpis.innerHTML = `
      <div class="kpi"><div class="label">Linhas NF</div><div class="value">${allNf.length}</div></div>
      <div class="kpi"><div class="label">SAP</div><div class="value">${sapLoaded ? 'Carregado' : '—'}</div></div>
      <div class="kpi${missing ? ' flag' : ''}"><div class="label">NF ausente SAP</div><div class="value">${missing}</div></div>
      <div class="kpi${valDiff ? ' flag' : ''}"><div class="label">Δ valor SAP</div><div class="value">${valDiff}</div></div>`;
  }

  const mesSel = $arm('nfFilterMes');
  if (mesSel && !mesSel._armBound) {
    mesSel._armBound = true;
    const opts = ['<option value="">Todos os meses</option>']
      .concat(months.map(m => `<option value="${armEsc(m.mesKey)}">${armEsc(m.mesLabel)}</option>`));
    mesSel.innerHTML = opts.join('');
    mesSel.addEventListener('change', renderArmNf);
    $arm('nfFilterIssue')?.addEventListener('change', renderArmNf);
    $arm('nfSearch')?.addEventListener('input', renderArmNf);
  }

  const thead = $arm('nfHead');
  if (thead) {
    thead.innerHTML = `
      ${sortTh('nf', 'mesLabel', 'Mês')}
      ${sortTh('nf', 'nf', 'NF')}
      <th>Data</th>
      ${sortTh('nf', 'valorNF', 'Valor NF Unilog', 'right')}
      ${sortTh('nf', 'sapValor', 'Valor SAP', 'right')}
      <th>Cliente SAP</th>
      ${sortTh('nf', 'fee', 'Taxa 5,5%', 'right')}
      ${sortTh('nf', 'feeDelta', 'Δ taxa', 'right')}
      <th>Estado</th>`;
  }

  const body = $arm('nfBody');
  if (body) {
    body.innerHTML = rows.map(r => {
      let cls = '';
      if (r.sapMissing) cls = 'nf-sap-missing';
      else if (r.valorDiff) cls = 'nf-val-mismatch';
      const estado = r.sapMissing ? '<span class="badge b-flag">NF ausente SAP</span>'
        : r.valorDiff ? '<span class="badge b-fix165">Δ valor</span>'
        : '<span class="badge b-ok">OK</span>';
      return `<tr class="${cls}">
        <td>${armEsc(r.mesLabel)}</td>
        <td>${armEsc(r.nf)}</td>
        <td>${armEsc(r.data)}</td>
        <td class="right">${armFmtMoney(r.valorNF)}</td>
        <td class="right">${r.sapFound ? armFmtMoney(r.sapValor) : '—'}</td>
        <td>${armEsc(r.sapCliente || '—')}</td>
        <td class="right">${armFmtMoney(r.fee)}</td>
        <td class="right">${Math.abs(r.feeDelta) > 0.02 ? armFmtMoney(r.feeDelta) : '—'}</td>
        <td>${estado}</td>
      </tr>`;
    }).join('');
  }
}

function renderArmAdicionais() {
  const empty = $arm('adicionaisEmpty');
  const content = $arm('adicionaisContent');
  const months = armPack?.months || [];
  const rows = [];
  months.forEach(m => {
    (m.adicionais || []).forEach(a => {
      rows.push({ ...a, mesKey: m.mesKey, mesLabel: m.mesLabel, fileName: m.fileName });
    });
  });
  if (!rows.length) {
    if (empty) empty.style.display = 'block';
    if (content) content.style.display = 'none';
    return;
  }
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'block';

  const filterMes = $arm('adFilterMes')?.value || '';
  let filtered = filterMes ? rows.filter(r => r.mesKey === filterMes) : rows;
  filtered = armApplySort(filtered, 'adicionais', {
    mesLabel: r => r.mesLabel,
    normName: r => r.normName,
    valor: r => r.valor,
    qtde: r => r.qtde
  });

  const mesSel = $arm('adFilterMes');
  if (mesSel && !mesSel._armBound) {
    mesSel._armBound = true;
    mesSel.innerHTML = '<option value="">Todos os meses</option>' +
      months.map(m => `<option value="${armEsc(m.mesKey)}">${armEsc(m.mesLabel)}</option>`).join('');
    mesSel.addEventListener('change', renderArmAdicionais);
  }

  const thead = $arm('adicionaisHead');
  if (thead) {
    thead.innerHTML = `
      ${sortTh('adicionais', 'mesLabel', 'Mês')}
      ${sortTh('adicionais', 'normName', 'Serviço')}
      <th>Catálogo</th>
      ${sortTh('adicionais', 'qtde', 'Qtde', 'right')}
      ${sortTh('adicionais', 'valor', 'Valor', 'right')}`;
  }

  const body = $arm('adicionaisBody');
  if (body) {
    body.innerHTML = filtered.map(r => {
      const cat = resolveCatalogCategory(r.rawName);
      return `<tr>
        <td>${armEsc(r.mesLabel)}</td>
        <td>${armEsc(r.normName)}</td>
        <td>${armEsc(catalogLabel(cat.id))}${cat.sure ? '' : ' <span class="badge b-min">?</span>'}</td>
        <td class="right">${fmtArmQty(r.qtde)}</td>
        <td class="right">${armFmtMoney(r.valor)}</td>
      </tr>`;
    }).join('');
  }
}

function renderArmCatalogo() {
  const months = armPack?.months || [];
  const unknown = collectUnknownServicos(months);
  const allNames = new Map();
  months.forEach(m => (m.servicos || []).forEach(s => {
    if (!allNames.has(s.normName)) allNames.set(s.normName, s.rawName);
  }));

  const rows = armApplySort([...allNames.entries()].map(([normName, rawName]) => ({
    normName, rawName, cat: resolveCatalogCategory(rawName)
  })), 'catalogo', {
    servico: r => r.normName,
    catalog: r => r.cat.id
  });

  const warn = $arm('catalogWarn');
  if (warn) {
    warn.style.display = unknown.length ? 'block' : 'none';
    warn.innerHTML = unknown.length
      ? `<strong>${unknown.length} serviço(s) por validar</strong> — confirma o mapeamento para o catálogo abaixo.`
      : '';
  }

  const body = $arm('catalogBody');
  if (body) {
    body.innerHTML = rows.map(r => {
      const opts = CATALOGO_DESPESAS.map(c =>
        `<option value="${c.id}" ${(armCatalogOverrides[r.normName] || r.cat.id) === c.id ? 'selected' : ''}>${armEsc(c.label)}</option>`
      ).join('');
      return `<tr class="${r.cat.sure ? '' : 'qz-diff-row'}">
        <td>${armEsc(r.normName)}</td>
        <td style="font-size:10px;color:var(--muted)">${armEsc(r.rawName)}</td>
        <td><select class="inp" data-norm="${armEsc(r.normName)}" onchange="armSetCatalog(this)">${opts}</select></td>
        <td>${r.cat.sure ? '<span class="badge b-ok">Auto</span>' : '<span class="badge b-flag">Validar</span>'}</td>
      </tr>`;
    }).join('');
  }

  const thead = $arm('catalogHead');
  if (thead) {
    thead.innerHTML = `
      ${sortTh('catalogo', 'servico', 'Serviço (normalizado)')}
      <th>Nome original</th>
      ${sortTh('catalogo', 'catalog', 'Catálogo')}
      <th>Estado</th>`;
  }
}

function armSetCatalog(sel) {
  const norm = sel.dataset.norm;
  armCatalogOverrides[norm] = sel.value;
  saveCatalogOverrides();
  if (armPack?.months?.length) {
    armPack.months.forEach(m => {
      (m.servicos || []).forEach(s => {
        if (s.normName === norm) {
          const cat = resolveCatalogCategory(s.rawName);
          s.catalogId = cat.id;
          s.catalogSure = cat.sure;
        }
      });
      m.adicionais = (m.servicos || []).filter(s => !s.isNf && s.catalogId !== 'impostos');
    });
    renderArmActiveTab();
  }
  armToast('Catálogo atualizado.');
}

function armSetProcessing(active, label) {
  const spin = $arm('procSpin');
  const btn = $arm('procBtn');
  const note = $arm('procNote');
  if (spin) {
    spin.hidden = !active;
    spin.style.display = active ? 'inline-flex' : 'none';
  }
  if (btn) btn.disabled = active || !armPendingFiles.length;
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
}

async function processOneArmFile(file) {
  const size = file.size || 0;
  if (size > ARM_MAX_FILE_BYTES) {
    throw new Error(
      `Ficheiro demasiado grande (${armFmtBytes(size)}). Limite ${armFmtBytes(ARM_MAX_FILE_BYTES)} — ` +
      'ex.: Novembro 2025 (~24 MB) trava o browser; usa meses mais recentes em .xlsx ou remove a folha Arm.'
    );
  }
  const { wb } = await armWithTimeout(readFileToWorkbook(file), ARM_PARSE_TIMEOUT_MS, file.name);
  await armYield();
  const record = await armWithTimeout(
    Promise.resolve(parseArmazemWorkbook(wb, file.name)),
    ARM_PARSE_TIMEOUT_MS,
    file.name
  );
  return { fileName: file.name, record };
}

async function processArmFiles() {
  if (armCompany() !== 'DFB') {
    armToast('Avaliação Armazém disponível apenas para DFB.', 'error');
    return;
  }
  if (!armPendingFiles.length) {
    armToast('Seleciona ficheiros Excel para processar.', 'error');
    return;
  }
  const pending = [...armPendingFiles];
  armSetProcessing(true, `A processar 0/${pending.length}…`);
  try {
    const results = [];
    for (let i = 0; i < pending.length; i++) {
      const file = pending[i];
      armSetProcessing(true, `A processar ${i + 1}/${pending.length}: ${file.name}`);
      try {
        results.push(await processOneArmFile(file));
      } catch (err) {
        const msg = String(err.message || err);
        console.error('[armazem] parse', file.name, err);
        results.push({ fileName: file.name, error: msg });
        armToast(`${file.name}: ${msg}`, 'error');
      }
      await armYield();
    }
    const built = buildArmPackFromFiles(results);
    const okN = built.months.length;
    const failN = built.failedFiles?.length || 0;
    if (!okN && failN) {
      armToast(`Nenhum mês processado (${failN} erro(s)).`, 'error');
      return;
    }
    if (!okN) {
      armToast('Nenhum dado extraído dos ficheiros.', 'error');
      return;
    }
    armPack = mergeArmPack(armPack, built);
    refreshAllSapOnNfs();
    armPendingFiles = [];
    updateArmFileZone();
    const saved = await persistArmPack(armPack);
    const totalMonths = armPack.months.length;
    armToast(
      `Processado: ${okN} novo(s), ${totalMonths} mês(es) no total` +
      (failN ? `, ${failN} erro(s)` : '') +
      (saved ? ' · guardado' : '') + '.',
      failN ? 'error' : 'success'
    );
    switchArmTab(totalMonths === 1 ? 'mensal' : 'resumo');
  } catch (err) {
    console.error('[armazem] processArmFiles', err);
    armToast('Erro ao processar: ' + (err.message || err), 'error');
  } finally {
    armSetProcessing(false);
    updateArmFileZone();
  }
}

async function loadSavedArmazem(silent) {
  if (typeof fetchExcelFiles !== 'function') return false;
  try {
    const m = await Promise.race([
      fetchExcelFiles([armSlot()]),
      new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 25_000))
    ]);
    const rec = m[armSlot()];
    const pack = parseArmPackFromRec(rec);
    if (pack?.months?.length) {
      armPack = pack;
      refreshAllSapOnNfs();
      updateArmFileZone();
      if (!silent) armToast(`${pack.months.length} mês(es) carregado(s) da cloud.`);
      return true;
    }
  } catch (e) {
    console.warn('[armazem] load saved', e);
  }
  return false;
}

function exportResumoSheetRows(rows, title) {
  const aoa = [
    [title || 'Resumo'],
    [],
    ['Tipo de Serviço', 'Qtde Serviço / Valor Faturação', 'Valor Unitário', 'Valor Calculado']
  ];
  rows.forEach(s => {
    aoa.push([
      s.rawName || s.normName,
      s.isNf ? s.qtde : s.qtde,
      s.isNf ? ARM_NF_RATE : s.valorUnit,
      s.valor
    ]);
  });
  const subtotal = rows.reduce((s, r) => s + (r.valor || 0), 0);
  aoa.push([], ['Subtotal', '', '', subtotal]);
  return aoa;
}

function exportArmazemWorkbook() {
  const months = armPack?.months || [];
  if (!months.length) { armToast('Sem dados para exportar.', 'error'); return; }
  const wb = XLSX.utils.book_new();
  const usedNames = new Set();

  months.forEach(m => {
    const rows = armResumoRows(m);
    let name = armMesSheetName(m.mesKey);
    let n = 2;
    while (usedNames.has(name)) { name = armMesSheetName(m.mesKey).slice(0, 28) + '_' + n++; }
    usedNames.add(name);
    const aoa = exportResumoSheetRows(rows, 'Resumo — ' + armMesNomeLong(m.mesKey, m.mesLabel));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
  });

  const agg = aggregateResumoByNorm(months);
  const totalAoa = exportResumoSheetRows(
    agg.map(r => ({ ...r, rawName: r.normName })),
    'Total — todos os meses'
  );
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(totalAoa), 'Total');

  const mensal = months.map(m => ({
    Mes: m.mesLabel, Total: m.totalServicos, Minimo: m.valorMinimo,
    AbaixoMinimo: m.belowMin ? 'SIM' : 'NÃO',
    NFs: m.nfRows?.length || 0,
    Adicionais: (m.adicionais || []).reduce((s, a) => s + a.valor, 0),
    Ficheiro: m.fileName
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mensal), 'Visão mensal');

  const nfs = months.flatMap(m => (m.nfRows || []).map(r => ({
    Mes: m.mesLabel, NF: r.nf, Data: r.data,
    ValorNF_Unilog: r.valorNF, ValorNF_SAP: r.sapValor ?? '',
    Taxa: r.fee, DeltaTaxa: r.feeDelta,
    SAP_OK: r.sapFound ? 'SIM' : (r.sapMissing ? 'AUSENTE' : '')
  })));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(nfs), 'NF 5.5%');

  const ad = months.flatMap(m => (m.adicionais || []).map(a => ({
    Mes: m.mesLabel, Servico: a.normName, Valor: a.valor, Qtde: a.qtde
  })));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ad), 'Adicionais');

  XLSX.writeFile(wb, `armazem_delta_${armCompany()}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  armToast('Excel exportado.');
}

async function reloadArmazemForCompany() {
  armPack = null;
  armPendingFiles = [];
  loadCatalogOverrides();
  await loadSavedArmazem(true);
  updateArmFileZone();
  renderArmDfbGate();
  if (armInited) renderArmActiveTab();
}

function initArmazem() {
  armSetProcessing(false);
  const initErr = $arm('initErr');
  if (initErr) initErr.style.display = 'none';

  if (armInited) {
    renderArmDfbGate();
    updateArmFileZone();
    refreshAllSapOnNfs();
    return;
  }
  armInited = true;
  loadCatalogOverrides();

  document.querySelectorAll('#page-armazem .arm-tab').forEach(t => {
    t.addEventListener('click', () => switchArmTab(t.dataset.tab));
  });

  const zone = $arm('fileZone');
  const input = $arm('fileInput');
  if (zone && input) {
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
    zone.addEventListener('drop', e => {
      e.preventDefault(); zone.classList.remove('drag');
      if (e.dataTransfer.files.length) {
        const files = armFilterExcelFiles(e.dataTransfer.files);
        if (!files.length) {
          armToast('Arrasta ficheiros .xls ou .xlsx.', 'error');
          return;
        }
        armPendingFiles = files;
        updateArmFileZone();
        armToast(`${files.length} ficheiro(s) selecionado(s).`, 'success');
      }
    });
    if (!input._armChangeBound) {
      input._armChangeBound = true;
      input.addEventListener('change', e => armOnFileSelected(e.target));
    }
  }

  $arm('procBtn')?.addEventListener('click', () => processArmFiles());
  $arm('loadBtn')?.addEventListener('click', async () => {
    const ok = await loadSavedArmazem(false);
    if (ok) switchArmTab((armPack?.months?.length || 0) === 1 ? 'mensal' : 'resumo');
  });
  $arm('exportBtn')?.addEventListener('click', () => exportArmazemWorkbook());

  renderArmDfbGate();
  updateArmFileZone();
  loadSavedArmazem(true).catch(e => console.warn('[armazem] load saved', e));
}

window.initArmazem = initArmazem;
window.armOnFileSelected = armOnFileSelected;
window.processArmFiles = processArmFiles;
window.armDoSort = armDoSort;
window.armSetCatalog = armSetCatalog;
window.exportArmazemWorkbook = exportArmazemWorkbook;
