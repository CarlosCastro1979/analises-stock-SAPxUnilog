// armazem.js v1.0.34
const ARMAZEM_JS_VERSION = '1.0.34';

const ARM_MINIMO_CONTRATUAL = 120000;
const ARM_NF_RATE = 0.055;
const ARM_IMPOSTOS_RATE = 0.1125;
const ARM_NET_FACTOR = 1 - ARM_IMPOSTOS_RATE; // 0,8875 — fator líquido Unilog
const ARM_IMPOSTOS_LABEL = 'Impostos — 2% (ISS) + 9,25% (PIS/COFINS) = 11,25% (gross-up s/ subtotal)';
const ARM_PERSIST_MAX_JSON_BYTES = 6 * 1024 * 1024;
const ARM_MAX_FILE_BYTES = 15 * 1024 * 1024;
const ARM_ABSOLUTE_MAX_FILE_BYTES = 40 * 1024 * 1024;
const ARM_PARSE_TIMEOUT_MS = 90_000;
const ARM_LARGE_PARSE_TIMEOUT_MS = 180_000;

const $arm = id => document.getElementById('arm-' + id);

let armInited = false;
let armPack = null;
let armPendingFiles = [];
let armActiveTab = 'lancamento';
let armVendasCache = null;
let armVendasCacheTs = 0;
let armLancamentoDraft = null;
let armLancamentoSummaryTimer = null;
let armNfUploadRows = [];
const ARM_VENDAS_CACHE_MS = 60_000;
let armCatalogOverrides = {};
let armSorts = {
  mensal: { col: 'mesKey', dir: 1 },
  acumulado: { col: 'valor', dir: -1 },
  nf: { col: 'nf', dir: 1 },
  nfUpload: { col: 'nf', dir: 1 },
  adicionais: { col: 'valor', dir: -1 },
  catalogo: { col: 'servico', dir: 1 }
};

/** Nome canónico — plain, [AG], [AT], [AREA TECNICA] e [AG/AT] são o mesmo serviço. */
const ARM_NORM_PALLET_POS = 'ARMAZENAGEM POR POSICAO PALLET (AREA TECNICA)';

/** Nome canónico — [INSUMOS], (INSUMOS) e variantes com INSUMOS no nome. */
const ARM_NORM_PALLET_INSUMOS = 'ARMAZENAGEM POR POSICAO PALLET (INSUMOS)';

/** Nome canónico — H.E, 22h–07h, domingo/feriado e variantes de acento são o mesmo serviço. */
const ARM_NORM_HORA_EXTRA = 'HORA EXTRA';

/** Nome canónico — [8], [8.2], acento e plain são o mesmo serviço. */
const ARM_NORM_READEQUACAO = 'READEQUACAO DE PRODUTOS - POR UNIDADE';

const ARM_SERVICO_DISPLAY_LABELS = {
  [ARM_NORM_PALLET_POS]: 'Armazenagem por posição pallet (Área técnica)',
  [ARM_NORM_PALLET_INSUMOS]: 'Armazenagem por posição pallet (Insumos)',
  [ARM_NORM_HORA_EXTRA]: 'Hora extra',
  [ARM_NORM_READEQUACAO]: 'Readequação de produtos - Por unidade'
};

/** Chaves legadas de overrides de catálogo → normName actual. */
const ARM_NORM_ALIASES = {
  'ARMAZENAGEM POR POSICAO PALLET [AG/AT]': ARM_NORM_PALLET_POS,
  'ARMAZENAGEM POR POSICAO PALLET': ARM_NORM_PALLET_POS,
  'ARMAZENAGEM POR POSICAO PALLET [INSUMOS]': ARM_NORM_PALLET_INSUMOS,
  'ARMAZENAGEM POR POSICAO PALLET (INSUMOS)': ARM_NORM_PALLET_INSUMOS,
  'HORA EXTRA - DOMINGO E FERIADO - HE': ARM_NORM_HORA_EXTRA,
  'HORA EXTRA - DOMINGO E FERIADO': ARM_NORM_HORA_EXTRA,
  'H.E - DOMINGO E FERIADO': ARM_NORM_HORA_EXTRA,
  'H.E 2A A 6A DE 22HS AS 07HS E AOS SABADOS': ARM_NORM_HORA_EXTRA,
  '2A A 6A DE 22HS AS 07HS E AOS SABADOS - H.E': ARM_NORM_HORA_EXTRA,
  'READEQUACAO DE PRODUTOS - POR UNIDADE': ARM_NORM_READEQUACAO,
  '[8] READEQUACAO DE PRODUTOS - POR UNIDADE': ARM_NORM_READEQUACAO,
  '[8.2] READEQUACAO DE PRODUTOS - POR UNIDADE': ARM_NORM_READEQUACAO
};

/** Catálogo essencial do dropdown (variantes obscuras via «+ novo serviço»). */
const ARM_DEFAULT_SERVICOS = [
  'PERCENTUAL SOBRE NF EXPEDIDA',
  ARM_NORM_PALLET_POS,
  ARM_NORM_PALLET_INSUMOS,
  'ARMAZENAGEM EXCEDENTE',
  ARM_NORM_HORA_EXTRA,
  'ETIQUETAGEM POR UNIDADE',
  ARM_NORM_READEQUACAO,
  'DESCARGA DE VEICULO',
  'CONFERENCIA DE MERCADORIA'
];

const CATALOGO_DESPESAS = [
  { id: 'nf_percentual', label: '5,5% sobre NF expedida' },
  { id: 'armazenagem', label: 'Armazenagem' },
  { id: 'armazenagem_insumos', label: 'Armazenagem (Insumos)' },
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

function armNormServicoUpper(raw) {
  return String(raw || '').trim()
    .replace(/^\[[\d.]+\]\s*/i, '')
    .replace(/:$/, '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

/** H.E / hora extra — mesma tarifa; variantes de acento, 22h–07h e domingo/feriado. */
function isHoraExtraVariant(u) {
  if (/HORA\s+EXTRA/.test(u)) return true;
  if (/(^|[\s\-])H\.?\s*E([\s\-]|$)/.test(u)) return true;
  if (/\bHE\b/.test(u) && /DOMINGO|FERIADO|22\s*HS|22HS|HORA|SABADOS?/.test(u)) return true;
  if (/22\s*HS|22HS/.test(u) && /(07\s*HS|07HS|SABADOS?|\b6\s*A\b)/.test(u)) return true;
  if (/DOMINGO\s+E\s+FERIADO/.test(u)) return true;
  return false;
}

function normalizeServicoName(raw) {
  const u = armNormServicoUpper(raw);
  if (!u) return '';

  if (/PERCENTUAL\s+SOBRE\s+NF\s+EXPEDID|ARMAZENAGEM\s+POR\s*%|5[,.]5\s*%/.test(u)) {
    return 'PERCENTUAL SOBRE NF EXPEDIDA';
  }
  if (isHoraExtraVariant(u)) {
    return ARM_NORM_HORA_EXTRA;
  }
  if (/ARMAZENAGEM\s+EXCEDENTE/.test(u) && /(\bNA\s+AT\b|AREA\s+TECNICA)/.test(u)) {
    return 'ARMAZENAGEM EXCEDENTE NA AT';
  }
  if (/^ARMAZENAGEM\s+EXCEDENTE/.test(u)) {
    return 'ARMAZENAGEM EXCEDENTE';
  }
  if (/^ARMAZENAGEM\s+POR\s+POSICAO\s+PALLET/.test(u)) {
    if (/\bINSUMOS\b/.test(u)) return ARM_NORM_PALLET_INSUMOS;
    return ARM_NORM_PALLET_POS;
  }
  if (/READEQUA/.test(u)) return ARM_NORM_READEQUACAO;
  if (/IMPOSTO|PIS|COFINS|ISS/.test(u)) return 'IMPOSTOS SERVICOS';

  // ETIQUETAGEM, DECARGA and other lines: keep distinct names (no cross-type merging).
  return u;
}

function armServicoDisplayLabel(normOrRaw) {
  const norm = normalizeServicoName(normOrRaw);
  return ARM_SERVICO_DISPLAY_LABELS[norm] || norm;
}

function armCatalogOverrideKey(norm) {
  return ARM_NORM_ALIASES[norm] || norm;
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

function armFmtResumoUnit(v, isNf, unitVaries) {
  if (isNf) return fmtArmPct(ARM_NF_RATE * 100);
  if (unitVaries || v == null) return '—';
  return fmtArmUnit(v);
}

function armResumoRowHtml(s, opts) {
  const isNf = !!s.isNf;
  const nome = armServicoDisplayLabel(s.normName || s.rawName);
  return `<tr>
    <td>${armEsc(nome)}</td>
    <td class="right">${armFmtResumoQty(s.qtde, isNf)}</td>
    <td class="right">${armFmtResumoUnit(s.valorUnit, isNf, s.unitVaries)}</td>
    <td class="right">${armFmtMoney(s.valor)}</td>
  </tr>`;
}

function aggregateResumoByNorm(months) {
  const map = {};
  months.forEach(m => {
    armResumoRows(m).forEach(s => {
      const norm = normalizeServicoName(s.rawName || s.normName);
      if (!norm) return;
      if (!map[norm]) {
        map[norm] = {
          normName: norm, qtde: 0, valor: 0, unitWeighted: 0, unitQty: 0,
          unitValues: new Set(), isNf: false, rawVariants: new Set()
        };
      }
      const row = map[norm];
      const q = s.qtde || 0;
      const u = s.valorUnit || 0;
      row.qtde += q;
      row.valor += s.valor || 0;
      if (q && u) {
        row.unitWeighted += q * u;
        row.unitQty += q;
        row.unitValues.add(Math.round(u * 10000) / 10000);
      }
      if (s.isNf) row.isNf = true;
      if (s.rawName) row.rawVariants.add(s.rawName);
    });
  });
  return Object.values(map).map(r => {
    const unitVaries = !r.isNf && r.unitValues.size > 1;
    let valorUnit = null;
    if (r.isNf) valorUnit = ARM_NF_RATE;
    else if (!unitVaries && r.unitQty) valorUnit = r.unitWeighted / r.unitQty;
    return {
      normName: r.normName,
      rawName: r.normName,
      qtde: r.qtde,
      valor: r.valor,
      valorUnit,
      unitVaries,
      isNf: r.isNf,
      rawVariants: [...r.rawVariants]
    };
  }).sort((a, b) => a.normName.localeCompare(b.normName, 'pt'));
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
  if (n === ARM_NORM_PALLET_INSUMOS) return { id: 'armazenagem_insumos', sure: true };
  if (/ARMAZENAGEM/.test(n)) return { id: 'armazenagem', sure: true };
  if (n === ARM_NORM_HORA_EXTRA || isHoraExtraVariant(n)) return { id: 'hora_extra', sure: true };
  if (/ETIQUETAGEM/.test(n)) return { id: 'etiquetagem', sure: true };
  if (n === ARM_NORM_READEQUACAO) return { id: 'readequacao', sure: true };
  if (/IMPOSTO/.test(n)) return { id: 'impostos', sure: true };
  return { id: 'outros', sure: false };
}

function resolveCatalogCategory(rawName) {
  const norm = normalizeServicoName(rawName);
  const overrideKey = armCatalogOverrideKey(norm);
  if (armCatalogOverrides[overrideKey] || armCatalogOverrides[norm]) {
    return { id: armCatalogOverrides[overrideKey] || armCatalogOverrides[norm], sure: true, norm, raw: rawName };
  }
  const g = guessCatalogCategory(norm);
  return { id: g.id, sure: g.sure, norm, raw: rawName };
}

function catalogLabel(id) {
  return CATALOGO_DESPESAS.find(c => c.id === id)?.label || id;
}

function armEnsurePack() {
  if (!armPack) armPack = { version: 3, updatedAt: new Date().toISOString(), months: [], customServices: [] };
  if (!armPack.customServices) armPack.customServices = [];
  if (!armPack.months) armPack.months = [];
  return armPack;
}

function armCollectServicosFromPack() {
  const names = new Set();
  (armPack?.months || []).forEach(m => (m.servicos || []).forEach(s => {
    const n = s.normName || normalizeServicoName(s.rawName);
    if (n) names.add(n);
  }));
  return names;
}

function armGetServiceCatalog() {
  const set = new Set(ARM_DEFAULT_SERVICOS);
  (armPack?.customServices || []).forEach(n => {
    const norm = normalizeServicoName(n);
    if (norm) set.add(norm);
  });
  return [...set].sort((a, b) => a.localeCompare(b, 'pt'));
}

function armServiceSubtotal(servicos) {
  return (servicos || []).filter(s => !/^IMPOSTO/i.test(s.normName || '')).reduce((a, s) => a + (Number(s.valor) || 0), 0);
}

/** Unilog Excel: =E11/0,8875-E11 (gross-up, não 11,25% simples do subtotal). */
function armCalcImpostos(subtotal) {
  const sub = armNum(subtotal);
  if (sub <= 0) return 0;
  return Math.round((sub / ARM_NET_FACTOR - sub) * 100) / 100;
}

function armGetMonthImpostos(month) {
  if (!month) return 0;
  if (month.impostosManual) return armNum(month.impostos);
  const pago = armGetMensalPagoSemImp(month);
  if (pago > 0) return armCalcImpostos(pago);
  return armCalcImpostos(month.totalServicos || month.valorApurado || 0);
}

function armGetMonthPagoComImp(month) {
  const pago = armGetMensalPagoSemImp(month) || armGetArmazenagemValor(month).pago;
  return pago + armGetMonthImpostos(month);
}

function armLancamentoServicosFromMonth(month) {
  return (month.servicos || [])
    .filter(s => !/^IMPOSTO/i.test(s.normName || ''))
    .map(s => ({ ...s }));
}

function armMonthsWithData() {
  return new Set((armPack?.months || [])
    .filter(m => m.mesKey && (m.servicos?.length || m.nfRows?.length))
    .map(m => m.mesKey));
}

function armBuildServicoFromEntry(rawName, qtde, valorUnit) {
  const norm = normalizeServicoName(rawName);
  const isNf = isNfPercentualService(rawName);
  const nums = armFinishServico(rawName, qtde, valorUnit, 0, isNf);
  const cat = resolveCatalogCategory(rawName);
  return {
    rawName: rawName || norm,
    normName: norm,
    codigo: rawName || norm,
    qtde: nums.qtde,
    valorUnit: nums.valorUnit,
    valor: nums.valor,
    catalogId: cat.id,
    catalogSure: cat.sure,
    isNf
  };
}

function armRecalcMonthTotals(month) {
  const servicos = month.servicos || [];
  const subtotal = armServiceSubtotal(servicos);
  month.totalServicos = subtotal;
  month.valorApurado = subtotal;
  month.impostos = month.impostosManual ? armNum(month.impostos) : armCalcImpostos(subtotal);
  month.valorTotal = subtotal + month.impostos;
  month.belowMin = subtotal < (month.valorMinimo || ARM_MINIMO_CONTRATUAL);
  month.adicionais = servicos.filter(s => !s.isNf && s.catalogId !== 'impostos');
  armRepairNfServico(month);
  return month;
}

function armTypologyGroups(servicos) {
  const groups = {};
  (servicos || []).forEach(s => {
    if (/^IMPOSTO/i.test(s.normName || '')) return;
    const id = s.catalogId || resolveCatalogCategory(s.rawName || s.normName).id;
    if (!groups[id]) groups[id] = { id, label: catalogLabel(id), valor: 0, count: 0 };
    groups[id].valor += Number(s.valor) || 0;
    groups[id].count++;
  });
  return Object.values(groups).sort((a, b) => a.label.localeCompare(b.label, 'pt'));
}

function armTypologySummaryHtml(servicos) {
  const groups = armTypologyGroups(servicos);
  if (!groups.length) {
    return '<p style="font-size:11px;color:var(--muted);margin:0;">Adiciona linhas de serviço para ver o resumo por tipologia.</p>';
  }
  const total = groups.reduce((s, g) => s + g.valor, 0);
  const rows = groups.map(g => `<tr class="arm-typo-row">
    <td>${armEsc(g.label)}</td>
    <td class="right">${g.count}</td>
    <td class="right">${armFmtMoney(g.valor)}</td>
  </tr>`).join('');
  return `<div class="tbl-wrap" style="max-height:280px;border:none;">
    <table><thead><tr>
      <th>Tipologia</th><th class="right">Linhas</th><th class="right">Total</th>
    </tr></thead><tbody>${rows}
      <tr class="arm-typo-total"><td>Total</td><td></td><td class="right">${armFmtMoney(total)}</td></tr>
    </tbody></table></div>`;
}

function armLancamentoMesKey() {
  const y = $arm('lancAno')?.value;
  const m = $arm('lancMes')?.value;
  if (!y || !m) return '';
  return `${y}-${m}`;
}

function armInitLancamentoSelectors() {
  const anoSel = $arm('lancAno');
  const mesSel = $arm('lancMes');
  if (!anoSel || anoSel._armInit) return;
  anoSel._armInit = true;
  if (!anoSel._armChangeBound) {
    anoSel._armChangeBound = true;
    anoSel.addEventListener('change', () => {
      armRefreshLancamentoMonthOptions();
      armLoadLancamentoDraft();
    });
  }
  if (mesSel && !mesSel._armChangeBound) {
    mesSel._armChangeBound = true;
    mesSel.addEventListener('change', () => armLoadLancamentoDraft());
  }
  armRefreshLancamentoMonthOptions();
}

function armRefreshLancamentoMonthOptions() {
  const anoSel = $arm('lancAno');
  const mesSel = $arm('lancMes');
  if (!anoSel || !mesSel) return;
  const dataMonths = armMonthsWithData();
  const now = new Date();
  const curY = now.getFullYear();
  const yearSet = new Set();
  dataMonths.forEach(mk => {
    const y = mk.slice(0, 4);
    if (y) yearSet.add(y);
  });
  for (let y = curY + 1; y >= curY - 3; y--) yearSet.add(String(y));
  const years = [...yearSet].map(Number).sort((a, b) => b - a);
  const prevY = anoSel.value || String(curY);
  anoSel.innerHTML = years.map(y =>
    `<option value="${y}" ${String(y) === prevY ? 'selected' : ''}>${y}</option>`
  ).join('');
  if (!anoSel.value) anoSel.value = String(curY);
  const y = anoSel.value;
  const curM = mesSel.value || String(now.getMonth() + 1).padStart(2, '0');
  mesSel.innerHTML = ARM_MESES_NOME.slice(1).map((nome, i) => {
    const num = String(i + 1).padStart(2, '0');
    const mk = `${y}-${num}`;
    const hasData = dataMonths.has(mk);
    const sel = num === curM ? ' selected' : '';
    const label = hasData ? `${nome} ●` : nome;
    return `<option value="${num}"${sel} title="${hasData ? 'Mês com dados guardados' : ''}">${label}</option>`;
  }).join('');
}

function armBlankLancamentoDraft(mesKey) {
  return {
    mesKey,
    mesLabel: armMesNomeLong(mesKey, ''),
    impostos: 0,
    impostosManual: false,
    entrySource: '',
    servicos: [armBuildServicoFromEntry('PERCENTUAL SOBRE NF EXPEDIDA', 0, ARM_NF_RATE)]
  };
}

function armLoadLancamentoDraft() {
  const mesKey = armLancamentoMesKey();
  if (!mesKey) {
    armToast('Seleciona ano e mês.', 'error');
    return;
  }
  const existing = (armPack?.months || []).find(m => m.mesKey === mesKey);
  if (existing) {
    const servicos = armLancamentoServicosFromMonth(existing);
    const subtotal = armServiceSubtotal(servicos);
    const impostosManual = !!existing.impostosManual;
    armLancamentoDraft = {
      mesKey,
      mesLabel: armMesLabel(existing),
      entrySource: existing.entrySource || (existing.format === 'manual' ? 'manual' : 'excel'),
      impostosManual,
      impostos: impostosManual ? armNum(existing.impostos) : armCalcImpostos(subtotal),
      servicos: servicos.length ? servicos : armBlankLancamentoDraft(mesKey).servicos
    };
    armNfUploadRows = (existing.nfRows || []).map(r => ({ ...r }));
  } else {
    armLancamentoDraft = armBlankLancamentoDraft(mesKey);
    armNfUploadRows = [];
  }
  renderArmLancamentoForm();
  renderArmNfUploadTable();
}

function armServiceSelectOptions(selected) {
  const cat = armGetServiceCatalog();
  const selNorm = normalizeServicoName(selected);
  let html = cat.map(name =>
    `<option value="${armEsc(name)}" ${normalizeServicoName(name) === selNorm ? 'selected' : ''}>${armEsc(armServicoDisplayLabel(name))}</option>`
  ).join('');
  if (selected && selNorm && !cat.some(n => normalizeServicoName(n) === selNorm)) {
    html += `<option value="${armEsc(selected)}" selected>${armEsc(armServicoDisplayLabel(selected))}</option>`;
  }
  return html;
}

function armLancamentoRowHtml(s, idx) {
  const isNf = !!s.isNf;
  const qtyVal = isNf ? fmtArmNum(s.qtde, 2) : fmtArmNum(s.qtde, 2);
  const unitVal = isNf ? fmtArmPct(ARM_NF_RATE * 100) : fmtArmNum(s.valorUnit, 2);
  const unitReadonly = isNf ? 'readonly tabindex="-1"' : '';
  return `<tr class="arm-entry-row" data-idx="${idx}">
    <td><select data-field="servico" onchange="armLancamentoServicoChange(${idx})">${armServiceSelectOptions(s.rawName || s.normName)}</select></td>
    <td class="right"><input type="text" data-field="qtde" value="${armEsc(qtyVal)}" oninput="armLancamentoRowPreview(${idx})" onblur="armLancamentoRowCommit(${idx})" style="text-align:right;"></td>
    <td class="right"><input type="text" data-field="unit" value="${armEsc(unitVal)}" ${unitReadonly} oninput="armLancamentoRowPreview(${idx})" onblur="armLancamentoRowCommit(${idx})" style="text-align:right;"></td>
    <td class="right arm-val-calc">${armFmtMoney(s.valor || 0)}</td>
    <td><button type="button" class="btn bg bs" style="padding:2px 6px;font-size:10px;" onclick="armLancamentoRemoveRow(${idx})" title="Remover">✕</button></td>
  </tr>`;
}

function renderArmLancamentoForm() {
  const draft = armLancamentoDraft;
  const body = $arm('lancEntryBody');
  if (!body || !draft) return;
  body.innerHTML = (draft.servicos || []).map((s, i) => armLancamentoRowHtml(s, i)).join('');
  const imp = $arm('lancImpostos');
  if (imp && document.activeElement !== imp) {
    imp.value = draft.impostos ? fmtArmNum(draft.impostos, 2) : '';
  }
  const impLbl = $arm('lancImpostosLabel');
  if (impLbl) {
    impLbl.textContent = draft.impostosManual
      ? `${ARM_IMPOSTOS_LABEL} (valor manual)`
      : ARM_IMPOSTOS_LABEL;
  }
  armRefreshLancamentoSummary();
  const note = $arm('lancSaveNote');
  if (note) {
    const saved = (armPack?.months || []).find(m => m.mesKey === draft.mesKey);
    if (saved) {
      const src = draft.entrySource === 'manual' ? 'manual' : 'Excel';
      note.textContent = `Dados carregados (${src}) — ${draft.mesLabel}`;
    } else {
      note.textContent = 'Mês ainda não guardado';
    }
  }
  if (typeof scheduleTableSort === 'function') scheduleTableSort();
}

function armRefreshLancamentoSummary() {
  const draft = armLancamentoDraft;
  if (!draft) return;
  const subtotal = armServiceSubtotal(draft.servicos);
  if (!draft.impostosManual) draft.impostos = armCalcImpostos(subtotal);
  const imp = armNum(draft.impostos);
  const typ = $arm('lancTypology');
  if (typ) typ.innerHTML = armTypologySummaryHtml(draft.servicos);
  const subEl = $arm('lancSubtotal');
  const impEl = $arm('lancImpostosDisp');
  const impSumLbl = $arm('lancImpostosSumLabel');
  const impLbl = $arm('lancImpostosLabel');
  const totEl = $arm('lancTotal');
  if (subEl) subEl.textContent = armFmtMoney(subtotal);
  if (impEl) impEl.textContent = armFmtMoney(imp);
  if (impSumLbl) {
    impSumLbl.textContent = draft.impostosManual ? 'Impostos (manual)' : ARM_IMPOSTOS_LABEL;
  }
  if (impLbl) {
    impLbl.textContent = draft.impostosManual
      ? `${ARM_IMPOSTOS_LABEL} (valor manual)`
      : ARM_IMPOSTOS_LABEL;
  }
  if (totEl) totEl.textContent = armFmtMoney(subtotal + imp);
  const impInp = $arm('lancImpostos');
  if (impInp && document.activeElement !== impInp && !draft.impostosManual) {
    impInp.value = imp > 0 ? fmtArmNum(imp, 2) : '';
  }
}

function armLancamentoRowValues(idx) {
  const row = document.querySelector(`#arm-lancEntryBody tr[data-idx="${idx}"]`);
  if (!row) return null;
  return {
    row,
    rawName: row.querySelector('[data-field=servico]')?.value || '',
    qtde: row.querySelector('[data-field=qtde]')?.value,
    unit: row.querySelector('[data-field=unit]')?.value
  };
}

/** Live preview while typing — updates draft + valor cell, never re-renders inputs. */
function armLancamentoRowPreview(idx) {
  if (!armLancamentoDraft?.servicos?.[idx]) return;
  const vals = armLancamentoRowValues(idx);
  if (!vals) return;
  armLancamentoDraft.servicos[idx] = armBuildServicoFromEntry(vals.rawName, vals.qtde, vals.unit);
  const s = armLancamentoDraft.servicos[idx];
  const valCell = vals.row.querySelector('.arm-val-calc');
  if (valCell) valCell.textContent = armFmtMoney(s.valor || 0);
  clearTimeout(armLancamentoSummaryTimer);
  armLancamentoSummaryTimer = setTimeout(armRefreshLancamentoSummary, 120);
}

/** Parse row on blur — format inputs, refresh summary; no full table re-render. */
function armLancamentoRowCommit(idx) {
  if (!armLancamentoDraft?.servicos?.[idx]) return;
  const vals = armLancamentoRowValues(idx);
  if (!vals) return;
  clearTimeout(armLancamentoSummaryTimer);
  armLancamentoDraft.servicos[idx] = armBuildServicoFromEntry(vals.rawName, vals.qtde, vals.unit);
  const s = armLancamentoDraft.servicos[idx];
  const qtInp = vals.row.querySelector('[data-field=qtde]');
  const unitInp = vals.row.querySelector('[data-field=unit]');
  if (qtInp) qtInp.value = fmtArmNum(s.qtde, 2);
  if (unitInp && !s.isNf) unitInp.value = fmtArmNum(s.valorUnit, 2);
  else if (unitInp && s.isNf) unitInp.value = fmtArmPct(ARM_NF_RATE * 100);
  const valCell = vals.row.querySelector('.arm-val-calc');
  if (valCell) valCell.textContent = armFmtMoney(s.valor || 0);
  armRefreshLancamentoSummary();
}

/** Serviço change may toggle NF unit readonly — re-render that row only. */
function armLancamentoServicoChange(idx) {
  armLancamentoRowCommit(idx);
  const draft = armLancamentoDraft;
  const body = $arm('lancEntryBody');
  if (!body || !draft?.servicos?.[idx]) return;
  const oldRow = body.querySelector(`tr[data-idx="${idx}"]`);
  if (!oldRow) return;
  const tmp = document.createElement('tbody');
  tmp.innerHTML = armLancamentoRowHtml(draft.servicos[idx], idx);
  oldRow.replaceWith(tmp.firstElementChild);
  if (typeof scheduleTableSort === 'function') scheduleTableSort();
}

function armLancamentoRowChange(idx) {
  armLancamentoRowCommit(idx);
}

function armLancamentoAddRow() {
  if (!armLancamentoDraft) armLoadLancamentoDraft();
  armLancamentoDraft.servicos.push(armBuildServicoFromEntry(ARM_NORM_HORA_EXTRA, 0, 0));
  renderArmLancamentoForm();
}

function armLancamentoRemoveRow(idx) {
  if (!armLancamentoDraft?.servicos) return;
  if (armLancamentoDraft.servicos.length <= 1) {
    armToast('Mínimo uma linha de serviço.', 'error');
    return;
  }
  armLancamentoDraft.servicos.splice(idx, 1);
  renderArmLancamentoForm();
}

function armLancamentoAddCustomService() {
  const inp = $arm('lancCustomSvc');
  const raw = String(inp?.value || '').trim();
  if (!raw) {
    armToast('Indica o nome do serviço.', 'error');
    return;
  }
  const norm = normalizeServicoName(raw);
  armEnsurePack();
  if (!armPack.customServices.includes(norm)) {
    armPack.customServices.push(norm);
    armPack.customServices.sort((a, b) => a.localeCompare(b, 'pt'));
  }
  if (!armLancamentoDraft) armLoadLancamentoDraft();
  armLancamentoDraft.servicos.push(armBuildServicoFromEntry(norm, 0, 0));
  if (inp) inp.value = '';
  renderArmLancamentoForm();
  armToast(`Serviço «${norm}» adicionado ao catálogo.`, 'success');
}

function armLancamentoCommitAllRows() {
  const draft = armLancamentoDraft;
  if (!draft?.servicos?.length) return;
  clearTimeout(armLancamentoSummaryTimer);
  draft.servicos.forEach((_, idx) => {
    const vals = armLancamentoRowValues(idx);
    if (!vals) return;
    draft.servicos[idx] = armBuildServicoFromEntry(vals.rawName, vals.qtde, vals.unit);
  });
}

async function armSaveLancamento() {
  if (armCompany() !== 'DFB') {
    armToast('Avaliação Armazém disponível apenas para DFB.', 'error');
    return;
  }
  if (!armLancamentoDraft) armLoadLancamentoDraft();
  armLancamentoCommitAllRows();
  const draft = armLancamentoDraft;
  if (!draft?.mesKey) {
    armToast('Seleciona mês de referência.', 'error');
    return;
  }
  const impInp = $arm('lancImpostos');
  draft.impostos = armNum(impInp?.value);
  if (!draft.impostosManual) {
    draft.impostos = armCalcImpostos(armServiceSubtotal(draft.servicos));
  }
  draft.servicos = (draft.servicos || []).filter(s => (s.valor || 0) > 0 || (s.qtde || 0) > 0);
  if (!draft.servicos.length) {
    armToast('Adiciona pelo menos uma linha com quantidade ou valor.', 'error');
    return;
  }

  const prevMonth = (armPack?.months || []).find(m => m.mesKey === draft.mesKey);
  const nfSource = armNfUploadRows.length ? armNfUploadRows : (prevMonth?.nfRows || []);

  const month = armRecalcMonthTotals(armNormalizeMonth({
    mesKey: draft.mesKey,
    mesLabel: draft.mesLabel,
    fileName: `manual_${draft.mesKey}`,
    format: 'manual',
    entrySource: 'manual',
    manualUpdatedAt: new Date().toISOString(),
    impostos: draft.impostos,
    impostosManual: !!draft.impostosManual,
    valorMinimo: ARM_MINIMO_CONTRATUAL,
    servicos: draft.servicos.map(s => ({ ...s })),
    nfRows: nfSource.map(r => applySapToArmNf({ ...r, mesKey: draft.mesKey }))
  }));

  if (month.nfRows?.length) {
    const nfSum = month.nfRows.reduce((a, r) => a + armNum(r.valorNF), 0);
    if (nfSum > 0) {
      let nfSvc = armFindNfServico(month);
      if (!nfSvc) {
        nfSvc = armBuildServicoFromEntry('PERCENTUAL SOBRE NF EXPEDIDA', nfSum, ARM_NF_RATE);
        month.servicos.unshift(nfSvc);
      } else {
        nfSvc.qtde = nfSum;
        nfSvc.valorUnit = ARM_NF_RATE;
        nfSvc.valor = nfSum * ARM_NF_RATE;
        nfSvc.isNf = true;
      }
      armRecalcMonthTotals(month);
    }
  }

  armEnsurePack();
  armPack.version = 3;
  armPack.updatedAt = new Date().toISOString();
  const others = (armPack.months || []).filter(m => m.mesKey !== month.mesKey);
  armPack.months = [...others, month].sort((a, b) => String(a.mesKey).localeCompare(String(b.mesKey)));
  refreshAllSapOnNfs();
  const saved = await persistArmPack(armPack);
  updateArmFileZone();
  armToast(`${armMesLabel(month)} guardado${saved ? ' na cloud' : ''}.`, 'success');
  armRefreshLancamentoMonthOptions();
  renderArmLancamentoForm();
}

function parseNfListFromWorkbook(wb, fileName, mesKey) {
  const nfRows = [];
  const nfSheet = wb.SheetNames.find(n => /PERCENTUAL SOBRE NF/i.test(n)) ||
    wb.SheetNames.find(n => /percentual.*nf/i.test(normSheetName(n)));
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
          fee: fee || valorNF * ARM_NF_RATE, mesKey, fileName
        });
      }
    }
  }
  if (!nfRows.length) {
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const m = sheetToMatrix(sheet);
    let hdr = -1;
    let colNf = 0;
    let colVal = 1;
    for (let r = 0; r < Math.min(30, m.length); r++) {
      const row = m[r] || [];
      for (let c = 0; c < row.length; c++) {
        const h = normHdrCell(row[c]);
        if (/^nf$|nota fiscal|documento|n[oº] doc/.test(h)) { hdr = r; colNf = c; }
        if (/valor nf|valor da nf|valor total|valor\b/.test(h)) colVal = c;
      }
      if (hdr >= 0) break;
    }
    if (hdr < 0) hdr = 0;
    for (let r = hdr + 1; r < m.length; r++) {
      const nf = cellText(m, r, colNf);
      const valorNF = armCellNum(m, r, colVal);
      if (!nf || nf === '-' || !valorNF) continue;
      if (!/^\d/.test(String(nf).trim()) && !/[A-Z0-9]{5,}/i.test(String(nf))) continue;
      nfRows.push({
        data: '', nf, valorNF, taxa: ARM_NF_RATE,
        fee: valorNF * ARM_NF_RATE, mesKey, fileName
      });
    }
  }
  return nfRows.map(applySapToArmNf);
}

async function armProcessNfUpload(file) {
  if (!file) return;
  const mesKey = armLancamentoMesKey();
  if (!mesKey) {
    armToast('Seleciona o mês de referência antes do upload.', 'error');
    return;
  }
  try {
    await armEnsureSapLoaded();
    const { wb } = await readFileToWorkbook(file, {});
    armNfUploadRows = parseNfListFromWorkbook(wb, file.name, mesKey);
    if (!armNfUploadRows.length) {
      armToast('Nenhuma NF encontrada no ficheiro — verifica formato.', 'error');
      return;
    }
    renderArmNfUploadTable();
    armToast(`${armNfUploadRows.length} NF(s) importada(s) — guarda o mês para persistir.`, 'success');
  } catch (e) {
    console.error('[armazem] nf upload', e);
    armToast('Erro ao ler Excel: ' + (e.message || e), 'error');
  }
}

function renderArmNfUploadTable() {
  const wrap = $arm('nfUploadSummary');
  const sapNote = $arm('nfUploadSapNote');
  if (sapNote) {
    sapNote.innerHTML = armSapStatusHtml();
    sapNote.style.display = 'block';
    sapNote.classList.toggle('data-warn', !sapApi()?.isSapLoaded?.());
  }
  if (!armNfUploadRows.length) {
    if (wrap) wrap.style.display = 'none';
    return;
  }
  if (wrap) wrap.style.display = 'block';

  const missing = armNfUploadRows.filter(r => r.sapMissing).length;
  const valDiff = armNfUploadRows.filter(r => r.valorDiff).length;
  const sumUni = armNfUploadRows.reduce((s, r) => s + armNum(r.valorNF), 0);
  const sumSap = armNfUploadRows.filter(r => r.sapFound).reduce((s, r) => s + armNum(r.sapValor), 0);

  const kpis = $arm('nfUploadKpis');
  if (kpis) {
    kpis.innerHTML = `
      <div class="kpi"><div class="label">NFs</div><div class="value">${armNfUploadRows.length}</div></div>
      <div class="kpi"><div class="label">Σ Unilog</div><div class="value">${armFmtMoney(sumUni)}</div></div>
      <div class="kpi"><div class="label">Σ SAP</div><div class="value">${sumSap > 0 ? armFmtMoney(sumSap) : '—'}</div></div>
      <div class="kpi${missing || valDiff ? ' flag' : ''}"><div class="label">Problemas</div><div class="value">${missing + valDiff}</div><div class="sub">${missing} ausente · ${valDiff} Δ valor</div></div>`;
  }

  let rows = armApplySort(armNfUploadRows, 'nfUpload', {
    nf: r => r.nf,
    valorNF: r => r.valorNF,
    sapValor: r => r.sapValor ?? -1,
    diff: r => (r.sapFound ? armNum(r.valorNF) - armNum(r.sapValor) : 0)
  });

  const thead = $arm('nfUploadHead');
  if (thead) {
    thead.innerHTML = `
      ${sortTh('nfUpload', 'nf', 'NF')}
      ${sortTh('nfUpload', 'valorNF', 'Valor Unilog', 'right')}
      ${sortTh('nfUpload', 'sapValor', 'Valor ZFACT/SAP', 'right')}
      ${sortTh('nfUpload', 'diff', 'Δ valor', 'right')}
      <th>Estado</th>`;
  }

  const body = $arm('nfUploadBody');
  if (body) {
    body.innerHTML = rows.map(r => {
      const diff = r.sapFound ? armNum(r.valorNF) - armNum(r.sapValor) : null;
      let cls = '';
      if (r.sapNotLoaded) cls = 'nf-sap-pending';
      else if (r.sapMissing) cls = 'nf-sap-missing';
      else if (r.valorDiff) cls = 'nf-val-mismatch';
      const estado = armNfEstadoBadge(r);
      return `<tr class="${cls}">
        <td>${armEsc(r.nf)}</td>
        <td class="right">${armFmtMoney(r.valorNF)}</td>
        <td class="right">${r.sapFound ? armFmtMoney(r.sapValor) : '—'}</td>
        <td class="right">${diff != null && Math.abs(diff) > 0.02 ? armFmtMoney(diff) : '—'}</td>
        <td>${estado}</td>
      </tr>`;
    }).join('');
  }
  if (typeof scheduleTableSort === 'function') scheduleTableSort();
}

function renderArmLancamento() {
  armInitLancamentoSelectors();
  armRefreshLancamentoMonthOptions();
  if (!armLancamentoDraft) armLoadLancamentoDraft();
  else renderArmLancamentoForm();
  renderArmNfUploadTable();
}

function armLancamentoResetImpostosAuto() {
  if (!armLancamentoDraft) return;
  armLancamentoDraft.impostosManual = false;
  armLancamentoDraft.impostos = armCalcImpostos(armServiceSubtotal(armLancamentoDraft.servicos));
  renderArmLancamentoForm();
}

function armCleanFileBase(fileName) {
  return String(fileName || '')
    .replace(/\.xlsx?$/i, '')
    .replace(/[_\s.-]+$/g, '')
    .trim()
    .replace(/^receita[_\s]*(?:e[_\s]*)?custo\s*delta\s*[-–—:]?\s*/i, '')
    .trim();
}

function parseMonthFromFileName(fileName) {
  const base = armCleanFileBase(fileName);
  const yearMatch = base.replace(/[_\s.-]+$/g, '').match(/(\d{4})\s*$/);
  const year = yearMatch ? parseInt(yearMatch[1], 10) : null;
  const low = base.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  let month = null;
  const meses = Object.entries(ARM_MESES_PT).sort((a, b) => b[0].length - a[0].length);
  for (const [name, num] of meses) {
    const key = name.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const re = new RegExp('(?:^|[^a-z])' + key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:[^a-z]|$)');
    if (re.test(low)) { month = num; break; }
  }
  if (!month || !year) return { mesKey: '', mesLabel: base || String(fileName || '') };
  const mesKey = `${year}-${String(month).padStart(2, '0')}`;
  const mesLabel = armMesNomeLong(mesKey, '');
  return { mesKey, mesLabel };
}

function armMesLabel(m) {
  if (!m) return '—';
  const obj = typeof m === 'string' ? { mesKey: m } : m;
  const mesKey = String(obj.mesKey || '').trim();
  if (/^\d{4}-\d{2}$/.test(mesKey)) return armMesNomeLong(mesKey, '');
  for (const src of [obj.fileName, obj.mesLabel]) {
    if (!src) continue;
    const parsed = parseMonthFromFileName(src);
    if (parsed.mesKey) return parsed.mesLabel;
  }
  const fallback = String(obj.mesLabel || obj.fileName || mesKey || '').trim();
  return fallback || '—';
}

function armMergeServicosByNorm(servicos) {
  const map = new Map();
  (servicos || []).forEach(s => {
    const norm = normalizeServicoName(s.rawName || s.normName);
    if (!norm) return;
    const prev = map.get(norm);
    if (!prev) {
      map.set(norm, { ...s, normName: norm });
      return;
    }
    const q1 = Number(prev.qtde) || 0;
    const q2 = Number(s.qtde) || 0;
    const v1 = Number(prev.valor) || 0;
    const v2 = Number(s.valor) || 0;
    const u1 = Number(prev.valorUnit) || 0;
    const u2 = Number(s.valorUnit) || 0;
    const qtde = q1 + q2;
    const valor = v1 + v2;
    let valorUnit = prev.valorUnit;
    if (!prev.isNf && !s.isNf && qtde > 0 && (u1 || u2)) {
      valorUnit = ((q1 * u1) + (q2 * u2)) / qtde;
    }
    map.set(norm, {
      ...prev,
      normName: norm,
      qtde,
      valor,
      valorUnit,
      isNf: prev.isNf || s.isNf
    });
  });
  return [...map.values()];
}

function armNormalizeMonth(m) {
  if (!m) return m;
  let mesKey = String(m.mesKey || '').trim();
  if (!/^\d{4}-\d{2}$/.test(mesKey)) {
    const fromFile = parseMonthFromFileName(m.fileName || m.mesLabel || '');
    if (fromFile.mesKey) mesKey = fromFile.mesKey;
    else if (m.dataInicial) mesKey = String(m.dataInicial).slice(0, 7);
  }
  m.mesKey = mesKey;
  m.mesLabel = armMesLabel({ mesKey, mesLabel: m.mesLabel, fileName: m.fileName });
  if (m.servicos?.length) {
    m.servicos = armMergeServicosByNorm(m.servicos.map(s => ({
      ...s,
      normName: normalizeServicoName(s.rawName || s.normName)
    })));
  }
  armRepairNfServico(m);
  if (m.vendasLiq != null && m.vendasLiq !== '') {
    const vl = armNum(m.vendasLiq);
    if (vl > 0) m.vendasLiq = vl;
    else delete m.vendasLiq;
  }
  return m;
}

function armMonthMergeKey(m) {
  const mk = String(m?.mesKey || '').trim();
  if (/^\d{4}-\d{2}$/.test(mk)) return mk;
  const fn = String(m?.fileName || '').trim().toLowerCase();
  return fn || mk || '';
}

function armMonthDataScore(m) {
  return (m?.servicos?.length || 0) * 100000
    + (m?.nfRows?.length || 0) * 100
    + (Number(m?.valorTotal) > 0 ? 10 : 0)
    + (Number(m?.totalServicos) > 0 ? 1 : 0);
}

function armDedupePackMonths(pack) {
  if (!pack?.months?.length) return pack;
  const byKey = new Map();
  for (const m of pack.months) {
    const k = armMonthMergeKey(m);
    if (!k) {
      byKey.set(`__anon_${byKey.size}`, m);
      continue;
    }
    const prev = byKey.get(k);
    if (!prev) {
      byKey.set(k, m);
      continue;
    }
    const keep = armMonthDataScore(m) >= armMonthDataScore(prev) ? m : prev;
    const drop = keep === m ? prev : m;
    if ((keep.servicos?.length || 0) > (drop.servicos?.length || 0)) {
      console.warn(`[armazem] dedupe ${k}: kept ${armMesLabel(keep)} (${keep.servicos.length} srv), dropped stale (${drop.fileName || '?'})`);
    }
    byKey.set(k, keep);
  }
  pack.months = [...byKey.values()].sort((a, b) => String(a.mesKey).localeCompare(String(b.mesKey)));
  return pack;
}

function armNormalizePackMonths(pack) {
  if (!pack?.months?.length) return pack;
  pack.months.forEach(m => {
    armNormalizeMonth(m);
    armRecalcMonthTotals(m);
  });
  return armDedupePackMonths(pack);
}

function armLogPackMonths(label) {
  const months = armPack?.months || [];
  console.log(`[armazem] ${label}: ${months.length} month(s)`,
    months.map(m => `${m.mesKey || '?'}:${m.servicos?.length || 0}srv`).join('; '));
  months.filter(m => !(m.servicos?.length)).forEach(m => {
    console.warn(`[armazem] ${armMesLabel(m)} sem servicos no resumo — processar de novo: ${m.fileName || '(sem ficheiro)'}`);
  });
}

function armInfoLabelValue(matrix, r, labelRe) {
  const row = matrix[r] || [];
  for (let c = 0; c < row.length; c++) {
    const label = cellText(matrix, r, c);
    if (!labelRe.test(label)) continue;
    for (let v = c + 1; v < Math.min(c + 4, row.length); v++) {
      const t = cellText(matrix, r, v);
      if (t && t !== '-') return t;
    }
  }
  return '';
}

function parseInfoSheetValues(info) {
  const out = {
    depositante: '', dataInicial: '', dataFinal: '',
    totalServicos: 0, valorMinimo: ARM_MINIMO_CONTRATUAL,
    valorApurado: 0, valorTotal: 0
  };
  if (!info?.length) return out;
  for (let r = 0; r < info.length; r++) {
    const dep = armInfoLabelValue(info, r, /Depositante/i);
    if (dep) out.depositante = dep;
    const di = armInfoLabelValue(info, r, /Data Inicial/i);
    if (di) out.dataInicial = parseBrDate(di);
    const df = armInfoLabelValue(info, r, /Data Final/i);
    if (df) out.dataFinal = parseBrDate(df);
    const ts = armInfoLabelValue(info, r, /Total dos Servi/i);
    if (ts) out.totalServicos = armNum(ts) || out.totalServicos;
    const vm = armInfoLabelValue(info, r, /Valor M[ií]nimo Contratual/i);
    if (vm) out.valorMinimo = armNum(vm) || out.valorMinimo;
    const va = armInfoLabelValue(info, r, /Valor Apurado/i);
    if (va) out.valorApurado = armNum(va) || out.valorApurado;
    const vt = armInfoLabelValue(info, r, /VALOR FINAL A FATURAR|Valor Total \(com impostos\)/i);
    if (vt) out.valorTotal = armNum(vt) || out.valorTotal;
  }
  return out;
}

function armGetArmazenagemValor(month) {
  const fromInfo = Number(month.totalServicos) || 0;
  const fromRows = armResumoRows(month).reduce((a, s) => a + (Number(s.valor) || 0), 0);
  const subtotal = fromInfo > 0 ? fromInfo : fromRows;
  const finalVal = Number(month.valorTotal) || 0;
  const pago = subtotal > 0 ? subtotal : finalVal;
  return { subtotal, final: finalVal, pago, primary: pago };
}

function armFindNfServico(month) {
  return (month?.servicos || []).find(s =>
    s.isNf || isNfPercentualService(s.rawName || s.normName)
  ) || null;
}

/** NF resumo: qtde = faturação expedida (base); valor = taxa 5,5%. Recover when fee was stored as qtde. */
function armNfBaseFromServico(row) {
  if (!row) return 0;
  const q = armNum(row.qtde);
  const fee = armNum(row.valor);
  if (q <= 0) {
    if (fee > 0) return fee / ARM_NF_RATE;
    return 0;
  }
  if (fee > 0) {
    const expectedFee = q * ARM_NF_RATE;
    const relErr = Math.abs(expectedFee - fee) / Math.max(expectedFee, fee, 1);
    if (relErr < 0.08) return q;
    // qtde ≈ valor → both read the fee column, not the NF base
    if (Math.abs(q - fee) / Math.max(q, fee, 1) < 0.03) return q / ARM_NF_RATE;
  }
  if (q >= 500000) return q;
  const asBase = q / ARM_NF_RATE;
  if (asBase >= 500000) return asBase;
  return q;
}

function armRepairNfServico(month) {
  const row = armFindNfServico(month);
  if (!row) return;
  row.isNf = true;
  const base = armNfBaseFromServico(row);
  if (base > 0) {
    row.qtde = base;
    row.valorUnit = ARM_NF_RATE;
    row.valor = base * ARM_NF_RATE;
  }
}

function armGetVendasLiq(month) {
  const v = armNum(month?.vendasLiq);
  return v > 0 ? v : null;
}

async function armSetVendasLiq(mesKey, rawVal) {
  const val = armNum(rawVal);
  const pack = armEnsurePack();
  const month = (pack.months || []).find(m => m.mesKey === mesKey);
  if (!month) return;
  const prev = armGetVendasLiq(month) || 0;
  if (val === prev) return;
  if (val > 0) month.vendasLiq = val;
  else delete month.vendasLiq;
  month.vendasLiqUpdatedAt = new Date().toISOString();
  pack.updatedAt = new Date().toISOString();
  const saved = await persistArmPack(pack);
  if (saved) armToast('Vendas Liq guardadas.');
  if (armActiveTab === 'mensal') renderArmMensal();
  if (armActiveTab === 'adicionais') renderArmAdicionais();
}

function armGetNfBase(month) {
  const fromServico = armNfBaseFromServico(armFindNfServico(month));
  const nfSum = (month?.nfRows || []).reduce((a, r) => a + armNum(r.valorNF), 0);
  if (fromServico > 0 && nfSum > 0) {
    if (nfSum > fromServico * 1.5) return nfSum;
    if (fromServico > nfSum * 1.5) return fromServico;
    return Math.max(fromServico, nfSum);
  }
  return fromServico > 0 ? fromServico : nfSum > 0 ? nfSum : 0;
}

/** Taxa 5,5% NF — valor (R$) da linha PERCENTUAL SOBRE NF EXPEDIDA, nunca a base NF. */
function armGetNfFee(month) {
  const row = armFindNfServico(month);
  if (row) {
    const fee = armNum(row.valor);
    const base = armNfBaseFromServico(row);
    if (fee > 0 && base > 0) {
      const expected = base * ARM_NF_RATE;
      const relErr = Math.abs(expected - fee) / Math.max(expected, fee, 1);
      if (relErr < 0.08) return fee;
      if (Math.abs(armNum(row.qtde) - fee) / Math.max(fee, 1) < 0.03) return expected;
      if (fee <= expected * 1.01) return fee;
      return expected;
    }
    if (fee > 0 && fee < 500000) return fee;
    if (base > 0) return Math.round(base * ARM_NF_RATE * 100) / 100;
  }
  const nfBase = armGetNfBase(month);
  return nfBase > 0 ? Math.round(nfBase * ARM_NF_RATE * 100) / 100 : 0;
}

function armGetAdicionaisSum(month) {
  const fromArr = (month?.adicionais || []).reduce((s, a) => s + (Number(a.valor) || 0), 0);
  if (fromArr > 0) return fromArr;
  return armResumoRows(month).filter(s => !s.isNf).reduce((s, r) => s + (Number(r.valor) || 0), 0);
}

/** Visão mensal: armazenagem 5,5% + adicionais (s/ impostos). */
function armGetMensalPagoSemImp(month) {
  return armGetNfFee(month) + armGetAdicionaisSum(month);
}

/** Valor cobrado na rubrica 5,5% sobre NF expedida (armazenagem principal). */
function armGetArmazenagemRubricaValor(month) {
  const nf = armFindNfServico(month);
  if (nf && Number(nf.valor) > 0) return Number(nf.valor);
  return (month?.servicos || [])
    .filter(s => s.isNf || s.catalogId === 'nf_percentual')
    .reduce((a, s) => a + (Number(s.valor) || 0), 0);
}

function armFmtPctGasto(armVal, nfBase) {
  if (!Number.isFinite(nfBase) || nfBase <= 0) return '—';
  if (!Number.isFinite(armVal) || armVal < 0) return '—';
  return fmtArmPct(armVal / nfBase * 100, 2);
}

function armVendasFromSapNf() {
  const api = sapApi();
  if (!api?.isSapLoaded?.()) return { byMes: {}, loaded: false, source: null, note: '' };
  const byMes = {};
  Object.values(api.getMap()).forEach(entry => {
    const mk = String(entry.dtEmissao || '').slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(mk)) return;
    if (!byMes[mk]) byMes[mk] = { valor: 0 };
    byMes[mk].valor += Number(entry.valorNF) || 0;
  });
  const loaded = Object.keys(byMes).length > 0;
  return {
    byMes, loaded, source: loaded ? 'sap_nf' : null,
    note: loaded ? 'Vendas estimadas via export SAP NF (Fretes)' : ''
  };
}

async function loadArmVendasByMonth(force) {
  const now = Date.now();
  if (!force && armVendasCache && (now - armVendasCacheTs) < ARM_VENDAS_CACHE_MS) return armVendasCache;
  armVendasCache = { byMes: {}, loaded: false, source: null, note: 'carregar vendas em Dados → Carregar Vendas' };
  armVendasCacheTs = now;

  if (typeof db !== 'undefined') {
    try {
      const { data, error } = await db.from('logistica_vendas').select('mes,ean,qt').limit(50000);
      if (!error && data?.length) {
        const precoMap = typeof buildSapPrecoMap === 'function' ? buildSapPrecoMap() : {};
        const hasPreco = Object.keys(precoMap).length > 0;
        const byMes = {};
        data.forEach(r => {
          const qt = parseFloat(r.qt) || 0;
          if (qt <= 0) return;
          if (!byMes[r.mes]) byMes[r.mes] = { valor: 0 };
          const preco = precoMap[r.ean] || 0;
          if (hasPreco && preco > 0) byMes[r.mes].valor += qt * preco;
        });
        if (hasPreco && Object.values(byMes).some(x => x.valor > 0)) {
          armVendasCache = { byMes, loaded: true, source: 'vendas', note: 'Vendas (qt × preço SAP, todos os canais)' };
          return armVendasCache;
        }
        armVendasCache.note = 'Vendas carregadas — falta SAP (preços) para calcular valor';
      }
    } catch (e) {
      console.warn('[armazem] vendas', e);
    }
  }

  const sap = armVendasFromSapNf();
  if (sap.loaded) {
    armVendasCache = sap;
    return armVendasCache;
  }
  return armVendasCache;
}

function armVendasForMes(vendasPack, mesKey) {
  const v = vendasPack?.byMes?.[mesKey];
  if (!v || !Number.isFinite(v.valor) || v.valor <= 0) return null;
  return v.valor;
}

function armInvalidateVendasCache() {
  armVendasCache = null;
  armVendasCacheTs = 0;
}

function renderArmResumoSections(months) {
  const byYear = {};
  months.forEach(m => {
    const y = (m.mesKey || '').slice(0, 4) || 'Outros';
    if (!byYear[y]) byYear[y] = [];
    byYear[y].push(m);
  });
  Object.keys(byYear).forEach(y => {
    byYear[y].sort((a, b) => String(a.mesKey).localeCompare(String(b.mesKey)));
  });

  let html = '';
  Object.keys(byYear).sort().forEach(year => {
    html += `<div class="arm-resumo-year"><div class="section-title arm-resumo-year-title">${armEsc(year)}</div>`;
    html += byYear[year].map(m => {
      const rows = armResumoRows(m);
      const titulo = armMesLabel(m);
      const emptyWarn = rows.length ? '' : (
        `<p class="data-warn" style="margin:0 0 8px;">Sem linhas de resumo guardadas — seleciona e processa de novo `
        + `<strong>${armEsc(m.fileName || titulo)}</strong> na aba Carregamento.</p>`
      );
      return `<div class="arm-resumo-month">
        <div class="section-title arm-resumo-month-title">${armEsc(titulo)}</div>
        ${emptyWarn}
        ${resumoTableHtml(rows)}
      </div>`;
    }).join('');
    html += '</div>';
  });

  const agg = aggregateResumoByNorm(months);
  html += `<div class="arm-resumo-month arm-resumo-total-block">
    <div class="section-title arm-resumo-month-title">Total geral</div>
    ${resumoTableHtml(agg, { useNorm: true, total: true })}
  </div>`;
  return html;
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

function armResumoNomeHdrRe(h) {
  return /tipo de servi|codigo servi|cod servi|descricao.*servi|nome.*servi|^servico$|^descricao$/.test(h);
}

function findResumoHeaderRow(matrix) {
  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r] || [];
    let hasNome = false;
    let hasQtde = false;
    for (let c = 0; c < row.length; c++) {
      const h = normHdrCell(row[c]);
      if (armResumoNomeHdrRe(h)) hasNome = true;
      if (/qtde\s*servi|qtde|quantidade|qtd\b/.test(h)) hasQtde = true;
    }
    if (hasNome && hasQtde) return r;
  }
  return -1;
}

function scanResumoCols(matrix, hdrRow, fallback) {
  const cols = {
    ...(fallback || { nome: 1, qtde: 2, unit: 3, valor: 4 }),
    nomeTipo: -1,
    nomeCodigo: -1,
    nomeDesc: -1
  };
  const row = matrix[hdrRow] || [];
  for (let c = 0; c < row.length; c++) {
    const h = normHdrCell(row[c]);
    if (!h) continue;
    if (/qtde\s*servi|qtde|quantidade|qtd\b/.test(h)) cols.qtde = c;
    else if (/valor unit|unitario/.test(h)) cols.unit = c;
    else if (/valor calc|calculado|valor total|valor servi/.test(h)) cols.valor = c;
    else if (/tipo de servi/.test(h)) cols.nomeTipo = c;
    else if (/codigo servi|cod servi/.test(h)) cols.nomeCodigo = c;
    else if (/descricao.*servi|^descricao$|nome.*servi|^servico$/.test(h)) cols.nomeDesc = c;
  }
  // Default nome col for resumoRowServiceName fallbacks (prefer descriptive columns at row level).
  if (cols.nomeDesc >= 0) cols.nome = cols.nomeDesc;
  else if (cols.nomeTipo >= 0) cols.nome = cols.nomeTipo;
  else if (cols.nomeCodigo >= 0) cols.nome = cols.nomeCodigo;
  return cols;
}

function resumoServiceNameScore(text) {
  const s = String(text ?? '').trim();
  if (!s || s === '-') return 0;
  const hasAlpha = /[A-Za-zÀ-ú]/.test(s);
  return s.length + (hasAlpha ? 250 : (/^\d+$/.test(s) ? -50 : 0));
}

function resumoRowServiceName(matrix, r, cols) {
  const seen = new Set();
  const candidates = [];
  const tryCol = (c) => {
    if (c == null || c < 0 || seen.has(c)) return;
    seen.add(c);
    const t = cellText(matrix, r, c);
    if (t === '' || t === null || t === undefined) return;
    candidates.push(String(t).trim());
  };
  for (const c of [cols.nomeDesc, cols.nomeTipo, cols.nomeCodigo, cols.nome, 1, 0, 2, 3, 4, 5, 6]) {
    tryCol(c);
  }
  if (!candidates.length) return '';
  candidates.sort((a, b) => resumoServiceNameScore(b) - resumoServiceNameScore(a));
  return candidates[0];
}

function armIsResumoFooterRow(nome) {
  return /^Apura|^Total\b|^Impostos/i.test(String(nome || '').trim());
}

function parseResumoServicosFromMatrix(matrix, opts) {
  const servicos = [];
  const hdr = findResumoHeaderRow(matrix);
  if (hdr < 0) return { servicos, hdr: -1, cols: null };
  const cols = scanResumoCols(matrix, hdr, opts?.colFallback);
  for (let r = hdr + 1; r < matrix.length; r++) {
    const nome = resumoRowServiceName(matrix, r, cols);
    if (!nome) continue;
    if (armIsResumoFooterRow(nome)) continue;
    const isNf = isNfPercentualService(nome);
    const nums = armFinishServico(
      nome,
      armCellNum(matrix, r, cols.qtde),
      armCellNum(matrix, r, cols.unit),
      armCellNum(matrix, r, cols.valor),
      isNf
    );
    if (!nums.valor && !nums.qtde) continue;
    const norm = normalizeServicoName(nome);
    const cat = resolveCatalogCategory(nome);
    servicos.push({
      rawName: nome, normName: norm, codigo: nome,
      qtde: nums.qtde, valorUnit: nums.valorUnit, valor: nums.valor,
      catalogId: cat.id, catalogSure: cat.sure, isNf
    });
  }
  return { servicos, hdr, cols };
}

function parseV1LegacyResumoServicos(matrix) {
  const servicos = [];
  let resumoHdr = -1;
  for (let r = 0; r < matrix.length; r++) {
    const t0 = cellText(matrix, r, 0);
    const t1 = cellText(matrix, r, 1);
    if (/^\s*Resumo\s*$/i.test(t1) || /^\s*Resumo\s*$/i.test(t0)) { resumoHdr = r; break; }
  }
  if (resumoHdr < 0) return servicos;
  const hdrRow = resumoHdr + 1;
  const cols = scanResumoCols(matrix, hdrRow, { nome: 1, qtde: 17, unit: 25, valor: 28 });
  for (let r = resumoHdr + 2; r < matrix.length; r++) {
    const nome = resumoRowServiceName(matrix, r, cols);
    if (!nome) continue;
    if (/^Apura/i.test(nome) || /IMPOSTO/i.test(nome)) continue;
    const isNf = isNfPercentualService(nome);
    const nums = armFinishServico(
      nome,
      armCellNum(matrix, r, cols.qtde),
      armCellNum(matrix, r, cols.unit),
      armCellNum(matrix, r, cols.valor),
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
  return servicos;
}

function armGuessWrongFileError(wb, fileName) {
  if (findResumoSheetName(wb)) return null;
  for (const n of wb.SheetNames || []) {
    const m = sheetToMatrix(wb.Sheets[n]);
    if (findResumoHeaderRow(m) >= 0) return null;
  }
  const fn = String(fileName || '');
  if (/zfact|z_fact/i.test(fn)) {
    return 'Ficheiro ZFACT (SAP NF) — carrega em Fretes → Excel SAP NF. Armazém espera Receita_Custo DELTA (Unilog).';
  }
  try {
    const sapLoad = sapApi()?.loadSapRowsFromWorkbook?.(wb);
    if (sapLoad?.valid >= 5) {
      return 'Parece export SAP NF (ZFACT) — carrega em Fretes CT-e, não no Armazém. Aqui usa Receita_Custo DELTA da Unilog.';
    }
  } catch (_) { /* ignore */ }
  const sheets = (wb.SheetNames || []).join(', ') || '(nenhuma)';
  return `Nenhum serviço no bloco Resumo — folhas: ${sheets}. Esperado folha "Resumo de Serviços" (v2) ou bloco Resumo (v1).`;
}

function armLogParsedMonth(record, fileName) {
  if (!record) return;
  const label = armMesLabel(record);
  const n = record.servicos?.length || 0;
  const sub = record.totalServicos || 0;
  const fin = record.valorTotal || 0;
  console.log(`[armazem] ${label} parsed: ${n} servicos`, {
    mesKey: record.mesKey, fileName: fileName || record.fileName,
    subtotal: sub, final: fin, format: record.format,
    nfBase: armGetNfBase(record)
  });
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

function normSheetName(name) {
  return String(name || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function findSheetName(wb, pattern) {
  const re = typeof pattern === 'string' ? new RegExp(pattern, 'i') : pattern;
  return wb.SheetNames.find(n => re.test(normSheetName(n))) || null;
}

function isArmResumoSheetName(name) {
  const s = normSheetName(name);
  return /^resumo(\s+(de\s+)?)?servi/.test(s) || s === 'resumo';
}

function findResumoSheetName(wb) {
  const names = wb.SheetNames || [];
  const byPattern = names.find(n => isArmResumoSheetName(n));
  if (byPattern) return byPattern;
  for (const n of names) {
    const m = sheetToMatrix(wb.Sheets[n]);
    if (findResumoHeaderRow(m) >= 0) return n;
  }
  return null;
}

function isArmBloatSheet(name) {
  const n = normSheetName(name);
  return /^(arm|stock)$/.test(n) || /^armazen/.test(n) || /estoque/.test(n);
}

function pickBillingSheets(sheetNames) {
  if (!sheetNames?.length) return [];
  const names = [...sheetNames];
  const v2Keep = names.filter(n => {
    if (isArmBloatSheet(n)) return false;
    const s = normSheetName(n);
    return isArmResumoSheetName(n) || /^informa/.test(s);
  });
  if (v2Keep.length) return v2Keep;
  const v1Keep = names.filter(n => !isArmBloatSheet(n));
  return v1Keep.length ? v1Keep : [names[0]];
}

function armShortMonthLabel(fileName, mesKey, mesLabel) {
  return armMesLabel({ mesKey, mesLabel, fileName });
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

function parseArmazemV2(wb, fileName, opts) {
  const skipNfDetail = !!opts?.skipNfDetail;
  const infoName = findSheetName(wb, /^Informa/);
  const resumoName = findResumoSheetName(wb);
  const diarioName = findSheetName(wb, /^Detalhamento Di/);
  if (!resumoName) throw new Error('Folha "Resumo de Serviços" não encontrada (formato v2).');

  const meta = parseMonthFromFileName(fileName);
  const info = infoName ? sheetToMatrix(wb.Sheets[infoName]) : [];
  const resumo = sheetToMatrix(wb.Sheets[resumoName]);

  const infoVals = parseInfoSheetValues(info);
  let dataInicial = infoVals.dataInicial;
  let dataFinal = infoVals.dataFinal;
  let totalServicos = infoVals.totalServicos;
  let valorMinimo = infoVals.valorMinimo;
  let valorApurado = infoVals.valorApurado;
  let valorTotal = infoVals.valorTotal;
  let depositante = infoVals.depositante;

  if (!meta.mesKey && dataInicial) meta.mesKey = dataInicial.slice(0, 7);
  if (meta.mesKey && !meta.mesLabel) {
    meta.mesLabel = armMesNomeLong(meta.mesKey, '');
  }

  const hdr = findResumoHeaderRow(resumo);
  if (hdr < 0) throw new Error('Cabeçalho do resumo não encontrado.');
  const cols = scanResumoCols(resumo, hdr);
  const parsedResumo = parseResumoServicosFromMatrix(resumo);
  const servicos = parsedResumo.servicos;

  const totalCalc = servicos.filter(s => !/IMPOSTO/i.test(s.normName)).reduce((a, s) => a + s.valor, 0);
  console.log(`[armazem] parseV2 resumo`, {
    fileName, mesKey: meta.mesKey, resumoSheet: resumoName, hdr, cols,
    servicos: servicos.length, subtotalCalc: totalCalc
  });

  const nfRows = [];
  if (!skipNfDetail) {
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
  }

  return buildMonthRecord({
    fileName, format: 'v2', meta, depositante, dataInicial, dataFinal,
    totalServicos: totalServicos || totalCalc,
    valorMinimo, valorApurado: valorApurado || totalCalc, valorTotal,
    servicos, nfRows,
    partialParse: skipNfDetail,
    nfSkipped: skipNfDetail,
    parseNote: skipNfDetail ? 'NF detalhe omitido (ficheiro grande)' : ''
  });
}

function parseArmazemV1(wb, fileName, opts) {
  const skipNfDetail = !!opts?.skipNfDetail;
  const meta = parseMonthFromFileName(fileName);

  let depositante = '';
  let servicos = [];
  let mainMatrix = sheetToMatrix(wb.Sheets[wb.SheetNames[0]]);
  for (const sheetName of wb.SheetNames || []) {
    if (isArmBloatSheet(sheetName)) continue;
    const m = sheetToMatrix(wb.Sheets[sheetName]);
    if (!depositante) {
      for (let r = 0; r < Math.min(20, m.length); r++) {
        for (let c = 0; c < 10; c++) {
          const v = cellText(m, r, c);
          if (/DELTA FOODS|14830817000100/i.test(v)) depositante = v;
        }
      }
    }
    const fromHeader = parseResumoServicosFromMatrix(m);
    const fromLegacy = parseV1LegacyResumoServicos(m);
    const found = fromHeader.servicos.length >= fromLegacy.length ? fromHeader.servicos : fromLegacy;
    if (found.length > servicos.length) {
      servicos = found;
      mainMatrix = m;
    }
  }
  const m = mainMatrix;

  const nfRows = [];
  if (!skipNfDetail) {
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
  }

  const totalCalc = servicos.filter(s => !/IMPOSTO/i.test(s.normName)).reduce((a, s) => a + s.valor, 0);
  let dataInicial = nfRows[0]?.data || '';
  let dataFinal = nfRows[nfRows.length - 1]?.data || '';

  return buildMonthRecord({
    fileName, format: 'v1', meta, depositante, dataInicial, dataFinal,
    totalServicos: totalCalc, valorMinimo: ARM_MINIMO_CONTRATUAL,
    valorApurado: totalCalc, valorTotal: totalCalc,
    servicos, nfRows,
    partialParse: skipNfDetail,
    nfSkipped: skipNfDetail,
    parseNote: skipNfDetail ? 'NF detalhe omitido (ficheiro grande)' : ''
  });
}

function buildMonthRecord(opts) {
  const nfRows = (opts.nfRows || []).map(applySapToArmNf);
  const belowMin = opts.totalServicos < (opts.valorMinimo || ARM_MINIMO_CONTRATUAL);
  return armNormalizeMonth({
    fileName: opts.fileName,
    format: opts.format,
    mesKey: opts.meta?.mesKey || '',
    mesLabel: opts.meta?.mesLabel || armMesNomeLong(opts.meta?.mesKey, '') || opts.fileName,
    dataInicial: opts.dataInicial || '',
    dataFinal: opts.dataFinal || '',
    depositante: opts.depositante || '',
    totalServicos: opts.totalServicos || 0,
    valorMinimo: opts.valorMinimo || ARM_MINIMO_CONTRATUAL,
    valorApurado: opts.valorApurado || 0,
    valorTotal: opts.valorTotal || 0,
    impostos: armNum(opts.impostos),
    entrySource: opts.entrySource || 'excel',
    belowMin,
    servicos: opts.servicos || [],
    nfRows,
    adicionais: (opts.servicos || []).filter(s => !s.isNf && s.catalogId !== 'impostos'),
    partialParse: !!opts.partialParse,
    nfSkipped: !!opts.nfSkipped,
    parseNote: opts.parseNote || ''
  });
}

function parseArmazemWorkbook(wb, fileName, opts) {
  const hasV2 = !!findResumoSheetName(wb);
  const rec = hasV2 ? parseArmazemV2(wb, fileName, opts) : parseArmazemV1(wb, fileName, opts);
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

let _armSapEnsurePromise = null;

/** Wait for shared Fretes SAP NF map (ZFACT) — no separate upload in Armazém. */
async function armEnsureSapLoaded() {
  if (sapApi()?.isSapLoaded?.()) return true;
  if (_armSapEnsurePromise) return _armSapEnsurePromise;
  _armSapEnsurePromise = (async () => {
    try {
      if (typeof loadSavedFretesFiles === 'function') {
        await loadSavedFretesFiles(true);
      } else if (sapApi()?.ensureLoaded) {
        await sapApi().ensureLoaded(true);
      }
    } catch (e) {
      console.warn('[armazem] ensure sap', e);
    }
    return !!sapApi()?.isSapLoaded?.();
  })().finally(() => { _armSapEnsurePromise = null; });
  return _armSapEnsurePromise;
}

function armGotoFretesPage(ev) {
  ev?.preventDefault?.();
  document.querySelector('.nt[data-page="fretes"]')?.click();
}

function armSapStatusHtml() {
  const api = sapApi();
  if (api?.isSapLoaded?.()) {
    const n = Object.keys(api.getMap()).length;
    return `✓ SAP NF (ZFACT) carregado — <strong>${n.toLocaleString('pt-BR')}</strong> NFs disponíveis (partilhado com Fretes, sem re-upload).`;
  }
  return '⚠ <strong>Carrega ZFACT em Dados/Fretes primeiro</strong> — menu '
    + '<a href="#" class="arm-link-fretes">Fretes CT-e</a> → Excel SAP NF. '
    + 'Não é necessário carregar outra vez aqui no Armazém.';
}

function armNfEstadoBadge(r) {
  if (r.sapNotLoaded) return '<span class="badge b-warn">SAP não carregado</span>';
  if (r.sapMissing) return '<span class="badge b-flag">Ausente ZFACT</span>';
  if (r.valorDiff) return '<span class="badge b-fix165">Δ valor</span>';
  return '<span class="badge b-ok">OK</span>';
}

function refreshArmazemSapValidation() {
  refreshAllSapOnNfs();
  armNfUploadRows = armNfUploadRows.map(r => applySapToArmNf(r));
  if (armInited) renderArmActiveTab();
}

async function armRefreshWithSap() {
  await armEnsureSapLoaded();
  refreshAllSapOnNfs();
  armNfUploadRows = armNfUploadRows.map(r => applySapToArmNf(r));
  renderArmActiveTab();
}

function applySapToArmNf(row) {
  const api = sapApi();
  const out = { ...row };
  out.valorUnilog = armNum(row.valorNF);
  out.feeUnilog = armNum(row.fee);
  out.feeExpected = out.valorUnilog * ARM_NF_RATE;
  out.feeDelta = out.feeUnilog - out.feeExpected;

  const sapLoaded = !!(api && api.isSapLoaded());
  out.sapLoaded = sapLoaded;
  out.sapNotLoaded = !sapLoaded;

  if (!sapLoaded) {
    out.sapFound = false;
    out.sapMissing = false;
    out.sapValor = null;
    out.valorDiff = false;
    return out;
  }

  const sap = api.lookupSapEntry(row.nf);
  out.sapFound = !!sap;
  out.sapMissing = !sap;
  out.sapValor = sap ? sap.valorNF : null;
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
  return findResumoSheetName(wb) ? 'v2' : 'v1';
}

async function readFileToWorkbook(file, opts) {
  const buf = await file.arrayBuffer();
  const readArr = (options) => {
    try {
      return XLSX.read(buf, { type: 'array', ...options });
    } catch (err) {
      const msg = typeof formatXlsxReadError === 'function' ? formatXlsxReadError(err) : String(err.message || err);
      throw new Error(msg);
    }
  };

  if (!opts?.lightweight) {
    return { wb: readArr(), buffer: buf };
  }

  const probe = readArr({ bookSheets: true });
  const allNames = probe.SheetNames || [];
  const toLoad = pickBillingSheets(allNames);
  const skipped = allNames.filter(n => !toLoad.includes(n));
  const wb = skipped.length ? readArr({ sheets: toLoad }) : readArr();
  return { wb, buffer: buf, sheetFilter: { loaded: toLoad, skipped } };
}

function mergeArmPack(prev, next) {
  if (!prev?.months?.length) return armNormalizePackMonths(next);
  if (!next?.months?.length) {
    return armNormalizePackMonths({
      ...prev,
      failedFiles: [...(prev.failedFiles || []), ...(next.failedFiles || [])]
    });
  }
  const replaceMes = new Set();
  const replaceFile = new Set();
  const prevByKey = new Map();
  prev.months.forEach(m => {
    const mk = String(m.mesKey || '').trim();
    if (/^\d{4}-\d{2}$/.test(mk)) prevByKey.set(mk, m);
  });
  next.months = (next.months || []).map(m => {
    const prevM = prevByKey.get(m.mesKey);
    if (prevM && armGetVendasLiq(prevM) && !armGetVendasLiq(m)) {
      return { ...m, vendasLiq: prevM.vendasLiq, vendasLiqUpdatedAt: prevM.vendasLiqUpdatedAt || '' };
    }
    return m;
  });
  next.months.forEach(m => {
    const mk = String(m.mesKey || '').trim();
    if (/^\d{4}-\d{2}$/.test(mk)) replaceMes.add(mk);
    const fn = String(m.fileName || '').trim().toLowerCase();
    if (fn) replaceFile.add(fn);
  });
  const months = [
    ...prev.months.filter(m => {
      const mk = String(m.mesKey || '').trim();
      const fn = String(m.fileName || '').trim().toLowerCase();
      if (mk && replaceMes.has(mk)) return false;
      if (fn && replaceFile.has(fn)) return false;
      return true;
    }),
    ...next.months
  ].sort((a, b) => String(a.mesKey).localeCompare(String(b.mesKey)));
  return armNormalizePackMonths({
    version: 3,
    updatedAt: new Date().toISOString(),
    months,
    failedFiles: [...(prev.failedFiles || []), ...(next.failedFiles || [])],
    customServices: [...new Set([...(prev.customServices || []), ...(next.customServices || [])])]
  });
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
  return armNormalizePackMonths({ version: 3, updatedAt: new Date().toISOString(), months, failedFiles });
}

function slimArmPackForPersist(pack) {
  if (!pack) return null;
  return {
    version: pack.version || 3,
    updatedAt: new Date().toISOString(),
    customServices: pack.customServices || [],
    months: (pack.months || []).map(m => {
      const n = armNormalizeMonth({ ...m });
      return {
      fileName: n.fileName, format: n.format, mesKey: n.mesKey, mesLabel: n.mesLabel,
      dataInicial: n.dataInicial, dataFinal: n.dataFinal,
      totalServicos: n.totalServicos, valorMinimo: n.valorMinimo,
      valorApurado: n.valorApurado, valorTotal: n.valorTotal, belowMin: n.belowMin,
      impostos: armNum(n.impostos),
      impostosManual: !!n.impostosManual,
      vendasLiq: armGetVendasLiq(n) || undefined,
      vendasLiqUpdatedAt: n.vendasLiqUpdatedAt || '',
      entrySource: n.entrySource || (n.format === 'manual' ? 'manual' : 'excel'),
      manualUpdatedAt: n.manualUpdatedAt || '',
      servicos: n.servicos, nfRows: n.nfRows, adicionais: n.adicionais,
      partialParse: n.partialParse, nfSkipped: n.nfSkipped, parseNote: n.parseNote
    };
    }),
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
      if (!pack.customServices) pack.customServices = [];
      pack.months.forEach(m => {
        m.impostos = armNum(m.impostos);
        m.nfRows = (m.nfRows || []).map(applySapToArmNf);
      });
      armNormalizePackMonths(pack);
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
      map[id].months.add(armMesLabel(m));
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

const ARM_MENSAL_COLS = 11;

function armEnsureTableCols(table, nCols) {
  if (!table || !nCols) return;
  table.dataset.managedSort = 'arm';
  delete table.dataset.domSortBound;
  let cg = table.querySelector('colgroup');
  if (!cg) {
    cg = document.createElement('colgroup');
    table.insertBefore(cg, table.firstChild);
  }
  while (cg.children.length < nCols) cg.appendChild(document.createElement('col'));
  while (cg.children.length > nCols) cg.removeChild(cg.lastChild);
}

function switchArmTab(tab) {
  armActiveTab = tab;
  document.querySelectorAll('#page-armazem .arm-tab').forEach(t => {
    t.classList.toggle('active', t.dataset.tab === tab);
  });
  ['lancamento', 'carregamento', 'mensal', 'resumo', 'nf', 'adicionais', 'catalogo'].forEach(id => {
    const el = $arm('tab-' + id);
    if (el) el.style.display = id === tab ? 'block' : 'none';
  });
  armRefreshWithSap().catch(e => console.warn('[armazem] refresh sap', e));
}

function renderArmActiveTab() {
  if (armActiveTab === 'lancamento') renderArmLancamento();
  if (armActiveTab === 'mensal') renderArmMensal();
  if (armActiveTab === 'resumo') renderArmAcumulado().catch(e => console.warn('[armazem] resumo', e));
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
    const lines = [];
    (armPack?.months || []).forEach(m => {
      let line = armMesLabel(m);
      if (m.fileName && m.fileName !== line) line += ` (${m.fileName})`;
      if (!(m.servicos?.length)) line += ' ⚠ sem resumo — processar de novo';
      if (m.partialParse) line += ' (resumo OK · NF omitido)';
      lines.push(line);
    });
    armPendingFiles.forEach(f => lines.push(f.name + ' (pendente — clicar Processar)'));
    (armPack?.failedFiles || []).forEach(f => {
      lines.push(`${f.fileName} (erro: ${f.error || 'desconhecido'})`);
    });
    if (lines.length) {
      list.style.display = 'block';
      list.innerHTML = lines.map(n => `<div>${armEsc(n)}</div>`).join('');
    } else {
      list.style.display = 'none';
      list.innerHTML = '';
    }
  }
  const sapNote = $arm('sapNote');
  if (sapNote) {
    sapNote.innerHTML = armSapStatusHtml();
    sapNote.style.display = 'block';
    sapNote.classList.toggle('data-warn', !sapApi()?.isSapLoaded?.());
  }
  const btn = $arm('procBtn');
  if (btn) btn.disabled = !armPendingFiles.length;
}

function armAggregateMensalTotals(monthRows) {
  const t = { vendasLiq: 0, armazenagem: 0, adicionais: 0, pago: 0, impostos: 0, pagoComImp: 0 };
  monthRows.forEach(m => {
    const vl = armGetVendasLiq(m);
    if (vl > 0) t.vendasLiq += vl;
    t.armazenagem += armGetNfFee(m);
    t.adicionais += armGetAdicionaisSum(m);
    t.pago += armGetMensalPagoSemImp(m);
    t.impostos += armGetMonthImpostos(m);
    t.pagoComImp += armGetMonthPagoComImp(m);
  });
  return t;
}

function buildArmMensalDisplayRows(sorted, groupByYear) {
  if (!groupByYear) {
    const out = sorted.map(m => ({ type: 'month', month: m }));
    if (sorted.length) {
      out.push({ type: 'total', label: 'Total geral', ...armAggregateMensalTotals(sorted) });
    }
    return out;
  }
  const out = [];
  const dated = sorted.filter(m => /^\d{4}-\d{2}$/.test(String(m.mesKey || '')));
  const other = sorted.filter(m => !/^\d{4}-\d{2}$/.test(String(m.mesKey || '')));
  let lastYear = null;
  let yearGroup = [];
  dated.forEach(m => {
    const year = String(m.mesKey).slice(0, 4);
    if (lastYear && year !== lastYear && yearGroup.length) {
      out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...armAggregateMensalTotals(yearGroup) });
      yearGroup = [];
    }
    if (year !== lastYear) {
      out.push({ type: 'year', label: year, year });
      lastYear = year;
    }
    out.push({ type: 'month', month: m });
    yearGroup.push(m);
  });
  if (yearGroup.length && lastYear) {
    out.push({ type: 'subtotal', label: `Total ${lastYear}`, year: lastYear, ...armAggregateMensalTotals(yearGroup) });
  }
  other.forEach(m => out.push({ type: 'month', month: m }));
  return out;
}

function armMensalTotalsRowHtml(label, t, trCls) {
  const pctArm = armFmtPctGasto(t.armazenagem, t.vendasLiq);
  const pctAdic = armFmtPctGasto(t.adicionais, t.vendasLiq);
  const pctVl = armFmtPctGasto(t.pago, t.vendasLiq);
  const pctVlCom = armFmtPctGasto(t.pagoComImp, t.vendasLiq);
  return `<tr class="${trCls}">
    <td><strong>${armEsc(label)}</strong></td>
    <td class="right">${t.vendasLiq > 0 ? armFmtMoney(t.vendasLiq) : '—'}</td>
    <td class="right">${t.armazenagem > 0 ? armFmtMoney(t.armazenagem) : '—'}</td>
    <td class="right" title="Armazenagem ÷ Vendas Liq">${pctArm}</td>
    <td class="right">${t.adicionais > 0 ? armFmtMoney(t.adicionais) : '—'}</td>
    <td class="right" title="Adicionais ÷ Vendas Liq">${pctAdic}</td>
    <td class="right">${t.pago > 0 ? armFmtMoney(t.pago) : '—'}</td>
    <td class="right" title="Pago s/ impostos ÷ Vendas Liq">${pctVl}</td>
    <td class="right">${t.impostos > 0 ? armFmtMoney(t.impostos) : '—'}</td>
    <td class="right">${t.pagoComImp > 0 ? armFmtMoney(t.pagoComImp) : '—'}</td>
    <td class="right" title="Total pago c/ impostos ÷ Vendas Liq">${pctVlCom}</td>
  </tr>`;
}

function armMensalMonthRowHtml(m) {
  const armFee = armGetNfFee(m);
  const addVal = armGetAdicionaisSum(m);
  const imp = armGetMonthImpostos(m);
  const pago = armGetMensalPagoSemImp(m);
  const pagoComImp = armGetMonthPagoComImp(m);
  const vendasLiq = armGetVendasLiq(m);
  const pctArm = armFmtPctGasto(armFee, vendasLiq);
  const pctAdic = armFmtPctGasto(addVal, vendasLiq);
  const pctVl = armFmtPctGasto(pago, vendasLiq);
  const pctVlCom = armFmtPctGasto(pagoComImp, vendasLiq);
  const partial = m.partialParse ? ' <span class="badge b-warn" title="' + armEsc(m.parseNote || '') + '">NF omitido</span>' : '';
  const vlDisplay = vendasLiq ? fmtArmNum(vendasLiq, 2) : '';
  return `<tr class="arm-mensal-month-row">
    <td>${armEsc(armMesLabel(m))}${partial}</td>
    <td class="right"><input type="text" class="fi arm-vendas-liq-input" data-mes="${armEsc(m.mesKey)}" value="${armEsc(vlDisplay)}" placeholder="R$ …" onblur="typeof armSetVendasLiq==='function'&&armSetVendasLiq('${armEsc(m.mesKey)}',this.value)" style="width:110px;text-align:right;font-size:11px;padding:4px 6px;"></td>
    <td class="right" title="Taxa 5,5% sobre NF expedida">${armFee > 0 ? armFmtMoney(armFee) : '—'}</td>
    <td class="right" title="Armazenagem ÷ Vendas Liq">${pctArm}</td>
    <td class="right">${addVal > 0 ? armFmtMoney(addVal) : '—'}</td>
    <td class="right" title="Adicionais ÷ Vendas Liq">${pctAdic}</td>
    <td class="right">${pago > 0 ? armFmtMoney(pago) : '—'}</td>
    <td class="right" title="Pago s/ impostos ÷ Vendas Liq">${pctVl}</td>
    <td class="right">${imp > 0 ? armFmtMoney(imp) : '—'}</td>
    <td class="right">${pagoComImp > 0 ? armFmtMoney(pagoComImp) : '—'}</td>
    <td class="right" title="Total pago c/ impostos ÷ Vendas Liq">${pctVlCom}</td>
  </tr>`;
}

function armMensalExportRowFromMonth(m) {
  const armFee = armGetNfFee(m);
  const adicionais = armGetAdicionaisSum(m);
  const pago = armGetMensalPagoSemImp(m);
  const vendasLiq = armGetVendasLiq(m);
  const pagoComImp = armGetMonthPagoComImp(m);
  return {
    Mes: armMesLabel(m),
    VendasLiq: vendasLiq || '',
    Armazenagem55: armFee,
    PctArmazenagemSobreVendasLiq: vendasLiq > 0 ? armFee / vendasLiq : '',
    Adicionais: adicionais,
    PctAdicionaisSobreVendasLiq: vendasLiq > 0 ? adicionais / vendasLiq : '',
    PagoSemImpostos: pago,
    PctSobreVendasLiq: vendasLiq > 0 ? pago / vendasLiq : '',
    Impostos: armGetMonthImpostos(m),
    TotalComImpostos: pagoComImp,
    PctComImpSobreVendasLiq: vendasLiq > 0 ? pagoComImp / vendasLiq : ''
  };
}

function armMensalExportRowFromTotals(t, label) {
  return {
    Mes: label,
    VendasLiq: t.vendasLiq || '',
    Armazenagem55: t.armazenagem || '',
    PctArmazenagemSobreVendasLiq: t.vendasLiq > 0 ? t.armazenagem / t.vendasLiq : '',
    Adicionais: t.adicionais,
    PctAdicionaisSobreVendasLiq: t.vendasLiq > 0 ? t.adicionais / t.vendasLiq : '',
    PagoSemImpostos: t.pago,
    PctSobreVendasLiq: t.vendasLiq > 0 ? t.pago / t.vendasLiq : '',
    Impostos: t.impostos,
    TotalComImpostos: t.pagoComImp,
    PctComImpSobreVendasLiq: t.vendasLiq > 0 ? t.pagoComImp / t.vendasLiq : ''
  };
}

function buildArmMensalExportRows(months) {
  const sorted = [...months].sort((a, b) => String(a.mesKey).localeCompare(String(b.mesKey)));
  const out = [];
  const dated = sorted.filter(m => /^\d{4}-\d{2}$/.test(String(m.mesKey || '')));
  const other = sorted.filter(m => !/^\d{4}-\d{2}$/.test(String(m.mesKey || '')));
  let lastYear = null;
  let yearGroup = [];
  dated.forEach(m => {
    const year = String(m.mesKey).slice(0, 4);
    if (lastYear && year !== lastYear && yearGroup.length) {
      out.push(armMensalExportRowFromTotals(armAggregateMensalTotals(yearGroup), `Total ${lastYear}`));
      yearGroup = [];
    }
    out.push(armMensalExportRowFromMonth(m));
    yearGroup.push(m);
    lastYear = year;
  });
  if (yearGroup.length && lastYear) {
    out.push(armMensalExportRowFromTotals(armAggregateMensalTotals(yearGroup), `Total ${lastYear}`));
  }
  other.forEach(m => out.push(armMensalExportRowFromMonth(m)));
  return out;
}

function renderArmMensal() {
  const empty = $arm('mensalEmpty');
  const content = $arm('mensalContent');
  const months = armNormalizePackMonths({ months: armPack?.months || [] }).months;
  if (!months.length) {
    if (empty) empty.style.display = 'block';
    if (content) content.style.display = 'none';
    return;
  }
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'block';

  const totalPago = months.reduce((s, m) => s + armGetMensalPagoSemImp(m), 0);
  const totalComImp = months.reduce((s, m) => s + armGetMonthPagoComImp(m), 0);
  const totalAdicionais = months.reduce((s, m) => s + armGetAdicionaisSum(m), 0);
  const totalArmFee = months.reduce((s, m) => s + armGetNfFee(m), 0);
  const kpis = $arm('mensalKpis');
  if (kpis) {
    kpis.innerHTML = `
      <div class="kpi"><div class="label">Meses</div><div class="value">${months.length}</div></div>
      <div class="kpi"><div class="label">Total pago (s/ imp.)</div><div class="value">${armFmtMoney(totalPago)}</div></div>
      <div class="kpi"><div class="label">Total pago (c/ imp.)</div><div class="value">${armFmtMoney(totalComImp)}</div></div>
      <div class="kpi"><div class="label">Armazenagem (5,5%)</div><div class="value">${armFmtMoney(totalArmFee)}</div><div class="sub">Taxa NF expedida</div></div>
      <div class="kpi"><div class="label">Adicionais</div><div class="value">${armFmtMoney(totalAdicionais)}</div></div>`;
  }

  const mensalSort = armSorts.mensal;
  const groupByYear = !mensalSort || mensalSort.col === 'mesKey' || mensalSort.col === 'mesLabel';
  const rows = armApplySort(months, 'mensal', {
    mesKey: r => r.mesKey,
    mesLabel: r => armMesLabel(r),
    vendasLiq: r => armGetVendasLiq(r) ?? -1,
    armazenagemFee: r => armGetNfFee(r),
    pctArmazenagemVl: r => {
      const vl = armGetVendasLiq(r);
      const a = armGetNfFee(r);
      return vl > 0 && a >= 0 ? a / vl : -1;
    },
    adicionais: r => armGetAdicionaisSum(r),
    pctAdicionaisVl: r => {
      const vl = armGetVendasLiq(r);
      const a = armGetAdicionaisSum(r);
      return vl > 0 && a >= 0 ? a / vl : -1;
    },
    pagoSemImp: r => armGetMensalPagoSemImp(r),
    pctVendasLiq: r => {
      const vl = armGetVendasLiq(r);
      const p = armGetMensalPagoSemImp(r);
      return vl > 0 && p >= 0 ? p / vl : -1;
    },
    impostos: r => armGetMonthImpostos(r),
    pagoComImp: r => armGetMonthPagoComImp(r),
    pctVendasLiqCom: r => {
      const vl = armGetVendasLiq(r);
      const t = armGetMonthPagoComImp(r);
      return vl > 0 && t >= 0 ? t / vl : -1;
    }
  });
  const displayRows = buildArmMensalDisplayRows(rows, groupByYear);

  const body = $arm('mensalBody');
  const thead = $arm('mensalHead');
  const table = thead?.closest('table') || body?.closest('table');
  if (table) armEnsureTableCols(table, ARM_MENSAL_COLS);

  if (thead) {
    thead.innerHTML = `
      ${sortTh('mensal', 'mesLabel', 'Mês')}
      ${sortTh('mensal', 'vendasLiq', 'Vendas Liq', 'right')}
      ${sortTh('mensal', 'armazenagemFee', 'Armazenagem (5,5%)', 'right')}
      ${sortTh('mensal', 'pctArmazenagemVl', '% arm. s/ vendas liq', 'right')}
      ${sortTh('mensal', 'adicionais', 'Adicionais', 'right')}
      ${sortTh('mensal', 'pctAdicionaisVl', '% adic. s/ vendas liq', 'right')}
      ${sortTh('mensal', 'pagoSemImp', 'Pago (s/ imp.)', 'right')}
      ${sortTh('mensal', 'pctVendasLiq', '% s/ vendas liq', 'right')}
      ${sortTh('mensal', 'impostos', 'Impostos', 'right')}
      ${sortTh('mensal', 'pagoComImp', 'Total pago (c/ imp.)', 'right')}
      ${sortTh('mensal', 'pctVendasLiqCom', '% c/ imp. s/ vendas liq', 'right')}`;
  }

  if (body) {
    body.innerHTML = displayRows.map(row => {
      if (row.type === 'year') {
        return `<tr class="month-year-row"><td colspan="${ARM_MENSAL_COLS}"><strong>${armEsc(row.label)}</strong></td></tr>`;
      }
      if (row.type === 'subtotal') {
        return armMensalTotalsRowHtml(row.label, row, 'month-subtotal-row');
      }
      if (row.type === 'total') {
        return armMensalTotalsRowHtml(row.label, row, 'month-total-row');
      }
      return armMensalMonthRowHtml(row.month);
    }).join('');
  }

  if (typeof scheduleTableSort === 'function') scheduleTableSort();
}

async function renderArmAcumulado() {
  const empty = $arm('acumuladoEmpty');
  const content = $arm('acumuladoContent');
  const sections = $arm('acumuladoSections');
  const months = armNormalizePackMonths({ months: armPack?.months || [] }).months;
  if (!months.length) {
    if (empty) empty.style.display = 'block';
    if (content) content.style.display = 'none';
    return;
  }
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'block';

  if (sections) sections.innerHTML = renderArmResumoSections(months);
  if (typeof scheduleColResize === 'function') scheduleColResize();
}

function renderArmNf() {
  const empty = $arm('nfEmpty');
  const content = $arm('nfContent');
  const months = armPack?.months || [];
  const allNf = months.flatMap(m => (m.nfRows || []).map(r => ({ ...r, mesKey: m.mesKey, mesLabel: armMesLabel(m) })));
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
  const notLoaded = allNf.some(r => r.sapNotLoaded);
  const missing = sapLoaded ? allNf.filter(r => r.sapMissing).length : 0;
  const valDiff = allNf.filter(r => r.valorDiff).length;

  const kpis = $arm('nfKpis');
  if (kpis) {
    kpis.innerHTML = `
      <div class="kpi"><div class="label">Linhas NF</div><div class="value">${allNf.length}</div></div>
      <div class="kpi${notLoaded ? ' flag' : ''}"><div class="label">SAP ZFACT</div><div class="value">${sapLoaded ? 'Carregado' : '—'}</div>${notLoaded ? '<div class="sub">Carrega em Fretes CT-e</div>' : ''}</div>
      <div class="kpi${missing ? ' flag' : ''}"><div class="label">NF ausente SAP</div><div class="value">${missing}</div></div>
      <div class="kpi${valDiff ? ' flag' : ''}"><div class="label">Δ valor SAP</div><div class="value">${valDiff}</div></div>`;
  }

  const mesSel = $arm('nfFilterMes');
  if (mesSel && !mesSel._armBound) {
    mesSel._armBound = true;
    const opts = ['<option value="">Todos os meses</option>']
      .concat(months.map(m => `<option value="${armEsc(m.mesKey)}">${armEsc(armMesLabel(m))}</option>`));
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
      ${sortTh('nf', 'valorNF', 'Valor Unilog', 'right')}
      ${sortTh('nf', 'sapValor', 'Valor ZFACT/SAP', 'right')}
      <th>Δ valor</th>
      <th>Cliente SAP</th>
      ${sortTh('nf', 'fee', 'Taxa 5,5%', 'right')}
      <th>Estado</th>`;
  }

  const body = $arm('nfBody');
  if (body) {
    body.innerHTML = rows.map(r => {
      const diff = r.sapFound ? armNum(r.valorNF) - armNum(r.sapValor) : null;
      let cls = '';
      if (r.sapNotLoaded) cls = 'nf-sap-pending';
      else if (r.sapMissing) cls = 'nf-sap-missing';
      else if (r.valorDiff) cls = 'nf-val-mismatch';
      const estado = armNfEstadoBadge(r);
      return `<tr class="${cls}">
        <td>${armEsc(r.mesLabel)}</td>
        <td>${armEsc(r.nf)}</td>
        <td>${armEsc(r.data)}</td>
        <td class="right">${armFmtMoney(r.valorNF)}</td>
        <td class="right">${r.sapFound ? armFmtMoney(r.sapValor) : '—'}</td>
        <td class="right">${diff != null && Math.abs(diff) > 0.02 ? armFmtMoney(diff) : '—'}</td>
        <td>${armEsc(r.sapCliente || '—')}</td>
        <td class="right">${armFmtMoney(r.fee)}</td>
        <td>${estado}</td>
      </tr>`;
    }).join('');
  }
}

function armMesShortLabel(mesKey) {
  const m = String(mesKey || '').match(/^(\d{4})-(\d{2})$/);
  if (!m) return mesKey || '—';
  const mi = parseInt(m[2], 10);
  return (ARM_MESES_NOME[mi] || m[2]).slice(0, 3);
}

function armAdicionaisPivotServices(months) {
  const byNorm = new Map();
  months.forEach(m => {
    (m.adicionais || []).forEach(a => {
      const norm = a.normName || normalizeServicoName(a.rawName);
      if (!byNorm.has(norm)) {
        byNorm.set(norm, {
          normName: norm,
          displayLabel: armServicoDisplayLabel(norm),
          byMes: {}
        });
      }
      const svc = byNorm.get(norm);
      svc.byMes[m.mesKey] = (svc.byMes[m.mesKey] || 0) + armNum(a.valor);
    });
  });
  return byNorm;
}

function renderArmAdicionaisMatrixHtml(year, mesKeys, monthsByKey, services) {
  const monthCells = mesKeys.map(mk => {
    const m = monthsByKey.get(mk);
    const title = armEsc(armMesNomeLong(mk, m ? armMesLabel(m) : ''));
    return `<th class="right" title="${title}">${armEsc(armMesShortLabel(mk))}</th>`;
  }).join('');

  const serviceRows = services.map(svc => {
    let rowTotal = 0;
    const cells = mesKeys.map(mk => {
      const v = svc.byMes[mk] || 0;
      if (v > 0) rowTotal += v;
      return `<td class="right">${v > 0 ? armFmtMoney(v) : '—'}</td>`;
    }).join('');
    return `<tr>
      <td>${armEsc(svc.displayLabel)}</td>
      ${cells}
      <td class="right arm-ad-subtotal-col">${rowTotal > 0 ? armFmtMoney(rowTotal) : '—'}</td>
    </tr>`;
  }).join('');

  let vlYearTotal = 0;
  const vlCells = mesKeys.map(mk => {
    const vl = armGetVendasLiq(monthsByKey.get(mk));
    if (vl) vlYearTotal += vl;
    return `<td class="right">${vl ? armFmtMoney(vl) : '—'}</td>`;
  }).join('');

  let grandTotal = 0;
  const totalCells = mesKeys.map(mk => {
    let colTotal = 0;
    services.forEach(svc => { colTotal += svc.byMes[mk] || 0; });
    grandTotal += colTotal;
    return `<td class="right">${colTotal > 0 ? armFmtMoney(colTotal) : '—'}</td>`;
  }).join('');

  return `<div class="arm-adicionais-year arm-resumo-year">
    <div class="section-title arm-resumo-year-title">${armEsc(year)}</div>
    <div class="tbl-wrap">
      <table class="arm-adicionais-matrix" data-managed-sort="1">
        <thead><tr>
          <th>Serviço</th>
          ${monthCells}
          <th class="right arm-ad-subtotal-col">Total</th>
        </tr></thead>
        <tbody>
          ${serviceRows}
          <tr class="arm-ad-vl-row">
            <td>Vendas Liq</td>
            ${vlCells}
            <td class="right arm-ad-subtotal-col">${vlYearTotal > 0 ? armFmtMoney(vlYearTotal) : '—'}</td>
          </tr>
          <tr class="arm-ad-total-row">
            <td>Total adicionais</td>
            ${totalCells}
            <td class="right arm-ad-subtotal-col">${grandTotal > 0 ? armFmtMoney(grandTotal) : '—'}</td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>`;
}

function renderArmAdicionais() {
  const empty = $arm('adicionaisEmpty');
  const content = $arm('adicionaisContent');
  const sections = $arm('adicionaisSections');
  const months = armPack?.months || [];
  const hasAdicionais = months.some(m => (m.adicionais || []).length > 0);
  if (!hasAdicionais) {
    if (empty) empty.style.display = 'block';
    if (content) content.style.display = 'none';
    return;
  }
  if (empty) empty.style.display = 'none';
  if (content) content.style.display = 'block';

  const serviceMap = armAdicionaisPivotServices(months);
  const byYear = {};
  months.forEach(m => {
    const y = (m.mesKey || '').slice(0, 4);
    if (!/^\d{4}$/.test(y)) return;
    if (!byYear[y]) byYear[y] = { mesKeys: new Set(), monthsByKey: new Map() };
    byYear[y].mesKeys.add(m.mesKey);
    byYear[y].monthsByKey.set(m.mesKey, m);
  });

  const html = Object.keys(byYear).sort().map(year => {
    const { mesKeys, monthsByKey } = byYear[year];
    const sortedMesKeys = [...mesKeys].sort();
    const yearServices = [...serviceMap.values()]
      .filter(s => sortedMesKeys.some(mk => (s.byMes[mk] || 0) > 0))
      .sort((a, b) => a.displayLabel.localeCompare(b.displayLabel, 'pt'));
    if (!yearServices.length) return '';
    return renderArmAdicionaisMatrixHtml(year, sortedMesKeys, monthsByKey, yearServices);
  }).filter(Boolean).join('');

  if (sections) sections.innerHTML = html;
  if (typeof scheduleColResize === 'function') scheduleColResize();
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
  const loadBtn = $arm('loadBtn');
  const note = $arm('procNote');
  if (spin) {
    spin.hidden = !active;
    spin.style.display = active ? 'inline-flex' : 'none';
  }
  if (btn) btn.disabled = active || !armPendingFiles.length;
  if (loadBtn) loadBtn.disabled = !!active;
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
  if (size > ARM_ABSOLUTE_MAX_FILE_BYTES) {
    throw new Error(
      `Ficheiro demasiado grande (${armFmtBytes(size)}). Limite absoluto ${armFmtBytes(ARM_ABSOLUTE_MAX_FILE_BYTES)}.`
    );
  }
  const lightweight = size >= ARM_MAX_FILE_BYTES;
  const timeoutMs = lightweight ? ARM_LARGE_PARSE_TIMEOUT_MS : ARM_PARSE_TIMEOUT_MS;
  const { wb, sheetFilter } = await armWithTimeout(
    readFileToWorkbook(file, lightweight ? { lightweight: true } : {}),
    timeoutMs,
    file.name
  );
  await armYield();
  const parseOpts = lightweight ? { skipNfDetail: true } : {};
  const record = await armWithTimeout(
    Promise.resolve(parseArmazemWorkbook(wb, file.name, parseOpts)),
    timeoutMs,
    file.name
  );
  if (lightweight) {
    const skipped = sheetFilter?.skipped?.length ? sheetFilter.skipped.join(', ') : 'folhas extra';
    const monthLbl = armShortMonthLabel(file.name, record.mesKey, record.mesLabel);
    record.partialParse = true;
    record.nfSkipped = true;
    record.parseNote = `Resumo OK; NF detalhe omitido (${armFmtBytes(size)}; ignorado: ${skipped})`;
    record._partialToast = `${monthLbl}: resumo OK, NF detalhe omitido (ficheiro grande).`;
  }
  if (!record.servicos?.length) {
    let hint = armGuessWrongFileError(wb, file.name);
    if (!hint && record.format === 'v2') {
      hint = 'Folha Resumo encontrada mas sem linhas válidas — confirma colunas Qtde / Valor Calculado no Excel Unilog.';
    }
    throw new Error(hint || 'Nenhum serviço no bloco Resumo — verifica o ficheiro ou formato.');
  }
  armLogParsedMonth(record, file.name);
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
      const last = results[results.length - 1];
      if (last?.record?._partialToast) {
        armToast(last.record._partialToast, 'info');
        delete last.record._partialToast;
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
    armLogPackMonths('after process');
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
  if (!silent) armSetProcessing(true, 'A carregar da cloud…');
  try {
    const m = await Promise.race([
      fetchExcelFiles([armSlot()]),
      new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), 25_000))
    ]);
    const rec = m[armSlot()];
    const pack = parseArmPackFromRec(rec);
    if (pack?.months?.length) {
      armPack = armNormalizePackMonths(pack);
      refreshAllSapOnNfs();
      armLogPackMonths('loaded from cloud');
      updateArmFileZone();
      if (armInited) renderArmActiveTab();
      const stale = (armPack.months || []).filter(m => !(m.servicos?.length)).length;
      if (!silent) {
        armToast(
          `${armPack.months.length} mês(es) carregado(s) da cloud` +
          (stale ? ` — ${stale} sem resumo (processar de novo)` : '') + '.',
          stale ? 'error' : 'success'
        );
      }
      return true;
    }
  } catch (e) {
    console.warn('[armazem] load saved', e);
  } finally {
    if (!silent) armSetProcessing(false);
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
    const aoa = exportResumoSheetRows(rows, 'Resumo — ' + armMesLabel(m));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
  });

  const agg = aggregateResumoByNorm(months);
  const totalAoa = exportResumoSheetRows(
    agg.map(r => ({ ...r, rawName: r.normName })),
    'Total — todos os meses'
  );
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(totalAoa), 'Total');

  const mensal = buildArmMensalExportRows(months);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mensal), 'Visão mensal');

  const nfs = months.flatMap(m => (m.nfRows || []).map(r => ({
    Mes: armMesLabel(m), NF: r.nf, Data: r.data,
    ValorNF_Unilog: r.valorNF, ValorNF_SAP: r.sapValor ?? '',
    Taxa: r.fee, DeltaTaxa: r.feeDelta,
    SAP_OK: r.sapFound ? 'SIM' : (r.sapMissing ? 'AUSENTE' : '')
  })));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(nfs), 'NF 5.5%');

  const ad = months.flatMap(m => (m.adicionais || []).map(a => ({
    Mes: armMesLabel(m), Servico: a.normName, Valor: a.valor, Qtde: a.qtde
  })));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ad), 'Adicionais');

  XLSX.writeFile(wb, `armazem_delta_${armCompany()}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  armToast('Excel exportado.');
}

async function reloadArmazemForCompany() {
  armPack = null;
  armPendingFiles = [];
  armLancamentoDraft = null;
  armNfUploadRows = [];
  armInvalidateVendasCache();
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
    armRefreshWithSap().catch(e => console.warn('[armazem] refresh sap', e));
    return;
  }
  armInited = true;
  loadCatalogOverrides();

  document.getElementById('page-armazem')?.addEventListener('click', e => {
    if (e.target?.closest?.('.arm-link-fretes')) armGotoFretesPage(e);
  });

  document.querySelectorAll('#page-armazem .arm-tab').forEach(t => {
    t.addEventListener('click', () => switchArmTab(t.dataset.tab));
  });

  $arm('lancLoadBtn')?.addEventListener('click', () => armLoadLancamentoDraft());
  $arm('lancAddRowBtn')?.addEventListener('click', () => armLancamentoAddRow());
  $arm('lancAddCustomBtn')?.addEventListener('click', () => armLancamentoAddCustomService());
  $arm('lancSaveBtn')?.addEventListener('click', () => armSaveLancamento());
  $arm('lancImpostos')?.addEventListener('input', () => {
    if (armLancamentoDraft) {
      armLancamentoDraft.impostosManual = true;
      armLancamentoDraft.impostos = armNum($arm('lancImpostos')?.value);
      armRefreshLancamentoSummary();
    }
  });
  $arm('lancImpostosAutoBtn')?.addEventListener('click', () => armLancamentoResetImpostosAuto());

  const nfZone = $arm('nfUploadZone');
  const nfInput = $arm('nfUploadInput');
  if (nfZone && nfInput) {
    nfZone.addEventListener('dragover', e => { e.preventDefault(); nfZone.classList.add('drag'); });
    nfZone.addEventListener('dragleave', () => nfZone.classList.remove('drag'));
    nfZone.addEventListener('drop', e => {
      e.preventDefault(); nfZone.classList.remove('drag');
      const f = e.dataTransfer.files?.[0];
      if (f) armProcessNfUpload(f);
    });
    nfInput.addEventListener('change', e => {
      const f = e.target.files?.[0];
      if (f) armProcessNfUpload(f);
      e.target.value = '';
    });
  }

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
  Promise.all([
    loadSavedArmazem(true),
    armEnsureSapLoaded()
  ]).then(() => {
    refreshAllSapOnNfs();
    if (armActiveTab === 'lancamento') armLoadLancamentoDraft();
    else renderArmActiveTab();
  }).catch(e => console.warn('[armazem] load saved', e));
}

window.initArmazem = initArmazem;
window.refreshArmazemSapValidation = refreshArmazemSapValidation;
window.armSetVendasLiq = armSetVendasLiq;
window.armEnsureSapLoaded = armEnsureSapLoaded;
window.armOnFileSelected = armOnFileSelected;
window.processArmFiles = processArmFiles;
window.armDoSort = armDoSort;
window.armSetCatalog = armSetCatalog;
window.exportArmazemWorkbook = exportArmazemWorkbook;
window.armInvalidateVendasCache = armInvalidateVendasCache;
window.armLancamentoRowChange = armLancamentoRowChange;
window.armLancamentoRowPreview = armLancamentoRowPreview;
window.armLancamentoRowCommit = armLancamentoRowCommit;
window.armLancamentoServicoChange = armLancamentoServicoChange;
window.armLancamentoRemoveRow = armLancamentoRemoveRow;
window.armLancamentoAddRow = armLancamentoAddRow;
window.armLancamentoAddCustomService = armLancamentoAddCustomService;
window.armSaveLancamento = armSaveLancamento;
window.armLoadLancamentoDraft = armLoadLancamentoDraft;
window.armLancamentoResetImpostosAuto = armLancamentoResetImpostosAuto;
window.armProcessNfUpload = armProcessNfUpload;
