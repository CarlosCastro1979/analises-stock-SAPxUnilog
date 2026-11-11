// total-brasil.js v1.0.0 — Total Brasil (DFB + QB) em Cobranças UNILOG
const TOTAL_BRASIL_JS_VERSION = '1.0.0';

const TB_COMPANIES = ['DFB', 'QB'];
/** Unilog report operator of each company's quinzenal pack. */
const TB_OPERATOR = { DFB: 'DELTA', QB: 'QBRASIL' };
const TB_AMOUNTS = ['vendasBrutas', 'vendasLiq', 'armComImp', 'freteB2b', 'freteB2c'];
const TB_QB_CACHE_MS = 120_000;

/** QB rows read read-only from the cloud: { co, at, rows, armMonths, qzFiles, error }. */
let tbQbCache = null;
let tbQbPromise = null;
let tbLoading = false;
let tbLastModel = null;
let tbSort = { col: 'mesKey', dir: 1 };
let tbCoSort = { col: 'co', dir: 1 };

function tbYield() {
  return new Promise(resolve => {
    try { requestAnimationFrame(() => setTimeout(resolve, 0)); } catch (_) { setTimeout(resolve, 0); }
  });
}

function tbCompany() {
  return typeof company !== 'undefined' ? company : 'DFB';
}

function tbIsActive() {
  return !!document.getElementById('page-totalBrasil')?.classList.contains('active');
}

function tbMoney(v) {
  return typeof cuFmtMoney === 'function' ? cuFmtMoney(v) : 'R$ ' + (Number(v) || 0).toFixed(2);
}

function tbPct(num, den) {
  if (!(den > 0) || !Number.isFinite(num)) return '—';
  return typeof cuFmtPct === 'function' ? cuFmtPct(num, den) : ((num / den) * 100).toFixed(2) + '%';
}

function tbEsc(s) {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function tbMesLabel(mesKey) {
  return typeof cuMesLabel === 'function' ? cuMesLabel(mesKey) : mesKey;
}

/** Same derived columns as Total Custo Unilog — percentages always from the (summed) amounts. */
function tbDerive(r) {
  const out = { ...r };
  TB_AMOUNTS.forEach(k => { out[k] = Number(out[k]) || 0; });
  out.total = out.armComImp + out.freteB2b;
  out.totalComB2c = out.total + out.freteB2c;
  out.pct = out.vendasLiq > 0 ? out.total / out.vendasLiq : null;
  out.pctVendasBrutas = out.vendasBrutas > 0 ? out.total / out.vendasBrutas : null;
  out.pctComB2c = out.vendasLiq > 0 ? out.totalComB2c / out.vendasLiq : null;
  return out;
}

function tbSum(rows) {
  const t = {};
  TB_AMOUNTS.forEach(k => { t[k] = rows.reduce((s, r) => s + (Number(r?.[k]) || 0), 0); });
  return tbDerive(t);
}

function tbHasAny(r) {
  return !!r && TB_AMOUNTS.some(k => (Number(r[k]) || 0) > 0);
}

async function tbFetchRec(co, slot) {
  const { data, error } = await db.from('logistica_excel_files')
    .select('slot,file_name,uploaded_at,file_data').eq('company', co).eq('slot', slot);
  if (error) throw error;
  return (data || [])[0] || null;
}

function tbDecodeJson(rec) {
  if (!rec?.file_data) return null;
  const b64 = rec.file_data;
  rec.file_data = null;
  return JSON.parse(new TextDecoder().decode(base64ToArrayBuffer(b64)));
}

/**
 * Monthly fretes of a stored quinzenal pack, as Total Custo Unilog reads them:
 * B2B = Σ pago per NF of the quinzenal B2B rows (b2bMonthTotals.totalPagoQz), B2C = b2cMonthTotals.totalFrete.
 * Only rows of the company's own operator; the pack is a private copy (never persisted).
 */
function tbFretesRowsFromQzPack(pack, operator) {
  if (!pack) return [];
  delete pack.fileBinaries;
  if (typeof cleanQuinzenalPackRows === 'function') cleanQuinzenalPackRows(pack);
  const own = x => qzOperatorOf(x) === operator;
  const byMes = {};
  const row = k => (byMes[k] = byMes[k] || { mesKey: k, freteB2b: 0, freteB2c: 0, byNf: {} });
  (pack.b2bRows || []).filter(own).forEach(r => {
    const k = r.mesKey || mesKeyFromQuinzenaKey(r.quinzenaKey);
    const nfKey = r.nfKey || normNFKey(r.nf);
    if (!k || !nfKey) return;
    const m = row(k);
    if (!m.byNf[nfKey]) m.byNf[nfKey] = { valorNF: 0, pago: 0, nCte: 0 };
    mergeQzB2BNf(m.byNf[nfKey], r);
  });
  const b2c = (pack.b2cRows || []).filter(own);
  b2c.forEach(r => enrichB2cRowMesFields(r));
  buildB2CMonthTotals(b2c).forEach(t => { row(t.mesKey).freteB2c = Number(t.totalFrete) || 0; });
  Object.values(byMes).forEach(m => {
    m.freteB2b = Object.values(m.byNf).reduce((s, x) => s + (Number(x.pago) || 0), 0);
    delete m.byNf;
  });
  return Object.values(byMes).filter(m => m.mesKey && m.mesKey !== 'sem-mes');
}

/** Read another company's Armazém + quinzenal packs from the cloud — the selected company is not touched. */
async function tbLoadCompanyReadOnly(co) {
  const out = { co, at: Date.now(), rows: [], armMonths: 0, qzFiles: 0, error: '' };
  const errors = [];
  const byMes = {};
  const get = k => (byMes[k] = byMes[k] || { mesKey: k });

  try {
    const { pack } = await armFetchPackForCompany(co);
    await tbYield();
    const armRows = pack ? armMonthlyCustoRowsFromPack(pack) : [];
    out.armMonths = armRows.length;
    out.nfRate = pack?.contract?.nfRate || null;
    armRows.forEach(r => {
      const m = get(r.mesKey);
      m.vendasBrutas = r.vendasBrutas || 0;
      m.vendasLiq = r.vendasLiq || 0;
      m.armComImp = r.armComImp || 0;
    });
  } catch (e) {
    console.warn('[totalBrasil] armazem', co, e);
    errors.push(`Armazém ${co}: ${e.message || e}`);
  }

  try {
    const rec = await tbFetchRec(co, typeof EXCEL_SLOTS !== 'undefined' ? EXCEL_SLOTS.FRETES_QUINZENAL : 'fretes_quinzenal');
    await tbYield();
    const pack = tbDecodeJson(rec);
    await tbYield();
    out.qzFiles = pack?.files?.length || 0;
    tbFretesRowsFromQzPack(pack, TB_OPERATOR[co]).forEach(r => {
      const m = get(r.mesKey);
      m.freteB2b = r.freteB2b;
      m.freteB2c = r.freteB2c;
    });
  } catch (e) {
    console.warn('[totalBrasil] fretes', co, e);
    errors.push(`Fretes ${co}: ${e.message || e}`);
  }

  out.rows = Object.values(byMes).map(tbDerive).filter(tbHasAny)
    .sort((a, b) => String(a.mesKey).localeCompare(String(b.mesKey)));
  out.error = errors.join(' · ');
  console.log('[totalBrasil] read-only', co, out.rows.length, 'months', { armMonths: out.armMonths, qzFiles: out.qzFiles, error: out.error });
  return out;
}

function tbEnsureQb(force) {
  if (!force && tbQbCache && Date.now() - tbQbCache.at < TB_QB_CACHE_MS) return Promise.resolve(tbQbCache);
  if (tbQbPromise) return tbQbPromise;
  tbQbPromise = tbLoadCompanyReadOnly('QB')
    .then(r => { tbQbCache = r; return r; })
    .finally(() => { tbQbPromise = null; });
  return tbQbPromise;
}

/** DFB = the selected company — same in-memory data as its Total Custo Unilog. */
async function tbEnsureDfb() {
  const tasks = [];
  if (typeof fteNeedsCloudReload === 'function' && fteNeedsCloudReload() && typeof loadSavedFretesFiles === 'function') {
    tasks.push(loadSavedFretesFiles(true).catch(e => console.warn('[totalBrasil] fretes DFB', e)));
  }
  if (typeof armEnsurePackLoaded === 'function' && !(typeof armIsCompanyPackLoaded === 'function' && armIsCompanyPackLoaded('DFB'))) {
    tasks.push(armEnsurePackLoaded('DFB').catch(e => console.warn('[totalBrasil] armazem DFB', e)));
  }
  await Promise.all(tasks);
}

function tbDfbRows() {
  const rows = typeof buildCustoUnilogRows === 'function' ? buildCustoUnilogRows() : [];
  return rows.filter(r => r.mesKey && r.mesKey !== 'sem-mes').map(tbDerive).filter(tbHasAny);
}

function tbMissingLabel(co, r) {
  if (!tbHasAny(r)) return `${co} sem dados`;
  const miss = [];
  if (!(r.armComImp > 0)) miss.push('armazém');
  if (!(r.freteB2b > 0)) miss.push('B2B');
  return miss.length ? `${co} s/ ${miss.join(' e ')}` : '';
}

function tbBuildModel() {
  const dfbRows = tbDfbRows();
  const qbRows = tbQbCache?.rows || [];
  const byCo = { DFB: new Map(dfbRows.map(r => [r.mesKey, r])), QB: new Map(qbRows.map(r => [r.mesKey, r])) };
  const keys = [...new Set([...byCo.DFB.keys(), ...byCo.QB.keys()])].sort();
  const months = keys.map(k => {
    const per = {};
    TB_COMPANIES.forEach(co => { per[co] = byCo[co].get(k) || tbDerive({ mesKey: k }); });
    const total = tbSum(TB_COMPANIES.map(co => per[co]));
    const missing = TB_COMPANIES.map(co => tbMissingLabel(co, per[co])).filter(Boolean);
    return { mesKey: k, mesLabel: tbMesLabel(k), total, per, missing };
  });
  return { months, dfbCount: dfbRows.length, qbCount: qbRows.length };
}

function tbFilteredMonths(model, year) {
  return year ? model.months.filter(m => m.mesKey.startsWith(year)) : model.months;
}

const TB_COLS = [
  { col: 'mesKey', label: 'Mês' },
  { col: 'vendasBrutas', label: 'Vendas Brutas', money: true, title: 'Base NF / faturação bruta Unilog (armazenagem)' },
  { col: 'vendasLiq', label: 'Vendas Liq', money: true },
  { col: 'armComImp', label: 'Armazenagem', money: true, title: 'Armazenagem total pago c/ impostos' },
  { col: 'freteB2b', label: 'Fretes B2B', money: true, title: 'Fretes B2B — quinzenal' },
  { col: 'total', label: 'Total (Arm+B2B)', money: true, title: 'Armazenagem + Fretes B2B' },
  { col: 'pct', label: '% s/ Vendas', pct: ['total', 'vendasLiq'], title: 'Total (Arm+B2B) ÷ Vendas Líquidas (somadas)' },
  { col: 'pctVendasBrutas', label: '% s/ Vendas Brutas', pct: ['total', 'vendasBrutas'], title: 'Total (Arm+B2B) ÷ Vendas Brutas (somadas)' },
  { col: 'freteB2c', label: 'Fretes B2C', money: true, title: 'Fretes B2C — quinzenal' },
  { col: 'totalComB2c', label: 'Total c/ B2C', money: true, title: 'Total Unilog + Fretes B2C' },
  { col: 'pctComB2c', label: '% c/ B2C', pct: ['totalComB2c', 'vendasLiq'], title: 'Total c/ B2C ÷ Vendas Líquidas (somadas)' }
];

function tbSortTh(state, fn, c, extraCls) {
  const active = state.col === c.col;
  const ind = active ? (state.dir === 1 ? '▲' : '▼') : '↕';
  const cls = ['sortable', c.col !== 'mesKey' && c.col !== 'co' ? 'right' : '', active ? (state.dir === 1 ? 'sorted-asc' : 'sorted-desc') : '', extraCls || '']
    .filter(Boolean).join(' ');
  const title = c.title ? ` title="${tbEsc(c.title)}"` : '';
  return `<th class="${cls}"${title} onclick="${fn}('${c.col}')">${tbEsc(c.label)}<span class="sort-ind">${ind}</span></th>`;
}

function tbCells(r) {
  return TB_COLS.slice(1).map(c => {
    if (c.money) return `<td class="right">${r[c.col] > 0 ? tbMoney(r[c.col]) : '—'}</td>`;
    return `<td class="right">${tbPct(r[c.pct[0]], r[c.pct[1]])}</td>`;
  }).join('');
}

function tbCoBadge(co) {
  return `<span class="tb-co tb-co-${co}">${co}</span>`;
}

function tbSortValue(r, col) {
  if (col === 'mesKey') return r.mesKey;
  const v = r[col];
  return v == null ? -Infinity : v;
}

function tbSortedMonths(months, view) {
  const pick = m => (view === 'DFB' || view === 'QB') ? m.per[view] : m.total;
  return [...months].sort((a, b) => {
    const va = tbSortValue({ ...pick(a), mesKey: a.mesKey }, tbSort.col);
    const vb = tbSortValue({ ...pick(b), mesKey: b.mesKey }, tbSort.col);
    if (va === vb) return 0;
    if (typeof va === 'number' && typeof vb === 'number') return tbSort.dir * (va - vb);
    return tbSort.dir * String(va).localeCompare(String(vb), 'pt', { numeric: true });
  });
}

function tbDoSort(col) {
  tbSort = tbSort.col === col ? { col, dir: -tbSort.dir } : { col, dir: col === 'mesKey' ? 1 : -1 };
  renderTotalBrasil();
}

function tbDoCoSort(col) {
  tbCoSort = tbCoSort.col === col ? { col, dir: -tbCoSort.dir } : { col, dir: col === 'co' ? 1 : -1 };
  renderTotalBrasil();
}

function tbMonthRowsHtml(m, view) {
  const miss = m.missing.map(t => `<span class="tb-miss">${tbEsc(t)}</span>`).join('');
  if (view === 'DFB' || view === 'QB') {
    const r = m.per[view];
    const lbl = tbMissingLabel(view, r);
    return `<tr><td>${tbEsc(m.mesLabel)}${lbl ? `<span class="tb-miss">${tbEsc(lbl)}</span>` : ''}</td>${tbCells(r)}</tr>`;
  }
  let html = `<tr class="tb-month-total"><td>${tbEsc(m.mesLabel)}${miss}</td>${tbCells(m.total)}</tr>`;
  if (view === 'detail') {
    html += TB_COMPANIES.map(co => `<tr class="tb-co-row"><td>${tbCoBadge(co)}${tbHasAny(m.per[co]) ? '' : '<span style="color:var(--muted)">sem dados</span>'}</td>${tbCells(m.per[co])}</tr>`).join('');
  }
  return html;
}

function tbRenderKpis(months) {
  const kpis = document.getElementById('tb-kpis');
  if (!kpis) return;
  const tot = tbSum(months.map(m => m.total));
  const per = {};
  TB_COMPANIES.forEach(co => { per[co] = tbSum(months.map(m => m.per[co])); });
  const split = k => TB_COMPANIES.map(co => `${co} ${tbMoney(per[co][k])}`).join(' · ');
  kpis.innerHTML = `
    <div class="kpi"><div class="label">Vendas Brutas</div><div class="value">${tbMoney(tot.vendasBrutas)}</div><div class="sub">${split('vendasBrutas')}</div></div>
    <div class="kpi"><div class="label">Vendas Líquidas</div><div class="value">${tbMoney(tot.vendasLiq)}</div><div class="sub">${split('vendasLiq')}</div></div>
    <div class="kpi"><div class="label">Custo Armazenagem</div><div class="value">${tbMoney(tot.armComImp)}</div><div class="sub">${split('armComImp')}</div></div>
    <div class="kpi"><div class="label">Custo Fretes B2B</div><div class="value">${tbMoney(tot.freteB2b)}</div><div class="sub">${split('freteB2b')}</div></div>
    <div class="kpi flag"><div class="label">Custo Total Unilog Brasil</div><div class="value">${tbMoney(tot.total)}</div><div class="sub">Arm + B2B · ${tbPct(tot.total, tot.vendasLiq)} s/ vendas · ${tbPct(tot.total, tot.vendasBrutas)} s/ brutas<br>${split('total')}</div></div>
    <div class="kpi"><div class="label">Custo Fretes B2C</div><div class="value">${tbMoney(tot.freteB2c)}</div><div class="sub">${split('freteB2c')}</div></div>
    <div class="kpi flag"><div class="label">Total Brasil c/ B2C</div><div class="value">${tbMoney(tot.totalComB2c)}</div><div class="sub">${tbPct(tot.totalComB2c, tot.vendasLiq)} s/ vendas<br>${split('totalComB2c')}</div></div>`;
}

function tbRenderCoTable(months, yearLabel) {
  const head = document.getElementById('tb-coHead');
  const body = document.getElementById('tb-coBody');
  if (!head || !body) return;
  const tot = tbSum(months.map(m => m.total));
  const cols = [{ col: 'co', label: 'Empresa' }, ...TB_COLS.slice(1), { col: 'share', label: 'Peso no Total c/ B2C', title: 'Total c/ B2C da empresa ÷ Total Brasil c/ B2C' }, { col: 'nMeses', label: 'Meses c/ dados' }];
  head.innerHTML = cols.map(c => tbSortTh(tbCoSort, 'tbDoCoSort', c)).join('');
  const rows = TB_COMPANIES.map(co => {
    const r = tbSum(months.map(m => m.per[co]));
    r.co = co;
    r.share = tot.totalComB2c > 0 ? r.totalComB2c / tot.totalComB2c : null;
    r.nMeses = months.filter(m => tbHasAny(m.per[co])).length;
    return r;
  }).sort((a, b) => {
    const va = a[tbCoSort.col] ?? -Infinity;
    const vb = b[tbCoSort.col] ?? -Infinity;
    if (typeof va === 'number' && typeof vb === 'number') return tbCoSort.dir * (va - vb);
    return tbCoSort.dir * String(va).localeCompare(String(vb), 'pt');
  });
  body.innerHTML = rows.map(r => `<tr><td>${tbCoBadge(r.co)}</td>${tbCells(r)}<td class="right">${r.share != null ? tbPct(r.totalComB2c, tot.totalComB2c) : '—'}</td><td class="right">${r.nMeses} / ${months.length}</td></tr>`).join('')
    + `<tr class="tb-total-row"><td>Total Brasil${yearLabel ? ' ' + tbEsc(yearLabel) : ''}</td>${tbCells(tot)}<td class="right">${tot.totalComB2c > 0 ? '100%' : '—'}</td><td class="right">${months.length}</td></tr>`;
}

function renderTotalBrasil() {
  const gate = document.getElementById('tb-gate');
  const content = document.getElementById('tb-content');
  const isDfb = tbCompany() === 'DFB';
  if (gate) gate.style.display = isDfb ? 'none' : 'block';
  if (content) content.style.display = isDfb ? '' : 'none';
  if (!isDfb) return;

  const status = document.getElementById('tb-status');
  const warn = document.getElementById('tb-warn');
  if (tbLoading && !tbLastModel) {
    if (status) status.innerHTML = '<span class="spin" style="width:12px;height:12px;border-width:2px;vertical-align:middle"></span> A carregar DFB e QB…';
    return;
  }
  const model = tbBuildModel();
  tbLastModel = model;

  const yearSel = document.getElementById('tb-yearFilter');
  const years = [...new Set(model.months.map(m => m.mesKey.slice(0, 4)).filter(y => /^\d{4}$/.test(y)))].sort();
  const prevYear = yearSel?.dataset.init ? yearSel.value : null;
  if (yearSel) {
    yearSel.innerHTML = ['<option value="">Todos os anos</option>'].concat(years.map(y => `<option value="${y}">${y}</option>`)).join('');
    if (prevYear != null && (prevYear === '' || years.includes(prevYear))) yearSel.value = prevYear;
    else if (years.length) yearSel.value = years[years.length - 1];
    yearSel.dataset.init = '1';
  }
  const year = yearSel?.value || '';
  const view = document.getElementById('tb-viewMode')?.value || 'detail';
  const months = tbFilteredMonths(model, year);

  if (status) {
    const qb = tbQbCache;
    const at = qb?.at ? new Date(qb.at).toLocaleTimeString('pt-PT', { hour: '2-digit', minute: '2-digit' }) : '';
    status.textContent = `${tbLoading ? 'A actualizar… · ' : ''}DFB: ${model.dfbCount} mês(es) · QB: ${model.qbCount} mês(es)`
      + (qb ? ` (QB lido da cloud ${at}: ${qb.armMonths} mês(es) armazém, ${qb.qzFiles} quinzenal(is))` : '');
  }
  if (warn) {
    const msgs = [];
    if (tbQbCache?.error) msgs.push(`Erro a ler QB: ${tbEsc(tbQbCache.error)}`);
    const missN = months.filter(m => m.missing.length).length;
    if (missN) msgs.push(`${missN} mês(es) com dados em falta numa empresa (assinalados na tabela) — a soma desses meses só inclui o que existe.`);
    warn.style.display = msgs.length ? 'block' : 'none';
    warn.innerHTML = msgs.join('<br>');
  }

  const empty = document.getElementById('tb-empty');
  const dash = document.getElementById('tb-dash');
  if (empty) empty.style.display = months.length ? 'none' : 'block';
  if (dash) dash.style.display = months.length ? 'block' : 'none';
  if (!months.length) return;

  tbRenderKpis(months);
  const head = document.getElementById('tb-head');
  if (head) head.innerHTML = TB_COLS.map(c => tbSortTh(tbSort, 'tbDoSort', c)).join('');
  const body = document.getElementById('tb-body');
  if (body) {
    const sorted = tbSortedMonths(months, view);
    const yearLabel = year ? ' ' + year : '';
    let foot;
    if (view === 'DFB' || view === 'QB') {
      foot = `<tr class="tb-total-row"><td>Total ${view}${yearLabel}</td>${tbCells(tbSum(months.map(m => m.per[view])))}</tr>`;
    } else {
      foot = `<tr class="tb-total-row"><td>Total Brasil${yearLabel}</td>${tbCells(tbSum(months.map(m => m.total)))}</tr>`;
      if (view === 'detail') {
        foot += TB_COMPANIES.map(co => `<tr class="tb-co-total-row"><td>${tbCoBadge(co)}Total${yearLabel}</td>${tbCells(tbSum(months.map(m => m.per[co])))}</tr>`).join('');
      }
    }
    body.innerHTML = sorted.map(m => tbMonthRowsHtml(m, view)).join('') + foot;
  }
  tbRenderCoTable(months, year);
  if (typeof scheduleColResize === 'function') scheduleColResize();
}

async function openTotalBrasil(force) {
  syncTotalBrasilTab();
  if (tbCompany() !== 'DFB') {
    renderTotalBrasil();
    return;
  }
  tbLoading = true;
  renderTotalBrasil();
  try {
    await Promise.all([
      tbEnsureDfb(),
      tbEnsureQb(!!force).catch(e => {
        console.warn('[totalBrasil] QB', e);
        tbQbCache = { co: 'QB', at: Date.now(), rows: [], armMonths: 0, qzFiles: 0, error: e.message || String(e) };
      })
    ]);
  } finally {
    tbLoading = false;
  }
  if (tbIsActive() && tbCompany() === 'DFB') renderTotalBrasil();
}

function refreshTotalBrasil() {
  tbQbCache = null;
  openTotalBrasil(true).catch(e => console.warn('[totalBrasil] refresh', e));
}

function exportTotalBrasilExcel() {
  const model = tbLastModel || tbBuildModel();
  const year = document.getElementById('tb-yearFilter')?.value || '';
  const months = tbFilteredMonths(model, year);
  if (!months.length) { if (typeof toast === 'function') toast('Sem dados para exportar.', 'error'); return; }
  const c = v => Math.round((Number(v) || 0) * 100) / 100;
  const line = (mes, empresa, r, extra) => ({
    Mes: mes,
    Empresa: empresa,
    VendasBrutas: c(r.vendasBrutas),
    VendasLiq: c(r.vendasLiq),
    ArmazenagemComImpostos: c(r.armComImp),
    FretesB2B: c(r.freteB2b),
    TotalArmB2B: c(r.total),
    PctSobreVendasLiq: r.pct ?? '',
    PctSobreVendasBrutas: r.pctVendasBrutas ?? '',
    FretesB2C: c(r.freteB2c),
    TotalComB2C: c(r.totalComB2c),
    PctComB2CSobreVendasLiq: r.pctComB2c ?? '',
    Observacao: extra || ''
  });
  const detail = [];
  months.forEach(m => {
    detail.push(line(m.mesLabel, 'Total Brasil', m.total, m.missing.join('; ')));
    TB_COMPANIES.forEach(co => detail.push(line(m.mesLabel, co, m.per[co], tbMissingLabel(co, m.per[co]))));
  });
  const yl = year ? `Total ${year}` : 'Total';
  detail.push(line(yl, 'Total Brasil', tbSum(months.map(m => m.total))));
  TB_COMPANIES.forEach(co => detail.push(line(yl, co, tbSum(months.map(m => m.per[co])))));
  const totalOnly = months.map(m => line(m.mesLabel, 'Total Brasil', m.total, m.missing.join('; ')));
  totalOnly.push(line(yl, 'Total Brasil', tbSum(months.map(m => m.total))));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(totalOnly), 'Total Brasil');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detail), 'Por empresa');
  XLSX.writeFile(wb, `total_brasil_DFB_QB_${year || 'todos'}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  if (typeof toast === 'function') toast('Excel Total Brasil exportado.', 'success');
}

/** Tab is DFB-only; switching to QB while on it moves to Total Custo Unilog (QB). */
function syncTotalBrasilTab() {
  const isDfb = tbCompany() === 'DFB';
  document.querySelectorAll('.cob-its .it[data-cob="totalBrasil"]').forEach(it => it.classList.toggle('cob-hidden', !isDfb));
}

function tbOnCompanyChange() {
  tbQbCache = null;
  tbLastModel = null;
  syncTotalBrasilTab();
  if (tbIsActive()) {
    if (tbCompany() !== 'DFB' && typeof activateCobrancasPage === 'function') activateCobrancasPage('custoUnilog');
    else openTotalBrasil().catch(() => {});
  }
}

document.getElementById('coSel')?.addEventListener('change', tbOnCompanyChange);
syncTotalBrasilTab();

window.openTotalBrasil = openTotalBrasil;
window.renderTotalBrasil = renderTotalBrasil;
window.refreshTotalBrasil = refreshTotalBrasil;
window.exportTotalBrasilExcel = exportTotalBrasilExcel;
window.tbDoSort = tbDoSort;
window.tbDoCoSort = tbDoCoSort;
window.tbLoadCompanyReadOnly = tbLoadCompanyReadOnly;
