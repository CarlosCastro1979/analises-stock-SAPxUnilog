/* insumos.js v1.0.0 — Unilog insumos CSV + SAP CX Papelão / ocupação por posição */
(function (global) {
  'use strict';

  const INSUMOS_JS_VERSION = '1.0.0';
  let insumosUniRows = [];
  let insumosFileName = '';
  let insumosBuffer = null;
  let insumosActiveTab = 'sap';
  let insumosCloudLoaded = false;
  let insumosSaving = false;

  function $(id) { return document.getElementById(id); }
  function toastMsg(msg, type) {
    if (typeof toast === 'function') toast(msg, type || 'success');
  }
  function fmtNum(n) {
    if (typeof fmt === 'function') return fmt(n);
    return (Number(n) || 0).toLocaleString('pt-BR');
  }
  function escHtml(s) {
    if (typeof esc === 'function') return esc(s);
    return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function slotKey() {
    return (typeof EXCEL_SLOTS !== 'undefined' && EXCEL_SLOTS.INSUMOS_UNILOG) ? EXCEL_SLOTS.INSUMOS_UNILOG : 'insumos_unilog';
  }

  function pickField(row, names) {
    for (const n of names) {
      if (row[n] != null && String(row[n]).trim() !== '') return String(row[n]).trim();
    }
    const keys = Object.keys(row || {});
    const want = names.map(n => n.toLowerCase().normalize('NFD').replace(/\p{M}/gu, '').replace(/[^a-z0-9]/g, ''));
    for (const k of keys) {
      const kn = k.toLowerCase().normalize('NFD').replace(/\p{M}/gu, '').replace(/[^a-z0-9]/g, '');
      if (want.includes(kn)) {
        const v = String(row[k] ?? '').trim();
        if (v) return v;
      }
    }
    return '';
  }

  function parseQt(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    let s = String(v).trim().replace(/\s/g, '');
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.includes(',') && !s.includes('.')) s = s.replace(',', '.');
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : 0;
  }

  function normInsumosUniRows(rawRows) {
    const out = [];
    for (const r of rawRows || []) {
      const ua = pickField(r, ['UA', 'Ua', 'Material', 'Cod', 'Código', 'Codigo', 'Produto Codigo', 'Produto Código']);
      const endereco = pickField(r, ['Endereço', 'Endereco', 'Endereço ', 'Endereco ', 'Localização', 'Localizacao', 'Posicao', 'Posição']);
      const desc = pickField(r, ['Descrição', 'Descricao', 'Descriçao', 'Produto', 'Desc']);
      const cliente = pickField(r, ['Cliente', 'Client']);
      const cd = pickField(r, ['CD', 'Cd', 'Centro']);
      const dt = pickField(r, ['Dt_Entrada', 'Dt Entrada', 'Data Entrada', 'Dt.Entrada', 'Data']);
      const qt = parseQt(pickField(r, ['Qtde', 'Qtd', 'Quantidade', 'Qt', 'Qtde ']));
      if (!ua && !endereco && !desc && !qt) continue;
      out.push({
        ua: String(ua || '').replace(/^0+(?=\d)/, '') || String(ua || ''),
        uaRaw: ua,
        desc,
        cliente,
        cd,
        endereco,
        qt,
        dt
      });
    }
    return out;
  }

  function parseCsvText(text) {
    const wb = XLSX.read(text, { type: 'string', FS: ';', raw: false });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval: '' });
  }

  function rowsFromArrayBuffer(buf, fileName) {
    const name = String(fileName || '').toLowerCase();
    if (name.endsWith('.csv') || name.endsWith('.txt')) {
      let text = '';
      try {
        text = new TextDecoder('utf-8').decode(buf);
        if (text.includes('\uFFFD') || /Descri..o|Endere..o/.test(text.slice(0, 80))) {
          text = new TextDecoder('windows-1252').decode(buf);
        }
      } catch (_) {
        text = new TextDecoder('windows-1252').decode(buf);
      }
      // Prefer semicolon (Unilog BR); fallback to comma via SheetJS auto
      if (text.split('\n')[0].includes(';')) {
        return parseCsvText(text);
      }
      const wb = XLSX.read(text, { type: 'string' });
      return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    }
    return xlsxRowsFromArrayBuffer(buf);
  }

  async function applyInsumosBuffer(buf, fileName, { autosave = true } = {}) {
    insumosBuffer = buf;
    insumosFileName = fileName || insumosFileName || 'insumos.csv';
    const raw = rowsFromArrayBuffer(buf, insumosFileName);
    insumosUniRows = normInsumosUniRows(raw);
    const fn = $('insumosFn');
    const zone = $('insumosZone');
    if (fn) fn.textContent = '✓ ' + insumosFileName + ' · ' + insumosUniRows.length + ' linhas';
    if (zone) zone.classList.add('loaded');
    renderInsumosPage();
    if (autosave) scheduleInsumosAutosave();
    toastMsg('Insumos Unilog: ' + insumosUniRows.length + ' linhas', 'success');
  }

  function loadInsumosFile(inp) {
    const f = inp?.files?.[0];
    if (!f) return;
    const fnEl = $('insumosFn');
    if (fnEl) fnEl.textContent = '⏳ A ler ' + f.name + '…';
    const r = new FileReader();
    r.onload = async e => {
      try {
        await applyInsumosBuffer(e.target.result, f.name, { autosave: true });
      } catch (err) {
        console.error('[insumos] load', err);
        if (fnEl) fnEl.textContent = '';
        toastMsg('Erro ao ler insumos: ' + (err.message || err), 'error');
      }
    };
    r.onerror = () => {
      if (fnEl) fnEl.textContent = '';
      toastMsg('Erro ao ler ficheiro', 'error');
    };
    r.readAsArrayBuffer(f);
  }

  function scheduleInsumosAutosave() {
    if (typeof scheduleAutosave !== 'function') {
      saveInsumosCloud().catch(() => {});
      return;
    }
    scheduleAutosave('insumosUni', () => saveInsumosCloud(), 1200);
  }

  async function saveInsumosCloud() {
    if (!insumosBuffer || !insumosFileName || insumosSaving) return;
    insumosSaving = true;
    const note = $('insumosSaveNote');
    try {
      if (typeof upsertExcelBinary === 'function') {
        await upsertExcelBinary(slotKey(), insumosFileName, insumosBuffer);
      }
      if (note) note.textContent = 'Guardado na cloud · ' + new Date().toLocaleTimeString('pt-BR');
      if (typeof showSaveIndicator === 'function') showSaveIndicator();
    } catch (e) {
      console.warn('[insumos] save cloud', e);
      if (typeof toastExcelFilesError === 'function') toastExcelFilesError('Insumos: ', e);
      else toastMsg('Falha ao guardar insumos na cloud', 'error');
    } finally {
      insumosSaving = false;
    }
  }

  async function loadInsumosCloud(opts = {}) {
    const silent = !!opts.silent;
    try {
      if (typeof fetchExcelFiles !== 'function') return false;
      const meta = await fetchExcelFiles([slotKey()]);
      const rec = meta[slotKey()];
      if (!rec?.file_data) {
        if (!silent) toastMsg('Sem ficheiro de insumos na cloud', 'error');
        return false;
      }
      const buf = base64ToArrayBuffer(rec.file_data);
      await applyInsumosBuffer(buf, rec.file_name || 'insumos.csv', { autosave: false });
      insumosCloudLoaded = true;
      const note = $('insumosSaveNote');
      if (note) note.textContent = 'Cloud: ' + (rec.file_name || '') + (rec.uploaded_at ? ' · ' + new Date(rec.uploaded_at).toLocaleDateString('pt-BR') : '');
      return true;
    } catch (e) {
      console.warn('[insumos] load cloud', e);
      if (!silent) toastMsg('Erro ao carregar insumos cloud', 'error');
      return false;
    }
  }

  async function ensureInsumosCloudLoaded(opts = {}) {
    if (insumosUniRows.length || insumosCloudLoaded) {
      renderInsumosPage();
      return true;
    }
    return loadInsumosCloud(opts);
  }

  function insumosTab(name, el) {
    insumosActiveTab = name;
    const root = $('insumosIts');
    if (root) root.querySelectorAll('.it').forEach(t => t.classList.remove('active'));
    if (el) el.classList.add('active');
    const sap = $('insumos-sap'), uni = $('insumos-uni'), pos = $('insumos-pos');
    if (sap) sap.style.display = name === 'sap' ? '' : 'none';
    if (uni) uni.style.display = name === 'uni' ? '' : 'none';
    if (pos) pos.style.display = name === 'pos' ? '' : 'none';
    renderInsumosPage();
    if (typeof scheduleColResize === 'function') scheduleColResize();
  }

  function isInsumoRow(ean, desc, cat, subcat) {
    if (typeof isInsumo === 'function' && isInsumo(ean, desc)) return true;
    if (typeof isCaixasPapelao === 'function' && isCaixasPapelao(ean, desc)) return true;
    if (typeof isInsumosCat === 'function' && isInsumosCat(cat)) return true;
    if (typeof isCaixasPapelaoCat === 'function' && (isCaixasPapelaoCat(cat) || isCaixasPapelaoCat(subcat))) return true;
    return false;
  }

  function buildSapInsumosRows() {
    const byEan = new Map();
    const params = (typeof occParams !== 'undefined' && occParams) ? occParams : [];
    const sap = (typeof sapData !== 'undefined' && sapData) ? sapData : [];
    for (const r of sap) {
      if (!r.EAN_norm) continue;
      const param = params.find(p => p.ean === r.EAN_norm);
      const cat = param?.cat || '';
      const subcat = param?.subcat || '';
      if (!isInsumoRow(r.EAN_norm, r.desc, cat, subcat)) continue;
      let g = byEan.get(r.EAN_norm);
      if (!g) {
        g = {
          ean: r.EAN_norm,
          desc: r.desc || '',
          material: r.material || (typeof getSapCode === 'function' ? getSapCode(r.EAN_norm) : ''),
          cat: cat || (typeof INSUMOS_CAT !== 'undefined' ? INSUMOS_CAT : 'Insumos'),
          subcat: subcat || (typeof CAIXAS_PAPELAO_SUBCAT !== 'undefined' ? CAIXAS_PAPELAO_SUBCAT : 'CX Papelao'),
          qt: 0,
          deps: new Set()
        };
        byEan.set(r.EAN_norm, g);
      }
      g.qt += r.qt || 0;
      if (r.deposito) g.deps.add(String(r.deposito).padStart(4, '0'));
      if (!g.material && r.material) g.material = r.material;
      if (!g.desc && r.desc) g.desc = r.desc;
    }
    // Also include occParams flagged as insumos even with 0 SAP stock
    for (const p of params) {
      if (!p?.ean || byEan.has(p.ean)) continue;
      if (!isInsumoRow(p.ean, p.descricao || p.desc || '', p.cat, p.subcat)) continue;
      byEan.set(p.ean, {
        ean: p.ean,
        desc: p.descricao || p.desc || '',
        material: typeof getSapCode === 'function' ? getSapCode(p.ean) : '',
        cat: p.cat || 'Insumos',
        subcat: p.subcat || 'CX Papelao',
        qt: 0,
        deps: new Set()
      });
    }
    return [...byEan.values()].map(r => ({
      ...r,
      deps: [...r.deps].sort().join(', '),
      depsArr: [...r.deps]
    }));
  }

  function buildPosAgg() {
    const byPos = new Map();
    const byPosUa = [];
    for (const r of insumosUniRows) {
      const end = r.endereco || '—';
      let g = byPos.get(end);
      if (!g) {
        g = { endereco: end, qt: 0, uas: new Set(), descs: [] };
        byPos.set(end, g);
      }
      g.qt += r.qt || 0;
      if (r.ua) g.uas.add(r.ua);
      if (r.desc && g.descs.length < 4 && !g.descs.includes(r.desc)) g.descs.push(r.desc);
      byPosUa.push({
        endereco: end,
        ua: r.ua || '—',
        desc: r.desc || '',
        qt: r.qt || 0
      });
    }
    const posRows = [...byPos.values()].map(g => ({
      endereco: g.endereco,
      skus: g.uas.size,
      qt: g.qt,
      descSample: g.descs.join(' · ')
    }));
    return { posRows, byPosUa };
  }

  function renderInsumosPage() {
    const sapRows = buildSapInsumosRows();
    const { posRows, byPosUa } = buildPosAgg();
    const posCount = posRows.filter(p => p.endereco && p.endereco !== '—').length;
    const qtUni = insumosUniRows.reduce((s, r) => s + (r.qt || 0), 0);
    const qtSap = sapRows.reduce((s, r) => s + (r.qt || 0), 0);
    const kpis = $('insumosKpis');
    if (kpis) {
      kpis.innerHTML = `
        <div class="kpi"><div class="kl">SAP Insumos</div><div class="kv">${fmtNum(sapRows.length)}</div><div class="ks">EANs · qt ${fmtNum(qtSap)}</div></div>
        <div class="kpi"><div class="kl">Unilog linhas</div><div class="kv">${fmtNum(insumosUniRows.length)}</div><div class="ks">extract completo</div></div>
        <div class="kpi"><div class="kl">Posições ocupadas</div><div class="kv">${fmtNum(posCount)}</div><div class="ks">endereços com stock</div></div>
        <div class="kpi"><div class="kl">Qt Unilog</div><div class="kv">${fmtNum(qtUni)}</div><div class="ks">soma qtde extract</div></div>
      `;
    }

    // SAP table
    const sapSrch = ($('insumosSapSrch')?.value || '').toLowerCase().trim();
    let sapList = [...sapRows];
    if (sapSrch) {
      sapList = sapList.filter(r =>
        String(r.ean).toLowerCase().includes(sapSrch) ||
        String(r.material).toLowerCase().includes(sapSrch) ||
        String(r.desc).toLowerCase().includes(sapSrch) ||
        String(r.cat).toLowerCase().includes(sapSrch)
      );
    }
    if (typeof sorts !== 'undefined' && sorts.insumosSap && typeof applySort === 'function') applySort('insumosSap', sapList);
    else sapList.sort((a, b) => String(a.desc || '').localeCompare(String(b.desc || ''), 'pt'));
    if (typeof updateSortHeaders === 'function') updateSortHeaders('insumosSap');
    const sapBody = $('insumosSapBody');
    const sapEmpty = $('insumosSapEmpty');
    if (sapBody) {
      if (!sapList.length) {
        sapBody.innerHTML = '';
        if (sapEmpty) sapEmpty.style.display = 'block';
      } else {
        if (sapEmpty) sapEmpty.style.display = 'none';
        sapBody.innerHTML = sapList.map(r => `<tr>
          <td style="font-size:11px">${escHtml(r.cat || '—')}</td>
          <td style="font-size:11px">${escHtml(r.subcat || '—')}</td>
          <td><span class="badge bgray" style="font-family:var(--mono);font-size:11px">${escHtml(r.material || '—')}</span></td>
          <td><span class="badge bneu">${escHtml(r.ean)}</span></td>
          <td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.desc)}">${escHtml(r.desc || '—')}</td>
          <td class="num">${fmtNum(r.qt)}</td>
          <td style="font-size:11px">${escHtml(r.deps || '—')}</td>
        </tr>`).join('');
      }
    }

    // Unilog table
    const uniSrch = ($('insumosUniSrch')?.value || '').toLowerCase().trim();
    let uniList = [...insumosUniRows];
    if (uniSrch) {
      uniList = uniList.filter(r =>
        String(r.ua).toLowerCase().includes(uniSrch) ||
        String(r.desc).toLowerCase().includes(uniSrch) ||
        String(r.endereco).toLowerCase().includes(uniSrch) ||
        String(r.cliente).toLowerCase().includes(uniSrch)
      );
    }
    if (typeof sorts !== 'undefined' && sorts.insumosUni && typeof applySort === 'function') applySort('insumosUni', uniList);
    else uniList.sort((a, b) => String(a.endereco || '').localeCompare(String(b.endereco || ''), 'pt'));
    if (typeof updateSortHeaders === 'function') updateSortHeaders('insumosUni');
    const uniBody = $('insumosUniBody');
    const uniEmpty = $('insumosUniEmpty');
    if (uniBody) {
      if (!uniList.length) {
        uniBody.innerHTML = '';
        if (uniEmpty) uniEmpty.style.display = 'block';
      } else {
        if (uniEmpty) uniEmpty.style.display = 'none';
        uniBody.innerHTML = uniList.map(r => `<tr>
          <td><span class="badge bgray" style="font-family:var(--mono);font-size:11px">${escHtml(r.ua || '—')}</span></td>
          <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.desc)}">${escHtml(r.desc || '—')}</td>
          <td style="font-size:11px">${escHtml(r.cliente || '—')}</td>
          <td style="font-family:var(--mono);font-size:11px">${escHtml(r.endereco || '—')}</td>
          <td class="num">${fmtNum(r.qt)}</td>
          <td style="font-size:11px;color:var(--muted)">${escHtml(r.dt || '—')}</td>
        </tr>`).join('');
      }
    }

    // Positions
    const posSrch = ($('insumosPosSrch')?.value || '').toLowerCase().trim();
    let posList = [...posRows];
    let posUaList = [...byPosUa];
    if (posSrch) {
      posList = posList.filter(r =>
        String(r.endereco).toLowerCase().includes(posSrch) ||
        String(r.descSample).toLowerCase().includes(posSrch)
      );
      posUaList = posUaList.filter(r =>
        String(r.endereco).toLowerCase().includes(posSrch) ||
        String(r.ua).toLowerCase().includes(posSrch) ||
        String(r.desc).toLowerCase().includes(posSrch)
      );
    }
    if (typeof sorts !== 'undefined' && sorts.insumosPos && typeof applySort === 'function') applySort('insumosPos', posList);
    else posList.sort((a, b) => b.qt - a.qt);
    if (typeof sorts !== 'undefined' && sorts.insumosPosUa && typeof applySort === 'function') applySort('insumosPosUa', posUaList);
    else posUaList.sort((a, b) => String(a.endereco).localeCompare(String(b.endereco), 'pt') || b.qt - a.qt);
    if (typeof updateSortHeaders === 'function') {
      updateSortHeaders('insumosPos');
      updateSortHeaders('insumosPosUa');
    }
    const posBody = $('insumosPosBody');
    const posUaBody = $('insumosPosUaBody');
    const posEmpty = $('insumosPosEmpty');
    const noPos = !insumosUniRows.length;
    if (posEmpty) posEmpty.style.display = noPos ? 'block' : 'none';
    if (posBody) {
      posBody.innerHTML = noPos ? '' : posList.map(r => `<tr>
        <td style="font-family:var(--mono);font-size:12px">${escHtml(r.endereco)}</td>
        <td class="num">${fmtNum(r.skus)}</td>
        <td class="num">${fmtNum(r.qt)}</td>
        <td style="font-size:11px;max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.descSample)}">${escHtml(r.descSample || '—')}</td>
      </tr>`).join('');
    }
    if (posUaBody) {
      posUaBody.innerHTML = noPos ? '' : posUaList.map(r => `<tr>
        <td style="font-family:var(--mono);font-size:11px">${escHtml(r.endereco)}</td>
        <td><span class="badge bgray" style="font-family:var(--mono);font-size:11px">${escHtml(r.ua)}</span></td>
        <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.desc)}">${escHtml(r.desc || '—')}</td>
        <td class="num">${fmtNum(r.qt)}</td>
      </tr>`).join('');
    }

    if (typeof scheduleColResize === 'function') scheduleColResize();
  }

  function exportInsumosXlsx() {
    if (typeof XLSX === 'undefined') { toastMsg('XLSX indisponível', 'error'); return; }
    const sapRows = buildSapInsumosRows();
    const { posRows, byPosUa } = buildPosAgg();
    if (!sapRows.length && !insumosUniRows.length) { toastMsg('Sem dados de insumos', 'error'); return; }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sapRows.map(r => ({
      Categoria: r.cat, Subcategoria: r.subcat, Cod_SAP: r.material, EAN: r.ean, Descricao: r.desc, Qt_SAP: r.qt, Depositos: r.deps
    }))), 'SAP Insumos');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(insumosUniRows.map(r => ({
      CD: r.cd, Cliente: r.cliente, Descricao: r.desc, UA: r.ua, Endereco: r.endereco, Qtde: r.qt, Dt_Entrada: r.dt
    }))), 'Unilog Extract');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(posRows.map(r => ({
      Endereco: r.endereco, UAs: r.skus, Qtde: r.qt, Materiais: r.descSample
    }))), 'Posicoes');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(byPosUa.map(r => ({
      Endereco: r.endereco, UA: r.ua, Descricao: r.desc, Qtde: r.qt
    }))), 'Qtd por Pos x UA');
    XLSX.writeFile(wb, 'insumos_' + new Date().toISOString().slice(0, 10) + '.xlsx');
    toastMsg('Excel exportado', 'success');
  }

  // Drag & drop on zone
  function initInsumosDnD() {
    const zone = $('insumosZone');
    if (!zone || zone.dataset.dnd) return;
    zone.dataset.dnd = '1';
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('drag');
      const f = e.dataTransfer?.files?.[0];
      if (!f) return;
      const inp = $('insumosFile');
      if (inp) {
        try {
          const dt = new DataTransfer();
          dt.items.add(f);
          inp.files = dt.files;
        } catch (_) {}
      }
      const reader = new FileReader();
      reader.onload = ev => applyInsumosBuffer(ev.target.result, f.name, { autosave: true }).catch(err => {
        toastMsg('Erro: ' + (err.message || err), 'error');
      });
      reader.readAsArrayBuffer(f);
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initInsumosDnD);
  } else {
    initInsumosDnD();
  }

  global.INSUMOS_JS_VERSION = INSUMOS_JS_VERSION;
  global.loadInsumosFile = loadInsumosFile;
  global.loadInsumosCloud = loadInsumosCloud;
  global.ensureInsumosCloudLoaded = ensureInsumosCloudLoaded;
  global.saveInsumosCloud = saveInsumosCloud;
  global.renderInsumosPage = renderInsumosPage;
  global.insumosTab = insumosTab;
  global.exportInsumosXlsx = exportInsumosXlsx;
  global.getInsumosUniRows = () => insumosUniRows;
})(typeof window !== 'undefined' ? window : globalThis);
