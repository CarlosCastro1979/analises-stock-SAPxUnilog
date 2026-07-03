/* Inventário — doc. de acerto (layout PT: folhas 0008, 0009, 0011, 0022, 0044) */
(function (global) {
  'use strict';

  const INVENTARIO_ACERTO_JS_VERSION = '1.0.2';
  const INVENTARIO_ACERTO_DEPOTS = ['8', '9', '11', '22', '44'];
  const INVENTARIO_ACERTO_SHEETS = { 8: '0008', 9: '0009', 11: '0011', 22: '0022', 44: '0044' };
  const DATA_START_ROW = 9;
  const HEADER_LAST_ROW = 7;
  const TEMPLATE_URL = 'assets/inventario_acerto_template.xlsx';

  let inventarioAcertoRows = [];
  let inventarioAcertoTemplateBuf = null;

  function iaDepotLabel(dk) {
    const m = typeof DEPOT_MAP !== 'undefined' ? DEPOT_MAP : {};
    return m[dk]?.label || dk;
  }

  function resolveSapMaterial(reconRow, depotKey) {
    const ean = reconRow?.EAN_norm;
    if (!ean) return '';

    if (reconRow.material) return String(reconRow.material).trim();

    if (typeof buildEanSapCodeMap === 'function' && typeof sapData !== 'undefined' && sapData) buildEanSapCodeMap();

    const sap = typeof sapData !== 'undefined' ? sapData : [];
    const depotHit = sap.find((x) => x.deposito === depotKey && x.EAN_norm === ean && x.material);
    if (depotHit) return String(depotHit.material).trim();

    if (typeof _eanSapCodeMap !== 'undefined' && _eanSapCodeMap?.[ean]) return [..._eanSapCodeMap[ean]][0];

    const anyEan = sap.find((x) => x.EAN_norm === ean && x.material);
    if (anyEan) return String(anyEan.material).trim();

    const asMaterial = sap.find((x) => x.material === ean);
    if (asMaterial) return String(asMaterial.material).trim();

    return '';
  }

  function getSapUmb(ean) {
    const row = (typeof sapData !== 'undefined' ? sapData : []).find((x) => x.EAN_norm === ean);
    const umb = row?.umb || row?.UMB;
    return umb ? String(umb).trim() : 'UN';
  }

  function iaAcertoHI(r) {
    if (r.qt_uni > 0 && r.qt_sap === 0) return { h: r.qt_uni, i: 0 };
    if (r.qt_sap > 0 && r.qt_uni === 0) return { h: 0, i: r.qt_sap };
    return { h: r.qt_uni, i: r.qt_sap };
  }

  function buildAcertoRowsForDepot(dk) {
    const rows = (typeof reconData !== 'undefined' ? reconData[dk] : null) || [];
    const precoMap = typeof buildSapPrecoMap === 'function' ? buildSapPrecoMap() : {};
    let skippedNoMaterial = 0;
    const out = rows
      .filter((r) => r.diff !== 0)
      .map((r) => {
        const hi = iaAcertoHI(r);
        const material = resolveSapMaterial(r, dk);
        return {
          depot: dk,
          depotSheet: INVENTARIO_ACERTO_SHEETS[dk] || dk.padStart(4, '0'),
          depotLabel: iaDepotLabel(dk),
          material,
          EAN_norm: r.EAN_norm,
          desc: r.desc || '',
          lote: r.Lote_norm || '',
          preco: precoMap[r.EAN_norm] || 0,
          umb: getSapUmb(r.EAN_norm),
          qt_wms: hi.h,
          qt_sap: hi.i,
          diff: hi.h - hi.i,
          _estado: hi.h > 0 && hi.i === 0 ? 'Só WMS' : hi.i > 0 && hi.h === 0 ? 'Só SAP' : 'Diverg.',
        };
      })
      .filter((r) => {
        if (r.material) return true;
        skippedNoMaterial++;
        return false;
      })
      .sort((a, b) => {
        const ma = a.material.localeCompare(b.material);
        if (ma) return ma;
        return (a.lote || '').localeCompare(b.lote || '');
      });
    out._skippedNoMaterial = skippedNoMaterial;
    return out;
  }

  function buildInventarioAcertoRows() {
    if (typeof company !== 'undefined' && company !== 'DFB') return [];
    const out = [];
    for (const dk of INVENTARIO_ACERTO_DEPOTS) out.push(...buildAcertoRowsForDepot(dk));
    inventarioAcertoRows = out;
    return out;
  }

  function iaClearSheetDataFromRow(ws, startRow) {
    if (!ws) return;
    const startR = startRow - 1;
    for (const addr of Object.keys(ws)) {
      if (addr.charAt(0) === '!') continue;
      try {
        const cell = XLSX.utils.decode_cell(addr);
        if (cell.r >= startR) delete ws[addr];
      } catch (_) {
        /* ignore malformed keys */
      }
    }
    if (ws['!rows'] && ws['!rows'].length > startR) ws['!rows'] = ws['!rows'].slice(0, startR);
    const headerEndR = startRow - 2;
    const base = ws['!ref']
      ? XLSX.utils.decode_range(ws['!ref'])
      : { s: { r: 0, c: 0 }, e: { r: headerEndR, c: 12 } };
    base.e.r = Math.max(base.s.r, headerEndR);
    base.e.c = Math.max(base.e.c, 12);
    ws['!ref'] = XLSX.utils.encode_range(base);
  }

  function iaSetCell(ws, addr, value, kind) {
    if (kind === 'f') {
      ws[addr] = { t: 'n', f: value };
      return;
    }
    if (kind === 'n') {
      const n = Number(value);
      ws[addr] = { t: 'n', v: Number.isFinite(n) ? n : 0 };
      return;
    }
    ws[addr] = { t: 's', v: String(value ?? '') };
  }

  function iaWriteDataRow(ws, rowNum, row) {
    const r = rowNum;
    iaSetCell(ws, 'A' + r, row.material);
    iaSetCell(ws, 'B' + r, row.desc);
    if (row.lote) iaSetCell(ws, 'C' + r, row.lote);
    if (row.preco > 0) iaSetCell(ws, 'D' + r, row.preco, 'n');
    iaSetCell(ws, 'E' + r, row.umb || 'UN');
    iaSetCell(ws, 'F' + r, '+$M$3', 'f');
    iaSetCell(ws, 'H' + r, row.qt_wms, 'n');
    iaSetCell(ws, 'I' + r, row.qt_sap, 'n');
    iaSetCell(ws, 'J' + r, 'D' + r + '*H' + r, 'f');
    iaSetCell(ws, 'K' + r, 'H' + r + '-I' + r, 'f');
    iaSetCell(ws, 'L' + r, 'H' + r + '-I' + r, 'f');
    iaSetCell(ws, 'M' + r, 'L' + r + '*D' + r, 'f');
  }

  function iaUpdateSheetRef(ws, lastRow) {
    const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: HEADER_LAST_ROW - 1, c: 12 } };
    range.e.r = Math.max(range.e.r, lastRow - 1);
    range.e.c = Math.max(range.e.c, 12);
    ws['!ref'] = XLSX.utils.encode_range(range);
  }

  function iaExcelSerialDate(d) {
    const dt = d || new Date();
    const utc = Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate());
    return (utc - Date.UTC(1899, 11, 30)) / 86400000;
  }

  async function loadInventarioAcertoTemplate() {
    if (inventarioAcertoTemplateBuf) return inventarioAcertoTemplateBuf;
    const resp = await fetch(TEMPLATE_URL);
    if (!resp.ok) throw new Error('Template não encontrado (' + TEMPLATE_URL + ')');
    inventarioAcertoTemplateBuf = await resp.arrayBuffer();
    return inventarioAcertoTemplateBuf;
  }

  function inventarioAcertoExportFilename() {
    const d = new Date().toISOString().slice(0, 10);
    return 'Inventario_DFB_acerto_inventario_' + d + '.xlsx';
  }

  async function exportInventarioAcertoXlsx() {
    if (typeof company !== 'undefined' && company !== 'DFB') {
      toast('Doc. de acerto disponível apenas para DFB', 'error');
      return;
    }
    if (!Object.keys(typeof reconData !== 'undefined' ? reconData : {}).length) {
      toast('Sem dados de conciliação — carrega SAP e Unilog', 'error');
      return;
    }
    buildInventarioAcertoRows();
    const totalLines = inventarioAcertoRows.length;
    if (!totalLines) {
      toast('Sem linhas com diferença nos depósitos 0008–0044 (ou sem código SAP)', 'info');
      return;
    }
    try {
      const buf = await loadInventarioAcertoTemplate();
      const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellStyles: true });
      let skippedNoMaterial = 0;
      for (const dk of INVENTARIO_ACERTO_DEPOTS) {
        const sheetName = INVENTARIO_ACERTO_SHEETS[dk];
        const ws = wb.Sheets[sheetName];
        if (!ws) continue;
        iaClearSheetDataFromRow(ws, DATA_START_ROW);
        ws['M3'] = { t: 's', v: sheetName };
        ws['L5'] = { t: 'n', v: iaExcelSerialDate(new Date()) };
        const rows = buildAcertoRowsForDepot(dk);
        skippedNoMaterial += rows._skippedNoMaterial || 0;
        let rn = DATA_START_ROW;
        for (const row of rows) {
          iaWriteDataRow(ws, rn, row);
          rn++;
        }
        iaUpdateSheetRef(ws, rows.length ? rn - 1 : HEADER_LAST_ROW);
      }
      XLSX.writeFile(wb, inventarioAcertoExportFilename());
      let msg = 'Exportado: ' + totalLines + ' linhas de acerto';
      if (skippedNoMaterial) msg += ' · ' + skippedNoMaterial + ' omitidas (sem código SAP)';
      toast(msg, 'success');
    } catch (e) {
      toast('Erro ao exportar: ' + (e.message || e), 'error');
    }
  }

  function renderInventarioAcertoKpis(rows) {
    const el = document.getElementById('iaKpis');
    if (!el) return;
    if (!rows.length) {
      el.innerHTML = '';
      return;
    }
    const units = rows.reduce((s, r) => s + Math.abs(r.diff), 0);
    const wmsOnly = rows.filter((r) => r.qt_wms > 0 && r.qt_sap === 0).length;
    const sapOnly = rows.filter((r) => r.qt_sap > 0 && r.qt_wms === 0).length;
    const both = rows.length - wmsOnly - sapOnly;
    el.innerHTML =
      '<div class="kpi"><div class="kl">Linhas acerto</div><div class="kv">' +
      (typeof fmtKvNum === 'function' ? fmtKvNum(rows.length) : rows.length) +
      '</div><div class="ks">depósitos 0008–0044 · só diferenças</div></div>' +
      '<div class="kpi"><div class="kl">Unidades |dif|</div><div class="kv">' +
      (typeof fmtKvNum === 'function' ? fmtKvNum(units) : units) +
      '</div><div class="ks">soma |WMS−SAP| por linha</div></div>' +
      '<div class="kpi"><div class="kl">Só WMS</div><div class="kv b">' +
      (typeof fmtKvNum === 'function' ? fmtKvNum(wmsOnly) : wmsOnly) +
      '</div><div class="ks">H físico · I=0</div></div>' +
      '<div class="kpi"><div class="kl">Só SAP / diverg.</div><div class="kv">' +
      (typeof fmtKvNum === 'function' ? fmtKvNum(sapOnly) : sapOnly) +
      ' / ' +
      (typeof fmtKvNum === 'function' ? fmtKvNum(both) : both) +
      '</div><div class="ks">SAP-only · ambos com stock</div></div>';
  }

  function renderInventarioAcertoTable(rows) {
    const body = document.getElementById('iaBody');
    const empty = document.getElementById('iaEmpty');
    const wrap = document.getElementById('iaTableWrap');
    if (!body) return;
    if (!rows.length) {
      body.innerHTML = '';
      if (empty) empty.style.display = 'block';
      if (wrap) wrap.style.display = 'none';
      return;
    }
    if (empty) empty.style.display = 'none';
    if (wrap) wrap.style.display = '';
    const sorted = typeof sorts !== 'undefined' && sorts.ia ? rows.slice() : rows;
    if (typeof sorts !== 'undefined' && sorts.ia && typeof applySort === 'function') applySort('ia', sorted);
    body.innerHTML = sorted
      .map((r) => {
        const estado =
          r.qt_wms > 0 && r.qt_sap === 0
            ? '<span class="badge binfo">Só WMS</span>'
            : r.qt_sap > 0 && r.qt_wms === 0
              ? '<span class="badge bwarn">Só SAP</span>'
              : '<span class="badge bwarn">Diverg.</span>';
        return (
          '<tr>' +
          '<td><span class="badge bgray">' +
          r.depotSheet +
          '</span></td>' +
          '<td style="font-family:var(--mono);font-size:11px">' +
          r.material +
          '</td>' +
          '<td class="col-desc">' +
          (r.desc || '—') +
          '</td>' +
          '<td style="font-family:var(--mono);font-size:11px">' +
          (r.lote || '—') +
          '</td>' +
          '<td class="num">' +
          (typeof fmtPtNum === 'function' ? fmtPtNum(r.qt_wms) : r.qt_wms) +
          '</td>' +
          '<td class="num">' +
          (typeof fmtPtNum === 'function' ? fmtPtNum(r.qt_sap) : r.qt_sap) +
          '</td>' +
          '<td class="num">' +
          (typeof fmtKvSigned === 'function' ? fmtKvSigned(r.diff) : r.diff) +
          '</td>' +
          '<td>' +
          estado +
          '</td>' +
          '</tr>'
        );
      })
      .join('');
    const table = document.getElementById('iaTable');
    if (typeof updateSortHeaders === 'function') updateSortHeaders('ia');
    if (table && typeof enableTableSort === 'function') enableTableSort(table);
    if (typeof scheduleTableSort === 'function') scheduleTableSort();
  }

  function renderInventarioAcertoPage() {
    const qbNote = document.getElementById('iaQbNote');
    const main = document.getElementById('iaMain');
    if (typeof company !== 'undefined' && company !== 'DFB') {
      if (qbNote) qbNote.style.display = 'block';
      if (main) main.style.display = 'none';
      return;
    }
    if (qbNote) qbNote.style.display = 'none';
    if (main) main.style.display = '';
    const rows = buildInventarioAcertoRows();
    renderInventarioAcertoKpis(rows);
    renderInventarioAcertoTable(rows);
    const depCounts = INVENTARIO_ACERTO_DEPOTS.map((dk) => {
      const n = buildAcertoRowsForDepot(dk).length;
      return INVENTARIO_ACERTO_SHEETS[dk] + ': ' + n;
    }).join(' · ');
    const sub = document.getElementById('iaDepotCounts');
    if (sub) sub.textContent = depCounts;
  }

  function syncInventarioAcertoNav() {
    const nt = document.getElementById('nt-inventario-acerto');
    if (!nt) return;
    nt.style.display = typeof company !== 'undefined' && company === 'DFB' ? '' : 'none';
  }

  function iaDoSort(col) {
    if (typeof doSort === 'function') doSort('ia', col);
  }

  global.INVENTARIO_ACERTO_JS_VERSION = INVENTARIO_ACERTO_JS_VERSION;
  global.INVENTARIO_ACERTO_DEPOTS = INVENTARIO_ACERTO_DEPOTS;
  global.buildInventarioAcertoRows = buildInventarioAcertoRows;
  global.renderInventarioAcertoPage = renderInventarioAcertoPage;
  global.exportInventarioAcertoXlsx = exportInventarioAcertoXlsx;
  global.inventarioAcertoExportFilename = inventarioAcertoExportFilename;
  global.syncInventarioAcertoNav = syncInventarioAcertoNav;
  global.iaDoSort = iaDoSort;
})(typeof window !== 'undefined' ? window : globalThis);
