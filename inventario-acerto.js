/* Inventário — doc. de acerto (layout PT: folhas 0008, 0009, 0011, 0022, 0044) */
(function (global) {
  'use strict';

  const INVENTARIO_ACERTO_JS_VERSION = '1.0.3';
  const INVENTARIO_ACERTO_DEPOTS = ['8', '9', '11', '22', '44'];
  const INVENTARIO_ACERTO_SHEETS = { 8: '0008', 9: '0009', 11: '0011', 22: '0022', 44: '0044' };
  const DATA_START_ROW = 9;
  const DATE_CELL = 'L5';
  const DEPOT_CELL = 'M3';
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

  function iaExcelSerialDate(d) {
    const dt = d || new Date();
    const utc = Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate());
    return (utc - Date.UTC(1899, 11, 30)) / 86400000;
  }

  function iaXmlEsc(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function iaFflate() {
    const f = typeof fflate !== 'undefined' ? fflate : global.fflate;
    if (!f?.unzip || !f?.zip || !f?.strFromU8 || !f?.strToU8) {
      throw new Error('fflate não carregado — recarrega a página');
    }
    return f;
  }

  function iaBuildSheetPathMap(files) {
    const { strFromU8 } = iaFflate();
    const wb = strFromU8(files['xl/workbook.xml']);
    const rels = strFromU8(files['xl/_rels/workbook.xml.rels']);
    const relMap = {};
    const relRe = /<Relationship[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?>/g;
    let m;
    while ((m = relRe.exec(rels))) {
      const target = m[2].replace(/^\//, '');
      relMap[m[1]] = target.startsWith('xl/') ? target : 'xl/' + target;
    }
    const sheetMap = {};
    const sheetRe = /<sheet[^>]*\bname="([^"]+)"[^>]*\br:id="([^"]+)"[^>]*\/?>/g;
    while ((m = sheetRe.exec(wb))) {
      const path = relMap[m[2]];
      if (path) sheetMap[m[1]] = path;
    }
    return sheetMap;
  }

  function iaExtractRowXml(sheetXml, rowNum) {
    const re = new RegExp('<row r="' + rowNum + '"[\\s\\S]*?</row>');
    return sheetXml.match(re)?.[0] || '';
  }

  function iaRemapRowXml(rowXml, srcRow, dstRow) {
    const src = String(srcRow);
    const dst = String(dstRow);
    let xml = rowXml.replace('r="' + src + '"', 'r="' + dst + '"');
    xml = xml.replace(new RegExp('r="([A-Z]+)' + src + '"', 'g'), 'r="$1' + dst + '"');
    xml = xml.replace(new RegExp('([A-Z]+)' + src + '(?=[:"<\\s])', 'g'), '$1' + dst);
    return xml;
  }

  function iaPatchCellXml(sheetXml, addr, innerXml, typeAttr) {
    const cellRe = new RegExp('<c r="' + addr + '"([^>/]*)(?:/>|>[\\s\\S]*?</c>)');
    const m = sheetXml.match(cellRe);
    if (!m) return sheetXml;
    const attrs = m[1].replace(/\s*t="[^"]*"/g, '') + (typeAttr || '');
    return sheetXml.replace(cellRe, '<c r="' + addr + '"' + attrs + '>' + innerXml + '</c>');
  }

  function iaReplaceCellInner(rowXml, col, rowNum, inner, extraAttrs) {
    const addr = col + rowNum;
    const cellRe = new RegExp('<c r="' + addr + '"([^>/]*)(?:/>|>[\\s\\S]*?</c>)');
    const m = rowXml.match(cellRe);
    if (!m) return rowXml;
    let attrs = m[1];
    if (extraAttrs && /t="/.test(extraAttrs)) attrs = attrs.replace(/\s*t="[^"]*"/g, '');
    attrs += extraAttrs || '';
    return rowXml.replace(cellRe, '<c r="' + addr + '"' + attrs + '>' + inner + '</c>');
  }

  function iaFillDataRow(rowXml, rowNum, row, sheetName, isFirst, lastRow) {
    const jv = row.preco > 0 ? row.preco * row.qt_wms : 0;
    const kv = row.qt_wms - row.qt_sap;
    const mv = kv * (row.preco > 0 ? row.preco : 0);

    rowXml = iaReplaceCellInner(rowXml, 'A', rowNum, '<is><t>' + iaXmlEsc(row.material) + '</t></is>', ' t="inlineStr"');
    rowXml = iaReplaceCellInner(rowXml, 'B', rowNum, '<is><t>' + iaXmlEsc(row.desc) + '</t></is>', ' t="inlineStr"');

    if (row.lote) {
      rowXml = iaReplaceCellInner(rowXml, 'C', rowNum, '<is><t>' + iaXmlEsc(row.lote) + '</t></is>', ' t="inlineStr"');
    } else {
      rowXml = iaReplaceCellInner(rowXml, 'C', rowNum, '', '');
    }

    if (row.preco > 0) {
      rowXml = iaReplaceCellInner(rowXml, 'D', rowNum, '<v>' + row.preco + '</v>', '');
    } else {
      rowXml = iaReplaceCellInner(rowXml, 'D', rowNum, '', '');
    }

    rowXml = iaReplaceCellInner(rowXml, 'E', rowNum, '<is><t>' + iaXmlEsc(row.umb || 'UN') + '</t></is>', ' t="inlineStr"');

    if (isFirst) {
      rowXml = iaReplaceCellInner(
        rowXml,
        'F',
        rowNum,
        '<f t="shared" ref="F' + rowNum + ':F' + lastRow + '" si="0">+$M$3</f><v>' + iaXmlEsc(sheetName) + '</v>',
        ' t="str"'
      );
      rowXml = iaReplaceCellInner(
        rowXml,
        'J',
        rowNum,
        '<f t="shared" ref="J' + rowNum + ':J' + lastRow + '" si="1">D' + rowNum + '*H' + rowNum + '</f><v>' + jv + '</v>',
        ''
      );
      rowXml = iaReplaceCellInner(
        rowXml,
        'K',
        rowNum,
        '<f t="shared" ref="K' + rowNum + ':K' + lastRow + '" si="2">H' + rowNum + '-I' + rowNum + '</f><v>' + kv + '</v>',
        ''
      );
      rowXml = iaReplaceCellInner(
        rowXml,
        'L',
        rowNum,
        '<f t="shared" ref="L' + rowNum + ':L' + lastRow + '" si="3">H' + rowNum + '-I' + rowNum + '</f><v>' + kv + '</v>',
        ''
      );
      rowXml = iaReplaceCellInner(
        rowXml,
        'M',
        rowNum,
        '<f t="shared" ref="M' + rowNum + ':M' + lastRow + '" si="4">L' + rowNum + '*D' + rowNum + '</f><v>' + mv + '</v>',
        ''
      );
    } else {
      rowXml = iaReplaceCellInner(
        rowXml,
        'F',
        rowNum,
        '<f t="shared" si="0"/><v>' + iaXmlEsc(sheetName) + '</v>',
        ' t="str"'
      );
      rowXml = iaReplaceCellInner(rowXml, 'J', rowNum, '<f t="shared" si="1"/><v>' + jv + '</v>', '');
      rowXml = iaReplaceCellInner(rowXml, 'K', rowNum, '<f t="shared" si="2"/><v>' + kv + '</v>', '');
      rowXml = iaReplaceCellInner(rowXml, 'L', rowNum, '<f t="shared" si="3"/><v>' + kv + '</v>', '');
      rowXml = iaReplaceCellInner(rowXml, 'M', rowNum, '<f t="shared" si="4"/><v>' + mv + '</v>', '');
    }

    rowXml = iaReplaceCellInner(rowXml, 'H', rowNum, '<v>' + row.qt_wms + '</v>', '');
    rowXml = iaReplaceCellInner(rowXml, 'I', rowNum, '<v>' + row.qt_sap + '</v>', '');
    rowXml = iaReplaceCellInner(rowXml, 'G', rowNum, '', '');
    return rowXml;
  }

  function iaPatchSheetXml(sheetXml, sheetName, dataRows, dateSerial) {
    sheetXml = iaPatchCellXml(
      sheetXml,
      DEPOT_CELL,
      '<is><t>' + iaXmlEsc(sheetName) + '</t></is>',
      ' t="inlineStr"'
    );
    sheetXml = iaPatchCellXml(sheetXml, DATE_CELL, '<v>' + dateSerial + '</v>', '');

    const protoFirst = iaExtractRowXml(sheetXml, DATA_START_ROW);
    const protoNext = iaExtractRowXml(sheetXml, DATA_START_ROW + 1) || protoFirst;
    if (!protoFirst) throw new Error('Template inválido — falta linha ' + DATA_START_ROW);

    const sheetDataRe = /(<sheetData>)([\s\S]*)(<\/sheetData>)/;
    const sdMatch = sheetXml.match(sheetDataRe);
    if (!sdMatch) throw new Error('Template inválido — sem sheetData');

    const keptRows = [];
    const rowRe = /<row r="(\d+)"[\s\S]*?<\/row>/g;
    let rm;
    while ((rm = rowRe.exec(sdMatch[2]))) {
      if (parseInt(rm[1], 10) < DATA_START_ROW) keptRows.push(rm[0]);
    }

    const lastRow = dataRows.length ? DATA_START_ROW + dataRows.length - 1 : DATA_START_ROW - 1;
    const newDataRows = [];
    dataRows.forEach((row, i) => {
      const rowNum = DATA_START_ROW + i;
      const proto = i === 0 ? protoFirst : protoNext;
      const srcRow = i === 0 ? DATA_START_ROW : DATA_START_ROW + 1;
      let rowXml = iaRemapRowXml(proto, srcRow, rowNum);
      rowXml = iaFillDataRow(rowXml, rowNum, row, sheetName, i === 0, Math.max(lastRow, DATA_START_ROW));
      newDataRows.push(rowXml);
    });

    const newSheetData = sdMatch[1] + keptRows.join('') + newDataRows.join('') + sdMatch[3];
    sheetXml = sheetXml.replace(sheetDataRe, newSheetData);

    sheetXml = sheetXml.replace(/<dimension ref="([^"]+)"/, function (_m, ref) {
      const parts = ref.split(':');
      const end = parts[1] || parts[0];
      const endCol = end.replace(/\d+/g, '');
      const endRow = Math.max(lastRow, DATA_START_ROW - 1);
      return '<dimension ref="' + (parts[0] || 'A1') + ':' + endCol + endRow + '"';
    });

    return sheetXml;
  }

  function iaPatchTemplateXlsx(templateBuf, depotRowsMap, dateSerial) {
    const { unzip, zip, strFromU8, strToU8 } = iaFflate();
    const files = unzip(new Uint8Array(templateBuf));
    const sheetPaths = iaBuildSheetPathMap(files);

    for (const dk of INVENTARIO_ACERTO_DEPOTS) {
      const sheetName = INVENTARIO_ACERTO_SHEETS[dk];
      const path = sheetPaths[sheetName];
      if (!path || !files[path]) continue;
      const xml = strFromU8(files[path]);
      files[path] = strToU8(iaPatchSheetXml(xml, sheetName, depotRowsMap[dk] || [], dateSerial));
    }

    return zip(files);
  }

  function iaDownloadXlsx(buf, filename) {
    const blob = new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
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
      const depotRowsMap = {};
      let skippedNoMaterial = 0;
      for (const dk of INVENTARIO_ACERTO_DEPOTS) {
        const rows = buildAcertoRowsForDepot(dk);
        skippedNoMaterial += rows._skippedNoMaterial || 0;
        depotRowsMap[dk] = rows;
      }
      const out = iaPatchTemplateXlsx(buf, depotRowsMap, iaExcelSerialDate(new Date()));
      iaDownloadXlsx(out, inventarioAcertoExportFilename());
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
