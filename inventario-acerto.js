/* Inventário — doc. de acerto (layout PT: folhas 0008, 0009, 0011, 0022, 0044) */
(function (global) {
  'use strict';

  const INVENTARIO_ACERTO_JS_VERSION = '1.0.9';
  const INVENTARIO_ACERTO_DEPOTS = ['8', '9', '11', '22', '44'];
  const INVENTARIO_ACERTO_SHEETS = { 8: '0008', 9: '0009', 11: '0011', 22: '0022', 44: '0044' };
  /** Headers 1–8; data from row 9. Cols: A Mat · B Desc · C Lote · D Preço · E UMB · F Dep · H Unilog · I SAP · J D×H · K/L H−I · M L×D */
  const DATA_START_ROW = 9;
  const DATE_CELL = 'L5';
  const DEPOT_CELL = 'M3';
  const TEMPLATE_URL = 'assets/inventario_acerto_template.xlsx?v=' + INVENTARIO_ACERTO_JS_VERSION;

  const IA_DRILL_LABELS = {
    wms: 'Só Unilog (H>0, I=0)',
    sap: 'Só SAP (H=0, I>0)',
    diverg: 'Divergência (ambos com stock)',
    subir: 'Subir SAP (Δ qty > 0)',
    descer: 'Descer SAP (Δ qty < 0)',
  };

  let inventarioAcertoRows = [];
  let inventarioAcertoTemplateBuf = null;
  /** @type {null|'wms'|'sap'|'diverg'|'subir'|'descer'} */
  let iaDrill = null;

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

  function iaResolveCat(ean, existingCat, existingSub) {
    if (existingCat || existingSub) return { cat: existingCat || '', subcat: existingSub || '' };
    const params = typeof occParams !== 'undefined' ? occParams : [];
    const p = params.find((x) => x.ean === ean);
    return { cat: p?.cat || '', subcat: p?.subcat || '' };
  }

  /** H = físico Unilog, I = SAP contabilístico. */
  function iaAcertoHI(r) {
    if (r.qt_uni > 0 && r.qt_sap === 0) return { h: r.qt_uni, i: 0 };
    if (r.qt_sap > 0 && r.qt_uni === 0) return { h: 0, i: r.qt_sap };
    return { h: r.qt_uni, i: r.qt_sap };
  }

  function iaRowEstado(hi) {
    if (hi.h > 0 && hi.i === 0) return 'Só WMS';
    if (hi.i > 0 && hi.h === 0) return 'Só SAP';
    return 'Diverg.';
  }

  function iaRowDir(diff) {
    if (diff > 0) return 'Subir';
    if (diff < 0) return 'Descer';
    return '—';
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
        const diff = hi.h - hi.i;
        const preco = precoMap[r.EAN_norm] || 0;
        const cs = iaResolveCat(r.EAN_norm, r.cat, r.subcat);
        const excl =
          typeof isExcluidoDifStock === 'function' && isExcluidoDifStock(r.EAN_norm, dk, r.desc);
        return {
          depot: dk,
          depotSheet: INVENTARIO_ACERTO_SHEETS[dk] || dk.padStart(4, '0'),
          depotLabel: iaDepotLabel(dk),
          material,
          EAN_norm: r.EAN_norm,
          desc: r.desc || '',
          lote: r.Lote_norm || '',
          cat: cs.cat,
          subcat: cs.subcat,
          preco,
          umb: getSapUmb(r.EAN_norm),
          qt_wms: hi.h,
          qt_sap: hi.i,
          diff,
          valor: excl || preco <= 0 ? 0 : diff * preco,
          _semValor: !!excl || preco <= 0,
          _estado: iaRowEstado(hi),
          _dir: iaRowDir(diff),
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

  function iaMatchSrch(srch, r) {
    if (!srch) return true;
    const s = srch.toLowerCase();
    return (
      String(r.material || '')
        .toLowerCase()
        .includes(s) ||
      String(r.EAN_norm || '')
        .toLowerCase()
        .includes(s) ||
      String(r.lote || '')
        .toLowerCase()
        .includes(s) ||
      String(r.desc || '')
        .toLowerCase()
        .includes(s)
    );
  }

  function iaApplyDrill(rows) {
    if (!iaDrill) return rows;
    if (iaDrill === 'wms') return rows.filter((r) => r._estado === 'Só WMS');
    if (iaDrill === 'sap') return rows.filter((r) => r._estado === 'Só SAP');
    if (iaDrill === 'diverg') return rows.filter((r) => r._estado === 'Diverg.');
    if (iaDrill === 'subir') return rows.filter((r) => r.diff > 0);
    if (iaDrill === 'descer') return rows.filter((r) => r.diff < 0);
    return rows;
  }

  /** Rows after cat/sub/search/(optional drill) — used by UI and Excel export. */
  function getInventarioAcertoFilteredRows(opts) {
    const skipDrill = !!(opts && opts.skipDrill);
    let rows = buildInventarioAcertoRows();

    const cats = [...new Set(rows.map((r) => r.cat).filter(Boolean))].sort();
    if (typeof msUpdate === 'function') msUpdate('iaCat', cats, 'Categorias', 'renderInventarioAcertoPage');
    const selCats = typeof msGet === 'function' ? msGet('iaCat') : new Set();
    const subSrc = selCats.size ? rows.filter((r) => selCats.has(r.cat)) : rows;
    const subs = [...new Set(subSrc.map((r) => r.subcat).filter(Boolean))].sort();
    if (typeof msUpdate === 'function') msUpdate('iaSub', subs, 'Subcategorias', 'renderInventarioAcertoPage');
    const selSubs = typeof msGet === 'function' ? msGet('iaSub') : new Set();

    if (selCats.size) rows = rows.filter((r) => selCats.has(r.cat));
    if (selSubs.size) rows = rows.filter((r) => selSubs.has(r.subcat));

    const srchEl = document.getElementById('iaSrch');
    const srch = (srchEl?.value || '').trim().toLowerCase();
    if (srch) rows = rows.filter((r) => iaMatchSrch(srch, r));

    if (!skipDrill) rows = iaApplyDrill(rows);
    return rows;
  }

  function iaFmtSignedQty(n) {
    if (n == null || !Number.isFinite(n)) return '—';
    const abs = typeof fmtPtNum === 'function' ? fmtPtNum(Math.abs(n)) : String(Math.abs(n));
    if (n > 0) return '+' + abs;
    if (n < 0) return '-' + abs;
    return abs;
  }

  function iaSetDrill(kind) {
    iaDrill = iaDrill === kind ? null : kind;
    renderInventarioAcertoPage();
  }

  function iaClearDrill() {
    iaDrill = null;
    renderInventarioAcertoPage();
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
    // Require Sync APIs — async zip returns a Promise and produces a corrupt download.
    const unzip = f && f.unzipSync;
    const zip = f && f.zipSync;
    if (!unzip || !zip || !f.strFromU8 || !f.strToU8) {
      throw new Error('fflate não carregado — recarrega a página');
    }
    return { unzip: unzip, zip: zip, strFromU8: f.strFromU8, strToU8: f.strToU8 };
  }

  /** Windows Zip may store backslashes; fflate keeps them → xl/workbook.xml lookups miss. */
  function iaNormalizeZipFiles(files) {
    const out = {};
    Object.keys(files).forEach(function (k) {
      const nk = String(k).replace(/\\/g, '/');
      if (!out[nk] || k === nk) out[nk] = files[k];
    });
    return out;
  }
  function iaZipGet(files, path) {
    if (files[path]) return files[path];
    return files[path.replace(/\//g, '\\')] || null;
  }

  function iaBuildSheetPathMap(files) {
    const { strFromU8 } = iaFflate();
    const wbU8 = iaZipGet(files, 'xl/workbook.xml');
    const relsU8 = iaZipGet(files, 'xl/_rels/workbook.xml.rels');
    if (!wbU8 || !relsU8) {
      throw new Error('Template ZIP inválido — workbook.xml / rels em falta');
    }
    const wb = strFromU8(wbU8);
    const rels = strFromU8(relsU8);
    const relMap = {};
    let m;
    const relRe = /<Relationship[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?>/g;
    while ((m = relRe.exec(rels))) {
      var target = m[2].replace(/^\//, '').replace(/\\/g, '/');
      relMap[m[1]] = target.indexOf('xl/') === 0 ? target : 'xl/' + target;
    }
    const relRe2 = /<Relationship[^>]*\bTarget="([^"]+)"[^>]*\bId="([^"]+)"[^>]*\/?>/g;
    while ((m = relRe2.exec(rels))) {
      if (relMap[m[2]]) continue;
      var t2 = m[1].replace(/^\//, '').replace(/\\/g, '/');
      relMap[m[2]] = t2.indexOf('xl/') === 0 ? t2 : 'xl/' + t2;
    }
    const sheetMap = {};
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

  function iaCellStyleAttr(rowXml, col, rowNum) {
    const m = rowXml.match(new RegExp('<c r="' + col + rowNum + '"([^>/]*)'));
    if (!m) return '';
    const sm = m[1].match(/\bs="(\d+)"/);
    return sm ? ' s="' + sm[1] + '"' : '';
  }

  function iaReplaceWholeCell(rowXml, col, rowNum, cellXml) {
    const re = new RegExp('<c r="' + col + rowNum + '"(?:[^>/]*/>|[^>]*>(?:[\\s\\S]*?)</c>)');
    if (re.test(rowXml)) return rowXml.replace(re, cellXml);
    return rowXml.replace('</row>', cellXml + '</row>');
  }

  function iaPatchCellXml(sheetXml, addr, innerXml, typeAttr) {
    const cellRe = new RegExp('<c r="' + addr + '"([^>/]*)(?:/>|>[\\s\\S]*?</c>)');
    const m = sheetXml.match(cellRe);
    if (!m) return sheetXml;
    const attrs = m[1].replace(/\s*t="[^"]*"/g, '') + (typeAttr || '');
    return sheetXml.replace(cellRe, '<c r="' + addr + '"' + attrs + '>' + innerXml + '</c>');
  }

  /**
   * Fill one data row with plain values (no shared formulas).
   * Shared-formula rewrite + stale xl/calcChain.xml was corrupting the xlsx
   * (Excel: "Só de Leitura - Reparado" / empty sheets).
   */
  function iaFillDataRow(rowXml, rowNum, row, sheetName) {
    // Drop any formula nodes left from the prototype row
    rowXml = rowXml.replace(/<f\b[^>]*\/>/g, '').replace(/<f\b[^>]*>[\s\S]*?<\/f>/g, '');

    const preco = Number(row.preco) > 0 ? Number(row.preco) : 0;
    const h = Number(row.qt_wms) || 0;
    const i = Number(row.qt_sap) || 0;
    const jv = preco > 0 ? preco * h : 0;
    const kv = h - i;
    const mv = preco > 0 ? kv * preco : 0;
    const st = function (col) {
      return iaCellStyleAttr(rowXml, col, rowNum);
    };

    rowXml = iaReplaceWholeCell(
      rowXml,
      'A',
      rowNum,
      '<c r="A' + rowNum + '"' + st('A') + ' t="inlineStr"><is><t>' + iaXmlEsc(row.material) + '</t></is></c>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'B',
      rowNum,
      '<c r="B' + rowNum + '"' + st('B') + ' t="inlineStr"><is><t>' + iaXmlEsc(row.desc || '') + '</t></is></c>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'C',
      rowNum,
      row.lote
        ? '<c r="C' + rowNum + '"' + st('C') + ' t="inlineStr"><is><t>' + iaXmlEsc(row.lote) + '</t></is></c>'
        : '<c r="C' + rowNum + '"' + st('C') + '/>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'D',
      rowNum,
      preco > 0
        ? '<c r="D' + rowNum + '"' + st('D') + '><v>' + preco + '</v></c>'
        : '<c r="D' + rowNum + '"' + st('D') + '/>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'E',
      rowNum,
      '<c r="E' + rowNum + '"' + st('E') + ' t="inlineStr"><is><t>' + iaXmlEsc(row.umb || 'UN') + '</t></is></c>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'F',
      rowNum,
      '<c r="F' + rowNum + '"' + st('F') + ' t="inlineStr"><is><t>' + iaXmlEsc(sheetName) + '</t></is></c>'
    );
    rowXml = iaReplaceWholeCell(rowXml, 'G', rowNum, '<c r="G' + rowNum + '"' + st('G') + '/>');
    rowXml = iaReplaceWholeCell(
      rowXml,
      'H',
      rowNum,
      '<c r="H' + rowNum + '"' + st('H') + '><v>' + h + '</v></c>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'I',
      rowNum,
      '<c r="I' + rowNum + '"' + st('I') + '><v>' + i + '</v></c>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'J',
      rowNum,
      preco > 0
        ? '<c r="J' + rowNum + '"' + st('J') + '><v>' + jv + '</v></c>'
        : '<c r="J' + rowNum + '"' + st('J') + '/>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'K',
      rowNum,
      '<c r="K' + rowNum + '"' + st('K') + '><v>' + kv + '</v></c>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'L',
      rowNum,
      '<c r="L' + rowNum + '"' + st('L') + '><v>' + kv + '</v></c>'
    );
    rowXml = iaReplaceWholeCell(
      rowXml,
      'M',
      rowNum,
      preco > 0
        ? '<c r="M' + rowNum + '"' + st('M') + '><v>' + mv + '</v></c>'
        : '<c r="M' + rowNum + '"' + st('M') + '/>'
    );
    return rowXml;
  }

  function iaVerifySheetPatched(sheetXml, sheetName, dataRows) {
    if (!dataRows.length) return;
    const a9 = sheetXml.match(
      new RegExp('<c r="A' + DATA_START_ROW + '"[^>]*>[\\s\\S]*?</c>|<c r="A' + DATA_START_ROW + '"[^>]*/>')
    );
    const hasMat = a9 && /<(?:v|t|is)>/.test(a9[0]);
    if (!hasMat) {
      throw new Error('Export sheet ' + sheetName + ': A' + DATA_START_ROW + ' sem dados após patch — abortado');
    }
    const rowRe = /<row r="(\d+)"[\s\S]*?<\/row>/g;
    let m;
    while ((m = rowRe.exec(sheetXml))) {
      if (parseInt(m[1], 10) < DATA_START_ROW) continue;
      if (/t="shared"/.test(m[0])) {
        throw new Error('Export sheet ' + sheetName + ': fórmulas partilhadas residuais — abortado');
      }
    }
  }

  function iaPatchSheetXml(sheetXml, sheetName, dataRows, dateSerial) {
    sheetXml = iaPatchCellXml(
      sheetXml,
      DEPOT_CELL,
      '<is><t>' + iaXmlEsc(sheetName) + '</t></is>',
      ' t="inlineStr"'
    );
    sheetXml = iaPatchCellXml(sheetXml, DATE_CELL, '<v>' + dateSerial + '</v>', '');

    // Stale autoFilter ranges from the PT sample cause Excel repair noise
    sheetXml = sheetXml.replace(/<autoFilter\b[^>]*\/>/g, '');
    sheetXml = sheetXml.replace(/<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/g, '');

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
    dataRows.forEach(function (row, i) {
      const rowNum = DATA_START_ROW + i;
      const proto = i === 0 ? protoFirst : protoNext;
      const srcRow = i === 0 ? DATA_START_ROW : DATA_START_ROW + 1;
      let rowXml = iaRemapRowXml(proto, srcRow, rowNum);
      rowXml = iaFillDataRow(rowXml, rowNum, row, sheetName);
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

    iaVerifySheetPatched(sheetXml, sheetName, dataRows);
    return sheetXml;
  }

  /** Drop calcChain — it still points at template shared formulas after we rewrite sheetData. */
  function iaStripCalcChain(files, strFromU8, strToU8) {
    Object.keys(files).forEach(function (k) {
      if (/calcChain\.xml$/i.test(String(k).replace(/\\/g, '/'))) delete files[k];
    });
    const relsPath = 'xl/_rels/workbook.xml.rels';
    const relsU8 = iaZipGet(files, relsPath);
    if (relsU8) {
      let rels = strFromU8(relsU8);
      rels = rels.replace(/<Relationship[^>]*calcChain[^>]*\/?>/gi, '');
      files[relsPath] = strToU8(rels);
    }
    const ctPath = '[Content_Types].xml';
    const ctU8 = iaZipGet(files, ctPath);
    if (ctU8) {
      let ct = strFromU8(ctU8);
      ct = ct.replace(/<Override[^>]*calcChain[^>]*\/>/gi, '');
      files[ctPath] = strToU8(ct);
    }
  }

  function iaStripStaleDefinedNames(files, strFromU8, strToU8) {
    const wbPath = 'xl/workbook.xml';
    const wbU8 = iaZipGet(files, wbPath);
    if (!wbU8) return;
    let wb = strFromU8(wbU8);
    wb = wb.replace(/<definedNames>[\s\S]*?<\/definedNames>/g, '<definedNames/>');
    files[wbPath] = strToU8(wb);
  }

  function iaPatchTemplateXlsx(templateBuf, depotRowsMap, dateSerial) {
    const { unzip, zip, strFromU8, strToU8 } = iaFflate();
    const files = iaNormalizeZipFiles(unzip(new Uint8Array(templateBuf)));
    const sheetPaths = iaBuildSheetPathMap(files);
    if (!Object.keys(sheetPaths).length) {
      throw new Error('Template sem folhas mapeadas (workbook sheet/rId)');
    }

    var patched = 0;
    for (const dk of INVENTARIO_ACERTO_DEPOTS) {
      const sheetName = INVENTARIO_ACERTO_SHEETS[dk];
      const path = sheetPaths[sheetName];
      const sheetU8 = path ? iaZipGet(files, path) : null;
      if (!path || !sheetU8) {
        if ((depotRowsMap[dk] || []).length) {
          throw new Error('Folha ' + sheetName + ' em falta no template ZIP');
        }
        continue;
      }
      files[path] = strToU8(
        iaPatchSheetXml(strFromU8(sheetU8), sheetName, depotRowsMap[dk] || [], dateSerial)
      );
      patched++;
    }
    if (!patched) throw new Error('Nenhuma folha do Acerto Inventário foi escrita');

    iaStripCalcChain(files, strFromU8, strToU8);
    iaStripStaleDefinedNames(files, strFromU8, strToU8);

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
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return 'Inventario_DFB_acerto_inventario_' + yyyy + '-' + mm + '-' + dd + '.xlsx';
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
    const filtered = getInventarioAcertoFilteredRows();
    const totalLines = filtered.length;
    if (!totalLines) {
      toast('Sem linhas com diferença nos filtros actuais (ou sem código SAP)', 'info');
      return;
    }
    try {
      // Always re-fetch template (avoid stale ArrayBuffer)
      inventarioAcertoTemplateBuf = null;
      const buf = await loadInventarioAcertoTemplate();
      const depotRowsMap = {};
      for (const dk of INVENTARIO_ACERTO_DEPOTS) depotRowsMap[dk] = [];
      for (const r of filtered) {
        if (!depotRowsMap[r.depot]) depotRowsMap[r.depot] = [];
        depotRowsMap[r.depot].push(r);
      }
      for (const dk of INVENTARIO_ACERTO_DEPOTS) {
        depotRowsMap[dk].sort((a, b) => {
          const ma = a.material.localeCompare(b.material);
          if (ma) return ma;
          return (a.lote || '').localeCompare(b.lote || '');
        });
      }
      const out = iaPatchTemplateXlsx(buf, depotRowsMap, iaExcelSerialDate(new Date()));
      iaDownloadXlsx(out, inventarioAcertoExportFilename());
      const valorTot = filtered.reduce((s, r) => s + (r.valor || 0), 0);
      const fmtV =
        typeof fmtBrlLiq === 'function' ? fmtBrlLiq(valorTot) : String(valorTot);
      let msg = 'Exportado: ' + totalLines + ' linhas · valor acerto ' + fmtV;
      const selCats = typeof msGet === 'function' ? msGet('iaCat') : new Set();
      const selSubs = typeof msGet === 'function' ? msGet('iaSub') : new Set();
      if (selCats.size || selSubs.size || iaDrill) msg += ' (filtrado)';
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
    // Saldo líquido assinado (mesmo sinal que o valor) — subir e descer anulam-se
    const units = rows.reduce((s, r) => s + (Number(r.diff) || 0), 0);
    const valorTot = rows.reduce((s, r) => s + (r.valor || 0), 0);
    const fmtV = typeof fmtKvBrlLiq === 'function' ? fmtKvBrlLiq : (v) => String(v);
    const fmtQ =
      typeof fmtKv === 'function' ? (n) => fmtKv(iaFmtSignedQty(n)) : iaFmtSignedQty;
    const liqCls = typeof kvLiqCls === 'function' ? kvLiqCls(valorTot) : '';
    const qtyCls = typeof kvLiqCls === 'function' ? kvLiqCls(units) : '';

    el.innerHTML =
      '<div class="kpi" style="border-color:' +
      (valorTot > 0 ? 'var(--yellow)' : valorTot < 0 ? 'var(--red)' : 'var(--border)') +
      ';grid-column:span 2">' +
      '<div class="kl">Valor final do acerto</div>' +
      '<div class="kv ' +
      liqCls +
      '">' +
      fmtV(valorTot) +
      '</div>' +
      '<div class="ks">Σ (Unilog−SAP)×preço SAP · vista filtrada · + = subir stock SAP</div></div>' +
      '<div class="kpi" style="border-color:' +
      (units > 0 ? 'var(--yellow)' : units < 0 ? 'var(--red)' : 'var(--border)') +
      ';grid-column:span 2">' +
      '<div class="kl">Saldo final (qtd)</div>' +
      '<div class="kv ' +
      qtyCls +
      '">' +
      fmtQ(units) +
      '</div>' +
      '<div class="ks">Σ (Unilog−SAP) · vista filtrada · + = subir stock SAP</div></div>';
  }

  function renderInventarioAcertoTable(rows) {
    const body = document.getElementById('iaBody');
    const empty = document.getElementById('iaEmpty');
    const emptyMsg = document.getElementById('iaEmptyMsg');
    const wrap = document.getElementById('iaTableWrap');
    if (!body) return;
    if (!rows.length) {
      body.innerHTML = '';
      if (empty) empty.style.display = 'block';
      if (wrap) wrap.style.display = 'none';
      if (emptyMsg) {
        const hasFilt =
          iaDrill ||
          (typeof msGet === 'function' && (msGet('iaCat').size || msGet('iaSub').size)) ||
          (document.getElementById('iaSrch')?.value || '').trim();
        emptyMsg.textContent = hasFilt
          ? 'Nenhuma linha com os filtros activos — limpa categoria/subcategoria ou o filtro do resumo'
          : 'Carrega SAP e Unilog e processa — ou sem diferenças nos 5 depósitos';
      }
      return;
    }
    if (empty) empty.style.display = 'none';
    if (wrap) wrap.style.display = '';
    const sorted = typeof sorts !== 'undefined' && sorts.ia ? rows.slice() : rows.slice();
    if (typeof sorts !== 'undefined' && sorts.ia && typeof applySort === 'function') applySort('ia', sorted);
    else sorted.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));

    const fmtNum = typeof fmtPtNum === 'function' ? fmtPtNum : (n) => n;
    body.innerHTML = sorted
      .map((r) => {
        const estadoBadge =
          r._estado === 'Só WMS'
            ? '<span class="badge binfo">Só Unilog</span>'
            : r._estado === 'Só SAP'
              ? '<span class="badge bwarn">Só SAP</span>'
              : '<span class="badge bwarn">Diverg.</span>';
        const dirBadge =
          r.diff > 0
            ? '<span class="badge bok">Subir</span>'
            : r.diff < 0
              ? '<span class="badge berr">Descer</span>'
              : '—';
        const diffCls = r.diff > 0 ? 'pos' : r.diff < 0 ? 'neg' : '';
        const valorCls = r.valor > 0 ? 'pos' : r.valor < 0 ? 'neg' : '';
        const valorTxt =
          r._semValor || !(r.preco > 0)
            ? '—'
            : typeof fmtBrlLiq === 'function'
              ? fmtBrlLiq(r.valor)
              : r.valor;
        return (
          '<tr>' +
          '<td><span class="badge bgray">' +
          r.depotSheet +
          '</span></td>' +
          '<td>' +
          (r.cat || '—') +
          '</td>' +
          '<td>' +
          (r.subcat || '—') +
          '</td>' +
          '<td style="font-family:var(--mono);font-size:11px">' +
          r.material +
          '</td>' +
          '<td style="font-family:var(--mono);font-size:11px">' +
          (r.EAN_norm || '—') +
          '</td>' +
          '<td class="col-desc">' +
          (r.desc || '—') +
          '</td>' +
          '<td style="font-family:var(--mono);font-size:11px">' +
          (r.lote || '—') +
          '</td>' +
          '<td class="num">' +
          fmtNum(r.qt_sap) +
          '</td>' +
          '<td class="num">' +
          fmtNum(r.qt_wms) +
          '</td>' +
          '<td class="num ' +
          diffCls +
          '">' +
          iaFmtSignedQty(r.diff) +
          '</td>' +
          '<td class="num ' +
          valorCls +
          '">' +
          valorTxt +
          '</td>' +
          '<td>' +
          dirBadge +
          '</td>' +
          '<td>' +
          estadoBadge +
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

    const beforeDrill = getInventarioAcertoFilteredRows({ skipDrill: true });
    const rows = iaApplyDrill(beforeDrill);

    const banner = document.getElementById('iaFilterBanner');
    const lbl = document.getElementById('iaFilterLbl');
    if (banner && lbl) {
      if (iaDrill) {
        banner.style.display = 'flex';
        lbl.textContent = IA_DRILL_LABELS[iaDrill] || iaDrill;
      } else {
        banner.style.display = 'none';
      }
    }

    renderInventarioAcertoKpis(rows);
    renderInventarioAcertoTable(rows);

    const depCounts = INVENTARIO_ACERTO_DEPOTS.map((dk) => {
      const n = rows.filter((r) => r.depot === dk).length;
      return INVENTARIO_ACERTO_SHEETS[dk] + ': ' + n;
    }).join(' · ');
    const sub = document.getElementById('iaDepotCounts');
    if (sub) sub.textContent = depCounts;
  }

  function syncInventarioAcertoNav() {
    const it = document.getElementById('it-inventario-acerto');
    if (!it) return;
    const show = typeof company !== 'undefined' && company === 'DFB';
    it.style.display = show ? '' : 'none';
    // If QB hides this tab while it is active, fall back to Por Lote
    if (!show && it.classList.contains('active')) {
      const loteTab = document.querySelector('#reconIts .it');
      if (loteTab && typeof reconTab === 'function') reconTab('lote', loteTab);
    }
  }

  function iaDoSort(col) {
    if (typeof doSort === 'function') doSort('ia', col);
  }

  global.INVENTARIO_ACERTO_JS_VERSION = INVENTARIO_ACERTO_JS_VERSION;
  global.INVENTARIO_ACERTO_DEPOTS = INVENTARIO_ACERTO_DEPOTS;
  global.buildInventarioAcertoRows = buildInventarioAcertoRows;
  global.getInventarioAcertoFilteredRows = getInventarioAcertoFilteredRows;
  global.renderInventarioAcertoPage = renderInventarioAcertoPage;
  global.exportInventarioAcertoXlsx = exportInventarioAcertoXlsx;
  global.inventarioAcertoExportFilename = inventarioAcertoExportFilename;
  global.syncInventarioAcertoNav = syncInventarioAcertoNav;
  global.iaDoSort = iaDoSort;
  global.iaSetDrill = iaSetDrill;
  global.iaClearDrill = iaClearDrill;
})(typeof window !== 'undefined' ? window : globalThis);
