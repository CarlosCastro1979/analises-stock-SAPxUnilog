/* Inventário — doc. de acerto (mesmo Excel template/pipeline do Mapa Acertos / acerto sem valor) */
(function (global) {
  'use strict';

  const INVENTARIO_ACERTO_JS_VERSION = '1.0.10';
  const INVENTARIO_ACERTO_DEPOTS = ['8', '9', '11', '22', '44'];
  const INVENTARIO_ACERTO_SHEETS = { 8: '0008', 9: '0009', 11: '0011', 22: '0022', 44: '0044' };
  /** Excel = assets/mapa_acertos_template.xlsx via maExportPatchedXlsx (cols A–O iguais ao acerto sem valor). */

  const IA_DRILL_LABELS = {
    wms: 'Só Unilog (físico>0, SAP=0)',
    sap: 'Só SAP (físico=0, SAP>0)',
    diverg: 'Divergência (ambos com stock)',
    subir: 'Subir SAP (Δ qty > 0)',
    descer: 'Descer SAP (Δ qty < 0)',
  };

  let inventarioAcertoRows = [];
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

  /** Map inventário rows → Mapa Acertos export shape (same workbook cols A–O). */
  function iaToMapaExportRows(rows) {
    return rows.map(function (r) {
      const adj = Number(r.diff) || 0;
      return {
        depot: r.depot,
        depotSheet: r.depotSheet,
        material: r.material,
        EAN_norm: r.EAN_norm,
        desc: r.desc || '',
        lote: r.lote || '',
        umb: r.umb || 'UN',
        preco: Number(r.preco) || 0,
        qt_uni: Number(r.qt_wms) || 0,
        qt_sap: Number(r.qt_sap) || 0,
        adj: adj,
        subir: adj > 0 ? adj : 0,
        descer: adj < 0 ? -adj : 0,
        justificacao: 'acerto inventário',
      };
    });
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
    if (typeof maExportPatchedXlsx !== 'function') {
      toast('Módulo Mapa Acertos a carregar — aguarde ou recarregue', 'error');
      return;
    }
    const filtered = getInventarioAcertoFilteredRows().slice().sort(function (a, b) {
      const d = String(a.depotSheet).localeCompare(String(b.depotSheet));
      if (d) return d;
      const ma = String(a.material).localeCompare(String(b.material));
      if (ma) return ma;
      return String(a.lote || '').localeCompare(String(b.lote || ''));
    });
    const totalLines = filtered.length;
    if (!totalLines) {
      toast('Sem linhas com diferença nos filtros actuais (ou sem código SAP)', 'info');
      return;
    }
    try {
      const exportRows = iaToMapaExportRows(filtered);
      const result = await maExportPatchedXlsx(exportRows, inventarioAcertoExportFilename());
      const valorTot = filtered.reduce(function (s, r) { return s + (r.valor || 0); }, 0);
      const fmtV = typeof fmtBrlLiq === 'function' ? fmtBrlLiq(valorTot) : String(valorTot);
      let msg = 'Exportado: ' + result.written + ' linhas · valor acerto ' + fmtV;
      const selCats = typeof msGet === 'function' ? msGet('iaCat') : new Set();
      const selSubs = typeof msGet === 'function' ? msGet('iaSub') : new Set();
      if (selCats.size || selSubs.size || iaDrill) msg += ' (filtrado)';
      if (result.truncated) msg += ' · aviso: limite ' + result.maxPerSheet + ' linhas/folha (template)';
      toast(msg, result.truncated ? 'info' : 'success');
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
