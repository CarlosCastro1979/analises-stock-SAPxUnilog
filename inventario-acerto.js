/* Inventário — doc. de acerto (mesmo Excel template/pipeline do Mapa Acertos / acerto sem valor)
 * Duas vistas:
 *  - Full (ia): SAP = Unilog 100% (todas as classes, incluindo FI)
 *  - Sem FI (ias): SAP = Unilog sem classe FI
 * Excluídos do acerto: máquinas/equip. profissionais (+ caixas de papelão via isExcluidoDifStock). */
(function (global) {
  'use strict';

  const INVENTARIO_ACERTO_JS_VERSION = '1.1.4';
  const INVENTARIO_ACERTO_DEPOTS = ['8', '99', '11', '22', '44'];
  const INVENTARIO_ACERTO_SHEETS = { 8: '0008', 99: '0099', 11: '0011', 22: '0022', 44: '0044' };

  const IA_DRILL_LABELS = {
    wms: 'Só Unilog (físico>0, SAP=0)',
    sap: 'Só SAP (físico=0, SAP>0)',
    diverg: 'Divergência (ambos com stock)',
    subir: 'Subir SAP (Δ qty > 0)',
    descer: 'Descer SAP (Δ qty < 0)',
    semMat: 'Sem código SAP (não exportável)',
  };

  function iaDepotLabel(dk) {
    const m = typeof DEPOT_MAP !== 'undefined' ? DEPOT_MAP : {};
    return m[dk]?.label || dk;
  }

  function resolveSapMaterial(reconRow, depotKey) {
    const ean = reconRow?.EAN_norm;
    if (!ean) return '';
    if (reconRow.material) return String(reconRow.material).trim();
    if (typeof buildEanSapCodeMap === 'function') buildEanSapCodeMap();
    const sap = typeof sapData !== 'undefined' ? sapData : [];
    const depotHit = sap.find((x) => x.deposito === depotKey && x.EAN_norm === ean && x.material);
    if (depotHit) return String(depotHit.material).trim();
    if (typeof _eanSapCodeMap !== 'undefined' && _eanSapCodeMap?.[ean]?.size) return [..._eanSapCodeMap[ean]][0];
    const anyEan = sap.find((x) => x.EAN_norm === ean && x.material);
    if (anyEan) return String(anyEan.material).trim();
    const asMaterial = sap.find((x) => x.material === ean);
    if (asMaterial) return String(asMaterial.material).trim();
    if (typeof getManualSapMaterial === 'function') {
      const ov = getManualSapMaterial(ean);
      if (ov) return ov;
    }
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

  /** Físico Unilog alvo do acerto; se excludeFI, remove qty da classe FI (falta de inventário). */
  function iaPhysicalQty(r, excludeFI) {
    let h = Number(r.qt_uni) || 0;
    if (excludeFI) {
      const by = r.qt_uni_by_class || {};
      const fi = Number(by.FI) || Number(by.fi) || 0;
      h = Math.max(0, h - fi);
    }
    return h;
  }

  /** Δ = Unilog(alvo) − SAP — subir SAP quando WMS > SAP. */
  function iaAcertoHI(r, excludeFI) {
    const h = iaPhysicalQty(r, excludeFI);
    const i = Number(r.qt_sap) || 0;
    return { h, i };
  }

  /** Máquinas profissionais (+ caixas papelão) — fora do acerto; não exigir no SAP. */
  function iaIsExcludedFromAcerto(r, dk) {
    if (typeof isExcluidoDifStock === 'function') return !!isExcluidoDifStock(r.EAN_norm, dk, r.desc);
    if (typeof isEquipProfissional === 'function') return !!isEquipProfissional(r.EAN_norm, dk);
    return false;
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

  function iaFmtSignedQty(n) {
    if (n == null || !Number.isFinite(n)) return '—';
    const abs = typeof fmtPtNum === 'function' ? fmtPtNum(Math.abs(n)) : String(Math.abs(n));
    if (n > 0) return '+' + abs;
    if (n < 0) return '-' + abs;
    return abs;
  }

  function buildAcertoRowsForDepot(dk, excludeFI) {
    const rows = (typeof reconData !== 'undefined' ? reconData[dk] : null) || [];
    const precoMap = typeof buildSapPrecoMap === 'function' ? buildSapPrecoMap() : {};
    let countSemMaterial = 0;
    let skippedExcluded = 0;
    const out = rows
      .map((r) => {
        if (iaIsExcludedFromAcerto(r, dk)) {
          skippedExcluded++;
          return null;
        }
        const hi = iaAcertoHI(r, excludeFI);
        const diff = hi.h - hi.i;
        if (diff === 0) return null;
        const material = resolveSapMaterial(r, dk);
        const semMat = !material;
        if (semMat) countSemMaterial++;
        const preco = precoMap[r.EAN_norm] || 0;
        const cs = iaResolveCat(r.EAN_norm, r.cat, r.subcat);
        const fiQty = Number(r.qt_uni_by_class?.FI) || Number(r.qt_uni_by_class?.fi) || 0;
        return {
          depot: dk,
          depotSheet: INVENTARIO_ACERTO_SHEETS[dk] || dk.padStart(4, '0'),
          depotLabel: iaDepotLabel(dk),
          material: material || '',
          EAN_norm: r.EAN_norm,
          desc: r.desc || '',
          lote: r.Lote_norm || '',
          cat: cs.cat,
          subcat: cs.subcat,
          preco,
          umb: getSapUmb(r.EAN_norm),
          qt_wms: hi.h,
          qt_sap: hi.i,
          qt_fi_excluded: excludeFI ? fiQty : 0,
          diff,
          valor: preco <= 0 ? 0 : diff * preco,
          _semValor: preco <= 0,
          _semMaterial: semMat,
          _estado: iaRowEstado(hi),
          _dir: iaRowDir(diff),
        };
      })
      .filter((r) => !!r)
      .sort((a, b) => {
        // Com material primeiro; sem cód. SAP no fim (ainda visíveis)
        if (!!a._semMaterial !== !!b._semMaterial) return a._semMaterial ? 1 : -1;
        const ma = String(a.material || '').localeCompare(String(b.material || ''));
        if (ma) return ma;
        return (a.lote || '').localeCompare(b.lote || '');
      });
    out._skippedNoMaterial = countSemMaterial; // legado: contagem sem material (agora visíveis)
    out._countSemMaterial = countSemMaterial;
    out._skippedExcluded = skippedExcluded;
    return out;
  }

  function iaMatchSrch(srch, r) {
    if (!srch) return true;
    const s = srch.toLowerCase();
    return (
      String(r.material || '').toLowerCase().includes(s) ||
      String(r.EAN_norm || '').toLowerCase().includes(s) ||
      String(r.lote || '').toLowerCase().includes(s) ||
      String(r.desc || '').toLowerCase().includes(s)
    );
  }

  function iaApplyDrill(rows, drill) {
    if (!drill) return rows;
    if (drill === 'wms') return rows.filter((r) => r._estado === 'Só WMS');
    if (drill === 'sap') return rows.filter((r) => r._estado === 'Só SAP');
    if (drill === 'diverg') return rows.filter((r) => r._estado === 'Diverg.');
    if (drill === 'subir') return rows.filter((r) => r.diff > 0);
    if (drill === 'descer') return rows.filter((r) => r.diff < 0);
    if (drill === 'semMat') return rows.filter((r) => r._semMaterial);
    return rows;
  }

  function iaToMapaExportRows(rows, note) {
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
        qt_sap_after: (Number(r.qt_sap) || 0) + adj,
        adj: adj,
        subir: adj > 0 ? adj : 0,
        descer: adj < 0 ? -adj : 0,
        justificacao: note || 'acerto inventário',
      };
    });
  }

  /**
   * @param {object} cfg
   * @param {string} cfg.prefix DOM id prefix (ia | ias)
   * @param {boolean} cfg.excludeFI
   * @param {string} cfg.sortKey sorts table id
   * @param {string} cfg.navId tab button id
   * @param {string} cfg.panelId panel id
   * @param {string} cfg.fileStem export filename stem
   * @param {string} cfg.justify Excel note
   */
  function createInventarioAcertoView(cfg) {
    const prefix = cfg.prefix;
    const excludeFI = !!cfg.excludeFI;
    const sortKey = cfg.sortKey;
    const catMs = prefix + 'Cat';
    const subMs = prefix + 'Sub';
    let drill = null;
    let cachedRows = [];

    function el(id) {
      return document.getElementById(prefix + id);
    }

    function buildRows() {
      if (typeof company !== 'undefined' && company !== 'DFB') return [];
      const out = [];
      let countSemMaterial = 0;
      let skippedExcluded = 0;
      for (const dk of INVENTARIO_ACERTO_DEPOTS) {
        const part = buildAcertoRowsForDepot(dk, excludeFI);
        countSemMaterial += part._countSemMaterial || 0;
        skippedExcluded += part._skippedExcluded || 0;
        out.push(...part);
      }
      out._countSemMaterial = countSemMaterial;
      out._skippedExcluded = skippedExcluded;
      cachedRows = out;
      return out;
    }

    function getFilteredRows(opts) {
      const skipDrill = !!(opts && opts.skipDrill);
      let rows = buildRows();
      const cats = [...new Set(rows.map((r) => r.cat).filter(Boolean))].sort();
      if (typeof msUpdate === 'function') msUpdate(catMs, cats, 'Categorias', cfg.renderName);
      const selCats = typeof msGet === 'function' ? msGet(catMs) : new Set();
      const subSrc = selCats.size ? rows.filter((r) => selCats.has(r.cat)) : rows;
      const subs = [...new Set(subSrc.map((r) => r.subcat).filter(Boolean))].sort();
      if (typeof msUpdate === 'function') msUpdate(subMs, subs, 'Subcategorias', cfg.renderName);
      const selSubs = typeof msGet === 'function' ? msGet(subMs) : new Set();
      if (selCats.size) rows = rows.filter((r) => selCats.has(r.cat));
      if (selSubs.size) rows = rows.filter((r) => selSubs.has(r.subcat));
      const srch = (el('Srch')?.value || '').trim().toLowerCase();
      if (srch) rows = rows.filter((r) => iaMatchSrch(srch, r));
      if (!skipDrill) rows = iaApplyDrill(rows, drill);
      return rows;
    }

    function setDrill(kind) {
      drill = drill === kind ? null : kind;
      render();
    }
    function clearDrill() {
      drill = null;
      render();
    }

    function exportFilename() {
      const d = new Date();
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      const dd = String(d.getDate()).padStart(2, '0');
      return cfg.fileStem + '_' + yyyy + '-' + mm + '-' + dd + '.xlsx';
    }

    async function exportXlsx() {
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
      const filteredAll = getFilteredRows().slice().sort(function (a, b) {
        const d = String(a.depotSheet).localeCompare(String(b.depotSheet));
        if (d) return d;
        const ma = String(a.material || '').localeCompare(String(b.material || ''));
        if (ma) return ma;
        return String(a.lote || '').localeCompare(String(b.lote || ''));
      });
      const skippedSemMat = filteredAll.filter(function (r) { return r._semMaterial; }).length;
      const filtered = filteredAll.filter(function (r) { return !r._semMaterial && r.material; });
      if (!filtered.length) {
        toast(
          skippedSemMat
            ? 'Nada exportável: ' + skippedSemMat + ' linha(s) sem código SAP (cria material no SAP ou mapeia EAN)'
            : 'Sem linhas com diferença nos filtros actuais',
          'info'
        );
        return;
      }
      try {
        const exportRows = iaToMapaExportRows(filtered, cfg.justify);
        const result = await maExportPatchedXlsx(exportRows, exportFilename());
        const valorTot = filtered.reduce(function (s, r) { return s + (r.valor || 0); }, 0);
        const fmtV = typeof fmtBrlLiq === 'function' ? fmtBrlLiq(valorTot) : String(valorTot);
        let msg = 'Exportado: ' + result.written + ' linhas · valor acerto ' + fmtV;
        if (skippedSemMat) msg += ' · ' + skippedSemMat + ' sem cód. SAP omitidas';
        const selCats = typeof msGet === 'function' ? msGet(catMs) : new Set();
        const selSubs = typeof msGet === 'function' ? msGet(subMs) : new Set();
        if (selCats.size || selSubs.size || drill) msg += ' (filtrado)';
        if (excludeFI) msg += ' · sem FI';
        if (result.truncated) msg += ' · aviso: limite ' + result.maxPerSheet + ' linhas/folha (template)';
        toast(msg, result.truncated || skippedSemMat ? 'info' : 'success');
      } catch (e) {
        toast('Erro ao exportar: ' + (e.message || e), 'error');
      }
    }

    function renderKpis(rows) {
      const kpiEl = el('Kpis');
      if (!kpiEl) return;
      if (!rows.length) {
        kpiEl.innerHTML = '';
        return;
      }
      const exportable = rows.filter((r) => !r._semMaterial);
      const semMatN = rows.length - exportable.length;
      const units = exportable.reduce((s, r) => s + (Number(r.diff) || 0), 0);
      const valorTot = exportable.reduce((s, r) => s + (r.valor || 0), 0);
      const fmtV = typeof fmtKvBrlLiq === 'function' ? fmtKvBrlLiq : (v) => String(v);
      const fmtQ = typeof fmtKv === 'function' ? (n) => fmtKv(iaFmtSignedQty(n)) : iaFmtSignedQty;
      const fmtN = typeof fmtKvNum === 'function' ? fmtKvNum : (n) => String(n);
      const liqCls = typeof kvLiqCls === 'function' ? kvLiqCls(valorTot) : '';
      const qtyCls = typeof kvLiqCls === 'function' ? kvLiqCls(units) : '';
      const alvoNote = excludeFI
        ? ' · alvo = Unilog sem FI (SAP = WMS excl. falta de inventário)'
        : ' · alvo = Unilog 100% incl. FI (SAP = WMS)';
      const semMatKpi = semMatN
        ? '<div class="kpi" style="border-color:var(--yellow);cursor:pointer" onclick="typeof ' +
          (prefix === 'ias' ? 'iasSetDrill' : 'iaSetDrill') +
          '===\'function\'&&' +
          (prefix === 'ias' ? 'iasSetDrill' : 'iaSetDrill') +
          '(\'semMat\')"><div class="kl">Sem código SAP</div><div class="kv y">' +
          fmtN(semMatN) +
          '</div><div class="ks">visíveis na tabela · não entram no Excel · clica para filtrar</div></div>'
        : '';
      kpiEl.innerHTML =
        '<div class="kpi" style="border-color:' +
        (valorTot > 0 ? 'var(--yellow)' : valorTot < 0 ? 'var(--red)' : 'var(--border)') +
        ';grid-column:span 2">' +
        '<div class="kl">Valor final do acerto</div>' +
        '<div class="kv ' + liqCls + '">' + fmtV(valorTot) + '</div>' +
        '<div class="ks">Σ (Unilog−SAP)×preço SAP · só com cód. SAP · + = subir stock SAP' +
        alvoNote +
        '</div></div>' +
        '<div class="kpi" style="border-color:' +
        (units > 0 ? 'var(--yellow)' : units < 0 ? 'var(--red)' : 'var(--border)') +
        ';grid-column:span 2">' +
        '<div class="kl">Saldo final (qtd)</div>' +
        '<div class="kv ' + qtyCls + '">' + fmtQ(units) + '</div>' +
        '<div class="ks">Σ (Unilog−SAP) · só com cód. SAP · + = subir stock SAP' +
        alvoNote +
        '</div></div>' +
        semMatKpi;
    }

    function renderTable(rows) {
      const body = el('Body');
      const empty = el('Empty');
      const emptyMsg = el('EmptyMsg');
      const wrap = el('TableWrap');
      if (!body) return;
      if (!rows.length) {
        body.innerHTML = '';
        if (empty) empty.style.display = 'block';
        if (wrap) wrap.style.display = 'none';
        if (emptyMsg) {
          const hasFilt =
            drill ||
            (typeof msGet === 'function' && (msGet(catMs).size || msGet(subMs).size)) ||
            (el('Srch')?.value || '').trim();
          emptyMsg.textContent = hasFilt
            ? 'Nenhuma linha com os filtros activos — limpa categoria/subcategoria ou o filtro do resumo'
            : excludeFI
              ? 'Carrega SAP e Unilog e processa — ou sem diferenças (já sem FI)'
              : 'Carrega SAP e Unilog e processa — ou sem diferenças nos 5 depósitos';
        }
        return;
      }
      if (empty) empty.style.display = 'none';
      if (wrap) wrap.style.display = '';
      const sorted = typeof sorts !== 'undefined' && sorts[sortKey] ? rows.slice() : rows.slice();
      if (typeof sorts !== 'undefined' && sorts[sortKey] && typeof applySort === 'function') applySort(sortKey, sorted);
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
          const dirBadge = r._semMaterial
            ? '<span class="badge bgray">Sem cód. SAP</span>'
            : r.diff > 0
              ? '<span class="badge bok">Subir</span>'
              : r.diff < 0
                ? '<span class="badge berr">Descer</span>'
                : '—';
          const diffCls = r.diff > 0 ? 'pos' : r.diff < 0 ? 'neg' : '';
          const valorCls = r.valor > 0 ? 'pos' : r.valor < 0 ? 'neg' : '';
          const valorTxt =
            r._semMaterial || r._semValor || !(r.preco > 0)
              ? '—'
              : typeof fmtBrlLiq === 'function'
                ? fmtBrlLiq(r.valor)
                : r.valor;
          const matCell = r._semMaterial
            ? '<span style="color:var(--muted)">—</span> <span class="badge bwarn" title="Existe no WMS mas sem material SAP mapeado">Sem cód. SAP</span>'
            : r.material;
          return (
            '<tr' + (r._semMaterial ? ' style="background:rgba(234,179,8,.08)"' : '') + '>' +
            '<td><span class="badge bgray">' + r.depotSheet + '</span></td>' +
            '<td>' + (r.cat || '—') + '</td>' +
            '<td>' + (r.subcat || '—') + '</td>' +
            '<td style="font-family:var(--mono);font-size:11px">' + matCell + '</td>' +
            '<td style="font-family:var(--mono);font-size:11px">' + (r.EAN_norm || '—') + '</td>' +
            '<td class="col-desc">' + (r.desc || '—') + '</td>' +
            '<td style="font-family:var(--mono);font-size:11px">' + (r.lote || '—') + '</td>' +
            '<td class="num">' + fmtNum(r.qt_sap) + '</td>' +
            '<td class="num">' + fmtNum(r.qt_wms) + '</td>' +
            '<td class="num ' + diffCls + '">' + iaFmtSignedQty(r.diff) + '</td>' +
            '<td class="num ' + valorCls + '">' + valorTxt + '</td>' +
            '<td>' + dirBadge + '</td>' +
            '<td>' + estadoBadge + '</td>' +
            '</tr>'
          );
        })
        .join('');
      const table = el('Table');
      if (typeof updateSortHeaders === 'function') updateSortHeaders(sortKey);
      if (table && typeof enableTableSort === 'function') enableTableSort(table);
      if (typeof scheduleTableSort === 'function') scheduleTableSort();
    }

    function render() {
      const qbNote = el('QbNote');
      const main = el('Main');
      if (typeof company !== 'undefined' && company !== 'DFB') {
        if (qbNote) qbNote.style.display = 'block';
        if (main) main.style.display = 'none';
        return;
      }
      if (qbNote) qbNote.style.display = 'none';
      if (main) main.style.display = '';

      const beforeDrill = getFilteredRows({ skipDrill: true });
      const rows = iaApplyDrill(beforeDrill, drill);

      const banner = el('FilterBanner');
      const lbl = el('FilterLbl');
      if (banner && lbl) {
        if (drill) {
          banner.style.display = 'flex';
          lbl.textContent = IA_DRILL_LABELS[drill] || drill;
        } else {
          banner.style.display = 'none';
        }
      }

      renderKpis(rows);
      renderTable(rows);

      const depCounts = INVENTARIO_ACERTO_DEPOTS.map((dk) => {
        const n = rows.filter((r) => r.depot === dk).length;
        return INVENTARIO_ACERTO_SHEETS[dk] + ': ' + n;
      }).join(' · ');
      const allBuilt = cachedRows || [];
      const semMatTot = allBuilt._countSemMaterial || allBuilt.filter((r) => r._semMaterial).length;
      const exclTot = allBuilt._skippedExcluded || 0;
      const extras = [];
      if (semMatTot) extras.push(semMatTot + ' sem cód. SAP (visíveis, não exportáveis)');
      if (exclTot) extras.push(exclTot + ' excl. equip.prof./cx.papelão');
      const sub = el('DepotCounts');
      if (sub) sub.textContent = depCounts + (extras.length ? ' · ' + extras.join(' · ') : '');
    }

    function syncNav() {
      const it = document.getElementById(cfg.navId);
      if (!it) return;
      const show = typeof company !== 'undefined' && company === 'DFB';
      it.style.display = show ? '' : 'none';
      if (!show && it.classList.contains('active')) {
        const loteTab = document.querySelector('#reconIts .it');
        if (loteTab && typeof reconTab === 'function') reconTab('lote', loteTab);
      }
    }

    function doSortCol(col) {
      if (typeof doSort === 'function') doSort(sortKey, col);
    }

    return {
      render,
      exportXlsx,
      syncNav,
      setDrill,
      clearDrill,
      doSortCol,
      getFilteredRows,
      buildRows,
      excludeFI,
    };
  }

  const iaFull = createInventarioAcertoView({
    prefix: 'ia',
    excludeFI: false,
    sortKey: 'ia',
    navId: 'it-inventario-acerto',
    panelId: 'recon-inventario',
    fileStem: 'Inventario_DFB_acerto_inventario',
    justify: 'acerto inventário',
    renderName: 'renderInventarioAcertoPage',
  });

  const iaSemFi = createInventarioAcertoView({
    prefix: 'ias',
    excludeFI: true,
    sortKey: 'ias',
    navId: 'it-inventario-acerto-sem-fi',
    panelId: 'recon-inventario-sem-fi',
    fileStem: 'Inventario_DFB_acerto_inventario_sem_FI',
    justify: 'acerto inventário sem FI',
    renderName: 'renderInventarioAcertoSemFiPage',
  });

  function syncInventarioAcertoNav() {
    iaFull.syncNav();
    iaSemFi.syncNav();
  }

  global.INVENTARIO_ACERTO_JS_VERSION = INVENTARIO_ACERTO_JS_VERSION;
  global.INVENTARIO_ACERTO_DEPOTS = INVENTARIO_ACERTO_DEPOTS;
  global.buildInventarioAcertoRows = function () { return iaFull.buildRows(); };
  global.getInventarioAcertoFilteredRows = function (opts) { return iaFull.getFilteredRows(opts); };
  global.renderInventarioAcertoPage = function () { iaFull.render(); };
  global.exportInventarioAcertoXlsx = function () { return iaFull.exportXlsx(); };
  global.inventarioAcertoExportFilename = function () {
    const d = new Date();
    return 'Inventario_DFB_acerto_inventario_' + d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') + '.xlsx';
  };
  global.syncInventarioAcertoNav = syncInventarioAcertoNav;
  global.iaDoSort = function (col) { iaFull.doSortCol(col); };
  global.iaSetDrill = function (k) { iaFull.setDrill(k); };
  global.iaClearDrill = function () { iaFull.clearDrill(); };

  global.buildInventarioAcertoSemFiRows = function () { return iaSemFi.buildRows(); };
  global.getInventarioAcertoSemFiFilteredRows = function (opts) { return iaSemFi.getFilteredRows(opts); };
  global.renderInventarioAcertoSemFiPage = function () { iaSemFi.render(); };
  global.exportInventarioAcertoSemFiXlsx = function () { return iaSemFi.exportXlsx(); };
  global.iasDoSort = function (col) { iaSemFi.doSortCol(col); };
  global.iasSetDrill = function (k) { iaSemFi.setDrill(k); };
  global.iasClearDrill = function () { iaSemFi.clearDrill(); };
})(typeof window !== 'undefined' ? window : globalThis);
