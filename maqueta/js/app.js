/* =========================================================
   Vista actual: Invoices Archive (tabla, filtros, rango de fechas, reporte)
   ========================================================= */
(function (global) {
  'use strict';

  const $ = sel => document.querySelector(sel);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const icon = (id, cls = 'icon') => `<svg class="${cls}"><use href="#${id}"/></svg>`;

  // ---------- Datos (del Excel del software, sin modificar) ----------
  // SAP / ACCOUNTANT: valores aleatorios solo para presentación (fijos por registro).
  // Regla: si ACCOUNTANT está marcado, SAP siempre está marcado.
  function demoChecks(i) {
    let t = (i + 1) * 0x6D2B79F5;
    const rnd = () => { t = (t + 0x6D2B79F5) | 0; let x = Math.imul(t ^ (t >>> 15), 1 | t); x ^= x + Math.imul(x ^ (x >>> 7), 61 | x); return ((x ^ (x >>> 14)) >>> 0) / 4294967296; };
    const sap = rnd() < 0.55;
    const acc = sap && rnd() < 0.7;
    return { sap: sap ? 'checked' : 'unchecked', acc: acc ? 'checked' : 'unchecked' };
  }
  const records = (global.SOFTWARE_RECORDS || []).map((r, i) => ({
    _id: i, _row: i + 2,
    type: r[0], supplier: r[1], invoiceId: r[2], issueDate: r[3], dueDate: r[4],
    amount: r[5], oc: r[6], ...demoChecks(i), pdf: r[9]
  }));

  // ---------- Estado ----------
  const state = {
    range: { start: '', end: '' },
    sort: { key: 'issueDate', dir: 'desc' },
    filters: {},            // key → texto / valor select
    cross: null             // { status: Map(rec→estado), meta }
  };

  const COLUMNS = [
    { key: 'type',      label: 'Type',           filter: 'text' },
    { key: 'supplier',  label: 'Supplier',       filter: 'text' },
    { key: 'invoiceId', label: 'Invoice ID',     filter: 'text' },
    { key: 'issueDate', label: 'Issue Date',     filter: 'text' },
    { key: 'dueDate',   label: 'Due Date',       filter: 'text' },
    { key: 'amount',    label: 'Invoice Amount', filter: 'text', numeric: true },
    { key: 'oc',        label: 'OC',             filter: 'text' },
    { key: 'sap',       label: 'SAP',            filter: 'check', center: true },
    { key: 'acc',       label: 'Accountant',     filter: 'check', center: true },
    { key: 'cross',     label: 'Cruce',          filter: 'cross', crossOnly: true },
    { key: 'actions',   label: 'Actions',        center: true, noSort: true }
  ];

  const CROSS_STATUS = {
    ok:    { cls: 'ok',    icon: 'i-check',    label: 'Coincide' },
    diff:  { cls: 'diff',  icon: 'i-diff',     label: 'Con diferencias' },
    alert: { cls: 'alert', icon: 'i-notfound', label: 'No está en externo' },
    dup:   { cls: 'warn',  icon: 'i-dup',      label: 'Duplicado' },
    error: { cls: 'error', icon: 'i-error',    label: 'Sin Nº factura' },
    none:  { cls: 'neutral', icon: 'i-info',   label: 'Fuera del cruce' }
  };
  function statusChip(key, extraTitle) {
    const s = CROSS_STATUS[key];
    return `<span class="status ${s.cls}" title="${esc(extraTitle || s.label)}">${icon(s.icon)}${s.label}</span>`;
  }

  // ---------- Fechas ----------
  const pad = n => String(n).padStart(2, '0');
  const iso = d => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  const parseIso = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
  function addMonths(s, n) { const d = parseIso(s); d.setMonth(d.getMonth() + n); return iso(d); }

  function defaultRange() {
    const end = iso(new Date());
    const range = { start: addMonths(end, -3), end };
    const inRange = records.some(r => r.issueDate >= range.start && r.issueDate <= range.end);
    if (!inRange && records.length) {           // si hoy está lejos de los datos, anclar al último registro
      const max = records.reduce((m, r) => r.issueDate > m ? r.issueDate : m, '');
      return { start: addMonths(max, -3), end: max };
    }
    return range;
  }

  // ---------- Consultas ----------
  /** Registros del reporte del software para el período (lo que genera REPORT). */
  function getPeriodRecords() {
    const { start, end } = state.range;
    return records.filter(r => r.issueDate >= start && r.issueDate <= end);
  }

  function crossKeyOf(rec) {
    if (!state.cross) return '';
    const st = state.cross.status.get(rec);
    return st === undefined ? 'none' : st;
  }

  function getVisibleRecords() {
    let rows = getPeriodRecords();
    const f = state.filters;
    COLUMNS.forEach(c => {
      const v = f[c.key];
      if (!v) return;
      if (c.filter === 'text') {
        const q = v.toLowerCase();
        rows = rows.filter(r => String(r[c.key] ?? '').toLowerCase().includes(q));
      } else if (c.filter === 'check') {
        rows = rows.filter(r => r[c.key] === v);
      } else if (c.filter === 'cross' && state.cross) {
        rows = rows.filter(r => {
          const k = crossKeyOf(r);
          if (v === 'dup') return k.includes('+dup');
          return k.split('+')[0] === v;
        });
      }
    });
    const { key, dir } = state.sort;
    const col = COLUMNS.find(c => c.key === key);
    const mul = dir === 'asc' ? 1 : -1;
    const val = r => col && col.numeric ? (CrossMatch.parseAmount(r[key]) ?? -Infinity) : String(r[key] ?? '').toLowerCase();
    rows = rows.slice().sort((a, b) => {
      const x = val(a), y = val(b);
      if (x < y) return -1 * mul;
      if (x > y) return 1 * mul;
      return a._id - b._id;
    });
    return rows;
  }

  // ---------- Render ----------
  function renderHead() {
    const cols = COLUMNS.filter(c => !c.crossOnly || state.cross);
    $('#colCross').style.display = state.cross ? '' : 'none';
    $('#headRow').innerHTML = cols.map(c => {
      const sorted = state.sort.key === c.key;
      const arrow = sorted ? icon(state.sort.dir === 'asc' ? 'i-arrow-up' : 'i-arrow-down', 'icon sort') : icon('i-arrow-down', 'icon sort');
      const label = c.noSort
        ? `<span class="th-label" style="cursor:default">${c.label}</span>`
        : `<button type="button" class="th-label ${sorted ? 'sorted' : ''}" data-sort="${c.key}" title="Ordenar por ${c.label}">${c.label}${arrow}</button>`;
      let filter = '';
      if (c.filter === 'text') {
        filter = `<input class="th-filter" data-filter="${c.key}" value="${esc(state.filters[c.key] || '')}" placeholder="Filtrar…" aria-label="Filtrar ${c.label}">`;
      } else if (c.filter === 'check') {
        const v = state.filters[c.key] || '';
        return `<th class="center"><div class="th-inner" style="align-items:flex-start"><span class="th-select-wrap"><span class="mini">${c.label.toUpperCase()}</span>
          <select class="th-select" data-filter="${c.key}" aria-label="Filtrar ${c.label}">
            <option value="" ${v === '' ? 'selected' : ''}>None</option>
            <option value="checked" ${v === 'checked' ? 'selected' : ''}>Checked</option>
            <option value="unchecked" ${v === 'unchecked' ? 'selected' : ''}>Unchecked</option>
          </select></span></div></th>`;
      } else if (c.filter === 'cross') {
        const v = state.filters.cross || '';
        const opts = [['', 'Todos'], ['ok', 'Coincide'], ['diff', 'Con diferencias'], ['alert', 'No está en externo'], ['dup', 'Duplicado'], ['error', 'Sin Nº factura'], ['none', 'Fuera del cruce']];
        return `<th><div class="th-inner"><span class="th-select-wrap"><span class="mini">CRUCE</span>
          <select class="th-select" data-filter="cross" aria-label="Filtrar por estado del cruce">
            ${opts.map(([k, l]) => `<option value="${k}" ${v === k ? 'selected' : ''}>${l}</option>`).join('')}
          </select></span></div></th>`;
      }
      return `<th class="${c.center ? 'center' : ''}"><div class="th-inner">${label}${filter}</div></th>`;
    }).join('');
  }

  function crossCell(rec) {
    const k = crossKeyOf(rec);
    const parts = k.split('+');
    let html = statusChip(parts[0] || 'none');
    if (parts[1] === 'dup') html += ' ' + statusChip('dup', 'Este Nº de factura aparece más de una vez en el reporte del software');
    return `<td>${html}</td>`;
  }

  function renderBody() {
    const rows = getVisibleRecords();
    const total = getPeriodRecords().length;
    const html = rows.map(r => `<tr data-id="${r._id}">
        <td>${esc(r.type)}</td>
        <td>${esc(r.supplier)}</td>
        <td>${r.invoiceId ? `<a class="inv-link" data-pdf="${r._id}" title="${esc(r.pdf ? 'Download ' + r.pdf : '')}">${esc(r.invoiceId)}</a>` : '<span class="empty-id">(sin Nº)</span>'}</td>
        <td class="nowrap">${esc(r.issueDate)}</td>
        <td class="nowrap">${esc(r.dueDate)}</td>
        <td class="num">${esc(r.amount)}</td>
        <td>${esc(r.oc)}</td>
        <td class="center"><input type="checkbox" class="cb" tabindex="-1" onclick="return false" data-check="sap" ${r.sap === 'checked' ? 'checked' : ''} aria-label="SAP"></td>
        <td class="center"><input type="checkbox" class="cb" tabindex="-1" onclick="return false" data-check="acc" ${r.acc === 'checked' ? 'checked' : ''} aria-label="Accountant"></td>
        ${state.cross ? crossCell(r) : ''}
        <td class="center"><button class="icon-btn" data-menu="${r._id}" aria-label="Acciones">${icon('i-more')}</button></td>
      </tr>`).join('');
    $('#tbody').innerHTML = html;
    $('#emptyState').hidden = rows.length > 0;
    const filtered = rows.length !== total;
    $('#rowCount').innerHTML = filtered
      ? `Mostrando <b>${rows.length.toLocaleString('es-CO')}</b> de ${total.toLocaleString('es-CO')} facturas del período`
      : `<b>${total.toLocaleString('es-CO')}</b> facturas en el período`;
    $('#clearFilters').hidden = !Object.values(state.filters).some(Boolean);
  }

  function render() { renderHead(); renderBody(); }

  function renderDateLabel() { $('#dateLabel').textContent = `${state.range.start} ~ ${state.range.end}`; }

  // ---------- Eventos: tabla ----------
  $('#headRow').addEventListener('click', e => {
    const b = e.target.closest('[data-sort]');
    if (!b) return;
    const key = b.dataset.sort;
    state.sort = state.sort.key === key ? { key, dir: state.sort.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'asc' };
    renderHead(); renderBody();
  });
  let filterTimer;
  $('#headRow').addEventListener('input', e => {
    const el = e.target.closest('input[data-filter]');
    if (!el) return;
    state.filters[el.dataset.filter] = el.value.trim();
    clearTimeout(filterTimer);
    filterTimer = setTimeout(renderBody, 150);
  });
  $('#headRow').addEventListener('change', e => {
    const el = e.target.closest('select[data-filter]');
    if (!el) return;
    state.filters[el.dataset.filter] = el.value;
    renderBody();
  });
  $('#clearFilters').addEventListener('click', () => { state.filters = {}; render(); });


  $('#tbody').addEventListener('click', e => {
    const link = e.target.closest('[data-pdf]');
    if (link) {
      const rec = records[+link.dataset.pdf];
      toast(rec.pdf ? `Descarga de ${rec.pdf} deshabilitada en la maqueta.` : 'Esta factura no tiene PDF asociado.');
      return;
    }
    const mb = e.target.closest('[data-menu]');
    if (mb) openRowMenu(records[+mb.dataset.menu], mb);
  });

  // ---------- Menú de fila ----------
  const menu = $('#rowMenu');
  function openRowMenu(rec, anchor) {
    const items = [
      `<button data-act="pdf">${icon('i-pdf')}Descargar PDF</button>`,
      `<button data-act="copy">${icon('i-dup')}Copiar Nº de factura</button>`
    ];
    if (state.cross) items.push(`<button data-act="cross">${icon('i-compare')}Ver resultado del cruce</button>`);
    menu.innerHTML = items.join('');
    menu.hidden = false;
    const r = anchor.getBoundingClientRect();
    const top = Math.min(r.bottom + 4, window.innerHeight - menu.offsetHeight - 8);
    menu.style.top = top + 'px';
    menu.style.left = Math.max(8, r.right - menu.offsetWidth) + 'px';
    menu.onclick = ev => {
      const act = ev.target.closest('[data-act]')?.dataset.act;
      menu.hidden = true;
      if (act === 'pdf') toast(rec.pdf ? `Descarga de ${rec.pdf} deshabilitada en la maqueta.` : 'Esta factura no tiene PDF asociado.');
      if (act === 'copy') { navigator.clipboard?.writeText(rec.invoiceId || '').catch(() => {}); toast('Nº de factura copiado.'); }
      if (act === 'cross') global.CrossUI && global.CrossUI.showResult();
    };
  }
  document.addEventListener('click', e => {
    if (!menu.hidden && !e.target.closest('#rowMenu') && !e.target.closest('[data-menu]')) menu.hidden = true;
  });
  $('#tableScroll').addEventListener('scroll', () => { menu.hidden = true; });

  // ---------- Rango de fechas (calendario) ----------
  const countInRange = (s, e) => records.filter(r => r.issueDate >= s && r.issueDate <= e).length;
  $('#dateField').addEventListener('click', () => {
    global.RangePicker.open({
      anchor: $('#dateField'), start: state.range.start, end: state.range.end, maxMonths: 3, count: countInRange,
      onApply(range) {
        state.range = range;
        renderDateLabel(); renderBody();
        global.CrossUI && global.CrossUI.onRangeChange(state.range);
      }
    });
  });

  // ---------- REPORT (funcionalidad existente: exporta el reporte del período) ----------
  $('#btnReport').addEventListener('click', () => {
    if (!global.XLSX) { toast('No se pudo cargar la librería de Excel.'); return; }
    const rows = getVisibleRecords();
    const aoa = [['Type', 'Supplier', 'Invoice ID', 'Issue Date', 'Due Date', 'Invoice Amount', 'OC', 'SAP', 'ACCOUNTANT']]
      .concat(rows.map(r => [r.type, r.supplier, r.invoiceId, r.issueDate, r.dueDate, r.amount, r.oc, r.sap, r.acc]));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Invoices History');
    XLSX.writeFile(wb, `invoices_history_${state.range.start}_to_${state.range.end}_.xlsx`);
    toast(`Reporte generado con ${rows.length.toLocaleString('es-CO')} facturas.`);
  });

  // ---------- Toast ----------
  let toastTimer;
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg; t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 3200);
  }

  // ---------- API para el módulo de cruce ----------
  global.App = {
    esc, icon, toast, statusChip,
    getAllRecords: () => records,
    getPeriodRecords,
    getRecordsInRange: (s, e) => records.filter(r => r.issueDate >= s && r.issueDate <= e),
    countInRange,
    getRange: () => ({ ...state.range }),
    setCross(status, meta) {
      state.cross = { status, meta };
      state.filters.cross = '';
      render();
    },
    clearCross() { state.cross = null; delete state.filters.cross; render(); },
    filterByCross(value) {
      if (!state.cross) return;
      state.filters.cross = value; render();
      $('#tableScroll').scrollTop = 0;
    },
    focusRecord(rec) {
      if (rec.issueDate < state.range.start || rec.issueDate > state.range.end) {
        toast('La factura está fuera del período seleccionado.'); return;
      }
      state.filters = {}; render();
      const tr = document.querySelector(`#tbody tr[data-id="${rec._id}"]`);
      if (tr) { tr.scrollIntoView({ block: 'center' }); tr.classList.add('flash'); setTimeout(() => tr.classList.remove('flash'), 1700); }
    }
  };

  // ---------- Inicio ----------
  state.range = defaultRange();
  renderDateLabel();
  $('#sourceNote').textContent = `Fuente: ${global.SOFTWARE_SOURCE || 'reporte del software'} · ${records.length.toLocaleString('es-CO')} registros`;
  render();
})(window);
