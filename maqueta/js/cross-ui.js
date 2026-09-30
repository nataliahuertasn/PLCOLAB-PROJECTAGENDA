/* =========================================================
   Cruce de información — modales (carga, análisis, proceso,
   resultado y detalle) e integración con la vista actual.
   ========================================================= */
(function (global) {
  'use strict';

  const $ = sel => document.querySelector(sel);
  const { esc, icon, toast } = global.App;
  const fmt = n => Number(n).toLocaleString('es-CO');
  const wait = ms => new Promise(r => setTimeout(r, ms));

  const MAP_FIELDS = [
    { key: 'invoiceId', label: 'Invoice number (Nro. Documento)', required: true },
    { key: 'supplier',  label: 'Supplier' },
    { key: 'amount',    label: 'Invoice amount' },
    { key: 'issueDate', label: 'Issue date' },
    { key: 'oc',        label: 'Purchase order (OC)' },
    { key: 'dueDate',   label: 'Due date' },
    { key: 'type',      label: 'Type' }
  ];

  const S = {
    step: 1,
    file: null,        // { name, size }
    wb: null,
    sheet: '',
    headerIndex: 0,
    headers: [],
    rowsRaw: [],
    rowsText: [],
    map: {},
    external: [],
    result: null,
    crossedRange: null,
    crossedAt: null,
    busy: false
  };

  // ---------- Modales ----------
  const backdrops = { cross: $('#crossBackdrop'), result: $('#resultBackdrop') };
  function open(name) { Object.entries(backdrops).forEach(([k, b]) => { b.hidden = k !== name; }); }
  function close(name) { backdrops[name].hidden = true; }
  document.addEventListener('click', e => {
    const c = e.target.closest('[data-close]');
    if (c) { if (c.dataset.close === 'cross' && S.busy) return; close(c.dataset.close); }
  });
  Object.entries(backdrops).forEach(([k, b]) => b.addEventListener('mousedown', e => {
    if (e.target === b && !(k === 'cross' && S.busy)) close(k);
  }));
  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    const openOne = Object.entries(backdrops).find(([, b]) => !b.hidden);
    if (openOne && !(openOne[0] === 'cross' && S.busy)) close(openOne[0]);
  });

  // ---------- Pasos ----------
  function renderCross() {
    const body = $('#crossBody'), foot = $('#crossFooter');
    if (S.step === 1) { body.innerHTML = stepUpload(); foot.innerHTML = footUpload(); bindUpload(); }
    if (S.step === 3) { body.innerHTML = stepProcessing(); foot.innerHTML = ''; }
  }

  // ----- Paso 1: carga -----
  function stepUpload() {
    let fileHtml = '';
    if (S.file) {
      const sheets = S.wb ? S.wb.SheetNames : [];
      fileHtml = `<div class="file-card">
        <div class="file-icon">${icon('i-sheet', 'icon lg')}</div>
        <div style="min-width:0">
          <div class="name">${esc(S.file.name)}</div>
          <div class="meta">${(S.file.size / 1024).toFixed(0)} KB · <b>${fmt(S.rowsRaw.length)}</b> records found · ${S.headers.filter(Boolean).length} columns</div>
        </div>
        <div class="spacer"></div>
        ${sheets.length > 1 ? `<div class="field" style="min-width:160px"><label>Sheet</label><select id="sheetSel">${sheets.map(s => `<option ${s === S.sheet ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></div>` : ''}
        <button class="icon-btn" id="removeFile" title="Remove file" aria-label="Remove file">${icon('i-close')}</button>
      </div>`;
    }
    return `<h3 class="section-title">Upload report to reconcile</h3>
      <p class="section-help">Select the PL Colab report. Each record will be compared against the Project Agenda invoices.</p>
      <div class="dropzone" id="dropzone">
        ${icon('i-upload')}
        <p>Drag the file here or</p>
        <button class="btn btn-primary" id="pickFile" type="button">Select file</button>
        <input type="file" id="fileInput" accept=".xlsx" hidden>
        <div class="formats">Supported format: .xlsx</div>
      </div>
      <div id="fileError"></div>
      ${fileHtml}
      <div class="period-block">
        <h3 class="section-title">Select period to reconcile</h3>
        <p class="section-help">Project Agenda invoices with an <b>Issue Date</b> in this period will be compared against the PL Colab report.</p>
        <div class="date-filter">
          ${icon('i-calendar')}
          <button class="outlined" id="recPeriod" type="button" aria-haspopup="dialog">
            <span class="legend">Issue Date (Max 3 months)</span>
            <span>${S.range ? `${S.range.start} ~ ${S.range.end}` : 'Select period'}</span>
          </button>
          ${S.range ? `<span class="period-count"><b>${fmt(global.App.countInRange(S.range.start, S.range.end))}</b> Project Agenda invoices${S.rangeSuggested ? ' · suggested from the PL Colab issue dates' : ''}</span>` : ''}
        </div>
      </div>`;
  }
  function footUpload() {
    return `<button class="btn btn-text" data-close="cross">Cancel</button>
      <button class="btn btn-primary btn-flat" id="toStep2" ${S.file && S.rowsRaw.length && S.map.invoiceId !== undefined && S.range ? '' : 'disabled'}>Continue</button>`;
  }

  /** Propone el período a partir de las fechas de emisión del archivo de PL Colab (si caben en 3 meses). */
  function suggestRange() {
    if (S.map.issueDate === undefined) return;
    const dates = buildExternal().map(r => r.issueDate).filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d || '')).sort();
    if (!dates.length) return;
    const start = dates[0], end = dates[dates.length - 1];
    const [y, m, d] = start.split('-').map(Number);
    const max = new Date(y, m - 1 + 3, d);
    const maxIso = max.getFullYear() + '-' + String(max.getMonth() + 1).padStart(2, '0') + '-' + String(max.getDate()).padStart(2, '0');
    if (end <= maxIso) { S.range = { start, end }; S.rangeSuggested = true; }
  }
  function bindUpload() {
    const input = $('#fileInput'), dz = $('#dropzone');
    $('#pickFile').onclick = () => input.click();
    input.onchange = () => input.files[0] && loadFile(input.files[0]);
    ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('drag'); }));
    ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('drag'); }));
    dz.addEventListener('drop', e => e.dataTransfer.files[0] && loadFile(e.dataTransfer.files[0]));
    const rm = $('#removeFile');
    if (rm) rm.onclick = () => { Object.assign(S, { file: null, wb: null, headers: [], rowsRaw: [], rowsText: [], map: {} }); renderCross(); };
    const ss = $('#sheetSel');
    if (ss) ss.onchange = () => { readSheet(ss.value); renderCross(); };
    const next = $('#toStep2');
    if (next) next.onclick = runCross;
    $('#recPeriod').onclick = () => global.RangePicker.open({
      anchor: $('#recPeriod'), start: S.range && S.range.start, end: S.range && S.range.end,
      maxMonths: 3, count: global.App.countInRange,
      onApply(range) { S.range = range; S.rangeSuggested = false; renderCross(); }
    });
  }

  function showFileError(msg) {
    const el = $('#fileError');
    if (el) el.innerHTML = `<div class="note err">${icon('i-error', 'icon sm')}<span>${esc(msg)}</span></div>`;
  }

  function loadFile(file) {
    if (!global.XLSX) { showFileError('The Excel reader library could not be loaded (offline).'); return; }
    if (!/\.xlsx$/i.test(file.name)) { showFileError('Unsupported format. Please use an .xlsx file.'); return; }
    const dz = $('#dropzone');
    dz.innerHTML = `<div class="spinner" style="margin:0 auto 8px;width:28px;height:28px"></div><p>Reading ${esc(file.name)}…</p>`;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        S.wb = XLSX.read(new Uint8Array(reader.result), { type: 'array', cellDates: true });
        S.file = { name: file.name, size: file.size };
        readSheet(S.wb.SheetNames[0]);
        suggestRange();
        renderCross();
        if (!S.rowsRaw.length) showFileError('The file has no records.');
        else if (S.map.invoiceId === undefined) showFileError('The file does not have a column named "' + CrossMatch.INVOICE_HEADER + '". It is required to reconcile.');
      } catch (err) {
        S.file = null; renderCross(); showFileError('The file could not be read: ' + err.message);
      }
    };
    reader.onerror = () => { renderCross(); showFileError('Error reading the file.'); };
    reader.readAsArrayBuffer(file);
  }

  function readSheet(name) {
    S.sheet = name;
    const ws = S.wb.Sheets[name];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: false });
    const text = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: false });
    S.headerIndex = CrossMatch.detectHeaderRow(text);
    S.headers = (text[S.headerIndex] || []).map(h => String(h).trim());
    const isEmpty = r => !r || r.every(c => String(c).trim() === '');
    S.rowsRaw = []; S.rowsText = []; S.rowNumbers = [];
    for (let i = S.headerIndex + 1; i < raw.length; i++) {
      if (isEmpty(text[i])) continue;
      S.rowsRaw.push(raw[i]); S.rowsText.push(text[i]); S.rowNumbers.push(i + 1);
    }
    // Número de fila real en Excel (considera filas vacías al inicio)
    const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r : 0;
    S.rowNumbers = S.rowNumbers.map(n => n + ref);
    S.map = CrossMatch.autoMap(S.headers);
  }

  // ----- Paso 2: análisis y configuración -----
  function buildExternal() {
    const m = S.map;
    const get = (arr, k) => m[k] === undefined ? undefined : arr[m[k]];
    return S.rowsRaw.map((raw, i) => {
      const text = S.rowsText[i];
      const str = v => v === undefined ? undefined : String(v).trim();
      const d = v => v === undefined ? undefined : CrossMatch.normalizeDate(v);
      return {
        _src: 'externo', _row: S.rowNumbers[i],
        invoiceId: str(get(text, 'invoiceId')) ?? '',
        supplier: str(get(text, 'supplier')) ?? '',
        amount: str(get(text, 'amount')),
        issueDate: d(get(raw, 'issueDate')),
        dueDate: d(get(raw, 'dueDate')),
        oc: str(get(text, 'oc')),
        type: str(get(text, 'type'))
      };
    });
  }

  // ----- Paso 3: proceso -----
  const PROC = ['Reading PL Colab records', 'Loading Project Agenda records for the period', 'Normalizing identifiers', 'Comparing record by record', 'Detecting duplicates and differences', 'Building result'];
  function stepProcessing() {
    return `<div class="processing">
      <div style="font-size:14px;font-weight:500">Reconciling…</div>
      <div class="progress"><div id="procBar"></div></div>
      <ul class="proc-steps" id="procSteps">${PROC.map(p => `<li><span class="pending"></span>${p}</li>`).join('')}</ul>
    </div>`;
  }
  async function runCross() {
    S.busy = true; S.step = 3; renderCross();
    const lis = [...document.querySelectorAll('#procSteps li')];
    const mark = (i, st) => {
      const li = lis[i]; li.className = st;
      li.firstElementChild.outerHTML = st === 'run' ? '<span class="spinner"></span>' : icon('i-check', 'icon sm');
      $('#procBar').style.width = ((st === 'done' ? i + 1 : i + 0.5) / PROC.length * 100) + '%';
    };
    let external, software, result;
    const tasks = [
      () => { external = buildExternal(); },
      () => { software = global.App.getRecordsInRange(S.range.start, S.range.end); },
      () => {},
      () => {
        const compareFields = ['amount', 'issueDate', 'oc'].filter(f => S.map[f] !== undefined);
        result = CrossMatch.run(software, external, {
          useSupplier: false,   // regla fija: solo Invoice ID vs Nro. Documento
          compareFields, softwareAll: global.App.getAllRecords()
        });
      },
      () => {}, () => {}
    ];
    try {
      for (let i = 0; i < tasks.length; i++) { mark(i, 'run'); await wait(90); tasks[i](); mark(i, 'done'); }
      await wait(150);
      S.external = external; S.result = result;
      S.crossedRange = { ...S.range }; S.crossedAt = new Date();
      S.busy = false; S.step = 1;
      showResult();
    } catch (err) {
      S.busy = false; S.step = 1; renderCross();
      $('#crossBody').insertAdjacentHTML('afterbegin', `<div class="note err" style="margin:0 0 12px">${icon('i-error', 'icon sm')}<span>Reconciliation error: ${esc(err.message)}</span></div>`);
    }
  }

  // ---------- Resultado: PL Colab & Project Agenda Invoice Reconciliation ----------
  // PL Colab = reporte externo cargado · Project Agenda = reporte del software (período de la vista).
  // La tabla lista las facturas de PL Colab que no se encontraron en Project Agenda.
  const TITLE = 'PL Colab & Project Agenda Invoice Reconciliation';

  function reconciliationRows() {
    const rows = [];
    S.result.onlyExternal.forEach(g => g.records.forEach(rec => rows.push({
      consecutive: rec.invoiceId, supplier: rec.supplier, issueDate: rec.issueDate || ''
    })));
    return rows;
  }

  function showResult() {
    const r = S.result;
    const rows = reconciliationRows();
    $('#recTitle').textContent = TITLE;
    $('#recCounts').innerHTML = `
      <div class="rec-stat">${icon('i-sheet')}<div><span>PL Colab Invoices</span><b>${fmt(r.external.rows)}</b></div></div>
      <div class="rec-stat">${icon('i-db')}<div><span>Project Agenda Invoices</span><b>${fmt(r.software.rows)}</b></div></div>`;
    $('#recListTitle').innerHTML = `Not found in Project Agenda <span class="rec-badge">${fmt(rows.length)}</span>`;
    $('#recTbody').innerHTML = rows.length
      ? rows.map(x => `<tr><td class="rec-id">${esc(x.consecutive)}</td><td>${x.supplier ? esc(x.supplier) : '<span class="rec-na">—</span>'}</td><td class="rec-date">${esc(x.issueDate)}</td></tr>`).join('')
      : `<tr><td colspan="3" class="rec-empty">${icon('i-check')}All PL Colab invoices were found in Project Agenda.</td></tr>`;
    open('result');
  }

  $('#recExport').onclick = () => {
    if (!global.XLSX) { toast('The Excel library could not be loaded.'); return; }
    const r = S.result;
    const rows = reconciliationRows();
    const aoa = [
      [TITLE], [],
      ['PL Colab Invoices', r.external.rows],
      ['Project Agenda Invoices', r.software.rows], [],
      ['Consecutive', 'Supplier', 'Issue Date'],
      ...rows.map(x => [x.consecutive, x.supplier, x.issueDate])
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 26 }, { wch: 48 }, { wch: 14 }];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 5, c: 0 }, e: { r: 5 + rows.length, c: 2 } }) };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Reconciliation');
    XLSX.writeFile(wb, `reconciliation_plcolab_vs_projectagenda_${S.crossedRange.start}_to_${S.crossedRange.end}.xlsx`);
    toast(`Excel exported with ${fmt(rows.length)} invoice(s).`);
  };

  // ---------- Entrada ----------
  function startFlow(reset) {
    if (reset) Object.assign(S, { file: null, wb: null, headers: [], rowsRaw: [], rowsText: [], map: {} });
    S.step = 1;
    if (!S.range) { S.range = global.App.getRange(); S.rangeSuggested = false; }
    open('cross');
    renderCross();
  }
  $('#btnCross').addEventListener('click', () => startFlow(false));

  global.CrossUI = {
    onRangeChange() {},
    showResult() { if (S.result) showResult(); }
  };
})(window);
