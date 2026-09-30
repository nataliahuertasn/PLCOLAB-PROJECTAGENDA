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
    { key: 'invoiceId', label: 'Invoice number', required: true },
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
    useSupplier: true,
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
    if (S.step === 2) { body.innerHTML = stepConfig(); foot.innerHTML = footConfig(); bindConfig(); }
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
      </div>
      <div class="note">${icon('i-info', 'icon sm')}<span>Detected columns: ${S.headers.filter(Boolean).map(h => `<b>${esc(h)}</b>`).join(', ')}</span></div>`;
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
      ${fileHtml}`;
  }
  function footUpload() {
    return `<button class="btn btn-text" data-close="cross">Cancel</button>
      <button class="btn btn-primary btn-flat" id="toStep2" ${S.file && S.rowsRaw.length ? '' : 'disabled'}>Continue</button>`;
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
    if (next) next.onclick = () => { S.step = 2; renderCross(); };
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
        renderCross();
        if (!S.rowsRaw.length) showFileError('The file has no records.');
        else if (S.map.invoiceId === undefined) showFileError('The invoice number column was not detected automatically. You can select it in the next step.');
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
    S.useSupplier = S.map.supplier !== undefined;
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

  function stepConfig() {
    const period = global.App.getPeriodRecords();
    const range = global.App.getRange();
    const ext = buildExternal();
    const noId = S.map.invoiceId === undefined ? ext.length : ext.filter(r => !CrossMatch.normalizeInvoice(r.invoiceId)).length;
    const colOpts = sel => `<option value="">— Not used —</option>` + S.headers.map((h, i) => h ? `<option value="${i}" ${sel === i ? 'selected' : ''}>${esc(h)}</option>` : '').join('');
    const sample = k => {
      if (S.map[k] === undefined) return '';
      const first = S.rowsText.find(r => String(r[S.map[k]] ?? '').trim());
      return first ? 'e.g. ' + esc(first[S.map[k]]) : 'Empty column';
    };
    return `
      <h3 class="section-title">File analysis</h3>
      <div class="scope-box">
        <div class="scope">
          <div class="lbl">${icon('i-db', 'icon sm')} Project Agenda</div>
          <div class="val">${fmt(period.length)}</div>
          <div class="desc">records · period ${range.start} ~ ${range.end}<br>(issue date filter of the view)</div>
        </div>
        <div class="scope-arrow">${icon('i-swap', 'icon lg')}</div>
        <div class="scope">
          <div class="lbl">${icon('i-sheet', 'icon sm')} PL Colab</div>
          <div class="val">${fmt(ext.length)}</div>
          <div class="desc">records · ${esc(S.file.name)}<br>(analyzed in full, not filtered by date)</div>
        </div>
      </div>

      <h3 class="section-title">PL Colab report columns</h3>
      <p class="section-help">Detected automatically. Adjust them if the report uses different column names.</p>
      <div class="map-grid">
        ${MAP_FIELDS.map(f => `<div class="field">
            <label for="map-${f.key}">${f.label}${f.required ? ' <span class="req">*</span>' : ''}</label>
            <select id="map-${f.key}" data-map="${f.key}">${colOpts(S.map[f.key])}</select>
            <span class="sample">${sample(f.key)}</span>
          </div>`).join('')}
      </div>

      <div class="field" style="margin-top:14px">
        <label style="display:flex;gap:8px;align-items:center;color:var(--text);font-size:12.5px">
          <input type="checkbox" id="useSupplier" ${S.useSupplier ? 'checked' : ''} ${S.map.supplier === undefined ? 'disabled' : ''}>
          Identify each invoice by <b>Supplier + Invoice number</b> (recommended)
        </label>
      </div>
      <div class="note">${icon('i-info', 'icon sm')}<span>
        ${S.useSupplier && S.map.supplier !== undefined
          ? 'Matching key: <b>Supplier + Invoice number</b>. The invoice number alone is not unique: the same number appears for different suppliers. Case, spaces, periods and company suffixes (S.A.S., S.A., LTDA…) are ignored.'
          : 'Matching key: <b>Invoice number</b> (normalized: no spaces, dashes or case differences).'}
        </span></div>
      ${S.map.invoiceId === undefined
        ? `<div class="note err">${icon('i-error', 'icon sm')}<span>Select the invoice number column to reconcile.</span></div>`
        : noId ? `<div class="note err">${icon('i-warn', 'icon sm')}<span>${fmt(noId)} PL Colab record(s) have no invoice number and cannot be reconciled.</span></div>` : ''}
    `;
  }
  function footConfig() {
    return `<button class="btn btn-text" id="back1">Back</button><span class="spacer"></span>
      <button class="btn btn-text" data-close="cross">Cancel</button>
      <button class="btn btn-primary btn-flat" id="startCross" ${S.map.invoiceId === undefined ? 'disabled' : ''}>${icon('i-compare', 'icon sm')}Start reconciliation</button>`;
  }
  function bindConfig() {
    document.querySelectorAll('[data-map]').forEach(sel => sel.onchange = () => {
      const k = sel.dataset.map;
      if (sel.value === '') delete S.map[k]; else S.map[k] = +sel.value;
      if (k === 'supplier') S.useSupplier = S.map.supplier !== undefined;
      renderCross();
    });
    $('#useSupplier').onchange = e => { S.useSupplier = e.target.checked; renderCross(); };
    $('#back1').onclick = () => { S.step = 1; renderCross(); };
    $('#startCross').onclick = runCross;
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
      () => { software = global.App.getPeriodRecords(); },
      () => {},
      () => {
        const compareFields = ['amount', 'issueDate', 'oc'].filter(f => S.map[f] !== undefined);
        result = CrossMatch.run(software, external, {
          useSupplier: S.useSupplier && S.map.supplier !== undefined,
          compareFields, softwareAll: global.App.getAllRecords()
        });
      },
      () => {}, () => {}
    ];
    try {
      for (let i = 0; i < tasks.length; i++) { mark(i, 'run'); await wait(280); tasks[i](); mark(i, 'done'); }
      await wait(250);
      S.external = external; S.result = result;
      S.crossedRange = global.App.getRange(); S.crossedAt = new Date();
      S.busy = false; S.step = 2;
      showResult();
    } catch (err) {
      S.busy = false; S.step = 2; renderCross();
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
      <dt>PL Colab Invoices</dt><dd>${fmt(r.external.rows)}</dd>
      <dt>Project Agenda Invoices</dt><dd>${fmt(r.software.rows)}</dd>`;
    $('#recTbody').innerHTML = rows.length
      ? rows.map(x => `<tr><td>${esc(x.consecutive)}</td><td>${esc(x.supplier)}</td><td>${esc(x.issueDate)}</td></tr>`).join('')
      : `<tr><td colspan="3" class="rec-empty">All PL Colab invoices were found in Project Agenda.</td></tr>`;
    $('#recNote').textContent = `${fmt(rows.length)} PL Colab invoice(s) not found in Project Agenda · ${S.file.name} · Period ${S.crossedRange.start} ~ ${S.crossedRange.end}`;
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
    S.step = S.file ? 2 : 1;
    open('cross');
    renderCross();
  }
  $('#btnCross').addEventListener('click', () => startFlow(false));

  global.CrossUI = {
    onRangeChange() {},
    showResult() { if (S.result) showResult(); }
  };
})(window);
