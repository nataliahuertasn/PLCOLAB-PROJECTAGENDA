/* =========================================================
   RangePicker — calendario de rango de fechas (estilo Material)
   Uso: RangePicker.open({ anchor, start, end, maxMonths, count, onApply })
   ========================================================= */
(function (global) {
  'use strict';

  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const MON = MONTHS.map(m => m.slice(0, 3));
  const pad = n => String(n).padStart(2, '0');
  const iso = (y, m, d) => y + '-' + pad(m + 1) + '-' + pad(d);
  const parse = s => { const [y, m, d] = s.split('-').map(Number); return { y, m: m - 1, d }; };
  function addMonths(s, n) { const p = parse(s); const dt = new Date(p.y, p.m + n, p.d); return iso(dt.getFullYear(), dt.getMonth(), dt.getDate()); }
  const short = s => { const p = parse(s); return MON[p.m] + ' ' + p.d; };

  let el = null, st = null;

  function build() {
    el = document.createElement('div');
    el.className = 'rp';
    el.setAttribute('role', 'dialog');
    el.setAttribute('aria-label', 'Select date range');
    el.hidden = true;
    document.body.appendChild(el);
    el.addEventListener('click', onClick);
    document.addEventListener('mousedown', e => {
      if (!el.hidden && !el.contains(e.target) && !(st && st.anchor.contains(e.target))) close();
    }, true);
    document.addEventListener('keydown', e => { if (!el.hidden && e.key === 'Escape') { e.stopPropagation(); close(); } }, true);
    window.addEventListener('resize', () => { if (!el.hidden) position(); });
  }

  function render() {
    const { vy, vm, start, end } = st;
    const first = new Date(vy, vm, 1).getDay();
    const days = new Date(vy, vm + 1, 0).getDate();
    const tooLong = start && end && end > addMonths(start, st.maxMonths);
    const title = start && end ? (start === end ? short(start) : short(start) + ' – ' + short(end))
      : start ? short(start) + ' – …' : 'Select dates';
    const year = start ? parse(start).y : vy;

    let cells = '';
    for (let i = 0; i < first; i++) cells += '<span></span>';
    for (let d = 1; d <= days; d++) {
      const v = iso(vy, vm, d);
      const edge = v === start || v === end;
      const inRange = start && end && v > start && v < end;
      const cls = ['rp-day', edge ? 'edge' : '', inRange ? 'in' : '', v === st.today ? 'today' : ''].join(' ');
      cells += `<button type="button" class="${cls}" data-day="${v}" aria-pressed="${edge || !!inRange}">${d}</button>`;
    }

    let hint = '';
    if (tooLong) hint = `<div class="rp-hint err">Maximum range is ${st.maxMonths} months.</div>`;
    else if (start && end && st.count) hint = `<div class="rp-hint">${st.count(start, end).toLocaleString('es-CO')} invoices with issue date in this range</div>`;
    else if (start && !end) hint = '<div class="rp-hint">Select the end date</div>';

    el.innerHTML = `
      <div class="rp-head"><div class="rp-year">${year}</div><div class="rp-title">${title}</div></div>
      <div class="rp-nav">
        <button type="button" class="rp-arrow" data-nav="-1" aria-label="Previous month">‹</button>
        <div class="rp-month">${MONTHS[vm]} ${vy}</div>
        <button type="button" class="rp-arrow" data-nav="1" aria-label="Next month">›</button>
      </div>
      <div class="rp-grid rp-wd">${['S', 'M', 'T', 'W', 'T', 'F', 'S'].map(w => `<span>${w}</span>`).join('')}</div>
      <div class="rp-grid">${cells}</div>
      ${hint}
      <div class="rp-actions">
        <button type="button" class="btn btn-text" data-act="cancel">Cancel</button>
        <button type="button" class="btn btn-text" data-act="ok" ${start && end && !tooLong ? '' : 'disabled'}>OK</button>
      </div>`;
  }

  function onClick(e) {
    const day = e.target.closest('[data-day]');
    if (day) {
      const v = day.dataset.day;
      if (!st.start || st.end) { st.start = v; st.end = null; }
      else if (v < st.start) { st.start = v; }
      else st.end = v;
      render(); return;
    }
    const nav = e.target.closest('[data-nav]');
    if (nav) {
      const dt = new Date(st.vy, st.vm + Number(nav.dataset.nav), 1);
      st.vy = dt.getFullYear(); st.vm = dt.getMonth(); render(); return;
    }
    const act = e.target.closest('[data-act]');
    if (act && act.dataset.act === 'cancel') close();
    if (act && act.dataset.act === 'ok' && !act.disabled) { const cb = st.onApply; const r = { start: st.start, end: st.end }; close(); cb(r); }
  }

  function position() {
    const r = st.anchor.getBoundingClientRect();
    const w = el.offsetWidth, h = el.offsetHeight;
    let top = r.bottom + 4;
    if (top + h > window.innerHeight - 8) top = Math.max(8, r.top - h - 4);
    el.style.top = top + 'px';
    el.style.left = Math.max(8, Math.min(r.left, window.innerWidth - w - 8)) + 'px';
  }

  function open(opts) {
    if (!el) build();
    const ref = parse(opts.end || opts.start || new Date().toISOString().slice(0, 10));
    const now = new Date();
    st = {
      anchor: opts.anchor, start: opts.start || null, end: opts.end || null,
      maxMonths: opts.maxMonths || 3, count: opts.count, onApply: opts.onApply,
      vy: ref.y, vm: ref.m, today: iso(now.getFullYear(), now.getMonth(), now.getDate())
    };
    render();
    el.hidden = false;
    position();
    st.anchor.classList.add('open');
  }

  function close() {
    if (!el || el.hidden) return;
    el.hidden = true;
    if (st) st.anchor.classList.remove('open');
  }

  global.RangePicker = { open, close };
})(window);
