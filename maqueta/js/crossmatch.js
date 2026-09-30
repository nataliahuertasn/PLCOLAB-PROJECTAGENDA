/* =========================================================
   CrossMatch — lógica pura del cruce (sin DOM)
   Compara registro contra registro usando una llave de documento:
     Proveedor normalizado + Nº de factura normalizado
   (o solo Nº de factura si el reporte externo no trae proveedor).
   ========================================================= */
(function (global) {
  'use strict';

  // ---------- Normalización ----------
  function stripAccents(s) {
    return s.normalize('NFD').replace(/[̀-ͯ]/g, '');
  }

  /** Nº de factura: mayúsculas, sin espacios, guiones ni puntos. "fe-40930 " → "FE40930" */
  function normalizeInvoice(v) {
    if (v === null || v === undefined) return '';
    return stripAccents(String(v)).toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  /** Proveedor: sin tildes, sin puntuación y sin sufijo societario.
   *  "TRANSCONT S.A.S" = "TRANSCONT SAS" = "Transcont S.A.S." → "TRANSCONT" */
  const LEGAL_SUFFIX = /\s(S\s?A\s?S|S\s?A|S\s?EN\s?C(\s?S)?|LTDA|LTD|LIMITED|INC|LLC|CORP|CORPORATION|CO|E\s?U|GMBH|SA\sDE\sCV)$/;
  function normalizeSupplier(v) {
    if (v === null || v === undefined) return '';
    let s = stripAccents(String(v)).toUpperCase().replace(/\./g, '').replace(/[^A-Z0-9]+/g, ' ').trim();
    let prev;
    do { prev = s; s = s.replace(LEGAL_SUFFIX, '').trim(); } while (s !== prev && s.length);
    return s.replace(/\s+/g, '');
  }

  /** Valor: interpreta formatos mixtos presentes en el Excel
   *  ("86,388,016.10", "19.782.023", "99.999,99", "421235.00"). Devuelve número o null. */
  function parseAmount(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    let s = String(v).trim().replace(/[\s$]/g, '').replace(/[A-Za-z]/g, '');
    if (!s) return null;
    const neg = /^-|^\(.*\)$/.test(s);
    s = s.replace(/[()\-]/g, '');
    const lastDot = s.lastIndexOf('.'), lastComma = s.lastIndexOf(',');
    if (lastDot > -1 && lastComma > -1) {
      const dec = lastDot > lastComma ? '.' : ',';
      const thou = dec === '.' ? ',' : '.';
      s = s.split(thou).join('').replace(dec, '.');
    } else if (lastDot > -1 || lastComma > -1) {
      const sep = lastDot > -1 ? '.' : ',';
      const parts = s.split(sep);
      if (parts.length > 2) s = parts.join('');                          // 19.782.023
      else if (parts[1].length === 3 && parts[0] !== '0') s = parts.join(''); // 999.999 → miles
      else s = parts.join('.');                                          // 421235.00 / 4.7064831
    }
    const n = parseFloat(s);
    return isFinite(n) ? (neg ? -n : n) : null;
  }

  function pad(n) { return String(n).padStart(2, '0'); }

  /** Fecha → "YYYY-MM-DD" (acepta Date, serial de Excel, ISO, dd/mm/yyyy). */
  function normalizeDate(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date && !isNaN(v)) {
      const d = new Date(v.getTime() + 60000);
      return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    }
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) return m[1] + '-' + pad(m[2]) + '-' + pad(m[3]);
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
    if (m) {
      let d = +m[1], mo = +m[2], y = +m[3];
      if (y < 100) y += 2000;
      if (mo > 12 && d <= 12) { const t = d; d = mo; mo = t; } // formato mm/dd
      return y + '-' + pad(mo) + '-' + pad(d);
    }
    return s;
  }

  function normalizeOC(v) {
    return v === null || v === undefined ? '' : String(v).trim().replace(/\.0+$/, '');
  }

  // ---------- Detección de encabezados y columnas ----------
  // Regla fija del cruce: Invoice ID del software ↔ columna "Nro. Documento" del reporte cargado.
  const INVOICE_HEADER = 'Nro. Documento';
  const HEADER_ALIASES = {
    invoiceId: [headerKey(INVOICE_HEADER)],
    supplier:  ['remitente', 'supplier', 'proveedor', 'nombreproveedor', 'vendor', 'tercero', 'nombretercero', 'razonsocial', 'emisor', 'nombreemisor', 'razonsocialemisor'],
    amount:    ['totaldocumento', 'invoiceamount', 'amount', 'valor', 'valortotal', 'valorfactura', 'monto', 'total', 'totalfactura', 'importe'],
    issueDate: ['issuedate', 'fechaemision', 'fechadeemision', 'fechafactura', 'fechadocumento', 'fecha', 'date'],
    dueDate:   ['duedate', 'fechavencimiento', 'fechadevencimiento', 'vencimiento'],
    oc:        ['oc', 'ordencompra', 'ordendecompra', 'po', 'purchaseorder', 'pedido', 'nooc'],
    type:      ['type', 'tipo', 'tipodocumento']
  };
  function headerKey(h) { return stripAccents(String(h || '')).toLowerCase().replace(/[^a-z0-9]/g, ''); }

  /** Fila de encabezados: la que contiene "Nro. Documento"; si no existe, la de más coincidencias de alias. */
  function detectHeaderRow(rows) {
    const inv = headerKey(INVOICE_HEADER);
    for (let i = 0; i < Math.min(rows.length, 25); i++) {
      if ((rows[i] || []).some(c => headerKey(c) === inv)) return i;
    }
    let best = { index: 0, score: -1 };
    const all = Object.values(HEADER_ALIASES).flat();
    for (let i = 0; i < Math.min(rows.length, 25); i++) {
      const score = (rows[i] || []).filter(c => all.includes(headerKey(c))).length;
      if (score > best.score) best = { index: i, score };
    }
    return best.index;
  }

  /** Propone el mapeo de columnas del archivo externo → campos lógicos. */
  function autoMap(headers) {
    const keys = headers.map(headerKey);
    const map = {};
    const used = new Set();
    Object.entries(HEADER_ALIASES).forEach(([field, aliases]) => {
      for (const alias of aliases) {                      // coincidencia exacta, en orden de prioridad
        const i = keys.findIndex((k, idx) => k === alias && !used.has(idx));
        if (i > -1) { map[field] = i; used.add(i); return; }
      }
    });
    if (map.supplier === undefined) {
      const i = keys.findIndex((k, idx) => !used.has(idx) && /(proveedor|supplier|tercero|razonsocial)/.test(k));
      if (i > -1) { map.supplier = i; used.add(i); }
    }
    return map;
  }

  // ---------- Cruce ----------
  const FIELD_LABELS = { amount: 'Valor', issueDate: 'Fecha emisión', oc: 'OC' };

  function makeKey(rec, useSupplier) {
    const inv = normalizeInvoice(rec.invoiceId);
    if (!inv) return '';
    return useSupplier ? normalizeSupplier(rec.supplier) + '::' + inv : inv;
  }

  function groupByKey(records, useSupplier, errors, sourceName) {
    const groups = new Map();
    records.forEach(rec => {
      const key = makeKey(rec, useSupplier);
      if (!key) {
        errors.push({ source: sourceName, rec, reason: 'Registro sin Nº de factura: no se puede cruzar' });
        return;
      }
      rec._key = key;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(rec);
    });
    return groups;
  }

  function differingFields(recs, fields) {
    return fields.filter(f => new Set(recs.map(r => compareValue(f, r))).size > 1);
  }

  function compareValue(field, rec) {
    if (field === 'amount') { const n = parseAmount(rec.amount); return n === null ? '' : Math.round(n * 100) / 100; }
    if (field === 'issueDate') return normalizeDate(rec.issueDate);
    if (field === 'oc') return normalizeOC(rec.oc);
    return String(rec[field] ?? '').trim();
  }

  /**
   * @param {Array} software    registros de la vista del software (filtrados por período)
   * @param {Array} external    registros del reporte externo
   * @param {Object} opts       { useSupplier, compareFields: ['amount','issueDate','oc'], softwareAll }
   */
  function run(software, external, opts) {
    const useSupplier = !!opts.useSupplier;
    const compareFields = opts.compareFields || [];
    const errors = [];

    const swGroups = groupByKey(software, useSupplier, errors, 'software');
    const exGroups = groupByKey(external, useSupplier, errors, 'externo');

    // Índices auxiliares para pistas ("hints")
    const allSw = opts.softwareAll || software;
    const swAllByKey = new Map();
    const swAllByInvoice = new Map();
    allSw.forEach(r => {
      const k = makeKey(r, useSupplier); if (!k) return;
      if (!swAllByKey.has(k)) swAllByKey.set(k, []);
      swAllByKey.get(k).push(r);
      const inv = normalizeInvoice(r.invoiceId);
      if (!swAllByInvoice.has(inv)) swAllByInvoice.set(inv, []);
      swAllByInvoice.get(inv).push(r);
    });
    const exByInvoice = new Map();
    external.forEach(r => {
      const inv = normalizeInvoice(r.invoiceId); if (!inv) return;
      if (!exByInvoice.has(inv)) exByInvoice.set(inv, []);
      exByInvoice.get(inv).push(r);
    });

    const matched = [], onlyExternal = [], onlySoftware = [], duplicates = [];

    exGroups.forEach((exRecs, key) => {
      const swRecs = swGroups.get(key);
      if (swRecs) {
        const diffs = [];
        // Hay diferencia en un campo solo si ningún registro del software comparte valor
        // con algún registro externo (evita falsos positivos cuando la factura está duplicada).
        compareFields.forEach(f => {
          const same = (a, b) => f === 'amount' && a !== '' && b !== '' ? Math.abs(a - b) <= 1 : a === b;
          const swVals = swRecs.map(r => compareValue(f, r)), exVals = exRecs.map(r => compareValue(f, r));
          const agree = swVals.some(a => exVals.some(b => same(a, b)));
          if (!agree) {
            const uniq = (recs) => [...new Set(recs.map(r => r[f] ?? ''))].join(' | ');
            diffs.push({ field: f, label: FIELD_LABELS[f] || f, software: uniq(swRecs), external: uniq(exRecs) });
          }
        });
        matched.push({ key, software: swRecs, external: exRecs, diffs });
      } else {
        let hint = '';
        const outOfPeriod = swAllByKey.get(key);
        if (outOfPeriod && outOfPeriod.length) {
          hint = 'Existe en el software con fecha de emisión ' + outOfPeriod[0].issueDate + ' (fuera del período seleccionado).';
        } else if (useSupplier) {
          const sameInv = swAllByInvoice.get(normalizeInvoice(exRecs[0].invoiceId));
          if (sameInv && sameInv.length) hint = 'El mismo Nº de factura está registrado en el software para: ' + [...new Set(sameInv.map(r => r.supplier))].join(', ') + '.';
        }
        onlyExternal.push({ key, records: exRecs, hint });
      }
    });

    swGroups.forEach((swRecs, key) => {
      if (exGroups.has(key)) return;
      let hint = '';
      if (useSupplier) {
        const sameInv = exByInvoice.get(normalizeInvoice(swRecs[0].invoiceId));
        if (sameInv && sameInv.length) hint = 'El mismo Nº de factura aparece en el reporte externo para: ' + [...new Set(sameInv.map(r => r.supplier))].join(', ') + '.';
      }
      onlySoftware.push({ key, records: swRecs, hint });
    });

    const dupFields = ['amount', 'issueDate', 'oc', 'supplier', 'type', 'dueDate'];
    [['software', swGroups], ['externo', exGroups]].forEach(([source, groups]) => {
      groups.forEach((recs, key) => {
        if (recs.length < 2) return;
        const varying = differingFields(recs, dupFields.filter(f => recs.some(r => r[f] !== undefined)));
        duplicates.push({ source, key, records: recs, count: recs.length, identical: varying.length === 0, varying });
      });
    });

    function sourceStats(records, groups, source) {
      const dups = duplicates.filter(d => d.source === source);
      return {
        rows: records.length,
        unique: groups.size,
        dupGroups: dups.length,
        dupRows: dups.reduce((s, d) => s + d.count, 0),
        dupExtra: dups.reduce((s, d) => s + d.count - 1, 0),
        noId: errors.filter(e => e.source === source).length
      };
    }

    // Estado por registro del software (para marcar la tabla principal)
    const softwareStatus = new Map();
    matched.forEach(m => m.software.forEach(r => softwareStatus.set(r, m.diffs.length ? 'diff' : 'ok')));
    onlySoftware.forEach(o => o.records.forEach(r => softwareStatus.set(r, 'alert')));
    duplicates.filter(d => d.source === 'software').forEach(d => d.records.forEach(r => {
      softwareStatus.set(r, (softwareStatus.get(r) || '') + '+dup');
    }));
    errors.filter(e => e.source === 'software').forEach(e => softwareStatus.set(e.rec, 'error'));

    const matchedRows = { software: 0, external: 0 };
    matched.forEach(m => { matchedRows.software += m.software.length; matchedRows.external += m.external.length; });

    return {
      useSupplier,
      compareFields,
      software: sourceStats(software, swGroups, 'software'),
      external: sourceStats(external, exGroups, 'externo'),
      matched,
      matchedRows,
      matchedWithDiffs: matched.filter(m => m.diffs.length),
      onlyExternal,
      onlySoftware,
      duplicates,
      errors,
      softwareStatus
    };
  }

  global.CrossMatch = {
    normalizeInvoice, normalizeSupplier, parseAmount, normalizeDate, normalizeOC,
    detectHeaderRow, autoMap, headerKey, run, FIELD_LABELS, INVOICE_HEADER
  };
})(window);
