/* Sift — swipe-based title/abstract screening that writes decisions back to Excel.
 * A project is a named review holding one or more spreadsheet files (e.g. one export per database).
 * Everything runs in the browser; projects live in IndexedDB on this device.
 * On desktop Chrome/Edge each Excel file is updated in the background (File System Access API).
 * Elsewhere (iPhone, Android, Safari, Firefox) "Save Excel file" hands you the updated files. */
(() => {
  'use strict';

  const XL = window.XLSX;
  const app = document.getElementById('app');
  const fileInput = document.getElementById('file-input');
  const toastEl = document.getElementById('toast');

  const INCLUDE = 1, EXCLUDE = 0, MAYBE = 2;
  const LABEL = { 1: 'Included', 0: 'Excluded', 2: 'Maybe' };
  const CLS = { 1: 'include', 0: 'exclude', 2: 'maybe' };
  const STAGE_NAME = { t: 'Title screening', a: 'Abstract screening' };
  const DEFAULT_HEADER = { t: 'Title screening', a: 'Abstract screening' };

  // Column finder: a Supabase Edge Function that holds the Claude key server-side.
  // These two values are public by design (they identify the project, they grant nothing secret).
  const AI_URL = 'https://cmerdamunvelqipstkbx.supabase.co/functions/v1/identify-columns';
  const AI_ANON = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNtZXJkYW11bnZlbHFpcHN0a2J4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTQ2ODM2NzAsImV4cCI6MjExMDI1OTY3MH0.KSksFl1TzAEC0OVUKi1xkJJgRj_LhNYJZE80Rnn_Sm0';

  const canLiveSave = typeof window.showOpenFilePicker === 'function' && (() => {
    try { return window.top === window.self; } catch { return false; }
  })();
  const isTouch = window.matchMedia('(pointer: coarse)').matches;

  /* ------------------------------------------------------------------ storage
   * projects: { id, name, fileIds[], pos:{t,a}, sample, createdAt, updatedAt, stats }
   * sources:  { id, projectId, fileName, ext, sheet, cols, newHeaders, fingerprint, decisions:{t,a}, unsaved, stats }
   * files:    { id (= source id), bytes, handle } */
  let idb = null;
  const mem = { projects: new Map(), sources: new Map(), files: new Map(), handles: new Map() };
  // Live-file handles live in their own store so opening a spreadsheet never has to read one, plus a per-visit cache.
  // (Chrome private windows can end the page when a stored handle is read back; see readHandle.)
  const handleCache = new Map();
  let skipStoredHandles = false;

  function convertPos(o) {
    const out = { t: null, a: null };
    for (const k of ['t', 'a']) {
      const v = o.pos && o.pos[k];
      out[k] = v == null ? null : v === -1 ? -1 : { sid: o.id, r: v };
    }
    return out;
  }
  function openIDB() {
    return new Promise(resolve => {
      try {
        const rq = indexedDB.open('sift-screener', 3);
        rq.onupgradeneeded = ev => {
          const d = rq.result;
          const tx = rq.transaction;
          if (ev.oldVersion < 1) {
            d.createObjectStore('projects', { keyPath: 'id' });
            d.createObjectStore('files', { keyPath: 'id' });
          }
          if (ev.oldVersion < 2) {
            const src = d.createObjectStore('sources', { keyPath: 'id' });
            if (ev.oldVersion >= 1) {
              // v1 stored one file per project: turn each into a project holding that one file.
              const ps = tx.objectStore('projects');
              ps.getAll().onsuccess = e => {
                for (const o of e.target.result || []) {
                  if (o.fileIds) continue;
                  const pid = 'P' + o.id;
                  const s = { ...o, projectId: pid, addedAt: o.createdAt || Date.now() };
                  delete s.pos; delete s.name; delete s.sample; delete s.live;
                  src.put(s);
                  ps.delete(o.id);
                  ps.put({ id: pid, name: o.name, fileIds: [o.id], pos: convertPos(o), sample: !!o.sample, stats: o.stats,
                    createdAt: o.createdAt || Date.now(), updatedAt: o.updatedAt || Date.now() });
                }
              };
            }
          }
          if (ev.oldVersion < 3) {
            const hs = d.createObjectStore('handles', { keyPath: 'id' });
            if (ev.oldVersion >= 1) {
              tx.objectStore('files').openCursor().onsuccess = e => {
                const cursor = e.target.result;
                if (!cursor) return;
                const rec = cursor.value;
                if (rec.handle) { hs.put({ id: rec.id, handle: rec.handle }); cursor.update({ id: rec.id, bytes: rec.bytes }); }
                cursor.continue();
              };
            }
          }
        };
        rq.onsuccess = () => resolve(rq.result);
        rq.onerror = () => resolve(null);
        rq.onblocked = () => resolve(null);
      } catch { resolve(null); }
    });
  }
  function req(store, mode, fn) {
    if (!idb) {
      const m = mem[store];
      return Promise.resolve(fn({
        getAll: () => ({ result: [...m.values()] }),
        get: k => ({ result: m.get(k) }),
        put: v => { m.set(v.id, v); return {}; },
        delete: k => { m.delete(k); return {}; },
      }).result);
    }
    return new Promise((resolve, reject) => {
      try {
        const t = idb.transaction(store, mode);
        const r = fn(t.objectStore(store));
        t.oncomplete = () => resolve(r ? r.result : undefined);
        t.onerror = () => reject(t.error);
        t.onabort = () => reject(t.error);
      } catch (e) { reject(e); }
    });
  }
  const db = {
    all: s => req(s, 'readonly', st => st.getAll()),
    get: (s, k) => req(s, 'readonly', st => st.get(k)),
    put: (s, v) => req(s, 'readwrite', st => st.put(v)),
    del: (s, k) => req(s, 'readwrite', st => st.delete(k)),
  };
  const storageFail = () => toast('Couldn’t save progress on this device. Check that site storage is allowed.', 5000);
  const saveProject = p => db.put('projects', p).catch(storageFail);
  const saveSource = s => db.put('sources', s).catch(storageFail);

  /* ------------------------------------------------------------------ helpers */
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = n => Number(n || 0).toLocaleString();
  const baseName = f => String(f || '').replace(/\.[^.]+$/, '');
  const extOf = f => (String(f || '').match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
  const uid = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const plural = (n, w) => `${fmt(n)} ${w}${n === 1 ? '' : 's'}`;

  let toastTimer = null;
  function toast(msg, ms = 2600) {
    toastEl.textContent = msg;
    toastEl.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toastEl.hidden = true; }, ms);
  }

  const ICON = {
    back: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 18l-6-6 6-6"/></svg>',
    undo: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 14L4 9l5-5"/><path d="M4 9h11a5 5 0 010 10h-3"/></svg>',
    x: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    check: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
    q: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 9a3 3 0 115 2.2c-1.2.8-2 1.4-2 2.8"/><path d="M12 18h.01"/></svg>',
    chev: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>',
    magic: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 20L15 9"/><path d="M13.5 7.5l3 3"/><path d="M18 2.5v3M16.5 4h3"/><path d="M20.5 9v2M19.5 10h2"/><path d="M9 3v2M8 4h2"/></svg>',
    criteria: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6h11M9 12h11M9 18h11"/><path d="M4 6l1 1 2-2M4 12l1 1 2-2M4 18l1 1 2-2"/></svg>',
    close: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    file: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round" aria-hidden="true"><path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/><path d="M9 12h6M9 16h6"/></svg>',
  };

  /* ------------------------------------------------------------------ spreadsheet */
  function readWorkbook(bytes) {
    return XL.read(new Uint8Array(bytes), { type: 'array', cellStyles: true, cellNF: true, cellDates: false });
  }
  function rangeOf(ws) { return XL.utils.decode_range(ws['!ref'] || 'A1:A1'); }
  function cellText(ws, r, c) {
    if (c == null || c < 0) return '';
    const cell = ws[XL.utils.encode_cell({ r, c })];
    if (!cell) return '';
    const v = cell.w != null ? cell.w : cell.v;
    return v == null ? '' : String(v).trim();
  }
  function cellDecision(ws, r, c) {
    if (c == null || c < 0) return undefined;
    const cell = ws[XL.utils.encode_cell({ r, c })];
    if (!cell || cell.v == null) return undefined;
    const raw = String(cell.v).trim();
    if (raw === '') return undefined; // blank or space-only cells are unscreened (Number('') would be 0)
    const n = Number(raw);
    return n === 0 || n === 1 || n === 2 ? n : undefined;
  }
  function headersOf(ws) {
    const rg = rangeOf(ws);
    const out = [];
    for (let c = rg.s.c; c <= rg.e.c; c++) {
      const name = cellText(ws, rg.s.r, c);
      out.push({ c, name, label: (name || 'Untitled') + '  (col ' + XL.utils.encode_col(c) + ')' });
    }
    return out;
  }
  // Rows that hold a record (title or abstract present). r = 0-based sheet row.
  function readRows(ws, cols) {
    const rg = rangeOf(ws);
    const rows = [];
    for (let r = rg.s.r + 1; r <= rg.e.r; r++) {
      const t = cellText(ws, r, cols.title);
      const a = cellText(ws, r, cols.abs);
      if (!t && !a) continue;
      rows.push({ r, t, a, ft: cellDecision(ws, r, cols.titleDec), fa: cellDecision(ws, r, cols.absDec) });
    }
    return rows;
  }
  function fingerprint(sheet, rows) {
    let h = 2166136261;
    const s = sheet + '|' + rows.length + '|' + rows.slice(0, 25).map(x => x.t).join('|');
    for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
    return (h >>> 0).toString(36);
  }
  function applyDecisions(ws, s) {
    const rg = rangeOf(ws);
    let maxC = rg.e.c;
    for (const [c, name] of Object.entries(s.newHeaders || {})) {
      const addr = XL.utils.encode_cell({ r: rg.s.r, c: +c });
      if (!ws[addr] || ws[addr].v == null || ws[addr].v === '') ws[addr] = { t: 's', v: name };
      maxC = Math.max(maxC, +c);
    }
    for (const [key, col] of [['t', s.cols.titleDec], ['a', s.cols.absDec]]) {
      maxC = Math.max(maxC, col);
      for (const [r, v] of Object.entries(s.decisions[key])) {
        ws[XL.utils.encode_cell({ r: +r, c: col })] = { t: 'n', v };
      }
    }
    rg.e.c = maxC;
    ws['!ref'] = XL.utils.encode_range(rg);
  }
  // File types Sift can write back into in place (background saving).
  const LIVE_TYPES = ['xlsx', 'xlsm', 'csv'];

  /* Write decisions into an .xlsx/.xlsm by editing only the decision cells inside the file's zip.
   * Every other part (styles, conditional formatting, filters, other sheets, macros) stays byte-for-byte.
   * Throws if the file has a shape this doesn't handle; the caller then falls back to a full rewrite. */
  function patchXlsx(bytes, s) {
    const F = window.fflate;
    if (!F) throw new Error('zip library missing');
    const files = F.unzipSync(new Uint8Array(bytes));
    const text = p => { if (!files[p]) throw new Error('missing ' + p); return F.strFromU8(files[p]); };
    const parse = str => {
      const d = new DOMParser().parseFromString(str, 'application/xml');
      if (d.getElementsByTagName('parsererror').length) throw new Error('xml parse error');
      return d;
    };
    const declOf = str => (str.match(/^\s*<\?xml[^>]*\?>\s*/) || [''])[0].trim();
    // XMLSerializer may emit its own declaration; keep exactly one (the file's original), or Excel reports damage.
    const serialize = (d, decl) => (decl ? decl + '\r\n' : '') + new XMLSerializer().serializeToString(d).replace(/^\s*<\?xml[^>]*\?>\s*/, '');
    const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    // Locate the sheet's XML part.
    const wbStr = text('xl/workbook.xml');
    const wbDoc = parse(wbStr);
    const sheetEl = [...wbDoc.getElementsByTagName('sheet')].find(e => e.getAttribute('name') === s.sheet);
    if (!sheetEl) throw new Error('sheet not found');
    const rid = sheetEl.getAttributeNS(R_NS, 'id') || sheetEl.getAttribute('r:id');
    const relsDoc = parse(text('xl/_rels/workbook.xml.rels'));
    const rel = [...relsDoc.getElementsByTagName('Relationship')].find(e => e.getAttribute('Id') === rid);
    if (!rel) throw new Error('sheet relationship not found');
    const target = rel.getAttribute('Target');
    const sheetPath = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');

    // Collect the edits: header cells for new decision columns, then every decision.
    const edits = new Map(); // "r,c" -> {r, c, v} (r 0-based)
    const headerRow = s.headerRow != null ? s.headerRow : 0;
    for (const [c, name] of Object.entries(s.newHeaders || {})) edits.set(headerRow + ',' + c, { r: headerRow, c: +c, v: String(name), header: true });
    for (const [key, col] of [['t', s.cols.titleDec], ['a', s.cols.absDec]]) {
      for (const [r, v] of Object.entries(s.decisions[key])) edits.set(r + ',' + col, { r: +r, c: col, v: Number(v) });
    }

    const shStr = text(sheetPath);
    const shDoc = parse(shStr);
    const NS = shDoc.documentElement.namespaceURI;
    const sheetData = shDoc.getElementsByTagNameNS(NS, 'sheetData')[0];
    if (!sheetData) throw new Error('no sheetData');
    const rowEls = new Map();
    for (const row of sheetData.getElementsByTagNameNS(NS, 'row')) {
      const n = +row.getAttribute('r');
      if (!n) throw new Error('row without number');
      rowEls.set(n, row);
    }
    let removedFormula = false, maxCol = -1, maxRow = -1;
    const byRow = new Map();
    for (const e of edits.values()) { if (!byRow.has(e.r)) byRow.set(e.r, []); byRow.get(e.r).push(e); }
    for (const [r, list] of byRow) {
      const rowNum = r + 1;
      let row = rowEls.get(rowNum);
      if (!row) {
        row = shDoc.createElementNS(NS, 'row');
        row.setAttribute('r', String(rowNum));
        const after = [...rowEls.keys()].filter(n => n > rowNum).sort((a, b) => a - b)[0];
        sheetData.insertBefore(row, after ? rowEls.get(after) : null);
        rowEls.set(rowNum, row);
      }
      row.removeAttribute('spans'); // only a layout hint; stale once a row gains cells
      const cells = [...row.getElementsByTagNameNS(NS, 'c')];
      if (cells.some(c => !c.getAttribute('r'))) throw new Error('cell without reference');
      for (const e of list) {
        const ref = XL.utils.encode_cell({ r: e.r, c: e.c });
        let cell = cells.find(c => c.getAttribute('r') === ref);
        if (cell) {
          const f = cell.getElementsByTagNameNS(NS, 'f')[0];
          if (f) {
            if (f.getAttribute('t') === 'shared' && f.getAttribute('ref')) throw new Error('shared formula master');
            removedFormula = true;
          }
          for (const a of [...cell.attributes]) if (a.name !== 'r' && a.name !== 's') cell.removeAttribute(a.name); // keep the cell's style
          while (cell.firstChild) cell.removeChild(cell.firstChild);
        } else {
          cell = shDoc.createElementNS(NS, 'c');
          cell.setAttribute('r', ref);
          const next = cells.find(c => XL.utils.decode_cell(c.getAttribute('r')).c > e.c);
          row.insertBefore(cell, next || null);
          cells.splice(next ? cells.indexOf(next) : cells.length, 0, cell);
        }
        if (e.header) {
          cell.setAttribute('t', 'inlineStr');
          const is = shDoc.createElementNS(NS, 'is'), t = shDoc.createElementNS(NS, 't');
          t.textContent = e.v; is.appendChild(t); cell.appendChild(is);
        } else {
          const v = shDoc.createElementNS(NS, 'v');
          v.textContent = String(e.v); cell.appendChild(v);
        }
        maxCol = Math.max(maxCol, e.c); maxRow = Math.max(maxRow, e.r);
      }
    }
    // Keep the sheet's declared size covering the cells written.
    const dim = shDoc.getElementsByTagNameNS(NS, 'dimension')[0];
    if (dim && dim.getAttribute('ref')) {
      const d = XL.utils.decode_range(dim.getAttribute('ref'));
      d.e.c = Math.max(d.e.c, maxCol); d.e.r = Math.max(d.e.r, maxRow);
      dim.setAttribute('ref', XL.utils.encode_range(d));
    }
    files[sheetPath] = F.strToU8(serialize(shDoc, declOf(shStr)));

    // Ask Excel to recalculate formulas on open (e.g. columns that combine two reviewers' decisions).
    let calc = wbDoc.getElementsByTagNameNS(wbDoc.documentElement.namespaceURI, 'calcPr')[0];
    if (!calc) {
      calc = wbDoc.createElementNS(wbDoc.documentElement.namespaceURI, 'calcPr');
      const later = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
      const before = [...wbDoc.documentElement.children].find(el => later.includes(el.localName));
      wbDoc.documentElement.insertBefore(calc, before || null);
    }
    calc.setAttribute('fullCalcOnLoad', '1');
    files['xl/workbook.xml'] = F.strToU8(serialize(wbDoc, declOf(wbStr)));

    // If a formula cell was overwritten, Excel's calculation chain would point at it; drop the chain (Excel rebuilds it).
    if (removedFormula && files['xl/calcChain.xml']) {
      delete files['xl/calcChain.xml'];
      for (const p of ['xl/_rels/workbook.xml.rels', '[Content_Types].xml']) {
        const str = text(p), d = parse(str);
        for (const el of [...d.documentElement.children]) {
          if ((el.getAttribute('Target') || '').endsWith('calcChain.xml') || (el.getAttribute('PartName') || '').endsWith('calcChain.xml')) el.parentNode.removeChild(el);
        }
        files[p] = F.strToU8(serialize(d, declOf(str)));
      }
    }
    return F.zipSync(files, { level: 6 });
  }

  function outputFile(src) {
    const s = src.s;
    const ws = src.wb.Sheets[s.sheet];
    applyDecisions(ws, s); // keeps the in-memory copy current
    if (s.ext === 'csv') {
      const csv = XL.utils.sheet_to_csv(ws);
      return { name: s.fileName, blob: new Blob(['﻿' + csv], { type: 'text/csv' }) };
    }
    const type = s.ext === 'xlsm' ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    if ((s.ext === 'xlsx' || s.ext === 'xlsm') && src.bytes) {
      try {
        return { name: s.fileName, blob: new Blob([patchXlsx(src.bytes, s)], { type }), preserved: true };
      } catch (e) {
        console.warn('Sift: in-place save not possible, rewriting the workbook instead.', e);
      }
    }
    // Fallback (and .xls): rebuild the workbook. Cell formatting is not kept on this path.
    // Old .xls column outlines arrive as "level", which is not a valid .xlsx attribute (outlineLevel is kept).
    for (const name of src.wb.SheetNames) for (const col of src.wb.Sheets[name]['!cols'] || []) if (col) delete col.level;
    const out = XL.write(src.wb, { bookType: 'xlsx', type: 'array', cellStyles: true, compression: true });
    const name = s.ext === 'xlsx' ? s.fileName : baseName(s.fileName) + '.xlsx';
    return { name, blob: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), preserved: false };
  }
  // Fold decisions already in the file into Sift's record. Returns true if Sift holds decisions the file lacks.
  function mergeFileDecisions(s, rows) {
    let fileBehind = false;
    for (const x of rows) {
      if (x.ft !== undefined && s.decisions.t[x.r] === undefined) s.decisions.t[x.r] = x.ft;
      if (x.fa !== undefined && s.decisions.a[x.r] === undefined) s.decisions.a[x.r] = x.fa;
      if (s.decisions.t[x.r] !== undefined && s.decisions.t[x.r] !== x.ft) fileBehind = true;
      if (s.decisions.a[x.r] !== undefined && s.decisions.a[x.r] !== x.fa) fileBehind = true;
    }
    return fileBehind;
  }

  /* ------------------------------------------------------------------ state */
  let projects = [];
  let cur = null;    // { p, srcs: [{ s, wb, rows, handle, live, reconnect }] }
  let screen = null; // { stage, queue: [{ k, x }], qi, busy }
  let view = 'home';

  const emptyStats = () => ({ t: { total: 0, done: 0, 1: 0, 0: 0, 2: 0 }, a: { total: 0, done: 0, 1: 0, 0: 0, 2: 0 } });
  function sourceStats(s, rows) {
    const d = s.decisions;
    const st = emptyStats();
    st.t.total = rows.length;
    for (const x of rows) {
      const tv = d.t[x.r];
      if (tv !== undefined) { st.t.done++; st.t[tv]++; }
      if (tv === INCLUDE || tv === MAYBE) {
        st.a.total++;
        const av = d.a[x.r];
        if (av !== undefined) { st.a.done++; st.a[av]++; }
      }
    }
    return st;
  }
  function addStats(a, b) {
    for (const k of ['t', 'a']) for (const f of ['total', 'done', 1, 0, 2]) a[k][f] += b[k][f];
    return a;
  }
  function refreshStats() {
    const total = emptyStats();
    for (const src of cur.srcs) { src.s.stats = sourceStats(src.s, src.rows); addStats(total, src.s.stats); }
    cur.p.stats = total;
    return total;
  }
  async function loadProjects() {
    projects = (await db.all('projects').catch(() => [])) || [];
    projects.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  }
  const touchProject = () => { cur.p.updatedAt = Date.now(); saveProject(cur.p); };

  /* ------------------------------------------------------------------ home */
  function barHTML(s) {
    if (!s || !s.total) return '<span class="bar"></span>';
    const w = k => (100 * s[k] / s.total).toFixed(2) + '%';
    return `<span class="bar"><i class="i" style="width:${w(1)}"></i><i class="m" style="width:${w(2)}"></i><i class="e" style="width:${w(0)}"></i></span>`;
  }
  // A real progress bar: share of records screened.
  function meterHTML(done, total) {
    const pct = total ? Math.round(100 * done / total) : 0;
    return `<div class="meter" role="progressbar" aria-valuemin="0" aria-valuemax="${total}" aria-valuenow="${done}" aria-label="${fmt(done)} of ${fmt(total)} screened"><i style="width:${total ? (100 * done / total).toFixed(2) : 0}%"></i></div>`;
  }
  // Donut chart of include / maybe / exclude, with a legend that carries the figures.
  const SPLIT = [[1, 'Include', 'inc'], [2, 'Maybe', 'may'], [0, 'Exclude', 'exc']];
  function splitHTML(st) {
    const decided = st[1] + st[2] + st[0];
    const R = 42, SW = 14, C = 2 * Math.PI * R, GAP = 2;
    const parts = SPLIT.filter(([k]) => st[k] > 0);
    let offset = 0, arcs = '';
    for (const [k, name, cls] of parts) {
      const frac = st[k] / decided;
      const len = Math.max(0.5, frac * C - (parts.length > 1 ? GAP : 0));
      const tip = `${name}: ${fmt(st[k])} (${Math.round(frac * 100)}%)`;
      arcs += `<circle class="arc ${cls}" cx="50" cy="50" r="${R}" stroke-width="${SW}" stroke-dasharray="${len.toFixed(2)} ${(C - len).toFixed(2)}" stroke-dashoffset="${(-offset).toFixed(2)}" data-tip="${esc(tip)}"><title>${esc(tip)}</title></circle>`;
      offset += frac * C;
    }
    const legend = SPLIT.map(([k, name, cls]) => `<li><span class="sw ${cls}" aria-hidden="true"></span><span class="lg-name">${name}</span><b class="num">${fmt(st[k])}</b><span class="lg-pct num">${decided ? Math.round(100 * st[k] / decided) + '%' : '–'}</span></li>`).join('');
    return `<div class="split">
      <div class="donut-wrap">
        <svg class="donut" viewBox="0 0 100 100" role="img" aria-label="${SPLIT.map(([k, n]) => `${n} ${st[k]}`).join(', ')}">
          <circle class="track" cx="50" cy="50" r="${R}" stroke-width="${SW}"/>
          <g transform="rotate(-90 50 50)">${arcs}</g>
        </svg>
        <div class="donut-label"><b class="num">${fmt(decided)}</b><span>${decided ? 'decided' : 'none yet'}</span></div>
        <div class="donut-tip" hidden></div>
      </div>
      <ul class="legend">${legend}</ul>
    </div>`;
  }
  // Hover / tap a slice to see its figures.
  app.addEventListener('pointerover', e => {
    const arc = e.target.closest && e.target.closest('.arc');
    if (!arc) return;
    const wrap = arc.closest('.donut-wrap'), tip = wrap.querySelector('.donut-tip');
    tip.textContent = arc.dataset.tip; tip.hidden = false;
    wrap.querySelectorAll('.arc').forEach(a => a.classList.toggle('dim', a !== arc));
  });
  app.addEventListener('pointerout', e => {
    const arc = e.target.closest && e.target.closest('.arc');
    if (!arc) return;
    const wrap = arc.closest('.donut-wrap');
    wrap.querySelector('.donut-tip').hidden = true;
    wrap.querySelectorAll('.arc').forEach(a => a.classList.remove('dim'));
  });

  function renderHome() {
    view = 'home'; cur = null; screen = null;
    const list = projects.map(p => {
      const s = p.stats || emptyStats();
      const n = (p.fileIds || []).length;
      return `<button class="project" data-open="${esc(p.id)}">
        <span class="p-name">${esc(p.name)}${p.sample ? ' <span class="tag">Example data</span>' : ''}</span>
        <span class="chev">${ICON.chev}</span>
        <span class="p-meta num">
          <span>${n ? plural(n, 'file') : 'No files yet'}</span>
          ${n ? `<span>${barHTML(s.t)}Titles ${fmt(s.t.done)} / ${fmt(s.t.total)}</span>
          <span>${barHTML(s.a)}Abstracts ${fmt(s.a.done)} / ${fmt(s.a.total)}</span>` : ''}
        </span>
      </button>`;
    }).join('');

    app.innerHTML = `<section class="view"><div class="wrap">
      <div class="brand"><b>Sift</b><span>Swipe to screen</span></div>
      ${projects.length ? `
        <div class="row" style="justify-content:space-between">
          <h2>Projects</h2>
          <button class="btn primary" data-act="new">New project</button>
        </div>
        <div class="projects">${list}</div>
      ` : `
        <div class="empty">
          <h1>Screen records by swiping</h1>
          <p class="note">Create a project for your review, then add your search results as Excel or CSV files. Swipe right to include, left to exclude, up for maybe. Every decision is written to your screening columns as 1, 0 or 2.</p>
          <div class="row">
            <button class="btn primary" data-act="new">New project</button>
            <button class="btn" data-act="sample">Try with example records</button>
          </div>
        </div>
      `}
      <p class="note">${canLiveSave
        ? 'On this browser, decisions are written straight into your Excel files as you swipe.'
        : 'Progress is saved on this device after every swipe. Use <b>Save Excel file</b> to get the updated spreadsheets. For automatic saving into the files themselves, open Sift in Chrome or Edge on a computer.'}</p>
      ${projects.length ? '<div><button class="link" data-act="sample">Add an example project</button></div>' : ''}
    </div></section>`;
  }

  /* ------------------------------------------------------------------ new project */
  function renderNewProject() {
    view = 'new';
    app.innerHTML = `<section class="view"><div class="wrap">
      <div><button class="link" data-act="home">${ICON.back} Projects</button></div>
      <h1>New project</h1>
      <form class="fields" id="np-form" autocomplete="off">
        <div class="field">
          <label for="np-name">Project name</label>
          <input id="np-name" class="text" type="text" maxlength="120" placeholder="e.g. Exercise for knee osteoarthritis" required>
          <span class="hint">You can add several files to a project, such as one export per database.</span>
        </div>
        ${criteriaFieldsHTML('np', { inc: '', exc: '' })}
        <div class="row"><button class="btn primary" id="np-go" type="submit" disabled>Create project</button></div>
      </form>
    </div></section>`;
    const input = document.getElementById('np-name');
    const go = document.getElementById('np-go');
    input.addEventListener('input', () => { go.disabled = !input.value.trim(); });
    document.getElementById('np-form').addEventListener('submit', async e => {
      e.preventDefault();
      const name = input.value.trim();
      if (!name) return;
      const p = await createProject(name, false, readCriteria('np'));
      cur = { p, srcs: [] };
      renderProject();
    });
    setTimeout(() => input.focus(), 50);
  }
  async function createProject(name, sample, criteria) {
    const p = { id: uid('P'), name, fileIds: [], pos: { t: null, a: null }, sample, criteria: criteria || { inc: '', exc: '' },
      showCriteria: false, createdAt: Date.now(), updatedAt: Date.now(), stats: emptyStats() };
    await saveProject(p);
    await loadProjects();
    return p;
  }

  /* ------------------------------------------------------------------ eligibility criteria */
  const hasCriteria = p => !!(p.criteria && (p.criteria.inc.trim() || p.criteria.exc.trim()));
  function criteriaFieldsHTML(prefix, c) {
    return `<div class="field">
        <label for="${prefix}-inc">Inclusion criteria <span class="faint">(optional)</span></label>
        <textarea id="${prefix}-inc" class="text area" rows="5" placeholder="e.g. Adults with knee osteoarthritis&#10;Any exercise-based intervention&#10;Randomised controlled trials">${esc(c.inc)}</textarea>
      </div>
      <div class="field">
        <label for="${prefix}-exc">Exclusion criteria <span class="faint">(optional)</span></label>
        <textarea id="${prefix}-exc" class="text area" rows="4" placeholder="e.g. Post-surgical rehabilitation&#10;Conference abstracts">${esc(c.exc)}</textarea>
        <span class="hint">You can pull these up beside each record while you screen.</span>
      </div>`;
  }
  const readCriteria = prefix => ({
    inc: document.getElementById(prefix + '-inc').value.trim(),
    exc: document.getElementById(prefix + '-exc').value.trim(),
  });
  function criteriaBlocksHTML(c) {
    const block = (cls, title, text) => text ? `<div class="crit-block ${cls}"><h4>${title}</h4><p>${esc(text)}</p></div>` : '';
    return block('inc', 'Include', c.inc) + block('exc', 'Exclude', c.exc);
  }
  function criteriaSectionHTML(p) {
    return `<div class="section" id="crit-section">
      <div class="row" style="justify-content:space-between"><h2>Eligibility criteria</h2>
        <button class="btn" data-act="edit-criteria">${hasCriteria(p) ? 'Edit criteria' : 'Add criteria'}</button></div>
      ${hasCriteria(p) ? `<div class="crit-preview">${criteriaBlocksHTML(p.criteria)}</div>`
        : '<p class="note">No criteria yet. Add your inclusion and exclusion criteria so you can check them while you screen.</p>'}
    </div>`;
  }
  function renderEditCriteria() {
    const slot = document.getElementById('crit-section');
    slot.innerHTML = `<h2>Eligibility criteria</h2>
      <form class="fields" id="cr-form" autocomplete="off">
        ${criteriaFieldsHTML('cr', cur.p.criteria || { inc: '', exc: '' })}
        <div class="row"><button class="btn primary" type="submit">Save criteria</button><button class="btn" type="button" data-act="project">Cancel</button></div>
      </form>`;
    document.getElementById('cr-inc').focus();
    document.getElementById('cr-form').addEventListener('submit', async e => {
      e.preventDefault();
      cur.p.criteria = readCriteria('cr');
      touchProject();
      await loadProjects();
      toast('Criteria saved');
      renderProject();
    });
  }
  function toggleCriteria(force) {
    if (!cur || !screen) return;
    const show = typeof force === 'boolean' ? force : !cur.p.showCriteria;
    cur.p.showCriteria = show;
    saveProject(cur.p);
    const panel = document.getElementById('crit');
    const main = document.getElementById('screen-main');
    const btn = document.getElementById('crit-btn');
    if (panel) panel.hidden = !show;
    if (main) main.classList.toggle('with-crit', show);
    if (btn) btn.setAttribute('aria-pressed', String(show));
  }

  /* ------------------------------------------------------------------ adding files */
  async function pickFile() {
    if (canLiveSave) {
      try {
        const [handle] = await window.showOpenFilePicker({
          multiple: false,
          types: [{ description: 'Spreadsheet', accept: {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
            'application/vnd.ms-excel.sheet.macroEnabled.12': ['.xlsm'],
            'application/vnd.ms-excel': ['.xls'],
            'text/csv': ['.csv'],
          } }],
        });
        const file = await handle.getFile();
        return startImport(file, handle);
      } catch (e) {
        if (e && e.name === 'AbortError') return;
      }
    }
    fileInput.value = '';
    fileInput.click();
  }
  fileInput.addEventListener('change', () => {
    const f = fileInput.files && fileInput.files[0];
    if (f) startImport(f, null);
  });

  async function startImport(file, handle) {
    if (!cur) return;
    const ext = extOf(file.name);
    if (!['xlsx', 'xls', 'csv', 'xlsm'].includes(ext)) { toast('Choose an .xlsx, .xlsm, .xls or .csv file.'); return; }
    let bytes, wb;
    try {
      bytes = await file.arrayBuffer();
      wb = readWorkbook(bytes);
    } catch {
      toast('That file couldn’t be read as a spreadsheet.', 4000);
      return;
    }
    if (!wb.SheetNames.length) { toast('That file has no sheets.'); return; }
    renderMap({ fileName: file.name, ext, bytes, handle, wb, sheet: wb.SheetNames.length > 1 ? null : wb.SheetNames[0], existing: null });
  }

  /* ------------------------------------------------------------------ column mapping */
  function guess(headers, test, avoid) {
    const h = headers.find(x => test(x.name) && !(avoid && avoid(x.name)));
    return h ? h.c : null;
  }
  const isDecisionName = n => /screen|decision|include|exclu|incl|verdict/i.test(n);

  function renderMap(draft) {
    view = 'map';
    const ex = draft.existing;
    const multiSheet = draft.wb.SheetNames.length > 1;
    const ws = draft.sheet ? draft.wb.Sheets[draft.sheet] : null; // null until the user picks a sheet
    const headers = ws ? headersOf(ws) : [];
    const prev = ex && ex.sheet === draft.sheet ? ex.cols : null;
    // Reuse the column choices of another file in this project with the same headers.
    const sibling = ws && !ex && cur.srcs.find(src => {
      const h = headersOf(src.wb.Sheets[src.s.sheet]).map(x => x.name).join('|');
      return h === headers.map(x => x.name).join('|');
    });
    const base = prev || (sibling ? sibling.s.cols : null);

    const sel = {
      title: base ? base.title : (guess(headers, n => /^\s*(article\s+|primary\s+)?title\s*$/i.test(n)) ?? guess(headers, n => /title/i.test(n), isDecisionName)),
      abs: base ? base.abs : (guess(headers, n => /^\s*abstract\s*$/i.test(n)) ?? guess(headers, n => /abstract|summary/i.test(n), isDecisionName)),
      titleDec: base ? base.titleDec : (guess(headers, n => /title/i.test(n) && isDecisionName(n)) ?? 'new'),
      absDec: base ? base.absDec : (guess(headers, n => /abstract/i.test(n) && isDecisionName(n)) ?? 'new'),
    };
    const nh = (ex && ex.newHeaders) || (sibling && sibling.s.newHeaders) || {};
    if (base) {
      if (nh[base.titleDec] && !headers.some(h => h.c === base.titleDec && h.name)) sel.titleDec = 'new';
      if (nh[base.absDec] && !headers.some(h => h.c === base.absDec && h.name)) sel.absDec = 'new';
    }

    const opts = (val, allowNew, newLabel) =>
      (allowNew ? `<option value="new"${val === 'new' ? ' selected' : ''}>+ New column “${esc(newLabel)}”</option>` : `<option value=""${val == null ? ' selected' : ''} disabled>Choose a column</option>`) +
      headers.map(h => `<option value="${h.c}"${String(val) === String(h.c) ? ' selected' : ''}>${esc(h.label)}</option>`).join('');

    app.innerHTML = `<section class="view"><div class="wrap">
      <div><button class="link" data-act="project">${ICON.back} ${esc(cur.p.name)}</button></div>
      <div style="display:grid;gap:6px">
        <h1>${ex ? 'Change columns' : 'Match your columns'}</h1>
        <p class="muted small">${esc(draft.fileName)}</p>
      </div>
      ${multiSheet ? `<div class="field"><label for="m-sheet">Sheet</label>
        <select id="m-sheet">${ws ? '' : '<option value="" selected disabled>Choose a sheet</option>'}${draft.wb.SheetNames.map(n => `<option${n === draft.sheet ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select>
        ${ws ? '' : '<span class="hint">This file has more than one sheet. Pick the one with your records.</span>'}</div>` : ''}
      <div class="ai-row">
        <button class="btn magic" id="m-ai" type="button"${ws ? '' : ' disabled'}>${ICON.magic}<span>Identify the columns for me</span></button>
        ${ws ? '' : '<p class="hint small faint">Choose a sheet first.</p>'}
        <p class="ai-note" id="m-ai-note" hidden></p>
      </div>
      ${ws ? `<div class="fields">
        <div class="pair">
          <div class="field"><label for="m-title">Title column</label><select id="m-title">${opts(sel.title, false)}</select></div>
          <div class="field"><label for="m-tdec">Title screening decision column</label><select id="m-tdec">${opts(sel.titleDec, true, DEFAULT_HEADER.t)}</select></div>
        </div>
        <div class="pair">
          <div class="field"><label for="m-abs">Abstract column</label><select id="m-abs">${opts(sel.abs, false)}</select></div>
          <div class="field"><label for="m-adec">Abstract screening decision column</label><select id="m-adec">${opts(sel.absDec, true, DEFAULT_HEADER.a)}</select></div>
        </div>
        <p class="hint small faint">Decisions are written as 1 = include, 0 = exclude, 2 = maybe. Values already in a decision column are kept and count as screened.</p>
      </div>
      <div class="preview" id="m-preview"></div>` : ''}
      <p class="error" id="m-err" hidden></p>
      <div class="row"><button class="btn primary" id="m-go"${ws ? '' : ' disabled'}>${ex ? 'Save columns' : 'Add file to project'}</button></div>
    </div></section>`;

    const $ = id => document.getElementById(id);
    if ($('m-sheet')) $('m-sheet').addEventListener('change', e => renderMap({ ...draft, sheet: e.target.value }));
    if (!ws) return; // nothing else to set up until a sheet is chosen

    const read = () => ({
      title: $('m-title').value === '' ? null : +$('m-title').value,
      abs: $('m-abs').value === '' ? null : +$('m-abs').value,
      titleDec: $('m-tdec').value === 'new' ? 'new' : +$('m-tdec').value,
      absDec: $('m-adec').value === 'new' ? 'new' : +$('m-adec').value,
    });
    const preview = () => {
      const s = read();
      const rg = rangeOf(ws);
      let r = rg.s.r + 1;
      for (; r <= rg.e.r; r++) if (cellText(ws, r, s.title) || cellText(ws, r, s.abs)) break;
      const t = cellText(ws, r, s.title), a = cellText(ws, r, s.abs);
      $('m-preview').innerHTML = `<h2>First record · ${fmt(rg.e.r - rg.s.r)} rows</h2>
        <div class="t">${t ? esc(t) : '<span class="faint">No title in this column</span>'}</div>
        <div class="a">${a ? esc(a) : '<span class="faint">No abstract in this column</span>'}</div>`;
    };
    preview();
    ['m-title', 'm-abs', 'm-tdec', 'm-adec'].forEach(id => $(id).addEventListener('change', preview));

    $('m-ai').addEventListener('click', async () => {
      const btn = $('m-ai'), note = $('m-ai-note');
      btn.disabled = true;
      btn.classList.add('busy');
      btn.querySelector('span').textContent = 'Reading your file…';
      note.hidden = true;
      $('m-err').hidden = true;
      try {
        const pick = await identifyColumns(ws, headers);
        const setSel = (id, pos, allowNew) => {
          const el = $(id);
          el.value = pos >= 0 && headers[pos] ? String(headers[pos].c) : (allowNew ? 'new' : '');
          el.closest('.field').classList.add('ai-picked');
        };
        setSel('m-title', pick.title, false);
        setSel('m-abs', pick.abstract, false);
        setSel('m-tdec', pick.title_decision, true);
        setSel('m-adec', pick.abstract_decision, true);
        preview();
        const missing = [pick.title < 0 && 'title', pick.abstract < 0 && 'abstract'].filter(Boolean);
        note.textContent = (missing.length
          ? `Claude couldn’t find a ${missing.join(' or ')} column. Pick it yourself, then tap Looks right.`
          : 'Claude picked the columns outlined in purple. Check them, then tap Looks right.') + (pick.note ? ' ' + pick.note : '');
        note.hidden = false;
        $('m-go').textContent = 'Looks right';
        btn.querySelector('span').textContent = 'Identify again';
      } catch (e) {
        $('m-err').textContent = e.message || 'The column finder didn’t work. Pick the columns yourself.';
        $('m-err').hidden = false;
        btn.querySelector('span').textContent = 'Identify the columns for me';
      }
      btn.disabled = false;
      btn.classList.remove('busy');
    });

    $('m-go').addEventListener('click', async () => {
      const s = read();
      const err = msg => { $('m-err').textContent = msg; $('m-err').hidden = false; };
      if (s.title == null) return err('Choose the column that holds titles.');
      if (s.abs == null) return err('Choose the column that holds abstracts.');
      if ([s.titleDec, s.absDec].some(v => v === s.title || v === s.abs)) return err('A decision column can’t be the same as the title or abstract column.');
      if (s.titleDec !== 'new' && s.titleDec === s.absDec) return err('Use different columns for title and abstract decisions.');
      const rg = rangeOf(ws);
      let next = rg.e.c + 1;
      const newHeaders = {};
      if (s.titleDec === 'new') { s.titleDec = next++; newHeaders[s.titleDec] = DEFAULT_HEADER.t; }
      if (s.absDec === 'new') { s.absDec = next++; newHeaders[s.absDec] = DEFAULT_HEADER.a; }
      await finishMapping(draft, s, newHeaders);
    });
  }

  // Send the chosen sheet's headers and a few sample values per column to the column finder.
  async function identifyColumns(ws, headers) {
    const rg = rangeOf(ws);
    const columns = headers.slice(0, 80).map(h => {
      const samples = [];
      for (let r = rg.s.r + 1; r <= rg.e.r && samples.length < 4; r++) {
        const v = cellText(ws, r, h.c);
        if (v) samples.push(v.length > 240 ? v.slice(0, 240) + '…' : v);
      }
      return { name: h.name, samples };
    });
    let res;
    try {
      res = await fetch(AI_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + AI_ANON, apikey: AI_ANON },
        body: JSON.stringify({ columns }),
      });
    } catch {
      throw new Error('Couldn’t reach the column finder. Check your connection, or pick the columns yourself.');
    }
    let body = null;
    try { body = await res.json(); } catch {}
    if (!res.ok || !body || body.error) throw new Error((body && body.error) || 'The column finder didn’t work. Pick the columns yourself.');
    return body;
  }

  async function finishMapping(draft, cols, newHeaders) {
    const ws = draft.wb.Sheets[draft.sheet];
    const rows = readRows(ws, cols);
    if (!rows.length) { toast('No records found in those columns.'); return; }
    const ex = draft.existing;
    const s = ex ? { ...ex } : {
      id: uid('S'), projectId: cur.p.id, fileName: draft.fileName, ext: draft.ext,
      addedAt: Date.now(), decisions: { t: {}, a: {} }, unsaved: 0,
    };
    if (ex && ex.sheet !== draft.sheet) {
      s.decisions = { t: {}, a: {} };
      for (const k of ['t', 'a']) if (cur.p.pos[k] && cur.p.pos[k].sid === s.id) cur.p.pos[k] = null;
    }
    s.sheet = draft.sheet;
    s.headerRow = rangeOf(ws).s.r;
    s.cols = cols;
    s.newHeaders = newHeaders;
    s.fingerprint = fingerprint(draft.sheet, rows);
    mergeFileDecisions(s, rows);
    s.stats = sourceStats(s, rows);
    s.updatedAt = Date.now();

    await saveSource(s);
    if (!ex) {
      const bytes = draft.bytes || XL.write(draft.wb, { bookType: 'xlsx', type: 'array' });
      await db.put('files', { id: s.id, bytes }).catch(() => {});
      cur.p.fileIds.push(s.id);
      const live = !!draft.handle && LIVE_TYPES.includes(s.ext);
      if (live) rememberHandle(s.id, draft.handle);
      cur.srcs.push({ s, wb: draft.wb, rows, bytes, handle: live ? draft.handle : null, live, reconnect: false });
    } else {
      const src = cur.srcs.find(x => x.s.id === s.id);
      Object.assign(src, { s, wb: draft.wb, rows });
    }
    refreshStats();
    touchProject();
    await loadProjects();
    const src = cur.srcs.find(x => x.s.id === s.id);
    if (src.live) queueLiveSave(src);
    renderProject();
    if (!ex) toast(`Added ${s.fileName}`);
  }

  /* ------------------------------------------------------------------ live-file handles */
  function rememberHandle(id, handle) {
    if (!handle) return;
    handleCache.set(id, handle);
    db.put('handles', { id, handle }).catch(() => {});
  }
  async function readHandle(id) {
    if (handleCache.has(id)) return handleCache.get(id);
    if (skipStoredHandles) return null;
    // If reading the handle ends the page, this marker is still here on the next load and we skip stored handles.
    try { sessionStorage.setItem('sift-reading-handle', id); } catch {}
    const rec = await db.get('handles', id).catch(() => null);
    try { sessionStorage.removeItem('sift-reading-handle'); } catch {}
    if (rec && rec.handle) handleCache.set(id, rec.handle);
    return rec ? rec.handle : null;
  }

  /* ------------------------------------------------------------------ open project */
  async function loadSource(id, askPermission) {
    const s = await db.get('sources', id).catch(() => null);
    const rec = await db.get('files', id).catch(() => null);
    if (!s || !rec) return null;
    let bytes = rec.bytes, handle = null, live = false, reconnect = false;
    if (canLiveSave && LIVE_TYPES.includes(s.ext)) handle = await readHandle(id); // old .xls files: Save Excel file instead
    if (handle && canLiveSave) {
      try {
        let perm = await handle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted' && askPermission) perm = await handle.requestPermission({ mode: 'readwrite' });
        if (perm === 'granted') {
          const file = await handle.getFile();
          bytes = await file.arrayBuffer();
          live = true;
          db.put('files', { id, bytes }).catch(() => {});
        } else reconnect = true;
      } catch { reconnect = true; }
    }
    let wb;
    try { wb = readWorkbook(bytes); } catch { return null; }
    if (!wb.Sheets[s.sheet]) return { s, wb, bytes, rows: [], handle, live, reconnect, broken: true };
    if (s.headerRow == null) s.headerRow = rangeOf(wb.Sheets[s.sheet]).s.r;
    const rows = readRows(wb.Sheets[s.sheet], s.cols);
    const fileBehind = mergeFileDecisions(s, rows);
    return { s, wb, bytes, rows, handle, live, reconnect, fileBehind };
  }

  async function openProject(id) {
    const p = projects.find(x => x.id === id);
    if (!p) return renderHome();
    const srcs = [];
    let ask = true;
    for (const fid of p.fileIds || []) {
      const src = await loadSource(fid, ask);
      if (src) {
        srcs.push(src);
        if (src.reconnect) ask = false; // one permission prompt per click; others get a Reconnect button
      }
    }
    p.fileIds = srcs.map(x => x.s.id);
    cur = { p, srcs };
    // e.g. the page closed moments after a swipe: bring the Excel file up to date now
    for (const src of srcs) if (src.live && src.fileBehind) queueLiveSave(src);
    refreshStats();
    saveProject(p);
    renderProject();
    if (srcs.some(x => x.broken)) toast('A sheet is missing from one of the files. Use Change columns on that file.', 4500);
  }

  /* ------------------------------------------------------------------ project view */
  function renderProject() {
    view = 'project'; screen = null;
    const { p, srcs } = cur;
    const s = refreshStats();
    const multi = srcs.length > 1;
    const stageCard = (k) => {
      const st = s[k];
      const left = st.total - st.done;
      const locked = k === 'a' && st.total === 0;
      const noun = k === 't' ? 'title' : 'abstract';
      const label = left > 0 ? `Screen ${fmt(left)} remaining ${noun}${left === 1 ? '' : 's'}` : `All ${noun}s screened · review`;
      return `<div class="stage">
        <h3 class="stage-title">${STAGE_NAME[k]}</h3>
        <div class="progress-block">
          <div class="prog-head num"><span class="big">${st.total ? Math.round(100 * st.done / st.total) : 0}%<small> screened</small></span>
            <span class="muted small">${fmt(st.done)} of ${fmt(st.total)} · ${fmt(left)} left</span></div>
          ${meterHTML(st.done, st.total)}
        </div>
        ${splitHTML(st)}
        <p class="small muted">${k === 'a' ? 'Records included or marked maybe at title stage.' : multi ? `Every record across ${srcs.length} files.` : 'Every record in the file.'}</p>
        <button class="btn ${locked ? '' : 'primary'}" data-stage="${k}" ${locked ? 'disabled' : ''}>${locked ? 'Include some titles first' : label}</button>
      </div>`;
    };
    const fileRow = (src) => {
      const f = src.s, st = f.stats || emptyStats();
      const status = src.broken ? '<span class="tag exclude">Sheet missing</span>'
        : src.reconnect ? `<button class="btn small-btn" data-reconnect="${esc(f.id)}">Reconnect file</button>`
        : src.live ? '<span class="tag include">Live file</span>'
        : f.unsaved ? `<span class="tag maybe">${fmt(f.unsaved)} not in Excel yet</span>` : '';
      return `<div class="file" data-file="${esc(f.id)}">
        <div class="f-head">
          <span class="f-icon">${ICON.file}</span>
          <div class="f-name"><b>${esc(f.fileName)}</b><span class="muted small num">${esc(f.sheet)} · ${plural(src.rows.length, 'record')}</span></div>
          ${status}
        </div>
        <div class="p-meta num">
          <span>${barHTML(st.t)}Titles ${fmt(st.t.done)} / ${fmt(st.t.total)}</span>
          <span>${barHTML(st.a)}Abstracts ${fmt(st.a.done)} / ${fmt(st.a.total)}</span>
        </div>
        <div class="f-actions" data-slot="${esc(f.id)}">
          ${src.live ? '' : `<button class="link small" data-save-one="${esc(f.id)}">Save Excel file</button>`}
          <button class="link small" data-columns="${esc(f.id)}">Change columns</button>
          <button class="link small danger-link" data-remove="${esc(f.id)}">Remove</button>
        </div>
      </div>`;
    };

    app.innerHTML = `<section class="view"><div class="wrap">
      <div><button class="link" data-act="home">${ICON.back} Projects</button></div>
      <div style="display:grid;gap:6px" id="title-slot">
        <h1>${esc(p.name)}</h1>
        <p class="muted small num">${srcs.length ? `${plural(srcs.length, 'file')} · ${plural(s.t.total, 'record')}` : 'No files yet'}${p.sample ? ' · <span class="tag">Example data</span>' : ''}</p>
      </div>
      ${srcs.length ? `<div class="stages">${stageCard('t')}${stageCard('a')}</div>` : `
        <div class="empty">
          <h3 class="stage-title" style="border:0;padding:0">Add your first file</h3>
          <p class="note">Add an Excel or CSV export of your search results. You can add more files later, for example one per database, and screen them together.</p>
          <button class="btn primary" data-act="add">Add Excel file</button>
        </div>`}
      ${criteriaSectionHTML(p)}
      ${srcs.length ? `
        <div class="section">
          <div class="row" style="justify-content:space-between"><h2>Files</h2><button class="btn" data-act="add">Add file</button></div>
          <div class="files">${srcs.map(fileRow).join('')}</div>
        </div>` : ''}
      <div class="row">
        ${srcs.length ? saveButtonHTML('btn') : ''}
        <button class="btn" data-act="rename">Rename</button>
        <span id="del-slot"><button class="btn danger" data-act="delete">Delete project</button></span>
      </div>
      ${srcs.length ? `<p class="note">${srcs.every(x => x.live)
        ? 'Live files: every decision is written into the Excel files on this computer a moment after you swipe.'
        : 'Progress is saved on this device after every swipe. <b>Save Excel file</b> gives you the spreadsheets with all decisions filled in.'}</p>` : ''}
    </div></section>`;
  }

  function renderRename() {
    const slot = document.getElementById('title-slot');
    slot.innerHTML = `<form class="row" id="rn-form" autocomplete="off">
      <input id="rn-name" class="text" type="text" maxlength="120" value="${esc(cur.p.name)}" aria-label="Project name" style="flex:1;min-width:200px">
      <button class="btn primary" type="submit">Save</button>
      <button class="btn" type="button" data-act="project">Cancel</button>
    </form>`;
    const input = document.getElementById('rn-name');
    input.focus(); input.select();
    document.getElementById('rn-form').addEventListener('submit', async e => {
      e.preventDefault();
      const name = input.value.trim();
      if (!name) return;
      cur.p.name = name;
      touchProject();
      await loadProjects();
      renderProject();
    });
  }

  /* ------------------------------------------------------------------ saving to Excel */
  const nonLive = () => cur.srcs.filter(x => !x.live && !x.broken);
  function saveButtonHTML(cls) {
    if (!cur.srcs.length) return '';
    const isBtn = cls === 'btn';
    const pending = nonLive();
    if (!pending.length) {
      const st = liveState;
      const text = st === 'saving' ? 'Saving to Excel file…' : st === 'error' ? 'Not saved · retry' : 'Changes saved to Excel file';
      return `<button class="${isBtn ? 'btn' : 'save'} ${st === 'error' ? 'err' : st === 'saving' ? 'pending' : 'ok'}" data-act="live-retry" id="save-btn">${text}</button>`;
    }
    const n = pending.reduce((a, x) => a + (x.s.unsaved || 0), 0);
    const label = pending.length > 1 ? 'Save Excel files' : 'Save Excel file';
    if (isBtn) return `<button class="btn" data-act="export" id="save-btn">${label}${n ? ` · ${fmt(n)} new` : ''}</button>`;
    return `<button class="save ${n ? 'pending' : 'ok'}" data-act="export" id="save-btn">${label}${n ? ` · ${fmt(n)}` : ''}</button>`;
  }
  function refreshSaveButton() {
    const b = document.getElementById('save-btn');
    if (!b || !cur) return;
    const tmp = document.createElement('div');
    tmp.innerHTML = saveButtonHTML(b.classList.contains('btn') ? 'btn' : 'save');
    if (tmp.firstElementChild) b.replaceWith(tmp.firstElementChild);
  }

  let liveTimer = null, liveState = 'ok', liveBusy = false;
  const liveDirty = new Set();
  function queueLiveSave(src) {
    if (!src || !src.live) return;
    liveDirty.add(src);
    if (liveState !== 'error') { liveState = 'saving'; refreshSaveButton(); } // don't claim "saved" while a write is pending
    clearTimeout(liveTimer);
    liveTimer = setTimeout(liveSave, 400);
  }
  async function liveSave() {
    if (liveBusy) { clearTimeout(liveTimer); liveTimer = setTimeout(liveSave, 400); return; }
    if (!liveDirty.size) { refreshSaveButton(); return; }
    liveBusy = true; liveState = 'saving'; refreshSaveButton();
    const batch = [...liveDirty]; liveDirty.clear();
    let failed = false;
    for (const src of batch) {
      try {
        const { blob } = outputFile(src);
        const w = await src.handle.createWritable();
        await w.write(blob);
        await w.close();
        src.s.unsaved = 0;
        saveSource(src.s);
        blob.arrayBuffer().then(bytes => db.put('files', { id: src.s.id, bytes })).catch(() => {});
      } catch (e) {
        failed = true;
        liveDirty.add(src);
        toast(e && e.name === 'NoModificationAllowedError'
          ? `${src.s.fileName} is locked by another app. Close it there and tap “retry”.`
          : `Couldn’t write to ${src.s.fileName}. Your progress is still saved here.`, 5000);
      }
    }
    liveBusy = false;
    liveState = failed ? 'error' : liveDirty.size ? 'saving' : 'ok'; // more swipes may have arrived during the write
    refreshSaveButton();
  }
  function retryLive() {
    if (!cur) return;
    cur.srcs.filter(x => x.live).forEach(x => liveDirty.add(x));
    liveSave();
  }

  async function exportFiles(targets) {
    if (!targets.length) return;
    let files;
    try { files = targets.map(src => ({ src, ...outputFile(src) })); } catch { toast('Couldn’t build the spreadsheet.'); return; }
    const done = () => {
      for (const f of files) { f.src.s.unsaved = 0; f.src.s.exportedAt = Date.now(); saveSource(f.src.s); }
      refreshSaveButton();
      if (view === 'project') renderProject();
    };
    // Phones: the share sheet lets you "Save to Files" / Drive / email, all files at once.
    if (isTouch && navigator.canShare) {
      const list = files.map(f => new File([f.blob], f.name, { type: f.blob.type }));
      if (navigator.canShare({ files: list })) {
        try { await navigator.share({ files: list }); done(); toast('Saved'); return; }
        catch (e) { if (e && e.name === 'AbortError') return; }
      }
    }
    files.forEach((f, i) => setTimeout(() => {
      const url = URL.createObjectURL(f.blob);
      const a = document.createElement('a');
      a.href = url; a.download = f.name;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
    }, i * 400));
    done();
    toast(files.length > 1 ? `Downloaded ${files.length} files` : `Downloaded ${files[0].name}`);
  }
  function exportAll() {
    const pending = nonLive();
    const changed = pending.filter(x => x.s.unsaved);
    return exportFiles(changed.length ? changed : pending);
  }

  /* ------------------------------------------------------------------ screening */
  const decOf = (stage, item) => cur.srcs[item.k].s.decisions[stage][item.x.r];
  function buildQueue(stage) {
    const q = [];
    cur.srcs.forEach((src, k) => {
      if (src.broken) return;
      const d = src.s.decisions;
      for (const x of src.rows) if (stage === 't' || d.t[x.r] === INCLUDE || d.t[x.r] === MAYBE) q.push({ k, x });
    });
    return q;
  }
  function nextUndecided(from) {
    const { queue, stage } = screen;
    for (let i = from; i < queue.length; i++) if (decOf(stage, queue[i]) === undefined) return i;
    for (let i = 0; i < Math.min(from, queue.length); i++) if (decOf(stage, queue[i]) === undefined) return i;
    return queue.length;
  }
  function posOf(i) {
    const { queue } = screen;
    if (i >= queue.length) return -1;
    return { sid: cur.srcs[queue[i].k].s.id, r: queue[i].x.r };
  }
  function startStage(stage) {
    const queue = buildQueue(stage);
    screen = { stage, queue, qi: 0, busy: false, review: false };
    const pos = cur.p.pos[stage];
    let qi = !pos || pos === -1 ? -1 : queue.findIndex(it => cur.srcs[it.k].s.id === pos.sid && it.x.r === pos.r);
    if (qi < 0) qi = nextUndecided(0);
    screen.qi = qi;
    renderScreen();
  }

  function renderScreen() {
    view = 'screen';
    const { stage, queue } = screen;
    app.innerHTML = `<section class="screen">
      <div class="topbar">
        <button class="link" data-act="project" aria-label="Back to project">${ICON.back}<span>Project</span></button>
        <div class="mid"><div>${STAGE_NAME[stage]}</div><span class="num" id="s-count"></span></div>
        <span id="s-save">${saveButtonHTML('save')}</span>
      </div>
      <div class="progress" id="s-bar"></div>
      <div class="screen-main${cur.p.showCriteria ? ' with-crit' : ''}" id="screen-main">
        <div class="deck" id="deck"></div>
        <aside class="crit" id="crit" aria-label="Eligibility criteria"${cur.p.showCriteria ? '' : ' hidden'}>
          <div class="crit-head"><h3>Eligibility criteria</h3>
            <button class="crit-x" data-act="criteria-close" aria-label="Hide criteria">${ICON.close}</button></div>
          <div class="crit-body" tabindex="0">${hasCriteria(cur.p) ? criteriaBlocksHTML(cur.p.criteria)
            : '<p class="muted small">No criteria for this project yet.</p><button class="link small" data-act="criteria-add">Add criteria</button>'}</div>
        </aside>
      </div>
      <div class="controls" id="controls">
        <button class="circle back" data-act="undo" aria-label="Previous record">${ICON.undo}</button>
        <button class="circle no" data-dec="0" aria-label="Exclude">${ICON.x}</button>
        <button class="circle may" data-dec="2" aria-label="Maybe">${ICON.q}</button>
        <button class="circle yes" data-dec="1" aria-label="Include">${ICON.check}</button>
        <button class="circle back crit-toggle" id="crit-btn" data-act="criteria" aria-label="Show eligibility criteria" aria-pressed="${!!cur.p.showCriteria}" title="Criteria (C)">${ICON.criteria}</button>
      </div>
      <div class="hints" id="hints">${isTouch
        ? 'Swipe right to include · left to exclude · up for maybe'
        : '→ include · ← exclude · ↑ maybe · Backspace for previous · C for criteria'}</div>
    </section>`;
    if (!queue.length) {
      document.getElementById('deck').innerHTML = `<div class="done"><div><h1>No records here yet</h1><p class="muted">Include titles first, then come back for abstracts.</p><button class="btn" data-act="project">Back to project</button></div></div>`;
      document.getElementById('controls').hidden = true;
      document.getElementById('hints').hidden = true;
      return;
    }
    renderDeck(false);
  }

  function updateHeader() {
    const { stage, queue, qi } = screen;
    const st = refreshStats()[stage];
    document.getElementById('s-count').textContent = qi < queue.length
      ? `${screen.review ? 'Reviewing' : 'Record'} ${fmt(qi + 1)} of ${fmt(queue.length)} · ${fmt(st.done)} screened`
      : `${fmt(st.done)} of ${fmt(queue.length)} screened`;
    document.getElementById('s-bar').innerHTML = meterHTML(st.done, st.total);
  }

  function cardHTML(item, idx, under) {
    const { stage, queue } = screen;
    const x = item.x;
    const v = decOf(stage, item);
    const isAbs = stage === 'a';
    const chip = v !== undefined ? `<span class="tag ${CLS[v]}">${LABEL[v]}</span>` : '';
    const where = cur.srcs.length > 1 ? `${esc(cur.srcs[item.k].s.fileName)} · row ${fmt(x.r + 1)}` : `Excel row ${fmt(x.r + 1)}`;
    return `<article class="card ${isAbs ? 'abs' : ''} ${under ? 'under' : ''}" ${under ? 'aria-hidden="true"' : ''}>
      <div class="meta num"><span class="where">${fmt(idx + 1)} / ${fmt(queue.length)} · ${where}</span>${chip}</div>
      <div class="body">
        <div class="inner ${isAbs ? 'top' : ''}">
          <div class="ttl">${x.t ? esc(x.t) : '<span class="faint">No title</span>'}</div>
          ${isAbs ? (x.a ? `<div class="abstract">${esc(x.a)}</div>` : '<div class="abstract none">No abstract in this record.</div>') : ''}
        </div>
      </div>
      ${under ? '' : `<span class="stamp s-in">INCLUDE</span><span class="stamp s-out">EXCLUDE</span><span class="stamp s-may">MAYBE</span>`}
      ${under || (v === undefined && !screen.review) ? '' : `<div class="row" style="justify-content:center"><button class="link small" data-act="resume">${screen.review ? 'Skip to next record' : 'Skip to next unscreened record'}</button></div>`}
    </article>`;
  }

  function renderDeck(animateIn) {
    const deck = document.getElementById('deck');
    const { queue, qi, stage } = screen;
    updateHeader();
    const controls = document.getElementById('controls');
    const hints = document.getElementById('hints');
    if (qi >= queue.length) {
      const st = cur.p.stats[stage];
      const left = queue.length - st.done;
      const allLive = !nonLive().length;
      screen.review = false;
      deck.innerHTML = `<div class="done"><div>
        <h1>${left ? `${fmt(left)} records still unscreened` : `All ${fmt(queue.length)} records screened`}</h1>
        ${splitHTML(st)}
        <div class="row" style="justify-content:center">
          ${allLive ? '' : `<button class="btn primary" data-act="export">${nonLive().length > 1 ? 'Save Excel files' : 'Save Excel file'}</button>`}
          ${stage === 't' && st[1] + st[2] > 0 ? `<button class="btn${allLive ? ' primary' : ''}" data-stage="a">Start abstract screening</button>` : ''}
          <button class="btn" data-act="review">Review from the start</button>
        </div>
      </div></div>`;
      controls.querySelectorAll('[data-dec]').forEach(b => { b.disabled = true; });
      hints.hidden = true;
      cur.p.pos[stage] = -1;
      saveProject(cur.p);
      return;
    }
    controls.querySelectorAll('[data-dec]').forEach(b => { b.disabled = false; });
    hints.hidden = false;
    let nextIdx = -1;
    if (screen.review) nextIdx = qi + 1 < queue.length ? qi + 1 : -1;
    else for (let i = qi + 1; i < queue.length; i++) if (decOf(stage, queue[i]) === undefined) { nextIdx = i; break; }
    deck.innerHTML = (nextIdx >= 0 ? cardHTML(queue[nextIdx], nextIdx, true) : '') + cardHTML(queue[qi], qi, false);
    const card = deck.querySelector('.card:not(.under)');
    if (animateIn) {
      card.classList.add('enter', 'dragging');
      requestAnimationFrame(() => { card.classList.remove('dragging'); requestAnimationFrame(() => card.classList.remove('enter')); });
    }
    attachGestures(card);
    cur.p.pos[stage] = posOf(qi);
  }

  function decide(v) {
    const { stage, queue, qi } = screen;
    const item = queue[qi];
    const src = cur.srcs[item.k];
    const prev = src.s.decisions[stage][item.x.r];
    src.s.decisions[stage][item.x.r] = v;
    if (prev !== v) src.s.unsaved = (src.s.unsaved || 0) + 1;
    src.s.updatedAt = Date.now();
    // Reviewing walks through every record in order; normal screening jumps to the next unscreened one.
    screen.qi = screen.review ? qi + 1 : nextUndecided(qi + 1);
    cur.p.pos[stage] = posOf(screen.qi);
    refreshStats();
    saveSource(src.s);
    touchProject();
    if (src.live) queueLiveSave(src);
    renderDeck(true);
    const s = document.getElementById('s-save');
    if (s) s.innerHTML = saveButtonHTML('save');
  }

  function flyOut(v, card) {
    if (!screen || screen.busy || screen.qi >= screen.queue.length) return;
    card = card || document.querySelector('#deck .card:not(.under)');
    if (!card) return;
    screen.busy = true;
    const W = window.innerWidth, H = window.innerHeight;
    const m = /translate\(([-\d.]+)px,\s*([-\d.]+)px\)/.exec(card.style.transform || '');
    const y0 = m ? +m[2] : 0;
    card.classList.remove('dragging');
    const show = cls => { const s = card.querySelector(cls); if (s) s.style.opacity = 1; };
    if (v === INCLUDE) { show('.s-in'); card.style.transform = `translate(${W + 200}px, ${y0}px) rotate(18deg)`; }
    else if (v === EXCLUDE) { show('.s-out'); card.style.transform = `translate(${-W - 200}px, ${y0}px) rotate(-18deg)`; }
    else { show('.s-may'); card.style.transform = `translate(0px, ${-H - 200}px)`; }
    card.style.opacity = '0';
    const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    setTimeout(() => { screen.busy = false; decide(v); }, reduce ? 20 : 240);
  }

  function goBack() {
    if (!screen || screen.busy) return;
    if (screen.qi <= 0 && screen.qi < screen.queue.length) { toast('This is the first record.'); return; }
    screen.qi = Math.max(0, Math.min(screen.qi, screen.queue.length) - 1);
    renderDeck(false);
    saveProject(cur.p);
  }

  /* ------------------------------------------------------------------ gestures */
  let momentum = null;
  function attachGestures(card) {
    const body = card.querySelector('.body');
    const sIn = card.querySelector('.s-in'), sOut = card.querySelector('.s-out'), sMay = card.querySelector('.s-may');
    let pid = null, sx = 0, sy = 0, ly = 0, lt = 0, vy = 0, vx = 0, lx = 0, mode = null, dx = 0, dy = 0;

    const thrX = () => Math.min(130, card.clientWidth * 0.26);
    const thrY = () => Math.min(140, card.clientHeight * 0.2);
    const reset = () => {
      card.classList.remove('dragging');
      card.style.transform = '';
      card.style.borderColor = '';
      [sIn, sOut, sMay].forEach(s => { if (s) s.style.opacity = 0; });
    };

    // iOS Safari: stop the page from claiming the gesture.
    card.addEventListener('touchmove', e => { if (e.cancelable) e.preventDefault(); }, { passive: false });

    card.addEventListener('pointerdown', e => {
      if (screen.busy || (e.pointerType === 'mouse' && e.button !== 0)) return;
      if (e.target.closest('button, a')) return;
      if (momentum) { cancelAnimationFrame(momentum); momentum = null; }
      pid = e.pointerId; sx = lx = e.clientX; sy = ly = e.clientY; lt = performance.now();
      mode = null; dx = dy = vx = vy = 0;
      try { card.setPointerCapture(pid); } catch {}
    });

    card.addEventListener('pointermove', e => {
      if (e.pointerId !== pid) return;
      const mx = e.clientX - sx, my = e.clientY - sy;
      const now = performance.now(), dt = Math.max(1, now - lt);
      vx = 0.8 * ((e.clientX - lx) / dt) + 0.2 * vx;
      vy = 0.8 * ((e.clientY - ly) / dt) + 0.2 * vy;
      if (!mode) {
        if (Math.hypot(mx, my) < 8) return;
        if (Math.abs(mx) > Math.abs(my)) mode = 'card';
        else {
          const scrollable = body.scrollHeight > body.clientHeight + 2;
          const atBottom = body.scrollTop + body.clientHeight >= body.scrollHeight - 2;
          mode = (!scrollable || (my < 0 && atBottom)) ? 'card' : 'scroll';
        }
        if (mode === 'card') card.classList.add('dragging');
      }
      if (mode === 'scroll') {
        body.scrollTop -= (e.clientY - ly);
      } else {
        dx = mx; dy = my;
        card.style.transform = `translate(${dx}px, ${dy}px) rotate(${dx / 20}deg)`;
        const vertical = -dy > Math.abs(dx);
        const pIn = vertical ? 0 : Math.max(0, Math.min(1, dx / thrX()));
        const pOut = vertical ? 0 : Math.max(0, Math.min(1, -dx / thrX()));
        const pMay = vertical ? Math.max(0, Math.min(1, -dy / thrY())) : 0;
        if (sIn) sIn.style.opacity = pIn;
        if (sOut) sOut.style.opacity = pOut;
        if (sMay) sMay.style.opacity = pMay;
        card.style.borderColor = pIn > .6 ? 'var(--include)' : pOut > .6 ? 'var(--exclude)' : pMay > .6 ? 'var(--maybe)' : '';
      }
      lx = e.clientX; ly = e.clientY; lt = now;
    });

    const end = e => {
      if (e.pointerId !== pid) return;
      pid = null;
      if (mode === 'scroll') {
        let v = vy, last = performance.now();
        const step = t => {
          const dt = t - last; last = t;
          body.scrollTop -= v * dt;
          v *= Math.pow(0.94, dt / 16);
          momentum = Math.abs(v) > 0.03 ? requestAnimationFrame(step) : null;
        };
        if (Math.abs(v) > 0.1) momentum = requestAnimationFrame(step);
      } else if (mode === 'card' && e.type === 'pointerup') {
        const vertical = -dy > Math.abs(dx);
        if (vertical && (-dy > thrY() || (vy < -0.7 && -dy > 40))) return flyOut(MAYBE, card);
        if (!vertical && (dx > thrX() || (vx > 0.7 && dx > 40))) return flyOut(INCLUDE, card);
        if (!vertical && (dx < -thrX() || (vx < -0.7 && dx < -40))) return flyOut(EXCLUDE, card);
        reset();
      } else if (mode === 'card') {
        reset();
      }
      mode = null;
    };
    card.addEventListener('pointerup', end);
    card.addEventListener('pointercancel', end);
  }

  document.addEventListener('keydown', e => {
    if (view !== 'screen' || !screen || e.metaKey || e.ctrlKey || e.altKey) return;
    if (e.target.closest && e.target.closest('select, input, textarea')) return;
    const k = e.key;
    if (k === 'ArrowRight') { e.preventDefault(); flyOut(INCLUDE); }
    else if (k === 'ArrowLeft') { e.preventDefault(); flyOut(EXCLUDE); }
    else if (k === 'ArrowUp') { e.preventDefault(); flyOut(MAYBE); }
    else if (k === 'Backspace' || k === 'z' || k === 'Z') { e.preventDefault(); goBack(); }
    else if (k === 'c' || k === 'C') { e.preventDefault(); toggleCriteria(); }
    else if (k === 'Escape') { e.preventDefault(); renderProject(); }
    else if (k === 'ArrowDown' || k === ' ') {
      const b = document.querySelector('#deck .card:not(.under) .body');
      if (b) { e.preventDefault(); b.scrollBy({ top: (k === ' ' && e.shiftKey ? -1 : 1) * b.clientHeight * 0.8, behavior: 'smooth' }); }
    }
  });

  /* ------------------------------------------------------------------ actions */
  const srcById = id => cur && cur.srcs.find(x => x.s.id === id);

  app.addEventListener('click', async e => {
    const t = e.target.closest('[data-act], [data-open], [data-stage], [data-dec], [data-columns], [data-remove], [data-remove-yes], [data-save-one], [data-reconnect]');
    if (!t || t.disabled) return;
    const d = t.dataset;
    if (d.open) return openProject(d.open);
    if (d.stage) return startStage(d.stage);
    if (d.dec) return flyOut(+d.dec);
    if (d.saveOne) { const src = srcById(d.saveOne); return src && exportFiles([src]); }
    if (d.columns) {
      const src = srcById(d.columns);
      if (!src) return;
      return renderMap({ fileName: src.s.fileName, ext: src.s.ext, bytes: null, handle: src.handle, wb: src.wb, sheet: src.wb.Sheets[src.s.sheet] ? src.s.sheet : src.wb.SheetNames[0], existing: src.s });
    }
    if (d.reconnect) {
      const src = srcById(d.reconnect);
      if (!src) return;
      const fresh = await loadSource(src.s.id, true);
      if (fresh && fresh.live) { Object.assign(src, fresh); queueLiveSave(src); toast('Reconnected'); }
      else toast('Sift still can’t edit that file. Allow access when Chrome asks.', 4000);
      return renderProject();
    }
    if (d.remove) {
      const slot = app.querySelector(`[data-slot="${CSS.escape(d.remove)}"]`);
      slot.innerHTML = `<span class="small muted">Remove this file and its decisions from the project? The Excel file itself isn’t touched.</span>
        <button class="link small danger-link" data-remove-yes="${esc(d.remove)}">Remove</button>
        <button class="link small" data-act="project">Cancel</button>`;
      return;
    }
    if (d.removeYes) {
      const id = d.removeYes;
      await db.del('sources', id).catch(() => {});
      await db.del('files', id).catch(() => {});
      await db.del('handles', id).catch(() => {}); handleCache.delete(id);
      cur.srcs = cur.srcs.filter(x => x.s.id !== id);
      cur.p.fileIds = cur.p.fileIds.filter(x => x !== id);
      for (const k of ['t', 'a']) if (cur.p.pos[k] && cur.p.pos[k].sid === id) cur.p.pos[k] = null;
      refreshStats();
      touchProject();
      await loadProjects();
      toast('File removed');
      return renderProject();
    }
    switch (d.act) {
      case 'home': await loadProjects(); return renderHome();
      case 'new': return renderNewProject();
      case 'add': return pickFile();
      case 'sample': return createSample();
      case 'project': return cur ? renderProject() : renderHome();
      case 'export': return exportAll();
      case 'live-retry': return retryLive();
      case 'undo': return goBack();
      case 'rename': return renderRename();
      case 'edit-criteria': return renderEditCriteria();
      case 'criteria': return toggleCriteria();
      case 'criteria-close': return toggleCriteria(false);
      case 'criteria-add': renderProject(); return renderEditCriteria();
      case 'resume': screen.qi = screen.review ? screen.qi + 1 : nextUndecided(screen.qi); return renderDeck(false);
      case 'review': screen.review = true; screen.qi = 0; return renderDeck(false);
      case 'delete':
        document.getElementById('del-slot').innerHTML = `<span class="confirm"><span class="small muted">Delete this project from this device? Your Excel files aren’t touched.</span>
          <button class="btn danger" data-act="delete-yes">Delete</button><button class="btn" data-act="project">Cancel</button></span>`;
        return;
      case 'delete-yes':
        for (const id of cur.p.fileIds) { await db.del('sources', id).catch(() => {}); await db.del('files', id).catch(() => {}); await db.del('handles', id).catch(() => {}); handleCache.delete(id); }
        await db.del('projects', cur.p.id).catch(() => {});
        await loadProjects();
        toast('Project deleted');
        return renderHome();
    }
  });

  /* ------------------------------------------------------------------ example project */
  async function createSample() {
    const rows = [
      ['ID', 'Title', 'Abstract', 'Authors', 'Year'],
      [1, 'Effect of an eight-week mindfulness programme on burnout in intensive care nurses: a randomised controlled trial',
        'Background: Burnout is common among intensive care nurses. Methods: 112 nurses were randomised to an eight-week mindfulness-based stress reduction course or a waiting list. The primary outcome was emotional exhaustion on the Maslach Burnout Inventory at 12 weeks. Results: Emotional exhaustion fell by 6.1 points in the intervention group compared with 1.2 points in controls (p < 0.01). Conclusions: A structured mindfulness course reduced burnout over three months.',
        'Example A, Example B', 2022],
      [2, 'Prevalence of burnout among junior doctors: a national cross-sectional survey',
        'We surveyed 3,410 junior doctors using the Copenhagen Burnout Inventory. Overall, 54% met the threshold for work-related burnout. Rotas with more than four night shifts per month were associated with higher scores. No intervention was assessed.',
        'Example C', 2021],
      [3, 'Mindfulness-based stress reduction for emergency physicians: a pilot feasibility study',
        'This single-arm pilot enrolled 24 emergency physicians in an adapted four-week mindfulness course. Recruitment, retention and acceptability were assessed. Retention was 83% and most participants rated the course as useful. Burnout scores were collected but the study was not powered to detect change.',
        'Example D, Example E', 2020],
      [4, 'Yoga and resilience in first-year medical students: a qualitative study',
        'Semi-structured interviews were held with 18 medical students after a term of weekly yoga classes. Themes included improved sleep, a sense of community and reduced exam anxiety. Burnout was not measured.',
        'Example F', 2023],
      [5, 'Online mindfulness training and emotional exhaustion in primary care staff: a cluster randomised trial',
        'Forty general practices were randomised to an app-based mindfulness programme or usual care. Across 612 staff, emotional exhaustion at six months was lower in the intervention arm (mean difference −3.4, 95% CI −5.1 to −1.7). Effects were larger for reception and nursing staff than for doctors.',
        'Example G, Example H', 2024],
      [6, 'The economics of physician turnover: a narrative review',
        '', 'Example I', 2019],
      [7, 'Peer support groups versus mindfulness for nurse burnout: a non-inferiority trial',
        'Nurses across six hospitals were randomised to facilitated peer support groups or a mindfulness course. Peer support was non-inferior to mindfulness for emotional exhaustion at 16 weeks. Attendance was higher in the peer support arm.',
        'Example J', 2023],
    ];
    const ws = XL.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 5 }, { wch: 60 }, { wch: 80 }, { wch: 20 }, { wch: 6 }];
    const wb = XL.utils.book_new();
    XL.utils.book_append_sheet(wb, ws, 'Records');
    const bytes = XL.write(wb, { bookType: 'xlsx', type: 'array' });
    const p = await createProject('Example review: mindfulness and burnout', true, {
      inc: 'Population: practising healthcare workers (doctors, nurses, allied health, support staff)\nIntervention: mindfulness-based programme of any length or format\nComparator: any (usual care, waiting list, active control)\nOutcome: burnout measured with a validated scale (e.g. Maslach Burnout Inventory)\nStudy design: randomised controlled trials, including cluster and pilot RCTs',
      exc: 'Students not yet in practice\nYoga, exercise or peer support without a mindfulness component\nStudies that do not measure burnout\nQualitative studies, reviews, editorials and conference abstracts',
    });
    cur = { p, srcs: [] };
    renderMap({ fileName: 'Example records.xlsx', ext: 'xlsx', bytes, handle: null, wb: readWorkbook(bytes), sheet: 'Records', existing: null });
  }

  /* ------------------------------------------------------------------ boot */
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden' && cur) {
      saveProject(cur.p);
      if (liveDirty.size) { clearTimeout(liveTimer); liveSave(); }
    }
  });

  async function boot() {
    if (!XL) { app.innerHTML = '<section class="view"><div class="wrap"><h1>Sift couldn’t start</h1><p class="note">The spreadsheet library didn’t load. Reload the page.</p></div></section>'; return; }
    try {
      if (sessionStorage.getItem('sift-reading-handle')) {
        sessionStorage.removeItem('sift-reading-handle');
        skipStoredHandles = true;
        setTimeout(() => toast('Sift couldn’t reconnect to your Excel file in this window, so automatic saving is off here. Use Save Excel file, or open Sift in a normal (not private) window.', 8000), 300);
      }
    } catch {}
    idb = await openIDB();
    if (!idb) toast('This browser is blocking storage, so progress will only last until you close the page.', 6000);
    try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch {}
    await loadProjects();
    renderHome();
    if ('serviceWorker' in navigator && location.protocol === 'https:') {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
  }
  boot();
})();
