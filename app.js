/* Sift Screener — swipe-based title/abstract screening that writes decisions back to Excel.
 * Everything runs in the browser. Projects live in IndexedDB on this device.
 * On desktop Chrome/Edge the Excel file itself is updated in the background (File System Access API).
 * Elsewhere (iPhone, Android, Safari, Firefox) "Save Excel" hands you the updated file. */
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

  const canLiveSave = typeof window.showOpenFilePicker === 'function' && (() => {
    try { return window.top === window.self; } catch { return false; }
  })();
  const isTouch = window.matchMedia('(pointer: coarse)').matches;

  /* ------------------------------------------------------------------ storage */
  let idb = null;
  const mem = { projects: new Map(), files: new Map() };

  function openIDB() {
    return new Promise(resolve => {
      try {
        const rq = indexedDB.open('sift-screener', 1);
        rq.onupgradeneeded = () => {
          const d = rq.result;
          d.createObjectStore('projects', { keyPath: 'id' });
          d.createObjectStore('files', { keyPath: 'id' });
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

  /* ------------------------------------------------------------------ helpers */
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = n => Number(n || 0).toLocaleString();
  const baseName = f => String(f || '').replace(/\.[^.]+$/, '');
  const extOf = f => (String(f || '').match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
  const uid = () => 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

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
    if (!cell || cell.v == null || cell.v === '') return undefined;
    const n = Number(String(cell.v).trim());
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
  function applyDecisions(ws, p) {
    const rg = rangeOf(ws);
    let maxC = rg.e.c;
    for (const [c, name] of Object.entries(p.newHeaders || {})) {
      const addr = XL.utils.encode_cell({ r: rg.s.r, c: +c });
      if (!ws[addr] || ws[addr].v == null || ws[addr].v === '') ws[addr] = { t: 's', v: name };
      maxC = Math.max(maxC, +c);
    }
    for (const [key, col] of [['t', p.cols.titleDec], ['a', p.cols.absDec]]) {
      maxC = Math.max(maxC, col);
      for (const [r, v] of Object.entries(p.decisions[key])) {
        ws[XL.utils.encode_cell({ r: +r, c: col })] = { t: 'n', v };
      }
    }
    rg.e.c = maxC;
    ws['!ref'] = XL.utils.encode_range(rg);
  }
  function outputFile(p, wb) {
    const ws = wb.Sheets[p.sheet];
    applyDecisions(ws, p);
    if (p.ext === 'csv') {
      const csv = XL.utils.sheet_to_csv(ws);
      return { name: p.fileName, blob: new Blob(['﻿' + csv], { type: 'text/csv' }) };
    }
    const out = XL.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true, compression: true });
    const name = p.ext === 'xlsx' ? p.fileName : baseName(p.fileName) + '.xlsx';
    return { name, blob: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }) };
  }

  /* ------------------------------------------------------------------ state */
  let projects = [];
  let cur = null;    // { p, wb, ws, rows, handle, live }
  let screen = null; // { stage, queue, qi }
  let view = 'home';

  function stats(p, rows) {
    const d = p.decisions;
    const s = { t: { total: rows.length, done: 0, 1: 0, 0: 0, 2: 0 }, a: { total: 0, done: 0, 1: 0, 0: 0, 2: 0 } };
    for (const x of rows) {
      const tv = d.t[x.r];
      if (tv !== undefined) { s.t.done++; s.t[tv]++; }
      if (tv === INCLUDE || tv === MAYBE) {
        s.a.total++;
        const av = d.a[x.r];
        if (av !== undefined) { s.a.done++; s.a[av]++; }
      }
    }
    return s;
  }
  function saveProject(p) {
    return db.put('projects', p).catch(() => toast('Couldn’t save progress on this device. Check that site storage is allowed.', 5000));
  }
  async function loadProjects() {
    projects = (await db.all('projects').catch(() => [])) || [];
    projects.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  }

  /* ------------------------------------------------------------------ home */
  function barHTML(s) {
    if (!s || !s.total) return '<span class="bar"></span>';
    const w = k => (100 * s[k] / s.total).toFixed(2) + '%';
    return `<span class="bar"><i class="i" style="width:${w(1)}"></i><i class="m" style="width:${w(2)}"></i><i class="e" style="width:${w(0)}"></i></span>`;
  }
  function renderHome() {
    view = 'home'; cur = null; screen = null;
    const list = projects.map(p => {
      const s = p.stats || { t: { total: 0, done: 0 }, a: { total: 0, done: 0 } };
      return `<button class="project" data-open="${esc(p.id)}">
        <span class="p-name">${esc(p.name)}${p.sample ? ' <span class="tag">Example data</span>' : ''}</span>
        <span class="chev">${ICON.chev}</span>
        <span class="p-meta num">
          <span>${barHTML(s.t)}Titles ${fmt(s.t.done)} / ${fmt(s.t.total)}</span>
          <span>${barHTML(s.a)}Abstracts ${fmt(s.a.done)} / ${fmt(s.a.total)}</span>
        </span>
      </button>`;
    }).join('');

    app.innerHTML = `<section class="view"><div class="wrap">
      <div class="brand"><b>Sift</b><span>Systematic review screening</span></div>
      ${projects.length ? `
        <div class="row" style="justify-content:space-between">
          <h2>Projects</h2>
          <button class="btn primary" data-act="new">New project</button>
        </div>
        <div class="projects">${list}</div>
      ` : `
        <div class="empty">
          <h1>Screen records by swiping</h1>
          <p class="note">Open your search results as an Excel or CSV file. Swipe right to include, left to exclude, up for maybe. Every decision is written to your screening columns as 1, 0 or 2.</p>
          <div class="row">
            <button class="btn primary" data-act="new">Open Excel file</button>
            <button class="btn" data-act="sample">Try with example records</button>
          </div>
        </div>
      `}
      <p class="note">${canLiveSave
        ? 'On this browser, decisions are written straight into your Excel file as you swipe.'
        : 'Progress is saved on this device after every swipe. Use <b>Save Excel</b> to get the updated spreadsheet. For automatic saving into the file itself, open Sift in Chrome or Edge on a computer.'}</p>
      ${projects.length ? '<div><button class="link" data-act="sample">Add an example project</button></div>' : ''}
    </div></section>`;
  }

  /* ------------------------------------------------------------------ opening files */
  async function pickFile() {
    if (canLiveSave) {
      try {
        const [handle] = await window.showOpenFilePicker({
          multiple: false,
          types: [{ description: 'Spreadsheet', accept: {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
            'application/vnd.ms-excel': ['.xls'],
            'text/csv': ['.csv'],
          } }],
        });
        const file = await handle.getFile();
        return startImport(file, handle);
      } catch (e) {
        if (e && e.name === 'AbortError') return;
        // fall through to the plain picker
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
    const ext = extOf(file.name);
    if (!['xlsx', 'xls', 'csv', 'xlsm'].includes(ext)) { toast('Choose an .xlsx, .xls or .csv file.'); return; }
    let bytes, wb;
    try {
      bytes = await file.arrayBuffer();
      wb = readWorkbook(bytes);
    } catch {
      toast('That file couldn’t be read as a spreadsheet.', 4000);
      return;
    }
    if (!wb.SheetNames.length) { toast('That file has no sheets.'); return; }
    renderMap({ fileName: file.name, ext, bytes, handle, wb, sheet: wb.SheetNames[0], existing: null, sample: false });
  }

  /* ------------------------------------------------------------------ column mapping */
  function guess(headers, test, avoid) {
    const h = headers.find(x => test(x.name) && !(avoid && avoid(x.name)));
    return h ? h.c : null;
  }
  const isDecisionName = n => /screen|decision|include|exclu|incl|verdict/i.test(n);

  function renderMap(draft) {
    view = 'map';
    const ws = draft.wb.Sheets[draft.sheet];
    const headers = headersOf(ws);
    const ex = draft.existing;
    const prev = ex && ex.sheet === draft.sheet ? ex.cols : null;

    const sel = {
      title: prev ? prev.title : (guess(headers, n => /^\s*(article\s+|primary\s+)?title\s*$/i.test(n)) ?? guess(headers, n => /title/i.test(n), isDecisionName)),
      abs: prev ? prev.abs : (guess(headers, n => /^\s*abstract\s*$/i.test(n)) ?? guess(headers, n => /abstract|summary/i.test(n), isDecisionName)),
      titleDec: prev ? prev.titleDec : (guess(headers, n => /title/i.test(n) && isDecisionName(n)) ?? 'new'),
      absDec: prev ? prev.absDec : (guess(headers, n => /abstract/i.test(n) && isDecisionName(n)) ?? 'new'),
    };
    if (prev && ex.newHeaders) {
      if (ex.newHeaders[prev.titleDec] && !headers.some(h => h.c === prev.titleDec && h.name)) sel.titleDec = 'new';
      if (ex.newHeaders[prev.absDec] && !headers.some(h => h.c === prev.absDec && h.name)) sel.absDec = 'new';
    }

    const opts = (val, allowNew, newLabel) =>
      (allowNew ? `<option value="new"${val === 'new' ? ' selected' : ''}>+ New column “${esc(newLabel)}”</option>` : `<option value=""${val == null ? ' selected' : ''} disabled>Choose a column</option>`) +
      headers.map(h => `<option value="${h.c}"${String(val) === String(h.c) ? ' selected' : ''}>${esc(h.label)}</option>`).join('');

    app.innerHTML = `<section class="view"><div class="wrap">
      <div><button class="link" data-act="${ex ? 'project' : 'home'}">${ICON.back} ${ex ? 'Back to project' : 'Projects'}</button></div>
      <div style="display:grid;gap:6px">
        <h1>${ex ? 'Change columns' : 'Match your columns'}</h1>
        <p class="muted small">${esc(draft.fileName)}</p>
      </div>
      <div class="fields">
        ${draft.wb.SheetNames.length > 1 ? `<div class="field"><label for="m-sheet">Sheet</label>
          <select id="m-sheet">${draft.wb.SheetNames.map(n => `<option${n === draft.sheet ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select></div>` : ''}
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
      <div class="preview" id="m-preview"></div>
      <p class="error" id="m-err" hidden></p>
      <div class="row"><button class="btn primary" id="m-go">${ex ? 'Save columns' : 'Create project'}</button></div>
    </div></section>`;

    const $ = id => document.getElementById(id);
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
      const total = rg.e.r - rg.s.r;
      $('m-preview').innerHTML = `<h2>First record · ${fmt(total)} rows</h2>
        <div class="t">${t ? esc(t) : '<span class="faint">No title in this column</span>'}</div>
        <div class="a">${a ? esc(a) : '<span class="faint">No abstract in this column</span>'}</div>`;
    };
    preview();
    ['m-title', 'm-abs', 'm-tdec', 'm-adec'].forEach(id => $(id).addEventListener('change', preview));
    if ($('m-sheet')) $('m-sheet').addEventListener('change', e => renderMap({ ...draft, sheet: e.target.value }));

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

  async function finishMapping(draft, cols, newHeaders) {
    const ws = draft.wb.Sheets[draft.sheet];
    const rows = readRows(ws, cols);
    if (!rows.length) { toast('No records found in those columns.'); return; }
    if (!draft.bytes) draft.bytes = XL.write(draft.wb, { bookType: 'xlsx', type: 'array' });
    const ex = draft.existing;
    const p = ex ? { ...ex } : {
      id: uid(), name: baseName(draft.fileName), fileName: draft.fileName, ext: draft.ext,
      createdAt: Date.now(), decisions: { t: {}, a: {} }, pos: { t: null, a: null }, unsaved: 0, sample: draft.sample,
    };
    if (ex && ex.sheet !== draft.sheet) { p.decisions = { t: {}, a: {} }; p.pos = { t: null, a: null }; }
    p.sheet = draft.sheet;
    p.cols = cols;
    p.newHeaders = newHeaders;
    p.fingerprint = fingerprint(draft.sheet, rows);
    mergeFileDecisions(p, rows);
    p.stats = stats(p, rows);
    p.updatedAt = Date.now();
    p.live = !!draft.handle;

    await saveProject(p);
    await db.put('files', { id: p.id, bytes: draft.bytes, handle: draft.handle || null }).catch(() => {});
    await loadProjects();
    cur = { p, wb: draft.wb, ws, rows, handle: draft.handle || null, live: !!draft.handle };
    if (cur.live) queueLiveSave();
    renderProject();
  }

  function mergeFileDecisions(p, rows) {
    for (const x of rows) {
      if (x.ft !== undefined && p.decisions.t[x.r] === undefined) p.decisions.t[x.r] = x.ft;
      if (x.fa !== undefined && p.decisions.a[x.r] === undefined) p.decisions.a[x.r] = x.fa;
    }
  }

  /* ------------------------------------------------------------------ open project */
  async function openProject(id) {
    const p = projects.find(x => x.id === id);
    if (!p) return renderHome();
    const rec = await db.get('files', id).catch(() => null);
    if (!rec) { toast('The spreadsheet for this project is missing on this device.', 4000); return; }

    let bytes = rec.bytes, handle = rec.handle || null, live = false;
    if (handle && canLiveSave) {
      try {
        let perm = await handle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted') perm = await handle.requestPermission({ mode: 'readwrite' });
        if (perm === 'granted') {
          const file = await handle.getFile();
          bytes = await file.arrayBuffer();
          live = true;
          db.put('files', { id, bytes, handle }).catch(() => {});
        }
      } catch {
        toast('Couldn’t reach the original file. Using the copy saved on this device.', 4000);
      }
    }

    let wb;
    try { wb = readWorkbook(bytes); } catch { toast('The saved spreadsheet couldn’t be read.'); return; }
    const ws = wb.Sheets[p.sheet];
    if (!ws) { toast(`Sheet “${p.sheet}” is no longer in the file. Pick the columns again.`, 4000); renderMap({ fileName: p.fileName, ext: p.ext, bytes, handle, wb, sheet: wb.SheetNames[0], existing: p, sample: p.sample }); return; }
    const rows = readRows(ws, p.cols);
    mergeFileDecisions(p, rows);
    p.stats = stats(p, rows);
    p.live = live;
    cur = { p, wb, ws, rows, handle, live };
    renderProject();
  }

  /* ------------------------------------------------------------------ project view */
  function renderProject() {
    view = 'project'; screen = null;
    const { p, rows } = cur;
    p.stats = stats(p, rows);
    const s = p.stats;
    const stageCard = (k) => {
      const st = s[k];
      const left = st.total - st.done;
      const locked = k === 'a' && st.total === 0;
      const verb = st.done === 0 ? 'Start' : left === 0 ? 'Review' : 'Continue';
      return `<div class="stage">
        <h2>${STAGE_NAME[k]}</h2>
        <div class="big num">${fmt(st.done)}<small> / ${fmt(st.total)} screened</small></div>
        ${barHTML(st)}
        <div class="tally num">
          <span class="ci"><b>${fmt(st[1])}</b>Include</span>
          <span class="ce"><b>${fmt(st[0])}</b>Exclude</span>
          <span class="cm"><b>${fmt(st[2])}</b>Maybe</span>
          <span><b>${fmt(left)}</b>Left</span>
        </div>
        ${k === 'a' ? `<p class="small muted">Records included or marked maybe at title stage.</p>` : `<p class="small muted">Every record in the sheet.</p>`}
        <button class="btn ${locked ? '' : 'primary'}" data-stage="${k}" ${locked ? 'disabled' : ''}>${locked ? 'Include some titles first' : verb}</button>
      </div>`;
    };
    app.innerHTML = `<section class="view"><div class="wrap">
      <div><button class="link" data-act="home">${ICON.back} Projects</button></div>
      <div style="display:grid;gap:6px">
        <h1>${esc(p.name)}</h1>
        <p class="muted small num">${esc(p.fileName)} · ${esc(p.sheet)} · ${fmt(rows.length)} records${p.sample ? ' · <span class="tag">Example data</span>' : ''}</p>
      </div>
      <div class="stages">${stageCard('t')}${stageCard('a')}</div>
      <div class="row">
        ${saveButtonHTML('btn')}
        <button class="btn" data-act="columns">Change columns</button>
        <span id="del-slot"><button class="btn danger" data-act="delete">Delete project</button></span>
      </div>
      <p class="note">${cur.live
        ? 'Live file: every decision is written into the Excel file on this computer a moment after you swipe.'
        : 'Progress is saved on this device after every swipe. <b>Save Excel</b> gives you the spreadsheet with all decisions filled in.'}</p>
    </div></section>`;
  }

  function saveButtonHTML(cls) {
    const p = cur.p;
    if (cur.live) {
      const st = liveState;
      const text = st === 'saving' ? 'Saving…' : st === 'error' ? 'Not saved · retry' : 'Saved to file';
      return `<button class="${cls === 'btn' ? 'btn' : 'save'} ${st === 'error' ? 'err' : st === 'saving' ? 'pending' : 'ok'}" data-act="live-retry" id="save-btn">${text}</button>`;
    }
    const n = p.unsaved || 0;
    if (cls === 'btn') return `<button class="btn" data-act="export" id="save-btn">Save Excel${n ? ` · ${fmt(n)} new` : ''}</button>`;
    return `<button class="save ${n ? 'pending' : 'ok'}" data-act="export" id="save-btn">${n ? `Save Excel · ${fmt(n)}` : 'Save Excel'}</button>`;
  }
  function refreshSaveButton() {
    const b = document.getElementById('save-btn');
    if (!b || !cur) return;
    const tmp = document.createElement('div');
    tmp.innerHTML = saveButtonHTML(b.classList.contains('btn') ? 'btn' : 'save');
    b.replaceWith(tmp.firstElementChild);
  }

  /* ------------------------------------------------------------------ saving to Excel */
  let liveTimer = null, liveState = 'ok', liveBusy = false, liveAgain = false;
  function queueLiveSave() {
    if (!cur || !cur.live) return;
    clearTimeout(liveTimer);
    liveTimer = setTimeout(liveSave, 700);
  }
  async function liveSave() {
    if (!cur || !cur.live) return;
    if (liveBusy) { liveAgain = true; return; }
    liveBusy = true; liveState = 'saving'; refreshSaveButton();
    const c = cur;
    try {
      const { blob } = outputFile(c.p, c.wb);
      const w = await c.handle.createWritable();
      await w.write(blob);
      await w.close();
      liveState = 'ok';
      c.p.unsaved = 0;
      saveProject(c.p);
      blob.arrayBuffer().then(bytes => db.put('files', { id: c.p.id, bytes, handle: c.handle })).catch(() => {});
    } catch (e) {
      liveState = 'error';
      toast(e && e.name === 'NoModificationAllowedError'
        ? 'The file is locked by another app. Close it there and tap “retry”.'
        : 'Couldn’t write to the Excel file. Your progress is still saved here.', 5000);
    }
    liveBusy = false;
    refreshSaveButton();
    if (liveAgain) { liveAgain = false; queueLiveSave(); }
  }

  async function exportExcel() {
    if (!cur) return;
    const { p, wb } = cur;
    let file;
    try { file = outputFile(p, wb); } catch { toast('Couldn’t build the spreadsheet.'); return; }
    const done = () => { p.unsaved = 0; p.exportedAt = Date.now(); saveProject(p); refreshSaveButton(); };

    // Phones: the share sheet lets you "Save to Files" / Drive / email.
    if (isTouch && navigator.canShare) {
      const f = new File([file.blob], file.name, { type: file.blob.type });
      if (navigator.canShare({ files: [f] })) {
        try { await navigator.share({ files: [f] }); done(); toast('Saved'); return; }
        catch (e) { if (e && e.name === 'AbortError') return; }
      }
    }
    const url = URL.createObjectURL(file.blob);
    const a = document.createElement('a');
    a.href = url; a.download = file.name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    done();
    toast(`Downloaded ${file.name}`);
  }

  /* ------------------------------------------------------------------ screening */
  function buildQueue(stage) {
    const d = cur.p.decisions;
    return stage === 't' ? cur.rows : cur.rows.filter(x => d.t[x.r] === INCLUDE || d.t[x.r] === MAYBE);
  }
  function nextUndecided(from) {
    const { queue, stage } = screen;
    const d = cur.p.decisions[stage];
    for (let i = from; i < queue.length; i++) if (d[queue[i].r] === undefined) return i;
    for (let i = 0; i < Math.min(from, queue.length); i++) if (d[queue[i].r] === undefined) return i;
    return queue.length;
  }
  function startStage(stage) {
    const queue = buildQueue(stage);
    screen = { stage, queue, qi: 0, busy: false };
    const pos = cur.p.pos[stage];
    let qi = pos == null || pos === -1 ? -1 : queue.findIndex(x => x.r === pos);
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
      <div class="deck" id="deck"></div>
      <div class="controls" id="controls">
        <button class="circle back" data-act="undo" aria-label="Previous record">${ICON.undo}</button>
        <button class="circle no" data-dec="0" aria-label="Exclude">${ICON.x}</button>
        <button class="circle may" data-dec="2" aria-label="Maybe">${ICON.q}</button>
        <button class="circle yes" data-dec="1" aria-label="Include">${ICON.check}</button>
        <span class="spacer40" aria-hidden="true"></span>
      </div>
      <div class="hints" id="hints">${isTouch
        ? 'Swipe right to include · left to exclude · up for maybe'
        : '→ include · ← exclude · ↑ maybe · Backspace for previous'}</div>
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
    const st = stats(cur.p, cur.rows)[stage];
    cur.p.stats = stats(cur.p, cur.rows);
    const pos = Math.min(qi + 1, queue.length);
    document.getElementById('s-count').textContent = qi < queue.length
      ? `Record ${fmt(pos)} of ${fmt(queue.length)} · ${fmt(st.done)} screened`
      : `${fmt(st.done)} of ${fmt(queue.length)} screened`;
    document.getElementById('s-bar').innerHTML = barHTML(st);
  }

  function cardHTML(x, idx, under) {
    const { stage, queue } = screen;
    const v = cur.p.decisions[stage][x.r];
    const isAbs = stage === 'a';
    const chip = v !== undefined ? `<span class="tag ${CLS[v]}">${LABEL[v]}</span>` : '';
    return `<article class="card ${isAbs ? 'abs' : ''} ${under ? 'under' : ''}" ${under ? 'aria-hidden="true"' : ''}>
      <div class="meta num"><span>${fmt(idx + 1)} / ${fmt(queue.length)} · Excel row ${fmt(x.r + 1)}</span>${chip}</div>
      <div class="body">
        <div class="inner ${isAbs ? 'top' : ''}">
          <div class="ttl">${x.t ? esc(x.t) : '<span class="faint">No title</span>'}</div>
          ${isAbs ? (x.a ? `<div class="abstract">${esc(x.a)}</div>` : '<div class="abstract none">No abstract in this record.</div>') : ''}
        </div>
      </div>
      ${under ? '' : `<span class="stamp s-in">INCLUDE</span><span class="stamp s-out">EXCLUDE</span><span class="stamp s-may">MAYBE</span>`}
      ${under || v === undefined ? '' : `<div class="row" style="justify-content:center"><button class="link small" data-act="resume">Skip to next unscreened record</button></div>`}
    </article>`;
  }

  function renderDeck(animateIn) {
    const deck = document.getElementById('deck');
    const { queue, qi } = screen;
    updateHeader();
    const controls = document.getElementById('controls');
    const hints = document.getElementById('hints');
    if (qi >= queue.length) {
      const st = stats(cur.p, cur.rows)[screen.stage];
      const left = queue.length - st.done;
      deck.innerHTML = `<div class="done"><div>
        <h1>${left ? `${fmt(left)} records still unscreened` : `All ${fmt(queue.length)} records screened`}</h1>
        <p class="muted num">${fmt(st[1])} included · ${fmt(st[0])} excluded · ${fmt(st[2])} maybe</p>
        <div class="row" style="justify-content:center">
          ${cur.live ? '' : `<button class="btn primary" data-act="export">Save Excel</button>`}
          ${screen.stage === 't' && st[1] + st[2] > 0 ? `<button class="btn${cur.live ? ' primary' : ''}" data-stage="a">Start abstract screening</button>` : ''}
          <button class="btn" data-act="review">Review from the start</button>
        </div>
      </div></div>`;
      controls.querySelectorAll('[data-dec]').forEach(b => { b.disabled = true; });
      hints.hidden = true;
      cur.p.pos[screen.stage] = -1;
      saveProject(cur.p);
      return;
    }
    controls.querySelectorAll('[data-dec]').forEach(b => { b.disabled = false; });
    hints.hidden = false;
    const nextIdx = (() => {
      const d = cur.p.decisions[screen.stage];
      for (let i = qi + 1; i < queue.length; i++) if (d[queue[i].r] === undefined) return i;
      return -1;
    })();
    deck.innerHTML = (nextIdx >= 0 ? cardHTML(queue[nextIdx], nextIdx, true) : '') + cardHTML(queue[qi], qi, false);
    const card = deck.querySelector('.card:not(.under)');
    if (animateIn) {
      card.classList.add('enter', 'dragging');
      requestAnimationFrame(() => { card.classList.remove('dragging'); requestAnimationFrame(() => card.classList.remove('enter')); });
    }
    attachGestures(card);
    cur.p.pos[screen.stage] = queue[qi].r;
  }

  function decide(v) {
    const { stage, queue, qi } = screen;
    const x = queue[qi];
    const prev = cur.p.decisions[stage][x.r];
    cur.p.decisions[stage][x.r] = v;
    if (prev !== v) cur.p.unsaved = (cur.p.unsaved || 0) + 1;
    cur.p.updatedAt = Date.now();
    screen.qi = nextUndecided(qi + 1);
    cur.p.pos[stage] = screen.qi < queue.length ? queue[screen.qi].r : -1;
    cur.p.stats = stats(cur.p, cur.rows);
    saveProject(cur.p);
    if (cur.live) queueLiveSave();
    renderDeck(true);
    const s = document.getElementById('s-save');
    if (s) s.innerHTML = saveButtonHTML('save');
  }

  function flyOut(v, card) {
    if (screen.busy || screen.qi >= screen.queue.length) return;
    card = card || document.querySelector('#deck .card:not(.under)');
    if (!card) return;
    screen.busy = true;
    const W = window.innerWidth, H = window.innerHeight;
    const cur_t = card.style.transform || '';
    const m = /translate\(([-\d.]+)px,\s*([-\d.]+)px\)/.exec(cur_t);
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
    else if (k === 'Escape') { e.preventDefault(); renderProject(); }
    else if (k === 'ArrowDown' || k === ' ') {
      const b = document.querySelector('#deck .card:not(.under) .body');
      if (b) { e.preventDefault(); b.scrollBy({ top: (k === ' ' && e.shiftKey ? -1 : 1) * b.clientHeight * 0.8, behavior: 'smooth' }); }
    }
  });

  /* ------------------------------------------------------------------ actions */
  app.addEventListener('click', async e => {
    const t = e.target.closest('[data-act], [data-open], [data-stage], [data-dec]');
    if (!t || t.disabled) return;
    if (t.dataset.open) return openProject(t.dataset.open);
    if (t.dataset.stage) return startStage(t.dataset.stage);
    if (t.dataset.dec) return flyOut(+t.dataset.dec);
    switch (t.dataset.act) {
      case 'home': await loadProjects(); return renderHome();
      case 'new': return pickFile();
      case 'sample': return createSample();
      case 'project': return cur ? renderProject() : renderHome();
      case 'export': return exportExcel();
      case 'live-retry': return liveSave();
      case 'undo': return goBack();
      case 'resume': screen.qi = nextUndecided(screen.qi); return renderDeck(false);
      case 'review': screen.qi = 0; return renderDeck(false);
      case 'columns': {
        const rec = await db.get('files', cur.p.id).catch(() => null);
        return renderMap({ fileName: cur.p.fileName, ext: cur.p.ext, bytes: rec ? rec.bytes : null, handle: cur.handle, wb: cur.wb, sheet: cur.p.sheet, existing: cur.p, sample: cur.p.sample });
      }
      case 'delete':
        document.getElementById('del-slot').innerHTML = `<span class="confirm"><span class="small muted">Delete from this device? Your Excel file is not touched.</span>
          <button class="btn danger" data-act="delete-yes">Delete</button><button class="btn" data-act="project">Cancel</button></span>`;
        return;
      case 'delete-yes':
        await db.del('projects', cur.p.id).catch(() => {});
        await db.del('files', cur.p.id).catch(() => {});
        await loadProjects();
        toast('Project deleted');
        return renderHome();
    }
  });

  /* ------------------------------------------------------------------ example project */
  function createSample() {
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
    const wb2 = readWorkbook(bytes);
    renderMap({ fileName: 'Example records - mindfulness and burnout.xlsx', ext: 'xlsx', bytes, handle: null, wb: wb2, sheet: 'Records', existing: null, sample: true });
  }

  /* ------------------------------------------------------------------ boot */
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden' && cur) {
      saveProject(cur.p);
      if (cur.live && liveTimer) { clearTimeout(liveTimer); liveSave(); }
    }
  });

  async function boot() {
    if (!XL) { app.innerHTML = '<section class="view"><div class="wrap"><h1>Sift couldn’t start</h1><p class="note">The spreadsheet library didn’t load. Reload the page.</p></div></section>'; return; }
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
