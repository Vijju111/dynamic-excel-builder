import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { HotTable } from '@handsontable/react';
import { registerAllModules } from 'handsontable/registry';
import { FileSpreadsheet, Download, Plus, Trash2, RefreshCw, Upload, AlertTriangle } from 'lucide-react';
import { parseExcel, importWorkingSheet, downloadExport } from './api.js';
import { clearSession, loadSession, saveSession } from './db.js';

registerAllModules();

const EMPTY_ROWS = 25;
const normalize = (v) => String(v ?? '').trim();
const lookupKey = (v) => normalize(v).toLocaleLowerCase();
const blankRow = (columns) => Object.fromEntries(columns.map((c) => [c.id, '']));

function makeColumnId(name, existing = new Set()) {
  let base = normalize(name).toLowerCase().replace(/\W+/gu, '_').replace(/^_+|_+$/g, '') || 'column';
  let id = base;
  let i = 1;
  while (existing.has(id)) id = `${base}_${i++}`;
  existing.add(id);
  return id;
}

function createInitialRows(columns, count = EMPTY_ROWS) {
  return Array.from({ length: count }, () => blankRow(columns));
}

function migrateRows(oldRows, oldColumns, newColumns) {
  const findOld = (newCol) => oldColumns.find((old) =>
    old.id === newCol.id ||
    (!!old.sourceColumn && old.sourceColumn === newCol.sourceColumn) ||
    old.displayName === newCol.displayName
  );
  return oldRows.map((row) => {
    const next = {};
    for (const col of newColumns) {
      const old = findOld(col);
      next[col.id] = old ? (row[old.id] ?? '') : '';
    }
    return next;
  });
}

function mergeColumns(configColumns, liveColumns) {
  const result = [...configColumns];
  for (const live of liveColumns) {
    const exists = result.some((c) => c.id === live.id || (!!c.sourceColumn && c.sourceColumn === live.sourceColumn) || c.displayName === live.displayName);
    if (!exists) result.push(live);
  }
  return result;
}

function App() {
  const hotRef = useRef(null);
  const [step, setStep] = useState(1);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [master, setMaster] = useState(null);
  const [headerRow, setHeaderRow] = useState(0);
  const [selected, setSelected] = useState([]);
  const [customColumns, setCustomColumns] = useState([]);
  const [lookupColumn, setLookupColumn] = useState('');
  const [duplicateMode, setDuplicateMode] = useState('first');
  const [overwriteMode, setOverwriteMode] = useState('protect');
  const [workingColumns, setWorkingColumns] = useState([]);
  const [rows, setRows] = useState([]);
  const [cellMeta, setCellMeta] = useState({});
  const [rowStatus, setRowStatus] = useState({});
  const [filename, setFilename] = useState('Generated_Sheet');
  const [savedAt, setSavedAt] = useState(null);

  useEffect(() => {
    loadSession().then((session) => {
      if (session && confirm('A previous local session was found. Restore it?')) {
        setStep(session.step || 1);
        setMaster(session.master || null);
        setHeaderRow(session.headerRow || 0);
        setSelected(session.selected || []);
        setCustomColumns(session.customColumns || []);
        setLookupColumn(session.lookupColumn || '');
        setDuplicateMode(session.duplicateMode || 'first');
        setOverwriteMode(session.overwriteMode || 'protect');
        setWorkingColumns(session.workingColumns || []);
        setRows(session.rows || []);
        setCellMeta(session.cellMeta || {});
        setRowStatus(session.rowStatus || {});
        setFilename(session.filename || 'Generated_Sheet');
        setSavedAt(session.savedAt || null);
      }
    }).catch(console.error);
  }, []);

  useEffect(() => {
    const handler = (e) => {
      if (master || rows.length) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [master, rows.length]);

  useEffect(() => {
    if (!master && !rows.length) return;
    const state = { step, master, headerRow, selected, customColumns, lookupColumn, duplicateMode, overwriteMode, workingColumns, rows, cellMeta, rowStatus, filename };
    const t = setTimeout(() => {
      saveSession(state).then(() => setSavedAt(new Date().toLocaleString())).catch(console.error);
    }, 500);
    return () => clearTimeout(t);
  }, [step, master, headerRow, selected, customColumns, lookupColumn, duplicateMode, overwriteMode, workingColumns, rows, cellMeta, rowStatus, filename]);

  const duplicateLookupCount = useMemo(() => {
    if (!master || !lookupColumn) return 0;
    const counts = new Map();
    for (const r of master.rows) {
      const key = lookupKey(r[lookupColumn]);
      if (key) counts.set(key, (counts.get(key) || 0) + 1);
    }
    let duplicates = 0;
    counts.forEach((v) => { if (v > 1) duplicates += 1; });
    return duplicates;
  }, [master, lookupColumn]);

  const lookupIndex = useMemo(() => {
    const map = new Map();
    if (!master || !lookupColumn) return map;
    for (const record of master.rows) {
      const key = lookupKey(record[lookupColumn]);
      if (!key) continue;
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(record);
    }
    return map;
  }, [master, lookupColumn]);

  async function handleUpload(file) {
    if (!file) return;
    setBusy(true); setError(''); setNotice('');
    try {
      const data = await parseExcel(file, headerRow);
      const existing = new Set();
      setMaster(data);
      const cols = data.columns.map((c) => ({ id: makeColumnId(c.displayName, existing), displayName: c.displayName, sourceColumn: c.originalName, isCustom: false }));
      setSelected(cols);
      setLookupColumn(data.columns[0]?.originalName || '');
      setStep(2);
      setNotice(`Loaded ${data.rowCount} rows and ${data.columnCount} columns from ${data.fileName}.`);
    } catch (e) { setError(e.message); }
    finally { setBusy(false); }
  }

  function updateSelected(id, patch) {
    setSelected((prev) => prev.map((c) => c.id === id ? { ...c, ...patch } : c));
  }
  function updateCustom(id, patch) {
    setCustomColumns((prev) => prev.map((c) => c.id === id ? { ...c, ...patch } : c));
  }
  function move(listSetter, index, dir) {
    listSetter((prev) => {
      const next = [...prev];
      const j = index + dir;
      if (j < 0 || j >= next.length) return prev;
      [next[index], next[j]] = [next[j], next[index]];
      return next;
    });
  }

  function addCustomColumnToConfig() {
    const col = createCustomColumn();
    if (!col) return;
    setCustomColumns((prev) => [...prev, col]);
  }

  function createCustomColumn(defaultName = '') {
    const name = prompt('Custom column name:', defaultName);
    if (!name) return null;
    const clean = name.trim();
    const exists = [...selected, ...customColumns, ...workingColumns].some((c) => c.displayName.toLowerCase() === clean.toLowerCase());
    if (exists) {
      alert('Duplicate column names are not allowed.');
      return null;
    }
    const existingIds = new Set([...selected, ...customColumns, ...workingColumns].map((c) => c.id));
    return { id: makeColumnId(clean, existingIds), displayName: clean, sourceColumn: null, isCustom: true };
  }

  function generateSheet() {
    setError('');
    if (!lookupColumn) return setError('Please select a lookup column before generating the sheet.');
    if (!selected.some((c) => c.sourceColumn === lookupColumn)) return setError('The lookup column must be included in selected columns so users can type lookup values.');
    const configColumns = [...selected, ...customColumns];
    const columns = rows.length ? mergeColumns(configColumns, workingColumns.filter((c) => c.isCustom)) : configColumns;
    const nextRows = rows.length ? migrateRows(rows, workingColumns, columns) : createInitialRows(columns);
    setWorkingColumns(columns);
    setRows(nextRows);
    setStep(3);
    setNotice(rows.length ? 'Sheet updated. Your previously entered data was preserved.' : 'Live working sheet generated.');
  }

  function goBackToConfigure() {
    // Keep live columns visible in Configure so data/columns are never lost when the user goes back.
    setSelected(workingColumns.filter((c) => !c.isCustom));
    setCustomColumns(workingColumns.filter((c) => c.isCustom));
    setStep(2);
  }

  const gridColumns = useMemo(() => workingColumns.map((c) => ({ data: c.id, type: 'text', wordWrap: true })), [workingColumns]);
  const gridHeaders = useMemo(() => workingColumns.map((c) => c.displayName), [workingColumns]);
  const lookupWorkingCol = useMemo(() => workingColumns.find((c) => c.sourceColumn === lookupColumn), [workingColumns, lookupColumn]);

  const chooseFromMatches = useCallback((matches, title) => {
    if (!matches.length) return null;
    if (matches.length === 1) return matches[0];
    const choices = matches.slice(0, 10).map((m, i) => `${i + 1}. ${Object.entries(m).slice(0, 6).map(([k, v]) => `${k}: ${v}`).join(' | ')}`).join('\n');
    const pick = Number(prompt(`${title}\nEnter choice number or cancel:\n${choices}`, '1'));
    if (!pick) return null;
    return matches[Math.max(0, Math.min(matches.length - 1, pick - 1))];
  }, []);

  const getMatch = useCallback((key) => {
    const exactKey = lookupKey(key);
    if (!exactKey) return null;
    const exactMatches = lookupIndex.get(exactKey) || [];
    if (exactMatches.length) {
      if (duplicateMode === 'latest') return exactMatches[exactMatches.length - 1];
      if (duplicateMode === 'merge') return Object.assign({}, ...[...exactMatches].reverse());
      if (duplicateMode === 'popup' && exactMatches.length > 1) return chooseFromMatches(exactMatches, 'Multiple exact records found in main Excel.');
      return exactMatches[0];
    }

    // Partial, case-insensitive match. Ask user before updating from main Excel.
    const partialMatches = [];
    for (const [indexedKey, records] of lookupIndex.entries()) {
      if (indexedKey.includes(exactKey) || exactKey.includes(indexedKey)) partialMatches.push(...records);
      if (partialMatches.length >= 10) break;
    }
    if (partialMatches.length) {
      return chooseFromMatches(partialMatches, 'No exact match found, but partial matching records are available in main Excel. Do you want to update from one of these?');
    }
    return null;
  }, [lookupIndex, duplicateMode, chooseFromMatches]);

  function applyAutofill(rowIndex, force = false, explicitKey = undefined) {
    if (!lookupWorkingCol) return;
    const key = explicitKey !== undefined ? explicitKey : rows[rowIndex]?.[lookupWorkingCol.id];
    if (!normalize(key)) return;
    const match = getMatch(key);
    setRows((prev) => {
      const next = [...prev];
      const row = { ...(next[rowIndex] || blankRow(workingColumns)) };
      row[lookupWorkingCol.id] = key;
      if (!match) {
        next[rowIndex] = row;
        setRowStatus((s) => ({ ...s, [rowIndex]: 'not_found' }));
        return next;
      }
      row[lookupWorkingCol.id] = match[lookupColumn] ?? key;
      for (const col of workingColumns) {
        if (col.isCustom || !col.sourceColumn || col.id === lookupWorkingCol.id) continue;
        const metaKey = `${rowIndex}:${col.id}`;
        const manual = cellMeta[metaKey]?.manuallyEdited;
        if (!force && manual && overwriteMode === 'protect') continue;
        if (!force && manual && overwriteMode === 'confirm' && !confirm(`Overwrite manually edited value in ${col.displayName}?`)) continue;
        if (!force && manual && overwriteMode === 'refresh') continue;
        row[col.id] = match[col.sourceColumn] ?? '';
      }
      next[rowIndex] = row;
      setRowStatus((s) => ({ ...s, [rowIndex]: 'matched' }));
      return next;
    });
    setCellMeta((prev) => {
      const next = { ...prev };
      for (const col of workingColumns) {
        if (!col.isCustom && col.sourceColumn && col.id !== lookupWorkingCol.id) {
          next[`${rowIndex}:${col.id}`] = { autoFilled: true, manuallyEdited: false };
        }
      }
      return next;
    });
  }

  function syncRowsFromHot() {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    setRows(hot.getSourceData().map((r) => ({ ...blankRow(workingColumns), ...r })));
  }

  function afterChange(changes, source) {
    if (!changes || source === 'loadData' || source === 'autofill-internal') return;
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    const updated = hot.getSourceData().map((r) => ({ ...blankRow(workingColumns), ...r }));
    setRows(updated);
    const lookupProp = lookupWorkingCol?.id;
    const newMeta = {};
    const lookupRows = [];
    for (const [r, prop, oldVal, newVal] of changes) {
      if (oldVal === newVal) continue;
      if (prop === lookupProp) lookupRows.push([r, newVal]);
      else newMeta[`${r}:${prop}`] = { ...(cellMeta[`${r}:${prop}`] || {}), manuallyEdited: true };
    }
    if (Object.keys(newMeta).length) setCellMeta((m) => ({ ...m, ...newMeta }));
    setTimeout(() => lookupRows.forEach(([r, value]) => applyAutofill(r, false, value)), 0);
  }

  function addRows(count = 10) {
    setRows((prev) => [...prev, ...createInitialRows(workingColumns, count)]);
  }

  function selectedRange() {
    const hot = hotRef.current?.hotInstance;
    const sel = hot?.getSelectedLast();
    if (!sel) return null;
    return {
      r1: Math.max(0, Math.min(sel[0], sel[2])),
      r2: Math.max(0, Math.max(sel[0], sel[2])),
      c1: Math.max(0, Math.min(sel[1], sel[3])),
      c2: Math.max(0, Math.max(sel[1], sel[3])),
    };
  }

  function parseClipboardText(text) {
    return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter((line, idx, arr) => !(idx === arr.length - 1 && line === '')).map((line) => line.split('\t'));
  }

  function pasteMatrix(matrix, startRow = 0, startCol = 0) {
    if (!matrix?.length) return;
    const existingIds = new Set(workingColumns.map((c) => c.id));
    const nextColumns = [...workingColumns];
    let addedColumns = [];
    const maxCols = Math.max(...matrix.map((r) => r.length));
    for (let c = 0; c < maxCols; c++) {
      if (!nextColumns[startCol + c]) {
        const displayName = `Pasted Column ${nextColumns.length + 1}`;
        const col = { id: makeColumnId(displayName, existingIds), displayName, sourceColumn: null, isCustom: true };
        nextColumns.push(col);
        addedColumns.push(col);
      }
    }
    const nextRows = rows.map((r) => ({ ...blankRow(nextColumns), ...r }));
    while (nextRows.length < startRow + matrix.length) nextRows.push(blankRow(nextColumns));
    const lookupRows = [];
    const lookupColIndex = nextColumns.findIndex((c) => c.sourceColumn === lookupColumn);
    for (let r = 0; r < matrix.length; r++) {
      for (let c = 0; c < matrix[r].length; c++) {
        const col = nextColumns[startCol + c];
        nextRows[startRow + r][col.id] = matrix[r][c];
        if (startCol + c === lookupColIndex) lookupRows.push([startRow + r, matrix[r][c]]);
      }
    }
    if (addedColumns.length) {
      setWorkingColumns(nextColumns);
      setCustomColumns((prev) => [...prev, ...addedColumns]);
    }
    setRows(nextRows);
    setTimeout(() => lookupRows.forEach(([r, value]) => applyAutofill(r, false, value)), 0);
  }

  async function pasteFromClipboardButton() {
    try {
      const text = await navigator.clipboard.readText();
      const range = selectedRange() || { r1: 0, c1: 0 };
      pasteMatrix(parseClipboardText(text), range.r1, range.c1);
    } catch (_) {
      alert('Clipboard permission blocked. Click a cell in the live sheet and press Ctrl+V / Cmd+V.');
    }
  }

  function deleteSelectedRows() {
    const range = selectedRange();
    if (!range) return alert('Please select row(s) first.');
    setRows((prev) => prev.filter((_, i) => i < range.r1 || i > range.r2));
    setRowStatus({});
    hotRef.current?.hotInstance?.deselectCell();
  }

  function clearSelectedCells() {
    const range = selectedRange();
    if (!range) return alert('Please select cell(s) first.');
    setRows((prev) => prev.map((row, ri) => {
      if (ri < range.r1 || ri > range.r2) return row;
      const next = { ...row };
      for (let ci = range.c1; ci <= range.c2; ci++) {
        const col = workingColumns[ci];
        if (col) next[col.id] = '';
      }
      return next;
    }));
  }

  function clearSelectedColumn() {
    const range = selectedRange();
    if (!range) return alert('Please select a column/cell first.');
    setRows((prev) => prev.map((row) => {
      const next = { ...row };
      for (let ci = range.c1; ci <= range.c2; ci++) {
        const col = workingColumns[ci];
        if (col) next[col.id] = '';
      }
      return next;
    }));
  }

  function deleteSelectedColumns() {
    const range = selectedRange();
    if (!range) return alert('Please select column(s) or cell(s) first.');
    if (workingColumns.length <= (range.c2 - range.c1 + 1)) return alert('At least one column must remain.');
    const removeIds = new Set(workingColumns.slice(range.c1, range.c2 + 1).map((c) => c.id));
    const removedLookup = workingColumns.slice(range.c1, range.c2 + 1).some((c) => c.sourceColumn === lookupColumn);
    const nextColumns = workingColumns.filter((c) => !removeIds.has(c.id));
    setWorkingColumns(nextColumns);
    setSelected((prev) => prev.filter((c) => !removeIds.has(c.id)));
    setCustomColumns((prev) => prev.filter((c) => !removeIds.has(c.id)));
    setRows((prev) => prev.map((row) => {
      const next = {};
      for (const col of nextColumns) next[col.id] = row[col.id] ?? '';
      return next;
    }));
    if (removedLookup) {
      const fallback = nextColumns.find((c) => c.sourceColumn)?.sourceColumn || master?.columns?.[0]?.originalName || '';
      setLookupColumn(fallback);
      setNotice('Lookup column was deleted, so a new available lookup column was selected.');
    }
  }

  function moveSelectedColumn(direction) {
    const range = selectedRange();
    if (!range) return alert('Please select a column/cell first.');
    const from = range.c1;
    const to = from + direction;
    if (to < 0 || to >= workingColumns.length) return;
    setWorkingColumns((prev) => {
      const next = [...prev];
      [next[from], next[to]] = [next[to], next[from]];
      return next;
    });
    setTimeout(() => hotRef.current?.hotInstance?.selectCell(range.r1, to), 0);
  }

  function syncColumnMove(movedColumns, finalIndex) {
    if (!movedColumns?.length) return;
    setWorkingColumns((prev) => {
      const moving = movedColumns.map((i) => prev[i]);
      let remaining = prev.filter((_, i) => !movedColumns.includes(i));
      const insertAt = Math.max(0, Math.min(finalIndex, remaining.length));
      remaining.splice(insertAt, 0, ...moving);
      return remaining;
    });
  }

  function undoGrid() {
    const hot = hotRef.current?.hotInstance;
    hot?.undo();
    setTimeout(syncRowsFromHot, 0);
  }

  function redoGrid() {
    const hot = hotRef.current?.hotInstance;
    hot?.redo();
    setTimeout(syncRowsFromHot, 0);
  }

  function refreshSelectedRow() {
    const range = selectedRange();
    if (!range) return alert('Please select a row first.');
    for (let r = range.r1; r <= range.r2; r++) applyAutofill(r, true);
  }

  function refreshAllRows() {
    rows.forEach((row, idx) => {
      const value = lookupWorkingCol ? row[lookupWorkingCol.id] : '';
      if (normalize(value)) applyAutofill(idx, true, value);
    });
  }

  function addLiveCustomColumn() {
    const col = createCustomColumn('New Column');
    if (!col) return;
    setWorkingColumns((prev) => [...prev, col]);
    setCustomColumns((prev) => [...prev, col]);
    setRows((prev) => prev.map((r) => ({ ...r, [col.id]: '' })));
  }

  function addLiveMasterColumn(sourceColumnName) {
    if (!sourceColumnName) return;
    const masterCol = master?.columns?.find((c) => c.originalName === sourceColumnName);
    if (!masterCol) return;
    if (workingColumns.some((c) => c.sourceColumn === sourceColumnName)) return alert('This master column already exists in the live sheet.');
    const existingIds = new Set(workingColumns.map((c) => c.id));
    const col = { id: makeColumnId(masterCol.displayName, existingIds), displayName: masterCol.displayName, sourceColumn: masterCol.originalName, isCustom: false };
    setWorkingColumns((prev) => [...prev, col]);
    setSelected((prev) => [...prev, col]);
    setRows((prev) => prev.map((r) => {
      const next = { ...r, [col.id]: '' };
      const currentLookupValue = lookupWorkingCol ? r[lookupWorkingCol.id] : '';
      const exactMatches = lookupIndex.get(lookupKey(currentLookupValue)) || [];
      const match = duplicateMode === 'latest' ? exactMatches[exactMatches.length - 1] : exactMatches[0];
      if (match) next[col.id] = match[sourceColumnName] ?? '';
      return next;
    }));
    setNotice(`Added master column: ${masterCol.displayName}. Existing rows were filled from main Excel where exact lookup matches were found.`);
  }

  function changeLookupInWork(newLookup) {
    setLookupColumn(newLookup);
    const alreadyInSheet = workingColumns.some((c) => c.sourceColumn === newLookup);
    if (!alreadyInSheet && confirm('This lookup column is not in the live sheet. Add it now?')) {
      addLiveMasterColumn(newLookup);
    }
    setNotice('Lookup column changed. Matching is case-insensitive. Existing rows can be updated using Refresh All Rows.');
  }

  async function exportFile(kind) {
    const payload = { filename, columns: workingColumns, rows };
    try { await downloadExport(kind, payload); }
    catch (e) { setError(e.message); }
  }

  async function importGenerated(file) {
    if (!file) return;
    setBusy(true); setError('');
    try {
      const data = await importWorkingSheet(file, 0);
      const existingIds = new Set();
      const cols = data.columns.map((name) => ({ id: makeColumnId(name, existingIds), displayName: name, sourceColumn: master?.columns?.find((c) => c.originalName === name)?.originalName || null, isCustom: !master?.columns?.some((c) => c.originalName === name) }));
      setWorkingColumns(cols);
      setSelected(cols.filter((c) => !c.isCustom));
      setCustomColumns(cols.filter((c) => c.isCustom));
      setRows(data.rows.map((arr) => Object.fromEntries(cols.map((c, i) => [c.id, arr[i] ?? '']))));
      setStep(3);
      setNotice(`Imported ${data.rowCount} rows.`);
    } catch (e) { setError(e.message); }
    finally { setBusy(false); }
  }

  async function resetAll() {
    if (!confirm('Clear local session and start over?')) return;
    await clearSession();
    location.reload();
  }

  useEffect(() => {
    if (step !== 3) return;
    const onPaste = (event) => {
      const active = document.activeElement;
      const grid = document.querySelector('.hot-wrap');
      const isFormField = ['INPUT', 'TEXTAREA', 'SELECT'].includes(active?.tagName || '');
      const gridIsActive = grid && (grid.contains(active) || active?.closest?.('.handsontable') || hotRef.current?.hotInstance?.getSelectedLast());
      if (!gridIsActive || (isFormField && !active?.closest?.('.handsontable'))) return;
      const text = event.clipboardData?.getData('text/plain');
      if (!text) return;
      const matrix = parseClipboardText(text);
      if (!matrix.length) return;
      const range = selectedRange() || { r1: 0, c1: 0 };
      event.preventDefault();
      pasteMatrix(matrix, range.r1, range.c1);
    };
    document.addEventListener('paste', onPaste, true);
    return () => document.removeEventListener('paste', onPaste, true);
  }, [step, rows, workingColumns, lookupColumn, lookupIndex]);

  const contextMenu = useMemo(() => ({
    items: {
      row_above: {},
      row_below: {},
      remove_row: { name: 'Delete row(s)' },
      hsep1: '---------',
      clear_cells: { name: 'Clear selected cell(s)', callback: clearSelectedCells },
      clear_column: { name: 'Clear selected column(s)', callback: clearSelectedColumn },
      delete_column: { name: 'Delete selected column(s)', callback: deleteSelectedColumns },
      move_left: { name: 'Move column left', callback: () => moveSelectedColumn(-1) },
      move_right: { name: 'Move column right', callback: () => moveSelectedColumn(1) },
      add_custom_column: { name: 'Insert custom column', callback: addLiveCustomColumn },
      hsep2: '---------',
      copy: {},
      cut: {},
      paste_clipboard: { name: 'Paste from clipboard', callback: pasteFromClipboardButton },
      undo: { name: 'Undo', callback: undoGrid },
      redo: { name: 'Redo', callback: redoGrid },
    }
  }), [workingColumns, rows]);

  return <div className="app">
    <header className="topbar">
      <div><h1>Dynamic Excel Builder</h1><p>Upload master data, build custom sheets, live auto-fill, export Excel/PDF.</p></div>
      <div className="top-actions"><span className="saved">{savedAt ? `Auto-saved ${savedAt}` : 'Local IndexedDB session'}</span><button onClick={resetAll} className="ghost">Start Over</button></div>
    </header>

    <nav className="steps"><span className={step===1?'active':''}>1 Upload</span><span className={step===2?'active':''}>2 Configure</span><span className={step===3?'active':''}>3 Work & Export</span></nav>
    {error && <div className="alert error"><AlertTriangle size={18}/>{error}</div>}
    {notice && <div className="alert ok">{notice}</div>}
    {busy && <div className="alert busy">Processing...</div>}

    {step === 1 && <section className="card upload-card">
      <FileSpreadsheet size={42}/><h2>Upload Master Excel File</h2>
      <p>Supported: .xlsx and .xls. Empty rows are ignored and duplicate headers are renamed automatically.</p>
      <label>Header row number <input type="number" min="1" value={headerRow + 1} onChange={(e)=>setHeaderRow(Math.max(0, Number(e.target.value)-1))}/></label>
      <input id="masterFileInput" type="file" accept=".xlsx,.xls" onChange={(e)=>handleUpload(e.target.files?.[0])}/>
      <div className="mini">If headers are on row 3, enter 3 before uploading.</div>
    </section>}

    {step === 2 && master && <section className="grid-layout">
      <div className="card"><h2>Master Columns</h2><p>{master.rowCount} rows · {master.columnCount} columns</p>
        <button onClick={() => document.getElementById('masterFileInput2')?.click()} className="ghost"><Upload size={15}/> Re-upload/Re-parse</button>
        <input id="masterFileInput2" hidden type="file" accept=".xlsx,.xls" onChange={(e)=>handleUpload(e.target.files?.[0])}/>
        <h3>Lookup Column</h3><select value={lookupColumn} onChange={(e)=>setLookupColumn(e.target.value)}>{master.columns.map((c)=><option key={c.originalName} value={c.originalName}>{c.displayName}</option>)}</select>
        {duplicateLookupCount > 0 && <div className="warn">Warning: {duplicateLookupCount} duplicate lookup value(s) found. Matching is case-insensitive.</div>}
        <label>Duplicate handling<select value={duplicateMode} onChange={(e)=>setDuplicateMode(e.target.value)}><option value="first">Use first match</option><option value="latest">Use latest match</option><option value="popup">Show selection popup</option><option value="merge">Merge records</option></select></label>
        <label>Overwrite behavior<select value={overwriteMode} onChange={(e)=>setOverwriteMode(e.target.value)}><option value="protect">Protect manual edits</option><option value="confirm">Confirm overwrite</option><option value="always">Always overwrite</option><option value="refresh">Only on manual refresh</option></select></label>
      </div>
      <div className="card wide"><h2>Selected Existing Columns</h2>
        <div className="column-list">{selected.map((c,i)=><div className="col-row" key={c.id}><input value={c.displayName} onChange={(e)=>updateSelected(c.id,{displayName:e.target.value})}/><span className="source">from {c.sourceColumn}</span><button onClick={()=>move(setSelected,i,-1)}>↑</button><button onClick={()=>move(setSelected,i,1)}>↓</button><button onClick={()=>setSelected((p)=>p.filter((x)=>x.id!==c.id))}><Trash2 size={14}/></button></div>)}</div>
        <details><summary>Add removed master columns</summary>{master.columns.filter((m)=>!selected.some((s)=>s.sourceColumn===m.originalName)).map((m)=><button key={m.originalName} className="pill" onClick={()=>setSelected((p)=>[...p,{id:makeColumnId(m.displayName,new Set(p.map(x=>x.id))),displayName:m.displayName,sourceColumn:m.originalName,isCustom:false}])}>{m.displayName}</button>)}</details>
        <h2>Custom Columns</h2><button onClick={addCustomColumnToConfig}><Plus size={16}/> Add Custom Column</button>
        <div className="column-list">{customColumns.map((c,i)=><div className="col-row" key={c.id}><input value={c.displayName} onChange={(e)=>updateCustom(c.id,{displayName:e.target.value})}/><span className="source">custom</span><button onClick={()=>move(setCustomColumns,i,-1)}>↑</button><button onClick={()=>move(setCustomColumns,i,1)}>↓</button><button onClick={()=>setCustomColumns((p)=>p.filter((x)=>x.id!==c.id))}><Trash2 size={14}/></button></div>)}</div>
        <button className="primary" onClick={generateSheet}>{rows.length ? 'Update Live Sheet Without Losing Data' : 'Generate Live Working Sheet'}</button>
      </div>
    </section>}

    {step === 3 && <section className="workbench">
      <div className="toolbar card">
        <button onClick={goBackToConfigure} className="ghost">Back to Configuration</button>
        <label className="inline-label">Lookup<select value={lookupColumn} onChange={(e)=>changeLookupInWork(e.target.value)}>{master?.columns?.map((c)=><option key={c.originalName} value={c.originalName}>{c.displayName}</option>)}</select></label>
        <button onClick={()=>addRows(10)}><Plus size={16}/> Add 10 Rows</button>
        <button onClick={addLiveCustomColumn}><Plus size={16}/> Add Column</button>
        <select onChange={(e)=>{ addLiveMasterColumn(e.target.value); e.target.value=''; }} defaultValue=""><option value="">Add removed master column...</option>{master?.columns?.filter((m)=>!workingColumns.some((c)=>c.sourceColumn===m.originalName)).map((m)=><option key={m.originalName} value={m.originalName}>{m.displayName}</option>)}</select>
        <button onClick={pasteFromClipboardButton}>Paste Clipboard</button>
        <button onClick={undoGrid}>Undo</button>
        <button onClick={redoGrid}>Redo</button>
        <button onClick={deleteSelectedRows}><Trash2 size={16}/> Delete Rows</button>
        <button onClick={deleteSelectedColumns}><Trash2 size={16}/> Delete Columns</button>
        <button onClick={() => moveSelectedColumn(-1)}>← Move Column</button>
        <button onClick={() => moveSelectedColumn(1)}>Move Column →</button>
        <button onClick={clearSelectedCells}>Clear Cells</button>
        <button onClick={clearSelectedColumn}>Clear Column</button>
        <button onClick={refreshSelectedRow}><RefreshCw size={16}/> Refresh Row(s)</button>
        <button onClick={refreshAllRows}><RefreshCw size={16}/> Refresh All Rows</button>
        <input value={filename} onChange={(e)=>setFilename(e.target.value)} placeholder="Export filename"/>
        <button onClick={()=>exportFile('excel')}><Download size={16}/> Excel</button>
        <button onClick={()=>exportFile('pdf')}><Download size={16}/> PDF</button>
        <label className="import-button"><Upload size={16}/> Import Sheet<input type="file" hidden accept=".xlsx,.xls" onChange={(e)=>importGenerated(e.target.files?.[0])}/></label>
      </div>
      <div className="statusbar">Lookup: <b>{lookupColumn}</b> · Case-insensitive matching enabled · Rows: {rows.length} · Matched: {Object.values(rowStatus).filter(s=>s==='matched').length} · Not found: {Object.values(rowStatus).filter(s=>s==='not_found').length}</div>
      <div className="hot-wrap"><HotTable ref={hotRef} data={rows} columns={gridColumns} colHeaders={gridHeaders} rowHeaders={true} width="100%" height="68vh" licenseKey="non-commercial-and-evaluation" stretchH="all" manualColumnResize={true} manualColumnMove={true} manualRowMove={true} dropdownMenu={true} filters={true} contextMenu={contextMenu} afterColumnMove={(movedColumns, finalIndex) => syncColumnMove(movedColumns, finalIndex)} multiColumnSorting={true} copyPaste={{ rowsLimit: 100000, columnsLimit: 1000, pasteMode: 'overwrite' }} undo={true} outsideClickDeselects={false} afterChange={afterChange} afterCreateRow={() => setTimeout(syncRowsFromHot, 0)} afterRemoveRow={() => setTimeout(syncRowsFromHot, 0)} afterUndo={() => setTimeout(syncRowsFromHot, 0)} afterRedo={() => setTimeout(syncRowsFromHot, 0)} cells={(row)=>({ className: rowStatus[row] === 'not_found' ? 'not-found-row' : rowStatus[row] === 'matched' ? 'matched-row' : '' })}/></div>
      <p className="mini">Type a value into the lookup column. Matching is instant and case-insensitive. If only a partial match is found, the app asks before filling from main Excel. Custom columns are never overwritten.</p>
    </section>}
  </div>;
}

export default App;
