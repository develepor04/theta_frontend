// Grid <-> Univer IWorkbookData conversion for Theta Sheets.
// Grid shape: { sheets: [{ name, headers, rows }] }
// Univer IWorkbookData shape: { id, name, sheetOrder, sheets: { [id]: {...} }, styles }

import { hasScheduleHeaders } from './thetaValidation';
import { DEFAULT_DATE_FORMAT, getActiveDateFormat } from './dateFormats';

// Full schedule/cost/productivity column schema, matching the business-logic
// reference workbook (Dashboard KPIs / Schedule Intelligence / Cost
// Intelligence formulas). Only 'Activity ID' and 'Activity Name' are actually
// required for a row to count as an activity -- the rest are optional and
// used by compute_metrics_from_sheet() when present.
export const REQUIRED_SHEET_COLUMNS = [
  'Activity ID',
  'Activity Name',
  'Phase',
  'Cost Category',
  'Period',
  'Baseline Start',
  'Baseline Finish',
  'Forecast Start',
  'Forecast Finish',
  'Actual Start',
  'Actual Finish',
  '% Complete',
  'Status',
  'Variance (Days)',
  'Root Cause',
  'Impact',
  'Budget Cost (AED)',
  'Actual Cost (AED)',
  'Forecast Cost (AED)',
  'Planned Hours',
  'Actual Hours',
  'Planned Output',
  'Actual Output',
  'Output Unit',
  'Productivity Index',
];

const HEADER_STYLE_ID = 'theta-header-style';
const DATE_STYLE_ID = 'theta-date-style';

/** Default Univer / Excel number format for schedule date columns. */
export const DATE_NUMBER_FORMAT = DEFAULT_DATE_FORMAT;

const KNOWN_DATE_HEADERS = new Set(
  REQUIRED_SHEET_COLUMNS
    .filter((h) => /start|finish|end|date/i.test(h))
    .map((h) => h.toLowerCase()),
);

/**
 * True when a header label is a date column (schedule start/finish/end/date fields).
 * @param {unknown} label
 */
export function isDateColumnLabel(label) {
  const t = String(label ?? '').trim().toLowerCase();
  if (!t) return false;
  if (KNOWN_DATE_HEADERS.has(t)) return true;
  return /start|finish|end|date/.test(t);
}

/**
 * @param {unknown[]} headers
 * @returns {number[]}
 */
export function getDateColumnIndexes(headers = []) {
  const indexes = [];
  (headers || []).forEach((h, i) => {
    if (isDateColumnLabel(h)) indexes.push(i);
  });
  return indexes;
}

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

/**
 * Convert ISO / Date / serial-like values to an Excel serial day number.
 * Number formats only affect numeric dates — string "2024-02-08" will not reformat.
 * @param {unknown} value
 * @returns {number | string | ''}
 */
export function toExcelSerialDate(value) {
  if (value === undefined || value === null || value === '') return '';
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const utc = Date.UTC(value.getFullYear(), value.getMonth(), value.getDate());
    return Math.round((utc - EXCEL_EPOCH_UTC) / 86400000);
  }

  const s = String(value).trim();
  if (!s) return '';

  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (iso) {
    const utc = Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return Math.round((utc - EXCEL_EPOCH_UTC) / 86400000);
  }

  const asNum = Number(s);
  if (Number.isFinite(asNum) && asNum > 20000 && asNum < 80000 && String(asNum) === s) {
    return asNum;
  }

  const parsed = new Date(s);
  if (!Number.isNaN(parsed.getTime())) {
    const utc = Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
    return Math.round((utc - EXCEL_EPOCH_UTC) / 86400000);
  }

  return value;
}

/**
 * Convert an Excel serial (or leave ISO strings alone) back to `yyyy-MM-dd` for API/grid storage.
 * @param {unknown} value
 * @returns {unknown}
 */
export function excelSerialToIsoDate(value) {
  if (value === undefined || value === null || value === '') return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatUtcYmd(value);
  }
  if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 80000) {
    return formatUtcYmd(new Date(EXCEL_EPOCH_UTC + Math.round(value) * 86400000));
  }
  const s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const asNum = Number(s);
  if (Number.isFinite(asNum) && asNum > 20000 && asNum < 80000 && String(asNum) === s) {
    return formatUtcYmd(new Date(EXCEL_EPOCH_UTC + Math.round(asNum) * 86400000));
  }
  return value;
}

function formatUtcYmd(d) {
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/**
 * Normalize date-column values in a grid sheet to ISO strings (for save/API).
 * Preserves optional `columnFormats` (header → numfmt pattern).
 * @param {{ name?: string, headers?: unknown[], rows?: unknown[][], columnFormats?: Record<string, string> }} sheet
 */
function normalizeDateColumnsInSheet(sheet) {
  const headers = sheet?.headers || [];
  const dateCols = getDateColumnIndexes(headers);
  if (!dateCols.length) return sheet;
  const rows = (sheet.rows || []).map((row) => {
    const next = Array.isArray(row) ? [...row] : [];
    for (const c of dateCols) {
      if (c < next.length) next[c] = excelSerialToIsoDate(next[c]);
    }
    return next;
  });
  return { ...sheet, rows };
}

/**
 * Read a cell's number-format pattern via facade helpers.
 * @param {object} fWorksheet
 * @param {number} row
 * @param {number} col
 * @returns {string}
 */
function readCellNumberFormat(fWorksheet, row, col) {
  try {
    const range = fWorksheet.getRange(row, col, 1, 1);
    let pattern = String(range.getNumberFormat?.() ?? '').trim();
    if (pattern) return pattern;
    // Fallbacks: some builds expose pattern on style data (`n`) rather than numberFormat.
    const styleData = range.getCellStyleData?.();
    pattern = String(styleData?.n?.pattern ?? styleData?.numberFormat?.pattern ?? '').trim();
    if (pattern) return pattern;
    const style = range.getCellStyle?.();
    pattern = String(style?.numberFormat?.pattern ?? style?.n?.pattern ?? '').trim();
    return pattern;
  } catch {
    return '';
  }
}

/**
 * Read per-column number formats from a live worksheet.
 * Uses the most common non-General pattern in the column (not only row 1),
 * so Formats applied to a selection still persist on save/reload.
 * Keys are header labels so formats survive column reorder within reason.
 * @param {object | null | undefined} fWorksheet
 * @returns {Record<string, string>}
 */
export function extractColumnFormats(fWorksheet) {
  /** @type {Record<string, string>} */
  const formats = {};
  if (!fWorksheet?.getRange) return formats;

  let lastCol = -1;
  let lastRow = -1;
  try {
    lastCol = typeof fWorksheet.getLastColumn === 'function' ? fWorksheet.getLastColumn() : -1;
    lastRow = typeof fWorksheet.getLastRow === 'function' ? fWorksheet.getLastRow() : -1;
  } catch {
    return formats;
  }
  if (!Number.isFinite(lastCol) || lastCol < 0) return formats;
  if (!Number.isFinite(lastRow) || lastRow < 1) lastRow = 1;

  // Cap scan for large sheets; Formats are almost always uniform per column.
  const scanEnd = Math.min(lastRow, 200);

  for (let c = 0; c <= lastCol; c += 1) {
    let header = '';
    try {
      header = String(fWorksheet.getRange(0, c, 1, 1).getValue?.() ?? '').trim();
    } catch {
      header = '';
    }
    if (!header) continue;

    /** @type {Map<string, number>} */
    const counts = new Map();
    for (let r = 1; r <= scanEnd; r += 1) {
      const pattern = readCellNumberFormat(fWorksheet, r, c);
      if (!pattern || /^general$/i.test(pattern)) continue;
      counts.set(pattern, (counts.get(pattern) || 0) + 1);
    }

    let best = '';
    let bestCount = 0;
    for (const [pattern, count] of counts) {
      if (count > bestCount) {
        best = pattern;
        bestCount = count;
      }
    }
    if (best) formats[header] = best;
  }
  return formats;
}

/**
 * Apply saved `columnFormats` (header → pattern) onto a live workbook.
 * @param {object | null | undefined} fWorkbook
 * @param {{ sheets?: Array<{ name?: string, headers?: unknown[], columnFormats?: Record<string, string> }> }} grid
 */
export function applyColumnFormatsFromGrid(fWorkbook, grid) {
  if (!fWorkbook || !grid?.sheets?.length) return;
  for (const sheetData of grid.sheets) {
    const formats = sheetData?.columnFormats;
    if (!formats || typeof formats !== 'object' || !Object.keys(formats).length) continue;
    const name = String(sheetData.name || '').trim();
    const fSheet = (name && fWorkbook.getSheetByName?.(name))
      || fWorkbook.getActiveSheet?.();
    if (!fSheet) continue;
    applyColumnFormatsToSheet(fSheet, formats, sheetData.headers);
  }
}

/**
 * @param {object | null | undefined} fWorksheet
 * @param {Record<string, string>} columnFormats
 * @param {unknown[]} [headersHint]
 */
export function applyColumnFormatsToSheet(fWorksheet, columnFormats, headersHint) {
  if (!fWorksheet?.getRange || !columnFormats) return;

  let lastCol = 0;
  let rowCount = 50;
  try {
    lastCol = typeof fWorksheet.getLastColumn === 'function' ? fWorksheet.getLastColumn() : 0;
    if (typeof fWorksheet.getMaxRows === 'function') rowCount = fWorksheet.getMaxRows();
    else if (typeof fWorksheet.getRowCount === 'function') rowCount = fWorksheet.getRowCount();
    else if (typeof fWorksheet.getLastRow === 'function') {
      rowCount = Math.max(fWorksheet.getLastRow() + 20, 50);
    }
  } catch {
    return;
  }
  if (!Number.isFinite(lastCol) || lastCol < 0) return;
  if (!Number.isFinite(rowCount) || rowCount < 2) rowCount = 50;

  /** @type {string[]} */
  const headers = [];
  if (Array.isArray(headersHint) && headersHint.length) {
    for (let c = 0; c <= lastCol; c += 1) {
      headers.push(String(headersHint[c] ?? '').trim());
    }
  } else {
    for (let c = 0; c <= lastCol; c += 1) {
      try {
        headers.push(String(fWorksheet.getRange(0, c, 1, 1).getValue?.() ?? '').trim());
      } catch {
        headers.push('');
      }
    }
  }

  for (let c = 0; c <= lastCol; c += 1) {
    const header = headers[c] || '';
    const fmt = String(
      (header && columnFormats[header])
      || columnFormats[String(c)]
      || columnFormats[c]
      || '',
    ).trim();
    if (!fmt || /^general$/i.test(fmt)) continue;
    try {
      const range = fWorksheet.getRange(1, c, Math.max(rowCount - 1, 1), 1);
      if (typeof range.setNumberFormat === 'function') {
        range.setNumberFormat(fmt);
      }
    } catch {
      // skip column
    }
  }
}

/**
 * Coerce string dates in live date columns to Excel serials so Formats can change display.
 * @param {object | null | undefined} fWorkbook
 */
export function coerceDateColumnValuesInWorkbook(fWorkbook) {
  if (!fWorkbook) return;
  const sheets = typeof fWorkbook.getSheets === 'function'
    ? fWorkbook.getSheets()
    : [];
  const list = Array.isArray(sheets) && sheets.length
    ? sheets
    : [fWorkbook.getActiveSheet?.()].filter(Boolean);

  for (const sheet of list) {
    coerceDateColumnValuesInSheet(sheet);
  }
}

/**
 * @param {object | null | undefined} fWorksheet
 */
export function coerceDateColumnValuesInSheet(fWorksheet) {
  if (!fWorksheet?.getRange) return;

  let lastCol = 0;
  let lastRow = 0;
  try {
    lastCol = typeof fWorksheet.getLastColumn === 'function' ? fWorksheet.getLastColumn() : 0;
    lastRow = typeof fWorksheet.getLastRow === 'function' ? fWorksheet.getLastRow() : 0;
  } catch {
    return;
  }
  if (!Number.isFinite(lastCol) || lastCol < 0 || !Number.isFinite(lastRow) || lastRow < 1) return;

  /** @type {string[]} */
  const headers = [];
  for (let c = 0; c <= lastCol; c += 1) {
    try {
      headers.push(String(fWorksheet.getRange(0, c, 1, 1).getValue?.() ?? '').trim());
    } catch {
      headers.push('');
    }
  }
  const dateCols = getDateColumnIndexes(headers);
  if (!dateCols.length) return;

  for (const col of dateCols) {
    for (let r = 1; r <= lastRow; r += 1) {
      try {
        const range = fWorksheet.getRange(r, col, 1, 1);
        const raw = range.getValue?.();
        if (raw === undefined || raw === null || raw === '') continue;
        if (typeof raw === 'number' && Number.isFinite(raw)) continue;
        const serial = toExcelSerialDate(raw);
        if (typeof serial !== 'number' || serial === raw) continue;
        if (typeof range.setValue === 'function') range.setValue(serial);
        else range.setValues?.([[serial]]);
      } catch {
        // skip cell
      }
    }
  }
}

/**
 * Apply a date number format to date columns on every sheet in a live workbook.
 * @param {object | null | undefined} fWorkbook - Univer FWorkbook facade
 * @param {string} [pattern]
 */
export function applyDateNumberFormatsToWorkbook(fWorkbook, pattern) {
  if (!fWorkbook) return;
  const fmt = String(pattern || getActiveDateFormat() || DATE_NUMBER_FORMAT).trim();
  const sheets = typeof fWorkbook.getSheets === 'function'
    ? fWorkbook.getSheets()
    : [];
  const list = Array.isArray(sheets) && sheets.length
    ? sheets
    : [fWorkbook.getActiveSheet?.()].filter(Boolean);

  for (const sheet of list) {
    applyDateNumberFormatsToSheet(sheet, fmt);
  }
}

/**
 * @param {object | null | undefined} fWorksheet - Univer FWorksheet facade
 * @param {string} [pattern]
 */
export function applyDateNumberFormatsToSheet(fWorksheet, pattern) {
  if (!fWorksheet?.getRange) return;
  const fmt = String(pattern || getActiveDateFormat() || DATE_NUMBER_FORMAT).trim();

  let lastCol = 0;
  try {
    lastCol = typeof fWorksheet.getLastColumn === 'function'
      ? fWorksheet.getLastColumn()
      : 0;
  } catch {
    return;
  }
  if (!Number.isFinite(lastCol) || lastCol < 0) return;

  /** @type {string[]} */
  const headers = [];
  for (let c = 0; c <= lastCol; c += 1) {
    try {
      headers.push(String(fWorksheet.getRange(0, c, 1, 1).getValue?.() ?? '').trim());
    } catch {
      headers.push('');
    }
  }

  const dateCols = getDateColumnIndexes(headers);
  if (!dateCols.length) return;

  let rowCount = 50;
  try {
    if (typeof fWorksheet.getMaxRows === 'function') {
      rowCount = fWorksheet.getMaxRows();
    } else if (typeof fWorksheet.getRowCount === 'function') {
      rowCount = fWorksheet.getRowCount();
    } else if (typeof fWorksheet.getLastRow === 'function') {
      rowCount = Math.max(fWorksheet.getLastRow() + 20, 50);
    }
  } catch {
    rowCount = 50;
  }
  if (!Number.isFinite(rowCount) || rowCount < 2) rowCount = 50;

  for (const col of dateCols) {
    try {
      const range = fWorksheet.getRange(1, col, Math.max(rowCount - 1, 1), 1);
      if (typeof range.setNumberFormat === 'function') {
        range.setNumberFormat(fmt);
      }
    } catch {
      // skip column if facade rejects the range
    }
  }
}

/**
 * Apply format to active selection, else to all date columns on the active sheet.
 * @param {object} univerAPI
 * @param {string} pattern
 * @returns {{ ok: boolean, mode?: 'selection' | 'columns', error?: string }}
 */
export function applyDateFormatToSelectionOrColumns(univerAPI, pattern) {
  const fmt = String(pattern || '').trim();
  if (!fmt) return { ok: false, error: 'Empty format pattern.' };
  if (!univerAPI) return { ok: false, error: 'Spreadsheet is not ready.' };

  const workbook = univerAPI.getActiveWorkbook?.();
  const sheet = workbook?.getActiveSheet?.();
  if (!sheet?.getRange) return { ok: false, error: 'No active sheet.' };

  try {
    const active = typeof sheet.getActiveRange === 'function'
      ? sheet.getActiveRange()
      : null;
    if (active && typeof active.setNumberFormat === 'function') {
      active.setNumberFormat(fmt);
      return { ok: true, mode: 'selection' };
    }
  } catch {
    // fall through
  }

  try {
    applyDateNumberFormatsToSheet(sheet, fmt);
    return { ok: true, mode: 'columns' };
  } catch (err) {
    return { ok: false, error: err?.message || 'Could not apply format.' };
  }
}

export function blankGrid(name = 'Theta Sheets') {
  return { name, sheets: [{ name: 'Schedule', headers: REQUIRED_SHEET_COLUMNS, rows: [] }] };
}

export function toUniverWorkbookData({ sheets = [], name = 'Theta Sheets' } = {}) {
  const sheetList = sheets.length ? sheets : [{ name: 'Schedule', headers: REQUIRED_SHEET_COLUMNS, rows: [] }];
  const sheetOrder = [];
  const sheetsObj = {};
  /** @type {Record<string, object>} */
  const styles = {
    [HEADER_STYLE_ID]: { bl: 1, bg: { rgb: '#F1F5F9' } },
    [DATE_STYLE_ID]: { n: { pattern: getActiveDateFormat() || DATE_NUMBER_FORMAT } },
  };

  sheetList.forEach((s, idx) => {
    const sheetId = `sheet-${idx}`;
    sheetOrder.push(sheetId);
    const headers = s.headers || [];
    const rows = s.rows || [];
    const dateCols = new Set(getDateColumnIndexes(headers));
    const columnFormats = (s.columnFormats && typeof s.columnFormats === 'object')
      ? s.columnFormats
      : {};
    const cellData = {};

    // Per-column style ids so saved Formats survive reload (not one shared date style).
    /** @type {Record<number, string>} */
    const colStyleIds = {};
    headers.forEach((h, c) => {
      const header = String(h ?? '').trim();
      const saved = String(
        (header && columnFormats[header])
        || columnFormats[String(c)]
        || '',
      ).trim();
      if (saved && !/^general$/i.test(saved)) {
        const styleId = `theta-col-${idx}-${c}`;
        styles[styleId] = { n: { pattern: saved } };
        colStyleIds[c] = styleId;
      } else if (dateCols.has(c)) {
        colStyleIds[c] = DATE_STYLE_ID;
      }
    });

    cellData[0] = {};
    headers.forEach((h, c) => {
      cellData[0][c] = { v: h, s: HEADER_STYLE_ID };
    });

    rows.forEach((row, r) => {
      cellData[r + 1] = {};
      row.forEach((val, c) => {
        let v = val === undefined || val === null ? '' : val;
        // Date columns must be Excel serial numbers for Formats / numfmt to work.
        if (dateCols.has(c) && v !== '') {
          v = toExcelSerialDate(v);
        }
        const cell = { v };
        if (colStyleIds[c]) {
          cell.s = colStyleIds[c];
        }
        if (dateCols.has(c) && typeof v === 'number' && Number.isFinite(v)) {
          // CellValueType.NUMBER — required so numfmt treats the value as a date serial.
          cell.t = 2;
        }
        cellData[r + 1][c] = cell;
      });
    });

    sheetsObj[sheetId] = {
      id: sheetId,
      name: s.name || `Sheet${idx + 1}`,
      cellData,
      rowCount: Math.max(rows.length + 20, 50),
      columnCount: Math.max(headers.length + 5, 12),
    };
  });

  return {
    id: `wb-${Date.now()}`,
    name,
    sheetOrder,
    sheets: sheetsObj,
    styles,
  };
}

/** Pull a displayable value out of an Univer ICellData (or plain value). */
export function cellDisplayValue(cell) {
  if (cell === undefined || cell === null) return '';
  if (typeof cell !== 'object') return cell;
  if (cell.v !== undefined && cell.v !== null) return cell.v;
  const stream = cell.p?.body?.dataStream;
  if (typeof stream === 'string') {
    return stream.replace(/\r\n$/g, '').replace(/\0/g, '').replace(/\r/g, '');
  }
  return '';
}

function trimTrailingEmpty(values) {
  const next = [...values];
  while (next.length && String(next[next.length - 1] ?? '').trim() === '') next.pop();
  return next;
}

/** Prefer a row that carries Activity ID + Activity Name (aliases ok); else densest row. */
export function findHeaderRowIndex(matrix, scanRows = 15) {
  const limit = Math.min(scanRows, matrix?.length || 0);
  if (limit <= 0) return 0;

  for (let i = 0; i < limit; i++) {
    const cells = (matrix[i] || []).map((c) => String(c ?? '').trim());
    if (hasScheduleHeaders(cells)) return i;
  }

  let bestIdx = 0;
  let bestCount = -1;
  for (let i = 0; i < limit; i++) {
    const count = (matrix[i] || []).filter((c) => String(c ?? '').trim() !== '').length;
    if (count > bestCount) {
      bestCount = count;
      bestIdx = i;
    }
  }
  return bestIdx;
}

function matrixToSheet(name, matrix) {
  if (!matrix?.length) return { name, headers: [], rows: [] };
  const headerIdx = findHeaderRowIndex(matrix);
  const headers = trimTrailingEmpty(
    (matrix[headerIdx] || []).map((h) => String(h ?? '').trim()),
  );
  const width = headers.length;
  const rows = matrix
    .slice(headerIdx + 1)
    .map((row) => Array.from({ length: width }, (_, c) => {
      const val = row?.[c];
      return val === undefined || val === null ? '' : val;
    }))
    .filter((row) => row.some((v) => String(v).trim() !== ''));
  return { name, headers, rows };
}

function sheetFromCellData(name, cellData) {
  const rowKeys = Object.keys(cellData || {}).map(Number).filter((n) => Number.isFinite(n)).sort((a, b) => a - b);
  if (rowKeys.length === 0) return { name, headers: [], rows: [] };

  let maxCol = 0;
  rowKeys.forEach((r) => {
    Object.keys(cellData[r] || {}).forEach((c) => {
      const n = Number(c);
      if (Number.isFinite(n) && n > maxCol) maxCol = n;
    });
  });

  const matrix = rowKeys.map((r) => {
    const row = cellData[r] || {};
    return Array.from({ length: maxCol + 1 }, (_, c) => cellDisplayValue(row[c]));
  });

  // Preserve true row positions for header detection when row 0 was omitted.
  if (rowKeys[0] !== 0) {
    const padded = [];
    let cursor = 0;
    for (let r = 0; r <= rowKeys[rowKeys.length - 1]; r++) {
      if (cursor < rowKeys.length && rowKeys[cursor] === r) {
        padded.push(matrix[cursor]);
        cursor += 1;
      } else {
        padded.push(Array.from({ length: maxCol + 1 }, () => ''));
      }
    }
    return matrixToSheet(name, padded);
  }

  return matrixToSheet(name, matrix);
}

function sheetFromFacade(fSheet) {
  const name = fSheet?.getSheetName?.() || 'Sheet1';
  let lastRow = -1;
  let lastCol = -1;
  try {
    lastRow = typeof fSheet.getLastRow === 'function' ? fSheet.getLastRow() : -1;
    lastCol = typeof fSheet.getLastColumn === 'function' ? fSheet.getLastColumn() : -1;
  } catch {
    return { name, headers: [], rows: [] };
  }

  if (!Number.isFinite(lastRow) || !Number.isFinite(lastCol) || lastRow < 0 || lastCol < 0) {
    return { name, headers: [], rows: [] };
  }

  const rowCount = lastRow + 1;
  const colCount = lastCol + 1;
  let matrix = null;

  try {
    const values = fSheet.getRange?.(0, 0, rowCount, colCount)?.getValues?.();
    if (Array.isArray(values) && values.length) matrix = values;
  } catch {
    matrix = null;
  }

  if (!matrix) {
    matrix = [];
    for (let r = 0; r < rowCount; r++) {
      const row = [];
      for (let c = 0; c < colCount; c++) {
        try {
          row.push(fSheet.getRange?.(r, c, 1, 1)?.getValue?.() ?? '');
        } catch {
          row.push('');
        }
      }
      matrix.push(row);
    }
  }

  const base = matrixToSheet(name, matrix);
  const columnFormats = extractColumnFormats(fSheet);
  return normalizeDateColumnsInSheet({
    ...base,
    ...(Object.keys(columnFormats).length ? { columnFormats } : {}),
  });
}

/**
 * Preferred extractor: read via the Univer facade (what the user sees),
 * fall back to parsing workbook.save() cellData.
 */
export function extractGridFromUniverWorkbook(workbook, fallbackName = 'Theta Sheets') {
  if (!workbook) return { name: fallbackName, sheets: [] };

  try {
    const facadeSheets = workbook.getSheets?.();
    if (Array.isArray(facadeSheets) && facadeSheets.length > 0) {
      const sheets = facadeSheets.map((s) => sheetFromFacade(s));
      const hasAnyHeaders = sheets.some((s) => (s.headers || []).some((h) => String(h).trim()));
      if (hasAnyHeaders || sheets.some((s) => (s.rows || []).length > 0)) {
        return {
          name: workbook.getName?.() || fallbackName,
          sheets,
        };
      }
    }
  } catch {
    // fall through to snapshot parse
  }

  const snapshot = typeof workbook.save === 'function' ? workbook.save() : null;
  const parsed = fromUniverWorkbookData(snapshot);
  return {
    name: snapshot?.name || fallbackName,
    sheets: parsed.sheets || [],
  };
}

export function fromUniverWorkbookData(snapshot) {
  if (!snapshot || !snapshot.sheets) return { sheets: [] };
  const order = snapshot.sheetOrder && snapshot.sheetOrder.length
    ? snapshot.sheetOrder
    : Object.keys(snapshot.sheets);

  const sheets = order
    .map((id) => snapshot.sheets[id])
    .filter(Boolean)
    .map((sheet) => normalizeDateColumnsInSheet(
      sheetFromCellData(sheet.name, sheet.cellData || {}),
    ));

  return { sheets };
}

/** True when a grid still has Activity ID / Activity Name (aliases ok) on any sheet. */
export function gridHasRequiredHeaders(grid) {
  const sheets = grid?.sheets || [];
  return sheets.some((s) => hasScheduleHeaders(s.headers));
}
