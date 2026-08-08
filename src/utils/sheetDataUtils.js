// Grid <-> Univer IWorkbookData conversion for Theta Sheets.
// Grid shape: { sheets: [{ name, headers, rows }] }
// Univer IWorkbookData shape: { id, name, sheetOrder, sheets: { [id]: {...} }, styles }

import { hasScheduleHeaders } from './thetaValidation';

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

export function blankGrid(name = 'Theta Sheets') {
  return { name, sheets: [{ name: 'Schedule', headers: REQUIRED_SHEET_COLUMNS, rows: [] }] };
}

export function toUniverWorkbookData({ sheets = [], name = 'Theta Sheets' } = {}) {
  const sheetList = sheets.length ? sheets : [{ name: 'Schedule', headers: REQUIRED_SHEET_COLUMNS, rows: [] }];
  const sheetOrder = [];
  const sheetsObj = {};

  sheetList.forEach((s, idx) => {
    const sheetId = `sheet-${idx}`;
    sheetOrder.push(sheetId);
    const headers = s.headers || [];
    const rows = s.rows || [];
    const cellData = {};

    cellData[0] = {};
    headers.forEach((h, c) => {
      cellData[0][c] = { v: h, s: HEADER_STYLE_ID };
    });

    rows.forEach((row, r) => {
      cellData[r + 1] = {};
      row.forEach((val, c) => {
        cellData[r + 1][c] = { v: val === undefined || val === null ? '' : val };
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
    styles: {
      [HEADER_STYLE_ID]: { bl: 1, bg: { rgb: '#F1F5F9' } },
    },
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

  return matrixToSheet(name, matrix);
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
    .map((sheet) => sheetFromCellData(sheet.name, sheet.cellData || {}));

  return { sheets };
}

/** True when a grid still has Activity ID / Activity Name (aliases ok) on any sheet. */
export function gridHasRequiredHeaders(grid) {
  const sheets = grid?.sheets || [];
  return sheets.some((s) => hasScheduleHeaders(s.headers));
}
