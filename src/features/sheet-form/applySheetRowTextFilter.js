/**
 * Show/hide data rows (row index >= 1) so only rows whose cell text
 * contains `query` remain visible. Header row 0 stays visible.
 *
 * An empty query does not walk the sheet (that was scrolling Univer to the
 * last row on every tab switch). It only bulk-unhides if a filter was applied.
 */

const sheetsWithTextFilter = new Set();

function sheetFilterKey(sheet) {
  return String(sheet?.getSheetId?.() || sheet?.getSheetName?.() || '');
}

/**
 * @param {object | null | undefined} univerAPI
 * @param {string} query
 * @returns {{ shown: number, hidden: number, supported: boolean }}
 */
export function applySheetRowTextFilter(univerAPI, query) {
  const sheet = univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.();
  if (!sheet) return { shown: 0, hidden: 0, supported: false };

  const key = sheetFilterKey(sheet);
  const lastRow = Number(sheet.getLastRow?.() ?? 0);
  const lastCol = Number(sheet.getLastColumn?.() ?? 0);
  const q = String(query || '').trim().toLowerCase();

  if (!q) {
    if (!key || !sheetsWithTextFilter.has(key)) {
      return { shown: 0, hidden: 0, supported: true };
    }
    const restored = unhideAllDataRows(sheet, lastRow);
    sheetsWithTextFilter.delete(key);
    return restored;
  }

  const hideFn = pickHideFn(sheet);
  const showFn = pickShowFn(sheet);
  let shown = 0;
  let hidden = 0;

  for (let r = 1; r <= lastRow; r++) {
    const match = rowText(sheet, r, lastCol).includes(q);
    try {
      if (match) {
        showFn(r);
        shown += 1;
      } else {
        hideFn(r);
        hidden += 1;
      }
    } catch {
      return { shown, hidden, supported: false };
    }
  }

  if (key) sheetsWithTextFilter.add(key);
  return { shown, hidden, supported: true };
}

function unhideAllDataRows(sheet, lastRow) {
  try {
    if (typeof sheet.unhideRows === 'function' && lastRow >= 1) {
      sheet.unhideRows(1, lastRow);
      return { shown: lastRow, hidden: 0, supported: true };
    }
    if (typeof sheet.showRows === 'function' && lastRow >= 1) {
      sheet.showRows(1, lastRow);
      return { shown: lastRow, hidden: 0, supported: true };
    }
    const showFn = pickShowFn(sheet);
    for (let r = 1; r <= lastRow; r++) showFn(r);
    return { shown: lastRow, hidden: 0, supported: true };
  } catch {
    return { shown: 0, hidden: 0, supported: false };
  }
}

function rowText(sheet, row, lastCol) {
  const parts = [];
  for (let c = 0; c <= lastCol; c++) {
    try {
      const v = sheet.getRange?.(row, c, 1, 1)?.getValue?.();
      if (v !== undefined && v !== null && v !== '') parts.push(String(v));
    } catch {
      /* skip */
    }
  }
  return parts.join(' ').toLowerCase();
}

function pickHideFn(sheet) {
  if (typeof sheet.hideRows === 'function') {
    return (r) => sheet.hideRows(r, r);
  }
  if (typeof sheet.hideRow === 'function') {
    return (r) => sheet.hideRow(r);
  }
  return (r) => {
    const range = sheet.getRange?.(r, 0, 1, 1);
    if (typeof range?.hideRows === 'function') {
      range.hideRows();
      return;
    }
    if (typeof range?.hide === 'function') {
      range.hide();
      return;
    }
    throw new Error('hide not supported');
  };
}

function pickShowFn(sheet) {
  if (typeof sheet.unhideRows === 'function') {
    return (r) => sheet.unhideRows(r, r);
  }
  if (typeof sheet.showRows === 'function') {
    return (r) => sheet.showRows(r, r);
  }
  if (typeof sheet.showRow === 'function') {
    return (r) => sheet.showRow(r);
  }
  return (r) => {
    const range = sheet.getRange?.(r, 0, 1, 1);
    if (typeof range?.unhideRows === 'function') {
      range.unhideRows();
      return;
    }
    if (typeof range?.show === 'function') {
      range.show();
      return;
    }
    throw new Error('show not supported');
  };
}
