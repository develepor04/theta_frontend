/**
 * Show/hide data rows (row index >= 1) so only rows whose cell text
 * contains `query` remain visible. Header row 0 stays visible.
 *
 * Uses Univer worksheet hide/show helpers when present.
 *
 * @param {object | null | undefined} univerAPI
 * @param {string} query
 * @returns {{ shown: number, hidden: number, supported: boolean }}
 */
export function applySheetRowTextFilter(univerAPI, query) {
  const sheet = univerAPI?.getActiveWorkbook?.()?.getActiveSheet?.();
  if (!sheet) return { shown: 0, hidden: 0, supported: false };

  const hideFn = pickHideFn(sheet);
  const showFn = pickShowFn(sheet);

  const lastRow = Number(sheet.getLastRow?.() ?? 0);
  const lastCol = Number(sheet.getLastColumn?.() ?? 0);
  const q = String(query || '').trim().toLowerCase();
  let shown = 0;
  let hidden = 0;

  for (let r = 1; r <= lastRow; r++) {
    const match = !q || rowText(sheet, r, lastCol).includes(q);
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

  return { shown, hidden, supported: true };
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
