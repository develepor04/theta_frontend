/**
 * Append one row to a worksheet from form state + FormConfig.
 *
 * @param {object} univerAPI - FUniver facade
 * @param {import('./types').FormConfig} config
 * @param {Record<number, string | number | boolean>} formState - keyed by column index
 * @returns {{ ok: true, targetRow: number } | { ok: false, error: string }}
 */
export function submitRecord(univerAPI, config, formState) {
  if (!univerAPI) {
    return { ok: false, error: 'Spreadsheet is not ready.' };
  }
  if (!config?.sheetName || !Array.isArray(config.fields)) {
    return { ok: false, error: 'Form configuration is incomplete.' };
  }

  const fWorkbook = resolveWorkbook(univerAPI, config.unitId);
  if (!fWorkbook) {
    return { ok: false, error: 'Workbook not found.' };
  }

  const active = fWorkbook.getActiveSheet?.() || null;
  const fWorksheet = fWorkbook.getSheetByName?.(config.sheetName)
    || (active?.getSheetName?.() === config.sheetName ? active : null)
    || active;

  if (!fWorksheet) {
    return { ok: false, error: `Sheet "${config.sheetName}" not found.` };
  }

  const totalCols = resolveTotalCols(fWorksheet, config.fields);
  /** @type {Array<string | number | boolean | null>} */
  const rowContents = Array.from({ length: totalCols }, () => null);
  let wrote = 0;

  for (const field of config.fields) {
    const col = Number(field.column);
    if (!Number.isFinite(col) || col < 0) continue;

    if (field.editable) {
      const value = formState?.[field.column];
      if (value !== undefined && value !== null && value !== '') {
        rowContents[col] = coerceValue(value);
        wrote += 1;
      } else if (typeof value === 'boolean') {
        rowContents[col] = value;
        wrote += 1;
      } else if (typeof value === 'number' && Number.isFinite(value)) {
        rowContents[col] = value;
        wrote += 1;
      }
    } else if (field.defaultValue !== undefined) {
      rowContents[col] = coerceValue(field.defaultValue);
      wrote += 1;
    }
  }

  if (wrote === 0) {
    return { ok: false, error: 'Nothing to submit.' };
  }

  try {
    // Prefer the dedicated append API (Univer 0.25+).
    if (typeof fWorksheet.appendRow === 'function') {
      fWorksheet.appendRow(rowContents);
      const lastRow = typeof fWorksheet.getLastRow === 'function'
        ? fWorksheet.getLastRow()
        : -1;
      return { ok: true, targetRow: Number.isFinite(lastRow) ? lastRow : -1 };
    }

    // Fallback: write cells on the next empty row.
    let lastRow = typeof fWorksheet.getLastRow === 'function' ? fWorksheet.getLastRow() : 0;
    if (!Number.isFinite(lastRow) || lastRow < 0) lastRow = 0;
    const targetRow = Math.max(1, lastRow + 1);
    for (let c = 0; c < rowContents.length; c += 1) {
      const value = rowContents[c];
      if (value === null || value === undefined || value === '') continue;
      const range = fWorksheet.getRange(targetRow, c, 1, 1);
      if (typeof range.setValue === 'function') {
        range.setValue(value);
      } else {
        range.setValues([[value]]);
      }
    }
    return { ok: true, targetRow };
  } catch (err) {
    return {
      ok: false,
      error: err?.message || 'Failed to write the new row.',
    };
  }
}

/**
 * @param {object} univerAPI
 * @param {string} [unitId]
 */
function resolveWorkbook(univerAPI, unitId) {
  if (unitId && typeof univerAPI.getWorkbook === 'function') {
    const byId = univerAPI.getWorkbook(unitId);
    if (byId) return byId;
  }
  return univerAPI.getActiveWorkbook?.() || null;
}

/**
 * @param {any} fWorksheet
 * @param {import('./types').FormFieldConfig[]} fields
 */
function resolveTotalCols(fWorksheet, fields) {
  let fromSheet = 0;
  try {
    if (typeof fWorksheet.getLastColumn === 'function') {
      fromSheet = fWorksheet.getLastColumn() + 1;
    }
  } catch {
    fromSheet = 0;
  }
  const fromFields = fields.reduce((max, f) => Math.max(max, (f.column ?? 0) + 1), 0);
  return Math.max(fromSheet, fromFields, 1);
}

/**
 * @param {string | number | boolean | Date} value
 * @returns {string | number | boolean}
 */
function coerceValue(value) {
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return value;
}
