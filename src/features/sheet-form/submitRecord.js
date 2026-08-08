/**
 * Append one sparse row to a worksheet from form state + FormConfig.
 * Columns omitted from the sparse matrix (e.g. formula-driven) are left untouched.
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

  const fWorksheet = fWorkbook.getSheetByName?.(config.sheetName)
    || (fWorkbook.getActiveSheet?.()?.getSheetName?.() === config.sheetName
      ? fWorkbook.getActiveSheet()
      : null);

  if (!fWorksheet) {
    return { ok: false, error: `Sheet "${config.sheetName}" not found.` };
  }

  const lastRow = typeof fWorksheet.getLastRow === 'function'
    ? fWorksheet.getLastRow()
    : -1;
  const targetRow = (Number.isFinite(lastRow) ? lastRow : -1) + 1;

  const totalCols = resolveTotalCols(fWorksheet, config.fields);
  /** @type {Record<number, Record<number, string | number | boolean>>} */
  const rowValues = { 0: {} };

  for (const field of config.fields) {
    if (field.editable) {
      const value = formState?.[field.column];
      if (value !== undefined && value !== null && value !== '') {
        rowValues[0][field.column] = coerceValue(value);
      } else if (typeof value === 'boolean') {
        rowValues[0][field.column] = value;
      } else if (typeof value === 'number' && Number.isFinite(value)) {
        rowValues[0][field.column] = value;
      }
      // editable but empty: omit (required validation should have blocked submit)
    } else if (field.defaultValue !== undefined) {
      rowValues[0][field.column] = field.defaultValue;
    }
    // else: omit entirely — leave formulas / system columns untouched
  }

  if (Object.keys(rowValues[0]).length === 0) {
    return { ok: false, error: 'Nothing to submit.' };
  }

  try {
    const range = fWorksheet.getRange(targetRow, 0, 1, totalCols);
    // Prefer sparse object-matrix (supported in @univerjs/sheets 0.25.0).
    // Fall back to per-cell writes if the dense-array-only path is forced.
    if (typeof range.setValues === 'function') {
      try {
        range.setValues(rowValues);
      } catch {
        writeSparseCells(fWorksheet, targetRow, rowValues[0]);
      }
    } else {
      writeSparseCells(fWorksheet, targetRow, rowValues[0]);
    }
  } catch (err) {
    return {
      ok: false,
      error: err?.message || 'Failed to write the new row.',
    };
  }

  return { ok: true, targetRow };
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
    } else {
      const header = fWorksheet.getRange?.('1:1');
      if (header?.getLastColumn) {
        fromSheet = header.getLastColumn() + 1;
      }
    }
  } catch {
    fromSheet = 0;
  }

  const fromFields = fields.reduce((max, f) => Math.max(max, (f.column ?? 0) + 1), 0);
  return Math.max(fromSheet, fromFields, 1);
}

/**
 * @param {any} fWorksheet
 * @param {number} row
 * @param {Record<number, string | number | boolean>} cells
 */
function writeSparseCells(fWorksheet, row, cells) {
  for (const [colKey, value] of Object.entries(cells)) {
    const col = Number(colKey);
    fWorksheet.getRange(row, col, 1, 1).setValues([[value]]);
  }
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
