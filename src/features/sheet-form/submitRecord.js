import { isDateColumnLabel } from '../../utils/sheetDataUtils';
import { getActiveDateFormat } from '../../utils/dateFormats';
import { isoToExcelSerial } from './rowValues';

/**
 * Append one row to a worksheet from form state + FormConfig.
 *
 * @param {object} univerAPI - FUniver facade
 * @param {import('./types').FormConfig} config
 * @param {Record<number, string | number | boolean>} formState - keyed by column index
 * @returns {{ ok: true, targetRow: number } | { ok: false, error: string }}
 */
export function submitRecord(univerAPI, config, formState) {
  const resolved = resolveSheet(univerAPI, config);
  if (!resolved.ok) return resolved;
  const { fWorksheet } = resolved;

  const built = buildRowContents(config, formState, fWorksheet, { fillDefaults: true });
  if (!built.ok) return built;
  const { rowContents, wrote } = built;

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
      const targetRow = Number.isFinite(lastRow) ? lastRow : -1;
      if (targetRow >= 1) {
        applyDateFormatsForFields(fWorksheet, config.fields, targetRow);
      }
      return { ok: true, targetRow };
    }

    // Fallback: write cells on the next empty row.
    let lastRow = typeof fWorksheet.getLastRow === 'function' ? fWorksheet.getLastRow() : 0;
    if (!Number.isFinite(lastRow) || lastRow < 0) lastRow = 0;
    const targetRow = Math.max(1, lastRow + 1);
    writeRowCells(fWorksheet, targetRow, rowContents, { clearEmpty: false });
    applyDateFormatsForFields(fWorksheet, config.fields, targetRow);
    return { ok: true, targetRow };
  } catch (err) {
    return {
      ok: false,
      error: err?.message || 'Failed to write the new row.',
    };
  }
}

/**
 * Update an existing worksheet row in place from form state.
 *
 * @param {object} univerAPI
 * @param {import('./types').FormConfig} config
 * @param {Record<number, string | number | boolean>} formState
 * @param {number} targetRow - 0-based sheet row
 * @returns {{ ok: true, targetRow: number } | { ok: false, error: string }}
 */
export function updateRecord(univerAPI, config, formState, targetRow) {
  if (!Number.isFinite(targetRow) || targetRow < 1) {
    return { ok: false, error: 'Invalid row to update.' };
  }

  const resolved = resolveSheet(univerAPI, config);
  if (!resolved.ok) return resolved;
  const { fWorksheet } = resolved;

  const built = buildRowContents(config, formState, fWorksheet, { fillDefaults: false });
  if (!built.ok) return built;
  const { rowContents } = built;

  try {
    // Only touch editable columns; clear emptied fields so edits stick.
    for (const field of config.fields) {
      if (!field.editable) continue;
      const col = Number(field.column);
      if (!Number.isFinite(col) || col < 0) continue;
      const value = rowContents[col];
      const range = fWorksheet.getRange(targetRow, col, 1, 1);
      const next = value === null || value === undefined ? '' : value;
      if (typeof range.setValue === 'function') {
        range.setValue(next);
      } else {
        range.setValues([[next]]);
      }
    }
    applyDateFormatsForFields(fWorksheet, config.fields, targetRow);
    return { ok: true, targetRow };
  } catch (err) {
    return {
      ok: false,
      error: err?.message || 'Failed to update the row.',
    };
  }
}

/**
 * Keep date cells numeric-formatted after form writes.
 * Prefer the format already on this column (user Formats choice) over the default.
 * @param {any} fWorksheet
 * @param {import('./types').FormFieldConfig[]} fields
 * @param {number} targetRow
 */
function applyDateFormatsForFields(fWorksheet, fields, targetRow) {
  if (!fWorksheet?.getRange || !Number.isFinite(targetRow) || targetRow < 1) return;
  for (const field of fields || []) {
    const isDate = field.inputType === 'date' || isDateColumnLabel(field.label);
    if (!isDate) continue;
    const col = Number(field.column);
    if (!Number.isFinite(col) || col < 0) continue;
    try {
      const range = fWorksheet.getRange(targetRow, col, 1, 1);
      if (typeof range.setNumberFormat !== 'function') continue;
      let fmt = getActiveDateFormat();
      try {
        const probeRow = targetRow > 1 ? targetRow - 1 : targetRow;
        const probe = fWorksheet.getRange(probeRow, col, 1, 1);
        const existing = probe.getNumberFormat?.() || range.getNumberFormat?.();
        if (existing && String(existing).trim() && !/^general$/i.test(String(existing).trim())) {
          fmt = String(existing).trim();
        }
      } catch {
        // keep default
      }
      range.setNumberFormat(fmt);
    } catch {
      // ignore per-cell format failures
    }
  }
}

/**
 * @param {object} univerAPI
 * @param {import('./types').FormConfig} config
 */
function resolveSheet(univerAPI, config) {
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

  return { ok: true, fWorksheet };
}

/**
 * @param {import('./types').FormConfig} config
 * @param {Record<number, string | number | boolean>} formState
 * @param {any} fWorksheet
 * @param {{ fillDefaults?: boolean }} [opts]
 */
function buildRowContents(config, formState, fWorksheet, opts = {}) {
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
        rowContents[col] = coerceFieldValue(field, value);
        wrote += 1;
      } else if (typeof value === 'boolean') {
        rowContents[col] = value;
        wrote += 1;
      } else if (typeof value === 'number' && Number.isFinite(value)) {
        rowContents[col] = value;
        wrote += 1;
      } else {
        // Explicit empty so update can clear the cell
        rowContents[col] = '';
      }
    } else if (opts.fillDefaults && field.defaultValue !== undefined) {
      rowContents[col] = coerceFieldValue(field, field.defaultValue);
      wrote += 1;
    }
  }

  return { ok: true, rowContents, wrote };
}

/**
 * @param {any} fWorksheet
 * @param {number} targetRow
 * @param {Array<string | number | boolean | null>} rowContents
 * @param {{ clearEmpty?: boolean }} [opts]
 */
function writeRowCells(fWorksheet, targetRow, rowContents, opts = {}) {
  for (let c = 0; c < rowContents.length; c += 1) {
    const value = rowContents[c];
    if (!opts.clearEmpty && (value === null || value === undefined || value === '')) continue;
    const range = fWorksheet.getRange(targetRow, c, 1, 1);
    const next = value === null || value === undefined ? '' : value;
    if (typeof range.setValue === 'function') {
      range.setValue(next);
    } else {
      range.setValues([[next]]);
    }
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
 * @param {import('./types').FormFieldConfig} field
 * @param {string | number | boolean | Date} value
 * @returns {string | number | boolean}
 */
function coerceFieldValue(field, value) {
  if (field?.inputType === 'date') {
    if (value instanceof Date) {
      const y = value.getFullYear();
      const m = String(value.getMonth() + 1).padStart(2, '0');
      const d = String(value.getDate()).padStart(2, '0');
      return isoToExcelSerial(`${y}-${m}-${d}`);
    }
    const s = String(value ?? '').trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      return isoToExcelSerial(s);
    }
  }
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return value;
}
