import { useMemo } from 'react';
import { resolveFieldInput } from './inferFieldType';

const REQUIRED_LABELS = new Set([
  'activity id', 'activityid', 'activity code', 'id',
  'activity name', 'activityname',
]);

/**
 * Resolve FormConfig for the Add Record form.
 *
 * There is no persisted form-config store in this app yet — callers pass a
 * `FormConfig` (or a partial) and optionally let us fill `unitId` / `sheetName`
 * from the live Univer workbook, and enrich field types from data validation.
 *
 * @param {object | null | undefined} univerAPI
 * @param {Partial<import('./types').FormConfig> | import('./types').FormConfig | null | undefined} config
 * @param {{ headerRow?: number }} [options]
 */
export function useFormConfig(univerAPI, config, options = {}) {
  const headerRow = options.headerRow ?? 0;
  const configSheetName = config?.sheetName || '';
  const configUnitId = config?.unitId || '';
  const configFields = config?.fields;

  return useMemo(() => {
    const fWorkbook = univerAPI?.getActiveWorkbook?.() || null;
    const activeSheet = fWorkbook?.getActiveSheet?.() || null;
    const sheetName = configSheetName
      || activeSheet?.getSheetName?.()
      || '';
    const unitId = configUnitId
      || fWorkbook?.getId?.()
      || fWorkbook?.getUnitId?.()
      || '';

    const fWorksheet = sheetName && fWorkbook?.getSheetByName
      ? fWorkbook.getSheetByName(sheetName)
      : activeSheet;

    const baseFields = Array.isArray(configFields) && configFields.length > 0
      ? configFields
      : buildFieldsFromHeader(fWorksheet, headerRow);

    const fields = baseFields.map((field) => {
      const fromHeader = inferFromHeaderLabel(field.label);
      const resolved = resolveFieldInput(
        { ...field, inputType: field.inputType || fromHeader.inputType, options: field.options || fromHeader.options },
        fWorksheet,
        headerRow,
      );
      return {
        ...field,
        required: field.required || REQUIRED_LABELS.has(String(field.label || '').trim().toLowerCase()),
        inputType: resolved.inputType || fromHeader.inputType || 'text',
        options: (resolved.options?.length ? resolved.options : fromHeader.options) || [],
      };
    });

    return {
      unitId,
      sheetName,
      fields,
      ready: Boolean(univerAPI && fWorkbook && sheetName && fields.length > 0),
    };
  }, [univerAPI, configSheetName, configUnitId, configFields, headerRow]);
}

/**
 * Fallback: treat row 0 values as labels for every non-empty header cell.
 * @param {any} fWorksheet
 * @param {number} headerRow
 * @returns {import('./types').FormFieldConfig[]}
 */
function buildFieldsFromHeader(fWorksheet, headerRow) {
  if (!fWorksheet?.getRange) return [];

  let lastCol = 0;
  try {
    lastCol = typeof fWorksheet.getLastColumn === 'function'
      ? fWorksheet.getLastColumn()
      : fWorksheet.getRange('1:1').getLastColumn();
  } catch {
    return [];
  }

  if (!Number.isFinite(lastCol) || lastCol < 0) return [];

  /** @type {import('./types').FormFieldConfig[]} */
  const fields = [];
  for (let col = 0; col <= lastCol; col += 1) {
    let label = '';
    try {
      label = String(fWorksheet.getRange(headerRow, col, 1, 1).getValue?.() ?? '').trim();
    } catch {
      label = '';
    }
    if (!label) continue;
    const fromHeader = inferFromHeaderLabel(label);
    const options = fromHeader.inputType === 'select'
      ? (collectColumnOptions(fWorksheet, headerRow, col) || fromHeader.options || [])
      : fromHeader.options;
    fields.push({
      column: col,
      label,
      editable: true,
      required: REQUIRED_LABELS.has(label.toLowerCase()),
      inputType: fromHeader.inputType,
      options,
    });
  }
  return fields;
}

/**
 * @param {string} label
 * @returns {{ inputType: import('./types').FormFieldConfig['inputType'], options?: string[] }}
 */
function inferFromHeaderLabel(label) {
  const t = String(label || '').trim().toLowerCase();
  if (!t) return { inputType: 'text' };

  if (t === 'status') {
    return {
      inputType: 'select',
      options: ['Not Started', 'In Progress', 'Delayed', 'Completed', 'At Risk'],
    };
  }
  if (t === 'impact') {
    return { inputType: 'select', options: ['None', 'Low', 'Medium', 'High', 'Critical'] };
  }
  if (t === 'phase') {
    return {
      inputType: 'select',
      options: ['Engineering', 'Procurement', 'Construction', 'Commissioning'],
    };
  }
  if (t === 'period') {
    return { inputType: 'select', options: ['P1', 'P2', 'P3', 'P4'] };
  }
  if (/start|finish|end|date/.test(t)) {
    return { inputType: 'date' };
  }
  if (/%|cost|hours|output|variance|index|complete|budget|forecast|actual/.test(t) && !/root cause|category|unit/.test(t)) {
    return { inputType: 'number' };
  }
  return { inputType: 'text' };
}

/**
 * @param {any} fWorksheet
 * @param {number} headerRow
 * @param {number} col
 */
function collectColumnOptions(fWorksheet, headerRow, col) {
  const opts = new Set();
  let lastRow = headerRow;
  try {
    lastRow = typeof fWorksheet.getLastRow === 'function' ? fWorksheet.getLastRow() : headerRow;
  } catch {
    return null;
  }
  const maxScan = Math.min(lastRow, headerRow + 80);
  for (let r = headerRow + 1; r <= maxScan; r += 1) {
    try {
      const v = String(fWorksheet.getRange(r, col, 1, 1).getValue?.() ?? '').trim();
      if (v) opts.add(v);
    } catch {
      // skip
    }
  }
  return opts.size ? [...opts] : null;
}
