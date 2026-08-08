import { useMemo } from 'react';
import { resolveFieldInput } from './inferFieldType';

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
      const resolved = resolveFieldInput(field, fWorksheet, headerRow);
      return {
        ...field,
        inputType: resolved.inputType,
        options: resolved.options,
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
    fields.push({
      column: col,
      label,
      editable: true,
      required: false,
    });
  }
  return fields;
}
