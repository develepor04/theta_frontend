/**
 * Map Univer data-validation criteria on a cell/column to a form input type.
 * Pure / Facade-only: expects `fRange.getDataValidation()` (needs sheets-data-validation plugin).
 *
 * @param {{ getDataValidation?: () => { getCriteriaType?: () => string, getCriteriaValues?: () => [string|undefined, string|undefined, string|undefined], rule?: { type?: string, formula1?: string } } | null | undefined }} fRange
 * @returns {{ inputType: 'text' | 'number' | 'select' | 'date' | 'checkbox' | null, options?: string[] }}
 */
export function inferFieldType(fRange) {
  if (!fRange || typeof fRange.getDataValidation !== 'function') {
    return { inputType: null };
  }

  let dv;
  try {
    dv = fRange.getDataValidation();
  } catch {
    return { inputType: null };
  }

  if (!dv) {
    return { inputType: null };
  }

  const criteriaType = String(
    dv.getCriteriaType?.() ?? dv.rule?.type ?? ''
  ).toLowerCase();

  if (criteriaType === 'list' || criteriaType === 'listmultiple') {
    const [, formula1] = dv.getCriteriaValues?.() ?? [];
    const options = parseListOptions(formula1 ?? dv.rule?.formula1);
    return options.length > 0
      ? { inputType: 'select', options }
      : { inputType: 'select' };
  }

  if (criteriaType === 'date' || criteriaType === 'time') {
    return { inputType: 'date' };
  }

  if (criteriaType === 'checkbox') {
    return { inputType: 'checkbox' };
  }

  if (criteriaType === 'decimal' || criteriaType === 'whole') {
    return { inputType: 'number' };
  }

  if (criteriaType === 'textlength' || criteriaType === 'custom' || criteriaType === 'any') {
    return { inputType: 'text' };
  }

  return { inputType: 'text' };
}

/**
 * Univer stores list options as JSON.stringify([...]) (see serializeListOptions),
 * with a comma-separated fallback.
 * @param {string | undefined} formula1
 * @returns {string[]}
 */
function parseListOptions(formula1) {
  if (formula1 == null || formula1 === '') return [];
  const raw = String(formula1).trim();
  // Range/formula lists (e.g. =Sheet1!$A$1:$A$10) can't be resolved here.
  if (raw.startsWith('=')) return [];

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && parsed.every((o) => typeof o === 'string')) {
      return parsed.filter(Boolean);
    }
  } catch {
    // fall through
  }

  return raw.split(',').map((s) => s.trim()).filter(Boolean);
}

/**
 * Resolve the effective input type + options for a field config against a worksheet.
 *
 * @param {import('./types').FormFieldConfig} field
 * @param {{ getRange: (row: number, col: number, h: number, w: number) => any }} fWorksheet
 * @param {number} [headerRow=0]
 */
export function resolveFieldInput(field, fWorksheet, headerRow = 0) {
  let inferred = { inputType: null };
  if (fWorksheet?.getRange) {
    try {
      const fRange = fWorksheet.getRange(headerRow, field.column, 1, 1);
      inferred = inferFieldType(fRange);
    } catch {
      inferred = { inputType: null };
    }
  }

  if (inferred.inputType) {
    return {
      inputType: inferred.inputType,
      options: inferred.options?.length ? inferred.options : (field.options || []),
    };
  }

  return {
    inputType: field.inputType || 'text',
    options: field.options || [],
  };
}
