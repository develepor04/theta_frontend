/**
 * @typedef {'text' | 'number' | 'select' | 'date' | 'checkbox'} FormFieldInputType
 */

/**
 * @typedef {Object} FormFieldConfig
 * @property {number} column - 0-indexed target column
 * @property {string} label
 * @property {boolean} editable - false = read-only / system-managed in the form
 * @property {boolean} [required]
 * @property {FormFieldInputType} [inputType] - if omitted, auto-detect from column data validation
 * @property {string | number | boolean} [defaultValue] - written when editable=false and a value is wanted
 * @property {string[]} [options] - for select, if not auto-detected from data validation
 */

/**
 * @typedef {Object} FormConfig
 * @property {string} unitId
 * @property {string} sheetName
 * @property {FormFieldConfig[]} fields
 */

/**
 * @typedef {Object} InferredFieldType
 * @property {FormFieldInputType | null} inputType - null means no DV rule; use config fallback
 * @property {string[]} [options]
 */

export {};
