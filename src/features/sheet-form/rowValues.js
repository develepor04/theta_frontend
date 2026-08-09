/**
 * Read / normalize a worksheet row for the Add/Edit Record form.
 */

/**
 * @param {object} univerAPI
 * @param {import('./types').FormConfig} config
 * @param {number} row - 0-based sheet row
 * @returns {Record<number, string | number | boolean> | null}
 */
export function readRowValues(univerAPI, config, row) {
  if (!univerAPI || !config?.fields?.length || !Number.isFinite(row) || row < 0) {
    return null;
  }

  const fWorkbook = resolveWorkbook(univerAPI, config.unitId);
  if (!fWorkbook) return null;

  const active = fWorkbook.getActiveSheet?.() || null;
  const fWorksheet = fWorkbook.getSheetByName?.(config.sheetName)
    || (active?.getSheetName?.() === config.sheetName ? active : null)
    || active;
  if (!fWorksheet?.getRange) return null;

  /** @type {Record<number, string | number | boolean>} */
  const values = {};
  let any = false;

  for (const field of config.fields) {
    const col = Number(field.column);
    if (!Number.isFinite(col) || col < 0) continue;
    let raw;
    try {
      raw = fWorksheet.getRange(row, col, 1, 1).getValue?.();
    } catch {
      raw = undefined;
    }
    const normalized = normalizeForForm(field, raw);
    values[col] = normalized;
    if (normalized !== '' && normalized !== null && normalized !== undefined) {
      any = true;
    }
  }

  return any ? values : null;
}

/**
 * True when a row looks like a real data row (not blank).
 * @param {Record<number, string | number | boolean> | null} values
 */
export function rowHasData(values) {
  if (!values) return false;
  return Object.values(values).some((v) => {
    if (v === undefined || v === null || v === '') return false;
    if (typeof v === 'string' && !v.trim()) return false;
    return true;
  });
}

/**
 * @param {import('./types').FormFieldConfig} field
 * @param {unknown} raw
 */
function normalizeForForm(field, raw) {
  if (raw === undefined || raw === null) {
    return field.inputType === 'checkbox' ? false : '';
  }

  if (field.inputType === 'checkbox') {
    if (typeof raw === 'boolean') return raw;
    const t = String(raw).trim().toLowerCase();
    return t === 'true' || t === '1' || t === 'yes' || t === 'y';
  }

  if (field.inputType === 'date') {
    return toDateInputValue(raw);
  }

  if (field.inputType === 'number') {
    if (typeof raw === 'number' && Number.isFinite(raw)) return raw;
    const n = Number(raw);
    return Number.isFinite(n) ? n : String(raw);
  }

  if (raw instanceof Date) {
    return toDateInputValue(raw);
  }

  return String(raw);
}

/**
 * Convert Date / Excel serial / ISO-ish string → `yyyy-MM-dd` for `<input type="date">`.
 * @param {unknown} raw
 */
export function toDateInputValue(raw) {
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return formatYmd(raw);
  }

  if (typeof raw === 'number' && Number.isFinite(raw)) {
    // Excel serial dates typically fall in this range for modern project schedules.
    if (raw > 20000 && raw < 80000) {
      const utc = Date.UTC(1899, 11, 30) + Math.round(raw) * 86400000;
      return formatYmd(new Date(utc));
    }
    return String(raw);
  }

  const s = String(raw ?? '').trim();
  if (!s) return '';

  const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;

  const parsed = new Date(s);
  if (!Number.isNaN(parsed.getTime())) return formatYmd(parsed);

  return s;
}

/**
 * @param {Date} d
 */
function formatYmd(d) {
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/**
 * @param {string} iso yyyy-MM-dd
 * @returns {number | string}
 */
export function isoToExcelSerial(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  if (!m) return iso;
  const utc = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const excelEpoch = Date.UTC(1899, 11, 30);
  return Math.round((utc - excelEpoch) / 86400000);
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
