/**
 * Fast Google Sheets–style date/time format presets for Univer numfmt.
 * Patterns use Excel/Univer tokens (compatible with setNumberFormat).
 */

export const DEFAULT_DATE_FORMAT = 'yyyy-mm-dd';

/** @type {{ id: string, label: string, pattern: string, group: 'date' | 'time' | 'datetime' }[]} */
export const GOOGLE_SHEETS_DATE_FORMATS = [
  // Date
  { id: 'date', label: 'Date', pattern: 'm/d/yyyy', group: 'date' },
  { id: 'date_iso', label: 'ISO date', pattern: 'yyyy-mm-dd', group: 'date' },
  { id: 'date_eu', label: 'Day/Month/Year', pattern: 'dd/mm/yyyy', group: 'date' },
  { id: 'date_us_padded', label: 'MM/DD/YYYY', pattern: 'mm/dd/yyyy', group: 'date' },
  { id: 'date_dash', label: 'DD-MMM-YYYY', pattern: 'dd-mmm-yyyy', group: 'date' },
  { id: 'date_long', label: 'Month Day, Year', pattern: 'mmmm d, yyyy', group: 'date' },
  { id: 'date_medium', label: 'MMM D, YYYY', pattern: 'mmm d, yyyy', group: 'date' },
  { id: 'date_full', label: 'Weekday, Month Day, Year', pattern: 'dddd, mmmm d, yyyy', group: 'date' },
  { id: 'date_month_year', label: 'Month Year', pattern: 'mmmm yyyy', group: 'date' },
  { id: 'date_mmm_yyyy', label: 'MMM-YYYY', pattern: 'mmm-yyyy', group: 'date' },
  { id: 'date_month_day', label: 'Month Day', pattern: 'mmmm d', group: 'date' },
  { id: 'date_mmm_d', label: 'MMM D', pattern: 'mmm d', group: 'date' },
  { id: 'date_year', label: 'Year', pattern: 'yyyy', group: 'date' },

  // Time
  { id: 'time', label: 'Time', pattern: 'h:mm:ss AM/PM', group: 'time' },
  { id: 'time_24', label: '24-hour time', pattern: 'hh:mm:ss', group: 'time' },
  { id: 'time_short', label: 'Hour:Minute', pattern: 'h:mm AM/PM', group: 'time' },
  { id: 'time_24_short', label: '24-hour short', pattern: 'hh:mm', group: 'time' },

  // Date time
  { id: 'datetime', label: 'Date time', pattern: 'm/d/yyyy h:mm:ss AM/PM', group: 'datetime' },
  { id: 'datetime_iso', label: 'ISO date time', pattern: 'yyyy-mm-dd hh:mm:ss', group: 'datetime' },
  { id: 'datetime_eu', label: 'DD/MM/YYYY HH:MM', pattern: 'dd/mm/yyyy hh:mm', group: 'datetime' },
];

let activeDateFormat = DEFAULT_DATE_FORMAT;

export function getActiveDateFormat() {
  return activeDateFormat || DEFAULT_DATE_FORMAT;
}

export function setActiveDateFormat(pattern) {
  const next = String(pattern || '').trim();
  if (!next) return getActiveDateFormat();
  activeDateFormat = next;
  try {
    localStorage.setItem('theta.dateNumberFormat', next);
  } catch {
    // ignore
  }
  return activeDateFormat;
}

export function loadStoredDateFormat() {
  try {
    const stored = localStorage.getItem('theta.dateNumberFormat');
    if (stored && stored.trim()) {
      activeDateFormat = stored.trim();
    }
  } catch {
    // ignore
  }
  return getActiveDateFormat();
}
