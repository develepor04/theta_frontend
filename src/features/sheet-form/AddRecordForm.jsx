import React, { useEffect, useState } from 'react';
import toast from 'react-hot-toast';
import { submitRecord } from './submitRecord';

/**
 * Field rendering + submit for the Add Record panel.
 * Uses native controls so input/submit works reliably outside Univer's design system.
 */
export default function AddRecordForm({
  univerAPI,
  config,
  onSuccess,
  onCancel,
  onSubmit,
  initialFormValues = null,
  submitLabel = 'Add to sheet',
  successMessage = 'Row added to sheet',
}) {
  const fields = config?.fields || [];
  const [values, setValues] = useState(() => ({
    ...initialValues(fields),
    ...(initialFormValues || {}),
  }));
  const [errors, setErrors] = useState({});
  const [submitting, setSubmitting] = useState(false);

  const initialKey = initialFormValues
    ? Object.keys(initialFormValues).sort().map((k) => `${k}:${initialFormValues[k]}`).join('|')
    : '';

  useEffect(() => {
    setValues({
      ...initialValues(fields),
      ...(initialFormValues || {}),
    });
    setErrors({});
    // initialFormValues intentionally keyed via initialKey to avoid object-identity loops
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [config?.sheetName, config?.unitId, fieldsKey(fields), initialKey]);

  const setFieldValue = (column, next) => {
    setValues((prev) => ({ ...prev, [column]: next }));
    setErrors((prev) => {
      if (!prev[column]) return prev;
      const copy = { ...prev };
      delete copy[column];
      return copy;
    });
  };

  const validate = () => {
    /** @type {Record<number, string>} */
    const nextErrors = {};
    for (const field of fields) {
      if (!field.editable || !field.required) continue;
      const value = values[field.column];
      const empty = value === undefined
        || value === null
        || value === ''
        || (typeof value === 'string' && value.trim() === '');
      if (empty && typeof value !== 'boolean') {
        nextErrors[field.column] = `${field.label} is required.`;
      }
    }
    setErrors(nextErrors);
    return Object.keys(nextErrors).length === 0;
  };

  const handleSubmit = async (e) => {
    e?.preventDefault?.();
    if (!validate()) {
      toast.error('Please fill required fields.');
      return;
    }
    setSubmitting(true);
    try {
      if (onSubmit) {
        await onSubmit(values);
      } else {
        const result = submitRecord(univerAPI, config, values);
        if (!result.ok) {
          toast.error(result.error || 'Could not add record');
          return;
        }
      }
      toast.success(successMessage);
      setValues(initialValues(fields));
      setErrors({});
      onSuccess?.();
    } catch (err) {
      toast.error(err?.response?.data?.error || err?.message || 'Could not save record');
    } finally {
      setSubmitting(false);
    }
  };

  if (!fields.length) {
    return (
      <div className="add-record-empty">
        No form fields configured for this sheet.
      </div>
    );
  }

  return (
    <form className="add-record-form" onSubmit={handleSubmit}>
      <div className="add-record-panel__body">
        {fields.map((field) => {
          if (!field.editable) {
            const raw = values[field.column] !== undefined && values[field.column] !== null && values[field.column] !== ''
              ? values[field.column]
              : field.defaultValue;
            const display = raw === undefined || raw === null || raw === ''
              ? '—'
              : String(raw);
            return (
              <div key={field.column} className="add-record-field">
                <span className="add-record-field__label">{field.label}</span>
                <div className="add-record-field__static">{display}</div>
              </div>
            );
          }

          return (
            <label key={field.column} className="add-record-field">
              <span className="add-record-field__label">
                {field.label}
                {field.required ? <span className="add-record-field__required">*</span> : null}
              </span>
              {renderControl(field, values[field.column], (next) => setFieldValue(field.column, next))}
              {errors[field.column] ? (
                <span className="add-record-field__error">{errors[field.column]}</span>
              ) : null}
            </label>
          );
        })}
      </div>

      <div className="add-record-panel__footer">
        <button
          type="button"
          className="add-record-btn add-record-btn--ghost"
          onClick={onCancel}
          disabled={submitting}
        >
          Cancel
        </button>
        <button
          type="submit"
          className="add-record-btn add-record-btn--primary"
          disabled={submitting || (config?.ready === false)}
        >
          {submitting ? (submitLabel.toLowerCase().includes('save') ? 'Saving…' : 'Adding…') : submitLabel}
        </button>
      </div>
    </form>
  );
}

function initialValues(fields) {
  /** @type {Record<number, string | number | boolean | undefined>} */
  const next = {};
  for (const field of fields) {
    if (!field.editable) continue;
    if (field.inputType === 'checkbox') next[field.column] = false;
    else next[field.column] = '';
  }
  return next;
}

function fieldsKey(fields) {
  return fields
    .map((f) => `${f.column}:${f.label}:${f.editable}:${f.inputType}:${(f.options || []).join('|')}`)
    .join(';');
}

function renderControl(field, value, onChange) {
  const inputType = field.inputType || 'text';
  const common = {
    className: 'add-record-input',
    value: value == null ? '' : String(value),
  };

  if (inputType === 'checkbox') {
    return (
      <input
        type="checkbox"
        className="add-record-checkbox"
        checked={Boolean(value)}
        onChange={(e) => onChange(e.target.checked)}
      />
    );
  }

  if (inputType === 'select') {
    return (
      <select
        className="add-record-input"
        value={value == null ? '' : String(value)}
        onChange={(e) => onChange(e.target.value)}
      >
        <option value="">Select…</option>
        {(field.options || []).map((opt) => (
          <option key={opt} value={opt}>{opt}</option>
        ))}
      </select>
    );
  }

  if (inputType === 'date') {
    return (
      <input
        type="date"
        {...common}
        onChange={(e) => onChange(e.target.value)}
      />
    );
  }

  if (inputType === 'number') {
    return (
      <input
        type="number"
        {...common}
        placeholder={field.label}
        onChange={(e) => {
          const next = e.target.value;
          if (next === '') {
            onChange('');
            return;
          }
          const num = Number(next);
          onChange(Number.isFinite(num) ? num : next);
        }}
      />
    );
  }

  return (
    <input
      type="text"
      {...common}
      placeholder={field.label}
      onChange={(e) => onChange(e.target.value)}
    />
  );
}
