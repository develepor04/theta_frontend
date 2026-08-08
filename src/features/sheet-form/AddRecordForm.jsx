import React, { useEffect, useState } from 'react';
import toast from 'react-hot-toast';
import {
  Button,
  Checkbox,
  DatePicker,
  FormLayout,
  Input,
  Select,
} from '@univerjs/design';
import { submitRecord } from './submitRecord';

/**
 * Field rendering + submit for the Add Record panel.
 *
 * @param {{
 *   univerAPI?: object | null,
 *   config: import('./types').FormConfig & { ready?: boolean },
 *   onSuccess?: () => void,
 *   onCancel?: () => void,
 *   onSubmit?: (values: Record<number, any>) => Promise<void> | void,
 *   initialFormValues?: Record<number, any> | null,
 *   submitLabel?: string,
 *   successMessage?: string,
 * }} props
 */
export default function AddRecordForm({
  univerAPI,
  config,
  onSuccess,
  onCancel,
  onSubmit,
  initialFormValues = null,
  submitLabel = 'Submit',
  successMessage = 'Record added',
}) {
  const fields = config?.fields || [];
  const [values, setValues] = useState(() => ({
    ...initialValues(fields),
    ...(initialFormValues || {}),
  }));
  const [errors, setErrors] = useState({});
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    setValues({
      ...initialValues(fields),
      ...(initialFormValues || {}),
    });
    setErrors({});
  }, [config?.sheetName, config?.unitId, fieldsKey(fields), JSON.stringify(initialFormValues || {})]);

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

  const handleSubmit = async () => {
    if (!validate()) return;
    setSubmitting(true);
    try {
      if (onSubmit) {
        await onSubmit(values);
      } else {
        const result = submitRecord(univerAPI, config, values);
        if (!result.ok) {
          toast.error(result.error);
          return;
        }
        toast.success(successMessage);
      }
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
    <>
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
            <FormLayout
              key={field.column}
              className="add-record-field"
              label={(
                <span className="add-record-field__label">
                  {field.label}
                  {field.required ? <span className="add-record-field__required">*</span> : null}
                </span>
              )}
              error={errors[field.column]}
            >
              {renderControl(field, values[field.column], (next) => setFieldValue(field.column, next))}
            </FormLayout>
          );
        })}
      </div>

      <div className="add-record-panel__footer">
        <Button type="button" variant="default" size="small" onClick={onCancel} disabled={submitting}>
          Cancel
        </Button>
        <Button
          type="button"
          variant="primary"
          size="small"
          onClick={handleSubmit}
          disabled={submitting || (config?.ready === false)}
        >
          {submitLabel}
        </Button>
      </div>
    </>
  );
}

/**
 * @param {import('./types').FormFieldConfig[]} fields
 */
function initialValues(fields) {
  /** @type {Record<number, string | number | boolean | Date | undefined>} */
  const next = {};
  for (const field of fields) {
    if (!field.editable) continue;
    if (field.inputType === 'checkbox') {
      next[field.column] = false;
    } else if (field.inputType === 'date') {
      next[field.column] = undefined;
    } else if (field.inputType === 'number') {
      next[field.column] = '';
    } else if (field.inputType === 'select') {
      next[field.column] = '';
    } else {
      next[field.column] = '';
    }
  }
  return next;
}

/**
 * @param {import('./types').FormFieldConfig[]} fields
 */
function fieldsKey(fields) {
  return fields
    .map((f) => `${f.column}:${f.label}:${f.editable}:${f.inputType}:${(f.options || []).join('|')}`)
    .join(';');
}

/**
 * @param {import('./types').FormFieldConfig} field
 * @param {any} value
 * @param {(next: any) => void} onChange
 */
function renderControl(field, value, onChange) {
  const inputType = field.inputType || 'text';

  if (inputType === 'checkbox') {
    return (
      <Checkbox
        checked={Boolean(value)}
        onChange={(checked) => onChange(Boolean(checked))}
      >
        {field.label}
      </Checkbox>
    );
  }

  if (inputType === 'select') {
    const options = (field.options || []).map((opt) => ({ label: opt, value: opt }));
    return (
      <Select
        value={value == null ? '' : String(value)}
        options={[{ label: 'Select…', value: '' }, ...options]}
        onChange={(next) => onChange(next)}
      />
    );
  }

  if (inputType === 'date') {
    const dateValue = value instanceof Date
      ? value
      : (value ? new Date(value) : undefined);
    return (
      <DatePicker
        value={dateValue && !Number.isNaN(dateValue.getTime()) ? dateValue : undefined}
        onValueChange={(date) => onChange(date)}
      />
    );
  }

  if (inputType === 'number') {
    return (
      <Input
        type="number"
        value={value == null ? '' : String(value)}
        placeholder={field.label}
        onChange={(next) => {
          if (next === '' || next == null) {
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
    <Input
      type="text"
      value={value == null ? '' : String(value)}
      placeholder={field.label}
      onChange={(next) => onChange(next)}
    />
  );
}
