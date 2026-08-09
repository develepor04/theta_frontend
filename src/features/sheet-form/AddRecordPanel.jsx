import React from 'react';
import './addRecordPanel.css';
import AddRecordForm from './AddRecordForm';
import { updateRecord } from './submitRecord';

/**
 * Flex-row layout sibling to the Univer mount element.
 * Opening the panel shrinks the sheet container width so Univer's ResizeObserver reflows the grid.
 *
 * @param {{
 *   open: boolean,
 *   onClose: () => void,
 *   title?: string,
 *   univerAPI: object | null,
 *   config: import('./types').FormConfig & { ready?: boolean },
 *   children: React.ReactNode,
 *   onRecordAdded?: () => void,
 *   editRow?: number | null,
 *   initialFormValues?: Record<number, string | number | boolean> | null,
 * }} props
 */
export default function AddRecordPanel({
  open,
  onClose,
  title,
  univerAPI,
  config,
  children,
  onRecordAdded,
  editRow = null,
  initialFormValues = null,
}) {
  const isEdit = Number.isFinite(editRow) && editRow >= 1;
  const panelTitle = title || (isEdit ? 'Edit Record' : 'Add Record');

  return (
    <div className="add-record-layout">
      <div className="add-record-layout__sheet">
        {children}
      </div>

      <aside
        className={`add-record-panel${open ? ' add-record-panel--open' : ''}`}
        aria-hidden={!open}
      >
        <div className="add-record-panel__inner">
          <header className="add-record-panel__header">
            <h2 className="add-record-panel__title">{panelTitle}</h2>
            <button
              type="button"
              className="add-record-panel__close"
              aria-label="Close"
              onClick={onClose}
            >
              ×
            </button>
          </header>

          {open ? (
            <AddRecordForm
              key={isEdit ? `edit-${editRow}` : 'add'}
              univerAPI={univerAPI}
              config={config}
              initialFormValues={initialFormValues}
              onCancel={onClose}
              onSubmit={isEdit
                ? async (values) => {
                  const result = updateRecord(univerAPI, config, values, editRow);
                  if (!result.ok) {
                    throw new Error(result.error || 'Could not update record');
                  }
                }
                : undefined}
              onSuccess={() => {
                onRecordAdded?.();
                onClose();
              }}
              submitLabel={isEdit ? 'Save changes' : 'Add to sheet'}
              successMessage={isEdit ? 'Row updated' : 'Row added to sheet'}
            />
          ) : null}
        </div>
      </aside>
    </div>
  );
}
