import React, { useEffect, useState } from 'react';
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
 *   readOnly?: boolean,
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
  readOnly = false,
}) {
  const isEdit = Number.isFinite(editRow) && editRow >= 1;
  const panelTitle = title
    || (isEdit ? 'Row details' : 'Add record');
  const [tab, setTab] = useState('details');

  useEffect(() => {
    setTab('details');
  }, [editRow, open]);

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
            <div className="add-record-panel__heading">
              <h2 className="add-record-panel__title">{panelTitle}</h2>
              {isEdit ? (
                <p className="add-record-panel__sub">
                  {(config?.sheetName || 'Sheet')}
                  {' · '}
                  Row {editRow + 1}
                </p>
              ) : null}
            </div>
            <button
              type="button"
              className="add-record-panel__close"
              aria-label="Close"
              onClick={onClose}
            >
              ×
            </button>
          </header>

          {open && isEdit ? (
            <div className="add-record-panel__tabs" role="tablist">
              <button
                type="button"
                role="tab"
                aria-selected={tab === 'details'}
                className={`add-record-panel__tab${tab === 'details' ? ' is-on' : ''}`}
                onClick={() => setTab('details')}
              >
                Details
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={tab === 'comments'}
                className={`add-record-panel__tab${tab === 'comments' ? ' is-on' : ''}`}
                onClick={() => setTab('comments')}
              >
                Comments
              </button>
            </div>
          ) : null}

          {open && tab === 'comments' && isEdit ? (
            <div className="add-record-panel__comments">
              <p>
                Row comments are not stored yet. Details on this row stay in the
                sheet; threaded comments will use a later database table.
              </p>
            </div>
          ) : null}

          {open && (tab === 'details' || !isEdit) ? (
            <AddRecordForm
              key={isEdit ? `edit-${editRow}` : 'add'}
              univerAPI={univerAPI}
              config={config}
              initialFormValues={initialFormValues}
              readOnly={readOnly}
              onCancel={onClose}
              onSubmit={isEdit && !readOnly
                ? async (values) => {
                  const result = updateRecord(univerAPI, config, values, editRow);
                  if (!result.ok) {
                    throw new Error(result.error || 'Could not update record');
                  }
                }
                : undefined}
              onSuccess={() => {
                onRecordAdded?.();
                if (!isEdit) onClose();
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
