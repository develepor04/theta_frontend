import React from 'react';
import './addRecordPanel.css';
import AddRecordForm from './AddRecordForm';

/**
 * Flex-row layout sibling to the Univer mount element.
 * Opening the panel shrinks the sheet container width so Univer's ResizeObserver reflows the grid.
 *
 * Usage:
 * ```jsx
 * <AddRecordPanel open={open} onClose={...} univerAPI={api} config={formConfig}>
 *   <div ref={containerRef} style={{ height: '100%' }} />
 * </AddRecordPanel>
 * ```
 *
 * @param {{
 *   open: boolean,
 *   onClose: () => void,
 *   title?: string,
 *   univerAPI: object | null,
 *   config: import('./types').FormConfig & { ready?: boolean },
 *   children: React.ReactNode,
 * }} props
 */
export default function AddRecordPanel({
  open,
  onClose,
  title = 'Add Record',
  univerAPI,
  config,
  children,
  onRecordAdded,
}) {
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
            <h2 className="add-record-panel__title">{title}</h2>
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
              univerAPI={univerAPI}
              config={config}
              onCancel={onClose}
              onSuccess={() => {
                onRecordAdded?.();
                onClose();
              }}
              submitLabel="Add to sheet"
              successMessage="Row added to sheet"
            />
          ) : null}
        </div>
      </aside>
    </div>
  );
}
