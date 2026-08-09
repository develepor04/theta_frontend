import React, { useEffect, useImperativeHandle, useRef, useState, forwardRef } from 'react';
import { createPortal } from 'react-dom';
import toast from 'react-hot-toast';
import * as XLSX from 'xlsx';
import { createUniver, ICommandService, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import UniverPresetSheetsDataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US';
import { IMenuManagerService, RibbonStartGroup } from '@univerjs/ui';
import '@univerjs/sheets-numfmt/lib/facade';
import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/sheets-numfmt-ui/lib/index.css';
import '@univerjs/design/lib/index.css';
import { sheetService } from '../services/api';
import {
  toUniverWorkbookData,
  extractGridFromUniverWorkbook,
  gridHasRequiredHeaders,
  blankGrid,
  applyDateNumberFormatsToWorkbook,
  applyColumnFormatsFromGrid,
  coerceDateColumnValuesInWorkbook,
} from '../utils/sheetDataUtils';
import {
  getActiveDateFormat,
  loadStoredDateFormat,
} from '../utils/dateFormats';
import { validateSheetGrid } from '../utils/thetaValidation';
import {
  AddRecordButton,
  AddRecordPanel,
  useFormConfig,
} from '../features/sheet-form';
import { readRowValues, rowHasData } from '../features/sheet-form/rowValues';

const SAVE_DEBOUNCE_MS = 500;

/** Univer numfmt ribbon items (from @univerjs/sheets-numfmt-ui). */
const NUMFMT_RIBBON_ITEMS = [
  ['sheet.operation.open.numfmt.panel', 3],
  ['sheet.command.numfmt.set.percent', 3.1],
  ['sheet.command.numfmt.set.currency', 3.2],
  ['sheet.command.numfmt.add.decimal.command', 3.3],
  ['sheet.command.numfmt.subtract.decimal.command', 3.4],
];

/**
 * Univer registers Number Format under Start → Layout (late order), so it
 * collapses into the "…" overflow on typical widths. Move the built-in controls
 * into Start → Format (after font size, before bold) like Excel / Google Sheets.
 *
 * Numfmt menus are registered in the plugin's onRendered lifecycle — after
 * createUniver returns — so we retry on menuChanged$ until they appear.
 */
function promoteNumfmtToolbarControls(univer) {
  try {
    const menuManager = univer.__getInjector?.()?.get?.(IMenuManagerService);
    if (!menuManager) return undefined;

    const findGroup = (node, key) => {
      if (!node || typeof node !== 'object') return null;
      if (Object.prototype.hasOwnProperty.call(node, key)) return node[key];
      for (const value of Object.values(node)) {
        if (!value || typeof value !== 'object' || value.menuItemFactory) continue;
        const found = findGroup(value, key);
        if (found) return found;
      }
      return null;
    };

    const tryPromote = () => {
      const root = menuManager._menu;
      const layout = findGroup(root, RibbonStartGroup.LAYOUT);
      const format = findGroup(root, RibbonStartGroup.FORMAT);
      const panelId = NUMFMT_RIBBON_ITEMS[0][0];

      // Already promoted into Format.
      if (format?.[panelId]?.menuItemFactory) return true;
      // Not registered yet (onRendered hasn't run).
      if (!layout?.[panelId]?.menuItemFactory) return false;

      if (format) {
        for (const [id, order] of NUMFMT_RIBBON_ITEMS) {
          const item = layout[id];
          if (!item?.menuItemFactory) continue;
          format[id] = { ...item, order };
          delete layout[id];
        }
        menuManager.menuChanged$?.next?.();
        return true;
      }

      // Fallback: keep them in Layout but pull ahead of align/wrap.
      menuManager.mergeMenu({
        [RibbonStartGroup.LAYOUT]: Object.fromEntries(
          NUMFMT_RIBBON_ITEMS.map(([id], index) => [id, { order: -5 + index * 0.1 }]),
        ),
      });
      return true;
    };

    if (tryPromote()) return undefined;

    const sub = menuManager.menuChanged$?.subscribe?.(() => {
      if (tryPromote()) sub?.unsubscribe?.();
    });
    // Safety timeout so we don't leak a subscription if menus never appear.
    const timer = setTimeout(() => sub?.unsubscribe?.(), 5000);
    return () => {
      clearTimeout(timer);
      sub?.unsubscribe?.();
    };
  } catch {
    // Toolbar customization is best-effort; formatting still works via API/form.
    return undefined;
  }
}

/**
 * Self-contained live spreadsheet editor (Univer). Saves to the API whenever
 * `sheetId` is set. Does not depend on any dashboard-specific code.
 *
 * Mount contract: render with `key={sheetId}` at the call site so switching
 * sheets forces a full unmount/remount instead of reusing one Univer instance.
 *
 * Omitting `sheetId` puts the editor in local-only mode: `doSave` no-ops (no
 * network call), so it's editable in place with no backing sheet yet — used
 * by the "Browse Theta Sheets" picker to let the user edit a sheet before it
 * has been saved. Callers in that mode read edits via the `getGrid()` ref
 * handle instead of `onSaved`.
 */
const SpreadsheetEditor = forwardRef(function SpreadsheetEditor({
  sheetId,
  version,
  initialData,
  onChange,
  onSaved,
  onValidation,
  onDirty,
  onSheetRenamed,
  onSheetsChange,
  onSheetDeleted,
  hideToolbar = false,
  height = '600px',
}, ref) {
  const containerRef = useRef(null);
  const univerRef = useRef(null);
  const univerAPIRef = useRef(null);
  const changeDisposableRef = useRef(null);
  const renameDisposableRef = useRef(null);
  const debounceTimerRef = useRef(null);
  const latestVersionRef = useRef(version);
  const savingRef = useRef(false);
  const dirtyRef = useRef(false);
  const fileInputRef = useRef(null);
  const [remountKey, setRemountKey] = useState(0);
  const [localData, setLocalData] = useState(() => initialData || blankGrid());
  // Excel-style sheet-tab context menu (portaled above Theta overlays)
  const [sheetMenu, setSheetMenu] = useState(null); // { x, y, sheetId, sheetName, canDelete }
  const [univerAPI, setUniverAPI] = useState(null);
  const [addRecordOpen, setAddRecordOpen] = useState(false);
  const [editingRow, setEditingRow] = useState(null);
  const [editFormValues, setEditFormValues] = useState(null);
  const [activeSheetName, setActiveSheetName] = useState('');
  const formConfig = useFormConfig(
    univerAPI,
    activeSheetName ? { sheetName: activeSheetName } : null,
  );
  const formConfigRef = useRef(formConfig);
  formConfigRef.current = formConfig;
  const editingRowRef = useRef(null);
  editingRowRef.current = editingRow;
  const addRecordOpenRef = useRef(false);
  addRecordOpenRef.current = addRecordOpen;

  const syncSheetsToParent = () => {
    const workbook = univerAPIRef.current?.getActiveWorkbook();
    if (!workbook) return;
    const grid = extractGridFromUniverWorkbook(workbook, localData?.name || 'Theta Sheets');
    onSheetsChange?.(grid.sheets || []);
  };

  useImperativeHandle(ref, () => ({
    getGrid: () => {
      const workbook = univerAPIRef.current?.getActiveWorkbook();
      if (!workbook) return null;
      return extractGridFromUniverWorkbook(workbook, localData?.name || 'Theta Sheets');
    },
    setActiveSheetByName: (name) => {
      const workbook = univerAPIRef.current?.getActiveWorkbook();
      const sheet = workbook?.getSheetByName(name);
      if (sheet) workbook.setActiveSheet(sheet);
    },
    renameSheetByName: (oldName, newName) => {
      const next = String(newName ?? '').trim();
      const prev = String(oldName ?? '').trim();
      if (!prev || !next || prev === next) return false;
      const workbook = univerAPIRef.current?.getActiveWorkbook();
      const sheet = workbook?.getSheetByName(prev);
      if (!sheet?.setName) return false;
      sheet.setName(next);
      onSheetRenamed?.({ oldName: prev, newName: next });
      onDirty?.();
      return true;
    },
  }), [onDirty, onSheetRenamed]);

  useEffect(() => {
    latestVersionRef.current = version;
  }, [version]);

  // Keep Univer popups/dropdowns above the Theta Sheets full-screen overlay (z=1650).
  // Formats / font / color menus use Radix portals with univer-z-[1080], which otherwise
  // open "invisibly" behind the overlay (click looks like a no-op).
  useEffect(() => {
    const style = document.createElement('style');
    style.setAttribute('data-theta-univer-z', '1');
    style.textContent = `
      [id^="univer-popup-portal"] { z-index: 10000 !important; }
      [id^="univer-popup-portal"] * { z-index: inherit; }
      [data-slot="dropdown-menu-content"],
      [data-slot="dropdown-menu-sub-content"],
      [data-slot="popover-content"],
      [data-slot="select-content"],
      [data-radix-popper-content-wrapper] {
        z-index: 10000 !important;
      }
    `;
    document.head.appendChild(style);
    return () => style.remove();
  }, []);

  useEffect(() => {
    if (!containerRef.current) return undefined;

    const mountEl = document.createElement('div');
    mountEl.style.height = '100%';
    mountEl.style.width = '100%';
    containerRef.current.appendChild(mountEl);

    const { univer, univerAPI: api } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: mergeLocales(
          UniverPresetSheetsCoreEnUS,
          UniverPresetSheetsDataValidationEnUS,
        ),
      },
      presets: [
        UniverSheetsCorePreset({
          container: mountEl,
        }),
        UniverSheetsDataValidationPreset(),
      ],
    });
    univerRef.current = univer;
    univerAPIRef.current = api;
    setUniverAPI(api);
    const disposeNumfmtPromote = promoteNumfmtToolbarControls(univer);

    // Restore last preferred date format for auto-applied date columns.
    loadStoredDateFormat();
    // Ignore bootstrap mutations (coerce/format restore) so open doesn't autosave.
    let suppressDirty = true;

    const workbookData = toUniverWorkbookData(localData);
    api.createWorkbook(workbookData);
    const liveWorkbook = api.getActiveWorkbook?.();
    // Dates must be Excel serials (not "yyyy-MM-dd" strings) for Formats to change display.
    coerceDateColumnValuesInWorkbook(liveWorkbook);
    // Defaults for date columns, then restore any Formats the user previously saved.
    applyDateNumberFormatsToWorkbook(liveWorkbook, getActiveDateFormat());
    applyColumnFormatsFromGrid(liveWorkbook, localData);
    setActiveSheetName(liveWorkbook?.getActiveSheet?.()?.getSheetName?.() || '');
    // Allow a tick for Univer to finish applying restored formats.
    setTimeout(() => { suppressDirty = false; }, 0);

    const syncActiveSheetName = () => {
      const name = api.getActiveWorkbook()?.getActiveSheet?.()?.getSheetName?.() || '';
      setActiveSheetName(name);
    };

    const markDirtyAndSave = () => {
      if (suppressDirty) return;
      dirtyRef.current = true;
      onDirty?.();
      scheduleDebouncedSave();
    };

    changeDisposableRef.current = api.addEvent(api.Event.SheetValueChanged, markDirtyAndSave);

    // Number-format changes use SetNumfmtMutation and do NOT fire SheetValueChanged.
    // Listen for numfmt commands so Formats edits autosave like cell edits.
    let numfmtCommandDisposable = null;
    try {
      const commandService = univer.__getInjector?.()?.get?.(ICommandService);
      numfmtCommandDisposable = commandService?.onCommandExecuted?.((commandInfo) => {
        const id = String(commandInfo?.id || '');
        if (!id.includes('numfmt')) return;
        markDirtyAndSave();
      });
    } catch {
      numfmtCommandDisposable = null;
    }

    const activatedDisposable = api.addEvent?.(api.Event.ActiveSheetChanged, syncActiveSheetName)
      || api.addEvent?.(api.Event.SheetActivated, syncActiveSheetName);

    const notifyRename = (oldName, newName) => {
      const prev = String(oldName ?? '').trim();
      const next = String(newName ?? '').trim();
      if (!prev || !next || prev === next) return;
      onSheetRenamed?.({ oldName: prev, newName: next });
      onDirty?.();
    };

    const beforeDisposable = api.addEvent(api.Event.BeforeSheetNameChange, (params) => {
      notifyRename(params?.oldName, params?.newName);
    });

    const changedDisposable = api.addEvent?.(api.Event.SheetNameChanged, (params) => {
      const next = params?.newName ?? params?.worksheet?.getSheetName?.();
      const prev = params?.oldName;
      if (prev && next) notifyRename(prev, next);
    });

    // Capture sheet-tab right-clicks and show our Excel-style menu.
    // Univer's native menu portals at ~z-index 1020 and sits under the Theta
    // browser overlay (z=1650), so it looks like "right-click does nothing".
    const onTabContextMenu = (event) => {
      const tab = event.target?.closest?.('[data-u-comp="slide-tab-item"]');
      if (!tab || !mountEl.contains(tab)) return;
      event.preventDefault();
      event.stopPropagation();

      const sheetIdFromTab = tab.dataset?.id;
      const workbook = api.getActiveWorkbook();
      if (!workbook || !sheetIdFromTab) return;

      const sheets = workbook.getSheets?.() || [];
      const target = sheets.find((s) => s.getSheetId?.() === sheetIdFromTab) || workbook.getActiveSheet?.();
      if (!target) return;

      workbook.setActiveSheet?.(target);
      setSheetMenu({
        x: event.clientX,
        y: event.clientY,
        sheetId: target.getSheetId?.(),
        sheetName: String(target.getSheetName?.() || ''),
        canDelete: sheets.length > 1,
      });
    };
    mountEl.addEventListener('contextmenu', onTabContextMenu, true);

    renameDisposableRef.current = {
      dispose: () => {
        beforeDisposable?.dispose?.();
        changedDisposable?.dispose?.();
        activatedDisposable?.dispose?.();
        mountEl.removeEventListener('contextmenu', onTabContextMenu, true);
      },
    };

    return () => {
      disposeNumfmtPromote?.();
      numfmtCommandDisposable?.dispose?.();
      changeDisposableRef.current?.dispose?.();
      changeDisposableRef.current = null;
      renameDisposableRef.current?.dispose?.();
      renameDisposableRef.current = null;
      clearTimeout(debounceTimerRef.current);

      // Flush pending Formats/value edits before Univer is disposed so reopen
      // doesn't lose the last change that was still waiting on debounce.
      if (dirtyRef.current && sheetId && univerAPIRef.current) {
        try {
          const workbook = univerAPIRef.current.getActiveWorkbook?.();
          if (workbook) {
            const grid = extractGridFromUniverWorkbook(workbook, localData?.name || 'Theta Sheets');
            const canSave = !gridHasRequiredHeaders(localData) || gridHasRequiredHeaders(grid);
            if (canSave && grid?.sheets?.length) {
              onChange?.(grid);
              sheetService.saveSheet(sheetId, grid, latestVersionRef.current)
                .then((saved) => {
                  latestVersionRef.current = saved.version;
                  dirtyRef.current = false;
                  onSaved?.(saved);
                })
                .catch(() => {});
            }
          }
        } catch {
          // best-effort flush
        }
      }

      setSheetMenu(null);
      setAddRecordOpen(false);
      setEditingRow(null);
      setEditFormValues(null);
      setActiveSheetName('');
      setUniverAPI(null);
      const univerInstance = univerRef.current;
      univerRef.current = null;
      univerAPIRef.current = null;
      setTimeout(() => {
        univerInstance?.dispose();
        mountEl.remove();
      }, 0);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheetId, remountKey]);

  // Open Edit Record panel when a data row is selected in the sheet.
  useEffect(() => {
    if (!univerAPI?.Event?.SelectionMoveEnd || typeof univerAPI.addEvent !== 'function') {
      return undefined;
    }

    const disposable = univerAPI.addEvent(univerAPI.Event.SelectionMoveEnd, (params) => {
      const cfg = formConfigRef.current;
      if (!cfg?.ready || !cfg.fields?.length) return;

      const selections = params?.selections;
      if (!Array.isArray(selections) || !selections.length) return;

      const range = selections[0];
      const row = Number(range?.startRow ?? NaN);
      // Row 0 is the header — ignore; only data rows open the editor.
      if (!Number.isFinite(row) || row < 1) return;

      // Same row already open for edit — keep panel as-is (don't reset while typing).
      if (editingRowRef.current === row && addRecordOpenRef.current) return;

      const values = readRowValues(univerAPI, cfg, row);
      if (!rowHasData(values)) return;

      setEditingRow(row);
      setEditFormValues(values);
      setAddRecordOpen(true);
    });

    return () => disposable?.dispose?.();
  }, [univerAPI]);

  const closeRecordPanel = () => {
    setAddRecordOpen(false);
    setEditingRow(null);
    setEditFormValues(null);
  };

  const openAddRecordPanel = () => {
    setEditingRow(null);
    setEditFormValues(null);
    setAddRecordOpen(true);
  };

  function scheduleDebouncedSave() {
    clearTimeout(debounceTimerRef.current);
    debounceTimerRef.current = setTimeout(doSave, SAVE_DEBOUNCE_MS);
  }

  async function doSave() {
    if (!sheetId || !univerAPIRef.current || savingRef.current) return;
    const workbook = univerAPIRef.current.getActiveWorkbook();
    if (!workbook) return;
    const grid = extractGridFromUniverWorkbook(workbook, localData?.name || 'Theta Sheets');

    // Never persist a bad extract over a sheet that previously had required
    // headers — that was wiping Activity ID / Activity Name on edit.
    if (gridHasRequiredHeaders(localData) && !gridHasRequiredHeaders(grid)) {
      onValidation?.(validateSheetGrid(grid));
      toast.error('Could not read sheet headers after edit. Changes were not saved — try again.');
      return;
    }

    onChange?.(grid);
    onValidation?.(validateSheetGrid(grid));

    savingRef.current = true;
    try {
      const saved = await sheetService.saveSheet(sheetId, grid, latestVersionRef.current);
      latestVersionRef.current = saved.version;
      setLocalData(grid);
      dirtyRef.current = false;
      onSaved?.(saved);
    } catch (err) {
      if (err?.response?.status === 409) {
        toast.error('Sheet was updated in another session. Please reload to see the latest version.');
      } else {
        try {
          const retried = await sheetService.saveSheet(sheetId, grid, latestVersionRef.current);
          latestVersionRef.current = retried.version;
          setLocalData(grid);
          dirtyRef.current = false;
          onSaved?.(retried);
        } catch {
          toast.error('Could not save changes. Please check your connection.');
        }
      }
    } finally {
      savingRef.current = false;
    }
  }

  function triggerImport() {
    fileInputRef.current?.click();
  }

  function handleImportFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const sheets = wb.SheetNames.map(name => {
          const ws = wb.Sheets[name];
          const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          let headerIdx = 0, bestCount = -1;
          for (let i = 0; i < Math.min(10, raw.length); i++) {
            const count = (raw[i] || []).filter(c => String(c ?? '').trim() !== '').length;
            if (count > bestCount) { bestCount = count; headerIdx = i; }
          }
          const headers = (raw[headerIdx] || []).map(h => String(h ?? '').trim());
          const rows = raw.slice(headerIdx + 1).filter(row => row.some(c => String(c ?? '').trim() !== ''));
          return { name, headers, rows };
        });
        setLocalData({ name: 'Theta Sheets', sheets });
        setRemountKey(k => k + 1);
        setTimeout(() => doSave(), 150);
      } catch {
        toast.error('Could not read that file. Please check the format and try again.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  }

  function closeSheetMenu() {
    setSheetMenu(null);
  }

  function startInlineTabRename(sheetId, oldName) {
    const root = containerRef.current;
    if (!root) return false;
    const tab = root.querySelector(`[data-u-comp="slide-tab-item"][data-id="${CSS.escape(sheetId)}"]`);
    const label = tab?.querySelector('span');
    if (!label) return false;

    const workbook = univerAPIRef.current?.getActiveWorkbook();
    const sheet = workbook?.getSheets?.()?.find((s) => s.getSheetId?.() === sheetId)
      || workbook?.getSheetByName(oldName);
    if (!sheet) return false;

    label.setAttribute('contenteditable', 'true');
    label.style.outline = '1px solid #86efac';
    label.style.borderRadius = '3px';
    label.style.padding = '0 2px';
    label.focus();

    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(label);
    selection?.removeAllRanges();
    selection?.addRange(range);

    let finished = false;
    const finish = (commit) => {
      if (finished) return;
      finished = true;
      label.removeAttribute('contenteditable');
      label.style.outline = '';
      label.style.borderRadius = '';
      label.style.padding = '';
      label.removeEventListener('keydown', onKeyDown);
      label.removeEventListener('focusout', onFocusOut);

      const trimmed = String(label.innerText || '').replace(/\s+/g, ' ').trim();
      if (!commit || !trimmed || trimmed === oldName) {
        label.innerText = oldName;
        return;
      }
      const names = (workbook.getSheets?.() || [])
        .map((s) => s.getSheetName?.())
        .filter((n) => n && n !== oldName);
      if (names.includes(trimmed)) {
        label.innerText = oldName;
        toast.error('A sheet with that name already exists');
        return;
      }
      sheet.setName(trimmed);
      onSheetRenamed?.({ oldName, newName: trimmed });
      onDirty?.();
      syncSheetsToParent();
    };

    const onKeyDown = (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') {
        e.preventDefault();
        finish(true);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        finish(false);
      }
    };
    const onFocusOut = () => finish(true);

    label.addEventListener('keydown', onKeyDown);
    label.addEventListener('focusout', onFocusOut);
    return true;
  }

  function handleMenuRename() {
    if (!sheetMenu) return;
    const { sheetId, sheetName } = sheetMenu;
    closeSheetMenu();
    const workbook = univerAPIRef.current?.getActiveWorkbook();
    const sheet = workbook?.getSheets?.()?.find((s) => s.getSheetId?.() === sheetId)
      || workbook?.getSheetByName(sheetName)
      || workbook?.getActiveSheet?.();
    if (!sheet) return;
    workbook.setActiveSheet?.(sheet);
    const id = sheet.getSheetId?.() || sheetId;
    const oldName = String(sheet.getSheetName?.() || sheetName);
    // Let the context menu unmount first, then edit the tab label inline.
    setTimeout(() => {
      if (!startInlineTabRename(id, oldName)) {
        toast.error('Could not start rename on this sheet tab');
      }
    }, 0);
  }

  function handleMenuDuplicate() {
    if (!sheetMenu) return;
    const { sheetName } = sheetMenu;
    closeSheetMenu();
    const workbook = univerAPIRef.current?.getActiveWorkbook();
    const sheet = workbook?.getSheetByName(sheetName) || workbook?.getActiveSheet?.();
    if (!workbook || !sheet) return;
    try {
      workbook.duplicateSheet?.(sheet) || workbook.duplicateActiveSheet?.();
      onDirty?.();
      // Let Univer finish activating the duplicated sheet, then sync names.
      setTimeout(() => syncSheetsToParent(), 0);
    } catch {
      toast.error('Could not duplicate sheet');
    }
  }

  function handleMenuDelete() {
    if (!sheetMenu) return;
    const { sheetName, sheetId: tabSheetId, canDelete } = sheetMenu;
    if (!canDelete) {
      toast.error('Cannot delete the only remaining sheet.');
      closeSheetMenu();
      return;
    }
    closeSheetMenu();

    const workbook = univerAPIRef.current?.getActiveWorkbook();
    if (!workbook) {
      toast.error('Could not delete sheet');
      return;
    }

    const target = (workbook.getSheets?.() || []).find((s) => s.getSheetId?.() === tabSheetId)
      || workbook.getSheetByName(sheetName);
    if (!target) {
      // Still remove from parent state if tab mapping is stale.
      onSheetDeleted?.(sheetName, null);
      return;
    }

    const nameToRemove = String(target.getSheetName?.() || sheetName);

    // Snapshot remaining sheets first (Univer delete can be blocked by permissions).
    const before = extractGridFromUniverWorkbook(workbook, localData?.name || 'Theta Sheets');
    const remaining = (before.sheets || []).filter((s) => s.name !== nameToRemove);
    if (remaining.length === 0) {
      toast.error('Cannot delete the only remaining sheet.');
      return;
    }

    let removedInUniver = false;
    try {
      const result = workbook.deleteSheet(target);
      removedInUniver = result !== false
        && !(workbook.getSheets?.() || []).some((s) => s.getSheetId?.() === tabSheetId);
    } catch {
      removedInUniver = false;
    }

    // Fallback: remount workbook without that sheet (bypasses Univer permission blocks).
    if (!removedInUniver) {
      setLocalData({ name: localData?.name || 'Theta Sheets', sheets: remaining });
      setRemountKey((k) => k + 1);
    }

    // Excel Online style: parent removes from left list + autosaves immediately.
    onSheetDeleted?.(nameToRemove, remaining);
    onDirty?.();
    if (sheetId) {
      setTimeout(() => scheduleDebouncedSave(), 50);
    }
  }

  const menuItemStyle = (disabled = false) => ({
    display: 'block',
    width: '100%',
    textAlign: 'left',
    padding: '8px 14px',
    border: 'none',
    background: 'transparent',
    color: disabled ? '#94a3b8' : '#0f172a',
    fontSize: 13,
    cursor: disabled ? 'not-allowed' : 'pointer',
  });

  return (
    <div style={{ height, width: '100%', display: 'flex', flexDirection: 'column', position: 'relative' }}>
      <div style={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'center', gap: 8, padding: '0 0 8px' }}>
        {!hideToolbar && (
          <>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleImportFile}
              style={{ display: 'none' }}
            />
            <button
              type="button"
              onClick={triggerImport}
              style={{
                padding: '6px 14px', background: '#f1f5f9', color: '#334155',
                border: '1px solid #e2e8f0', borderRadius: 7, fontSize: 12.5,
                fontWeight: 600, cursor: 'pointer',
              }}
            >
              Import file
            </button>
          </>
        )}
        <AddRecordButton
          disabled={!formConfig.ready}
          onClick={openAddRecordPanel}
        />
      </div>

      <AddRecordPanel
        open={addRecordOpen}
        onClose={closeRecordPanel}
        univerAPI={univerAPI}
        config={formConfig}
        editRow={editingRow}
        initialFormValues={editFormValues}
        onRecordAdded={() => {
          // Ensure local browser + autosave paths pick up the new/edited row even if
          // Univer doesn't emit SheetValueChanged for sparse setValues.
          onDirty?.();
          syncSheetsToParent();
          if (sheetId) scheduleDebouncedSave();
        }}
      >
        <div ref={containerRef} style={{ height: '100%', width: '100%', minHeight: 0 }} />
      </AddRecordPanel>

      {sheetMenu && createPortal(
        <>
          <div
            onMouseDown={(e) => {
              e.preventDefault();
              closeSheetMenu();
            }}
            onContextMenu={(e) => { e.preventDefault(); closeSheetMenu(); }}
            style={{ position: 'fixed', inset: 0, zIndex: 20000 }}
          />
          <div
            role="menu"
            onMouseDown={(e) => e.stopPropagation()}
            style={{
              position: 'fixed',
              left: Math.min(sheetMenu.x, window.innerWidth - 200),
              top: Math.min(sheetMenu.y, window.innerHeight - 180),
              zIndex: 20001,
              minWidth: 180,
              background: '#fff',
              border: '1px solid #e2e8f0',
              borderRadius: 8,
              boxShadow: '0 10px 30px rgba(15, 23, 42, 0.18)',
              padding: '4px 0',
            }}
          >
            <button
              type="button"
              role="menuitem"
              disabled={!sheetMenu.canDelete}
              onMouseDown={(e) => e.stopPropagation()}
              onClick={(e) => {
                e.preventDefault();
                e.stopPropagation();
                handleMenuDelete();
              }}
              style={menuItemStyle(!sheetMenu.canDelete)}
              onMouseEnter={(e) => { if (sheetMenu.canDelete) e.currentTarget.style.background = '#f1f5f9'; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = 'transparent'; }}
            >
              Delete
            </button>
            <button
              type="button"
              role="menuitem"
              onMouseDown={(e) => e.stopPropagation()}
              onClick={(e) => {
                e.preventDefault();
                e.stopPropagation();
                handleMenuDuplicate();
              }}
              style={menuItemStyle()}
              onMouseEnter={(e) => { e.currentTarget.style.background = '#f1f5f9'; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = 'transparent'; }}
            >
              Duplicate
            </button>
            <button
              type="button"
              role="menuitem"
              onMouseDown={(e) => e.stopPropagation()}
              onClick={(e) => {
                e.preventDefault();
                e.stopPropagation();
                handleMenuRename();
              }}
              style={menuItemStyle()}
              onMouseEnter={(e) => { e.currentTarget.style.background = '#f1f5f9'; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = 'transparent'; }}
            >
              Rename
            </button>
          </div>
        </>,
        document.body
      )}
    </div>
  );
});

export default SpreadsheetEditor;
