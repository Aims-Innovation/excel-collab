import React, { memo, useCallback, useState } from 'react';
import { MdUndo, MdRedo } from 'react-icons/md';
import { Menu, MenuItem, Button } from '../../components';
import { importXLSX, exportToXLSX, exportToCsv, importCSV } from '../Excel';
import styles from './index.module.css';
import { Theme } from './Theme';
import i18n from '../../i18n';
import { I18N } from './I18N';
import { saveAs } from '../../util';
import { useExcel, useCoreStore } from '../store';
import { User } from './User';
import { File } from './File';
import { v4 } from 'uuid';

type Props = {
  leftChildren?: React.ReactNode;
  rightChildren?: React.ReactNode;
  hideNewFile?: boolean;
  hideRenameFile?: boolean;
  hideUser?: boolean;
  hideI18n?: boolean;
  hideTheme?: boolean;
  /** When false, hide all Export menu items (XLSX / CSV / JSON). Defaults to true. */
  canExport?: boolean;
  /** Hide the entire File menu (and the clickable filename widget on the
   *  left). Wins over hideNewFile / hideRenameFile / canExport when set --
   *  intended for view-only sessions where the user lacks edit permission. */
  hideFileMenu?: boolean;
};

export const MenuBarContainer: React.FunctionComponent<Props> = memo(
  ({
    leftChildren,
    rightChildren,
    hideNewFile,
    hideRenameFile,
    hideUser,
    hideI18n,
    hideTheme,
    canExport = true,
    hideFileMenu,
  }) => {
    const { controller, provider } = useExcel();
    const canUndo = useCoreStore((s) => s.canUndo);
    const canRedo = useCoreStore((s) => s.canRedo);
    const [visible, setVisible] = useState(false);
    const handleUndo = useCallback(() => controller.undo(), [controller]);
    const handleRedo = useCallback(() => controller.redo(), [controller]);
    const handleExportXLSX = useCallback(() => {
      exportToXLSX(`excel_${Date.now()}.xlsx`, controller);
    }, []);
    const handleExportCSV = useCallback(() => {
      const text = exportToCsv(controller);
      const blob = new Blob([text], {
        type: 'text/csv;charset=utf-8;',
      });
      saveAs(blob, `excel_${Date.now()}.csv`);
    }, []);
    const handleImportXLSX = useCallback(
      async (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (!file) {
          return;
        }
        const model = await importXLSX(file);
        await controller.fromJSON(model);
        event.target.value = '';
        event.target.blur();
      },
      [],
    );
    const handleImportCSV = useCallback(
      async (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (!file) {
          return;
        }
        await importCSV(file, controller);
        event.target.value = '';
        event.target.blur();
      },
      [],
    );
    const handleExportJSON = useCallback(() => {
      const blob = new Blob([JSON.stringify(controller.toJSON())], {
        type: 'application/json',
      });
      saveAs(blob, `excel_${Date.now()}.json`);
    }, []);
    const handleAddDocument = useCallback(() => {
      const docId = v4();
      provider?.addDocument?.(docId);
    }, []);
    return (
      <div className={styles['menubar-container']} data-testid="menubar">
        <div className={styles['menubar-menu']}>
          <div className={styles['quick-access']} data-testid="menubar-quick-access">
            <Button
              type="toolbar"
              disabled={!canUndo}
              onClick={handleUndo}
              testId="menubar-undo"
              title={i18n.t('toolbar-undo')}
              className={styles['quick-access-button']}
            >
              <MdUndo className={styles['quick-access-icon']} />
            </Button>
            <Button
              type="toolbar"
              disabled={!canRedo}
              onClick={handleRedo}
              testId="menubar-redo"
              title={i18n.t('toolbar-redo')}
              className={styles['quick-access-button']}
            >
              <MdRedo className={styles['quick-access-icon']} />
            </Button>
          </div>
          {!hideFileMenu && !hideRenameFile && (
            <File visible={visible} setVisible={setVisible} />
          )}
          {!hideFileMenu && (
            <Menu
              label={i18n.t('file')}
              className={styles.menu}
              testId="menubar-excel"
            >
              {!hideNewFile && (
                <MenuItem
                  onClick={handleAddDocument}
                  testId="menubar-new-excel"
                >
                  {i18n.t('new-file')}
                </MenuItem>
              )}
              {!hideRenameFile && (
                <MenuItem
                  onClick={() => setVisible(true)}
                  testId="menubar-rename-excel"
                >
                  {i18n.t('rename-file')}
                </MenuItem>
              )}
              <MenuItem testId="menubar-import-xlsx">
                <input
                  type="file"
                  hidden
                  onChange={handleImportXLSX}
                  accept=".xlsx"
                  id="import_xlsx"
                  data-testid="menubar-import-xlsx-input"
                />
                <label htmlFor="import_xlsx">
                  {i18n.t('import', { format: 'XLSX' })}
                </label>
              </MenuItem>
              <MenuItem testId="menubar-import-csv">
                <input
                  type="file"
                  hidden
                  onChange={handleImportCSV}
                  accept=".csv"
                  id="import_csv"
                  data-testid="menubar-import-csv-input"
                />
                <label htmlFor="import_csv">
                  {i18n.t('import', { format: 'CSV' })}
                </label>
              </MenuItem>
              {canExport && (
                <MenuItem
                  onClick={handleExportXLSX}
                  testId="menubar-export-xlsx"
                >
                  {i18n.t('export', { format: 'XLSX' })}
                </MenuItem>
              )}
              {canExport && (
                <MenuItem
                  testId="menubar-export-csv"
                  onClick={handleExportCSV}
                >
                  {i18n.t('export', { format: 'CSV' })}
                </MenuItem>
              )}
              {canExport && (
                <MenuItem
                  testId="menubar-export-json"
                  onClick={handleExportJSON}
                >
                  {i18n.t('export', { format: 'JSON' })}
                </MenuItem>
              )}
            </Menu>
          )}
          {leftChildren}
        </div>
        {rightChildren}
        {!hideUser && <User />}
        {!hideI18n && <I18N />}
        {!hideTheme && <Theme />}
      </div>
    );
  },
);

MenuBarContainer.displayName = 'MenuBarContainer';

export default MenuBarContainer;
