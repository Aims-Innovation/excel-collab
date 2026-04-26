import styles from './index.module.css';
import { memo, useEffect, useState, useSyncExternalStore } from 'react';
import FormulaBarContainer from './FormulaBar';
import ToolbarContainer from './ToolBar';
import CanvasContainer from './canvas';
import SheetBarContainer from './SheetBar';
import MenuBarContainer from './MenuBar';
import { useExcel, useUserInfo } from './store';
import { UserItem, SYNC_FLAG } from '../types';
import { modelToChangeSet } from '../util';
import { Loading } from '../components';
import i18n from '../i18n';
import { collaborationLog, perfMeasure } from '../util/debug';

// Membership check for transactions that originate from the Controller
// (via the @transaction() decorator). Kept as a module-level Set so the
// doc.on('update') handler can cheap-check every update without
// rebuilding it on each event.
const LOCAL_ORIGINS: ReadonlySet<unknown> = new Set<unknown>(
  Object.values(SYNC_FLAG),
);

function useCollaboration() {
  const [isLoading, setIsLoading] = useState(true);
  const setFileInfo = useUserInfo((s) => s.setFileInfo);
  const { provider, controller, awareness } = useExcel();
  useEffect(() => {
    let cancelled = false;
    async function init() {
      collaborationLog('useCollaboration init start', { hasProvider: !!provider });
      if (!provider) {
        if (controller.getSheetList().length === 0) {
          controller.addFirstSheet();
        }
        if (!cancelled) setIsLoading(false);
        return;
      }
      setIsLoading(true);
      const doc = controller.getHooks().doc;
      const file = await perfMeasure('provider.getDocument', () =>
        provider.getDocument?.(doc.guid),
      );
      if (cancelled) return;
      const content = file?.content ?? '';
      if (content) {
        // Diagnostic for the "initial load shows no formatting"
        // investigation: the persisted JSON the provider returns
        // *should* include style fields inline on each cell
        // (ModelCellType extends Partial<StyleType>). If styles are
        // missing here, the bug is upstream of the lib (consumer
        // backend dropping fields on save). If styles are present
        // here but don't render, the bug is in the lib's hydrate
        // -> repaint path. Opt in with:
        //   localStorage.setItem('debug', 'collaboration')
        try {
          const parsed = JSON.parse(content);
          const sheets = parsed?.worksheets ?? {};
          const sheetIds = Object.keys(sheets);
          const firstSheet = sheets[sheetIds[0]] ?? {};
          const cellEntries = Object.entries(firstSheet);
          const sampleCell = cellEntries[0];
          const STYLE_KEYS = [
            'fontColor',
            'fillColor',
            'fontFamily',
            'fontSize',
            'isBold',
            'isItalic',
            'underline',
            'numberFormat',
            'borderLeft',
            'borderRight',
            'borderTop',
            'borderBottom',
            'horizontalAlign',
            'verticalAlign',
            'isWrapText',
          ];
          const styledCellCount = cellEntries.reduce((acc, [, c]) => {
            const cell = c as Record<string, unknown>;
            return STYLE_KEYS.some((k) => cell?.[k] !== undefined)
              ? acc + 1
              : acc;
          }, 0);
          collaborationLog('hydrate-content shape', {
            topLevelKeys: Object.keys(parsed),
            sheetCount: sheetIds.length,
            firstSheetCellCount: cellEntries.length,
            firstSheetStyledCellCount: styledCellCount,
            sampleCellKey: sampleCell?.[0],
            sampleCellShape: sampleCell?.[1]
              ? Object.keys(sampleCell[1] as Record<string, unknown>)
              : null,
            sampleCellStyleFieldsSet: sampleCell?.[1]
              ? STYLE_KEYS.filter(
                  (k) =>
                    (sampleCell[1] as Record<string, unknown>)[k] !==
                    undefined,
                )
              : [],
          });
        } catch (err) {
          collaborationLog('hydrate-content shape: parse failed', err);
        }
        await perfMeasure('controller.fromJSON', () =>
          controller.fromJSON(JSON.parse(content)),
        );
      }
      setFileInfo(file?.id ?? doc.guid, file?.name ?? '');
      if (controller.getSheetList().length === 0) {
        controller.addFirstSheet();
      }
      collaborationLog('useCollaboration init done', {
        sheetCount: controller.getSheetList().length,
      });
      setIsLoading(false);
    }
    init();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!awareness) {
      collaborationLog(
        'awareness-wiring: skipped (no awareness provided yet); ' +
          'doc.on(update) is NOT registered until awareness arrives',
      );
      return;
    }
    const doc = controller.getHooks().doc;
    collaborationLog('awareness-wiring: registering doc.on(update)', {
      docGuid: doc.guid,
      docClientId: doc.clientID,
    });
    const awarenessHandler = () => {
      const list: UserItem[] = [];
      for (const item of awareness.getStates().entries()) {
        const [key, value] = item;
        if (!value.range || key === doc.clientID) {
          continue;
        }
        list.push({ clientId: key, range: value.range });
      }
      useUserInfo.getState().setUsers(list);
    };
    awareness.on('update', awarenessHandler);

    const docHandler = (_a: unknown, _b: unknown, _c: unknown, tran: any) => {
      // Skip updates that the Controller originated via @transaction().
      // Controller.emitChange() already dispatches renderChange for its
      // own mutations; without this filter every local write would be
      // followed by a duplicate render -- doubling work and, under
      // rapid interactions like column drag, spiraling into an
      // unresponsive render queue.
      if (LOCAL_ORIGINS.has(tran?.origin)) {
        collaborationLog('docHandler: skip (local origin)', {
          origin: tran?.origin,
        });
        return;
      }
      const changeSet = modelToChangeSet(tran);
      collaborationLog('docHandler: fired', {
        origin: tran?.origin,
        changeSetSize: changeSet.size,
        changeSet: Array.from(changeSet),
        changedTypeCount: tran?.changed?.size ?? 0,
        changedParentTypeCount: tran?.changedParentTypes?.size ?? 0,
      });
      controller.emit('renderChange', { changeSet });
    };
    doc.on('update', docHandler);

    const offRangeChange = controller.on('rangeChange', (range) => {
      awareness.setLocalStateField('range', range);
    });

    return () => {
      awareness.off('update', awarenessHandler);
      doc.off('update', docHandler);
      offRangeChange();
    };
  }, [awareness, controller]);

  return {
    isLoading,
  };
}

const getLanguageSnapshot = () => i18n.current;
const getLanguageServerSnapshot = () => 'en-US';

export type EditorProps = {
  style?: React.CSSProperties;
  menubarLeftChildren?: React.ReactNode;
  menubarRightChildren?: React.ReactNode;
  toolbarChildren?: React.ReactNode;
  sheetBarChildren?: React.ReactNode;
  /** Hide the "New File" entry in the File menu (host app owns document creation). */
  hideNewFile?: boolean;
  /** Hide the "Rename File" menu entry and the clickable filename widget (host app owns document naming). */
  hideRenameFile?: boolean;
  /** Hide the user / avatar widget on the right of the menubar. */
  hideUser?: boolean;
  /** Hide the language switcher on the right of the menubar. */
  hideI18n?: boolean;
  /** Hide the light / dark theme toggle on the right of the menubar. */
  hideTheme?: boolean;
  /** When false, hide all Export menu items (XLSX / CSV / JSON). Defaults to true. */
  canExport?: boolean;
};
const ExcelEditor: React.FunctionComponent<EditorProps> = memo(
  ({
    style,
    menubarLeftChildren,
    menubarRightChildren,
    toolbarChildren,
    sheetBarChildren,
    hideNewFile,
    hideRenameFile,
    hideUser,
    hideI18n,
    hideTheme,
    canExport,
  }) => {
    const { isLoading } = useCollaboration();
    // Subscribe to language here so only the i18n-rendering memo'd children
    // remount when locale changes. CanvasContainer is deliberately NOT keyed
    // — the worker + OffscreenCanvas pipeline must survive; canvas-side
    // i18n strings (context menu, font picker) render on-demand anyway.
    const language = useSyncExternalStore(
      i18n.subscribe,
      getLanguageSnapshot,
      getLanguageServerSnapshot,
    );

    if (isLoading) {
      return <Loading />;
    }

    return (
      <div
        className={styles['app-container']}
        data-testid="app-container"
        style={style}
      >
        <MenuBarContainer
          key={`menubar-${language}`}
          leftChildren={menubarLeftChildren}
          rightChildren={menubarRightChildren}
          hideNewFile={hideNewFile}
          hideRenameFile={hideRenameFile}
          hideUser={hideUser}
          hideI18n={hideI18n}
          hideTheme={hideTheme}
          canExport={canExport}
        />
        <ToolbarContainer key={`toolbar-${language}`}>
          {toolbarChildren}
        </ToolbarContainer>
        <FormulaBarContainer key={`formula-${language}`} />
        <CanvasContainer />
        <SheetBarContainer key={`sheetbar-${language}`}>
          {sheetBarChildren}
        </SheetBarContainer>
      </div>
    );
  },
);

ExcelEditor.displayName = 'ExcelEditor';

export * from './store';
export * from './MenuBar';
export {
  ExcelEditor,
  useCollaboration,
  ToolbarContainer,
  FormulaBarContainer,
  CanvasContainer,
  SheetBarContainer,
};
// Re-export the XLSX / CSV import + export helpers so consumers can
// do their own seed-from-blob / save-to-blob flows without having to
// reach for a subpath import (which the package.json `exports` field
// blocks). v0.1.13.22+.
export { importXLSX, exportToXLSX, importCSV, exportToCsv } from './Excel';
