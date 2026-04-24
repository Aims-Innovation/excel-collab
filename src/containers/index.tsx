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
      return;
    }
    const doc = controller.getHooks().doc;
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
        return;
      }
      const changeSet = modelToChangeSet(tran);
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
