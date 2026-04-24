import React, {
  useRef,
  useEffect,
  Fragment,
  useState,
  memo,
  useCallback,
} from 'react';
import type { IController, EventData, ModalValue } from '../../types';
import { getHitInfo, DEFAULT_POSITION, CELL_HEIGHT, CELL_WIDTH } from '../../util';
import styles from './index.module.css';
import { useCoreStore, useExcel } from '../../containers/store';
import { ScrollBar } from './ScrollBar';
import { ContextMenu } from './ContextMenu';
import { initCanvas } from './util';
import { BottomBar } from './BottomBar';
import FloatElementContainer from '../FloatElement';
import handlerList from './event';
import Modal from './modal';
import { Collaboration } from './Collaboration';
import { getResizeHit, type ResizeHit } from '../../canvas/resize';
import { MainCanvas } from '../../canvas';

function getEventData(
  event: React.PointerEvent<HTMLCanvasElement>,
  controller: IController,
): EventData {
  const rect = controller.getCanvasSize();
  const { clientX = 0, clientY = 0 } = event;
  const x = clientX - rect.left;
  const y = clientY - rect.top;

  const position = getHitInfo(controller, x, y);
  const result: EventData = {
    position,
    x,
    y,
    controller,
  };

  return result;
}

type ResizeDrag = {
  hit: ResizeHit;
  startPointer: number; // clientX (col) or clientY (row)
  startSize: number;
};

const MIN_RESIZE_SIZE = 4;

export const CanvasContainer = memo(() => {
  const { controller } = useExcel();
  const [modalState, setModalState] = useState<ModalValue | null>(null);
  const activeUuid = useCoreStore((s) => s.activeUuid);
  const floatElementList = useCoreStore((s) => s.drawings);
  const [menuPosition, setMenuPosition] = useState({
    top: DEFAULT_POSITION,
    left: DEFAULT_POSITION,
  });
  const [cursor, setCursor] = useState<string>('');

  const ref = useRef<HTMLCanvasElement>(null);
  const resizeDragRef = useRef<ResizeDrag | null>(null);
  useEffect(() => {
    if (!ref.current) {
      return;
    }
    return initCanvas(controller, ref.current);
  }, []);
  const handleContextMenu = useCallback(
    (event: React.MouseEvent<HTMLCanvasElement>) => {
      event.preventDefault();
      setMenuPosition({
        top: event.clientY,
        left: event.clientX,
      });
    },
    [],
  );
  const hideContextMenu = () => {
    setMenuPosition({
      top: DEFAULT_POSITION,
      left: DEFAULT_POSITION,
    });
  };
  const handlePointerMove = useCallback(
    (event: React.PointerEvent<HTMLCanvasElement>) => {
      // Resize-drag in progress: update size live and skip other handlers.
      const drag = resizeDragRef.current;
      if (drag) {
        const delta =
          drag.hit.axis === 'col'
            ? event.clientX - drag.startPointer
            : event.clientY - drag.startPointer;
        const next = Math.max(MIN_RESIZE_SIZE, drag.startSize + delta);
        if (drag.hit.axis === 'col') {
          controller.setColWidth(drag.hit.index, next);
        } else {
          controller.setRowHeight(drag.hit.index, next);
        }
        return;
      }
      // Hover cursor feedback when not pressing a button.
      if (event.buttons <= 0) {
        const { x, y } = eventCoords(event, controller);
        const hit = getResizeHit(controller, x, y);
        const nextCursor = hit
          ? hit.axis === 'col'
            ? 'col-resize'
            : 'row-resize'
          : '';
        setCursor((prev) => (prev === nextCursor ? prev : nextCursor));
        return;
      }
      /* jscpd:ignore-start */
      const data = getEventData(event, controller);
      for (const handler of handlerList) {
        const r = handler.pointerMove(data, event);
        if (r) {
          if (typeof r !== 'boolean') {
            setModalState(r);
          }
          break;
        }
      }
      /* jscpd:ignore-end */
    },
    [],
  );
  const handlePointerDown = useCallback(
    (event: React.PointerEvent<HTMLCanvasElement>) => {
      if (event.buttons <= 0) {
        return;
      }
      // Resize hit-test takes precedence over selection/filter handlers.
      const { x, y } = eventCoords(event, controller);
      const hit = getResizeHit(controller, x, y);
      if (hit) {
        const startSize =
          hit.axis === 'col'
            ? controller.getColWidth(hit.index)
            : controller.getRowHeight(hit.index);
        resizeDragRef.current = {
          hit,
          startPointer: hit.axis === 'col' ? event.clientX : event.clientY,
          startSize,
        };
        event.currentTarget.setPointerCapture?.(event.pointerId);
        event.preventDefault();
        return;
      }
      setModalState(null);
      const data = getEventData(event, controller);
      for (const handler of handlerList) {
        const r = handler.pointerDown(data, event);
        if (r) {
          if (typeof r !== 'boolean') {
            setModalState(r);
          }
          break;
        }
      }
    },
    [],
  );
  const handlePointerUp = useCallback(
    (event: React.PointerEvent<HTMLCanvasElement>) => {
      if (resizeDragRef.current) {
        resizeDragRef.current = null;
        event.currentTarget.releasePointerCapture?.(event.pointerId);
      }
    },
    [],
  );
  const handleDoubleClick = useCallback(
    (event: React.MouseEvent<HTMLCanvasElement>) => {
      const { x, y } = eventCoords(event, controller);
      const hit = getResizeHit(controller, x, y);
      if (!hit) return;
      event.preventDefault();
      const mc = MainCanvas.instance;
      if (!mc) return;
      if (hit.axis === 'col') {
        const measured = mc.getMeasuredColWidth(hit.index);
        controller.setColWidth(hit.index, measured ?? CELL_WIDTH);
      } else {
        const measured = mc.getMeasuredRowHeight(hit.index);
        controller.setRowHeight(hit.index, measured ?? CELL_HEIGHT);
      }
    },
    [],
  );
  return (
    <Fragment>
      <div
        className={styles['canvas-container']}
        data-testid="canvas-container"
      >
        <canvas
          className={styles['canvas-content']}
          onContextMenu={handleContextMenu}
          onPointerMove={handlePointerMove}
          onPointerDown={handlePointerDown}
          onPointerUp={handlePointerUp}
          onPointerCancel={handlePointerUp}
          onDoubleClick={handleDoubleClick}
          ref={ref}
          style={cursor ? { cursor } : undefined}
          data-testid="canvas-main"
        />
        <ScrollBar />
        <BottomBar />
        {floatElementList.length > 0 && (
          <FloatElementContainer
            floatElementList={floatElementList}
            activeUuid={activeUuid}
          />
        )}
        <Collaboration />
      </div>
      {menuPosition.top >= 0 && menuPosition.left >= 0 && (
        <ContextMenu {...menuPosition} hideContextMenu={hideContextMenu} />
      )}
      {modalState && <Modal {...modalState} hide={() => setModalState(null)} />}
    </Fragment>
  );
});

function eventCoords(
  event:
    | React.PointerEvent<HTMLCanvasElement>
    | React.MouseEvent<HTMLCanvasElement>,
  controller: IController,
) {
  const rect = controller.getCanvasSize();
  const { clientX = 0, clientY = 0 } = event;
  return { x: clientX - rect.left, y: clientY - rect.top };
}
CanvasContainer.displayName = 'CanvasContainer';

export default CanvasContainer;
