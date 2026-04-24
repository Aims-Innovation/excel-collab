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

type ResizeState =
  | { phase: 'primed'; hit: ResizeHit; startPointer: number; startSize: number; downTime: number }
  | { phase: 'dragging'; hit: ResizeHit; startPointer: number; startSize: number; pointerId: number };

const MIN_RESIZE_SIZE = 4;
const DRAG_THRESHOLD = 3;
const DOUBLE_CLICK_MS = 300;

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
  const resizeRef = useRef<ResizeState | null>(null);
  // Remembers the last resize-adjacent pointerdown so we can detect a
  // double-click ourselves (React's onDoubleClick is unreliable when a
  // pointerdown elsewhere calls preventDefault or captures the pointer).
  const lastResizeDownRef = useRef<{ hit: ResizeHit; time: number } | null>(
    null,
  );
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
  const applyResizeAutofit = useCallback((hit: ResizeHit) => {
    const mc = MainCanvas.instance;
    if (!mc) return;
    if (hit.axis === 'col') {
      const measured = mc.getMeasuredColWidth(hit.index);
      controller.setColWidth(hit.index, measured ?? CELL_WIDTH);
    } else {
      const measured = mc.getMeasuredRowHeight(hit.index);
      controller.setRowHeight(hit.index, measured ?? CELL_HEIGHT);
    }
  }, []);
  const handlePointerMove = useCallback(
    (event: React.PointerEvent<HTMLCanvasElement>) => {
      const state = resizeRef.current;
      if (state) {
        const pointer =
          state.hit.axis === 'col' ? event.clientX : event.clientY;
        const delta = pointer - state.startPointer;
        if (state.phase === 'primed') {
          // Only commit to a resize once the user has actually moved.
          // Under the threshold this is treated as a click, and the
          // selection handler will run on pointerup instead.
          if (Math.abs(delta) < DRAG_THRESHOLD) {
            return;
          }
          resizeRef.current = {
            phase: 'dragging',
            hit: state.hit,
            startPointer: state.startPointer,
            startSize: state.startSize,
            pointerId: event.pointerId,
          };
        }
        const next = Math.max(MIN_RESIZE_SIZE, state.startSize + delta);
        if (state.hit.axis === 'col') {
          controller.setColWidth(state.hit.index, next);
        } else {
          controller.setRowHeight(state.hit.index, next);
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
      const { x, y } = eventCoords(event, controller);
      const hit = getResizeHit(controller, x, y);
      if (hit) {
        const now = event.timeStamp;
        const prev = lastResizeDownRef.current;
        const isDoubleClick =
          prev !== null &&
          now - prev.time < DOUBLE_CLICK_MS &&
          prev.hit.axis === hit.axis &&
          prev.hit.index === hit.index;
        if (isDoubleClick) {
          lastResizeDownRef.current = null;
          resizeRef.current = null;
          applyResizeAutofit(hit);
          return;
        }
        lastResizeDownRef.current = { hit, time: now };
        const startSize =
          hit.axis === 'col'
            ? controller.getColWidth(hit.index)
            : controller.getRowHeight(hit.index);
        resizeRef.current = {
          phase: 'primed',
          hit,
          startPointer: hit.axis === 'col' ? event.clientX : event.clientY,
          startSize,
          downTime: now,
        };
        // Capture the pointer so pointermove/up events still reach us if
        // the cursor leaves the canvas during a drag. We do NOT
        // preventDefault — that suppresses the subsequent click event in
        // some browsers, which broke double-click-to-autofit previously
        // and prevented the "click to select header" selection path.
        try {
          event.currentTarget.setPointerCapture?.(event.pointerId);
        } catch {
          /* capture failed, continue without */
        }
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
    [applyResizeAutofit],
  );
  const handlePointerUp = useCallback(
    (event: React.PointerEvent<HTMLCanvasElement>) => {
      const state = resizeRef.current;
      if (!state) return;
      try {
        event.currentTarget.releasePointerCapture?.(event.pointerId);
      } catch {
        /* not captured — fine */
      }
      const wasDragging = state.phase === 'dragging';
      resizeRef.current = null;
      if (wasDragging) {
        return;
      }
      // Primed but never dragged: treat as a click on the header and
      // fall through to the selection handler now.
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
