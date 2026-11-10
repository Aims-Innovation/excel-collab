import React from 'react';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { initController } from '../controller';
import { FormulaBarContainer } from '../containers/FormulaBar';
import { StateContext, useActiveCell, useCoreStore } from '../containers/store';
import { registerGlobalEvent } from '../canvas/event';
import { MainHandler } from '../containers/canvas/event/main';
import { EditorStatus, IController, EventData } from '../types';

let controller: IController;
let unregister: (() => void) | undefined;

beforeAll(() => {
  // initController's default worker is its synchronous formula implementation.
  // Only the instanceof check needs the browser constructor in this DOM test.
  global.Worker = class {} as unknown as typeof Worker;
});

beforeEach(() => {
  controller = initController();
  controller.addFirstSheet();
  controller.setCellValue('Order ID', controller.getActiveRange().range);
  useCoreStore.setState({ editorStatus: EditorStatus.NONE });
  useActiveCell.setState({ value: 'Order ID', displayValue: 'Order ID' });
});

afterEach(() => {
  unregister?.();
  unregister = undefined;
  cleanup();
  controller.getHooks().doc.destroy();
});

function mountFormulaBar() {
  return render(
    <StateContext.Provider value={{ controller }}>
      <FormulaBarContainer />
    </StateContext.Provider>,
  );
}

function doubleClickCell() {
  const handler = new MainHandler();
  const data: EventData = {
    controller,
    x: 100,
    y: 100,
    position: { row: 0, col: 0, marginX: 0, marginY: 0 },
  };
  for (const timeStamp of [1000, 1100]) {
    handler.pointerDown(data, { timeStamp } as React.PointerEvent<HTMLCanvasElement>);
  }
}

describe('Y1-182 read-only cell entry', () => {
  it('does not open a formula editor in Preview, and restores editing when allowed', () => {
    controller.setReadOnly(true);
    mountFormulaBar();
    fireEvent.click(screen.getByTestId('formula-editor-trigger'));
    expect(screen.queryByTestId('formula-editor')).not.toBeInTheDocument();
    expect(useCoreStore.getState().editorStatus).toBe(EditorStatus.NONE);

    act(() => controller.setReadOnly(false));
    fireEvent.click(screen.getByTestId('formula-editor-trigger'));
    expect(screen.getByTestId('formula-editor')).toHaveValue('Order ID');
  });

  it.each(['F2', 'x'])('does not enter text mode on %s while read-only', (key) => {
    controller.setReadOnly(true);
    mountFormulaBar();
    unregister = registerGlobalEvent(controller, () => {});
    fireEvent.keyDown(document.body, { key });
    expect(useCoreStore.getState().editorStatus).toBe(EditorStatus.NONE);
    expect(screen.queryByTestId('formula-editor')).not.toBeInTheDocument();
  });

  it('blocks double-click entry but allows it again after permission restoration', () => {
    controller.setReadOnly(true);
    mountFormulaBar();
    act(doubleClickCell);
    expect(screen.queryByTestId('formula-editor')).not.toBeInTheDocument();
    expect(useCoreStore.getState().editorStatus).toBe(EditorStatus.NONE);

    act(() => controller.setReadOnly(false));
    act(doubleClickCell);
    expect(screen.getByTestId('formula-editor')).toHaveValue('Order ID');
  });

  it('discards an open draft on permission loss and does not resurrect it', () => {
    const { unmount } = mountFormulaBar();
    fireEvent.click(screen.getByTestId('formula-editor-trigger'));
    fireEvent.change(screen.getByTestId('formula-editor'), { target: { value: 'unsaved draft' } });
    act(() => controller.setReadOnly(true));
    expect(screen.queryByTestId('formula-editor')).not.toBeInTheDocument();
    expect(controller.getCell(controller.getActiveRange().range)?.value).toBe('Order ID');
    act(() => controller.setReadOnly(false));
    expect(screen.queryByTestId('formula-editor')).not.toBeInTheDocument();
    fireEvent.click(screen.getByTestId('formula-editor-trigger'));
    expect(screen.getByTestId('formula-editor')).toHaveValue('Order ID');
    unmount();
    expect(controller.getEventLength('readOnlyChange')).toBe(0);
  });

  it('blocks a stale Enter event immediately after revocation, before React commits', () => {
    mountFormulaBar();
    fireEvent.click(screen.getByTestId('formula-editor-trigger'));
    const editor = screen.getByTestId('formula-editor');
    fireEvent.change(editor, { target: { value: 'forbidden draft' } });
    const range = controller.getActiveRange().range;
    act(() => {
      controller.setReadOnly(true);
      editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    });
    expect(controller.getCell(range)?.value).toBe('Order ID');
    expect(controller.getActiveRange().range.row).toBe(range.row);
    expect(screen.queryByTestId('formula-editor')).not.toBeInTheDocument();
  });

  it('keeps read-only arrow navigation available', () => {
    controller.setReadOnly(true);
    unregister = registerGlobalEvent(controller, () => {});
    fireEvent.keyDown(document.body, { key: 'ArrowDown' });
    expect(controller.getActiveRange().range.row).toBe(1);
    expect(useCoreStore.getState().editorStatus).toBe(EditorStatus.NONE);
  });

  it('emits permission changes only when standing changes', () => {
    const listener = jest.fn();
    controller.on('readOnlyChange', listener);
    controller.setReadOnly(true);
    controller.setReadOnly(true);
    controller.setReadOnly(false);
    expect(listener.mock.calls).toEqual([[true], [false]]);
  });
});
