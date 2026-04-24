import {
  KeyboardEventItem,
  IController,
  EUnderLine,
  EditorStatus,
  IRange,
  ScrollValue,
} from '../types';
import {
  isMac,
  SHEET_ITEM_TEST_ID_PREFIX,
  FORMULA_EDITOR_ROLE,
  MERGE_CELL_LINE_BREAK,
  LINE_BREAK,
  isMergeContent,
} from '../util';
import { useCoreStore } from '../containers/store';

export function handleTabClick(controller: IController) {
  controller.transaction(() => {
    checkActiveElement(controller);
    controller.setNextActiveCell('right');
    recalculateScroll(controller);
  });
}

export function handleEnterClick(controller: IController) {
  controller.transaction(() => {
    checkActiveElement(controller);
    controller.setNextActiveCell('down');
    recalculateScroll(controller);
  });
}

export function computeScrollRowAndCol(
  controller: IController,
  left: number,
  top: number,
) {
  const oldScroll = controller.getScroll();
  let { row, col } = oldScroll;
  if (oldScroll.top !== top) {
    row = 0;
    let t = top;
    while (t > 0) {
      const a = controller.getRowHeight(row);
      if (a > t) {
        break;
      }
      t -= a;
      row++;
    }
  }
  if (oldScroll.left !== left) {
    col = 0;
    let t = left;
    while (t > 0) {
      const a = controller.getColWidth(col);
      if (a > t) {
        break;
      }
      t -= a;
      col++;
    }
  }
  return {
    row,
    col,
  };
}

export function scrollSheetToView(sheetId: string) {
  const selector = `div[data-testid="${SHEET_ITEM_TEST_ID_PREFIX}${sheetId}"]`;
  const dom = document.querySelector(selector);
  if (dom && typeof dom.scrollIntoView === 'function') {
    dom.scrollIntoView();
  }
}

export function scrollToView(controller: IController, range: IRange) {
  controller.transaction(() => {
    const sheetId = range.sheetId || controller.getCurrentSheetId();
    if (sheetId !== controller.getCurrentSheetId()) {
      controller.setCurrentSheetId(sheetId);
    }
    const sheetInfo = controller.getSheetInfo(sheetId);
    if (!sheetInfo) {
      return true;
    }
    if (
      range.row < 0 ||
      range.col < 0 ||
      range.row >= sheetInfo.rowCount ||
      range.col >= sheetInfo.colCount
    ) {
      return true;
    }
    const scroll = controller.getScroll(sheetId);
    const old = controller.computeCellPosition({
      row: scroll.row,
      col: scroll.col,
      colCount: 1,
      rowCount: 1,
      sheetId: sheetId,
    });
    const size = controller.getCanvasSize();
    const headerSize = controller.getHeaderSize();
    const { top, left } = controller.computeCellPosition(range);
    const minTop = old.top;
    const minLeft = old.left;
    const maxTop = old.top + size.height - headerSize.height;
    const maxLeft = old.left + size.width - headerSize.width;
    if (top >= minTop && top < maxTop && left >= minLeft && left <= maxLeft) {
      controller.setActiveRange(range);
      return true;
    }
    const oldPosition = controller.computeCellPosition(
      controller.getActiveRange().range,
    );

    scrollBar(controller, left - oldPosition.left, top - oldPosition.top);
    controller.setActiveRange(range);
    return true;
  });
}

export function computeScrollPosition(controller: IController) {
  const contentSize = 30;
  const BOTTOM_BUFF = 200;
  const sheetViewSize = controller.getSheetViewSize();
  const canvasRect = controller.getCanvasSize();
  const maxHeight = sheetViewSize.height - canvasRect.height + BOTTOM_BUFF;
  const maxWidth = sheetViewSize.width - canvasRect.width + BOTTOM_BUFF;
  const maxScrollHeight = canvasRect.height - contentSize;
  const maxScrollWidth = canvasRect.width - contentSize;

  return {
    maxHeight,
    maxWidth,
    maxScrollHeight,
    maxScrollWidth,
  };
}

export function scrollBar(
  controller: IController,
  scrollX: number,
  scrollY: number,
) {
  const oldScroll = controller.getScroll();

  const size = computeScrollPosition(controller);

  const { maxHeight, maxWidth, maxScrollHeight, maxScrollWidth } = size;
  let top = oldScroll.top + Math.ceil(scrollY);

  let left = oldScroll.left + Math.ceil(scrollX);

  if (top < 0) {
    top = 0;
  } else if (top > maxHeight) {
    top = maxHeight;
  }

  if (left < 0) {
    left = 0;
  } else if (left > maxWidth) {
    left = maxWidth;
  }
  const scrollTop = Math.floor((top * maxScrollHeight) / maxHeight);
  const scrollLeft = Math.floor((left * maxScrollWidth) / maxWidth);

  const { row, col } = computeScrollRowAndCol(controller, left, top);

  const newValue: ScrollValue = {
    left,
    top,
    scrollLeft,
    scrollTop,
    row,
    col,
  };
  controller.setScroll(newValue);
}

export function recalculateScroll(controller: IController) {
  const activeCell = controller.getActiveRange().range;
  const temp = {
    row: activeCell.row,
    col: activeCell.col,
    colCount: 1,
    rowCount: 1,
    sheetId: '',
  };
  const position = controller.computeCellPosition(temp);
  const cellSize = controller.getCellSize(temp);
  const domRect = controller.getCanvasSize();
  const oldScroll = controller.getScroll();
  const sheetInfo = controller.getSheetInfo(controller.getCurrentSheetId());
  if (!sheetInfo) {
    return;
  }
  const headerSize = controller.getHeaderSize();
  const buff = 5;
  const size = computeScrollPosition(controller);
  const { maxHeight, maxWidth, maxScrollHeight, maxScrollWidth } = size;
  if (position.left + cellSize.width + buff > domRect.width) {
    if (oldScroll.col <= sheetInfo.colCount - 2) {
      const left = oldScroll.left + controller.getCol(oldScroll.col).len;
      const scrollLeft = Math.floor((left * maxScrollWidth) / maxWidth);
      controller.setScroll({
        ...oldScroll,
        col: oldScroll.col + 1,
        left,
        scrollLeft,
      });
    }
  }

  if (position.left - headerSize.width < domRect.left + buff) {
    if (oldScroll.col >= 1) {
      const left = oldScroll.left - controller.getCol(oldScroll.col).len;
      const scrollLeft = Math.floor((left * maxScrollWidth) / maxWidth);
      controller.setScroll({
        ...oldScroll,
        col: oldScroll.col - 1,
        left,
        scrollLeft,
      });
    }
  }
  if (position.top + cellSize.height + buff > domRect.height) {
    if (oldScroll.row <= sheetInfo.rowCount - 2) {
      const top = oldScroll.top + controller.getRowHeight(oldScroll.row);
      const scrollTop = Math.floor((top * maxScrollHeight) / maxHeight);
      controller.setScroll({
        ...oldScroll,
        row: oldScroll.row + 1,
        top,
        scrollTop,
      });
    }
  }

  if (position.top - headerSize.height < domRect.top + buff) {
    if (oldScroll.row >= 1) {
      const top = oldScroll.top - controller.getRowHeight(oldScroll.row);
      const scrollTop = Math.floor((top * maxScrollHeight) / maxHeight);
      controller.setScroll({
        ...oldScroll,
        row: oldScroll.row - 1,
        top,
        scrollTop,
      });
    }
  }
}

export function checkFocus() {
  const dom = document.activeElement;
  if (!dom || dom.getAttribute('data-role') !== FORMULA_EDITOR_ROLE) {
    return false;
  }
  return true;
}

export function setActiveCellValue(controller: IController) {
  const inputDom = document.activeElement as
    | HTMLInputElement
    | HTMLTextAreaElement;
  const { range, isMerged } = controller.getActiveRange();
  const cellData = controller.getCell(range);
  let value = inputDom.value;
  if (
    typeof cellData?.value === 'string' &&
    isMergeContent(isMerged, cellData?.value)
  ) {
    value = value.replaceAll(LINE_BREAK, MERGE_CELL_LINE_BREAK);
  }
  controller.setCellValue(value, range);
  inputDom.value = '';
  inputDom.blur();
  useCoreStore.getState().setEditorStatus(EditorStatus.NONE);
}

function checkActiveElement(controller: IController) {
  if (!checkFocus()) {
    return;
  }
  setActiveCellValue(controller);
}

type JumpDirection = 'up' | 'down' | 'left' | 'right';

/**
 * Excel-style Ctrl+Arrow jump-to-edge: find the next "boundary" in the
 * given direction starting from the active cell. If the active cell is
 * populated, walk while neighbours are also populated and stop at the
 * last populated cell. If the active cell is empty, walk while
 * neighbours are also empty and stop at the first populated cell. In
 * either case, clamp to sheet bounds.
 */
function jumpToEdge(
  controller: IController,
  direction: JumpDirection,
): { row: number; col: number } | null {
  const sheetId = controller.getCurrentSheetId();
  const sheetInfo = controller.getSheetInfo(sheetId);
  if (!sheetInfo) return null;
  const active = controller.getActiveRange().range;
  let row = active.row;
  let col = active.col;

  const dRow = direction === 'down' ? 1 : direction === 'up' ? -1 : 0;
  const dCol = direction === 'right' ? 1 : direction === 'left' ? -1 : 0;
  const maxRow = sheetInfo.rowCount - 1;
  const maxCol = sheetInfo.colCount - 1;

  const inBounds = (r: number, c: number) =>
    r >= 0 && r <= maxRow && c >= 0 && c <= maxCol;
  const hasData = (r: number, c: number) => {
    const cell = controller.getCell({
      row: r,
      col: c,
      rowCount: 1,
      colCount: 1,
      sheetId,
    });
    const v = cell?.value;
    return v !== undefined && v !== null && v !== '';
  };

  const startHasData = hasData(row, col);
  let nr = row + dRow;
  let nc = col + dCol;

  if (startHasData) {
    // Walk while adjacent cell still has data; stop at last populated.
    while (inBounds(nr, nc) && hasData(nr, nc)) {
      row = nr;
      col = nc;
      nr += dRow;
      nc += dCol;
    }
    // If the very next neighbour was empty AND we didn't move, jump past
    // the empty gap to the next populated cell (Excel behaviour).
    if (row === active.row && col === active.col && inBounds(nr, nc)) {
      while (inBounds(nr, nc) && !hasData(nr, nc)) {
        nr += dRow;
        nc += dCol;
      }
      if (inBounds(nr, nc)) {
        row = nr;
        col = nc;
      } else {
        // Ran off the sheet while scanning the gap; clamp to the edge.
        row = Math.max(0, Math.min(maxRow, nr - dRow));
        col = Math.max(0, Math.min(maxCol, nc - dCol));
      }
    }
  } else {
    // Active is empty: walk to first populated cell in direction.
    while (inBounds(nr, nc) && !hasData(nr, nc)) {
      nr += dRow;
      nc += dCol;
    }
    if (inBounds(nr, nc)) {
      row = nr;
      col = nc;
    } else {
      // No populated cell; jump to sheet edge.
      row = Math.max(0, Math.min(maxRow, nr - dRow));
      col = Math.max(0, Math.min(maxCol, nc - dCol));
    }
  }

  return { row, col };
}

function extendActiveRangeTo(
  controller: IController,
  targetRow: number,
  targetCol: number,
) {
  const active = controller.getActiveRange().range;
  // Anchor is the existing top-left corner of the current selection.
  const anchorRow = active.row;
  const anchorCol = active.col;
  const minRow = Math.min(anchorRow, targetRow);
  const maxRow = Math.max(anchorRow, targetRow);
  const minCol = Math.min(anchorCol, targetCol);
  const maxCol = Math.max(anchorCol, targetCol);
  controller.setActiveRange({
    row: minRow,
    col: minCol,
    rowCount: maxRow - minRow + 1,
    colCount: maxCol - minCol + 1,
    sheetId: active.sheetId || controller.getCurrentSheetId(),
  });
}

const modifierKey: KeyboardEventItem['modifierKey'] = [
  isMac() ? 'meta' : 'ctrl',
];
const modifierKeyShift: KeyboardEventItem['modifierKey'] = [
  isMac() ? 'meta' : 'ctrl',
  'shift',
];

function makeJumpHandler(direction: JumpDirection) {
  return (controller: IController) => {
    if (checkFocus()) return;
    controller.transaction(() => {
      checkActiveElement(controller);
      const target = jumpToEdge(controller, direction);
      if (!target) return true;
      controller.setActiveRange({
        row: target.row,
        col: target.col,
        rowCount: 1,
        colCount: 1,
        sheetId: controller.getCurrentSheetId(),
      });
      recalculateScroll(controller);
      return true;
    });
  };
}

function makeExtendHandler(direction: JumpDirection) {
  return (controller: IController) => {
    if (checkFocus()) return;
    controller.transaction(() => {
      checkActiveElement(controller);
      const target = jumpToEdge(controller, direction);
      if (!target) return true;
      extendActiveRangeTo(controller, target.row, target.col);
      recalculateScroll(controller);
      return true;
    });
  };
}

/* jscpd:ignore-start */
export const keyboardEventList: KeyboardEventItem[] = [
  {
    key: 'Enter',
    modifierKey: [],
    handler: handleEnterClick,
  },
  {
    key: 'Tab',
    modifierKey: [],
    handler: handleTabClick,
  },
  {
    key: 'a',
    modifierKey,
    handler: (controller) => {
      if (checkFocus()) return;
      controller.setActiveRange({
        row: 0,
        col: 0,
        rowCount: 0,
        colCount: 0,
        sheetId: controller.getCurrentSheetId(),
      });
    },
  },
  {
    key: 'ArrowDown',
    modifierKey,
    handler: makeJumpHandler('down'),
  },
  {
    key: 'ArrowUp',
    modifierKey,
    handler: makeJumpHandler('up'),
  },
  {
    key: 'ArrowRight',
    modifierKey,
    handler: makeJumpHandler('right'),
  },
  {
    key: 'ArrowLeft',
    modifierKey,
    handler: makeJumpHandler('left'),
  },
  {
    key: 'ArrowDown',
    modifierKey: modifierKeyShift,
    handler: makeExtendHandler('down'),
  },
  {
    key: 'ArrowUp',
    modifierKey: modifierKeyShift,
    handler: makeExtendHandler('up'),
  },
  {
    key: 'ArrowRight',
    modifierKey: modifierKeyShift,
    handler: makeExtendHandler('right'),
  },
  {
    key: 'ArrowLeft',
    modifierKey: modifierKeyShift,
    handler: makeExtendHandler('left'),
  },
  {
    key: 'ArrowDown',
    modifierKey: [],
    handler: (controller) => {
      if (checkFocus()) {
        return;
      }
      handleEnterClick(controller);
    },
  },
  {
    key: 'ArrowUp',
    modifierKey: [],
    handler: (controller) => {
      if (checkFocus()) {
        return;
      }
      controller.transaction(() => {
        checkActiveElement(controller);
        controller.setNextActiveCell('up');
        recalculateScroll(controller);
        return true;
      });
    },
  },
  {
    key: 'ArrowRight',
    modifierKey: [],
    handler: (controller) => {
      if (checkFocus()) {
        return;
      }
      handleTabClick(controller);
    },
  },
  {
    key: 'ArrowLeft',
    modifierKey: [],
    handler: (controller) => {
      if (checkFocus()) {
        return;
      }
      controller.transaction(() => {
        checkActiveElement(controller);
        controller.setNextActiveCell('left');
        recalculateScroll(controller);
        return true;
      });
    },
  },
  {
    key: 'b',
    modifierKey,
    handler: (controller) => {
      if (checkFocus()) {
        return;
      }
      const cellData = controller.getCell(controller.getActiveRange().range);
      controller.updateCellStyle(
        { isBold: !cellData?.isBold },
        controller.getActiveRange().range,
      );
    },
  },
  {
    key: 'i',
    modifierKey,
    handler: (controller) => {
      if (checkFocus()) {
        return;
      }

      const cellData = controller.getCell(controller.getActiveRange().range);
      controller.updateCellStyle(
        { isItalic: !cellData?.isItalic },
        controller.getActiveRange().range,
      );
    },
  },
  {
    key: '5',
    modifierKey,
    handler: (controller) => {
      if (checkFocus()) {
        return;
      }
      const cellData = controller.getCell(controller.getActiveRange().range);
      controller.updateCellStyle(
        { isStrike: !cellData?.isStrike },
        controller.getActiveRange().range,
      );
    },
  },
  {
    key: 'u',
    modifierKey,
    handler: (controller) => {
      if (checkFocus()) {
        return;
      }
      const cellData = controller.getCell(controller.getActiveRange().range);
      const underline = cellData?.underline;
      let newUnderline = EUnderLine.NONE;
      if (underline === undefined || underline === EUnderLine.NONE) {
        newUnderline = EUnderLine.SINGLE;
      } else {
        newUnderline = EUnderLine.NONE;
      }
      controller.updateCellStyle(
        { underline: newUnderline },
        controller.getActiveRange().range,
      );
    },
  },
  {
    key: 'z',
    modifierKey,
    handler: (controller) => {
      controller.undo();
    },
  },
  {
    key: 'y',
    modifierKey,
    handler: (controller) => {
      controller.redo();
    },
  },
  {
    key: 'Backspace',
    modifierKey: [],
    handler: (controller) => {
      controller.deleteCell(controller.getActiveRange().range);
    },
  },
  {
    key: 'Delete',
    modifierKey: [],
    handler: (controller) => {
      controller.deleteCell(controller.getActiveRange().range);
    },
  },
];
/* jscpd:ignore-end */
