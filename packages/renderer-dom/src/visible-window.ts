export interface VisibleWindowInput {
  scrollTop: number;
  scrollLeft: number;
  viewportHeight: number;
  viewportWidth: number;
  rowCount: number;
  columnCount: number;
  rowHeight: number;
  columnWidth: number;
  buffer: number;
}

export interface VisibleWindow {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

const clamp = (value: number, min: number, max: number): number =>
  Math.min(Math.max(value, min), max);

export const getVisibleCellWindow = ({
  scrollTop,
  scrollLeft,
  viewportHeight,
  viewportWidth,
  rowCount,
  columnCount,
  rowHeight,
  columnWidth,
  buffer
}: VisibleWindowInput): VisibleWindow => {
  const maxRowIndex = Math.max(0, rowCount - 1);
  const maxColumnIndex = Math.max(0, columnCount - 1);
  const startRow = clamp(Math.floor(scrollTop / rowHeight) - buffer, 0, maxRowIndex);
  const endRow = clamp(
    Math.ceil((scrollTop + viewportHeight) / rowHeight) + buffer,
    startRow,
    maxRowIndex
  );
  const startCol = clamp(Math.floor(scrollLeft / columnWidth) - buffer, 0, maxColumnIndex);
  const endCol = clamp(
    Math.ceil((scrollLeft + viewportWidth) / columnWidth) + buffer,
    startCol,
    maxColumnIndex
  );

  return {
    rowStart: startRow,
    rowEnd: endRow,
    colStart: startCol,
    colEnd: endCol
  };
};