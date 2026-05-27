import { describe, expect, it } from "vitest";
import { getVisibleCellWindow } from "@excelsior/renderer-dom";

describe("visible window calculation", () => {
  it("clamps viewport coordinates to sheet bounds", () => {
    expect(
      getVisibleCellWindow({
        scrollTop: 999999,
        scrollLeft: 999999,
        viewportHeight: 600,
        viewportWidth: 800,
        rowCount: 20,
        columnCount: 10,
        rowHeight: 30,
        columnWidth: 100,
        buffer: 4
      })
    ).toEqual({
      rowStart: 19,
      rowEnd: 19,
      colStart: 9,
      colEnd: 9
    });
  });
});