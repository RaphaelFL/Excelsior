import type { CellRange, SpreadsheetOperation, WorkbookModel } from "../domain/types";

interface HistoryEntry {
  before: WorkbookModel;
  after: WorkbookModel;
  operations: SpreadsheetOperation[];
  affectedRanges: CellRange[];
}

export class HistoryManager {
  private past: HistoryEntry[] = [];

  private future: HistoryEntry[] = [];

  constructor(private readonly maxHistorySize: number) {}

  push(entry: HistoryEntry): void {
    this.past.push(entry);
    if (this.past.length > this.maxHistorySize) {
      this.past.shift();
    }
    this.future = [];
  }

  undo(): HistoryEntry | undefined {
    const entry = this.past.pop();
    if (!entry) {
      return undefined;
    }
    this.future.push(entry);
    return entry;
  }

  redo(): HistoryEntry | undefined {
    const entry = this.future.pop();
    if (!entry) {
      return undefined;
    }
    this.past.push(entry);
    return entry;
  }

  hasPast(): boolean {
    return this.past.length > 0;
  }

  hasFuture(): boolean {
    return this.future.length > 0;
  }
}