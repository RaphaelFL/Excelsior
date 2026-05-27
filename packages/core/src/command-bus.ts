import { cloneValue } from "./utils/clone";
import type { FormulaEngine, SpreadsheetEventMap, SpreadsheetOperation, WorkbookModel } from "./domain/types";
import type { SpreadsheetCommand } from "./commands/command-types";
import { TypedEventEmitter } from "./events/typed-event-emitter";
import { HistoryManager } from "./history/history-manager";
import type { ValidationRegistry } from "./validation/validation-registry";

export class CommandBus {
  private revision = 0;

  constructor(
    private workbook: WorkbookModel,
    private readonly historyManager: HistoryManager,
    private readonly events: TypedEventEmitter<SpreadsheetEventMap>,
    private readonly formulaEngine?: FormulaEngine,
    private readonly validationRegistry?: ValidationRegistry
  ) {}

  getSnapshot(): WorkbookModel {
    return cloneValue(this.workbook);
  }

  peekWorkbook(): Readonly<WorkbookModel> {
    return this.workbook;
  }

  getRevision(): number {
    return this.revision;
  }

  replaceWorkbook(workbook: WorkbookModel): void {
    this.workbook = cloneValue(workbook);
    this.revision += 1;
  }

  execute(command: SpreadsheetCommand): { workbook: WorkbookModel; operations: SpreadsheetOperation[] } {
    const startedAt = performance.now();
    const previous = this.getSnapshot();

    try {
      const result = command.execute({
        workbook: previous,
        formulaEngine: this.formulaEngine,
        validationRegistry: this.validationRegistry!,
        now: () => Date.now()
      });

      this.workbook = cloneValue(result.workbook);
  this.revision += 1;

      if (result.recordHistory) {
        this.historyManager.push({
          before: previous,
          after: cloneValue(result.workbook),
          operations: result.operations,
          affectedRanges: result.affectedRanges
        });
      }

      this.events.emit("command:completed", {
        timestamp: Date.now(),
        workbookId: this.workbook.id,
        sheetId: command.sheetId,
        durationMs: performance.now() - startedAt,
        commandType: command.type,
        operations: result.operations
      });

      return {
        workbook: this.getSnapshot(),
        operations: result.operations
      };
    } catch (error) {
      const resolvedError = error instanceof Error ? error : new Error("Command execution failed.");

      this.events.emit("command:failed", {
        timestamp: Date.now(),
        workbookId: this.workbook.id,
        sheetId: command.sheetId,
        durationMs: performance.now() - startedAt,
        commandType: command.type,
        errorCode: resolvedError.name || "COMMAND_FAILED"
      });

      throw error;
    }
  }

  undo(): WorkbookModel | undefined {
    const entry = this.historyManager.undo();
    if (!entry) {
      return undefined;
    }
    this.workbook = cloneValue(entry.before);
    this.revision += 1;
    return this.getSnapshot();
  }

  redo(): WorkbookModel | undefined {
    const entry = this.historyManager.redo();
    if (!entry) {
      return undefined;
    }
    this.workbook = cloneValue(entry.after);
    this.revision += 1;
    return this.getSnapshot();
  }
}