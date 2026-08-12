import type { CommandResult, SheetModel, WorkbookDataInput } from "../domain/types";
import { cloneValue } from "../utils/clone";
import { createId } from "../utils/id";
import type { CommandContext, SpreadsheetCommand } from "./command-types";

const createSheetModel = (
  input: WorkbookDataInput | undefined,
  workbook: CommandContext["workbook"]
): SheetModel => ({
  id: input?.id ?? createId(`sheet-${workbook.sheets.length + 1}`),
  name: input?.name ?? `Sheet${workbook.sheets.length + 1}`,
  cells: cloneValue(input?.cells ?? {}),
  merges: cloneValue(input?.merges ?? []),
  conditionalFormats: cloneValue(input?.conditionalFormats ?? []),
  frozenRows: Math.max(0, input?.frozenRows ?? 0),
  frozenColumns: Math.max(0, input?.frozenColumns ?? 0),
  columns: cloneValue(input?.columns ?? {}),
  rows: cloneValue(input?.rows ?? {}),
  rowCount: Math.min(input?.rowCount ?? 200, workbook.settings.maxRows),
  columnCount: Math.min(input?.columnCount ?? 26, workbook.settings.maxColumns),
  selection: {
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 }
  },
  charts: cloneValue(input?.charts ?? []),
  metadata: cloneValue(input?.metadata ?? {})
});

export class AddSheetCommand implements SpreadsheetCommand {
  readonly type = "AddSheetCommand";

  constructor(private readonly input?: WorkbookDataInput) {}

  execute(context: CommandContext): CommandResult {
    const workbook = cloneValue(context.workbook);
    const sheet = createSheetModel(this.input, workbook);
    workbook.sheets.push(sheet);
    workbook.activeSheetId = sheet.id;

    return {
      workbook,
      operations: [
        {
          op: "addSheet",
          id: sheet.id,
          path: ["sheets", workbook.sheets.length - 1],
          value: sheet
        }
      ],
      affectedRanges: [sheet.selection],
      recordHistory: true
    };
  }
}