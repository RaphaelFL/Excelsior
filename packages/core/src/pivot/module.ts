import type {
  PivotBuildAsyncOptions,
  PivotInferenceInput,
  PivotModule,
  PivotSheetInput,
  WorkbookDataInput,
  WorkbookModel
} from "../domain/types";
import { buildPivotSheet, buildPivotSheetAsync, inferPivotSheetInput } from "./engine";

export const defaultPivotModule: PivotModule = {
  createPivotSheet(workbook: Readonly<WorkbookModel>, input: PivotSheetInput): WorkbookDataInput {
    return buildPivotSheet(workbook, input);
  },
  createPivotSheetAsync(
    workbook: Readonly<WorkbookModel>,
    input: PivotSheetInput,
    options?: PivotBuildAsyncOptions
  ): Promise<WorkbookDataInput> {
    return buildPivotSheetAsync(workbook, input, options);
  },
  inferPivotSheet(workbook: Readonly<WorkbookModel>, input: PivotInferenceInput): PivotSheetInput {
    return inferPivotSheetInput(workbook, input);
  }
};