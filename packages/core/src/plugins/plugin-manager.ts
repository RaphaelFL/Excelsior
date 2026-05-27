import { cloneValue } from "../utils/clone";
import { createCoreOperationError } from "../errors/spreadsheet-operation-error";
import type { WorkbookEngine } from "../workbook-engine";
import type { GridPlugin, PluginContext, PluginDisposer, PluginState, RegisteredGridPlugin } from "./types";

interface ActivePluginSession {
  dispose: PluginDisposer;
}

const createPluginApi = (engine: WorkbookEngine) => ({
  setCellValue: engine.setCellValue.bind(engine),
  updateCells: engine.updateCells.bind(engine),
  applyCellTransaction: engine.applyCellTransaction.bind(engine),
  setRemoteGroupExpanded: engine.setRemoteGroupExpanded.bind(engine),
  setCellStyle: engine.setCellStyle.bind(engine),
  setCellValidation: engine.setCellValidation.bind(engine),
  setConditionalFormattingRules: engine.setConditionalFormattingRules.bind(engine),
  freezeRows: engine.freezeRows.bind(engine),
  freezeColumns: engine.freezeColumns.bind(engine),
  setRowsHidden: engine.setRowsHidden.bind(engine),
  setColumnsHidden: engine.setColumnsHidden.bind(engine),
  resizeColumn: engine.resizeColumn.bind(engine),
  resizeRow: engine.resizeRow.bind(engine),
  mergeCells: engine.mergeCells.bind(engine),
  unmergeCells: engine.unmergeCells.bind(engine),
  selectRange: engine.selectRange.bind(engine),
  applyOperations: engine.applyOperations.bind(engine),
  getSelection: engine.getSelection.bind(engine),
  getMerge: engine.getMerge.bind(engine),
  getColumnSchema: engine.getColumnSchema.bind(engine),
  getRowSchema: engine.getRowSchema.bind(engine),
  addSheet: engine.addSheet.bind(engine),
  deleteSheet: engine.deleteSheet.bind(engine),
  insertRows: engine.insertRows.bind(engine),
  deleteRows: engine.deleteRows.bind(engine),
  insertColumns: engine.insertColumns.bind(engine),
  deleteColumns: engine.deleteColumns.bind(engine),
  undo: engine.undo.bind(engine),
  redo: engine.redo.bind(engine),
  reportSecurityEvent: engine.reportSecurityEvent.bind(engine),
  validateCellValue: engine.validateCellValue.bind(engine),
  getCellValidation: engine.getCellValidation.bind(engine),
  getConditionalFormattingRules: engine.getConditionalFormattingRules.bind(engine),
  getConditionalStyle: engine.getConditionalStyle.bind(engine),
  getFrozenPane: engine.getFrozenPane.bind(engine),
  registerValidator: engine.registerValidator.bind(engine),
  unregisterValidator: engine.unregisterValidator.bind(engine)
});

export class PluginManager {
  private readonly plugins = new Map<string, GridPlugin>();

  private readonly activePlugins = new Map<string, ActivePluginSession>();

  private readonly pluginState = new Map<string, PluginState>();

  constructor(private readonly engine: WorkbookEngine) {}

  register(plugin: GridPlugin, enabled = true): void {
    if (this.plugins.has(plugin.id)) {
      throw createCoreOperationError("CORE_PLUGIN_ALREADY_REGISTERED", `Plugin already registered: ${plugin.id}`, {
        pluginId: plugin.id
      });
    }

    this.plugins.set(plugin.id, plugin);

    if (enabled) {
      this.enable(plugin.id);
    }
  }

  unregister(pluginId: string): void {
    if (!this.plugins.has(pluginId)) {
      throw createCoreOperationError("CORE_PLUGIN_NOT_FOUND", `Plugin not found: ${pluginId}`, { pluginId });
    }

    this.disable(pluginId);
    this.plugins.delete(pluginId);
    this.pluginState.delete(pluginId);
  }

  enable(pluginId: string): void {
    if (this.activePlugins.has(pluginId)) {
      return;
    }

    const plugin = this.plugins.get(pluginId);
    if (!plugin) {
      throw createCoreOperationError("CORE_PLUGIN_NOT_FOUND", `Plugin not found: ${pluginId}`, { pluginId });
    }

    const cleanup = new Set<PluginDisposer>();
    const context = this.createContext(pluginId, cleanup);

    try {
      const disposer = plugin.setup(context);
      if (typeof disposer === "function") {
        cleanup.add(disposer);
      }
    } catch (error) {
      for (const dispose of Array.from(cleanup).reverse()) {
        dispose();
      }
      throw error;
    }

    this.activePlugins.set(pluginId, {
      dispose: () => {
        for (const dispose of Array.from(cleanup).reverse()) {
          dispose();
        }
      }
    });
  }

  disable(pluginId: string): void {
    const active = this.activePlugins.get(pluginId);
    if (!active) {
      return;
    }

    this.activePlugins.delete(pluginId);
    active.dispose();
  }

  clear(): void {
    for (const pluginId of Array.from(this.activePlugins.keys())) {
      this.disable(pluginId);
    }
    this.plugins.clear();
    this.pluginState.clear();
  }

  list(): RegisteredGridPlugin[] {
    return Array.from(this.plugins.keys()).map((id) => ({
      id,
      enabled: this.activePlugins.has(id)
    }));
  }

  isEnabled(pluginId: string): boolean {
    return this.activePlugins.has(pluginId);
  }

  getState<TState extends PluginState = PluginState>(pluginId: string): TState | undefined {
    const state = this.pluginState.get(pluginId);
    return state ? (cloneValue(state) as TState) : undefined;
  }

  setState<TState extends PluginState = PluginState>(pluginId: string, nextState: TState): TState {
    const cloned = cloneValue(nextState) as TState;
    this.pluginState.set(pluginId, cloned);
    return cloneValue(cloned) as TState;
  }

  private createContext(pluginId: string, cleanup: Set<PluginDisposer>): PluginContext {
    return {
      pluginId,
      workbookId: this.engine.getSnapshot().id,
      commands: createPluginApi(this.engine),
      getSnapshot: () => cloneValue(this.engine.getSnapshot()),
      getState: () => this.getState(pluginId),
      setState: <TState extends PluginState = PluginState>(
        nextState: TState | ((previousState: TState | undefined) => TState)
      ) => {
        const previousState = this.getState<TState>(pluginId);
        const resolvedState =
          typeof nextState === "function"
            ? (nextState as (previousState: TState | undefined) => TState)(previousState)
            : nextState;
        return this.setState(pluginId, resolvedState);
      },
      on: (eventName, listener) => {
        const unsubscribe = this.engine.on(eventName, listener);
        cleanup.add(unsubscribe);

        return () => {
          cleanup.delete(unsubscribe);
          unsubscribe();
        };
      }
    };
  }
}