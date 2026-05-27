import { describe, expect, it } from "vitest";
import { WorkbookEngine } from "../src/index";
import type { GridPlugin } from "../src/index";

describe("plugin engine", () => {
  it("registers plugins by unique id and activates them per workbook instance", () => {
    const engine = new WorkbookEngine();
    const events: string[] = [];
    const plugin: GridPlugin = {
      id: "audit-plugin",
      setup(context) {
        events.push(context.pluginId);
        context.setState({ activations: 1 });
      }
    };

    engine.registerPlugin(plugin);

    expect(engine.isPluginEnabled("audit-plugin")).toBe(true);
    expect(engine.getRegisteredPlugins()).toContainEqual({ id: "audit-plugin", enabled: true });
    expect(events).toEqual(["audit-plugin"]);
    expect(engine.getPluginState<{ activations: number }>("audit-plugin")).toEqual({ activations: 1 });
    expect(() => engine.registerPlugin(plugin)).toThrow(/already registered/i);
  });

  it("removes plugin listeners when disabled and unregisters cleanly", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();
    let updates = 0;

    engine.registerPlugin({
      id: "listener-plugin",
      setup(context) {
        return context.on("cell:updated", () => {
          updates += 1;
        });
      }
    });

    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "A" });
    expect(updates).toBe(1);

    engine.disablePlugin("listener-plugin");
    engine.setCellValue({ sheetId: sheet.id, row: 0, col: 1, value: "B" });
    expect(updates).toBe(1);

    engine.unregisterPlugin("listener-plugin");
    expect(engine.getRegisteredPlugins()).toEqual([]);
  });

  it("isolates plugin state between workbooks", () => {
    const first = new WorkbookEngine();
    const second = new WorkbookEngine();

    const plugin: GridPlugin = {
      id: "selection-counter",
      setup(context) {
        context.setState({ selections: 0 });
        context.on("selection:changed", () => {
          context.setState<{ selections: number }>((previous) => ({
            selections: (previous?.selections ?? 0) + 1
          }));
        });
      }
    };

    first.registerPlugin(plugin);
    second.registerPlugin(plugin, false);
    const firstSheet = first.getActiveSheet();
    const secondSheet = second.getActiveSheet();

    first.selectRange({ sheetId: firstSheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 2 });
    first.selectRange({ sheetId: firstSheet.id, rowStart: 1, rowEnd: 1, colStart: 0, colEnd: 1 });
    second.selectRange({ sheetId: secondSheet.id, rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });

    expect(first.getPluginState<{ selections: number }>("selection-counter")).toEqual({ selections: 2 });
    expect(second.getPluginState<{ selections: number }>("selection-counter")).toBeUndefined();

    second.enablePlugin("selection-counter");
    second.selectRange({ sheetId: secondSheet.id, rowStart: 2, rowEnd: 2, colStart: 0, colEnd: 0 });
    expect(second.getPluginState<{ selections: number }>("selection-counter")).toEqual({ selections: 1 });
  });

  it("does not break the main flow when a listener throws", () => {
    const engine = new WorkbookEngine();
    const sheet = engine.getActiveSheet();

    engine.registerPlugin({
      id: "unstable-plugin",
      setup(context) {
        context.on("cell:updated", () => {
          throw new Error("listener failure");
        });
      }
    });

    expect(() => {
      engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: "safe" });
    }).not.toThrow();
    expect(engine.getDisplayValue(sheet.id, 0, 0)).toBe("safe");
  });
});