import { describe, expect, it } from "vitest";
import { createFigureFromSpreadsheetRange } from "../src";

describe("spreadsheet adapter", () => {
  it("creates cartesian traces from spreadsheet range input", () => {
    const figure = createFigureFromSpreadsheetRange({
      sheetName: "Resumo",
      headers: ["Mes", "Receita", "Custo"],
      rows: [
        ["Jan", 120, 80],
        ["Fev", 160, 95],
        ["Mar", 170, 110]
      ]
    });

    expect(figure.data).toHaveLength(2);
    expect(figure.data[0].name).toBe("Receita");
    expect(figure.data[1].name).toBe("Custo");
    expect(figure.layout?.xAxis?.title).toBe("Mes");
  });

  it("supports explicit x and series column selection", () => {
    const figure = createFigureFromSpreadsheetRange(
      {
        headers: ["Dia", "Pedidos", "Taxa", "Observacao"],
        rows: [
          ["01", 10, 0.2, "ok"],
          ["02", 12, 0.18, "ok"],
          ["03", 9, 0.24, "atencao"]
        ]
      },
      {
        xColumn: "Dia",
        seriesColumns: ["Pedidos"],
        traceType: "bar"
      }
    );

    expect(figure.data).toHaveLength(1);
    expect(figure.data[0].type).toBe("bar");
    expect(figure.data[0].name).toBe("Pedidos");
  });
});
