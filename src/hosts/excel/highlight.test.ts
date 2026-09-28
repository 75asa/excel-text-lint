import { describe, expect, it } from "vitest";
import type { ExcelLocation } from "../../core/locations";
import type { TextUnit, Violation } from "../../core/types";
import { stubExcelWorkbook } from "../../testing/excel-mock";
import { ExcelAdapter } from "./adapter";
import { HIGHLIGHT_FILL_COLOR, HIGHLIGHT_FORMULA, highlightCells } from "./highlight";

function unit(sheet: string, address: string, text = "テキスト"): TextUnit<ExcelLocation> {
  return {
    id: `${sheet}!${address}`,
    text,
    location: { host: "excel", sheet, address, row: 0, col: 0 },
  };
}

function violation(unitId: string): Violation {
  return {
    unitId,
    ruleId: "rule",
    message: "message",
    severity: "error",
    range: [0, 1],
    displayRange: [0, 1],
    line: 1,
    column: 1,
  };
}

const units = new Map(
  [unit("Sheet1", "A1"), unit("Sheet1", "B2"), unit("Sheet2", "C3"), unit("Gone", "A1")].map(
    (u) => [u.id, u],
  ),
);

/** ユーザーが自分で付けていた条件付き書式。 */
const userFormat = {
  id: "user",
  type: "Custom" as const,
  address: "A1",
  formula: "=A1>100",
  fillColor: "#FF0000",
  priority: 0,
};

describe("ExcelAdapter.highlight / clearHighlight", () => {
  it("違反のあるセルごとに条件付き書式を 1 つ足し、ユーザーの書式には触れずに解除できる", async () => {
    const mock = stubExcelWorkbook([
      { name: "Sheet1", cells: { A1: { value: "テキスト" } } },
      { name: "Sheet2" },
    ]);
    mock.sheet("Sheet1").conditionalFormats.push({ ...userFormat });
    const adapter = new ExcelAdapter();

    await adapter.highlight(
      [
        violation("Sheet1!A1"),
        violation("Sheet1!A1"), // 同じセルに 2 件
        violation("Sheet2!C3"),
        violation("Gone!A1"), // 削除されたシート
        violation("unknown"), // units にない
      ],
      units,
    );

    const sheet1 = mock.sheet("Sheet1").conditionalFormats;
    expect(sheet1).toHaveLength(2);
    expect(sheet1[0]).toMatchObject({
      address: "A1",
      formula: HIGHLIGHT_FORMULA,
      fillColor: HIGHLIGHT_FILL_COLOR,
      priority: 0,
    });
    expect(mock.sheet("Sheet2").conditionalFormats).toMatchObject([{ address: "C3" }]);
    // セルの値は変えない
    expect(mock.sheet("Sheet1").cells.A1).toEqual({ value: "テキスト" });

    await adapter.clearHighlight();

    expect(mock.sheet("Sheet1").conditionalFormats).toEqual([{ ...userFormat, priority: 1 }]);
    expect(mock.sheet("Sheet2").conditionalFormats).toEqual([]);
  });

  it("もう一度 highlight すると、前回のハイライトを消してから付け直す", async () => {
    const mock = stubExcelWorkbook([{ name: "Sheet1" }, { name: "Sheet2" }]);

    await expect(highlightCells([violation("Sheet1!A1")], units)).resolves.toBe(1);
    await expect(highlightCells([violation("Sheet1!B2")], units)).resolves.toBe(1);

    expect(mock.sheet("Sheet1").conditionalFormats).toMatchObject([{ address: "B2" }]);
  });

  it("タスクペインを開き直したあと（別のアダプタ）でも、ブックに残ったハイライトを解除できる", async () => {
    const mock = stubExcelWorkbook([{ name: "Sheet1" }, { name: "Sheet2" }]);
    await new ExcelAdapter().highlight([violation("Sheet1!A1"), violation("Sheet2!C3")], units);
    // ユーザーがハイライトされたセルをコピーした
    const [highlighted] = mock.sheet("Sheet1").conditionalFormats;
    expect(highlighted).toBeDefined();
    if (highlighted) {
      mock.sheet("Sheet2").conditionalFormats.push({ ...highlighted, id: "copied", address: "D4" });
    }

    await new ExcelAdapter().clearHighlight();

    expect(mock.sheet("Sheet1").conditionalFormats).toEqual([]);
    expect(mock.sheet("Sheet2").conditionalFormats).toEqual([]);
  });

  it("違反がなければ、解除だけする", async () => {
    const mock = stubExcelWorkbook([{ name: "Sheet1" }]);
    await highlightCells([violation("Sheet1!A1")], units);

    await expect(highlightCells([], units)).resolves.toBe(0);
    expect(mock.sheet("Sheet1").conditionalFormats).toEqual([]);
  });
});
