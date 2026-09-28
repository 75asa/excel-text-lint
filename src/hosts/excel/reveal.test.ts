import { describe, expect, it } from "vitest";
import type { ExcelLocation } from "../../core/locations";
import type { TextUnit } from "../../core/types";
import { stubExcelWorkbook } from "../../testing/excel-mock";
import { ExcelAdapter } from "./adapter";

const unit: TextUnit<ExcelLocation> = {
  id: "Sheet2!B3",
  text: "することができる。",
  location: { host: "excel", sheet: "Sheet2", address: "B3", row: 2, col: 1 },
};

describe("ExcelAdapter.reveal", () => {
  it("シートをアクティブにしてセルを選択し、unit を返す", async () => {
    const mock = stubExcelWorkbook([
      { name: "Sheet1" },
      { name: "Sheet2", cells: { B3: { value: "することができる。" } } },
    ]);

    await expect(new ExcelAdapter().reveal(unit, [6, 7])).resolves.toBe("unit");
    expect(mock.activeSheet).toBe("Sheet2");
    expect(mock.selected).toBe("Sheet2!B3");
  });

  it("シートが削除されていたら none を返し、何も選択しない", async () => {
    const mock = stubExcelWorkbook([{ name: "Sheet1" }]);

    await expect(new ExcelAdapter().reveal(unit)).resolves.toBe("none");
    expect(mock.activeSheet).toBeNull();
    expect(mock.selected).toBeNull();
  });

  it("セルのテキストが変わっていたら（行の挿入などでずれたら）none を返し、別のセルを選ばない", async () => {
    const mock = stubExcelWorkbook([{ name: "Sheet2", cells: { B3: { value: "別の文字列" } } }]);

    await expect(new ExcelAdapter().reveal(unit)).resolves.toBe("none");
    expect(mock.selected).toBeNull();
  });

  it("Excel の API がエラーを返したら none を返す", async () => {
    const mock = stubExcelWorkbook([{ name: "Sheet2", cells: { B3: { value: unit.text } } }]);
    mock.failNextSync = new Error("InvalidOperation");

    await expect(new ExcelAdapter().reveal(unit)).resolves.toBe("none");
  });
});
