import { describe, expect, it } from "vitest";
import { stubExcel } from "../../testing/excel-mock";
import { columnName, readSelectedCells } from "./selection";

describe("readSelectedCells", () => {
  it("空でないセルを行優先で、シート名つきの A1 形式のアドレスと一緒に返す", async () => {
    const mock = stubExcel({
      sheetName: "Sheet1",
      rowIndex: 2, // 3 行目
      columnIndex: 1, // B 列
      text: [
        ["りんご", ""],
        ["", "みかん"],
        ["ぶどう", "もも"],
      ],
    });

    await expect(readSelectedCells()).resolves.toEqual([
      { address: "Sheet1!B3", text: "りんご" },
      { address: "Sheet1!C4", text: "みかん" },
      { address: "Sheet1!B5", text: "ぶどう" },
      { address: "Sheet1!C5", text: "もも" },
    ]);
    expect(mock.runCount).toBe(1);
    expect(mock.syncCount).toBe(1);
  });

  it("すべて空なら空の配列を返す", async () => {
    stubExcel({ sheetName: "Sheet1", rowIndex: 0, columnIndex: 0, text: [[""]] });

    await expect(readSelectedCells()).resolves.toEqual([]);
  });
});

describe("columnName", () => {
  it.each([
    [0, "A"],
    [25, "Z"],
    [26, "AA"],
    [51, "AZ"],
    [52, "BA"],
    [701, "ZZ"],
    [702, "AAA"],
    [16383, "XFD"], // Excel の最終列
  ])("%i → %s", (index, name) => {
    expect(columnName(index)).toBe(name);
  });
});
