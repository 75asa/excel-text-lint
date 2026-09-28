import { describe, expect, it } from "vitest";
import type { ExcelLocation } from "../../core/locations";
import type { Fix, TextUnit } from "../../core/types";
import { type FakeCell, stubExcelWorkbook } from "../../testing/excel-mock";
import { ExcelAdapter } from "./adapter";
import { applyFixesToCell, applyFixesToText, wouldBeReinterpreted } from "./fix";

const text = "することができる。";
const fix: Fix = { range: [0, 8], text: "できる" };

function unitOf(cellText: string): TextUnit<ExcelLocation> {
  return {
    id: "Sheet1!B3",
    text: cellText,
    location: { host: "excel", sheet: "Sheet1", address: "B3", row: 2, col: 1 },
  };
}

function stubCell(cell: FakeCell) {
  return stubExcelWorkbook([{ name: "Sheet1", cells: { B3: cell } }]);
}

describe("applyFixesToText", () => {
  it("後ろの位置から順に適用するので、前の修正案の位置がずれない", () => {
    const result = applyFixesToText("abcdef", [
      { range: [0, 1], text: "AA" },
      { range: [4, 6], text: "" },
      { range: [2, 3], text: "CCC" },
    ]);
    expect(result.text).toBe("AAbCCCd");
    expect(result.applied.map((f) => f.range)).toEqual([
      [4, 6],
      [2, 3],
      [0, 1],
    ]);
    expect(result.skipped).toEqual([]);
  });

  it("範囲が重なる修正案と、テキストの外の修正案は飛ばす", () => {
    const overlapping: Fix = { range: [1, 3], text: "x" };
    const outside: Fix = { range: [5, 9], text: "y" };
    const result = applyFixesToText("abcd", [{ range: [2, 4], text: "Z" }, overlapping, outside]);
    expect(result.text).toBe("abZ");
    expect(result.skipped).toEqual([outside, overlapping]);
  });
});

describe("ExcelAdapter.applyFix", () => {
  it("fix.range を fix.text に置き換えてセルに書き戻す", async () => {
    const mock = stubCell({ value: `これを${text}` });

    await expect(
      new ExcelAdapter().applyFix(unitOf(`これを${text}`), { ...fix, range: [3, 11] }),
    ).resolves.toEqual({ status: "applied", text: "これをできる。" });
    expect(mock.writes).toEqual([{ address: "Sheet1!B3", values: [["これをできる。"]] }]);
  });

  it("セルのテキストが collect したときと違えば stale を返し、書き込まない", async () => {
    const mock = stubCell({ value: "書き換えられた" });

    await expect(new ExcelAdapter().applyFix(unitOf(text), fix)).resolves.toEqual({
      status: "stale",
      currentText: "書き換えられた",
    });
    expect(mock.writes).toEqual([]);
  });

  it("シートが削除されていたら stale を返す", async () => {
    stubExcelWorkbook([{ name: "Sheet2" }]);

    await expect(new ExcelAdapter().applyFix(unitOf(text), fix)).resolves.toMatchObject({
      status: "stale",
    });
  });

  it("数式のセルは unsupported にする", async () => {
    const mock = stubCell({ value: text, formula: '="すること"&"ができる。"' });

    await expect(new ExcelAdapter().applyFix(unitOf(text), fix)).resolves.toMatchObject({
      status: "unsupported",
    });
    expect(mock.writes).toEqual([]);
  });

  it("数値のセルや、表示形式で見た目が値と違うセルは unsupported にする", async () => {
    stubCell({ value: 1234, text: "1,234" });
    await expect(
      new ExcelAdapter().applyFix(unitOf("1,234"), { range: [1, 2], text: "" }),
    ).resolves.toMatchObject({ status: "unsupported" });

    stubCell({ value: "本文", text: "本文（注）" });
    await expect(
      new ExcelAdapter().applyFix(unitOf("本文（注）"), { range: [0, 1], text: "文" }),
    ).resolves.toMatchObject({ status: "unsupported" });
  });

  it("修正後の文字列が数式や数値として解釈されるなら unsupported にする", async () => {
    const mock = stubCell({ value: "a=1" });

    await expect(
      new ExcelAdapter().applyFix(unitOf("a=1"), { range: [0, 1], text: "" }),
    ).resolves.toMatchObject({ status: "unsupported" });
    expect(mock.writes).toEqual([]);
  });

  it("Excel の API がエラーを返したら failed を返す", async () => {
    const mock = stubCell({ value: text });
    const error = new Error("GeneralException");
    mock.failNextSync = error;

    await expect(new ExcelAdapter().applyFix(unitOf(text), fix)).resolves.toEqual({
      status: "failed",
      error,
    });
  });
});

describe("applyFixesToCell", () => {
  it("1 つのセルの複数の修正案を、後ろから順に適用して 1 回で書き戻す", async () => {
    const cell = "１２３を行なう。することができる。";
    const mock = stubCell({ value: cell });

    const result = await applyFixesToCell(unitOf(cell), [
      { range: [0, 3], text: "123" },
      { range: [4, 7], text: "行う" },
      { range: [8, 17], text: "できる。" },
    ]);

    expect(result).toMatchObject({ status: "applied", text: "123を行う。できる。", skipped: [] });
    expect(mock.writes).toEqual([{ address: "Sheet1!B3", values: [["123を行う。できる。"]] }]);
  });
});

describe("wouldBeReinterpreted", () => {
  it.each([
    ["=SUM(A1)", true],
    ["+81", true],
    ["-", true],
    ["'文字列", true],
    ["@name", true],
    ["123", true],
    ["1,234.5", true],
    ["50%", true],
    ["1e3", true],
    ["123円", false],
    ["することができる。", false],
    ["", false],
  ])("%s → %s", (value, expected) => {
    expect(wouldBeReinterpreted(value)).toBe(expected);
  });
});
