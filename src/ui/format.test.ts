import { describe, expect, it } from "vitest";
import {
  containerOf,
  formatExcelAddress,
  formatLocation,
  isFormulaLocation,
  quoteSheetName,
  splitRuleId,
} from "./format";

describe("quoteSheetName", () => {
  it.each([
    ["Sheet1", "Sheet1"],
    ["売上", "売上"],
    ["集計_2024", "集計_2024"],
  ])("英数字・日本語・_ だけの名前 %s はそのまま", (name, expected) => {
    expect(quoteSheetName(name)).toBe(expected);
  });

  it.each([
    ["月次 売上", "'月次 売上'"],
    ["2024", "'2024'"],
    ["A1", "'A1'"],
    ["R1C1", "'R1C1'"],
    ["a-b", "'a-b'"],
    ["Bob's", "'Bob''s'"],
  ])("空白・記号・番地と紛らわしい名前 %s は ' で囲む", (name, expected) => {
    expect(quoteSheetName(name)).toBe(expected);
  });
});

describe("formatLocation", () => {
  it("Excel はシート名 + セル番地", () => {
    const location = { host: "excel", sheet: "Sheet1", address: "B12", row: 11, col: 1 };
    expect(formatLocation(location)).toBe("Sheet1!B12");
    expect(formatExcelAddress("月次 売上", "C3")).toBe("'月次 売上'!C3");
  });

  it("ほかのホストも 1 行で表す", () => {
    expect(
      formatLocation({
        host: "powerpoint",
        slideId: "s",
        slideIndex: 1,
        shapeId: "x",
        shapeName: "タイトル 1",
      }),
    ).toBe("スライド 2 / タイトル 1");
    expect(formatLocation({ host: "word", part: "body", paragraphIndex: 4 })).toBe("本文 段落 5");
    expect(formatLocation({ host: "text", paragraphIndex: 0 })).toBe("段落 1");
  });

  it("知らない location なら fallback を返す", () => {
    expect(formatLocation(undefined, "u1")).toBe("u1");
    expect(formatLocation({ host: "unknown" }, "u2")).toBe("u2");
  });
});

describe("containerOf", () => {
  it("Excel はシートでまとめる", () => {
    expect(containerOf({ host: "excel", sheet: "Sheet2", address: "A1", row: 0, col: 0 })).toEqual({
      key: "excel:Sheet2",
      label: "Sheet2",
    });
  });

  it("location が不明なら 1 つにまとめる", () => {
    expect(containerOf(null)).toEqual({ key: "", label: "（場所不明）" });
  });
});

describe("splitRuleId", () => {
  it("プリセット名とルール名に分ける", () => {
    expect(splitRuleId("ja-technical-writing/ja-no-redundant-expression")).toEqual({
      preset: "ja-technical-writing",
      name: "ja-no-redundant-expression",
    });
    expect(splitRuleId("prh")).toEqual({ preset: "", name: "prh" });
  });
});

describe("isFormulaLocation", () => {
  it("isFormula が true のときだけ true", () => {
    expect(isFormulaLocation({ host: "excel", isFormula: true })).toBe(true);
    expect(isFormulaLocation({ host: "excel", isFormula: false })).toBe(false);
    expect(isFormulaLocation({ host: "excel" })).toBe(false);
    expect(isFormulaLocation(null)).toBe(false);
  });
});
