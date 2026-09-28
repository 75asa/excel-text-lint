import { describe, expect, it } from "vitest";
import { applyFixToText, computeDisplayRange, toSeverity, toViolation } from "./violation";

describe("computeDisplayRange", () => {
  it("2 文字以上の range はそのまま使う", () => {
    expect(computeDisplayRange("データを取得する", [0, 3])).toEqual([0, 3]);
  });

  it("1 文字の range は fix.range と合わせて広げる（PoC の例）", () => {
    const text = "データを取得することができる。";
    expect(computeDisplayRange(text, [6, 7], { range: [6, 15] })).toEqual([6, 15]);
  });

  it("fix.range が range と隣り合っていれば合わせる", () => {
    expect(computeDisplayRange("abcdef", [2, 3], { range: [3, 5] })).toEqual([2, 5]);
  });

  it("fix.range が range から離れていれば使わない", () => {
    expect(computeDisplayRange("abcdefgh", [1, 2], { range: [5, 7] })).toEqual([1, 2]);
  });

  it("fix が挿入（空の範囲）のときは 1 文字の range のまま", () => {
    expect(computeDisplayRange("abc", [1, 2], { range: [1, 1] })).toEqual([1, 2]);
  });

  it("空の range は次の 1 文字に広げる", () => {
    expect(computeDisplayRange("abc", [1, 1])).toEqual([1, 2]);
  });

  it("末尾の空の range は前の 1 文字に広げる", () => {
    expect(computeDisplayRange("abc", [3, 3])).toEqual([2, 3]);
  });

  it("空のテキストでは空の範囲を返す", () => {
    expect(computeDisplayRange("", [0, 0])).toEqual([0, 0]);
  });

  it("テキストの外の range はテキストの中に収める", () => {
    expect(computeDisplayRange("abc", [2, 10])).toEqual([2, 3]);
    expect(computeDisplayRange("abc", [5, 6])).toEqual([2, 3]);
  });

  it("サロゲートペアの途中で切らない", () => {
    const text = "a𩸽b"; // 𩸽 は 2 コード単位
    expect(computeDisplayRange(text, [1, 2])).toEqual([1, 3]);
    expect(computeDisplayRange(text, [2, 3])).toEqual([1, 3]);
  });
});

describe("toSeverity", () => {
  it("textlint の数値を文字列にする", () => {
    expect(toSeverity(1)).toBe("warning");
    expect(toSeverity(2)).toBe("error");
    expect(toSeverity(3)).toBe("info");
    expect(toSeverity(99)).toBe("error");
  });
});

describe("toViolation", () => {
  it("textlint のメッセージを Violation にする", () => {
    const text = "データを取得することができる。";
    const violation = toViolation("Sheet1!A1", text, {
      ruleId: "ja-technical-writing/ja-no-redundant-expression",
      message: "冗長な表現です",
      severity: 2,
      index: 6,
      line: 1,
      column: 7,
      range: [6, 7],
      fix: { range: [6, 15], text: "できる。" },
    });
    expect(violation).toEqual({
      unitId: "Sheet1!A1",
      ruleId: "ja-technical-writing/ja-no-redundant-expression",
      message: "冗長な表現です",
      severity: "error",
      range: [6, 7],
      displayRange: [6, 15],
      line: 1,
      column: 7,
      fix: { range: [6, 15], text: "できる。" },
    });
  });

  it("range がなければ index から 1 文字の範囲を作る", () => {
    const violation = toViolation("u", "abc", {
      ruleId: "r",
      message: "m",
      severity: 1,
      index: 1,
      line: 1,
      column: 2,
    });
    expect(violation.range).toEqual([1, 2]);
    expect(violation.fix).toBeUndefined();
    expect("fix" in violation).toBe(false);
  });
});

describe("applyFixToText", () => {
  it("範囲を置き換える", () => {
    expect(
      applyFixToText("データを取得することができる。", { range: [6, 15], text: "できる。" }),
    ).toBe("データを取得できる。");
  });

  it("挿入と削除", () => {
    expect(applyFixToText("abc", { range: [1, 1], text: "X" })).toBe("aXbc");
    expect(applyFixToText("abc", { range: [0, 2], text: "" })).toBe("c");
  });

  it("範囲がテキストの外なら例外", () => {
    expect(() => applyFixToText("abc", { range: [2, 5], text: "" })).toThrow(RangeError);
  });
});
