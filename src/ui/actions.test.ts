import { describe, expect, it } from "vitest";
import type { HostCapabilities, TextUnit, Violation } from "../core/types";
import {
  applyAllMessage,
  applyFixMessage,
  itemActions,
  markFixed,
  markStale,
  mergeFixes,
  planFixes,
  revealMessage,
} from "./actions";
import { buildItems } from "./results";

const none: HostCapabilities = {
  scopes: ["selection"],
  reveal: "none",
  highlight: false,
  applyFix: "none",
  selectionTracking: false,
};
const excel: HostCapabilities = {
  scopes: ["selection", "sheet", "workbook"],
  reveal: "unit",
  highlight: true,
  applyFix: "unit",
  selectionTracking: true,
};

function v(unitId: string, start: number, end: number, fixText?: string): Violation {
  return {
    unitId,
    ruleId: "r",
    message: "m",
    severity: "error",
    range: [start, end],
    displayRange: [start, end],
    line: 1,
    column: start + 1,
    ...(fixText !== undefined ? { fix: { range: [start, end] as const, text: fixText } } : {}),
  };
}

const units: TextUnit[] = [
  { id: "u1", text: "することができる。することができる。", location: null },
  { id: "u2", text: "abc", location: null },
];
const items = buildItems(units, [
  v("u1", 2, 8, "できる"),
  v("u1", 11, 17, "できる"),
  v("u1", 0, 2),
  v("u2", 0, 1, "A"),
]);

describe("itemActions", () => {
  it("capabilities が none なら、移動も修正も出さない", () => {
    const actions = itemActions(none, items[0]!.violation, undefined);
    expect(actions).toMatchObject({ reveal: false, fix: false });
  });

  it("修正案があって applyFix できるときだけ修正を出す", () => {
    const withFix = items.find((i) => i.violation.fix)!;
    const withoutFix = items.find((i) => !i.violation.fix)!;
    expect(itemActions(excel, withFix.violation, undefined)).toMatchObject({
      reveal: true,
      revealLabel: "移動",
      fix: true,
    });
    expect(itemActions(excel, withoutFix.violation, undefined).fix).toBe(false);
    expect(itemActions(excel, withFix.violation, "fixed").fix).toBe(false);
    expect(itemActions(excel, withFix.violation, "stale").fix).toBe(false);
  });

  it("reveal が container のホストでは「開く」", () => {
    expect(
      itemActions({ ...none, reveal: "container" }, items[0]!.violation, undefined),
    ).toMatchObject({ reveal: true, revealLabel: "開く" });
  });
});

describe("markFixed / markStale", () => {
  it("適用した違反は fixed、同じ TextUnit のほかの違反は stale、ほかの TextUnit は変えない", () => {
    const first = items.find((i) => i.violation.unitId === "u1")!;
    const states = markFixed(items, new Map(), [first.key]);
    const u1 = items.filter((i) => i.violation.unitId === "u1");
    expect(u1.map((i) => states.get(i.key))).toEqual(["fixed", "stale", "stale"]);
    const u2 = items.find((i) => i.violation.unitId === "u2")!;
    expect(states.has(u2.key)).toBe(false);
  });

  it("markStale は fixed を上書きしない", () => {
    const first = items.find((i) => i.violation.fix && i.violation.unitId === "u1")!;
    const states = markStale(items, markFixed(items, new Map(), [first.key]), "u1");
    expect(states.get(first.key)).toBe("fixed");
  });
});

describe("mergeFixes", () => {
  const text = "することができる。することができる。";

  it("同じテキストへの複数の修正を 1 つにまとめる", () => {
    const merged = mergeFixes(text, [
      { range: [11, 17], text: "できる" },
      { range: [2, 8], text: "できる" },
    ]);
    expect(merged?.used).toEqual([1, 0]);
    expect(merged?.fix).toEqual({ range: [2, 17], text: "できる。するできる" });
    const [start, end] = merged!.fix.range;
    expect(text.slice(0, start) + merged!.fix.text + text.slice(end)).toBe(
      "するできる。するできる。",
    );
  });

  it("重なる修正は先のものだけ、範囲外は捨てる", () => {
    const merged = mergeFixes("abcdef", [
      { range: [1, 3], text: "X" },
      { range: [2, 4], text: "Y" },
      { range: [5, 99], text: "Z" },
    ]);
    expect(merged).toEqual({ fix: { range: [1, 3], text: "X" }, used: [0] });
    expect(mergeFixes("abc", [])).toBeNull();
  });
});

describe("planFixes", () => {
  it("TextUnit ごとにまとめ、修正済み・要再チェックは除く", () => {
    const batches = planFixes(items, new Map());
    expect(batches.map((b) => [b.item.unit.id, b.keys.length])).toEqual([
      ["u1", 2],
      ["u2", 1],
    ]);
    const u2 = items.find((i) => i.violation.unitId === "u2")!;
    expect(planFixes(items, new Map([[u2.key, "fixed" as const]]))).toHaveLength(1);
  });
});

describe("メッセージ", () => {
  it("reveal: 期待どおりなら何も言わず、セルが見つからなければ再チェックを促す", () => {
    expect(revealMessage("unit", "unit", "Sheet1!A1", "セル")).toBeNull();
    expect(revealMessage("none", "unit", "Sheet1!A1", "セル")?.text).toContain(
      "セルが見つかりませんでした",
    );
    expect(revealMessage("none", "none", "Sheet1!A1", "セル")?.text).toContain("コピー");
  });

  it("applyFix: 結果ごとに案内する（skipped も扱う）", () => {
    expect(applyFixMessage({ status: "applied", text: "x" }, "A1").tone).toBe("success");
    expect(applyFixMessage({ status: "stale", currentText: "y" }, "A1").text).toContain(
      "もう一度チェック",
    );
    expect(applyFixMessage({ status: "unsupported", reason: "数式のセル" }, "A1").text).toContain(
      "（数式のセル）",
    );
    expect(applyFixMessage({ status: "failed", error: new Error("boom") }, "A1").text).toContain(
      "boom",
    );
    const skipped = { status: "skipped", reason: "同じ内容" } as unknown as Parameters<
      typeof applyFixMessage
    >[0];
    expect(applyFixMessage(skipped, "A1")).toMatchObject({ tone: "info" });
  });

  it("すべて適用: 件数をまとめる", () => {
    expect(applyAllMessage({ applied: 3 }, "セル")).toEqual({
      tone: "success",
      text: "修正の結果: 適用 3 セル",
    });
    const mixed = applyAllMessage({ applied: 1, stale: 1, unsupported: 2 }, "セル");
    expect(mixed.tone).toBe("warning");
    expect(mixed.text).toBe(
      "修正の結果: 適用 1 セル、内容が変わっていたため未適用 1 セル、適用できない 2 セル。もう一度チェックしてください。",
    );
  });
});
