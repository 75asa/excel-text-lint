import { describe, expect, it } from "vitest";
import type { ExcelLocation } from "../core/locations";
import type { Severity, TextUnit, Violation } from "../core/types";
import {
  buildItems,
  emptyFilter,
  filterItems,
  flattenGroups,
  groupItems,
  isFiltered,
  positionOf,
  stepKey,
  summarize,
  toggleSeverity,
} from "./results";

function cell(sheet: string, address: string, text: string): TextUnit<ExcelLocation> {
  return {
    id: `${sheet}!${address}`,
    text,
    location: { host: "excel", sheet, address, row: 0, col: 0 },
  };
}

function violation(
  unitId: string,
  ruleId: string,
  start: number,
  severity: Severity = "error",
  fix = false,
): Violation {
  return {
    unitId,
    ruleId,
    message: `${ruleId} の違反`,
    severity,
    range: [start, start + 1],
    displayRange: [start, start + 1],
    line: 1,
    column: start + 1,
    ...(fix ? { fix: { range: [start, start + 1] as const, text: "x" } } : {}),
  };
}

const units = [
  cell("Sheet1", "A1", "あいうえお"),
  cell("Sheet1", "A2", "かきくけこ"),
  cell("月次 売上", "B3", "さしすせそ"),
];

// エンジンは入力の順に返すが、同じセルの中の順は保証しないので、あえて崩しておく
const violations = [
  violation("Sheet1!A1", "preset/rule-b", 3, "warning"),
  violation("Sheet1!A1", "preset/rule-a", 1, "error", true),
  violation("月次 売上!B3", "preset/rule-a", 0, "error"),
  violation("Sheet1!A2", "prh", 2, "info", true),
  violation("消えたセル", "prh", 0),
];

const items = buildItems(units, violations);

describe("buildItems", () => {
  it("TextUnit の順 → テキスト内の位置の順に並べ、番地を付ける", () => {
    expect(items.map((item) => [item.location, item.violation.ruleId])).toEqual([
      ["Sheet1!A1", "preset/rule-a"],
      ["Sheet1!A1", "preset/rule-b"],
      ["Sheet1!A2", "prh"],
      ["'月次 売上'!B3", "preset/rule-a"],
    ]);
  });

  it("TextUnit のない違反は捨て、キーは一意", () => {
    expect(items).toHaveLength(4);
    expect(new Set(items.map((item) => item.key)).size).toBe(4);
  });
});

describe("filterItems", () => {
  it("既定では絞り込まない", () => {
    expect(filterItems(items, emptyFilter())).toHaveLength(4);
    expect(isFiltered(emptyFilter())).toBe(false);
  });

  it("重大度・ルール・シートで絞り込む", () => {
    const base = emptyFilter();
    const severities = toggleSeverity(base.severities, "error");
    expect(filterItems(items, { ...base, severities }).map((i) => i.violation.severity)).toEqual([
      "warning",
      "info",
    ]);
    expect(filterItems(items, { ...base, ruleId: "preset/rule-a" })).toHaveLength(2);
    expect(filterItems(items, { ...base, containerKey: "excel:月次 売上" })).toHaveLength(1);
    expect(isFiltered({ ...base, severities })).toBe(true);
  });

  it("toggleSeverity は元の Set を変えない", () => {
    const base = emptyFilter().severities;
    const next = toggleSeverity(base, "info");
    expect(base.has("info")).toBe(true);
    expect(next.has("info")).toBe(false);
    expect(toggleSeverity(next, "info").has("info")).toBe(true);
  });
});

describe("groupItems", () => {
  it("シート別は現れた順", () => {
    const groups = groupItems(items, "container");
    expect(groups.map((g) => [g.label, g.items.length])).toEqual([
      ["Sheet1", 3],
      ["月次 売上", 1],
    ]);
  });

  it("ルール別は件数の多い順で、プリセット名を補足にする", () => {
    const groups = groupItems(items, "rule");
    expect(groups.map((g) => [g.label, g.detail, g.items.length])).toEqual([
      ["rule-a", "preset", 2],
      ["prh", undefined, 1],
      ["rule-b", "preset", 1],
    ]);
  });

  it("まとめないときは 1 グループ、空なら 0 グループ", () => {
    expect(groupItems(items, "none")).toHaveLength(1);
    expect(groupItems([], "none")).toEqual([]);
  });

  it("flattenGroups はグループの表示順に並べる", () => {
    const order = flattenGroups(groupItems(items, "rule")).map((i) => i.violation.ruleId);
    expect(order).toEqual(["preset/rule-a", "preset/rule-a", "prh", "preset/rule-b"]);
  });
});

describe("summarize", () => {
  it("重大度・セル・修正案・ルール・シートごとの件数を数える", () => {
    const summary = summarize(items);
    expect(summary.total).toBe(4);
    expect(summary.bySeverity).toEqual({ error: 2, warning: 1, info: 1 });
    expect(summary.units).toBe(3);
    expect(summary.fixable).toBe(2);
    expect(summary.rules[0]).toEqual({ key: "preset/rule-a", label: "rule-a", count: 2 });
    expect(summary.containers.map((c) => [c.label, c.count])).toEqual([
      ["Sheet1", 3],
      ["月次 売上", 1],
    ]);
  });

  it("0 件", () => {
    expect(summarize([])).toMatchObject({ total: 0, units: 0, rules: [], containers: [] });
  });
});

describe("stepKey / positionOf", () => {
  const order = ["a", "b", "c"];

  it("次へ・前へで移動し、端では反対側に回る", () => {
    expect(stepKey(order, "a", 1)).toBe("b");
    expect(stepKey(order, "c", 1)).toBe("a");
    expect(stepKey(order, "a", -1)).toBe("c");
  });

  it("未選択（や一覧にないキー）なら、次へは先頭、前へは末尾", () => {
    expect(stepKey(order, null, 1)).toBe("a");
    expect(stepKey(order, "z", -1)).toBe("c");
    expect(stepKey([], null, 1)).toBeNull();
  });

  it("位置は 1 始まり、なければ 0", () => {
    expect(positionOf(order, "b")).toBe(2);
    expect(positionOf(order, null)).toBe(0);
    expect(positionOf(order, "z")).toBe(0);
  });
});
