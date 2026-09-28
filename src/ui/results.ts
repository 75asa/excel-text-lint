/**
 * 違反一覧の表示用ロジック（並べ替え・フィルタ・グルーピング・件数・前後の移動）。
 *
 * DOM に依存しない純関数だけを置く。
 */

import type { Severity, TextUnit, Violation } from "../core/types";
import { type Container, containerOf, formatLocation, SEVERITIES, splitRuleId } from "./format";

/** 一覧の 1 行。 */
export interface ResultItem<L = unknown> {
  /** 1 回の lint の結果の中で一意なキー。 */
  key: string;
  violation: Violation;
  unit: TextUnit<L>;
  /** 表示用の場所（例: `Sheet1!B12`）。 */
  location: string;
  container: Container;
}

/**
 * 違反を、TextUnit の順（collect の順。Excel なら行優先）→ テキスト内の位置の順に並べて一覧の行にする。
 *
 * 対応する TextUnit がない違反は捨てる。
 */
export function buildItems<L>(
  units: readonly TextUnit<L>[],
  violations: readonly Violation[],
): ResultItem<L>[] {
  const order = new Map<string, { unit: TextUnit<L>; index: number }>();
  units.forEach((unit, index) => {
    order.set(unit.id, { unit, index });
  });

  const rows: { item: ResultItem<L>; unitIndex: number; seq: number }[] = [];
  violations.forEach((violation, seq) => {
    const found = order.get(violation.unitId);
    if (!found) return;
    rows.push({
      item: {
        key: `${violation.unitId}#${seq}`,
        violation,
        unit: found.unit,
        location: formatLocation(found.unit.location, found.unit.id),
        container: containerOf(found.unit.location),
      },
      unitIndex: found.index,
      seq,
    });
  });
  rows.sort(
    (a, b) =>
      a.unitIndex - b.unitIndex ||
      a.item.violation.displayRange[0] - b.item.violation.displayRange[0] ||
      a.seq - b.seq,
  );
  return rows.map((row) => row.item);
}

// ---------------------------------------------------------------------------
// フィルタ
// ---------------------------------------------------------------------------

export interface ResultFilter {
  /** 表示する重大度。 */
  severities: ReadonlySet<Severity>;
  /** 表示するルール。null ならすべて。 */
  ruleId: string | null;
  /** 表示する入れ物（シートなど）のキー。null ならすべて。 */
  containerKey: string | null;
}

export const ALL_SEVERITIES: ReadonlySet<Severity> = new Set(SEVERITIES);

export function emptyFilter(): ResultFilter {
  return { severities: ALL_SEVERITIES, ruleId: null, containerKey: null };
}

export function filterItems<L>(
  items: readonly ResultItem<L>[],
  filter: ResultFilter,
): ResultItem<L>[] {
  return items.filter(
    (item) =>
      filter.severities.has(item.violation.severity) &&
      (filter.ruleId === null || item.violation.ruleId === filter.ruleId) &&
      (filter.containerKey === null || item.container.key === filter.containerKey),
  );
}

/** フィルタで何か絞り込んでいるか。 */
export function isFiltered(filter: ResultFilter): boolean {
  return (
    filter.ruleId !== null ||
    filter.containerKey !== null ||
    SEVERITIES.some((s) => !filter.severities.has(s))
  );
}

/** 重大度のフィルタを 1 つ切り替える。 */
export function toggleSeverity(
  severities: ReadonlySet<Severity>,
  severity: Severity,
): ReadonlySet<Severity> {
  const next = new Set(severities);
  if (next.has(severity)) next.delete(severity);
  else next.add(severity);
  return next;
}

// ---------------------------------------------------------------------------
// グルーピング
// ---------------------------------------------------------------------------

/** `container`: シート（スライド・ページ）別、`rule`: ルール別、`none`: まとめない。 */
export type GroupBy = "container" | "rule" | "none";

export interface ResultGroup<L = unknown> {
  key: string;
  label: string;
  /** 補足（ルール別のときのプリセット名など）。 */
  detail?: string;
  items: ResultItem<L>[];
}

/**
 * 一覧をグループに分ける。
 *
 * - `container`: 最初に現れた順（collect の順）
 * - `rule`: 件数の多い順、同じならルール名の順
 * - `none`: 1 つのグループ（key は空）
 *
 * グループの中は元の順（`buildItems` の順）を保つ。
 */
export function groupItems<L>(items: readonly ResultItem<L>[], by: GroupBy): ResultGroup<L>[] {
  if (by === "none") return items.length === 0 ? [] : [{ key: "", label: "", items: [...items] }];

  const groups = new Map<string, ResultGroup<L>>();
  for (const item of items) {
    const key = by === "container" ? item.container.key : item.violation.ruleId;
    let group = groups.get(key);
    if (!group) {
      if (by === "container") {
        group = { key, label: item.container.label, items: [] };
      } else {
        const { preset, name } = splitRuleId(key);
        group = { key, label: name, items: [] };
        if (preset) group.detail = preset;
      }
      groups.set(key, group);
    }
    group.items.push(item);
  }

  const list = [...groups.values()];
  if (by === "rule") list.sort((a, b) => b.items.length - a.items.length || byLabel(a, b));
  return list;
}

/** グループを表示の順に並べた一覧（前へ / 次への順番）。 */
export function flattenGroups<L>(groups: readonly ResultGroup<L>[]): ResultItem<L>[] {
  return groups.flatMap((group) => group.items);
}

// ---------------------------------------------------------------------------
// 件数
// ---------------------------------------------------------------------------

export interface CountEntry {
  key: string;
  label: string;
  count: number;
}

export interface ResultSummary {
  total: number;
  bySeverity: Record<Severity, number>;
  /** 違反のあった TextUnit（セルなど）の数。 */
  units: number;
  /** 修正案のある違反の数。 */
  fixable: number;
  /** ルールごとの件数（多い順）。 */
  rules: CountEntry[];
  /** 入れ物（シートなど）ごとの件数（現れた順）。 */
  containers: CountEntry[];
}

export function summarize(items: readonly ResultItem[]): ResultSummary {
  const bySeverity: Record<Severity, number> = { error: 0, warning: 0, info: 0 };
  const units = new Set<string>();
  const rules = new Map<string, CountEntry>();
  const containers = new Map<string, CountEntry>();
  let fixable = 0;

  for (const { violation, container } of items) {
    bySeverity[violation.severity] += 1;
    units.add(violation.unitId);
    if (violation.fix) fixable += 1;
    increment(rules, violation.ruleId, splitRuleId(violation.ruleId).name);
    increment(containers, container.key, container.label);
  }

  return {
    total: items.length,
    bySeverity,
    units: units.size,
    fixable,
    rules: [...rules.values()].sort((a, b) => b.count - a.count || byLabel(a, b)),
    containers: [...containers.values()],
  };
}

function increment(map: Map<string, CountEntry>, key: string, label: string): void {
  const entry = map.get(key);
  if (entry) entry.count += 1;
  else map.set(key, { key, label, count: 1 });
}

// ---------------------------------------------------------------------------
// 前へ / 次へ
// ---------------------------------------------------------------------------

/**
 * 表示順の一覧 `order`（キーの並び）の中で、`current` から `step` 個先のキーを返す。端では反対側に回る。
 *
 * `current` が一覧にない（未選択・フィルタで隠れた）ときは、次へなら先頭、前へなら末尾を返す。
 * 一覧が空なら null。
 */
export function stepKey(
  order: readonly string[],
  current: string | null,
  step: 1 | -1,
): string | null {
  if (order.length === 0) return null;
  const index = current === null ? -1 : order.indexOf(current);
  const next =
    index < 0 ? (step > 0 ? 0 : order.length - 1) : (index + step + order.length) % order.length;
  return order[next] ?? null;
}

/** `key` が `order` の何番目か（1 始まり）。見つからなければ 0。 */
export function positionOf(order: readonly string[], key: string | null): number {
  return key === null ? 0 : order.indexOf(key) + 1;
}

function byLabel(a: { key: string; label: string }, b: { key: string; label: string }): number {
  return a.label.localeCompare(b.label) || a.key.localeCompare(b.key);
}
