/**
 * ホストの能力（capabilities）と違反の状態から、行や詳細に出す操作を決める純関数。
 *
 * UI はここの結果だけを見てボタンを出し分ける。アダプタが reveal / applyFix を実装したら、
 * capabilities を変えるだけでボタンが出る。
 */

import type {
  ApplyFixResult,
  Fix,
  HostCapabilities,
  RevealPrecision,
  Violation,
} from "../core/types";
import type { ResultItem } from "./results";

/** 違反ごとの、操作のあとの状態。 */
export type ItemState =
  /** 修正を適用した。 */
  | "fixed"
  /** 同じ TextUnit の別の違反を修正した、またはホストのテキストが変わっていた。もう一度 lint が必要。 */
  | "stale";

export interface ItemActions {
  /** 「移動」ボタンを出すか。 */
  reveal: boolean;
  /** 「移動」ボタンの文言。 */
  revealLabel: string;
  /** 「修正」ボタンを出すか。 */
  fix: boolean;
}

export function itemActions(
  capabilities: HostCapabilities,
  violation: Violation,
  state: ItemState | undefined,
): ItemActions {
  return {
    reveal: capabilities.reveal !== "none",
    revealLabel: capabilities.reveal === "container" ? "開く" : "移動",
    fix: capabilities.applyFix !== "none" && violation.fix !== undefined && state === undefined,
  };
}

/** applyFix の単位についての注意書き。なければ null。 */
export function applyFixNote(capabilities: HostCapabilities, unitNoun: string): string | null {
  switch (capabilities.applyFix) {
    case "unit":
      return `${unitNoun}の値を丸ごと置き換えます。${unitNoun}内の一部だけに付けた書式は失われることがあります。`;
    case "selection":
      return "ホストで該当箇所を選択してから適用してください。選択範囲を置き換えます。";
    default:
      return null;
  }
}

/**
 * 修正を適用したあとの状態を返す。
 *
 * 適用した違反（`keys`）は `fixed`。同じ TextUnit のほかの違反は、テキストが変わって範囲がずれるので `stale` にする。
 */
export function markFixed(
  items: readonly ResultItem[],
  states: ReadonlyMap<string, ItemState>,
  keys: readonly string[],
): Map<string, ItemState> {
  const next = new Map(states);
  const fixed = new Set(keys);
  const units = new Set(
    items.filter((item) => fixed.has(item.key)).map((item) => item.violation.unitId),
  );
  for (const item of items) {
    if (!units.has(item.violation.unitId)) continue;
    next.set(item.key, fixed.has(item.key) ? "fixed" : "stale");
  }
  return next;
}

/** 同じ TextUnit の違反をすべて `stale` にする（ホストのテキストが変わっていたとき）。 */
export function markStale(
  items: readonly ResultItem[],
  states: ReadonlyMap<string, ItemState>,
  unitId: string,
): Map<string, ItemState> {
  const next = new Map(states);
  for (const item of items) {
    if (item.violation.unitId === unitId && next.get(item.key) !== "fixed")
      next.set(item.key, "stale");
  }
  return next;
}

// ---------------------------------------------------------------------------
// すべて適用
// ---------------------------------------------------------------------------

/** 1 つの TextUnit にまとめて適用する修正。 */
export interface FixBatch<L = unknown> {
  item: ResultItem<L>;
  /** 同じ TextUnit の修正を 1 つにまとめたもの。 */
  fix: Fix;
  /** まとめた違反のキー。 */
  keys: string[];
}

/**
 * 同じテキストへの複数の修正案を、1 つの修正にまとめる。
 *
 * applyFix は 1 回ごとに stale チェックをするので、同じセルに 2 回適用すると 2 回目は必ず stale になる。
 * そのため、範囲の重ならない修正を開始位置の順に並べ、最初の開始位置から最後の終了位置までを
 * 1 つの置き換えにする。重なる修正は、先に来たものだけを使う。範囲がテキストの外にある修正は使わない。
 *
 * @returns まとめた修正と、使った修正の添字。使える修正がなければ null
 */
export function mergeFixes(
  text: string,
  fixes: readonly Fix[],
): { fix: Fix; used: number[] } | null {
  const order = fixes
    .map((fix, index) => ({ fix, index }))
    .filter(
      ({ fix: { range } }) => 0 <= range[0] && range[0] <= range[1] && range[1] <= text.length,
    )
    .sort((a, b) => a.fix.range[0] - b.fix.range[0] || a.fix.range[1] - b.fix.range[1]);

  const used: { fix: Fix; index: number }[] = [];
  for (const entry of order) {
    const last = used[used.length - 1];
    // 重なる修正は捨てる（同じ位置への挿入が 2 つある場合も、後のものは捨てる）
    if (last && entry.fix.range[0] < last.fix.range[1]) continue;
    if (
      last &&
      entry.fix.range[0] === last.fix.range[0] &&
      entry.fix.range[0] === entry.fix.range[1]
    )
      continue;
    used.push(entry);
  }
  const first = used[0];
  const last = used[used.length - 1];
  if (!first || !last) return null;

  let replacement = "";
  let cursor = first.fix.range[0];
  for (const { fix } of used) {
    replacement += text.slice(cursor, fix.range[0]) + fix.text;
    cursor = fix.range[1];
  }
  return {
    fix: { range: [first.fix.range[0], last.fix.range[1]], text: replacement },
    used: used.map(({ index }) => index),
  };
}

/** 修正案のある違反を TextUnit ごとにまとめる（すでに修正済み・要再チェックのものは除く）。 */
export function planFixes<L>(
  items: readonly ResultItem<L>[],
  states: ReadonlyMap<string, ItemState>,
): FixBatch<L>[] {
  const byUnit = new Map<string, ResultItem<L>[]>();
  for (const item of items) {
    if (!item.violation.fix || states.has(item.key)) continue;
    const list = byUnit.get(item.violation.unitId);
    if (list) list.push(item);
    else byUnit.set(item.violation.unitId, [item]);
  }
  const batches: FixBatch<L>[] = [];
  for (const group of byUnit.values()) {
    const first = group[0];
    if (!first) continue;
    const merged = mergeFixes(
      first.unit.text,
      group.map((item) => item.violation.fix as Fix),
    );
    if (!merged) continue;
    batches.push({
      item: first,
      fix: merged.fix,
      keys: merged.used.flatMap((index) => group[index]?.key ?? []),
    });
  }
  return batches;
}

export type Tone = "info" | "success" | "warning" | "error";

export interface Message {
  tone: Tone;
  text: string;
}

/**
 * reveal の結果に応じた案内。何も言う必要がなければ null。
 *
 * `expected`（capabilities.reveal）が `none` でないのに `none` が返ったときは、シートの削除やセルのずれで
 * 見つからなかったということなので、もう一度チェックするよう促す。
 */
export function revealMessage(
  result: RevealPrecision,
  expected: RevealPrecision,
  location: string,
  unitNoun: string,
): Message | null {
  switch (result) {
    case "exact":
    case "unit":
      return null;
    case "container":
      return {
        tone: "info",
        text: `${location} を含む場所を開きました。該当箇所は「コピー」して検索してください。`,
      };
    case "none":
      return expected === "none"
        ? {
            tone: "warning",
            text: `${location} へは移動できません。該当箇所は「コピー」して検索してください。`,
          }
        : {
            tone: "warning",
            text: `${location} の${unitNoun}が見つかりませんでした。シートや${unitNoun}が移動・削除された可能性があります。もう一度チェックしてください。`,
          };
  }
}

/**
 * applyFix の結果の種類。core の `ApplyFixResult` に加えて、アダプタが返しうる `skipped`（適用する必要がなかった）も扱う。
 */
export type FixStatus = ApplyFixResult["status"] | "skipped";

/** applyFix の結果に応じた案内。 */
export function applyFixMessage(result: ApplyFixResult, location: string): Message {
  const status = result.status as FixStatus;
  const reason = "reason" in result && result.reason ? `（${result.reason}）` : "";
  switch (status) {
    case "applied":
      return { tone: "success", text: `${location} に修正を適用しました。` };
    case "stale":
      return {
        tone: "warning",
        text: `${location} の内容がチェックしたときから変わっていたので、適用しませんでした。もう一度チェックしてください。`,
      };
    case "unsupported":
      return { tone: "warning", text: `${location} には修正を適用できません${reason}。` };
    case "skipped":
      return {
        tone: "info",
        text: `${location} は変更の必要がなかったので、適用しませんでした${reason}。`,
      };
    case "failed":
      return {
        tone: "error",
        text: `修正の適用に失敗しました: ${errorMessage("error" in result ? result.error : result)}`,
      };
  }
}

/** 「すべて適用」の結果の件数から案内を作る。`counts` は TextUnit の数。 */
export function applyAllMessage(
  counts: Partial<Record<FixStatus, number>>,
  unitNoun: string,
): Message {
  const labels: [FixStatus, string][] = [
    ["applied", "適用"],
    ["stale", "内容が変わっていたため未適用"],
    ["unsupported", "適用できない"],
    ["skipped", "変更不要"],
    ["failed", "失敗"],
  ];
  const parts = labels
    .filter(([status]) => (counts[status] ?? 0) > 0)
    .map(([status, label]) => `${label} ${counts[status]} ${unitNoun}`);
  const text =
    parts.length > 0 ? `修正の結果: ${parts.join("、")}` : "適用する修正はありませんでした。";
  const tone: Tone = counts.failed
    ? "error"
    : counts.stale || counts.unsupported
      ? "warning"
      : counts.applied
        ? "success"
        : "info";
  return { tone, text: counts.stale ? `${text}。もう一度チェックしてください。` : text };
}

export function errorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return String(error);
}
