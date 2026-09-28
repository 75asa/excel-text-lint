/**
 * 実行の状態（収集・エンジンの起動・辞書の読み込み・lint・完了・中止・エラー）と、その表示の仕方。
 */

import type { Scope } from "../core/types";
import { scopeLabel } from "./format";

export type RunPhase =
  | { kind: "idle" }
  | { kind: "collecting"; scope: Scope; collected: number; total?: number }
  | { kind: "starting" }
  | {
      kind: "linting";
      done: number;
      total: number;
      violations: number;
      /** 最初の lint が終わって、辞書の読み込みが済んでいるか。 */
      warm: boolean;
      /** 残り時間の見積もり（ミリ秒）。見積もれないときは省略。 */
      remainingMs?: number;
    }
  | { kind: "done"; scope: Scope; units: number; violations: number }
  | { kind: "empty"; scope: Scope }
  | { kind: "cancelled" }
  | { kind: "error"; message: string };

/** textlint のエンジン（Worker）の状態。 */
export type EngineState = "starting" | "ready" | "failed";

export type PhaseTone = "neutral" | "busy" | "success" | "warning" | "error";

export interface PhaseView {
  text: string;
  /** 補足（初回の辞書の読み込みの説明など）。 */
  hint?: string;
  tone: PhaseTone;
  /** 進捗バー。null なら出さない。`max` がなければ不確定（アニメーション）の表示。 */
  progress: { value: number; max: number } | { indeterminate: true } | null;
  /** 実行中か（実行ボタンを止め、中止ボタンを出す）。 */
  busy: boolean;
}

/** 辞書のおおよその大きさ（kuromoji の辞書。gzip のまま約 15MB）。 */
const DICT_SIZE = "約 15MB";

/**
 * 実行の状態を、ステータス行と進捗バーの表示に変換する。
 *
 * @param unitNoun TextUnit の数え方（Excel なら「セル」）
 */
export function describePhase(phase: RunPhase, engine: EngineState, unitNoun: string): PhaseView {
  switch (phase.kind) {
    case "idle":
      if (engine === "starting")
        return {
          text: "textlint を準備しています…",
          tone: "busy",
          progress: null,
          busy: false,
        };
      if (engine === "failed")
        return {
          text: "textlint の起動に失敗しました。実行するともう一度試します。",
          tone: "warning",
          progress: null,
          busy: false,
        };
      return {
        text: "範囲を選んで「チェック」を押してください。",
        tone: "neutral",
        progress: null,
        busy: false,
      };
    case "collecting":
      return {
        text:
          phase.collected > 0
            ? `${scopeLabel(phase.scope)}のテキストを読み取っています…（${phase.collected}${phase.total ? ` / ${phase.total}` : ""} ${unitNoun}）`
            : `${scopeLabel(phase.scope)}のテキストを読み取っています…`,
        tone: "busy",
        progress:
          phase.total && phase.total > 0
            ? { value: phase.collected, max: phase.total }
            : { indeterminate: true },
        busy: true,
      };
    case "starting":
      return {
        text: "textlint を起動しています…",
        tone: "busy",
        progress: { indeterminate: true },
        busy: true,
      };
    case "linting":
      if (!phase.warm && phase.done === 0)
        return {
          text: `辞書（${DICT_SIZE}）を読み込んでいます…`,
          hint: "初回だけ時間がかかります。2 回目以降はキャッシュを使います。",
          tone: "busy",
          progress: { indeterminate: true },
          busy: true,
        };
      return {
        text: `チェックしています… ${phase.done} / ${phase.total} ${unitNoun}（違反 ${phase.violations} 件）`,
        ...(phase.remainingMs !== undefined
          ? { hint: `残り ${formatDuration(phase.remainingMs)}` }
          : {}),
        tone: "busy",
        progress: { value: phase.done, max: phase.total },
        busy: true,
      };
    case "done":
      return phase.violations === 0
        ? {
            text: `${phase.units} ${unitNoun}をチェックしました。問題は見つかりませんでした。`,
            tone: "success",
            progress: null,
            busy: false,
          }
        : {
            text: `${phase.units} ${unitNoun}をチェックし、${phase.violations} 件の違反が見つかりました。`,
            tone: "warning",
            progress: null,
            busy: false,
          };
    case "empty":
      return {
        text: `${scopeLabel(phase.scope)}にチェックする文字列がありません。`,
        tone: "neutral",
        progress: null,
        busy: false,
      };
    case "cancelled":
      return { text: "中止しました。", tone: "neutral", progress: null, busy: false };
    case "error":
      return {
        text: `チェックに失敗しました: ${phase.message}`,
        tone: "error",
        progress: null,
        busy: false,
      };
  }
}

/** 進捗の記録（何件目がいつ終わったか）。 */
export interface ProgressSample {
  done: number;
  /** ミリ秒（`performance.now()` など）。 */
  at: number;
}

/**
 * これまでの速さから、残り時間（ミリ秒）を見積もる。
 *
 * `start` は辞書を読み込んだあとの最初の記録にする（辞書の読み込みの時間を速さに含めないため）。
 * 件数か経過時間が少なすぎて当てにならないときは null。
 */
export function estimateRemainingMs(
  start: ProgressSample,
  now: ProgressSample,
  total: number,
): number | null {
  const done = now.done - start.done;
  const elapsed = now.at - start.at;
  if (done < 5 || elapsed < 1000 || now.done >= total) return null;
  return ((total - now.done) * elapsed) / done;
}

/** 残り時間を「約 20 秒」「約 3 分」のように丸める。 */
export function formatDuration(ms: number): string {
  const seconds = Math.max(1, Math.round(ms / 1000));
  if (seconds < 60) return `約 ${seconds < 10 ? seconds : Math.round(seconds / 5) * 5} 秒`;
  return `約 ${Math.round(seconds / 60)} 分`;
}

/**
 * 実行の前に出す、範囲についての注意。なければ null。
 *
 * シート・ブック全体はセルが多くなりやすく、時間のほとんどは lint にかかる（1 万セルで約 7 秒、5 万セルで約 30 秒。#17）。
 */
export function scopeNotice(scope: Scope, unitNoun: string): string | null {
  switch (scope) {
    case "workbook":
    case "presentation":
    case "section":
    case "document":
      return `${scopeLabel(scope)}全体は${unitNoun}が多いと時間がかかります（目安: 1 万${unitNoun}で約 7 秒）。途中で中止できます。`;
    case "sheet":
      return `${unitNoun}が多いシートは時間がかかることがあります。途中で中止できます。`;
    default:
      return null;
  }
}
