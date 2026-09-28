/**
 * textlint の結果を `Violation` に変換する純関数と、修正案の適用。
 */

import type { Fix, Severity, TextRange, Violation } from "./types";

/** textlint の `TextlintMessage` のうち、ここで使う部分。 */
export interface TextlintMessageLike {
  ruleId: string;
  message: string;
  /** 0: none, 1: warning, 2: error, 3: info */
  severity: number;
  index: number;
  line: number;
  column: number;
  range?: readonly [number, number];
  fix?: { range: readonly [number, number]; text: string };
}

/** textlint の severity（数値）を文字列にする。未知の値は `error` として扱う。 */
export function toSeverity(severity: number): Severity {
  switch (severity) {
    case 1:
      return "warning";
    case 3:
      return "info";
    default:
      return "error";
  }
}

/**
 * 表示やハイライトに使う範囲を決める。
 *
 * textlint の `range` は 1 文字（や 0 文字）だけのことがある（例: 「することができる。」の違反が `[6, 7]`）。
 * そのままでは何が悪いのか分かりにくいので、次の順で補正する。
 *
 * 1. `range` が 2 文字以上なら、そのまま使う
 * 2. `fix.range` が `range` に接している（重なるか隣り合う）なら、両方を合わせた範囲にする
 * 3. それでも空なら、1 文字分に広げる（末尾なら前の 1 文字）
 *
 * どの場合も `text` の長さに収め、サロゲートペアの途中で切れないようにする。
 */
export function computeDisplayRange(
  text: string,
  range: TextRange,
  fix?: { range: TextRange },
): TextRange {
  const length = text.length;
  let start = clamp(Math.min(range[0], range[1]), 0, length);
  let end = clamp(Math.max(range[0], range[1]), 0, length);

  if (end - start < 2 && fix) {
    const fixStart = clamp(Math.min(fix.range[0], fix.range[1]), 0, length);
    const fixEnd = clamp(Math.max(fix.range[0], fix.range[1]), 0, length);
    if (fixStart <= end && start <= fixEnd) {
      start = Math.min(start, fixStart);
      end = Math.max(end, fixEnd);
    }
  }

  if (start === end && length > 0) {
    if (end < length) end += 1;
    else start -= 1;
  }

  // サロゲートペアの途中で切らない
  if (
    start > 0 &&
    isLowSurrogate(text.charCodeAt(start)) &&
    isHighSurrogate(text.charCodeAt(start - 1))
  )
    start -= 1;
  if (
    end < length &&
    isLowSurrogate(text.charCodeAt(end)) &&
    isHighSurrogate(text.charCodeAt(end - 1))
  )
    end += 1;

  return [start, end];
}

/** textlint の 1 件のメッセージを Violation にする。 */
export function toViolation(unitId: string, text: string, message: TextlintMessageLike): Violation {
  const range: TextRange = message.range
    ? [message.range[0], message.range[1]]
    : [message.index, message.index + 1];
  const fix: Fix | undefined = message.fix
    ? { range: [message.fix.range[0], message.fix.range[1]], text: message.fix.text }
    : undefined;
  return {
    unitId,
    ruleId: message.ruleId,
    message: message.message,
    severity: toSeverity(message.severity),
    range,
    displayRange: computeDisplayRange(text, range, fix),
    line: message.line,
    column: message.column,
    ...(fix ? { fix } : {}),
  };
}

/** 修正案をテキストに適用した結果を返す。範囲がテキストの外なら例外を投げる。 */
export function applyFixToText(text: string, fix: Fix): string {
  const [start, end] = fix.range;
  if (!(0 <= start && start <= end && end <= text.length)) {
    throw new RangeError(
      `修正案の範囲 [${start}, ${end}] がテキスト（${text.length} 文字）の外にあります`,
    );
  }
  return text.slice(0, start) + fix.text + text.slice(end);
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function isHighSurrogate(code: number): boolean {
  return code >= 0xd800 && code <= 0xdbff;
}

function isLowSurrogate(code: number): boolean {
  return code >= 0xdc00 && code <= 0xdfff;
}
