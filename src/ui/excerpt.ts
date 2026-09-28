/**
 * 違反箇所を前後の文脈つきで見せるための抜粋を作る純関数。
 */

import type { Fix, TextRange } from "../core/types";

export interface Excerpt {
  /** 違反箇所の前の文脈。 */
  before: string;
  /** 違反箇所（強調して表示する部分）。 */
  match: string;
  /** 違反箇所の後の文脈。 */
  after: string;
  /** 前を切り詰めたか（先頭に「…」を付ける）。 */
  clippedStart: boolean;
  /** 後ろを切り詰めたか（末尾に「…」を付ける）。 */
  clippedEnd: boolean;
}

export interface ExcerptOptions {
  /** 前後に付ける文脈の文字数（UTF-16 のコード単位。サロゲートペアの途中では切らない）。既定は 16。 */
  context?: number;
  /** 改行を `↵` に置き換えて 1 行にする（一覧用）。既定は false。 */
  singleLine?: boolean;
}

/**
 * `text` の `range` の部分を、前後 `context` 文字の文脈と一緒に取り出す。
 *
 * 範囲はテキストの長さに収める。文脈の切れ目がサロゲートペアの途中に来るときは、ペアを落とす側に寄せる
 * （文字化けした半端な文字を出さない）。
 */
export function buildExcerpt(
  text: string,
  range: TextRange,
  options: ExcerptOptions = {},
): Excerpt {
  const context = Math.max(0, options.context ?? 16);
  const length = text.length;
  const start = clamp(Math.min(range[0], range[1]), 0, length);
  const end = clamp(Math.max(range[0], range[1]), 0, length);

  let from = Math.max(0, start - context);
  if (from > 0 && isLowSurrogate(text.charCodeAt(from))) from += 1;
  let to = Math.min(length, end + context);
  if (to < length && isLowSurrogate(text.charCodeAt(to))) to -= 1;

  const shape = options.singleLine ? toSingleLine : (s: string) => s;
  return {
    before: shape(text.slice(from, start)),
    match: shape(text.slice(start, end)),
    after: shape(text.slice(end, to)),
    clippedStart: from > 0,
    clippedEnd: to < length,
  };
}

export interface FixPreview {
  before: string;
  /** 消える部分。 */
  removed: string;
  /** 入る部分。 */
  inserted: string;
  after: string;
  clippedStart: boolean;
  clippedEnd: boolean;
}

/** 修正案の適用前後を見せるための抜粋。`removed` を `inserted` に置き換える形で表示する。 */
export function buildFixPreview(text: string, fix: Fix, options: ExcerptOptions = {}): FixPreview {
  const { before, match, after, clippedStart, clippedEnd } = buildExcerpt(text, fix.range, options);
  const inserted = options.singleLine ? toSingleLine(fix.text) : fix.text;
  return { before, removed: match, inserted, after, clippedStart, clippedEnd };
}

function toSingleLine(value: string): string {
  return value.replace(/\r\n|\r|\n/g, "↵");
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function isLowSurrogate(code: number): boolean {
  return code >= 0xdc00 && code <= 0xdfff;
}
