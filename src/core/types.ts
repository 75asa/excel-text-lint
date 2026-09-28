/**
 * ホストに依存しないドキュメント抽象と、ホストアダプタのインターフェース（#9）。
 *
 * 設計の意図は `src/core/README.md` を参照。ここでは Office.js の型を使わないこと。
 */

import type { HostName } from "./locations";

// ---------------------------------------------------------------------------
// テキストの単位
// ---------------------------------------------------------------------------

/**
 * lint の単位になるテキスト（Excel のセル、Word の段落、PowerPoint のシェイプ、OneNote の段落など）。
 *
 * lint は TextUnit 1 件ずつ行う。複数の TextUnit を連結して 1 回で lint してはいけない（件数が増えると超線形に遅くなる。#2 の PoC を参照）。
 *
 * @typeParam L ホストごとの位置情報（`ExcelLocation` など）。
 */
export interface TextUnit<L = unknown> {
  /** 1 回の collect の中で一意な ID。`Violation.unitId` から逆引きするのに使う。 */
  id: string;
  /** lint するテキスト。collect した時点の値で、applyFix の stale チェックにも使う。 */
  text: string;
  /** ホストの中での位置。reveal / highlight / applyFix で使う。core はこの中身を見ない。 */
  location: L;
}

/**
 * `TextUnit.text` の中の位置。UTF-16 のコード単位での `[start, end)`。
 *
 * textlint の `message.range` と同じ単位。
 */
export type TextRange = readonly [start: number, end: number];

// ---------------------------------------------------------------------------
// lint の結果
// ---------------------------------------------------------------------------

/** textlint の severity を文字列にしたもの。 */
export type Severity = "error" | "warning" | "info";

/** 修正案。`range` の部分を `text` に置き換える。 */
export interface Fix {
  range: TextRange;
  text: string;
}

/** 1 件の違反。 */
export interface Violation {
  /** 違反が見つかった `TextUnit.id`。 */
  unitId: string;
  /** ルールの ID（例: `ja-technical-writing/ja-no-redundant-expression`）。 */
  ruleId: string;
  message: string;
  severity: Severity;
  /** textlint が返した範囲そのもの。1 文字だけのこともある。 */
  range: TextRange;
  /**
   * 表示やハイライトに使う範囲。
   *
   * `range` が 1 文字以下のときは `fix.range` と合わせて広げる（`computeDisplayRange`）。
   * 違反箇所を人に見せるときはこちらを使う。
   */
  displayRange: TextRange;
  /** 1 始まりの行番号（セル内改行・段落内改行があるとき）。 */
  line: number;
  /** 1 始まりの列番号。 */
  column: number;
  /** 自動修正できる場合の修正案。 */
  fix?: Fix;
}

// ---------------------------------------------------------------------------
// ホストアダプタ
// ---------------------------------------------------------------------------

/**
 * collect の対象範囲。
 *
 * - `selection`: 選択範囲（全ホスト）
 * - `sheet` / `workbook`: Excel
 * - `document`: Word
 * - `slide` / `presentation`: PowerPoint
 * - `page` / `section`: OneNote
 */
export type Scope =
  | "selection"
  | "sheet"
  | "workbook"
  | "document"
  | "slide"
  | "presentation"
  | "page"
  | "section";

/**
 * reveal（違反箇所への移動）がどこまでできたか。
 *
 * - `exact`: 違反の文字範囲まで選択できた（Word など）
 * - `unit`: TextUnit（セル・段落・シェイプ）までは選択できた（Excel のセルなど）
 * - `container`: TextUnit を含む入れ物（シート・スライド・ページ）を開くところまで（OneNote など）
 * - `none`: 何もできなかった
 *
 * UI は結果を見て、文脈表示や検索語のコピーなどの代わりの手段を出す。
 */
export type RevealPrecision = "exact" | "unit" | "container" | "none";

/**
 * applyFix がどの単位で書き換えられるか。
 *
 * - `range`: 違反の文字範囲だけを書き換えられる（書式を保てる）
 * - `unit`: TextUnit のテキスト全体を置き換える（Excel のセルなど。セル内の部分書式は失われうる）
 * - `selection`: ユーザーの選択範囲を置き換える形でだけできる（OneNote の `setSelectedDataAsync`）
 * - `none`: できない（修正案の表示とコピーだけ）
 */
export type ApplyFixMode = "range" | "unit" | "selection" | "none";

/** ホストごとの能力。UI はこれを見てボタンの有無や説明文を切り替える。 */
export interface HostCapabilities {
  /** collect に渡せる範囲。 */
  scopes: readonly Scope[];
  /** reveal で期待できる最良の精度。実際の結果は reveal の戻り値で返る。 */
  reveal: RevealPrecision;
  /** highlight / clearHighlight が意味を持つか。false のときは何もしない実装でよい。 */
  highlight: boolean;
  applyFix: ApplyFixMode;
  /** `onSelectionChanged` を提供しているか。 */
  selectionTracking: boolean;
}

/** collect の進捗。 */
export interface CollectProgress {
  /** ここまでに集めた TextUnit の数。 */
  collected: number;
  /** 全体の数（分かる場合だけ）。 */
  total?: number;
}

export interface CollectOptions {
  /** 中断用。中断されたら `AbortError`（`signal.reason`）で reject する。 */
  signal?: AbortSignal;
  onProgress?: (progress: CollectProgress) => void;
}

/**
 * applyFix の結果。失敗しうる前提で扱う。
 *
 * - `applied`: 適用した
 * - `stale`: lint したときからテキストが変わっていたので適用しなかった（もう一度 lint を促す）
 * - `unsupported`: このホスト・この単位では適用できない
 * - `failed`: ホストの API がエラーを返した
 */
export type ApplyFixResult =
  | { status: "applied"; text: string }
  | { status: "stale"; currentText: string }
  | { status: "unsupported"; reason?: string }
  | { status: "failed"; error: unknown };

export interface Disposable {
  dispose(): void;
}

/**
 * ホストごとの差分を閉じ込めるアダプタ。
 *
 * 実装は `src/hosts/<host>/` に置く。できない操作は capabilities で宣言し、
 * メソッドは「何もしない」「`none` / `unsupported` を返す」で済ませてよい。
 *
 * @typeParam L そのホストの location 型。
 */
export interface HostAdapter<L = unknown> {
  readonly host: HostName;
  readonly capabilities: HostCapabilities;

  /** 対象範囲のテキストを集める。空のテキストは含めない。 */
  collect(scope: Scope, options?: CollectOptions): Promise<TextUnit<L>[]>;

  /**
   * 違反箇所へ移動（選択）する。`range` を省略したときは TextUnit 全体。
   * @returns 実際にできた精度
   */
  reveal(unit: TextUnit<L>, range?: TextRange): Promise<RevealPrecision>;

  /** 違反箇所をホスト上で目立たせる。capabilities.highlight が false なら何もしない実装でよい。 */
  highlight(
    violations: readonly Violation[],
    units: ReadonlyMap<string, TextUnit<L>>,
  ): Promise<void>;

  /** highlight で付けた書式を元に戻す。 */
  clearHighlight(): Promise<void>;

  /**
   * 修正案を適用する。
   *
   * 実装は、適用の直前にホストから現在のテキストを読み直し、`unit.text` と一致するかを確かめること
   * （一致しなければ `stale`）。書き換え後のテキストは `applyFixToText` で作れる。
   */
  applyFix(unit: TextUnit<L>, fix: Fix): Promise<ApplyFixResult>;

  /**
   * 選択（カーソル位置）が変わったときに、その位置にある TextUnit の ID を知らせる。
   * capabilities.selectionTracking が true のときだけ実装する。
   */
  onSelectionChanged?(callback: (unitId: string | null) => void): Disposable;
}
