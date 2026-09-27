/**
 * ホストごとの location 型（`TextUnit<L>` の `L`）。
 *
 * core はこの中身を見ない。アダプタと、ホストごとに表示を変えたい UI だけが使う。
 * どれも純粋なデータ（Office.js のプロキシオブジェクトを持たない）にすること。
 * `context.sync()` の外に持ち出せ、Worker とのやり取りや保存にも使えるようにするため。
 *
 * Excel 以外はまだ実装がない。各ホストの API の制約から考えた「案」で、実装のときに見直す。
 */

export type HostName = "excel" | "word" | "powerpoint" | "onenote" | "text";

/** Excel のセル。 */
export interface ExcelLocation {
  host: "excel";
  /** ワークシート名。 */
  sheet: string;
  /** シート名を含まない A1 形式のアドレス（例: `B3`）。 */
  address: string;
  /** 0 始まりの行番号。 */
  row: number;
  /** 0 始まりの列番号。 */
  col: number;
}

/**
 * Word の段落（案、#21 で見直す）。
 *
 * 段落は `Paragraph.uniqueLocalId`（WordApi 1.6）で探し、取れない環境では paragraphIndex とテキストの一致で探し直す。
 * 違反の範囲は `Paragraph.search()` か `getRange()` と文字数で選択する想定。
 */
export interface WordLocation {
  host: "word";
  part: "body" | "header" | "footer" | "footnote" | "endnote";
  paragraphId?: string;
  /** 本文の中の段落の添字。表の中の段落も通し番号で数える。 */
  paragraphIndex: number;
  table?: { tableIndex: number; row: number; cell: number };
}

/**
 * PowerPoint のシェイプ内のテキスト（案、#22 で見直す）。
 *
 * PowerPoint の API で選択できるのはスライドとシェイプまで（文字範囲の選択は `TextRange.select()` が使える環境だけ）。
 */
export interface PowerPointLocation {
  host: "powerpoint";
  slideId: string;
  /** 0 始まりのスライド番号（表示用）。 */
  slideIndex: number;
  shapeId: string;
  shapeName: string;
  /** 表のセルの場合。 */
  table?: { row: number; col: number };
  /** ノート（スピーカーノート）の場合 true。 */
  notes?: boolean;
}

/**
 * OneNote の段落（案、docs/research/onenote.md の「#9 への提言」から）。
 *
 * 段落の ID は編集で変わる可能性がある。ID が見つからないときは path とテキストの一致で探し直す。
 */
export interface OneNoteLocation {
  host: "onenote";
  pageId: string;
  /** `navigateToPageWithClientUrl` 用。 */
  pageClientUrl: string;
  pageTitle: string;
  kind: "title" | "paragraph" | "listItem" | "tableCell";
  outlineId?: string;
  paragraphId?: string;
  /** outline / 段落 / 子段落の添字の並び（ID が変わったときの予備）。 */
  path: number[];
  table?: { row: number; cell: number };
  list?: { type: "Number" | "Bullet"; index: number };
}

/**
 * ホストを持たないテキスト（Loop 向けの貼り付け UI やテスト用）。docs/research/loop.md を参照。
 */
export interface PlainTextLocation {
  host: "text";
  /** 貼り付けたテキストの中の段落番号（0 始まり）。 */
  paragraphIndex: number;
}

export type HostLocation = ExcelLocation | WordLocation | PowerPointLocation | OneNoteLocation | PlainTextLocation;
