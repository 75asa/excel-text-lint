/**
 * 表示用の文字列を作る純関数（場所・ルール名・重大度・範囲の名前）。
 *
 * DOM にも Office.js にも依存しない。ホストごとの location の中身を見るのはここだけにする。
 */

import type { HostLocation, HostName } from "../core/locations";
import type { Scope, Severity } from "../core/types";

/** グルーピングに使う「入れ物」（Excel のシート、PowerPoint のスライド、OneNote のページなど）。 */
export interface Container {
  /** グループを区別するキー。 */
  key: string;
  /** 表示名。 */
  label: string;
}

/**
 * Excel の数式と同じ規則で、必要ならシート名を `'` で囲む。
 *
 * 英数字・アンダースコア・日本語などだけからなる名前はそのまま。空白や記号を含む名前、数字で始まる名前、
 * セル番地（`A1`・`R1C1`）と紛らわしい名前は囲む。名前の中の `'` は `''` にする。
 */
export function quoteSheetName(name: string): string {
  const plain =
    /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) &&
    !/^[A-Za-z]{1,3}\d+$/.test(name) &&
    !/^[Rr]\d*[Cc]\d*$/.test(name);
  return plain ? name : `'${name.replaceAll("'", "''")}'`;
}

/** Excel のセルの番地（例: `Sheet1!B12`、`'月次 売上'!C3`）。 */
export function formatExcelAddress(sheet: string, address: string): string {
  return `${quoteSheetName(sheet)}!${address}`;
}

const WORD_PARTS: Record<string, string> = {
  body: "本文",
  header: "ヘッダー",
  footer: "フッター",
  footnote: "脚注",
  endnote: "文末脚注",
};

/**
 * 違反の場所を 1 行で表す（例: `Sheet1!B12`、`スライド 2 / タイトル 1`）。
 *
 * location が既知のホストの型でなければ `fallback`（ふつうは TextUnit の ID）を返す。
 */
export function formatLocation(location: unknown, fallback = ""): string {
  const loc = asHostLocation(location);
  if (!loc) return fallback;
  switch (loc.host) {
    case "excel":
      return formatExcelAddress(loc.sheet, loc.address);
    case "word": {
      const part = WORD_PARTS[loc.part] ?? loc.part;
      const table = loc.table
        ? ` / 表 ${loc.table.tableIndex + 1} (${loc.table.row + 1}, ${loc.table.cell + 1})`
        : "";
      return `${part} 段落 ${loc.paragraphIndex + 1}${table}`;
    }
    case "powerpoint":
      return `スライド ${loc.slideIndex + 1} / ${loc.notes ? "ノート" : loc.shapeName}`;
    case "onenote":
      return `${loc.pageTitle || "（無題のページ）"} / 段落 ${loc.path.map((i) => i + 1).join("-")}`;
    case "text":
      return `段落 ${loc.paragraphIndex + 1}`;
  }
}

/** 違反をまとめる入れ物（シート・スライド・ページなど）を返す。 */
export function containerOf(location: unknown): Container {
  const loc = asHostLocation(location);
  if (!loc) return { key: "", label: "（場所不明）" };
  switch (loc.host) {
    case "excel":
      return { key: `excel:${loc.sheet}`, label: loc.sheet };
    case "word":
      return { key: `word:${loc.part}`, label: WORD_PARTS[loc.part] ?? loc.part };
    case "powerpoint":
      return { key: `powerpoint:${loc.slideId}`, label: `スライド ${loc.slideIndex + 1}` };
    case "onenote":
      return { key: `onenote:${loc.pageId}`, label: loc.pageTitle || "（無題のページ）" };
    case "text":
      return { key: "text", label: "テキスト" };
  }
}

/** ホストごとの用語（「シート別」「セル」など）。 */
export interface HostTerms {
  /** 入れ物の名前（グルーピングの見出しに使う）。 */
  container: string;
  /** TextUnit の数え方（「12 セル」の「セル」）。 */
  unit: string;
}

const HOST_TERMS: Record<HostName, HostTerms> = {
  excel: { container: "シート", unit: "セル" },
  word: { container: "部分", unit: "段落" },
  powerpoint: { container: "スライド", unit: "シェイプ" },
  onenote: { container: "ページ", unit: "段落" },
  text: { container: "テキスト", unit: "段落" },
};

export function hostTerms(host: HostName): HostTerms {
  return HOST_TERMS[host];
}

const SCOPE_LABELS: Record<Scope, string> = {
  selection: "選択範囲",
  sheet: "シート",
  workbook: "ブック",
  document: "文書",
  slide: "スライド",
  presentation: "プレゼンテーション",
  page: "ページ",
  section: "セクション",
};

export function scopeLabel(scope: Scope): string {
  return SCOPE_LABELS[scope];
}

const SEVERITY_LABELS: Record<Severity, string> = {
  error: "エラー",
  warning: "警告",
  info: "情報",
};

export function severityLabel(severity: Severity): string {
  return SEVERITY_LABELS[severity];
}

/** 重大度の並び順（重いものが先）。 */
export const SEVERITIES: readonly Severity[] = ["error", "warning", "info"];

/**
 * ルール ID をプリセット名とルール名に分ける。
 *
 * `ja-technical-writing/ja-no-redundant-expression` → `{ preset: "ja-technical-writing", name: "ja-no-redundant-expression" }`。
 * プリセットを通さないルール（`prh` など）は `preset` が空になる。
 */
export function splitRuleId(ruleId: string): { preset: string; name: string } {
  const slash = ruleId.lastIndexOf("/");
  if (slash < 0) return { preset: "", name: ruleId };
  return { preset: ruleId.slice(0, slash), name: ruleId.slice(slash + 1) };
}

/**
 * 数式のセル（など、値をそのまま書き換えられない TextUnit）か。
 *
 * `ExcelLocation.isFormula`（#17 で追加）を見る。持っていない location では false。
 */
export function isFormulaLocation(location: unknown): boolean {
  return (
    typeof location === "object" &&
    location !== null &&
    (location as { isFormula?: unknown }).isFormula === true
  );
}

function asHostLocation(location: unknown): HostLocation | undefined {
  if (typeof location !== "object" || location === null || !("host" in location)) return undefined;
  const { host } = location as { host: unknown };
  return typeof host === "string" && host in HOST_TERMS ? (location as HostLocation) : undefined;
}
