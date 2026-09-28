import type { ExcelLocation } from "../../core/locations";
import type { ApplyFixResult, Fix, TextUnit } from "../../core/types";
import { applyFixToText } from "../../core/violation";

/** 複数の修正案をまとめて適用した結果。 */
export interface FixesAppliedToText {
  text: string;
  /** 適用した修正案（後ろの位置から順）。 */
  applied: Fix[];
  /** ほかの修正案と範囲が重なる、またはテキストの外にあるので適用しなかった修正案。 */
  skipped: Fix[];
}

/**
 * 1 つのテキストに複数の修正案を適用する。
 *
 * 後ろの位置から順に適用するので、前にある修正案の範囲がずれない。
 * すでに適用した修正案と範囲が重なるものは飛ばす（`skipped`）。同じ位置への挿入は、元の並びの順に並ぶ。
 */
export function applyFixesToText(text: string, fixes: readonly Fix[]): FixesAppliedToText {
  const ordered = fixes
    .map((fix, index) => ({ fix, index }))
    .sort(
      (a, b) =>
        b.fix.range[0] - a.fix.range[0] || b.fix.range[1] - a.fix.range[1] || b.index - a.index,
    );
  const applied: Fix[] = [];
  const skipped: Fix[] = [];
  let result = text;
  /** ここより後ろはすでに書き換えた。 */
  let limit = text.length;
  for (const { fix } of ordered) {
    const [start, end] = fix.range;
    if (!(0 <= start && start <= end && end <= limit)) {
      skipped.push(fix);
      continue;
    }
    result = applyFixToText(result, fix);
    applied.push(fix);
    limit = start;
  }
  return { text: result, applied, skipped };
}

/** `applyFixesToCell` の結果。`applied` のときは、適用しなかった修正案も返す。 */
export type ApplyFixesResult =
  | Exclude<ApplyFixResult, { status: "applied" }>
  | { status: "applied"; text: string; applied: Fix[]; skipped: Fix[] };

/** 1 件の修正案をセルに適用する（#19）。`applyFixesToCell` を参照。 */
export async function applyFixToCell(
  unit: TextUnit<ExcelLocation>,
  fix: Fix,
): Promise<ApplyFixResult> {
  const result = await applyFixesToCell(unit, [fix]);
  if (result.status !== "applied") return result;
  if (result.applied.length === 0) {
    return { status: "unsupported", reason: "修正案の範囲がセルのテキストの外にあります" };
  }
  return { status: "applied", text: result.text };
}

/**
 * セルに修正案を適用する（#19）。1 つのセルに複数の修正案があるときは、後ろの位置から順に適用して 1 回で書き戻す。
 *
 * 適用の前にセルを読み直し、次のときは書き込まない。
 *
 * - 数式のセル → `unsupported`（計算結果の文字列を直しても意味がないため）
 * - 表示されている文字列が `unit.text` と違う → `stale`（もう一度 lint してもらう）
 * - 文字列でない値（数値・日付・真偽値）や、表示形式で見た目が値と違うセル → `unsupported`
 * - 書き換え後のテキストが数式や数値として解釈されてしまう → `unsupported`
 *
 * セルの値全体を置き換えるので、セル内の一部の文字だけに付けた書式（部分書式）は失われうる。
 */
export async function applyFixesToCell(
  unit: TextUnit<ExcelLocation>,
  fixes: readonly Fix[],
): Promise<ApplyFixesResult> {
  const { sheet: sheetName, address } = unit.location;
  try {
    return await Excel.run(async (context): Promise<ApplyFixesResult> => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      sheet.load("isNullObject");
      await context.sync();
      if (sheet.isNullObject) return { status: "stale", currentText: "" };

      const range = sheet.getRange(address);
      range.load(["text", "values", "formulas", "valueTypes"]);
      await context.sync();
      const text = range.text[0]?.[0] ?? "";
      const value = range.values[0]?.[0];
      const formula = range.formulas[0]?.[0];

      if (typeof formula === "string" && formula.startsWith("=")) {
        return { status: "unsupported", reason: "数式のセルは自動修正できません" };
      }
      if (text !== unit.text) return { status: "stale", currentText: text };
      if (range.valueTypes[0]?.[0] !== "String" || value !== text) {
        return {
          status: "unsupported",
          reason: "文字列でない値や、表示形式で見た目が変わるセルは自動修正できません",
        };
      }

      const result = applyFixesToText(text, fixes);
      if (result.applied.length === 0) return { status: "applied", ...result };
      if (wouldBeReinterpreted(result.text)) {
        return {
          status: "unsupported",
          reason: "修正後の文字列が数式や数値として解釈されてしまうため、自動修正できません",
        };
      }

      range.values = [[result.text]];
      await context.sync();
      return { status: "applied", ...result };
    });
  } catch (error) {
    return { status: "failed", error };
  }
}

/**
 * `Range.values` に書き込むと、文字列のままにならない（おそれがある）値かどうか。
 *
 * Excel はユーザーの入力と同じように解釈するので、`=` で始まれば数式、数値に見えれば数値になる。
 * `+` / `-` / `@` で始まるものも数式として扱われることがあり、先頭の `'` は文字列の印として消える。
 * 安全側に倒して、これらはすべて書き込まない。日付に見える文字列も変換されるが、書式が地域設定に依存するので見ない。
 */
export function wouldBeReinterpreted(text: string): boolean {
  if (/^[=+\-@']/.test(text)) return true;
  return /^\s*[\d,]*\.?\d+(?:[eE][-+]?\d+)?\s*%?\s*$/.test(text);
}
