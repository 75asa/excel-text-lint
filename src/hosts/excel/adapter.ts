import type { ExcelLocation } from "../../core/locations";
import type {
  ApplyFixResult,
  CollectOptions,
  HostAdapter,
  HostCapabilities,
  RevealPrecision,
  Scope,
  TextUnit,
} from "../../core/types";
import { columnName } from "./selection";

/**
 * Excel のアダプタ。
 *
 * 今は collect("selection") だけを実装している。
 * - reveal（該当セルへの移動）: #15
 * - highlight / clearHighlight: #16
 * - applyFix: #19
 * - シート・ブック全体の collect: #13
 */
export class ExcelAdapter implements HostAdapter<ExcelLocation> {
  readonly host = "excel";
  readonly capabilities: HostCapabilities = {
    scopes: ["selection"],
    reveal: "none",
    highlight: false,
    applyFix: "none",
    selectionTracking: false,
  };

  async collect(scope: Scope, options: CollectOptions = {}): Promise<TextUnit<ExcelLocation>[]> {
    if (scope !== "selection")
      throw new Error(`Excel アダプタはまだ範囲「${scope}」に対応していません`);
    options.signal?.throwIfAborted();
    const units = await collectSelection();
    options.signal?.throwIfAborted();
    options.onProgress?.({ collected: units.length, total: units.length });
    return units;
  }

  async reveal(): Promise<RevealPrecision> {
    return "none";
  }

  async highlight(): Promise<void> {}

  async clearHighlight(): Promise<void> {}

  async applyFix(): Promise<ApplyFixResult> {
    return { status: "unsupported", reason: "Excel の自動修正は未実装です（#19）" };
  }
}

/**
 * 選択範囲のセルのうち、空でないものを行優先で返す。
 *
 * 表示形式を適用した文字列（`Range.text`）を読む。数式のセルは計算結果になる。
 * 列全体などの大きな選択に備えて、使用範囲（used range）との重なりだけを読む。
 */
async function collectSelection(): Promise<TextUnit<ExcelLocation>[]> {
  return Excel.run(async (context) => {
    const selected = context.workbook.getSelectedRange();
    const sheet = selected.worksheet;
    const range = selected.getIntersectionOrNullObject(sheet.getUsedRange(true));
    range.load(["isNullObject", "text", "rowIndex", "columnIndex"]);
    sheet.load("name");
    await context.sync();
    if (range.isNullObject) return [];

    const units: TextUnit<ExcelLocation>[] = [];
    range.text.forEach((cells, r) => {
      cells.forEach((text, c) => {
        if (text === "") return;
        const row = range.rowIndex + r;
        const col = range.columnIndex + c;
        const address = `${columnName(col)}${row + 1}`;
        units.push({
          id: `${sheet.name}!${address}`,
          text,
          location: { host: "excel", sheet: sheet.name, address, row, col },
        });
      });
    });
    return units;
  });
}
