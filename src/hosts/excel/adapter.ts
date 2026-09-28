import type { ExcelLocation } from "../../core/locations";
import type {
  ApplyFixResult,
  CollectOptions,
  Fix,
  HostAdapter,
  HostCapabilities,
  RevealPrecision,
  Scope,
  TextRange,
  TextUnit,
  Violation,
} from "../../core/types";
import { applyFixToCell } from "./fix";
import { clearHighlights, highlightCells } from "./highlight";
import { revealCell } from "./reveal";
import { columnName } from "./selection";

/**
 * Excel のアダプタ。
 *
 * - collect: 今は "selection" だけ（シート・ブック全体は #13）
 * - reveal（該当セルへの移動）: `reveal.ts`（#15）
 * - highlight / clearHighlight: `highlight.ts`（#16）
 * - applyFix: `fix.ts`（#19）
 */
export class ExcelAdapter implements HostAdapter<ExcelLocation> {
  readonly host = "excel";
  readonly capabilities: HostCapabilities = {
    scopes: ["selection"],
    reveal: "unit",
    highlight: true,
    applyFix: "unit",
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

  /** セルの中の文字位置までは選択できないので、`range` は使わない（精度は最良でも `unit`）。 */
  reveal(unit: TextUnit<ExcelLocation>, _range?: TextRange): Promise<RevealPrecision> {
    return revealCell(unit);
  }

  async highlight(
    violations: readonly Violation[],
    units: ReadonlyMap<string, TextUnit<ExcelLocation>>,
  ): Promise<void> {
    await highlightCells(violations, units);
  }

  async clearHighlight(): Promise<void> {
    await clearHighlights();
  }

  applyFix(unit: TextUnit<ExcelLocation>, fix: Fix): Promise<ApplyFixResult> {
    return applyFixToCell(unit, fix);
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
