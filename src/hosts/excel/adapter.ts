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
import { collectExcel, EXCEL_SCOPES, type ExcelCollectSettings, isExcelScope } from "./collect";

/**
 * Excel のアダプタ。
 *
 * collect（選択範囲・アクティブシート・ブック全体）は `collect.ts`（#13、#17）。
 * - reveal（該当セルへの移動）: #15
 * - highlight / clearHighlight: #16
 * - applyFix: #19
 */
export class ExcelAdapter implements HostAdapter<ExcelLocation> {
  readonly host = "excel";
  readonly capabilities: HostCapabilities = {
    scopes: EXCEL_SCOPES,
    reveal: "none",
    highlight: false,
    applyFix: "none",
    selectionTracking: false,
  };

  /**
   * @param collectSettings 非表示のセルや数式のセルを対象にするかなど（`ExcelCollectSettings`）。
   *   collect のたびに読むので、UI から書き換えてよい。
   */
  constructor(readonly collectSettings: ExcelCollectSettings = {}) {}

  async collect(scope: Scope, options: CollectOptions = {}): Promise<TextUnit<ExcelLocation>[]> {
    if (!isExcelScope(scope)) throw new Error(`Excel アダプタは範囲「${scope}」に対応していません`);
    return collectExcel(scope, { ...this.collectSettings, ...options });
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
