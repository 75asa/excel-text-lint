/**
 * Office.js なしでタスクペインの UI を表示するデモ（目視確認・スクリーンショット用）。
 *
 * 本番の build（vite.config.ts の input）には含めない。使い方は README.md を参照。
 * クエリで状態を切り替える:
 * - `?theme=dark|light`: テーマを固定する
 * - `?caps=none`: 今の Excel アダプタと同じ能力（移動・修正・ハイライトなし、範囲は選択範囲だけ）
 * - `?data=clean`: 違反が 0 件
 * - `?fail=1`: lint がエラーになる
 * - `?slow=1`: 辞書の読み込みを長く待たせる（進捗の表示の確認用）
 */

import type { LintOptions } from "../core/engine";
import type { ExcelLocation } from "../core/locations";
import type {
  ApplyFixResult,
  Fix,
  HostAdapter,
  HostCapabilities,
  RevealPrecision,
  TextUnit,
  Violation,
} from "../core/types";
import { applyFixToText, toViolation } from "../core/violation";
import { mountTaskpane } from "../ui/app";
import { applyTheme } from "../ui/theme";

const params = new URLSearchParams(location.search);
const theme = params.get("theme");
applyTheme(document.documentElement, theme === "dark" || theme === "light" ? theme : null);

const CELLS: [sheet: string, address: string, text: string][] = [
  ["Sheet1", "A2", "この機能を利用することができる。"],
  ["Sheet1", "B2", "見積りは、来週までに提出します。"],
  [
    "Sheet1",
    "B5",
    "サーバーの設定を行う必要があります。問題が発生した場合には、担当者に連絡を行ってください。",
  ],
  ["Sheet1", "C7", "ユーザー様の情報を適切に管理する事が重要です"],
  ["Sheet1", "D9", "合計"],
  ["月次 売上", "B3", "４月の売上は前年比１２０％でした。"],
  ["月次 売上", "C12", "詳細については、別紙を参照して下さい。"],
];

/** textlint の結果の代わり（ルールと範囲は実際の結果に近いもの）。 */
const MESSAGES: Record<
  string,
  { ruleId: string; message: string; severity: number; range: [number, number]; fix?: Fix }[]
> = {
  "Sheet1!A2": [
    {
      ruleId: "ja-technical-writing/ja-no-redundant-expression",
      message:
        '【dict5】 "することができる"は冗長な表現です。"することが"を省き簡潔な表現にすると文章が明瞭になります。',
      severity: 2,
      range: [7, 8],
      fix: { range: [7, 15], text: "できる" },
    },
  ],
  "Sheet1!B2": [
    {
      ruleId: "prh",
      message: "見積り => 見積もり",
      severity: 2,
      range: [0, 3],
      fix: { range: [0, 3], text: "見積もり" },
    },
  ],
  "Sheet1!B5": [
    {
      ruleId: "ja-technical-writing/ja-no-redundant-expression",
      message:
        '【dict3】 "設定を行う"は冗長な表現です。"設定する"など簡潔な表現にすると文章が明瞭になります。',
      severity: 2,
      range: [5, 10],
    },
    {
      ruleId: "ja-technical-writing/ja-no-redundant-expression",
      message:
        '【dict3】 "連絡を行って"は冗長な表現です。"連絡して"など簡潔な表現にすると文章が明瞭になります。',
      severity: 2,
      range: [34, 40],
    },
    {
      ruleId: "ja-technical-writing/no-doubled-joshi",
      message: '一文に二回以上利用されている助詞 "に" がみつかりました。',
      severity: 1,
      range: [33, 34],
    },
  ],
  "Sheet1!C7": [
    {
      ruleId: "ja-technical-writing/ja-no-mixed-period",
      message: '文末が"。"で終わっていません。',
      severity: 1,
      range: [21, 22],
      fix: { range: [22, 22], text: "。" },
    },
    {
      ruleId: "prh",
      message: "事 => こと",
      severity: 2,
      range: [16, 17],
      fix: { range: [16, 17], text: "こと" },
    },
  ],
  "月次 売上!B3": [
    {
      ruleId: "jtf-style/2.1.8.算用数字",
      message: "算用数字は「半角」で表記します。",
      severity: 2,
      range: [0, 1],
      fix: { range: [0, 1], text: "4" },
    },
    {
      ruleId: "jtf-style/2.1.8.算用数字",
      message: "算用数字は「半角」で表記します。",
      severity: 2,
      range: [9, 12],
      fix: { range: [9, 12], text: "120" },
    },
  ],
  "月次 売上!C12": [
    {
      ruleId: "prh",
      message: "下さい => ください",
      severity: 3,
      range: [15, 18],
      fix: { range: [15, 18], text: "ください" },
    },
  ],
};

const fullCaps: HostCapabilities = {
  scopes: ["selection", "sheet", "workbook"],
  reveal: "unit",
  highlight: true,
  applyFix: "unit",
  selectionTracking: false,
};
const currentExcelCaps: HostCapabilities = {
  scopes: ["selection"],
  reveal: "none",
  highlight: false,
  applyFix: "none",
  selectionTracking: false,
};

class DemoAdapter implements HostAdapter<ExcelLocation> {
  readonly host = "excel";
  readonly capabilities = params.get("caps") === "none" ? currentExcelCaps : fullCaps;
  readonly #texts = new Map(CELLS.map(([sheet, address, text]) => [`${sheet}!${address}`, text]));

  async collect(): Promise<TextUnit<ExcelLocation>[]> {
    await delay(150);
    return CELLS.map(([sheet, address], index) => ({
      id: `${sheet}!${address}`,
      text: this.#texts.get(`${sheet}!${address}`) ?? "",
      // #17 で ExcelLocation に isFormula が加わる。デモでは B3 を数式のセルにしておく
      location: {
        host: "excel",
        sheet,
        address,
        row: index,
        col: 0,
        ...{ isFormula: address === "B3" },
      },
    }));
  }

  async reveal(unit: TextUnit<ExcelLocation>): Promise<RevealPrecision> {
    console.info("reveal", unit.id);
    return unit.location.sheet === "月次 売上" && unit.location.address === "C12" ? "none" : "unit";
  }

  async highlight(violations: readonly Violation[]): Promise<void> {
    console.info("highlight", violations.length);
  }

  async clearHighlight(): Promise<void> {
    console.info("clearHighlight");
  }

  async applyFix(unit: TextUnit<ExcelLocation>, fix: Fix): Promise<ApplyFixResult> {
    await delay(100);
    const current = this.#texts.get(unit.id);
    if (current !== unit.text) return { status: "stale", currentText: current ?? "" };
    const text = applyFixToText(current, fix);
    if (/^[=+\-@']/.test(text)) return { status: "unsupported", reason: "数式として解釈される値" };
    this.#texts.set(unit.id, text);
    return { status: "applied", text };
  }
}

class DemoEngine {
  #warm = false;

  async init(): Promise<void> {
    await delay(400);
  }

  async lint(units: readonly TextUnit[], options: LintOptions = {}): Promise<Violation[]> {
    if (params.get("fail")) {
      await delay(300);
      throw new Error("textlint の Worker でエラーが発生しました（デモ）");
    }
    if (!this.#warm) await delay(params.get("slow") ? 60_000 : 1200, options.signal);
    this.#warm = true;
    const clean = params.get("data") === "clean";
    const result: Violation[] = [];
    let done = 0;
    for (const unit of units) {
      await delay(120, options.signal);
      const found = clean
        ? []
        : (MESSAGES[unit.id] ?? []).map((message) =>
            toViolation(unit.id, unit.text, {
              ...message,
              index: message.range[0],
              line: 1,
              column: message.range[0] + 1,
            }),
          );
      result.push(...found);
      done += 1;
      options.onProgress?.({ done, total: units.length, violations: result.length });
    }
    return result;
  }
}

function delay(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    signal?.throwIfAborted();
    const timer = setTimeout(resolve, ms);
    signal?.addEventListener(
      "abort",
      () => {
        clearTimeout(timer);
        reject(signal.reason);
      },
      { once: true },
    );
  });
}

const app = mountTaskpane({
  root: document.getElementById("app") as HTMLElement,
  adapter: new DemoAdapter(),
  engine: new DemoEngine(),
  title: "excel-text-lint",
});
Object.assign(window, { demoApp: app });
