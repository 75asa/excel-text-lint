/**
 * LintEngine のベンチマーク用のページ（`bench/lint-engine/run.mjs` がビルドしてヘッドレス Chrome で開く）。
 */
import { LintEngine } from "../../src/core/engine";
import type { TextUnit } from "../../src/core/types";

const TEXTS = [
  "サーバーを再起動して下さい。",
  "データを取得することができる。",
  "ユーザー登録が完了しました!",
  "１２３件のデータがあります",
  "これは問題ありません。",
  "詳細は、以下の手順を参照してください。、",
  "このAPIはJSON形式でデータを返します",
  "〜について、検討を行う事とする。",
  "エラーが発生した場合、ログを確認してください。",
  "このツールは、とても便利ですが、しかし、設定が難しいです。",
];

interface BenchArgs {
  cells: number;
  maxInFlight: number;
}

async function runBench({ cells, maxInFlight }: BenchArgs) {
  const engine = new LintEngine({
    workerUrl: "/textlint/loader.js",
    dictBaseUrl: "/dict/",
    maxInFlight,
  });
  let start = performance.now();
  await engine.init();
  const initMs = performance.now() - start;

  start = performance.now();
  await engine.lintUnit({ id: "warmup", text: TEXTS[1]!, location: null });
  const firstLintMs = performance.now() - start;

  const units: TextUnit[] = Array.from({ length: cells }, (_, i) => ({
    id: String(i),
    text: TEXTS[i % TEXTS.length]!,
    location: null,
  }));
  let firstResultMs: number | undefined;
  start = performance.now();
  const violations = await engine.lint(units, {
    onUnitResult: () => {
      firstResultMs ??= performance.now() - start;
    },
  });
  const lintMs = performance.now() - start;

  // キャンセルが効くまでの時間（lint の途中で中断して、reject されるまで）
  const controller = new AbortController();
  const running = engine.lint(units.slice(0, 2_000), { signal: controller.signal });
  await new Promise((resolve) => setTimeout(resolve, 300));
  start = performance.now();
  controller.abort(new DOMException("中止", "AbortError"));
  await running.catch(() => undefined);
  const abortMs = performance.now() - start;

  engine.dispose();
  return {
    cells,
    maxInFlight,
    initMs: Math.round(initMs),
    firstLintMs: Math.round(firstLintMs),
    lintMs: Math.round(lintMs),
    perCellMs: Math.round((lintMs / cells) * 1000) / 1000,
    firstResultMs: Math.round(firstResultMs ?? 0),
    abortMs: Math.round(abortMs * 10) / 10,
    violations: violations.length,
  };
}

Object.assign(globalThis, { runBench, benchReady: true });
