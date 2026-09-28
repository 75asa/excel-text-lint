/**
 * collect の処理時間のベンチマーク（#17）。`npm run bench:collect` で実行する。
 *
 * Office.js を使わず、フェイクの Excel（`stubWorkbook`）で大きなシートを読む。
 * フェイクの `context.sync()` は一瞬で終わるので、ここで測れるのは JavaScript 側の処理（チャンク分割・
 * 値の走査・TextUnit の生成）と、sync の回数・読んだセルの数だけ。実機の sync の時間は含まない。
 * 結果は `src/hosts/excel/README.md` の「ベンチマーク」に書く。
 */
import { expect, it } from "vitest";
import { collectExcel } from "../src/hosts/excel/collect";
import { type FakeCellValue, stubWorkbook } from "../src/testing/excel-mock";

const TEXTS = [
  "サーバーを再起動して下さい。",
  "データを取得することができる。",
  "ユーザー登録が完了しました!",
  "１２３件のデータがあります",
  "詳細は、以下の手順を参照してください。、",
];

/** 10 列の表。1 列目は番号、2 列目は日付（数値）、残りは文字列。1 割のセルは空。 */
function cell(row: number, col: number): FakeCellValue {
  if (col === 0) return row + 1;
  if (col === 1) return 45_000 + row;
  if ((row * 7 + col) % 10 === 0) return null;
  return TEXTS[(row + col) % TEXTS.length]!;
}

const CASES = [
  { cells: 10_000, hiddenEvery: 0 },
  { cells: 50_000, hiddenEvery: 0 },
  { cells: 100_000, hiddenEvery: 0 },
  { cells: 50_000, hiddenEvery: 10 }, // 10 行に 1 行が非表示（フィルター後など）
];

it.each(CASES)("$cells セル（非表示の行: $hiddenEvery 行ごと）", async ({ cells, hiddenEvery }) => {
  const columnCount = 10;
  const rowCount = cells / columnCount;
  const hiddenRows =
    hiddenEvery > 0
      ? Array.from({ length: Math.floor(rowCount / hiddenEvery) }, (_, i) => i * hiddenEvery + 1)
      : [];
  const runs: number[] = [];
  let result: { units: number; syncs: number } | undefined;
  for (let i = 0; i < 5; i++) {
    const mock = stubWorkbook({
      sheets: [{ name: "S", cells: { rowCount, columnCount, get: cell }, hiddenRows }],
    });
    const start = performance.now();
    const units = await collectExcel("sheet");
    runs.push(performance.now() - start);
    result = { units: units.length, syncs: mock.syncCount };
  }
  runs.sort((a, b) => a - b);
  console.log(
    JSON.stringify({
      cells,
      hiddenEvery,
      ...result,
      medianMs: Math.round(runs[2]! * 10) / 10,
      minMs: Math.round(runs[0]! * 10) / 10,
    }),
  );
  expect(result?.units).toBeGreaterThan(0);
});
