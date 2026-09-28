# hosts/excel

Excel 用のアダプタ。Excel JavaScript API（`Excel.run`）でセルのテキストを集め、違反箇所へ移動・ハイライトする処理を置く。

ホストに依存しない処理（lint の実行、結果の整形など）は `src/core/` に置き、ここからは呼び出すだけにする。

- `adapter.ts`: `HostAdapter<ExcelLocation>` の実装（reveal は #15、highlight は #16、applyFix は #19）
- `collect.ts`: セルのテキストを集める（#13、#17）。方針は下の「collect」
- `selection.ts`: 選択範囲のセルのテキストを読む最初の実装（雛形のもの）と、列名の変換（`columnName`）

## collect（#13、#17）

```ts
const adapter = new ExcelAdapter({ includeHidden: false, formulas: "include" });
const units = await adapter.collect("workbook", { signal, onProgress });
```

### 範囲（scope）

`capabilities.scopes` は `["selection", "sheet", "workbook"]`。

| scope | 読む範囲 |
|---|---|
| `selection` | 選択範囲（`getSelectedRanges()`。Ctrl で選んだ複数の領域も可）と、そのシートの使用範囲の重なり。領域が重なったセルは 1 回だけ返す |
| `sheet` | アクティブシートの使用範囲 |
| `workbook` | すべてのシートの使用範囲（シートの並び順） |

- 使用範囲は `getUsedRangeOrNullObject(true)`（値のあるセルだけ。書式だけのセルは含めない）。列全体や行全体を選択しても、使用範囲の分しか読まない
- 並びは、シートの順、その中では行優先（1 行が `chunkCells` を超える幅のシートだけは、チャンクごとの行優先）

### 対象にするセル

| セル | 扱い | 理由 |
|---|---|---|
| 文字列（`valueTypes` が `String`） | 対象 | |
| 数値・日付・時刻・真偽値・エラー | 除く | 日付や時刻も Excel では数値。表示形式を適用した文字列は lint しても意味がない |
| 空・空白だけ | 除く | |
| 数式で、結果が文字列 | **既定で対象**（`location.isFormula: true`）。`formulas: "exclude"` で除く | 読み手が目にするのは計算結果の文字列で、`="…"&A1` のように数式の中に文章を書くこともあるため。ただし参照先のセルと同じ違反が重ねて出ることがある |
| 結合セル | 左上のセルだけ対象 | Excel では左上以外のセルの値は空なので、特別な処理なしで左上だけになる |
| セル内改行 | そのまま 1 つの TextUnit | 違反の位置は `line` / `column` で表せる |
| 非表示の行・列・シート | **既定で除く**。`includeHidden: true` で含める | 読み手に見えない。reveal（#15）で選択しても画面に出せない。フィルターで隠れた行も非表示として扱う |

- テキストは `values`（セルの値）を使う。`text`（表示形式を適用した文字列）は使わない。applyFix（#19）で書き込むのは値で、stale の判定も値で行うため。表示形式で文字を足している（`@"様"` など）ときは、その部分は lint しない
- 数式かどうかは `formulas` で判定する。定数のセルでは `formulas` は値と同じになるので、「`=` で始まり、値と違う」ものを数式とする（`'=abc` と入力した文字列の定数を数式と間違えないため）
- **applyFix（#19）は `location.isFormula` が true のセルでは `unsupported` を返すこと**。値を書き込むと数式が消えるため
- 非表示のシートは `visibility` が `Visible` 以外（`Hidden` / `VeryHidden`）のもの。`selection` / `sheet` はアクティブシートなので常に表示されている

### TextUnit の id と location

- id は `<シート ID>!<番地>`（例: `{00000000-0001-0000-0000-000000000000}!B3`。`excelUnitId`）。シート ID（`Worksheet.id`）はシート名を変えても変わらないので、collect をやり直しても同じセルは同じ id になる
- location（`ExcelLocation`）は `sheetId`、`sheet`（名前、表示用）、`address`（`B3`）、`row` / `col`（0 始まり）、`isFormula`

### 大きなブック（#17）

- 範囲を「チャンク」に分けて読む。1 チャンクは `chunkCells`（既定 5,000）セル以下の長方形で、基本は行で分ける（10 列なら 500 行ずつ）
- チャンクごとに `values` / `valueTypes` / `formulas` / `rowHidden` / `columnHidden` を 1 回の `context.sync()` でまとめて読む。小さなシートや選択範囲の領域は、合計が `chunkCells` 以下になるまで 1 回の sync にまとめる
- 非表示の判定: `rowHidden` / `columnHidden` は「すべて非表示 → true、すべて表示 → false、まざる → null」を返す。まざっているチャンクだけ、もう 1 回の sync で `getSpecialCells("Visible")` の領域を読む。使えない環境（ExcelApi 1.9 未満など）では、行・列ごとに `rowHidden` / `columnHidden` を読む
- sync の回数: `selection` / `sheet` は 1 + チャンクの数、`workbook` は 2 + チャンクの数（非表示の行・列がまざるチャンクは +1）
- 応答が大きすぎて失敗したとき（`…PayloadSizeLimitExceeded`。Excel on the web の上限）は、チャンクを半分にして読み直す
- `signal` は sync の合間で確かめ、中断されたら次の sync をせずに `signal.reason` で reject する
- `onProgress` はチャンクを読むたびに `{ collected }`、最後に `{ collected, total }` を渡す（セル数は走査し終えるまで分からないので、途中の `total` はない）
- `collectExcel` の `onChunk` で、チャンクごとに見つかった TextUnit を受け取れる（下の「lint を collect と重ねるか」）

使っている API と要求セット: `getSelectedRanges`・`getSpecialCells`（ExcelApi 1.9）、`getRangeByIndexes`（1.7）、`getUsedRangeOrNullObject`（1.4）。manifest の `ExcelApi 1.1` は、実機での確認のあとで 1.9 に上げるのがよい。

### ベンチマーク

2026-09-28、Apple Silicon の Mac で計測。

**collect（フェイクの Excel）**: `npm run bench:collect`（`bench/collect.bench.test.ts`）

10 列の表（1 列目は番号、2 列目は日付、残りは文字列、1 割は空）をアクティブシートとして collect した。フェイクの `context.sync()` は一瞬で終わるので、JavaScript 側の処理時間と sync の回数だけを測っている。

| セル | TextUnit | sync | 時間（中央値） |
|---|---|---|---|
| 10,000 | 7,200 | 3 | 5 ms |
| 50,000 | 36,000 | 11 | 49 ms |
| 100,000 | 72,000 | 21 | 199 ms |
| 50,000（10 行に 1 行が非表示） | 32,500 | 21 | 38 ms |

**lint（LintEngine + 本物の Worker、ヘッドレス Chrome）**: `node bench/lint-engine/run.mjs`（事前に `npm run build:worker` と `npm i --no-save --ignore-scripts playwright-core`）

PoC と同じ短い日本語の文（違反のある文とない文）を 1 セル 1 件で lint した。

| セル | maxInFlight | lint の時間 | 1 セルあたり | 最初の結果まで |
|---|---|---|---|---|
| 10,000 | 1 | 9.7 秒 | 0.97 ms | 6 ms |
| 10,000 | 8（既定） | 8.5 秒 | 0.85 ms | 12 ms |
| 10,000 | 32 | 9.0 秒 | 0.90 ms | 12 ms |
| 50,000 | 1 | 44.7 秒 | 0.89 ms | 6 ms |
| 50,000 | 8（既定） | 40.1 秒 | 0.80 ms | 10 ms |
| 50,000 | 32 | 40.3 秒 | 0.81 ms | 6 ms |

- Worker の起動は約 80 ms、最初の lint（辞書の読み込みを含む）は約 1 秒（辞書が IndexedDB にキャッシュされたあとは約 0.2 秒）
- lint はセル数にほぼ線形で、1 セル約 0.8 ms。`maxInFlight` を 8 より増やしても速くならない。中断は `signal` の abort から reject まで 1 ms 未満（Worker に送った最大 8 件の処理が残るだけ）
- **LintEngine の設定は変えなくてよい**（`maxInFlight` は既定の 8 のまま）

**合わせた見積もり**（実機の sync の時間は未計測。1 回 50〜200 ms と仮定）

| セル（うち文字列 7 割） | collect | lint | 合計 |
|---|---|---|---|
| 10,000 | 約 0.2〜0.6 秒（sync 3 回） | 約 6 秒（7,200 件） | 約 7 秒 |
| 50,000 | 約 0.6〜2.2 秒（sync 11 回） | 約 29 秒（36,000 件） | 約 30 秒 |

時間の大半は lint で、collect は数 % にとどまる。

### lint を collect と重ねるか

`collectExcel` はチャンクごとに `onChunk` を呼ぶので、collect の完了を待たずに lint を始めることはできる。ただし、上の見積もりのとおり collect は全体の数 % なので、今は `HostAdapter.collect`（すべて集めてから返す）のままにしている。重ねるなら次のようにする（core の変更が要る）。

- `CollectOptions` に `onChunk?(units)` を足し、アダプタから渡す
- `LintEngine.lint` が `AsyncIterable<TextUnit[]>` も受け取れるようにし、進捗の `total` は collect が終わるまで未定にする
