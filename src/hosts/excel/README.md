# hosts/excel

Excel 用のアダプタ。Excel JavaScript API（`Excel.run`）でセルのテキストを集め、違反箇所へ移動・ハイライトする処理を置く。

ホストに依存しない処理（lint の実行、結果の整形など）は `src/core/` に置き、ここからは呼び出すだけにする。

| ファイル | 内容 |
|---|---|
| `adapter.ts` | `HostAdapter<ExcelLocation>` の実装。各操作は下のファイルの関数を呼ぶだけ |
| `selection.ts` | 選択範囲のセルのテキストを読む最初の実装（雛形のもの） |
| `reveal.ts` | 違反のセルへ移動する（#15） |
| `highlight.ts` | 違反のセルのハイライトと解除（#16） |
| `fix.ts` | 修正案の適用（#19） |

capabilities は `reveal: "unit"`・`highlight: true`・`applyFix: "unit"`。

テストは `src/testing/excel-mock.ts` の `stubExcelWorkbook`（シート・セル・条件付き書式を持つブックのフェイク）で書く。

## reveal（#15）

`worksheet.activate()` と `range.select()`（ExcelApi 1.1）でセルを選択する。

- Excel の API ではセルの中の文字位置までは選択できない（セルの編集モードに入る API もない）。精度は最良でも `unit` で、`range` 引数は使わない。文字位置は UI 側で文脈表示などで補う
- 次のときは何も選択せずに `none` を返す
  - シートが削除された・名前が変わった（`getItemOrNullObject` が null）
  - セルの表示テキストが `TextUnit.text` と違う。行や列の挿入・削除でセルがずれたとき、別のセルを選んでしまわないようにするため
  - API がエラーを返した（セルの編集中など）

## highlight / clearHighlight（#16）

### 方式の比較

| 方式 | 元の書式を壊さずに解除できるか | タスクペインを閉じたあとの解除 | 見た目・そのほか |
|---|---|---|---|
| 塗りつぶし（`format.fill.color`） | △ 元の色を覚えて戻す必要がある。`fill.color` はテーマの色・網掛け・グラデーションを表せず、完全には戻せない（`getCellProperties`（1.9）で読んでも、テーマの色や tint を含めて戻すのは複雑）。ハイライト中にユーザーが色を変えると、解除で上書きしてしまう | 覚えた元の色をどこかに保存する必要がある | 分かりやすい。ExcelApi 1.1 |
| **条件付き書式（採用）** | ✅ セルの書式そのものには触れない。付けた条件付き書式を消すだけで元どおり | ✅ 印の入った数式で見分けられるので、保存が要らない | 塗りつぶしと同じ見た目。条件付き書式の塗りはセルの塗りより優先して表示される。ExcelApi 1.6 |
| ノート / コメント | ✅ 別のオブジェクトなので書式は変えない | 付けたものを覚えておく必要がある | 違反の内容を書けるが、既存のノート・コメントとぶつかる。コメントは共同編集者に通知が飛ぶ。赤い三角だけで目立たない。ノートは ExcelApi 1.18、コメントは 1.10 |
| 部分書式（セル内の一部の文字だけ色を変える） | — | — | **Excel JavaScript API ではできない**（下記） |

条件付き書式を選んだ。「元の書式を壊さずに解除できる」を満たす方式のうち、見た目が塗りつぶしと同じで目立ち、解除のための状態を持たなくてよいため。

### 実装

- 違反のあるセルごとに、Custom 型の条件付き書式を 1 つ足す（同じセルに違反が複数あっても 1 つ）。数式は `=ISTEXT("excel-text-lint")`（文字列を与えているので常に真で、セルの値に依存しない）、塗りは `#FFEB9C`、優先順位はいちばん上（`priority = 0`）
- `highlight` は、前回のハイライトをすべて消してから付け直す
- `clearHighlight` は、ブックのすべてのシートの条件付き書式（`worksheet.getRange().conditionalFormats`）から、数式に `excel-text-lint` を含むものを消す。ユーザーの条件付き書式やセルの書式には触れない

### 解除の情報の保存先

**どこにも保存しない**。条件付き書式の数式に印（`excel-text-lint`）を入れておき、解除のときはブックを走査して印のあるものを消す。

- メモリ上に持つ方式は、タスクペインを閉じる・再読み込みする・ブックを開き直すと解除できなくなる
- ブックの設定（`Office.context.document.settings` やカスタム XML）に ID を保存する方式は、保存と実際の状態がずれうる（ユーザーがセルをコピーして条件付き書式が増えた、ハイライト中にブックを保存せずに閉じた、など）
- 印で見分ける方式なら、コピーで増えたものも含めて消せ、タスクペインを閉じたあとでも、開き直して `clearHighlight` を呼べば消せる

残る課題（UI 側、#18 で検討）:

- ハイライトしたままブックを保存すると、条件付き書式がファイルに残り、ほかの人にも見える。タスクペインを閉じる前に解除するよう促す、または起動時に残っているハイライトを消すボタンを出す
- 違反のセルが非常に多いと、条件付き書式も同じ数だけ増える。多い場合は件数に上限を設けるか、`getRanges`（ExcelApi 1.9）で複数のセルを 1 つの条件付き書式にまとめることを検討する
- ExcelApi 1.6 未満の Excel（Excel 2016 の買い切り版など）では条件付き書式の API がない

### セル内の一部の文字だけを色付けできるか

**できない**（2026 年 9 月時点）。

- Excel JavaScript API の書式（`Range.format.font`、`getCellProperties` / `setCellProperties` の `format.font`）はセル単位で、セル内の文字範囲を指定する方法がない。VBA の `Range.Characters(start, length).Font` に相当する API はない
- `CellPropertiesFont` も 1 セルに 1 つの値で、リッチテキスト（文字ごとの書式の列）は読み書きできない。Microsoft 365 Developer Platform にリッチテキスト対応の要望は出ているが、まだ提供されていない
- そのため、違反の文字範囲は UI（タスクペイン）の文脈表示で示す

## applyFix（#19）

`fix.ts` の `applyFixToCell`（1 件）と `applyFixesToCell`（1 つのセルの複数件）。

1. セルの `text` / `values` / `formulas` / `valueTypes` を読み直す
2. 数式のセル（`formulas` が `=` で始まる）は `unsupported`
3. 表示テキストが `TextUnit.text` と違えば `stale`（シートが削除されていても `stale`）
4. 文字列でない値（数値・日付・真偽値）や、表示形式で見た目が値と違うセル（`@"（注）"` など）は `unsupported`。表示テキストを直して書き戻すと、値が変わってしまうため
5. 修正案を適用する。複数のときは後ろの位置から順に適用し（前の修正案の位置がずれないように）、すでに適用したものと範囲が重なる修正案は飛ばす（`skipped`）
6. 書き換え後の文字列が数式や数値として解釈されてしまう（`=`・`+`・`-`・`@`・`'` で始まる、数値に見える）なら `unsupported`
7. `range.values` に書き込む（1 回の `sync`）

セルの値全体を書き換えるので、セル内の部分書式（手で付けた一部の文字の太字や色など）は失われうる（`ApplyFixMode` が `unit` なのはこのため）。

### Undo の方針

- Excel は 2025 年 9 月に、Office.js のアドインの操作を Ctrl+Z で元に戻せるようにした（[Undo capabilities with the Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities)）。取り消しの単位は `context.sync()` 1 回。ここで使う API（`values` の書き込み・条件付き書式の追加と削除・`activate` / `select`）は、元に戻せない API の一覧に入っていない
  - `Excel.run` に `mergeUndoGroup: true`（ExcelApi 1.20）を渡すと、複数の sync を 1 回の Undo にまとめられる
  - 古いビルドでは、アドインが書き込むと Undo の履歴がすべて消える
- applyFix は書き込みを 1 回の `sync` で行うので、新しい Excel では Ctrl+Z 1 回で修正前に戻る。**アドイン独自の取り消し機能は作らない**
  - 古い Excel で Undo が効かないことに備え、UI（#18）は修正前後のテキストを表示し、「すべて適用」の前に確認を出す
- highlight は「解除の sync」と「追加の sync」に分かれるので、Ctrl+Z で解除とハイライトが別々に戻る。ハイライトは `clearHighlight` で消すのが正しい手順とする（`mergeUndoGroup` でまとめるかは実機で確かめてから）
