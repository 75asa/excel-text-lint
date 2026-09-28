# hosts/excel

Excel 用のアダプタ。Excel JavaScript API（`Excel.run`）でセルのテキストを集め、違反箇所へ移動・ハイライトする処理を置く。

ホストに依存しない処理（lint の実行、結果の整形など）は `src/core/` に置き、ここからは呼び出すだけにする。

- `adapter.ts`: `HostAdapter<ExcelLocation>` の実装。今は collect("selection") だけ（reveal は #15、highlight は #16、applyFix は #19、シート・ブックの collect は #13）
- `selection.ts`: 選択範囲のセルのテキストを読む最初の実装（雛形のもの）
