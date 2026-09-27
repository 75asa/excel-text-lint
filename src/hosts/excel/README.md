# hosts/excel

Excel 用のアダプタ。Excel JavaScript API（`Excel.run`）でセルのテキストを集め、違反箇所へ移動・ハイライトする処理を置く。

ホストに依存しない処理（lint の実行、結果の整形など）は `src/core/` に置き、ここからは呼び出すだけにする。
