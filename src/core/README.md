# core

ホスト（Excel / Word / PowerPoint / OneNote）に依存しない層。

今後ここに置くもの:

- textlint の Worker とのやり取り（#8）
- lint 結果の型と整形

Office.js（`Office` / `Excel` などのグローバル）には依存しないこと。ホストごとの処理は `src/hosts/<host>/` に置く。
