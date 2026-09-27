# PoC: textlint をブラウザ（Web Worker）で動かす

Issue #2 の検証用コード。アプリ本体ではない。本体の構成は #3 で決める。

`@textlint/script-compiler` で textlint 本体と日本語プリセットを 1 つの Worker スクリプトにまとめる。それをヘッドレス Chrome で動かし、初期化時間・lint 時間・辞書の転送量を計測する。

## 実行

```sh
npm install
npm run build              # dist/textlint-worker.js を生成
CELLS=1000 npm run bench   # CHROME=<path> で Chrome の場所を変えられる
```

- ルール: `.textlintrc.json`（preset-ja-technical-writing と preset-ja-spacing）
- 入力: Excel のセルを想定した短い日本語の文（`public/index.html` の `TEMPLATES`）を指定の件数だけ繰り返す
- モード
  - `cdn`: 辞書をデフォルトの jsdelivr から取得する
  - `self-host`: `public/self-host-worker.js` で辞書の URL を自前の配信先に書き換える

## 結果（2026-09-28、Apple Silicon の Mac、ヘッドレス Chrome）

| 項目 | 結果 |
|---|---|
| Worker バンドル | 1.33 MB（gzip 後 357 KB） |
| kuromoji 辞書 | 約 14.7 MB（`.dat.gz` 12 ファイル。バンドルには含まれず、実行時に fetch される） |
| Worker の起動（`init` まで） | 約 80 ms |
| 最初の lint（辞書のロードを含む） | 0.9〜1.1 秒 |
| 2 回目以降の起動と最初の lint（辞書は IndexedDB にキャッシュ済み） | 0.2〜0.3 秒 |
| 1,000 セルを 1 セル 1 メッセージで lint | 約 0.8 秒 |
| 10,000 セルを 1 セル 1 メッセージで lint | 約 6.9 秒（ほぼ線形） |
| 1,000 セルを連結して 1 回で lint | 約 1.3 秒 |
| 10,000 セルを連結して 1 回で lint | **約 96 秒（超線形に遅くなる）** |
| `fix` コマンド | 動作する（`output` と `remainingMessages` が返る） |

違反 1 件あたりの結果の例:

```json
{
  "ruleId": "ja-technical-writing/ja-no-redundant-expression",
  "message": "【dict2】 \"することができる。\"は冗長な表現です。…",
  "line": 1, "column": 7, "index": 6,
  "range": [6, 7],
  "fix": { "range": [6, 15], "text": "できる。" }
}
```

## わかったこと

1. **ブラウザ上で実用的に動く。** 最大のリスクだった「kuromoji を含む日本語プリセットがブラウザで動くか」は問題なかった。ADR 0001 の方針は変えなくてよい
2. **セルは 1 つずつ lint する。** 連結して 1 回で投げる方式は、件数が増えると超線形に遅くなるため使わない。#8 のエンジンは「セル単位でメッセージを投げ、Worker の中で順に処理する」形にする。進捗表示とキャンセルもセル単位で入れられる
3. **辞書の取得先は jsdelivr にハードコードされている。**
   - バンドルに入る `process` の polyfill のせいで、`window.kuromojin.dicPath` による上書きは効かない
   - Worker の `fetch` をラップして URL を書き換えれば、自前の配信先から読める（`public/self-host-worker.js`）
   - 社内ネットワークで CDN がブロックされる場合に備え、本番は自前配信にする（#7 のホスティングで辞書も配信する）
4. **辞書は初回だけ約 15 MB をダウンロードする。** 以降は IndexedDB のキャッシュから読むので速い。初回ロード中の進捗表示が必要（#18 の UI）
5. **違反の位置情報がそのまま使える。**
   - `index`、`range`、`line`、`column` が取れる。セル内改行は `line` と `column` で表せる
   - ただし `range` は 1 文字だけのことがある。ハイライトや文脈表示では、`fix.range` があればそちらを使うか、前後の文字を足して表示する（#14、#16）
6. **ルールの選定はまだ要調整。**
   - 今回の文では preset-ja-spacing のルールに当たる違反が出なかった
   - 全角数字（`１２３`）や「下さい」は、preset-jtf-style や表記ゆれ辞書（prh）でないと検出できない
   - 同梱するルールは #10 で選び直す
