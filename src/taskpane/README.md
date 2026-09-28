# taskpane

タスクペインの入口。Office.js の初期化と、ホストのアダプタ・lint エンジンを用意して `src/ui/` の UI に渡すだけにしている。UI 自体はホストに依存しないので、Word / PowerPoint / OneNote でもそのまま使う（#18）。

| ファイル | 内容 |
|---|---|
| `taskpane.html` / `taskpane.ts` | 本番の入口。`Office.onReady` のあと `mountTaskpane()` を呼ぶ。Office のテーマも見る |
| `demo.html` / `demo.ts` | Office.js なしで UI だけを表示するデモ（目視確認用）。本番の build には含めない |
| `../ui/app.ts` | UI（DOM の組み立てとイベントの配線） |
| `../ui/results.ts` | 並べ替え・フィルタ・グルーピング・件数・前へ / 次へ（純関数） |
| `../ui/excerpt.ts` | 違反箇所の前後の文脈つきの抜粋、修正前後の抜粋（純関数） |
| `../ui/format.ts` | 番地（`Sheet1!B12`、`'月次 売上'!C3`）・ルール名・範囲名などの整形（純関数） |
| `../ui/actions.ts` | capabilities から出すボタンを決める、修正の一括適用の計画、結果の案内文（純関数） |
| `../ui/status.ts` | 実行の状態（収集・起動・辞書の読み込み・lint・完了・中止・エラー）の表示（純関数） |
| `../ui/theme.ts` | Office のテーマの背景色からライト / ダークを決める |
| `../ui/ui.css` | スタイル（Fluent 2 のトークンの値を CSS 変数で持つ） |

## UI フレームワークの決定（#18）

**Fluent UI は使わず、素の TypeScript + CSS で Fluent 2 の見た目に合わせて作る。**

候補を、使う予定の部品（ボタン・ドロップダウン・バッジ・進捗バー・ダイアログ・MessageBar・アコーディオン）だけを import して Vite（本番 build）でバンドルし、大きさを比べた（2026-09 時点）。

| 方式 | JS（min） | JS（gzip） | 依存 |
|---|---|---|---|
| Fluent UI React v9（9.74）+ React 18 | 444 kB | 130 kB | react / react-dom / @fluentui/react-components（Griffel など約 100 パッケージ） |
| Fluent UI Web Components v3（3.1） | 99 kB | 27 kB | @fluentui/web-components / @microsoft/fast-element / @fluentui/tokens |
| **素の TS + CSS（採用）** | タスクペイン全体で 37.5 kB（CSS は 14.2 kB） | 13.7 kB（CSS は 3.5 kB） | なし |

素の TS + CSS の数字は、lint エンジンのクライアントと Excel アダプタを含むタスクペイン全体の大きさ（UI を作り込む前は JS 8.5 kB / gzip 3.8 kB）。Fluent の 2 つは部品だけの大きさで、これにタスクペインのコードが足される。

理由:

- **バンドルサイズ**: textlint の Worker（数 MB）と辞書（約 15MB）が別にあり、初回の待ち時間はすでに長い。タスクペインの本体まで重くしたくない。React 版は約 10 倍になる
- **Office との見た目の一貫性**: 見た目の一貫性は、部品のライブラリよりもデザイントークン（色・角丸・文字サイズ・余白）で決まる。Fluent 2 の `webLightTheme` / `webDarkTheme` の値を `ui.css` の CSS 変数に写して、ボタン・MessageBar・バッジ・進捗バーなどを同じ寸法で作った。フォントは Segoe UI（なければ OS の日本語 UI フォント）
- **依存の少なさ**: 画面は 3 つ（実行 / 結果一覧 / 詳細）と小さく、状態も 1 つのクラスで持てる量。フレームワークの更新（React のメジャー更新、Web Components v3 の API の変更）に追従する手間を持たない。Office.js と同じく、テストも純関数に寄せれば DOM なしで書ける

採用しなかった場合の不利な点と対処:

- アクセシビリティは自分で担保する必要がある。ラジオグループ（範囲）・`aria-pressed`（重大度の絞り込み）・`aria-expanded`（グループ）・`role="status"` / `role="alert"`（進捗・案内）・`aria-modal` と `inert`（確認ダイアログ）を付け、キーボード（一覧で ↑↓ と Enter、詳細で Esc）でも操作できるようにした
- 部品が増えて手に負えなくなったら、Fluent UI Web Components v3 に置き換える。CSS 変数の名前を Fluent のトークンに寄せてあるので、移行しても見た目は変わらない

## 画面構成

1 つのページの中で、次の 3 つを切り替える。

1. **実行**（常に上に出す）: 範囲（`capabilities.scopes` が 2 つ以上のときだけ出す。Excel は選択範囲 / シート / ブック）、「〇〇をチェック」、中止、状態の表示と進捗バー
2. **結果一覧**: 件数のサマリ（重大度別・セル数・修正案の数）、重大度のチップ（押すと絞り込み）、まとめ方（シート別 / ルール別 / まとめない）とシート・ルールの絞り込み、一括操作（修正案をすべて適用・ハイライト・ハイライトを解除）、違反の一覧。行を押すと詳細を開く
3. **詳細**: 番地・重大度・メッセージ・ルール・行と列、前後 40 文字の文脈つきの該当箇所、修正案（修正前 → 修正後）、修正を適用・セルへ移動・該当箇所をコピー、セルの全文

下には「前へ / 次へ」のバーを固定で出す。移動すると、ホストが対応していれば該当セルを選択する（reveal）。

設定の画面は #11 で作る。今はヘッダーの「設定」ボタン（入口）だけで、押すと「準備中」と出す（`mountTaskpane` の `onOpenSettings` で差し替える）。

### capabilities による出し分け

UI は `HostAdapter` のインターフェースと `capabilities` だけを見る。アダプタの実装が進めば、UI を変えずにボタンが出る。

| capabilities | UI |
|---|---|
| `scopes` が 1 つ | 範囲の選択を出さない |
| `reveal: "none"` | 「移動」を出さない。詳細に「コピーして検索してください」と出す |
| `reveal: "container"` | ボタンの文言を「開く」にする（OneNote） |
| reveal の結果が `none`（期待は `unit` なのに） | 「セルが見つかりませんでした。もう一度チェックしてください」 |
| `applyFix: "none"` | 「修正」「すべて適用」を出さない（修正案の表示だけ） |
| `applyFix: "unit"` | 「セルの値を丸ごと置き換える」注意を詳細と確認ダイアログに出す |
| `highlight: false` | ハイライトのボタンを出さない |
| `selectionTracking: true` | ホストで選んだセルの違反を一覧で選択する |

applyFix の結果（`applied` / `stale` / `unsupported` / `failed` / `skipped`）は、それぞれ案内を出す。適用した違反は「修正済み」、同じセルのほかの違反は範囲がずれるので「要再チェック」にして、修正のボタンを消す。「すべて適用」は同じセルの修正を 1 つにまとめて（`mergeFixes`）セルごとに 1 回だけ applyFix を呼ぶ。実行前に確認ダイアログを出し、Ctrl+Z で戻せること（古い Office では戻せないこと）を書いておく。

ハイライトは書式としてファイルに残るので、自動では付けない。結果があるあいだは「ハイライトを解除」を常に出す。再実行するときと、ペインを閉じるとき（`pagehide`。完了は保証されない）には解除する。200 セルより多いときは確認してから付ける。

### 状態の表示

- 起動直後: Worker の起動を先に始め、「textlint を準備しています…」
- 初回の lint で最初のセルの結果が出るまで: 「辞書（約 15MB）を読み込んでいます…」と不確定の進捗バー（kuromoji の辞書は最初の lint で読み込まれるため）
- lint 中: 「n / N セル（違反 m 件）」と進捗バー。中止ボタンで止められる
- 完了: 違反が 0 件なら「問題は見つかりませんでした」
- エラー: 赤字で理由を出す。エンジンの起動に失敗したときは、次の実行でもう一度起動する

### レイアウトとテーマ

- 幅 320px で横スクロールが出ないようにしている（長い番地・ルール名は省略、文脈は 2 行まで）
- 一覧は 200 件ずつ描き、残りは「さらに表示」で出す。行には `content-visibility: auto` を付けている
- テーマは `prefers-color-scheme` に従う。`Office.context.officeTheme` が取れるときは、その背景色の明るさから決めて `<html data-theme>` で上書きし、`OfficeThemeChanged` にも追従する

## デモ（Office.js なしでの目視確認）

`demo.html` は本番の build（`vite.config.ts` の input）に含めていない。Vite の dev サーバーは使わず（開発用の証明書を作ってしまうため）、Vite の build API で別のディレクトリに出力して、静的に配信して開く。

```js
// 例: scratch/build-demo.mjs（configFile: false にして vite.config.ts を読まない）
import { build } from "vite";
await build({
  configFile: false,
  root: "src",
  base: "./",
  build: { outDir: "/tmp/demo-dist", rollupOptions: { input: "src/taskpane/demo.html" } },
});
```

出力を任意の静的サーバーで配信して `taskpane/demo.html` を開く。クエリで状態を切り替えられる（`?theme=dark`、`?caps=none`（今の Excel アダプタと同じ能力）、`?data=clean`（0 件）、`?fail=1`（エラー）、`?slow=1`（辞書の読み込み待ち））。
