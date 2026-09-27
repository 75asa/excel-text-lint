# ADR 0001: Office との連携方式

- Status: Proposed
- Date: 2026-09-28
- Issue: #1

## 背景

[textlint](https://textlint.github.io/) を Office のドキュメントに対して実行し、違反箇所を具体的に（どのセル・段落・スライドの、どの文字列が、どのルールに違反しているか）確認できるようにしたい。

対象は Excel から始め、最終的には Word / PowerPoint / OneNote / Loop に広げる。

textlint は JavaScript で書かれている。そのため、Office からどうやって JavaScript の実行環境につなぐかが最初の設計判断になる。

## 選択肢

### A. Office アドイン（Office.js / タスクペイン）

Web 技術（HTML/JS）で作るアドイン。Office がタスクペインに Web ページを読み込み、Office.js 経由でドキュメントを読み書きする。

- textlint は [`@textlint/script-compiler`](https://github.com/textlint/editor/tree/master/packages/@textlint/script-compiler) で「textlint 本体 + ルール + 設定」を 1 つの Web Worker スクリプトにまとめられる。Node もシェル呼び出しも要らない（[textlint editor](https://github.com/textlint/editor) で実績あり）
- 対応ホストと環境（[Office Add-ins availability](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-in-availability)、2026-08 時点）:

  | ホスト | Windows | Mac | Web | iPad |
  |---|---|---|---|---|
  | Excel | ✅ | ✅ | ✅ | ✅ |
  | Word | ✅ | ✅ | ✅ | ✅ |
  | PowerPoint | ✅ | ✅ | ✅ | △（一部） |
  | OneNote | ❌ | ❌ | ✅ | ❌ |
  | Loop | ❌（アドイン非対応） | | | |

- 同じコード（lint エンジン・UI）をホスト間で共有できる。違いは「テキストの集め方」と「違反箇所への移動・ハイライト」だけで、この部分をアダプタに閉じ込められる
- 配布: sideload（開発時）、Microsoft 365 管理センターからの組織展開、AppSource
- 制約:
  - HTTPS で配信するホスティングが必要
  - 大きなドキュメントでは `context.sync()` の往復がボトルネックになる
  - 日本語ルールで使う形態素解析の辞書（kuromoji、数 MB）の読み込みコストがかかる

### B. VBA マクロ + Node の textlint CLI

VBA から `WScript.Shell` で `npx textlint --format json` を起動し、結果をシートに書き出す。

- 利点: 既存の textlint 設定（`.textlintrc`、Node 用ルール）がそのまま使える。すぐ作れる
- 欠点:
  - Windows デスクトップ版でしか現実的に動かない。Mac はサンドボックスのため `AppleScriptTask` 経由になり、Web 版では VBA 自体が動かない
  - 利用者全員に Node と textlint のインストールが必要
  - インターネットから入手したファイルのマクロは既定でブロックされるため、配布が難しい
  - Word / PowerPoint は別の VBA を用意することになり、OneNote / Loop には対応できない

### C. Office Scripts（Excel の TypeScript オートメーション）

- npm パッケージを import できず、textlint と kuromoji を載せるのは現実的でない
- Excel に閉じた仕組みで、ほかのホストに広げられない

## 決定（案）

**A. Office アドインを採用する。**

- 最終目標（Excel → Word / PowerPoint → OneNote）を 1 つの仕組みでカバーできるのは A だけである
- 「違反箇所を具体的に見る」という要件は、タスクペインの一覧と、そこから該当セル・段落へ移動する UI で自然に満たせる
- B は icebox の Issue（#28）として残す。Windows 限定で、手元の `.textlintrc` をそのまま使いたい場合の補助手段とする

### 付随する方針

- **manifest**: add-in only manifest（XML）で始める。
  - [unified manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview) は Excel / Word / PowerPoint では Web、Windows（M365 の 2501 以降）、Mac（16.103 以降）で使える
  - 一方で OneNote は非対応で、買い切り版の Office on Windows でも使えない
  - 対応ホストが揃った時点で移行を検討する（#4 で最終判断）
- **Loop**: アドインでは対応できない。2026 年時点で Loop のページ内容を読み書きする API はない。
  - 候補は次の3つ:
    - Graph で `.loop` ファイルを HTML に変換して読み取る
    - Adaptive Card ベースの Loop コンポーネント
    - テキストを貼り付けて lint する Web UI
  - どれにするかは #25 で調査する
- **Excel の補助機能（将来）**: Excel だけの機能として、カスタム関数（例: `=TEXTLINT(A1)` で違反件数を返す）も作れる。MVP の後に検討する

## 影響

- M0 の技術検証は「ブラウザ（Worker）で textlint と日本語プリセットが動くか」（#2）が最大のリスクになる。ここで辞書サイズや読み込み時間が許容できないと判明した場合は、次のどちらかに切り替える。その際はこの ADR を改訂する
  - ルールを絞る
  - サーバー側で lint する構成（その場合はプライバシー面の再検討が必要）
- 開発スタックは npm エコシステム（Office のツール群、textlint）に合わせる。Fresh テンプレートは撤去する（#3）
