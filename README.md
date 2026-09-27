# excel-text-lint

[textlint](https://textlint.github.io/) で Office のドキュメントの文章をチェックする Office アドイン（Office.js・タスクペイン）です。まず Excel に対応し、その後 Word / PowerPoint / OneNote に広げます（方針は [ADR 0001](docs/adr/0001-office-integration.md)）。

> 現在は雛形の段階です。タスクペインには「選択範囲のセルのテキストを一覧表示する」ボタンだけがあり、textlint はまだ組み込んでいません。

## 必要なもの

- Node.js 22 以上、npm
- Excel（Mac / Windows のデスクトップ版、または Excel on the web）

## ディレクトリ構成

```
manifest.xml          add-in only manifest（XML）。開発用に https://localhost:3000 を指す
public/assets/        アイコン（仮）
src/
  taskpane/           タスクペインの HTML / TS / CSS
  core/               ホストに依存しない層（lint エンジンなど。今後追加）
  hosts/excel/        Excel アダプタ（Excel JavaScript API を使う処理）
docs/adr/             設計判断の記録
poc/                  技術検証のコード（本体とは別の npm プロジェクト）
```

## npm scripts

| コマンド | 内容 |
|---|---|
| `npm run dev` | Vite の開発サーバーを https://localhost:3000 で起動する |
| `npm start` | 開発サーバーを起動し、デスクトップ版 Excel に sideload して開く（`office-addin-debugging`） |
| `npm stop` | sideload を解除し、開発サーバーを止める |
| `npm run build` | `dist/` に本番用のファイルを出力する |
| `npm run typecheck` | TypeScript の型チェック |
| `npm run validate` | `manifest.xml` を検証する（`office-addin-manifest validate`。Microsoft のサービスに問い合わせるためネットワークが必要） |
| `npm run certs` | 開発用の HTTPS 証明書を作成・信頼登録する（`office-addin-dev-certs install`） |

## 開発の始め方

```sh
npm install
npm run certs   # 初回のみ。OS の証明書ストアに開発用 CA を登録するため、パスワードを求められる
npm run dev     # https://localhost:3000/taskpane/taskpane.html
```

- Office アドインは HTTPS で配信する必要があります。証明書は [office-addin-dev-certs](https://www.npmjs.com/package/office-addin-dev-certs) が `~/.office-addin-dev-certs` に作成し、`vite.config.ts` が読み込みます。`npm run certs` を先に実行していなくても、`npm run dev` の初回起動時に作成・登録されます
- 証明書を削除するときは `npx office-addin-dev-certs uninstall` を実行します
- ブラウザで直接タスクペインを開くと、Office の外なので「未対応のホスト」と表示されます。動作確認は Excel に sideload して行います

## sideload（Excel に読み込む）

どの方法でも、先に開発サーバー（`npm run dev`）を起動しておきます。`npm start` はサーバーの起動も行います。

### Excel on Mac

**自動:** `npm start` を実行すると、manifest を所定のフォルダにコピーして Excel を起動します。`npm stop` で解除します。

**手動:**

1. `manifest.xml` を `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` にコピーする（フォルダがなければ作る）
2. Excel を起動し（起動中なら再起動し）、ブックを開く
3. [ホーム] タブ → [アドイン]（または [挿入] → [アドイン] → [個人用アドイン]）から「excel-text-lint (dev)」を選ぶ
4. [ホーム] タブの [textlint を開く] ボタンでタスクペインが開く

参考: [Sideload Office Add-ins on Mac for testing](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)

### Excel on Windows

**自動:** `npm start` を実行すると、manifest をレジストリに登録して Excel を起動します。`npm stop` で解除します。

**手動（共有フォルダ経由）:**

1. `manifest.xml` を置いたフォルダをネットワーク共有にする（例: `\\<PC名>\addins`）
2. Excel の [ファイル] → [オプション] → [トラスト センター] → [トラスト センターの設定] → [信頼できるアドイン カタログ] で、共有フォルダのパスを追加し [メニューに表示する] にチェックを入れる
3. Excel を再起動し、[ホーム] → [アドイン] → [その他のアドイン] → [共有フォルダー] から「excel-text-lint (dev)」を追加する

参考:
- [Test and debug Office Add-ins（sideload の方法の一覧）](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins)
- [Sideload Office Add-ins for testing from a network share](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)

### Excel on the web

1. 開発サーバー（`npm run dev`）を起動しておく。ブラウザが https://localhost:3000 の証明書を信頼している必要がある（`npm run certs` 済みなら OK）
2. [Office on the web](https://www.office.com/) で Excel のブックを開く
3. [ホーム] → [アドイン] → [その他のアドイン]（Office アドインのダイアログ）を開く
4. [マイ アドイン] タブ → [個人用アドインを管理] → [マイ アドインのアップロード] を選び、`manifest.xml` をアップロードする
5. [ホーム] タブの [textlint を開く] ボタンでタスクペインが開く

参考: [Manually sideload Office Add-ins in Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)

### うまくいかないとき

- タスクペインが白いまま・読み込めない: 開発サーバーが動いているか、ブラウザで https://localhost:3000/taskpane/taskpane.html を開いて証明書の警告が出ないかを確認する
- manifest を変えたのに反映されない: Office のキャッシュを消す（[Clear the Office cache](https://learn.microsoft.com/office/dev/add-ins/testing/clear-cache)）

## 参考

- [Office Add-ins のドキュメント](https://learn.microsoft.com/office/dev/add-ins/)
- [Excel JavaScript API の概要](https://learn.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Office Add-ins with an add-in only manifest](https://learn.microsoft.com/office/dev/add-ins/develop/xml-manifest-overview)
