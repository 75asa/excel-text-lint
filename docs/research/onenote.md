# 調査: OneNote アダプタで使える API と制約

- Issue: #24（最初のステップ「API の制約調査」）
- 関連: #9（`TextUnit` / `HostAdapter` のインターフェース）、[ADR 0001](../adr/0001-office-integration.md)
- 調査日: 2026-09-28
- 方法: Microsoft Learn のリファレンスと GitHub の Issue を読んで整理した。実機ではまだ確かめていない。実機で確かめる項目は「実機で確認すること」にまとめた

## 結論

- **読み取り（collect）は、アクティブなページに限ればできる。** 本文の段落・入れ子の段落・表のセル・ページタイトルを、段落単位のプレーンテキストで取れる。リストかどうかも `getParagraphInfo()` でわかる
- **違反箇所へのジャンプ（reveal）は、ページ単位までしかできない。** 段落や文字範囲を選択・スクロールする API がない。そのため、タスクペイン側の工夫で補う（後述）
- **ハイライト（highlight）はできない。** 文字書式を変える API がない。段落を HTML で作り直せば色は付けられるが、破壊的な変更になるので採用しない
- **自動修正（applyFix）は、文字範囲の置換ができない。** `RichText.text` は読み取り専用。段落を丸ごと作り直す方法と、ユーザーの選択範囲を置き換える方法しかない。MVP では修正案を表示してコピーできるようにするだけにする
- **走査の範囲は「現在のページ」が基本になる。** ほかのページの内容は、そのページを開かない限り読めない（`Accessing page contents of inactive page is not allowed`）。セクション全体を走査するには、ページを 1 枚ずつ開いていく必要がある
- Microsoft Graph を併用すれば、ほかのページも読める。ただし OneNote は NAA（nested app authentication）の対象外で認証が重く、書き込みも段落単位になる。そのため初期実装では使わない

## HostAdapter の操作ごとの可否

凡例: ✅ できる / △ 制約付きでできる / ❌ できない

| 操作 | 可否 | 使う API | 制約・備考 |
|---|---|---|---|
| collect: 本文の段落 | ✅ | `Page.contents` → `PageContent.outline` → `Outline.paragraphs` → `Paragraph.richText.text` | `PageContent.type` は `Outline` / `Image` / `Other`。`Paragraph.type` は `RichText` / `Image` / `Table` / `Ink` / `Other` で、`RichText` だけを対象にする |
| collect: 入れ子の段落（インデント） | ✅ | `Paragraph.paragraphs` | 階層ごとに `load` と `sync` が要る。深さぶん往復が増える |
| collect: 表のセル | ✅ | `Paragraph.table` → `rows` → `cells` → `TableCell.paragraphs` | セルの中も段落の集まりなので、同じ処理を再帰で使える。`rowIndex` と `cellIndex` で位置を表せる |
| collect: リスト | ✅ | `Paragraph.getParagraphInfo()` | `listType`（`None` / `Number` / `Bullet`）と `index` が取れる。`indentationLevel` は API set 1.8 と書かれている（後述） |
| collect: 見出し | △ | `RichText.getHtml()`、`RichText.style` | 見出しかどうかを直接表すプロパティは 1.1 にない。`style`（`ParagraphStyle`）は API set 1.8 と書かれている。`getHtml()` の出力で見分けられるかは実機で確かめる |
| collect: ページタイトル | ✅ | `Page.title`、`Outline.isTitle()` | タイトルは `Page.title` で読める（書き込みもできる） |
| collect: 手書き（Ink）・画像 | ❌（対象外） | `Page.inkAnalysisOrNull` | 手書きの認識結果は取れる可能性があるが、lint の対象にしない |
| collect: 選択範囲 | △ | `Office.context.document.getSelectedDataAsync(CoercionType.Text)` | テキストは取れるが、ページのどこにあるかはわからない |
| reveal: ページを開く | ✅ | `Application.navigateToPage(page)`、`navigateToPageWithClientUrl(url)` | ページ単位まで |
| reveal: 段落・文字範囲の選択 | ❌ | なし | `Paragraph` にも `RichText` にも `select()` がない。代わりの UX は後述 |
| reveal の代わり: カーソル位置の追跡 | ✅ | `Office.EventType.DocumentSelectionChanged`、`Application.getActiveParagraphOrNull()` | ユーザーがカーソルを置いた段落がわかるので、タスクペインでその段落の違反だけを表示できる |
| highlight: 文字範囲の色付け | ❌ | なし | 文字書式（背景色など）を変える API がない |
| highlight: 段落を HTML で作り直して色付け | △（非推奨） | `RichText.getHtml()` → `Paragraph.insertHtmlAsSibling()` → `Paragraph.delete()` | 元に戻すにはもう一度作り直す必要がある。書式が落ちる・段落 ID が変わる・共同編集と競合する。採用しない |
| highlight: 表のセルの背景色 | △ | `TableCell.shadingColor`（読み書きできる） | 表のセルだけ。元の色を覚えておけば戻せる |
| highlight: ノートタグを付ける | △（非推奨） | `Paragraph.addNoteTag()` | 付けられるが、外す API がない（`NoteTag` に削除メソッドがない）。元に戻せないので使わない |
| clearHighlight | —（何もしない） | — | highlight をしないので不要 |
| applyFix: 文字範囲の置換 | ❌ | なし | `RichText.text` は読み取り専用で、範囲を指定して置き換える API もない |
| applyFix: 段落の作り直し | △（非推奨） | `getHtml()` → HTML 内の文字列を置換 → `insertHtmlAsSibling("Before", html)` → `delete()` | 書き込める HTML はサブセットで、空白がまとめられる。リンクやタグ、書式が落ちる可能性がある |
| applyFix: 選択範囲の置換 | △ | `getSelectedDataAsync` → 照合 → `setSelectedDataAsync(text, { coercionType: Text })` | ユーザーが違反箇所を選択している場合だけ使える。選択テキストが違反文字列と一致するか確かめてから置き換える |
| 範囲: 現在のページ | ✅ | `Application.getActivePage()` | — |
| 範囲: セクション | △ | `Section.pages` + ページごとに `navigateToPage` | メタデータ（`title`、`id`、`clientUrl`）は全ページ分取れる。内容はページを開かないと読めないので、画面が次々に切り替わる |
| 範囲: ノートブック | △ | `Notebook.sections`、`Notebook.sectionGroups` + 上と同じ | OneNote on the web では同時に開けるノートブックは 1 つだけ（`Application.notebooks` の説明による） |
| 範囲: セクション・ノートブック（Graph 経由） | △ | `GET /me/onenote/pages/{id}/content` | 画面を切り替えずに読める。ただし認証が要る（後述） |

### 実装イメージ（collect）

```ts
await OneNote.run(async (context) => {
  const page = context.application.getActivePage();
  page.load("id,title,clientUrl");
  const contents = page.contents;
  contents.load("items/type,items/id,items/outline/id");
  await context.sync();

  for (const pc of contents.items.filter((c) => c.type === "Outline")) {
    const paras = pc.outline.paragraphs;
    paras.load("items/id,items/type,items/richText/text,items/richText/languageId");
    // Table の場合は items/table/rows/items/cells/items/paragraphs を、
    // 子段落は items/paragraphs を、それぞれ次の sync で再帰的に load する
  }
  await context.sync();
});
```

`load` のパスでまとめて指定すれば、往復は「段落の深さ + 表の入れ子の深さ」回ほどで済む見込み。大きなページでの往復回数と時間は実機で測る。

## 違反箇所へジャンプできない場合の代わりの UX

1. **ページを開く。** 別のページの結果を選んだときは `navigateToPage` でそのページを開く
2. **文脈を見せる。** タスクペインに、違反を含む段落のテキストを違反部分を強調して表示する。段落の種類（タイトル・本文・リストの何番目・表の何行何列）も添える
3. **検索語をコピーする。** 違反箇所を含む短い文字列をクリップボードにコピーするボタンを置く。ユーザーは OneNote の検索（Ctrl+F）でその場所に移動できる
4. **カーソルに追従する。** `DocumentSelectionChanged` で `getActiveParagraphOrNull()` を読み、ユーザーがカーソルを置いた段落の違反を一覧の先頭に出す。一覧から探すのではなく、本文を読みながら違反を確認できる
5. **選択して直す（将来）。** ユーザーが違反箇所を選択すると、タスクペインが選択テキストを照合し、「この選択を修正案で置き換える」ボタンを有効にする（`setSelectedDataAsync`）

## Microsoft Graph の OneNote API を併用する案

| 項目 | 内容 |
|---|---|
| 読み取り | `GET /me/onenote/pages/{id}/content?includeIDs=true` でページの HTML を取れる。ページを開かなくてよい |
| ページ一覧 | `GET /me/onenote/sections/{id}/pages`（既定 20 件、`top` は最大 100 件。`@odata.nextLink` でページング） |
| Office.js との対応 | `Page.getRestApiId()`、`Section.getRestApiId()` で Graph の ID が取れる。SharePoint 上のノートブックは、ノートブックの baseUrl を使う必要がある（リファレンスのサンプルのコメントによる） |
| 書き込み | `PATCH /me/onenote/pages/{id}/content` に JSON の変更リストを送る。`p` / `li` / `h1`〜`h6` は生成 ID を指定して **要素ごと replace** できる |
| 書き込みの制約 | `span`、`a`、`tr`、`td` は対象にできない。表は表全体を replace するしかない。生成 ID は更新のたびに変わりうる |
| 権限 | `Notes.Read`（読み取り）、`Notes.ReadWrite`（書き込み）。アプリのみの認証は 2025-03-31 に廃止され、委任アクセス許可だけが使える |
| 認証 | NAA の信頼済みブローカー（`brk-multihub`）は Word / Excel / PowerPoint / Outlook / Teams で、**OneNote は含まれていない**。Office ダイアログ API を使った MSAL のポップアップ認証などが必要になる |

**判断:** 初期実装では使わない。

- 認証まわり（Entra ID のアプリ登録、ダイアログでのサインイン）の実装と運用の負担が大きい
- 書き込みは段落単位の replace にとどまり、Office.js で段落を作り直す方法と比べて利点が小さい
- 編集中のページでは、サーバー側の内容と画面の内容がずれる可能性がある（要確認）

「セクション・ノートブック全体を画面を切り替えずに lint したい」という要望が出た段階で、改めて検討する。

## sideload（OneNote on the web）

- manifest は add-in only manifest（XML）だけが使える。unified manifest には対応していない（ADR 0001 のとおり）
- manifest の `<Hosts>` には `<Host Name="Notebook"/>` を書く。`<Requirements>` には `<Set Name="OneNoteApi" MinVersion="1.1"/>` を書く
- 手動で sideload する手順:
  1. OneNote on the web でノートブックを開く
  2. **ホーム** > **アドイン** > **その他の設定** を選ぶ
  3. **Office アドイン** ダイアログで **マイ アドインのアップロード** を選ぶ
  4. manifest を選んでアップロードする
- Yeoman で作ったプロジェクトなら、`npm run start -- web --document {共有リンクの URL}` で自動的に sideload できる（Excel / OneNote / PowerPoint / Word が対象）
- sideload した manifest はブラウザの local storage に保存される。キャッシュを消したり、別のブラウザを使ったりしたときは、もう一度 sideload する

## Web Worker と IndexedDB（辞書のキャッシュ）

- OneNote on the web のアドインは iframe の中で動く（公式ドキュメントに「In OneNote on the web, the web application displays in a webview control or iframe.」とある）。アドインのページは自分のオリジンで読み込まれる
- Excel on the web と同じ仕組みなので、Web Worker（同じオリジンのスクリプト）と IndexedDB は **使える見込み**。ただし OneNote に固有の資料は見つからなかったので、実機で確認する
- 注意点: Chromium 115 以降はストレージのパーティション分割が有効。ストレージは「トップレベルのサイト + アドインのオリジン」ごとに分かれる。そのため、**Excel on the web でキャッシュした辞書（約 15 MB）は OneNote では使えず、OneNote で初めて開いたときにもう一度ダウンロードする**。初回ロードの進捗表示（#18）は OneNote でも必要になる
- ブラウザの設定でサードパーティのストレージがブロックされている場合、IndexedDB が使えない可能性がある。その場合は毎回ダウンロードする形に切り替える（キャッシュなしでも動くようにしておく）

## 既知の制限と不具合

- **公開されている requirement set は OneNoteApi 1.1 だけ。** 一方で、リファレンスには 1.2 / 1.8 / 1.9 と書かれたメンバーがある（`Section.isEncrypted` は 1.2、`RichText.style` と `ParagraphInfo.indentationLevel` は 1.8、`Application.getSelectedInkStrokes` は 1.9）。requirement set の一覧には 1.1 しか載っていないので、これらに頼るときは `isSetSupported("OneNoteApi", "1.x")` で確かめ、使えない前提で作る
- **ほかのページの内容は読めない。** `Accessing page contents of inactive page is not allowed` というエラーになる（OfficeDev/office-js#1373、2020 年から open のまま）。公式ドキュメントにも「Page Content にアクセスできるのはアクティブなページだけ」とある。ただしタイトルなどのメタデータは読める
- **Common API はごく一部だけ使える。** `getSelectedDataAsync`（`Text` と `Matrix`）、`setSelectedDataAsync`（`Text`、`Image`、`Html`）、`DocumentSelectionChanged` イベント、それに設定（content add-in のみ）
- **HTML を挿入すると空白がまとめられ、1 つの outline に貼り付けられる。** そのため、段落を作り直して修正する方法は元の見た目を保てない
- **OneNote on the web で content add-in のサイズを変えるとエラーダイアログが出る**（OfficeDev/office-js#6775、2026-06 報告）。タスクペインのアドインなので直接は関係しないが、OneNote のアドイン基盤はまだ不安定なところがある
- **NAA が OneNote に対応していない。** Graph を使う場合の認証が重くなる（前述）
- **型定義の注意。** `RichText.text` と `Paragraph.richText` は読み取り専用。`Paragraph.set()` はあるが、更新できるプロパティは実質的にない

## #9（HostAdapter のインターフェース）への提言

OneNote はできることが最も少ないホストなので、共通インターフェースは「できない操作がある」ことを前提に作る必要がある。

1. **ホストごとの能力（capabilities）を宣言する。**
   ```ts
   interface HostCapabilities {
     scopes: Scope[];                                    // OneNote: ["selection", "page"]（"section" は任意）
     reveal: "range" | "unit" | "container" | "none";    // OneNote: "container"（ページ単位）
     highlight: boolean;                                 // OneNote: false
     applyFix: "range" | "unit" | "selection" | "none";  // OneNote: "none"（将来は "selection"）
     selectionTracking: boolean;                         // OneNote: true
   }
   ```
   UI は capabilities を見て、ボタンの有無や説明文を切り替える。
2. **reveal はどこまでできたかを返す。** `reveal(unit, range): Promise<"exact" | "unit" | "container" | "none">` にする。UI は結果に応じて、文脈表示や検索語のコピーを出す。
3. **highlight / clearHighlight は省略できるようにする。** 何もしない実装を許し、capabilities で無効だと伝える。
4. **applyFix は失敗しうる前提にする。**
   - 戻り値で結果（適用した・対象が変わっていた・非対応）を返す
   - 適用前に対象の現在のテキストを読み直し、lint したときと一致するか確かめる（stale チェック）。この仕組みは全ホスト共通で持つ
5. **選択の変化を知らせる任意のイベントを追加する。** `onSelectionChanged?(cb: (unitId: string | null) => void): Disposable`。OneNote では reveal の代わりになる。Word と Excel でも「カーソル位置の違反を先頭に出す」機能に使える。
6. **collect は進捗とキャンセルに対応させる。**
   - OneNote でセクションを走査する場合は、ページを開くたびに `sync` が走り、時間がかかる
   - `collect(scope, { signal, onProgress })` にするか、AsyncIterable で unit を順に返す形にする
7. **OneNote の location の型（案）**
   ```ts
   type OneNoteLocation = {
     host: "onenote";
     pageId: string;
     pageClientUrl: string;      // navigateToPageWithClientUrl 用
     pageTitle: string;
     kind: "title" | "paragraph" | "listItem" | "tableCell";
     outlineId?: string;
     paragraphId?: string;
     path: number[];             // outline / 段落 / 子段落の添字の並び（ID が変わったときの予備）
     table?: { row: number; cell: number };
     list?: { type: "Number" | "Bullet"; index: number };
   };
   ```
   段落の ID は編集で変わる可能性がある。ID が見つからないときは path とテキストの一致で探し直す。

## 推奨する実装方針

**フェーズ 1（#24 の残りの範囲、読み取り専用）**

- 対象範囲は「現在のページ」と「選択範囲」
- collect: ページタイトル、本文の段落（入れ子を含む）、表のセル、リストの情報を集める。`RichText` 以外の段落は飛ばす
- reveal: 別のページなら `navigateToPage` で開き、あとは文脈表示・検索語のコピー・カーソル追従で補う
- highlight / applyFix: 実装しない。修正案は表示し、コピーできるようにする
- lint は PoC（#2）の結論どおり、段落 1 つを 1 メッセージとして Worker に送る

**フェーズ 2（任意）**

- 選択範囲を置き換える修正（`setSelectedDataAsync`）。選択テキストと違反文字列が一致するときだけ有効にする
- 表のセルに限って `shadingColor` でハイライトする（元の色を覚えておき、clearHighlight で戻す）

**見送り**

- 段落を HTML で作り直すハイライトと修正（破壊的で、元に戻せない）
- Graph の併用（認証の負担が大きい。要望が出たら再検討する）
- セクションやノートブック全体の走査（ページを次々に開く UX が悪い。要望が出たら、確認ダイアログ付きで検討する）

## 実機で確認すること

- [ ] OneNote on the web のタスクペインで Web Worker が起動し、IndexedDB に辞書をキャッシュできるか（Chrome / Edge / Safari）
- [ ] `RichText.getHtml()` の出力で見出し（h1〜h6 相当のスタイル）を見分けられるか
- [ ] `RichText.style` と `ParagraphInfo.indentationLevel`（API set 1.8）が実際に使えるか。`isSetSupported` の結果も見る
- [ ] 100 段落・表を含むページで、collect の往復回数と所要時間
- [ ] `DocumentSelectionChanged` と `getActiveParagraphOrNull()` で、カーソルのある段落を安定して取れるか
- [ ] `getSelectedDataAsync` / `setSelectedDataAsync`（Text）で、選択範囲を修正案に置き換えたとき書式が保たれるか。元に戻す（Ctrl+Z）が効くか
- [ ] 段落を編集したあとも、`Paragraph.id` が変わらないか

## 出典

- [OneNote JavaScript API programming overview](https://learn.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)（2025-09-24 更新）
- [Work with OneNote page content](https://learn.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins-page-content)（アクティブなページしか読めないこと、対応する HTML）
- [OneNote JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets)
- API リファレンス: [Application](https://learn.microsoft.com/en-us/javascript/api/onenote/onenote.application) / [Page](https://learn.microsoft.com/en-us/javascript/api/onenote/onenote.page) / [Section](https://learn.microsoft.com/en-us/javascript/api/onenote/onenote.section) / [Outline](https://learn.microsoft.com/en-us/javascript/api/onenote/onenote.outline) / [Paragraph](https://learn.microsoft.com/en-us/javascript/api/onenote/onenote.paragraph) / [RichText](https://learn.microsoft.com/en-us/javascript/api/onenote/onenote.richtext) / [TableCell](https://learn.microsoft.com/en-us/javascript/api/onenote/onenote.tablecell) / [ParagraphInfo](https://learn.microsoft.com/en-us/javascript/api/onenote/onenote.paragraphinfo) / [NoteTag](https://learn.microsoft.com/en-us/javascript/api/onenote/onenote.notetag)
- [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- [Persist add-in state and settings](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings)（ブラウザのストレージとパーティション分割）
- [Enable SSO with nested app authentication](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in)（`brk-multihub` の対象ホスト）
- [Get OneNote content and structure (Microsoft Graph)](https://learn.microsoft.com/en-us/graph/onenote-get-content)
- [Update OneNote page content (Microsoft Graph)](https://learn.microsoft.com/en-us/graph/onenote-update-page)
- [OneNote API overview (Microsoft Graph)](https://learn.microsoft.com/en-us/graph/integrate-with-onenote)（アプリのみの認証の廃止）
- [OfficeDev/office-js#1373: OneNote: Fetching contents of non-active page](https://github.com/OfficeDev/office-js/issues/1373)
- [OfficeDev/office-js#6775: OneNote on the web shows host error dialog when resizing a minimal content add-in](https://github.com/OfficeDev/office-js/issues/6775)
