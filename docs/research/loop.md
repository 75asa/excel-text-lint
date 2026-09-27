# 調査: Loop で textlint を使う方法

- Issue: #25
- 調査日: 2026-09-28
- 関連: [ADR 0001](../adr/0001-office-integration.md)、[PoC: textlint をブラウザで動かす](../../poc/textlint-worker/README.md)

## 結論（先に要約）

- **Loop のページを直接読み書きできる公式の手段は、2026-09 時点でもない。** Loop は Office アドインに対応しておらず、Loop 専用の API もない（[Q&A 2026-02](https://learn.microsoft.com/en-us/answers/questions/5760312/can-ms-provide-a-microsoft-loop-api-for-programmat)）。
- 読み取りだけなら Graph の `?format=html` 変換が使える。ただし、Loop アプリのワークスペースのページは SharePoint Embedded（SPE）に保存される。SPE にサードパーティのアプリがアクセスするには **app-only 権限、管理者による PowerShell の登録、サーバー側のシークレット** が要る。「ユーザーのブラウザの中だけで lint する」というこのプロジェクトの前提とは相性が悪い。
- **書き戻しはどの公式手段でもできない。**
- 推奨は段階的に進める案:
  1. **貼り付け Web UI**（全アカウントで確実に動き、コストも最小）
  2. **ブラウザ拡張の PoC**（Loop の中で違反を見られる唯一の方式。先に textlint editor 拡張で手早く検証する）
  3. 必要があれば、**Graph の委任権限で .loop を読む機能**（OneDrive / SharePoint にある Loop コンポーネントに限る）

  Adaptive Card、Copilot、Fluid Framework は採用しない。

## Loop のデータがどこにあるか（前提知識）

[Overview of Loop storage](https://learn.microsoft.com/en-us/microsoft-365/loop/loop-storage)（2026-05 更新）より:

| 作成した場所 | 保存先 |
|---|---|
| Loop アプリ（My workspace・共有ワークスペース） | SharePoint Embedded コンテナー |
| Teams のチャットノート、チャネルのワークスペース | SharePoint Embedded コンテナー |
| Teams のプライベートチャット・会議 | 作成者の OneDrive |
| Teams のチャネル | SharePoint サイト |
| Outlook、OneNote、Whiteboard | 作成者の OneDrive |

- ファイルは `.loop`（古いものは `.fluid`）。中身は Fluid Framework のスナップショットで、形式は公開されていない
- My workspace は Copilot Pages / Copilot Notebooks と同じユーザー所有の SPE コンテナーにある
- 個人の Microsoft アカウント（MSA）でも Loop アプリは使える（無償）。ただし Outlook などへのコンポーネント埋め込みはできない（[Q&A](https://learn.microsoft.com/en-us/answers/questions/5333208/does-loop-work-with-a-personal-account-in-microsof)、[Loop access via Microsoft 365 subscriptions](https://support.microsoft.com/en-us/office/loop-access-via-microsoft-365-subscriptions-92915461-4b14-49a4-9cd4-d1c259292afa)）。MSA のワークスペースが Graph からどう見えるか（`/me/drive` に出るか）は公式の記述がなく、**未確認**

## 比較表

| 観点 | A. Graph `?format=html` | B. Adaptive Card Loop コンポーネント | C. 貼り付け Web UI | D. ブラウザ拡張 | E. Copilot / エージェント | F. Fluid Framework |
|---|---|---|---|---|---|---|
| 実現性 | △ 読み取りのみ。SPE は管理者の設定が必須 | × Loop ページの本文は読めない | ◎ | ○ 要 PoC（DOM 次第） | × lint に向かない | × 非公開・非サポート |
| 権限 | OneDrive / SP: 委任 `Files.Read`（＋共有リンクなら `Files.Read.All`）。SPE: app-only ＋ `FileStorageContainer.Selected` ＋ ゲストアプリ登録（管理者） | Entra アプリ ＋ Bot 登録、組織へのアプリ展開（管理者） | 不要 | 拡張のインストール（組織では管理者のポリシー次第） | Copilot ライセンス ＋ `Files.Read.All` / `Sites.Read.All` など | — |
| Loop 内で違反が見えるか | × 別画面に一覧 | × 別カードに表示 | × 別画面 | ◎ 下線・ポップアップを出せる | × | — |
| 違反箇所へのジャンプ | × | × | ×（Web UI 内のみ） | ○ | × | — |
| 書き戻し | × | × | △ 修正後のテキストをコピー | △ DOM への入力で可能だが壊れやすい | × | — |
| 個人アカウント | △ 未確認 | × Teams / Outlook の組織アカウントが前提 | ◎ | ◎ | × | — |
| 主なリスク | テナント全体の Loop を読める強い権限、サーバー必須、HTML の構造が非公開 | Mac とモバイルで非対応、Loop ページ上では静的なサムネイルになる | 手間がかかる | Loop の DOM 変更で壊れる、Fluid との整合 | LLM がテキストを改変する、コスト | 規約・互換性 |
| 実装コスト | 中〜大 | 大 | 小 | 中 | 大 | 特大 |

## 各方式の詳細

### A. Microsoft Graph で .loop を HTML に変換して読む

- **API**: `GET /drives/{drive-id}/items/{item-id}/content?format=html`。`302` で事前認証付きの URL（数分で失効）が返る（[Convert to other formats](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format)、v1.0、2026-06 更新）
  - `html` に対応する拡張子は `loop, fluid, wbtx, whiteboard`
  - 取れるのは現行版だけで、バージョン指定の HTML 化はできない（[Q&A](https://learn.microsoft.com/en-us/answers/questions/2105327/get-specific-loop-file-version-exported-and-conver)）
- **権限**:
  - OneDrive / SharePoint にある `.loop`（Teams のプライベートチャット、Outlook などで作ったコンポーネント）: 委任 `Files.Read` で読める。個人アカウントも委任 `Files.Read` が表の上では対象
  - 他人が作ったコンポーネントを共有リンクから開くには `/shares/{encoded-url}/driveItem` で解決する。この場合は `Files.Read.All` が必要で、テナントによっては管理者の同意が要る
  - **SPE にある Loop ページ（Loop アプリのワークスペース）**: `FileStorageContainer.Selected` に加えて、Loop のコンテナー種別（`a187e399-0c36-4b98-8f04-1edc167a0996`）へのゲストアプリ登録が必要。これは SharePoint Embedded 管理者が `Set-SPOApplicationPermission` で行う
    - **ゲストアプリは app-only 権限のみ対応で、委任権限は非対応**（[Set-SPOApplicationPermission](https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/set-spoapplicationpermission)）
    - つまり「ログインしたユーザーの権限で自分のページだけ読む」ことができず、テナント中の Loop を読めるサーバーアプリになる
    - Microsoft 自身も、この経路を eDiscovery・エクスポート・移行ツール向けとして案内している（[Loop compliance summary](https://learn.microsoft.com/en-us/microsoft-365/loop/loop-compliance-summary)）
    - 実例として、エクスポートツールが app-only の `Files.Read.All`、`Sites.Read.All`、`FileStorageContainer.Selected` と Windows の PowerShell で登録している（[wals.pro の技術ガイド](https://wals.pro/en-us/blogs/news/microsoft-loop-to-html-export-technical-guide)、非公式）
- **UX**:
  - ユーザーは Loop のリンクを貼るかファイルを選び、別画面（Web UI やタスクペイン）で違反一覧を見る
  - HTML と Loop 上の位置を対応づける公式の ID はないので、Loop 内へのジャンプはできない。表示できるのは「どのブロックの、どの文字列か」まで
- **書き戻し**: できない。`.loop` は Fluid のスナップショットで、HTML からの逆変換 API はない
- **制約とリスク**:
  - SPE のために app-only の広い権限を持つサーバーを置くのは、ADR 0001 の「ブラウザ内で完結する」方針とプライバシーの前提に反する
  - HTML の構造は仕様として公開されておらず、変わりうる
  - Loop の URL（`loop.cloud.microsoft/p/...`）から driveItem を引く正式な方法が文書化されていない（**要検証**）
- **実装コスト**:
  - OneDrive / SharePoint に限定して委任権限・SPA（MSAL.js）で作るなら中程度
  - SPE まで対応する場合はサーバー、管理者向けの手順書、Windows の PowerShell が要り、大きい

### B. Adaptive Card ベースの Loop コンポーネント（メッセージ拡張）

- **仕組み**: Teams のメッセージ拡張に、リンク展開と Universal Actions（`refresh`）を実装する。manifest 1.13 以降で Microsoft 365 に拡張すると、Teams と Outlook でカードが Loop コンポーネントのように動く（[Loop Component in Adaptive Cards](https://learn.microsoft.com/en-us/microsoftteams/platform/m365-apps/cards-loop-component)、2026-07 更新）
- **実現性**: 「Loop の文章を lint する」目的には合わない
  - カードは自分のアプリのデータを表示・更新するもので、**周囲の Loop ページや他のコンポーネントの本文は読めない**
  - サードパーティのカードは Loop アプリのページに貼っても **静的なサムネイルになり、操作できない**。`.loop` ファイルも生成されない（[Q&A 2025-09](https://learn.microsoft.com/en-us/answers/questions/5551968/embedding-third-party-adaptive-card-loop-component)）
  - **Teams と Outlook の Mac 版・モバイル版では使えない**（上記 Learn の Note）
- **権限**: Entra アプリの登録、Azure Bot（Microsoft 365 チャネル）、SSO。組織への配布には Teams 管理者の承認が要る。個人アカウントは対象外
- **UX**: できるのは「テキストを入力するダイアログ → 結果をカードで返す」程度で、C の Web UI を Teams の中に置いただけになる。Adaptive Card では JS が動かないため、lint はダイアログの Web ページかボットのサーバーで行うことになる
- **書き戻し**: できない（カード自身の内容だけ）
- **実装コスト**: 大（ボットのホスティング、manifest、SSO、審査）

### C. テキストを貼り付けて lint する Web UI

- **仕組み**:
  - Loop でテキストを選んでコピーし、Web ページに貼り付けて lint する
  - クリップボードの `text/html` を読めば、見出しやリストなどの構造を保ったままブロック単位に分けられる。PoC の結論（セル単位で投げる）に合わせ、ブロック単位で Worker に投げる
- **権限**: 不要。サインインも要らない。テキストはブラウザの外に出ない
- **UX**:
  - 違反一覧と、原文上のハイライトは Web UI 内で見られる
  - Loop 側でのジャンプはできない。代わりに違反の前後の文字列を表示し、Loop 側で検索（Ctrl+F）しやすくする
- **書き戻し**: `fix` を適用したテキストをコピーし、ユーザーが貼り直す。Loop の書式は失われうるので、ブロック単位でコピーできるようにする
- **個人・組織**: 違いなし。Loop 以外（OneNote のデスクトップ版、メールの下書きなど）にも使える
- **制約とリスク**: 毎回のコピーと貼り付けの手間。大きなページでは貼り付けの範囲をユーザーが選ぶ
- **実装コスト**: 小
  - lint エンジン（#8）と結果表示の UI（#18）は Office アドインと共有できる
  - アドインのタスクペインと同じ静的ホスティング（#7）で、別のエントリーポイントとして配信できる

### D. ブラウザ拡張（content script）で loop.cloud.microsoft を lint する

- **仕組み**:
  - [textlint editor](https://github.com/textlint/editor) のブラウザ拡張は `textarea` と `contenteditable`（Google Docs なども）に対応している
  - `@textlint/script-compiler` で作った `textlint-worker.js` を読み込んで使う
  - このプロジェクトの PoC は同じ形式の Worker を作るので、**既存の拡張に自前の Worker を読ませるだけで、Loop 上で動くかをすぐ試せる**
- **権限**:
  - Graph の権限は不要
  - 組織の端末では、Edge / Chrome の拡張許可ポリシーで管理者の許可が要る場合がある
  - 個人は自由に入れられる
- **UX**: Loop の編集画面の上で下線や修正候補を出せる。違反箇所がそのまま見えるので、ジャンプの問題もない
- **書き戻し**: 技術的には入力イベントを模擬して置換できる。ただし Loop の編集は Fluid で同期されるため、DOM を直接書き換えると不整合や巻き戻りが起きうる。最初は表示だけにし、修正は「候補をコピーする」にとどめるのが安全
- **制約とリスク**:
  - Loop の DOM は非公開で、予告なく変わる
  - デスクトップの Loop アプリ（PWA）、モバイル、Teams / Outlook に埋め込まれたコンポーネントでは動かない（Teams / Outlook の Web 版なら動く可能性はある）
  - Loop には Microsoft Editor の校正が組み込まれており、その表示と重なる
- **実装コスト**: textlint editor の既存拡張で検証するだけなら小。自前の拡張を配布・保守するなら中

### E. Copilot / Microsoft 365 エージェント経由

- 宣言型エージェントには SharePoint / OneDrive を知識として与えられる。Copilot Notebooks は `.loop` を参照資料にできる（[Compare Loop, Copilot Pages, Notebooks](https://support.microsoft.com/en-us/microsoft-365-copilot/compare-microsoft-loop-copilot-pages-and-copilot-notebooks)）
- ただし、LLM を通すと本文がそのまま届く保証がない。textlint のように決定的なルールで全文を検査する用途には向かない
- [Copilot Retrieval API](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/retrieval/overview) もクエリに関連する断片を返すもので、全文は取れない
  - `.loop` は lexical 検索のみ対応
  - Copilot ライセンスか従量課金が要る
  - `Files.Read.All` と `Sites.Read.All` が必要
- エージェントの API プラグインから lint API を呼ぶ構成も考えられるが、lint をサーバー側で行うことになる
- **結論**: 採用しない（コスト、ライセンス、決定性、プライバシーの問題）

### F. Fluid Framework

- Loop は Fluid Framework 上に作られ、Fluid のクライアントライブラリは OSS として公開されている
- ただし Microsoft はサードパーティ向けの拡張として「自前のアプリで Fluid と Azure Fluid Relay を使う」ことを案内しており、Loop の文書に参加する方法は案内していない（[Stay in sync with Microsoft Loop](https://devblogs.microsoft.com/microsoft365dev/stay-in-sync-with-microsoft-loop/)、2021 年。その後も更新なし）
- Loop のドキュメントスキーマと、SharePoint へのドライバの使い方は非公開。Loop に接続するのは非サポートで、利用規約と互換性の面でリスクが大きい
- **結論**: 採用しない

## Loop コンポーネントを Word / Outlook / Teams に埋め込んだ場合

- **Word**:
  - Word for the web でのコンポーネントの挿入と表示は **2025-09-01 に廃止** された。既存のものは、Loop を開くリンク付きの読み取り専用プレースホルダーになる。デスクトップ版はもともとプレースホルダー表示（[Loop requirements](https://learn.microsoft.com/en-us/microsoft-365/loop/loop-requirements)、MC1107493）
  - Office.js（Word API）に Loop コンポーネントの中身を取る API はない。Word アドインから読めるのは、せいぜいリンクまで
  - リンクを取り出して A の Graph 経路で読むことはできる。ただしリンク先が SPE なら A と同じ制約を受ける
- **OneNote**: Web 版でのコンポーネントは 2026-09 中旬から順次廃止され、リンクになる（MC1454385）。Windows のデスクトップ版では残るが、OneNote のアドインは Web 版のみ（ADR 0001 の表）なので、アドインからは読めない
- **Outlook**: Outlook アドインで取れるメール本文にも、Loop コンポーネントの中身は含まれない前提で考えるべき（公式の記述なし、**未確認**）。実体は送信者の OneDrive にある `.loop` なので、読むなら A の委任権限の経路になる
- **Teams**: Teams のメッセージ拡張・タブから、チャットにある他の Loop コンポーネントの中身を読む API はない。実体はプライベートチャットなら OneDrive、チャットノートなら SPE

→ **「Word アドインで Loop も読む」は成り立たない。** Loop コンポーネントを読むなら Graph（A）かブラウザ拡張（D）になる。

## 推奨する方式（段階的な案）

1. **Step 1: 貼り付け Web UI（C）を M4 の最小対応とする**
   - lint エンジンと結果 UI を Office アドインと共有し、別のエントリーポイントとして配信する
   - 権限もアカウントの種類も問わず、すべてのユーザーが使える
   - `text/html` から構造を保ってブロックに分け、fix を適用したテキストをコピーできるようにする
2. **Step 2: ブラウザ拡張の PoC（D）**
   - まず既存の textlint editor 拡張に、このプロジェクトの `textlint-worker.js` を読ませる
   - loop.cloud.microsoft のページ、Teams / Outlook の Web 版に埋め込まれたコンポーネントで、下線が出るか・位置が合うかを確かめる
   - 動けば、ドキュメントに「推奨の使い方」として載せるだけで済む（自前の拡張は不要）
   - 動かなければ、自前の拡張を作るかどうかを判断する
3. **Step 3（任意）: Graph の委任権限で読む機能（A）**
   - Web UI に「Loop のリンクから読み込む」を追加する。MSAL.js を使い、クライアントだけで動かす
   - 対象は OneDrive / SharePoint にある `.loop` に限る
   - SPE（Loop アプリのワークスペース）を委任権限で読めるかを検証し、読めなければ SPE は対象外と明記する
   - **app-only のサーバー構成は取らない**。ADR 0001 のプライバシー方針を変えることになるので、必要になったら ADR を改訂する
4. 採用しない: B（Adaptive Card）、E（Copilot）、F（Fluid）
   - 公式の拡張ポイントが増えたら見直す（Microsoft 365 ロードマップの Loop の項目を M4 の開始時に再確認する）

根拠:

- Loop 内で違反を見る要件を満たせる可能性があるのは D だけ
- 全ユーザーに確実に届くのは C だけ
- A は読み取り専用で、Loop 本体（SPE）に対しては管理者の設定とサーバーが必須になる。費用対効果が低い

## Issue 案（人間が判断してから作成する）

1. **Loop: テキストを貼り付けて lint する Web UI**
   - 共有の lint エンジンと結果 UI を使った単体ページ
   - クリップボードの `text/html` をブロックに分けて lint する
   - 違反一覧、原文のハイライト、fix を適用したテキストのコピーを実装する
   - 配信はアドインと同じホスティング
2. **Loop: textlint editor 拡張 + 自前 Worker で loop.cloud.microsoft を lint できるか検証する（spike）**
   - Loop のページ、Teams / Outlook の Web 版に埋め込まれたコンポーネントで、表示、位置、パフォーマンス、Microsoft Editor との共存を確認する
   - 結果を `docs/research/loop.md` に追記する
3. **Loop: 自前のブラウザ拡張を作るか判断する**（2 の結果次第）
   - Loop に特化したアダプタ（ブロックの取得、ハイライト）
   - 修正を安全に反映できるかの検証
   - 配布（Edge Add-ons / Chrome Web Store）
4. **Loop: Graph（委任権限）で .loop を HTML として読み込む（任意）**
   - MSAL.js で `Files.Read` / `Files.Read.All` を取得する
   - Loop の URL（`/shares`）から driveItem を解決し、`?format=html` で取得して lint する
   - SPE のページを委任権限で読めるかの検証を含む
5. **ADR 0001 の Loop の項を更新する**
   - この調査結果（貼り付け UI を基本に、拡張を補助とし、app-only は取らない）を反映する

## 出典

- [Convert to other formats（driveItem get content format）- Microsoft Graph v1.0](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format)
- [Overview of Loop storage](https://learn.microsoft.com/en-us/microsoft-365/loop/loop-storage)
- [Requirements for Loop components and Loop workspaces](https://learn.microsoft.com/en-us/microsoft-365/loop/loop-requirements)
- [Summary of governance, lifecycle, and compliance capabilities for Loop](https://learn.microsoft.com/en-us/microsoft-365/loop/loop-compliance-summary)
- [Set-SPOApplicationPermission](https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/set-spoapplicationpermission)
- [SharePoint Embedded authentication and authorization](https://learn.microsoft.com/en-us/sharepoint/dev/embedded/development/auth)
- [Loop Component in Adaptive Cards - Teams](https://learn.microsoft.com/en-us/microsoftteams/platform/m365-apps/cards-loop-component)
- [Q&A: Embedding third-party Adaptive Card Loop components in Microsoft Loop（2025-09）](https://learn.microsoft.com/en-us/answers/questions/5551968/embedding-third-party-adaptive-card-loop-component)
- [Q&A: Loop API for programmatic page & component retrieval（2026-02）](https://learn.microsoft.com/en-us/answers/questions/5760312/can-ms-provide-a-microsoft-loop-api-for-programmat)
- [Q&A: Get specific Loop file version exported and converted in html format](https://learn.microsoft.com/en-us/answers/questions/2105327/get-specific-loop-file-version-exported-and-conver)
- [Q&A: Does Loop work with a personal account](https://learn.microsoft.com/en-us/answers/questions/5333208/does-loop-work-with-a-personal-account-in-microsof)
- [Loop access via Microsoft 365 subscriptions](https://support.microsoft.com/en-us/office/loop-access-via-microsoft-365-subscriptions-92915461-4b14-49a4-9cd4-d1c259292afa)
- [Microsoft Loop storage limits](https://support.microsoft.com/en-us/loop/microsoft-loop-storage-limits)
- [Microsoft 365 Copilot Retrieval API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/retrieval/overview)
- [Compare Microsoft Loop, Copilot Pages, and Copilot Notebooks](https://support.microsoft.com/en-us/microsoft-365-copilot/compare-microsoft-loop-copilot-pages-and-copilot-notebooks)
- [Stay in sync with Microsoft Loop - Microsoft 365 Developer Blog（2021-12）](https://devblogs.microsoft.com/microsoft365dev/stay-in-sync-with-microsoft-loop/)
- [textlint/editor](https://github.com/textlint/editor)
- 非公式: [Microsoft Loop to HTML Export — Technical Guide（wals.pro）](https://wals.pro/en-us/blogs/news/microsoft-loop-to-html-export-technical-guide)
- 非公式: [Retirement of Loop Component Rendering in Word for the Web（MC1107493）](https://m365admin.handsontek.net/retirement-loop-component-rendering-word-web/)

## 未確認事項（Issue 2・4 で検証する）

- 個人アカウントの Loop ワークスペースが Graph（`/me/drive` など）から見えるか
- SPE 上の Loop ページを、委任権限（`FileStorageContainer.Selected` の委任）で本人が読めるか。ゲストアプリは委任非対応と明記されているので、読めない見込み
- `loop.cloud.microsoft/p/...` の URL から driveItem を引く方法
- Loop の編集領域が `contenteditable` で、textlint editor 拡張が位置を正しく取れるか
