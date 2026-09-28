# タスクペインの HTTPS ホスティング

- Issue: #7
- 日付: 2026-09-28
- 決定: **GitHub Pages**（`https://75asa.github.io/excel-text-lint/`）

## 何を配信するか

Office アドインのタスクペインは、HTTPS で配信される静的ページとして読み込まれる。本番では次のファイルを配信する。

| パス | 中身 | 大きさ |
|---|---|---|
| `taskpane/taskpane.html`、`assets/*` | タスクペイン（Vite の build 成果物）とアイコン | 数 KB |
| `textlint/loader-<hash>.js`、`textlint/textlint-worker-<hash>.js` | textlint の Worker とローダー（ファイル名にハッシュが付く。下記「Worker のキャッシュ対策」） | 約 1.3 MB |
| `dict/*.dat.gz` | kuromoji の辞書（12 ファイル） | 約 14.7 MB |
| `manifest.prod.xml` | 本番用の manifest（配布用。sideload や組織展開で使う） | 数 KB |

辞書も自前で配信する。kuromoji の辞書の取得先は jsdelivr にハードコードされていて、社内ネットワークで CDN がブロックされる場合に備えるためである（`poc/textlint-worker/README.md` の「わかったこと」3）。Worker の `fetch` をラップして、`https://cdn.jsdelivr.net/npm/kuromoji@0.1.2/dict/...` を `<配信先>/dict/...` に書き換える（`poc/textlint-worker/public/self-host-worker.js`）。

### `.dat.gz` の配信で守ること

kuromoji は `.dat.gz` を **gzip のままのバイト列** として受け取り、自前の Gunzip（zlibjs）で展開する。そのため、配信先は次の条件を満たす必要がある。

- **`Content-Encoding: gzip` を付けない。** 付けると、ブラウザが通信の段階で展開してしまう。kuromoji には展開済みのデータが渡り、Gunzip が「incorrect header check」などで失敗する
  - 「拡張子が `.gz` なら `Content-Encoding: gzip` を付ける」挙動や、「事前圧縮ファイルとして扱う」挙動を持つサーバー・CDN は避けるか、設定で止める
- `Content-Type` は `application/gzip` か `application/octet-stream` にする（バイナリとして扱われればよい）
- 配信先で再圧縮しない（gzip のファイルをさらに圧縮しても小さくならない）

デプロイのワークフロー（`.github/workflows/deploy.yml`）では、デプロイの後に `dict/base.dat.gz` を実際に取得して確認する。確認するのは、`Content-Encoding` が付いていないことと、先頭が gzip のマジックナンバー（`1f 8b`）であることの 2 点である。

## 比較

| 観点 | GitHub Pages | Cloudflare Pages | Deno Deploy | Azure Static Web Apps |
|---|---|---|---|---|
| 無料枠 | 公開リポジトリなら無料。サイト 1 GB、帯域 100 GB/月（ソフトリミット） | 帯域無制限、ビルド 500 回/月、1 ファイル 25 MiB まで | リクエスト数と転送量に月の上限あり | Free プランは帯域 100 GB/月、アプリ 250 MB まで |
| カスタムドメイン | ○（HTTPS 証明書は自動） | ○ | ○ | ○（Free は 2 つまで） |
| `.gz` の扱い | `Content-Type: application/gzip`、`Content-Encoding` なし（実測。下記） | 要検証。`_headers` で `Content-Type` は指定できるが、エッジの圧縮と `.gz` の扱いは実際に確かめる必要がある | 自分でサーバーを書く（`serveDir` など）ので制御できる。その分コードの管理が要る | 要検証。事前圧縮ファイル（`*.gz`）を自動で使う機能があり、`.dat.gz` への影響を確かめる必要がある |
| ヘッダーの制御 | できない（`Cache-Control: max-age=600` 固定） | `_headers` で自由に設定できる | 自由に設定できる | `staticwebapp.config.json` で設定できる |
| CORS | `Access-Control-Allow-Origin: *` が付く（実測） | 設定次第 | 設定次第 | 設定次第 |
| プレビュー環境 | なし（本番 1 つだけ） | ブランチ・PR ごとに自動 | ブランチ・PR ごとに自動 | PR ごとに自動（Free は 3 つまで） |
| 設定の手間 | 同じリポジトリで完結する。Settings で Source を GitHub Actions にするだけ | 外部アカウントと API トークン（Secrets）が必要 | 外部アカウントと GitHub 連携が必要 | Azure のサブスクリプションとデプロイトークンが必要 |

GitHub Pages の `.gz` の扱いは、GitHub Pages で kuromoji の辞書を配信している既存のサイト（`takuyaa.github.io/kuromoji.js`）に対して、`Accept-Encoding: gzip, br` を付けてリクエストして確かめた。結果は `content-type: application/gzip` で、`content-encoding` は付かず、`access-control-allow-origin: *` が付いていた。

CORS について補足する。辞書はタスクペインと同じオリジンから配信するので、どの配信先でも CORS の設定は要らない。ほかのオリジンから辞書を読む構成にする場合だけ問題になる。

## 決定

**GitHub Pages を使う。**

- 同じリポジトリの中で完結する。外部のアカウント、API トークン、Secrets が要らない
- `.dat.gz` をそのまま（`Content-Encoding` なしで）返すことを実測で確認した
- 無料枠（帯域 100 GB/月）で当面は足りる。辞書は初回に約 15 MB を取得したあと IndexedDB にキャッシュされるので、1 利用者あたりの転送量は小さい
- カスタムドメインを後から付けられる

### 受け入れる欠点

- **プレビュー環境がない。** PR ごとの動作確認は、開発サーバー（`npm run dev`）に sideload して行う。本番と同じ構成で確認したくなったら、Cloudflare Pages か Azure Static Web Apps をプレビュー専用に追加することを検討する（その場合は `.gz` の扱いを先に検証する）
- **ヘッダーを変えられない。** `Cache-Control: max-age=600` で固定なので、ファイル名にハッシュが付いた JS / CSS でも長期キャッシュにならない。辞書は IndexedDB にキャッシュされるので影響は小さい。逆に、ハッシュの付かないファイルは、デプロイ後も最大 10 分は古いものが使われうる。textlint の Worker はファイル名にハッシュを付けて対処した（下記「Worker のキャッシュ対策」）
- **リポジトリのサブパス（`/excel-text-lint/`）で配信される。** Vite の `base` をこれに合わせる（下記）

## 構成

### URL

| 用途 | URL |
|---|---|
| タスクペイン | `https://75asa.github.io/excel-text-lint/taskpane/taskpane.html` |
| 辞書 | `https://75asa.github.io/excel-text-lint/dict/<名前>.dat.gz` |
| 本番用 manifest | `https://75asa.github.io/excel-text-lint/manifest.prod.xml` |

配信先の URL は `scripts/hosting.ts` の `DEFAULT_PROD_BASE_URL` で決まる。環境変数 `ADDIN_BASE_URL` で上書きできる。Vite の `base`（build のときだけ）と本番用 manifest の URL は、どちらもこの値から作る。カスタムドメインに移すときは、リポジトリの Variables に `ADDIN_BASE_URL`（例: `https://textlint.example.com/`）を設定する。

### build（`npm run build`）

`npm run build:worker` で Worker を生成してから `vite build` を実行する。`vite build` では、`scripts/` の Vite プラグインが次の処理をする。

1. `scripts/textlint-worker.ts` のプラグインが、ローダー（`src/textlint/loader.js`）と Worker（`npm run build:worker` が生成した `src/textlint/textlint-worker.js`）を、ファイル名にハッシュを付けて `dist/textlint/` に出力する（下記「Worker のキャッシュ対策」）
   - ローダーの `KUROMOJI_CDN`、Worker が参照する辞書の URL、インストールされている kuromoji のバージョンが一致しなければ build を失敗させる（下記「辞書の更新」）
2. `scripts/hosting.ts` のプラグインが、`node_modules/kuromoji/dict/*.dat.gz` を `dist/dict/` に **そのまま** コピーする（展開も再圧縮もしない）。kuromoji のバージョンは、textlint のバンドルが参照する jsdelivr の URL（`kuromoji@0.1.2`）に合わせて固定する
3. `scripts/hosting.ts` のプラグインが、`manifest.xml`（開発用）から `dist/manifest.prod.xml`（本番用）を生成する
   - `https://localhost:3000/` を配信先の URL に置き換える
   - `Id` を本番用の Id に置き換える。開発用と別の Id にすることで、両方を同時に sideload できる
   - `DisplayName` から ` (dev)` を外す
   - `localhost` が残っていたら build を失敗させる

できあがる `dist/` の構成:

```
dist/
  taskpane/taskpane.html
  assets/                 アイコンと、Vite が出力する JS / CSS
  textlint/loader-<hash>.js
  textlint/textlint-worker-<hash>.js
  dict/*.dat.gz           kuromoji の辞書（12 ファイル）
  manifest.prod.xml       本番用 manifest
```

本番用 manifest は `npm run validate:prod` で検証する（`npm run build` の後に実行する）。

### デプロイ（`.github/workflows/deploy.yml`）

main への push（と手動実行）で動く。

1. build: `npm ci`、`npm run build`、`npm run validate:prod` を実行し、`dist/` を Pages のアーティファクトとしてアップロードする
2. deploy: `actions/deploy-pages` で公開する
3. smoke-test: タスクペイン、manifest、辞書を実際に取得して、辞書のヘッダーとマジックナンバーを確認する。Worker は、タスクペインの HTML から JS → ローダー → Worker と実際の URL をたどって、JavaScript の `Content-Type` で取得できることを確認する（`scripts/smoke-test-worker.ts`）

### 事前に必要な設定（リポジトリの管理者が行う）

1. Settings > Pages > Build and deployment > Source を **GitHub Actions** にする
2. Settings > Environments > `github-pages` の Deployment branches が `main` を許可していることを確認する（Source を GitHub Actions にすると自動で作られる）
3. （任意）カスタムドメインを使う場合は Settings > Pages > Custom domain を設定し、Settings > Secrets and variables > Actions > Variables に `ADDIN_BASE_URL` を追加する

## Worker のキャッシュ対策（#45）

textlint の Worker はルールや prh の辞書を変えるたびに中身が変わる。GitHub Pages は `Cache-Control: max-age=600` で固定なので、ファイル名が同じままだと、デプロイ後も利用者のブラウザ（Office の WebView）に古い Worker が残ることがある。そこで、ローダーと Worker のファイル名にコンテンツハッシュを付ける。

### 比較

| 方式 | 内容 | 評価 |
|---|---|---|
| (a) Vite の worker / asset として取り込む | `new URL("./textlint-worker.js", import.meta.url)` や `?worker&url` / `?url` で import し、Vite にハッシュを付けさせる | `?worker` は Worker を Vite（Rolldown）でバンドルし直す。script-compiler の出力は約 1.3 MB の classic worker で、ローダーから `importScripts` で読むので、バンドルし直す意味がない。`?url` / `new URL()` ならそのままコピーされるが、ローダーの中の Worker の名前（`importScripts` の引数）は書き換えられない。dev では `src/` の下の JS が Vite の変換（import の解析など）を通ってしまう |
| (b) 自前のプラグインでハッシュ付きの名前にする | build のときに内容のハッシュで名前を決め、ローダーの中の Worker の名前を書き換えてから出力する。タスクペインには仮想モジュールでパスを渡す | 採用。ローダーのハッシュは Worker の名前を埋め込んだ後の内容で取るので、Worker が変わればローダーの名前も変わる。dev ではハッシュを付けず、変換も通さずに配信できる |
| (c) クエリ文字列でバージョンを付ける（`loader.js?v=<hash>`） | ファイル名は変えず、URL のクエリを変える | 不採用。GitHub Pages の CDN（Fastly）は **クエリを無視してキャッシュする**（実測。下記）。また、ローダーから `importScripts` する Worker の URL にもクエリを引き継ぐ必要がある |

(c) の実測（2026-09-28）: `https://75asa.github.io/excel-text-lint/manifest.prod.xml?v=<乱数>` を取得すると `x-cache: MISS` だった。直後にクエリなしの URL と、別のクエリ（`?v=x<乱数>`）の URL を取得すると、どちらも `x-cache: HIT` で、`x-github-request-id`（オリジンへのリクエストの ID）が最初のものと同じだった。CDN のキャッシュのキーにクエリが含まれていないので、クエリを変えても CDN に残った古いファイルが返りうる。

### 仕組み

- `scripts/textlint-worker.ts` の Vite プラグイン
  - build: Worker を `textlint/textlint-worker-<Worker のハッシュ>.js` に、ローダーの中の `"./textlint-worker.js"` をその名前に書き換えたものを `textlint/loader-<書き換えた後のローダーのハッシュ>.js` に出力する。ハッシュは SHA-256 の先頭 8 桁で、同じ内容なら同じ名前になる
  - dev: ハッシュを付けずに `/textlint/loader.js`・`/textlint/textlint-worker.js` で配信する。Vite の変換は通さない。起動時に Worker が無いか、`.textlintrc.json`・`package-lock.json`・`prh/` より古ければ `npm run build:worker` を実行する
  - タスクペインには仮想モジュール `virtual:textlint-worker` でローダーのパスを渡す。タスクペインは `textlintWorkerUrl()`（`src/core/worker-url.ts`）で URL を組み立てる
- タスクペインの HTML（`taskpane/taskpane.html`）にはハッシュが付かないので、デプロイ後の最大 10 分は古い HTML が使われうる。古い HTML は古い JS → 古いローダー → 古い Worker を参照するが、Pages はデプロイのたびにサイト全体を置き換えるので、古いファイルは 404 になる。その場合はタスクペインを開き直せば（10 分以内に）新しいものが読まれる

## 辞書の更新

kuromoji の辞書（`dist/dict/*.dat.gz`）は、Worker の中の kuromoji のローダーが **IndexedDB にキャッシュする**。HTTP のキャッシュ（`Cache-Control`）やファイル名のハッシュとは関係なく動く。

- キャッシュの場所: タスクペインのオリジン（`https://75asa.github.io`）の IndexedDB、データベース `@textlint/runtime-helper` の `kvs`
- キー: `kuromoji::https://cdn.jsdelivr.net/npm/kuromoji@0.1.2/dict/<名前>.dat.gz`。ローダーが書き換える **前** の URL なので、辞書の配信先を変えてもキャッシュは使い回される
- 値: 展開した後の辞書。12 ファイルで約 96 MiB（圧縮した状態では約 17 MB）
- 有効期限はない。一度キャッシュされると、同じキーでは二度とネットワークから取得しない。取得や展開に失敗したものはキャッシュされない

### 方針

- **辞書のバージョンは kuromoji@0.1.2 に固定する。** npm に公開されたパッケージの中身は変わらないので、同じバージョンのままなら辞書を更新する必要はない。`package.json` の `kuromoji` は範囲指定なしの `0.1.2` にしている
- 辞書の中身を差し替えたいとき（同じバージョンのまま別の辞書を配信する）は、キーが変わらないので既存の利用者には届かない。この運用はしない

### kuromoji のバージョンが変わったとき

textlint のルール（kuromojin）が別のバージョンの kuromoji を参照するようになると、Worker が取得する辞書の URL（`kuromoji@<新しいバージョン>/dict/...`）が変わる。

- キャッシュのキーも変わるので、利用者は初回に辞書を取得し直す（約 17 MB）。古いバージョンのキャッシュ（約 96 MiB）は IndexedDB に残り、自動では消えない。Office のキャッシュを消せば消える（[Clear the Office cache](https://learn.microsoft.com/office/dev/add-ins/testing/clear-cache)）
- ローダーの `KUROMOJI_CDN` が古いままだと、辞書の URL が書き換わらず jsdelivr から取得してしまう（CDN がブロックされる社内ネットワークでは lint できない）。これを防ぐため、build は次の 3 つが一致しなければ失敗する（`scripts/textlint-worker.ts` の `checkKuromojiDictUrl`）
  - `src/textlint/loader.js` の `KUROMOJI_CDN`
  - Worker が参照している辞書の URL
  - インストールされている `kuromoji` のバージョン（`dist/dict/` にコピーする辞書）
- 更新の手順: `package.json` の `kuromoji` を新しいバージョンに（範囲指定なしで）変え、`src/textlint/loader.js` の `KUROMOJI_CDN` とこの文書のバージョンを合わせて、`npm run build` が通ることを確かめる。古いキャッシュを消す処理を入れるかは、そのときに検討する
