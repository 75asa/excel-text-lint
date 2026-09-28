# タスクペインの HTTPS ホスティング

- Issue: #7
- 日付: 2026-09-28
- 決定: **GitHub Pages**（`https://75asa.github.io/excel-text-lint/`）

## 何を配信するか

Office アドインのタスクペインは、HTTPS で配信される静的ページとして読み込まれる。本番では次のファイルを配信する。

| パス | 中身 | 大きさ |
|---|---|---|
| `taskpane/taskpane.html`、`assets/*` | タスクペイン（Vite の build 成果物）とアイコン | 数 KB |
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
- **ヘッダーを変えられない。** `Cache-Control: max-age=600` で固定なので、ファイル名にハッシュが付いた JS / CSS でも長期キャッシュにならない。辞書は IndexedDB にキャッシュされるので影響は小さい
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

`vite build` の後に、`scripts/hosting.ts` の Vite プラグインが次の処理をする。

1. `node_modules/kuromoji/dict/*.dat.gz` を `dist/dict/` に **そのまま** コピーする（展開も再圧縮もしない）。kuromoji のバージョンは、textlint のバンドルが参照する jsdelivr の URL（`kuromoji@0.1.2`）に合わせて固定する
2. `manifest.xml`（開発用）から `dist/manifest.prod.xml`（本番用）を生成する
   - `https://localhost:3000/` を配信先の URL に置き換える
   - `Id` を本番用の Id に置き換える。開発用と別の Id にすることで、両方を同時に sideload できる
   - `DisplayName` から ` (dev)` を外す
   - `localhost` が残っていたら build を失敗させる

できあがる `dist/` の構成:

```
dist/
  taskpane/taskpane.html
  assets/                 アイコンと、Vite が出力する JS / CSS
  dict/*.dat.gz           kuromoji の辞書（12 ファイル）
  manifest.prod.xml       本番用 manifest
```

本番用 manifest は `npm run validate:prod` で検証する（`npm run build` の後に実行する）。

### デプロイ（`.github/workflows/deploy.yml`）

main への push（と手動実行）で動く。

1. build: `npm ci`、`npm run build`、`npm run validate:prod` を実行し、`dist/` を Pages のアーティファクトとしてアップロードする
2. deploy: `actions/deploy-pages` で公開する
3. smoke-test: タスクペイン、manifest、辞書を実際に取得して、辞書のヘッダーとマジックナンバーを確認する

### 事前に必要な設定（リポジトリの管理者が行う）

1. Settings > Pages > Build and deployment > Source を **GitHub Actions** にする
2. Settings > Environments > `github-pages` の Deployment branches が `main` を許可していることを確認する（Source を GitHub Actions にすると自動で作られる）
3. （任意）カスタムドメインを使う場合は Settings > Pages > Custom domain を設定し、Settings > Secrets and variables > Actions > Variables に `ADDIN_BASE_URL` を追加する
