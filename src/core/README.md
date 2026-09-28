# core

ホスト（Excel / Word / PowerPoint / OneNote）に依存しない層。Office.js（`Office` / `Excel` などのグローバル）には依存しないこと。ホストごとの処理は `src/hosts/<host>/` に置く。

| ファイル | 内容 |
|---|---|
| `types.ts` | `TextUnit` / `Violation` / `HostAdapter` / `HostCapabilities` など（#9） |
| `locations.ts` | ホストごとの location 型（Excel は実装済み、ほかは案） |
| `violation.ts` | textlint の結果を `Violation` にする純関数、表示用の範囲の計算、修正案の適用 |
| `engine.ts` | textlint の Worker を包むクライアント `LintEngine`（#8） |
| `worker-url.ts` | Worker（ローダー）の URL を組み立てる `textlintWorkerUrl()`。build ではファイル名にハッシュが付く（#45） |

## ドキュメント抽象（#9）

### TextUnit

セル・段落・シェイプなどを `TextUnit<L> { id, text, location }` として同じように扱う。

- `text` は collect した時点のテキスト。lint の入力で、applyFix の前に「ホストのテキストが変わっていないか」を確かめる基準にもなる
- `location` はホストごとの型（`ExcelLocation` など）で、core は中身を見ない。Office.js のプロキシオブジェクトは入れず、純粋なデータにする（`context.sync()` の外に持ち出せるように）
- lint は TextUnit 1 件ずつ行う。**連結して 1 回で lint してはいけない**（1 万セルで約 96 秒と超線形に遅くなる。#2 の PoC）

### Violation

`{ unitId, ruleId, message, severity, range, displayRange, line, column, fix? }`

- `range` は textlint が返した範囲そのもの（`TextUnit.text` の中の UTF-16 のオフセット）
- `displayRange` は表示・ハイライト用。textlint の `range` は 1 文字だけのことがあるので（「することができる。」が `[6, 7]` など）、`fix.range` と接していれば合わせて広げる。サロゲートペアの途中では切らない（`computeDisplayRange`）

### HostAdapter

ホストごとの差分（テキストの集め方と、違反箇所への移動・ハイライト・修正）を閉じ込める。OneNote のようにできることが少ないホストがあるので、**「できない操作がある」ことを前提にしている**（docs/research/onenote.md の「#9 への提言」）。

- `capabilities` で、使える範囲（`scopes`）、reveal の最良の精度、highlight の可否、applyFix の単位、選択の追従の可否を宣言する。UI はこれを見てボタンや説明を切り替える
- `collect(scope, { signal, onProgress })` は進捗とキャンセルに対応する（OneNote のセクション走査や大きなブックで時間がかかるため）
- `reveal(unit, range?)` は、実際にどこまでできたか（`exact` / `unit` / `container` / `none`）を返す。UI は結果に応じて文脈表示や検索語のコピーで補う
- `highlight` / `clearHighlight` は、何もしない実装でよい（`capabilities.highlight: false`）
- `applyFix(unit, fix)` は失敗しうる前提で、結果（`applied` / `stale` / `unsupported` / `failed`）を返す。実装は適用の直前に現在のテキストを読み直し、`unit.text` と違えば `stale` を返す
- `onSelectionChanged?` は任意。OneNote では reveal の代わりに、Excel / Word では「カーソル位置の違反を先に出す」機能に使う

各ホストでの想定:

| | scopes | reveal | highlight | applyFix | selectionTracking |
|---|---|---|---|---|---|
| Excel（#13〜#19） | selection / sheet / workbook | unit（セルを選択） | ✅（条件付き書式。解除で元どおり） | unit（セルの値を置き換え） | ✅ |
| Word（#21） | selection / document | exact | ✅ | range | ✅ |
| PowerPoint（#22） | selection / slide / presentation | unit（シェイプ） | △ | range | ✅ |
| OneNote（#24） | selection / page | container（ページを開く） | ❌ | none（将来は selection） | ✅ |

Excel アダプタ（`src/hosts/excel/`）の reveal / highlight / applyFix の方式は `src/hosts/excel/README.md` を参照。

## lint エンジン（#8）

### Worker のビルド

`npm run build:worker` で、`@textlint/script-compiler` がルート直下の `.textlintrc.json` と node_modules のルールから `src/textlint/textlint-worker.js` を生成する（git には入れない）。

- `npm run build` はこれを実行してから Vite で build する。Vite プラグイン（`scripts/textlint-worker.ts`）が、ローダーと Worker のファイル名にコンテンツハッシュを付けて `dist/textlint/loader-<hash>.js`・`dist/textlint/textlint-worker-<hash>.js` に出力する（#45。理由は `docs/hosting.md` の「Worker のキャッシュ対策」）
- dev サーバーは起動時に、Worker が無いか `.textlintrc.json`・`package-lock.json`・`prh/` より古ければ `npm run build:worker` を実行する。ハッシュは付けず `/textlint/loader.js`・`/textlint/textlint-worker.js` で（Vite の変換を通さずに）配信する
- タスクペインは `textlintWorkerUrl()`（`worker-url.ts`）でローダーの URL を組み立てて `LintEngine` に渡す。パスは仮想モジュール `virtual:textlint-worker` から受け取る（テストでは `src/testing/textlint-worker-stub.ts` に差し替える）

- `.textlintrc.json` は #10（`docs/rules.md`）で選んだ推奨のルールセット（preset-ja-technical-writing・preset-jtf-style・prh と `prh/business-ja.yml`。Excel で誤検出の多いルールは無効）。ルールを変えたら、そのパッケージを devDependencies に入れて `npm run build:worker` をやり直す。prh の辞書はビルド時に Worker に埋め込まれる（辞書の `imports:` は埋め込まれない）
- Worker は `src/textlint/loader.js`（ローダー）経由で起動する。生成された Worker の中の kuromoji は、辞書の URL が jsdelivr にハードコードされていて外から変えられない。ローダーが `fetch` をラップして、辞書の URL だけを自前の配信先に書き換える
- 辞書の配信先は、Worker の URL のクエリ `dict` で起動時に渡す（`LintEngine` の `dictBaseUrl`）。省略すると jsdelivr から取得する。タスクペインでは次のようにしている
  - `VITE_TEXTLINT_DICT_BASE_URL` があればそれ
  - build した成果物では `<base>/dict/`（辞書は build 時に `dist/dict/` へコピーする。#7）
  - dev サーバーでは jsdelivr
- 辞書の IndexedDB のキャッシュのキーは書き換える前の URL なので、配信先を変えてもキャッシュは使い回される

### LintEngine

```ts
const engine = new LintEngine({ workerUrl, dictBaseUrl });
await engine.init(); // 省略可（lint のときに起動する）
const violations = await engine.lint(units, { signal, onProgress });
const { output, remaining } = await engine.fix(unit);
```

- Worker とのやり取りは script-compiler の仕様どおり（送信 `lint` / `fix` / `merge-config`、受信 `init` / `lint:result` / `fix:result` / `error`）。リクエストごとに ID を振り、応答の順が入れ替わっても対応づける
- `lint` は TextUnit を 1 件ずつ送る。同時に送っておく数は `maxInFlight`（既定 8）まで。Worker の中では順に処理されるので、送りすぎるとキャンセルが効くまでが遅くなる
- `signal` で中断すると、まだ送っていない分は送らず、`signal.reason` で reject する。送った分の応答は捨てる
- id のない `error`（Worker 全体のエラー）が来たら、どのリクエストのものか分からないので、処理中のものをすべて失敗にする
- 空白だけのテキストは Worker に送らない
- `fix` はテキストを返すだけで、ホストには書き込まない（書き込みは `HostAdapter.applyFix`）
