# 同梱するルールセット

Issue #10 の調査結果。アドインが最初から使えるルールセットを決める。

- 調査日: 2026-09-28
- 検証コード: `poc/textlint-worker/`（`npm run compare` と `npm run bench`）
- 推奨設定: [`poc/textlint-worker/.textlintrc.json`](../poc/textlint-worker/.textlintrc.json) と [`poc/textlint-worker/prh/business-ja.yml`](../poc/textlint-worker/prh/business-ja.yml)

## 結論

次の 3 つを組み合わせ、Excel で誤検出が多いルールを無効にする。

| 構成要素 | 役割 |
|---|---|
| `textlint-rule-preset-ja-technical-writing` | 文章の質（冗長な表現、ら抜き、読点の数、助詞の連続、二重否定、半角カナなど） |
| `textlint-rule-preset-jtf-style` | 表記の統一（全角英数字、半角の `!` や `()`、全角と半角の間のスペースなど）。ほぼすべて自動修正できる |
| `textlint-rule-prh` + 自前の辞書 `business-ja.yml` | 業務文書でよく出る漢字をひらがなにそろえる（「下さい」「致します」「頂く」「出来る」「事」など） |

```json
{
  "rules": {
    "preset-ja-technical-writing": {
      "ja-no-mixed-period": false,
      "max-kanji-continuous-len": false,
      "no-exclamation-question-mark": false
    },
    "preset-jtf-style": {
      "1.2.1.句点(。)と読点(、)": false,
      "2.2.2.算用数字と漢数字の使い分け": false,
      "4.2.7.コロン(：)": false
    },
    "prh": {
      "rulePaths": ["./prh/business-ja.yml"]
    }
  }
}
```

理由:

1. **ブラウザ（script-compiler の Worker）で動く。** 候補のうち、実行時に `fs` を読むルール（SmartHR のプリセット、ひらがな化ルールの一部、web-plus-db）は Worker で読み込みに失敗した。推奨の 3 つは問題なく動く
2. **PoC で検出できなかったものを検出できる。** 全角数字（`１２３`）は jtf-style、「下さい」は自前の prh 辞書で検出できる。文章のセル 26 個のうち 19 個で違反を検出した（PoC の設定では 13 個）
3. **短いセルで誤検出しない。** 見出しや単語だけの 10 セルで、PoC の設定は 8 セルで違反を出した（うち 7 セルは句点の要求）。推奨の設定で出るのは 2 セル（`要確認!` と `第３四半期`）だけで、どちらも半角の `!` と全角数字を正しく検出したもの
4. **バンドルサイズは PoC とほぼ同じ。** jtf-style と prh は technical-writing がすでに内部で依存しているため、technical-writing 単体より 8 KB 増えるだけ。ja-spacing をやめた分、PoC の設定より 4 KB 小さい（後述）
5. **自動修正できる違反が多い。** jtf-style と prh はほぼすべて fixable。#19 の自動修正と相性がよい
6. **ライセンスはすべて MIT。** 辞書はリポジトリで作るので、再配布の問題もない（#12）

## 候補の比較

### 一覧

サイズは、各ルールだけを入れて `@textlint/script-compiler` でビルドした Worker のもの。textlint 本体だけ（ルールなし）で 252 KB（gzip 後 70 KB）ある。

| 候補 | 最新版 | 最終更新 | ライセンス | ブラウザで動くか | サイズ（gzip 後） | 本体からの増分（gzip 後） | fixable |
|---|---|---|---|---|---|---|---|
| textlint 本体のみ（比較用） | 15.8.0 | - | MIT | 動く | 252 KB（70 KB） | - | - |
| preset-ja-technical-writing | 12.0.2 | 2025-01 | MIT | 動く | 1,291 KB（348 KB） | +1,039 KB（+278 KB） | 一部（冗長な表現、半角カナなど） |
| preset-japanese | 10.0.4 | 2025-01 | MIT | 動く | 629 KB（178 KB） | +377 KB（+108 KB） | なし（今回の文では） |
| preset-jtf-style | 3.0.3 | 2025-09 | MIT | 動く | 856 KB（227 KB） | +604 KB（+157 KB） | ほぼすべて |
| preset-ja-spacing | 3.0.3 | 2026-08 | MIT | 動く | 291 KB（80 KB） | +39 KB（+10 KB） | ほぼすべて |
| prh + 自前の辞書 | 6.1.0 | 2025-04 | MIT（辞書は自作） | 動く（辞書はビルド時にインライン化） | 776 KB（205 KB） | +524 KB（+135 KB） | すべて |
| prh + WEB+DB PRESS 辞書（prh/rules） | 6.1.0 | 辞書は 2026-07 | ルールは MIT、辞書は LICENSE ファイルなし（後述） | 動く（同上） | 863 KB（224 KB） | +611 KB（+154 KB） | すべて |
| textlint-rule-web-plus-db | 1.1.5 | 2022-06 | MIT | **動かない**（実行時に `fs.readFileSync`） | 380 KB（107 KB） | - | - |
| preset-smarthr | 1.38.1 | 2026-09 | MIT | **動かない**（依存する ja-hiragana-hojodoushi と ja-hiragana-daimeishi が `fs.readFileSync`。ルールを無効にしても読み込み時に失敗する） | 1,463 KB（392 KB） | - | - |
| ja-hiragana-hojodoushi / ja-hiragana-fukushi | 1.1.0 / 1.3.0 | 2023-02 / 2022-06 | MIT | **動かない**（同上） | - | - | - |
| ja-hiragana-keishikimeishi | 1.1.0 | 2022-06 | MIT | 動く | 342 KB（96 KB） | +90 KB（+26 KB） | すべて |
| ja-no-orthographic-variants | 2.0.0 | 2022-05 | MIT | 動く | **1,953 KB**（346 KB） | +1,701 KB（+276 KB） | なし |
| no-doubled-joshi（単体） | 5.1.1 | 2025-09 | MIT | 動く | 574 KB（162 KB） | +322 KB（+92 KB） | なし |
| **推奨の組み合わせ** | - | - | MIT | 動く | **1,299 KB（350 KB）** | +1,047 KB（+280 KB） | 多い |

- no-doubled-joshi は technical-writing と japanese の両方に入っているため、単体では足さない
- npm の `prh-rules` は prh の辞書ではない。npm のセキュリティ保持用の空パッケージ（`0.0.1-security`）なので、**インストールしないこと**。prh の辞書集は GitHub の [prh/rules](https://github.com/prh/rules) にだけある
- `TEXLINT_COMPILER_INLINING=1` を付けてビルドすると、`fs.readFileSync` のインライン化（babel-plugin-static-fs）が有効になる。web-plus-db は動くようになったが、ひらがな化ルールと SmartHR のプリセットはパスを動的に組み立てているため、インライン化できずに失敗した

### ブラウザでの動作と prh の辞書

- script-compiler は webpack で `fs` を空にしてバンドルする（`resolve.fallback.fs = false`）。実行時にファイルを読むルールは、Worker で `readFileSync is not a function` になる
- prh の `rulePaths` は、script-compiler が内部で使う `@textlint/config-inliner` が **ビルド時に YAML を読み、`ruleContents` に書き換えてバンドルに埋め込む**。そのため、ブラウザでもファイルパスで辞書を指定できる
  - この書き換えは、トップレベルの `prh` ルールにだけ効く。プリセットの中の prh には効かない（jtf-style や SmartHR のプリセットは、もともと辞書を文字列で埋め込んでいるので問題ない）
  - 辞書の `imports:` はインライン化されない。辞書は 1 ファイルで完結させる
  - ユーザーが辞書を追加・編集する場合（#11）は、ビルド時のインライン化は使えない。YAML の文字列を `ruleContents` として渡し、Worker の中でルール設定を組み立てる形にする（#8）
- 古い web-plus-db に同梱の辞書は、現在の prh（v5 以降）では正規表現の `u` フラグのエラーで読み込めない。Node でも動かない

### 検出力（業務文書の日本語 36 セル）

入力は `poc/textlint-worker/corpus.mjs`。S01〜S26 は文章のセル、H01〜H10 は見出し・単語・値だけのセル。1 セルずつ Worker で lint した結果。

| セル | 内容 | technical-writing | japanese | jtf-style | ja-spacing | prh + 自前の辞書 | prh + WEB+DB | keishikimeishi | orthographic-variants | **推奨** |
|---|---|---|---|---|---|---|---|---|---|---|
| S01 | `サーバーを再起動して下さい。` | - | - | - | - | prh | prh | - | - | prh |
| S02 | `１２３件のデータがあります。` | - | - | jtf 2.1.8 | - | - | - | - | - | jtf 2.1.8 |
| S03 | `データを取得することができる。` | ja-no-redundant-expression | - | - | - | - | prh | - | - | ja-no-redundant-expression |
| S04 | `出来る限り早く対応致します。` | - | - | - | - | prh | prh | - | - | prh |
| S05 | `内容を確認の上、ご連絡させて頂きます。` | - | - | - | - | prh | prh | - | - | prh |
| S06 | `この画面からは設定が見れない。` | no-dropping-the-ra | no-dropping-the-ra | - | - | - | - | - | - | no-dropping-the-ra |
| S07 | `先方に確認し、その上で、担当者と相談し、改めて、ご連絡します。` | max-ten | max-ten | - | - | - | prh | - | - | max-ten |
| S08 | `材料不足で代替素材で製品を作った。` | no-doubled-joshi | no-doubled-joshi | - | - | - | - | - | - | no-doubled-joshi |
| S09 | `資料を作成した。明日提出します。` | - | - | - | - | - | - | - | - | - |
| S10 | `この方法で問題ないと思われます。` | - | - | - | - | - | - | - | - | - |
| S11 | `ユーザー登録が完了しました!` | no-exclamation-question-mark | - | jtf 4.2.1 | - | - | - | - | - | jtf 4.2.1 |
| S12 | `見積書(案)を添付します。` | - | - | jtf 4.3.1 | - | - | - | - | - | jtf 4.3.1 |
| S13 | `Excelファイルを共有フォルダに保存してください。` | - | - | - | - | - | - | - | - | - |
| S14 | `Excel ファイルを共有フォルダに保存してください。` | - | - | jtf 3.1.1 | ja-space-between-half-and-full-width | - | - | - | - | jtf 3.1.1 |
| S15 | `ＡＢＣ社のＷｅｂサイトを更新します。` | ja-unnatural-alphabet | - | jtf 2.1.9 | - | - | prh | - | - | jtf 2.1.9<br>ja-unnatural-alphabet |
| S16 | `今期の売上は前年比１１０％です。` | - | - | jtf 2.1.8 | - | - | - | - | - | jtf 2.1.8 |
| S17 | `但し、例外として事前申請が必要な場合もある。` | - | - | - | - | prh | prh | - | - | prh |
| S18 | `〇〇について検討を行う事とする。` | ja-no-redundant-expression | - | - | - | prh | - | ja-hiragana-keishikimeishi | - | ja-no-redundant-expression<br>prh |
| S19 | `宜しくお願い致します。` | - | - | - | - | prh | - | - | - | prh |
| S20 | `ｻｰﾊﾞｰの設定を変更しました。` | no-hankaku-kana | - | - | - | - | - | - | - | no-hankaku-kana |
| S21 | `サーバの設定とサーバーの設定を比較する。` | - | - | - | - | - | prh | - | - | - |
| S22 | `javascriptとGithubを使って開発する。` | - | - | - | - | - | prh | - | - | - |
| S23 | `問題が無いわけではない。` | no-double-negative-ja | no-double-negative-ja | - | - | - | prh | - | ja-no-orthographic-variants | no-double-negative-ja |
| S24 | `この件はこれは問題ありません` | no-doubled-joshi<br>ja-no-mixed-period | no-doubled-joshi | - | - | - | - | - | - | no-doubled-joshi |
| S25 | `詳細は、以下の手順を参照してください。、` | ja-no-mixed-period | - | - | - | - | - | - | - | - |
| S26 | `1行目の説明です⏎2行目の説明です` | ja-no-mixed-period | - | - | - | - | - | - | - | - |
| H01 | `売上集計` | ja-no-mixed-period | - | - | - | - | - | - | - | - |
| H02 | `対応済` | ja-no-mixed-period | - | - | - | - | - | - | - | - |
| H03 | `1.概要` | ja-no-mixed-period | - | jtf 1.2.1 | - | - | - | - | - | - |
| H04 | `担当者:山田` | ja-no-mixed-period | - | jtf 4.2.7 | - | - | - | - | - | - |
| H05 | `業務効率化推進委員会議事録` | max-kanji-continuous-len<br>ja-no-mixed-period | - | - | - | - | - | - | - | - |
| H06 | `2026/04/01` | - | - | - | - | - | - | - | - | - |
| H07 | `¥10,000` | - | - | - | - | - | - | - | - | - |
| H08 | `要確認!` | no-exclamation-question-mark | - | jtf 4.2.1 | - | - | - | - | - | jtf 4.2.1 |
| H09 | `第３四半期` | ja-no-mixed-period | - | jtf 2.1.8 | - | - | - | - | - | jtf 2.1.8 |
| H10 | `〜3月末` | ja-no-mixed-period | - | - | - | - | - | - | - | - |

`⏎` はセル内改行。SmartHR のプリセット、ひらがな化ルール（hojodoushi / fukushi）、web-plus-db はブラウザで動かないため表にない。参考までに Node で lint すると、ひらがな化ルールは S01、S04、S05、S07（「改めて」）、S18、S19 を検出した。

読み取れること:

- **technical-writing は文章の質に強いが、表記には弱い。** 全角数字、「下さい」、半角の `()` は検出しない。いっぽうで短いセルにも句点を求めるため、見出し・単語だけの 10 セルのうち 7 セルで ja-no-mixed-period が出る
- **japanese は technical-writing の一部。** 検出した違反は technical-writing でもすべて出る。technical-writing を入れるなら不要
- **jtf-style は表記の統一に強い。** 全角英数字、半角の `!`、`()`、スペースを検出し、ほぼすべて自動修正できる。ただし `1.概要` の `.` と `担当者:山田` の `:` にも反応する
- **ja-spacing は jtf-style 3.1.1 と重なる。** 今回の文で出たのは S14 だけで、jtf-style でも同じ箇所が出る。同時に入れると同じ違反が 2 回出るので入れない
- **prh の辞書で「下さい」などを拾える。** WEB+DB PRESS の辞書は技術用語（JavaScript、GitHub）に強いが、`サーバー` を `サーバ` に直すなど、業務文書の一般的な表記とは方針が違う
- **表記ゆれはセル単位では検出しにくい。** ja-no-orthographic-variants は 1 つの文書の中での揺れを見るルールで、1 セルずつ lint すると役に立たない（S21 の「サーバ」と「サーバー」も出ない）。そのうえバンドルが 1.7 MB 増える

### 自動修正（fix）の結果（推奨の設定）

| 入力 | fix の出力 |
|---|---|
| `サーバーを再起動して下さい。` | `サーバーを再起動してください。` |
| `１２３件のデータがあります。` | `123件のデータがあります。` |
| `データを取得することができる。` | `データを取得できる。` |
| `出来る限り早く対応致します。` | `できる限り早く対応いたします。` |
| `見積書(案)を添付します。` | `見積書（案）を添付します。` |
| `Excel ファイルを共有フォルダに保存してください。` | `Excelファイルを共有フォルダに保存してください。` |
| `ＡＢＣ社のＷｅｂサイトを更新します。` | `ABC社のWｅｂサイトを更新します。`（小文字の全角英字が直らない） |
| `ｻｰﾊﾞｰの設定を変更しました。` | `サーバーの設定を変更しました。` |
| `宜しくお願い致します。` | `よろしくお願いいたします。` |

fixable でないルール: no-dropping-the-ra、max-ten、no-doubled-joshi、no-double-negative-ja、ja-unnatural-alphabet（technical-writing）、japanese のすべて、ja-no-orthographic-variants。ja-no-redundant-expression は辞書の項目によって fixable なものとそうでないものがある。

## Excel 向けに無効化・調整するルール

| ルール | 設定 | 理由 |
|---|---|---|
| technical-writing / ja-no-mixed-period | 無効 | 文末に句点を求める。見出し・単語・値だけのセルでほぼ必ず誤検出する（H01〜H10 のうち 7 セル）。オプションに文字数の下限などはない |
| technical-writing / max-kanji-continuous-len | 無効 | 漢字 7 文字以上の連続を禁じる。部署名・会議名・製品名などの見出しで誤検出する（H05） |
| technical-writing / no-exclamation-question-mark | 無効 | `!` `?` を一律で禁じる。業務のメモでは全角の `！` `？` を使うことがある。半角は jtf-style 4.2.1 / 4.2.2 で全角に直す |
| jtf-style / 1.2.1 句点と読点 | 無効 | `1.概要` のような番号付き見出しの `.` に反応する（H03） |
| jtf-style / 4.2.7 コロン | 無効 | `担当者:山田` のような「項目:値」のセルで全角を求める（H04）。Excel ではよく使う書き方 |
| jtf-style / 2.2.2 算用数字と漢数字 | 無効 | technical-writing が同じルールを `arabic-kanji-numbers` として含むため、同じ違反が 2 回出る |
| preset-ja-spacing | 入れない | jtf-style 3.1.x と重なる |

## 推奨の設定で bench を回した結果

2026-09-28、Apple Silicon の Mac、ヘッドレス Chrome。`CELLS=1000 npm run bench` を 2 回ずつ回した範囲。PoC の設定（technical-writing + ja-spacing）も同じマシンで計測し直した。

| 項目 | PoC の設定 | 推奨の設定 |
|---|---|---|
| Worker バンドル | 1,334,633 B（gzip 後 357 KB） | 1,330,184 B（gzip 後 358 KB） |
| Worker の起動（`init` まで） | 77〜96 ms | 76〜84 ms |
| 最初の lint（辞書のロードを含む、self-host） | 0.88〜0.97 秒 | 0.92〜0.93 秒 |
| 2 回目以降の起動と最初の lint | 0.20〜0.23 秒 | 0.22〜0.24 秒 |
| 1,000 セルを 1 セル 1 メッセージで lint | 0.73〜0.79 秒 | 0.89〜1.07 秒 |
| bench の 10 文に対する違反の数（1,000 セル分） | 600 | 600 |

- サイズは変わらない。jtf-style と prh は technical-writing の依存としてすでにバンドルに入っている（technical-writing は jtf-style の 2.2.2 を使い、ja-no-abusage が prh を使う）。ja-spacing をやめた分と相殺している
- lint は 1,000 セルで 0.2 秒ほど（2〜3 割）遅くなる。jtf-style のルール数（38）と prh の辞書照合が増えたため。1 セルあたり 1 ms 程度で、#17 の分割処理の前提は変わらない
- 違反の数が同じなのは偶然。内訳は変わっている（ja-no-mixed-period の誤検出が消え、全角数字と「下さい」が出る）

## 相談したい点

推奨案を先に書き、ほかの選択肢を並べる。

1. **ja-no-mixed-period を完全に無効にするか**
   - 推奨: 無効にする
   - 失うもの: `。、` のような句読点の連続（S25）や、文章のセルでの句点の付け忘れを検出できなくなる
   - 選択肢: エンジン（#8）で「句点を含むセル」や「一定の文字数以上のセル」にだけ適用する。textlint のオプションでは表現できないので、ルールの結果をエンジン側で間引く形になる
2. **自前の辞書（`business-ja.yml`）で始めるか**
   - 推奨: 自前の小さな辞書で始める。項目は 14 個で、誤検出しにくいものだけにした（補助動詞の「下さい」「頂く」、「致します」「出来る」「宜しく」「但し」「予め」、形式名詞の「事」「時」「所」、誤記の「シュミレーション」など）。辞書の中身は PR で育てる
   - 選択肢 A: prh/rules の WEB+DB PRESS 辞書を同梱する。技術用語に強いが、`サーバー` → `サーバ` のように業務文書と方針が違う項目がある。また、リポジトリに LICENSE ファイルがなく（package.json は `"license": "MIT"`, `"private": true`）、WEB+DB PRESS の辞書は「公開許可はもらってある」という注記だけ。同梱するなら #12 で確認が必要
   - 選択肢 B: ユーザーが辞書を追加できるようにし（#11）、WEB+DB PRESS などは「追加できる辞書の例」として案内する
3. **ひらがな化ルール（ja-hiragana-hojodoushi / fukushi）を使うか**
   - 形態素解析を使うので、prh の正規表現より精度が高い。ただし実行時に `fs.readFileSync` で辞書を読むため、Worker では動かない
   - 推奨: いまは使わず、prh の辞書で代わりをする
   - 選択肢: upstream に「辞書を JS として埋め込む」PR を出す。またはエンジン（#8）で自前のビルド設定を持つなら、webpack で辞書を差し替える
4. **半角の `!` と `:` の扱い**
   - 推奨: `!` は全角に直す（jtf 4.2.1 を有効）。`:` はそのまま（jtf 4.2.7 を無効）
   - 業務の文書では「半角のままでよい」という方針もありうる。組織ごとの設定（#11）で切り替えられればよい
5. **セルをまたいだ表記ゆれ**
   - 「サーバ」と「サーバー」が別のセルにある場合、1 セルずつ lint する方式では検出できない。ルールの問題ではなく、エンジンの設計の問題
   - 推奨: #10 の範囲外とし、別 Issue で「シート全体の表記ゆれの集計」を検討する

## 既知の制約

- jtf-style 2.1.9 は小文字の全角英字を直さない（`Ｗｅｂ` → `Wｅｂ`）。全角の `％` も検出しない
- である調とですます調の混在は、1 セルの中でしか検出できない。「した。」は である調として扱われないため、S09 も検出しない
- ja-hiragana-keishikimeishi など、古いルールは kuromojin 2.x に依存し、ほかのルールの kuromojin 3.x と重複してバンドルされる。推奨の組み合わせには含めていない

## ライセンス（#12 の前調査）

- 推奨の 3 つ（technical-writing、jtf-style、prh）と、その依存（kuromoji を含め約 280 パッケージ）を調べた。MIT が 271、それ以外は BSD-2-Clause（structured-source、boundary、esprima）、BSD-3-Clause（diff、sprintf-js）、ISC（graceful-fs など 3 つ）、Apache-2.0（kuromoji）
- kuromoji の辞書（IPADIC 由来）は、実行時に fetch する別ファイルとして配信する。IPADIC のライセンス表記（kuromoji の `NOTICE.md`）は、#7 で辞書を自前配信するときに同梱する
- jtf-style の辞書は JTF 日本語標準スタイルガイドに基づくが、パッケージとしては MIT で配布されている
- 自前の辞書 `business-ja.yml` はこのリポジトリのライセンスに従う
- prh/rules の辞書を同梱する場合は、上の「相談したい点」2 のとおり確認が必要

## 再現方法

```sh
cd poc/textlint-worker
npm install
npm run compare                           # configs/*.json と .textlintrc.json を比較（results/compare.json）
CONFIGS=jtf-style,recommended npm run compare
npm run build && CELLS=1000 npm run bench # 推奨の設定で計測
TEXTLINTRC=configs/technical-writing.json OUT_DIR=dist npm run build  # 別の設定でビルド
```

`configs/prh-webdb.json` は、[prh/rules](https://github.com/prh/rules) の `media/WEB+DB_PRESS.yml` を `poc/textlint-worker/results/prh-rules/` に置いてから使う（ライセンスが不明なのでリポジトリには入れていない）。
