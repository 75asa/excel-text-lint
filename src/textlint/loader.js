// textlint の Worker（textlint-worker.js）を起動するローダー。classic worker として読み込む。
//
// textlint-worker.js は @textlint/script-compiler が生成する（npm run build:worker）。
// その中の kuromoji は辞書の URL が jsdelivr にハードコードされていて、外から変えられない（#2 の PoC を参照）。
// そこで fetch をラップして、辞書の URL だけを自前の配信先に書き換える。
//
// 配信先は Worker の URL のクエリ `dict` で渡す（例: loader.js?dict=https://example.com/dict）。
// 省略したときは書き換えず、jsdelivr から取得する。
// IndexedDB のキャッシュのキーは書き換える前の URL なので、配信先を変えてもキャッシュは使い回される。
//
// build では、このファイルと textlint-worker.js のファイル名にコンテンツハッシュを付けて dist/textlint/ に出力する。
// そのとき、末尾の importScripts に渡している Worker の相対パス（文字列リテラル）もハッシュ付きの名前に書き換える
// （scripts/textlint-worker.ts）。このリテラルと KUROMOJI_CDN は build で検査するので、形を変えないこと。

const KUROMOJI_CDN = "https://cdn.jsdelivr.net/npm/kuromoji@0.1.2/dict";

const dict = new URL(self.location.href).searchParams.get("dict");
if (dict) {
  const base = new URL(dict, self.location.href).href.replace(/\/+$/, "");
  const originalFetch = self.fetch.bind(self);
  self.fetch = (input, init) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
    return url.startsWith(KUROMOJI_CDN)
      ? originalFetch(base + url.slice(KUROMOJI_CDN.length), init)
      : originalFetch(input, init);
  };
}

importScripts(new URL("./textlint-worker.js", self.location.href).href);
