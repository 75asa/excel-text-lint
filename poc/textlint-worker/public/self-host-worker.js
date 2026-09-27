// kuromoji の辞書 URL（jsdelivr にハードコード）を自前の配信先へ書き換えてから textlint worker を読み込む
const CDN = "https://cdn.jsdelivr.net/npm/kuromoji@0.1.2/dict";
const LOCAL = new URL("./dict", self.location.href).href;
const originalFetch = self.fetch;
self.fetch = (input, init) => {
  const url = typeof input === "string" ? input : input.url;
  return originalFetch(url.startsWith(CDN) ? LOCAL + url.slice(CDN.length) : input, init);
};
importScripts("./textlint-worker.js");
