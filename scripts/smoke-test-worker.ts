// デプロイ後の smoke-test（.github/workflows/deploy.yml）で使う。#45
//
// Worker のファイル名にはハッシュが付く（scripts/textlint-worker.ts）ので、決め打ちの URL では確かめられない。
// デプロイされたタスクペインの HTML から、実際に使われる URL を順にたどって取得できることを確かめる。
//
//   taskpane.html → <script type="module"> の JS → textlint/loader-<hash>.js → textlint-worker-<hash>.js
//
// 使い方: node scripts/smoke-test-worker.ts <配信先の URL（末尾は />
// 例:     node scripts/smoke-test-worker.ts https://75asa.github.io/excel-text-lint/

const LOADER_PATTERN = /textlint\/loader-[0-9a-f]{8}\.js/g;
const WORKER_PATTERN = /textlint-worker-[0-9a-f]{8}\.js/g;

const RETRIES = 5;
const RETRY_DELAY_MS = 10_000;

async function get(url: URL): Promise<{ body: string; contentType: string }> {
  let lastError: unknown;
  for (let attempt = 0; attempt <= RETRIES; attempt++) {
    if (attempt > 0) await new Promise((resolve) => setTimeout(resolve, RETRY_DELAY_MS));
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
      return {
        body: await response.text(),
        contentType: response.headers.get("content-type") ?? "",
      };
    } catch (error) {
      lastError = error;
      console.log(`retry ${attempt + 1}/${RETRIES}: ${url.href}: ${String(error)}`);
    }
  }
  throw new Error(`${url.href} を取得できません: ${String(lastError)}`);
}

/** classic worker と importScripts は JavaScript の MIME タイプでないと読み込めない */
async function getScript(url: URL): Promise<string> {
  const { body, contentType } = await get(url);
  if (!/javascript/i.test(contentType)) {
    throw new Error(`${url.href} の Content-Type が JavaScript ではありません: ${contentType}`);
  }
  console.log(`ok: ${url.href} (${contentType}, ${body.length} bytes)`);
  return body;
}

function unique(matches: Iterable<RegExpMatchArray>, where: string): string {
  const found = [...new Set([...matches].map((match) => match[0]))];
  const [name] = found;
  if (name === undefined || found.length !== 1) {
    throw new Error(
      `${where} にハッシュ付きの名前がちょうど 1 つ必要です: ${JSON.stringify(found)}`,
    );
  }
  return name;
}

async function main(pageUrl: string): Promise<void> {
  const base = new URL(pageUrl.endsWith("/") ? pageUrl : `${pageUrl}/`);
  const htmlUrl = new URL("taskpane/taskpane.html", base);
  const { body: html } = await get(htmlUrl);
  console.log(`ok: ${htmlUrl.href}`);

  // Vite が出力するエントリ（type="module"）と、分割されたチャンク（modulepreload）
  const scripts = [
    ...html.matchAll(/<script\b[^>]*\btype="module"[^>]*\bsrc="([^"]+)"/g),
    ...html.matchAll(/<link\b[^>]*\brel="modulepreload"[^>]*\bhref="([^"]+)"/g),
  ].map((match) => new URL(match[1] ?? "", htmlUrl));
  if (scripts.length === 0)
    throw new Error(`${htmlUrl.href} に type="module" の script がありません`);

  const bundles = await Promise.all(scripts.map(getScript));
  // タスクペインは base からの相対パス（textlint/loader-<hash>.js）で参照している（src/core/worker-url.ts）
  const loaderUrl = new URL(
    unique(
      bundles.flatMap((js) => [...js.matchAll(LOADER_PATTERN)]),
      "タスクペインの JS",
    ),
    base,
  );
  const loader = await getScript(loaderUrl);

  const workerUrl = new URL(
    `./${unique(loader.matchAll(WORKER_PATTERN), loaderUrl.href)}`,
    loaderUrl,
  );
  const worker = await getScript(workerUrl);
  if (!worker.includes("kuromoji@")) {
    throw new Error(
      `${workerUrl.href} が textlint の Worker ではないようです（kuromoji を含まない）`,
    );
  }
}

const [pageUrl] = process.argv.slice(2);
if (!pageUrl) {
  console.error("使い方: node scripts/smoke-test-worker.ts <配信先の URL>");
  process.exit(2);
}
main(pageUrl).catch((error: unknown) => {
  console.log(`::error::${error instanceof Error ? error.message : String(error)}`);
  process.exit(1);
});
