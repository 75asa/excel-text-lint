// dist の worker を静的配信し、ヘッドレス Chrome で初期化時間・lint 時間・辞書の転送量を計測する
import { createServer } from "node:http";
import { readFile, cp } from "node:fs/promises";
import { extname, join } from "node:path";
import { chromium } from "playwright-core";

const ROOT = join(import.meta.dirname, "public");
await cp(join(import.meta.dirname, "dist/textlint-worker.js"), join(ROOT, "textlint-worker.js"));
await cp(join(import.meta.dirname, "node_modules/kuromoji/dict"), join(ROOT, "dict"), { recursive: true });

const TYPES = { ".html": "text/html", ".js": "text/javascript", ".gz": "application/octet-stream" };
const server = createServer(async (req, res) => {
  try {
    const body = await readFile(join(ROOT, decodeURIComponent(new URL(req.url, "http://x").pathname)));
    res.writeHead(200, { "content-type": TYPES[extname(req.url)] ?? "application/octet-stream" }).end(body);
  } catch {
    res.writeHead(404).end();
  }
}).listen(0);
const base = `http://localhost:${server.address().port}`;

const executablePath = process.env.CHROME ?? "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const cells = Number(process.env.CELLS ?? 1000);
const browser = await chromium.launch({ executablePath, headless: true });
try {
  for (const [mode, workerUrl] of [["cdn", "textlint-worker.js"], ["self-host", "self-host-worker.js"]]) {
    const ctx = await browser.newContext(); // キャッシュを分けるためにモードごとに作る
    const page = await ctx.newPage();
    const dict = { bytes: 0, hosts: new Set() };
    page.on("requestfinished", async (r) => {
      if (!r.url().includes("dict")) return;
      dict.hosts.add(new URL(r.url()).host);
      dict.bytes += (await r.sizes()).responseBodySize;
    });
    await page.goto(`${base}/index.html`);
    await page.waitForFunction(() => window.benchReady);
    const result = await page.evaluate((a) => window.runBench(a), { workerUrl, cells });
    console.log(JSON.stringify({ mode, cells, dictBytes: dict.bytes, dictHosts: [...dict.hosts], ...result }, null, 2));
    await ctx.close();
  }
} finally {
  await browser.close();
  server.close();
}
