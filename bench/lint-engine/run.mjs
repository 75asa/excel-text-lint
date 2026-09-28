// LintEngine の lint の時間を、ヘッドレス Chrome で測る（#17）。
//
// 事前に:
//   npm run build:worker                        # public/textlint/textlint-worker.js を作る
//   npm i --no-save --ignore-scripts playwright-core
// 実行:
//   node bench/lint-engine/run.mjs              # CELLS=10000,50000 MAX_IN_FLIGHT=1,8,32 CHROME=<path> で変えられる
//
// 静的ファイルを配るだけの http サーバーを立てる（Vite の dev サーバーや開発用証明書は使わない）。
import { mkdir, readFile } from "node:fs/promises";
import { createServer } from "node:http";
import { extname, join } from "node:path";
import { chromium } from "playwright-core";
import { build } from "rolldown";

const HERE = import.meta.dirname;
const ROOT = join(HERE, "../..");
const OUT = join(HERE, ".out");

await mkdir(OUT, { recursive: true });
await build({
  input: join(HERE, "page.ts"),
  output: { file: join(OUT, "page.js"), format: "esm" },
  logLevel: "warn",
});

const ROUTES = [
  ["/textlint/", join(ROOT, "public/textlint")],
  ["/dict/", join(ROOT, "node_modules/kuromoji/dict")],
  ["/page.js", join(OUT, "page.js")],
  ["/", join(HERE, "index.html")],
];
const TYPES = { ".html": "text/html", ".js": "text/javascript" };
const server = createServer(async (req, res) => {
  const path = decodeURIComponent(new URL(req.url, "http://x").pathname);
  const [prefix, target] = ROUTES.find(([p]) => path.startsWith(p));
  const file =
    prefix.endsWith("/") && prefix !== "/" ? join(target, path.slice(prefix.length)) : target;
  try {
    const body = await readFile(file);
    res.writeHead(200, { "content-type": TYPES[extname(file)] ?? "application/octet-stream" });
    res.end(body);
  } catch {
    res.writeHead(404).end();
  }
}).listen(0);
const base = `http://localhost:${server.address().port}`;

const executablePath =
  process.env.CHROME ?? "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const cellsList = (process.env.CELLS ?? "10000,50000").split(",").map(Number);
const inFlightList = (process.env.MAX_IN_FLIGHT ?? "1,8,32").split(",").map(Number);

const browser = await chromium.launch({ executablePath, headless: true });
try {
  const page = await browser.newPage();
  page.on("pageerror", (error) => console.error(error));
  await page.goto(`${base}/`);
  await page.waitForFunction(() => globalThis.benchReady);
  for (const cells of cellsList) {
    for (const maxInFlight of inFlightList) {
      const result = await page.evaluate((args) => globalThis.runBench(args), {
        cells,
        maxInFlight,
      });
      console.log(JSON.stringify(result));
    }
  }
} finally {
  await browser.close();
  server.close();
}
