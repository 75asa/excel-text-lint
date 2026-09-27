// configs/*.json のルールセットごとに worker をビルドし、ヘッドレス Chrome で corpus.mjs の各セルを lint して比較する
//   CONFIGS=a,b  対象を絞る（既定は configs/ のすべてと .textlintrc.json）
//   REBUILD=1    ビルド済みでも作り直す
// 結果は results/compare.json に書き、要約を標準出力に出す
import { createServer } from "node:http";
import { readFile, readdir, cp, mkdir, writeFile, access } from "node:fs/promises";
import { execFileSync } from "node:child_process";
import { gzipSync } from "node:zlib";
import { basename, extname, join } from "node:path";
import { chromium } from "playwright-core";
import { CORPUS } from "./corpus.mjs";

const DIR = import.meta.dirname;
const PUBLIC = join(DIR, "public");
await cp(join(DIR, "node_modules/kuromoji/dict"), join(PUBLIC, "dict"), { recursive: true });

const all = [
  ...(await readdir(join(DIR, "configs"))).filter((f) => f.endsWith(".json")).map((f) => ({ name: basename(f, ".json"), rc: join("configs", f) })),
  { name: "recommended", rc: ".textlintrc.json" },
];
const only = process.env.CONFIGS?.split(",");
const configs = only ? all.filter((c) => only.includes(c.name)) : all;

// ビルドとサイズ計測
for (const c of configs) {
  const out = join("dist", c.name);
  const file = join(DIR, out, "textlint-worker.js");
  const exists = await access(file).then(() => true, () => false);
  if (!exists || process.env.REBUILD) {
    try {
      execFileSync("npm", ["run", "build"], { cwd: DIR, env: { ...process.env, TEXTLINTRC: c.rc, OUT_DIR: out }, stdio: "pipe" });
    } catch (e) {
      c.buildError = String(e.stderr ?? e).slice(0, 500);
      continue;
    }
  }
  const buf = await readFile(file);
  c.bytes = buf.length;
  c.gzipBytes = gzipSync(buf).length;
}

const TYPES = { ".html": "text/html", ".js": "text/javascript", ".mjs": "text/javascript" };
const server = createServer(async (req, res) => {
  const path = decodeURIComponent(new URL(req.url, "http://x").pathname);
  try {
    const body = await readFile(path.startsWith("/dist/") ? join(DIR, path) : join(PUBLIC, path));
    res.writeHead(200, { "content-type": TYPES[extname(path)] ?? "application/octet-stream" }).end(body);
  } catch {
    res.writeHead(404).end();
  }
}).listen(0);
const base = `http://localhost:${server.address().port}`;

const executablePath = process.env.CHROME ?? "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const browser = await chromium.launch({ executablePath, headless: true });
const ctx = await browser.newContext(); // 辞書の IndexedDB キャッシュを共有する
try {
  for (const c of configs) {
    if (c.buildError) continue;
    const page = await ctx.newPage();
    const pageErrors = [];
    page.on("console", (m) => m.type() === "error" && pageErrors.push(m.text()));
    await page.goto(`${base}/index.html`);
    await page.waitForFunction(() => window.benchReady);
    const r = await page.evaluate(async ({ workerUrl, corpus }) => {
      const timeout = (p, ms) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej("timeout"), ms))]);
      const w = window.makeWorker(workerUrl);
      const out = { cells: {} };
      try {
        await timeout(w.ready, 30000);
        const t = performance.now();
        for (const { id, text } of corpus) {
          const res = await timeout(w.lint(text), 30000);
          const fixed = await timeout(w.fix(text), 30000);
          out.cells[id] = {
            messages: res.messages.map((m) => ({ ruleId: m.ruleId, message: m.message, fixable: Boolean(m.fix) })),
            fixed: fixed.output,
          };
        }
        out.lintMs = performance.now() - t;
      } catch (e) {
        out.error = e?.message ? [e.message, e.cause?.message].filter(Boolean).join(": ") : String(e);
      } finally {
        w.terminate();
      }
      return out;
    }, { workerUrl: `self-host-worker.js?worker=/dist/${c.name}/textlint-worker.js`, corpus: CORPUS });
    Object.assign(c, r, pageErrors.length ? { pageErrors } : {});
    await page.close();
  }
} finally {
  await browser.close();
  server.close();
}

await mkdir(join(DIR, "results"), { recursive: true });
await writeFile(join(DIR, "results/compare.json"), JSON.stringify({ corpus: CORPUS, configs }, null, 2));

// 要約: 設定ごとのサイズ・エラー、セルごとの検出ルール
const kb = (n) => (n / 1024).toFixed(0) + " KB";
for (const c of configs) {
  console.log(`\n## ${c.name}  size=${c.bytes ? kb(c.bytes) : "-"} gzip=${c.gzipBytes ? kb(c.gzipBytes) : "-"}${c.buildError ? " BUILD ERROR: " + c.buildError : ""}${c.error ? " RUNTIME ERROR: " + c.error : ""}`);
  if (!c.cells) continue;
  for (const { id, kind } of CORPUS) {
    const ms = c.cells[id]?.messages ?? [];
    if (!ms.length) continue;
    const rules = ms.map((m) => m.ruleId.replace(/^[^/]+\//, "") + (m.fixable ? "*" : ""));
    console.log(`${id}(${kind}) ${rules.join(", ")}`);
  }
}
