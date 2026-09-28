// textlint の Worker（ローダーと、script-compiler が生成した textlint-worker.js）を配信する Vite プラグイン。
// 方針は docs/hosting.md の「Worker のキャッシュ対策」を参照（#45）。
//
// - build: ファイル名にコンテンツハッシュを付けて dist/textlint/ に出力する。
//   ローダーの中の "./textlint-worker.js" もハッシュ付きの名前に書き換える
// - dev: ハッシュを付けずに /textlint/*.js で配信する。Worker が無いか古ければ、起動時に生成する
// - タスクペインは仮想モジュール `virtual:textlint-worker` からローダーのパスを受け取る
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { readdir, readFile, stat } from "node:fs/promises";
import { createRequire } from "node:module";
import { join, resolve } from "node:path";
import type { Plugin } from "vite";

export const VIRTUAL_MODULE_ID = "virtual:textlint-worker";
const RESOLVED_VIRTUAL_MODULE_ID = `\0${VIRTUAL_MODULE_ID}`;

/** ローダーと生成した Worker を置くディレクトリ（プロジェクトのルートからの相対パス） */
export const WORKER_SRC_DIR = "src/textlint";
export const LOADER_FILE = "loader.js";
export const WORKER_FILE = "textlint-worker.js";
/** 配信するときのディレクトリ（base からの相対パス） */
const OUT_DIR = "textlint";

/** ローダーの中で Worker を指している文字列リテラル。build でハッシュ付きの名前に置き換える */
const WORKER_REFERENCE = `"./${WORKER_FILE}"`;

/** Worker の生成に使う入力。これらが Worker より新しければ、dev の起動時に作り直す */
const WORKER_INPUTS = [".textlintrc.json", "package-lock.json", "prh"];

export interface HashedAsset {
  /** base からの相対パス（例: `textlint/loader-0123abcd.js`） */
  fileName: string;
  source: string;
}

/** 内容から 8 桁のハッシュ（SHA-256 の先頭、16 進数）を作る */
export function contentHash(source: string): string {
  return createHash("sha256").update(source).digest("hex").slice(0, 8);
}

/**
 * ローダーと Worker に、内容のハッシュを付けた名前を付ける。
 * ローダーは Worker の名前を埋め込んだ後の内容でハッシュを取るので、Worker が変わればローダーの名前も変わる。
 */
export function hashTextlintWorker(
  loaderSource: string,
  workerSource: string,
): { loader: HashedAsset; worker: HashedAsset } {
  const workerName = `textlint-worker-${contentHash(workerSource)}.js`;
  const count = loaderSource.split(WORKER_REFERENCE).length - 1;
  if (count !== 1) {
    throw new Error(
      `${LOADER_FILE} の中に ${WORKER_REFERENCE} がちょうど 1 つ必要です（${count} 個ありました）`,
    );
  }
  const loader = loaderSource.replace(WORKER_REFERENCE, `"./${workerName}"`);
  return {
    loader: { fileName: `${OUT_DIR}/loader-${contentHash(loader)}.js`, source: loader },
    worker: { fileName: `${OUT_DIR}/${workerName}`, source: workerSource },
  };
}

/**
 * ローダーが書き換える辞書の URL（KUROMOJI_CDN）が、Worker の中の URL と、
 * dist/dict にコピーする kuromoji のバージョンに一致しているかを確かめる。
 * 一致しないと、辞書の URL が書き換わらず jsdelivr から取得してしまう（社内ネットワークでは失敗する）。
 */
export function checkKuromojiDictUrl(
  loaderSource: string,
  workerSource: string,
  kuromojiVersion: string,
): string {
  const cdn = loaderSource.match(/const KUROMOJI_CDN = "([^"]+)";/)?.[1];
  if (!cdn) throw new Error(`${LOADER_FILE} に KUROMOJI_CDN が見つかりません`);
  const expected = `https://cdn.jsdelivr.net/npm/kuromoji@${kuromojiVersion}/dict`;
  if (cdn !== expected) {
    throw new Error(
      `${LOADER_FILE} の KUROMOJI_CDN（${cdn}）が、インストールされている kuromoji（${kuromojiVersion}）と一致しません`,
    );
  }
  if (!workerSource.includes(cdn)) {
    throw new Error(
      `${WORKER_FILE} が ${cdn} から辞書を取得していません。textlint 側の kuromoji のバージョンが変わった可能性があります（docs/hosting.md の「辞書の更新」を参照）`,
    );
  }
  return cdn;
}

/** 仮想モジュール `virtual:textlint-worker` のコード */
export function virtualModuleCode(loaderPath: string): string {
  return `export const textlintLoaderPath = ${JSON.stringify(loaderPath)};\n`;
}

/** Worker が無いか、入力（WORKER_INPUTS）のどれかより古ければ true */
export async function isWorkerStale(projectRoot: string): Promise<boolean> {
  const workerMtime = await mtimeOf(resolve(projectRoot, WORKER_SRC_DIR, WORKER_FILE));
  if (workerMtime === undefined) return true;
  for (const input of WORKER_INPUTS) {
    if ((await newestMtime(resolve(projectRoot, input))) > workerMtime) return true;
  }
  return false;
}

async function mtimeOf(path: string): Promise<number | undefined> {
  try {
    return (await stat(path)).mtimeMs;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined;
    throw error;
  }
}

/** ファイルならその mtime、ディレクトリなら中のファイルの最新の mtime。無ければ 0 */
async function newestMtime(path: string): Promise<number> {
  let info: Awaited<ReturnType<typeof stat>>;
  try {
    info = await stat(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return 0;
    throw error;
  }
  if (!info.isDirectory()) return info.mtimeMs;
  let newest = 0;
  for (const entry of await readdir(path)) {
    newest = Math.max(newest, await newestMtime(join(path, entry)));
  }
  return newest;
}

/** `npm run build:worker` を実行する */
function generateWorker(projectRoot: string): Promise<void> {
  return new Promise((resolvePromise, reject) => {
    const child = spawn("npm", ["run", "build:worker"], {
      cwd: projectRoot,
      stdio: "inherit",
      shell: process.platform === "win32",
    });
    child.on("error", reject);
    child.on("exit", (code) =>
      code === 0
        ? resolvePromise()
        : reject(new Error(`npm run build:worker が終了コード ${code} で失敗しました`)),
    );
  });
}

function kuromojiVersion(): string {
  const require = createRequire(import.meta.url);
  const pkg = require("kuromoji/package.json") as { version: string };
  return pkg.version;
}

export function textlintWorker(options: { projectRoot: string }): Plugin {
  const srcDir = resolve(options.projectRoot, WORKER_SRC_DIR);
  const read = (name: string) => readFile(join(srcDir, name), "utf8");
  let command: "build" | "serve" = "serve";
  let assets: { loader: HashedAsset; worker: HashedAsset } | undefined;

  return {
    name: "excel-text-lint:textlint-worker",

    configResolved(config) {
      command = config.command;
    },

    async buildStart() {
      if (command !== "build") return;
      let workerSource: string;
      try {
        workerSource = await read(WORKER_FILE);
      } catch {
        this.error(
          `${WORKER_SRC_DIR}/${WORKER_FILE} がありません。先に npm run build:worker を実行してください（npm run build は自動で実行します）`,
        );
      }
      if (await isWorkerStale(options.projectRoot)) {
        this.warn(
          `${WORKER_SRC_DIR}/${WORKER_FILE} が .textlintrc.json などより古いです。npm run build:worker をやり直してください`,
        );
      }
      const loaderSource = await read(LOADER_FILE);
      checkKuromojiDictUrl(loaderSource, workerSource, kuromojiVersion());
      assets = hashTextlintWorker(loaderSource, workerSource);
    },

    resolveId(id) {
      return id === VIRTUAL_MODULE_ID ? RESOLVED_VIRTUAL_MODULE_ID : undefined;
    },

    load(id) {
      if (id !== RESOLVED_VIRTUAL_MODULE_ID) return undefined;
      // dev ではハッシュを付けない（configureServer のミドルウェアがそのまま返す）
      return virtualModuleCode(assets?.loader.fileName ?? `${OUT_DIR}/${LOADER_FILE}`);
    },

    generateBundle() {
      if (!assets) return;
      for (const asset of [assets.worker, assets.loader]) {
        this.emitFile({ type: "asset", fileName: asset.fileName, source: asset.source });
      }
      console.log(`[textlint-worker] ${assets.loader.fileName}, ${assets.worker.fileName}`);
    },

    async configureServer(server) {
      const { logger } = server.config;
      if (await isWorkerStale(options.projectRoot)) {
        logger.info(
          `[textlint-worker] ${WORKER_SRC_DIR}/${WORKER_FILE} が無いか古いので、npm run build:worker を実行します`,
        );
        try {
          await generateWorker(options.projectRoot);
        } catch (error) {
          logger.error(
            `[textlint-worker] Worker を生成できませんでした。lint は動きません: ${String(error)}`,
          );
        }
      }

      // Vite の変換を通さず、そのまま返す（Worker は classic worker で、importScripts で読み込むため）
      const routes = new Map(
        [LOADER_FILE, WORKER_FILE].map((name) => [`${server.config.base}${OUT_DIR}/${name}`, name]),
      );
      server.middlewares.use(async (req, res, next) => {
        const name = routes.get((req.url ?? "").split("?")[0] ?? "");
        if (!name) return next();
        try {
          const body = await read(name);
          res.setHeader("Content-Type", "text/javascript; charset=utf-8");
          res.setHeader("Cache-Control", "no-cache");
          res.end(body);
        } catch (error) {
          logger.error(`[textlint-worker] ${name} を読み込めません: ${String(error)}`);
          res.statusCode = 404;
          res.end();
        }
      });
    },
  };
}
