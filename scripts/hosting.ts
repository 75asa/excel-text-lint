// 本番の配信（GitHub Pages）に必要なファイルを build 時に dist へ出力する。
// 方針は docs/hosting.md を参照。
import { copyFile, mkdir, readdir, readFile, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { dirname, join, resolve } from "node:path";
import type { Plugin } from "vite";

/** 開発用 manifest（manifest.xml）が指す配信元 */
const DEV_ORIGIN = "https://localhost:3000/";

/**
 * 本番の配信先。末尾は `/`。カスタムドメインに移すときは環境変数 ADDIN_BASE_URL で上書きする
 * （例: ADDIN_BASE_URL=https://textlint.example.com/ npm run build）。
 */
const DEFAULT_PROD_BASE_URL = "https://75asa.github.io/excel-text-lint/";

/**
 * 本番用 manifest の Id。開発用（manifest.xml）とは別の Id にして、両方を同時に sideload できるようにする。
 * 組織展開した後に変えると別のアドインとして扱われるので、変更しないこと。
 */
const PROD_ADDIN_ID = "94cc4eba-3f73-4562-9444-1b201a9343c8";

/** kuromoji の辞書を置く dist 内のディレクトリ（`<base>/dict/*.dat.gz` で配信される） */
const DICT_DIR = "dict";

export function prodBaseUrl(): URL {
  const raw = process.env.ADDIN_BASE_URL || DEFAULT_PROD_BASE_URL;
  const url = new URL(raw.endsWith("/") ? raw : `${raw}/`);
  if (url.protocol !== "https:") {
    throw new Error(`ADDIN_BASE_URL は https である必要があります: ${raw}`);
  }
  return url;
}

/** 開発用 manifest から本番用 manifest を作る */
export function toProdManifest(devManifest: string, baseUrl: URL): string {
  const replaced = devManifest
    .replaceAll(DEV_ORIGIN, baseUrl.href)
    .replace(/<Id>[^<]*<\/Id>/, `<Id>${PROD_ADDIN_ID}</Id>`)
    .replace(/(<DisplayName DefaultValue="[^"]*?) \(dev\)"/, '$1"')
    .replace(
      /^(<\?xml[^>]*\?>\n)/,
      "$1<!-- 自動生成（scripts/hosting.ts）。直接編集せず manifest.xml を編集すること -->\n",
    );

  if (replaced.includes("localhost")) {
    throw new Error("本番用 manifest に localhost が残っています");
  }
  if (replaced.includes("(dev)")) {
    throw new Error("本番用 manifest の DisplayName に (dev) が残っています");
  }
  if (!replaced.includes(`<Id>${PROD_ADDIN_ID}</Id>`)) {
    throw new Error("本番用 manifest の Id を置き換えられませんでした");
  }
  return replaced;
}

/** node_modules/kuromoji/dict の *.dat.gz を outDir/dict にそのままコピーする */
async function copyKuromojiDict(outDir: string): Promise<number> {
  const require = createRequire(import.meta.url);
  const srcDir = join(dirname(require.resolve("kuromoji/package.json")), "dict");
  const destDir = join(outDir, DICT_DIR);
  const files = (await readdir(srcDir)).filter((name) => name.endsWith(".dat.gz"));
  if (files.length === 0) {
    throw new Error(`kuromoji の辞書が見つかりません: ${srcDir}`);
  }
  await mkdir(destDir, { recursive: true });
  // 圧縮したまま置く。kuromoji が自前で gunzip するので、展開も再圧縮もしない
  await Promise.all(files.map((name) => copyFile(join(srcDir, name), join(destDir, name))));
  return files.length;
}

/**
 * build のときだけ動く Vite プラグイン。
 * - kuromoji の辞書を dist/dict にコピーする
 * - manifest.xml から本番用の dist/manifest.prod.xml を生成する
 */
export function hostingAssets(options: { projectRoot: string }): Plugin {
  let outDir = "";
  return {
    name: "excel-text-lint:hosting-assets",
    apply: "build",
    configResolved(config) {
      outDir = resolve(config.root, config.build.outDir);
    },
    async closeBundle() {
      const dictCount = await copyKuromojiDict(outDir);

      const baseUrl = prodBaseUrl();
      const devManifest = await readFile(resolve(options.projectRoot, "manifest.xml"), "utf8");
      await writeFile(join(outDir, "manifest.prod.xml"), toProdManifest(devManifest, baseUrl));

      console.log(`[hosting-assets] dict: ${dictCount} files, manifest.prod.xml: ${baseUrl.href}`);
    },
  };
}
