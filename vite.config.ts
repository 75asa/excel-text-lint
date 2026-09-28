import { resolve } from "node:path";
import { defineConfig } from "vite";
import { hostingAssets, prodBaseUrl } from "./scripts/hosting.ts";

const projectRoot = import.meta.dirname;
const root = resolve(projectRoot, "src");

export default defineConfig(async ({ command }) => ({
  root,
  // 本番（GitHub Pages）はリポジトリのサブパス（/excel-text-lint/）で配信される。docs/hosting.md を参照
  base: command === "build" ? prodBaseUrl().pathname : "/",
  publicDir: resolve(projectRoot, "public"),
  plugins: [hostingAssets({ projectRoot })],
  build: {
    outDir: resolve(projectRoot, "dist"),
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: resolve(root, "taskpane/taskpane.html"),
      },
    },
  },
  server: {
    port: 3000,
    strictPort: true,
    fs: {
      // Vite 既定の deny には "**/.git/**" が含まれ、git worktree を .git/wt/ の下に
      // 置いている場合（git-wt の wt.basedir = .git/wt）にすべてのファイルが 403 になる。
      // パターンをプロジェクトのルートに固定して、祖先ディレクトリの .git には反応しないようにする。
      deny: [".env", ".env.*", "*.{crt,pem}", `${projectRoot}/**/.git/**`],
    },
    // Office アドインは HTTPS 必須。開発用証明書は office-addin-dev-certs が
    // ~/.office-addin-dev-certs に作成・信頼登録する（初回はパスワードを求められる）。
    // build 時には証明書が要らないので、serve のときだけ読み込む。
    https: command === "serve" ? await getHttpsOptions() : undefined,
  },
}));

async function getHttpsOptions() {
  const { getHttpsServerOptions } = await import("office-addin-dev-certs");
  const { key, cert, ca } = await getHttpsServerOptions();
  return { key, cert, ca };
}
