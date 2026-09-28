import { resolve } from "node:path";
import { defineConfig } from "vitest/config";

// vite.config.ts を読み込ませないために、Vitest 用の設定を分けている。
// vite.config.ts は serve のときに office-addin-dev-certs で開発用証明書を作成・信頼登録するため、
// テストの実行で証明書が作られたり、キーチェーンに CA が入ったりしないようにする。
export default defineConfig({
  resolve: {
    alias: {
      // scripts/textlint-worker.ts のプラグインは使わない（dev サーバー用の処理で Worker を生成してしまうため）
      "virtual:textlint-worker": resolve(
        import.meta.dirname,
        "src/testing/textlint-worker-stub.ts",
      ),
    },
  },
  test: {
    include: ["src/**/*.test.ts", "scripts/**/*.test.ts"],
    environment: "node",
    // vi.stubGlobal で差し込んだ Excel / Office のスタブを、テストごとに元に戻す。
    unstubGlobals: true,
  },
});
