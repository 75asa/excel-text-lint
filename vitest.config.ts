import { defineConfig } from "vitest/config";

// vite.config.ts を読み込ませないために、Vitest 用の設定を分けている。
// vite.config.ts は serve のときに office-addin-dev-certs で開発用証明書を作成・信頼登録するため、
// テストの実行で証明書が作られたり、キーチェーンに CA が入ったりしないようにする。
export default defineConfig({
  test: {
    include: ["src/**/*.test.ts"],
    environment: "node",
    // vi.stubGlobal で差し込んだ Excel / Office のスタブを、テストごとに元に戻す。
    unstubGlobals: true,
  },
});
