/**
 * textlint の Worker（ローダー）の URL を組み立てる（#45）。
 *
 * ローダーのパスは Vite プラグイン（`scripts/textlint-worker.ts`）の仮想モジュールから受け取る。
 * build ではファイル名にコンテンツハッシュが付く（例: `textlint/loader-0123abcd.js`）ので、
 * ルールや Worker を変えてデプロイすると URL も変わり、古い Worker がキャッシュから使われることはない。
 * dev サーバーではハッシュを付けない（`textlint/loader.js`）。
 */
import { textlintLoaderPath } from "virtual:textlint-worker";

/**
 * ローダーの絶対 URL を返す。`LintEngine` の `workerUrl` に渡す。
 *
 * @param baseUrl Vite の `base`（既定は `import.meta.env.BASE_URL`）
 * @param href 相対な `baseUrl` の基準（既定は `location.href`）
 */
export function textlintWorkerUrl(
  baseUrl: string = import.meta.env.BASE_URL,
  href: string = globalThis.location.href,
): URL {
  const base = baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`;
  return new URL(`${base}${textlintLoaderPath}`, href);
}
