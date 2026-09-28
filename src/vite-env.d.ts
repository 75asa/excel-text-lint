/// <reference types="vite/client" />

interface ImportMetaEnv {
  /** kuromoji の辞書の配信先（`base.dat.gz` などが並ぶディレクトリの URL）。未設定なら jsdelivr から取得する。 */
  readonly VITE_TEXTLINT_DICT_BASE_URL?: string;
}

/** ローダーのパス（base からの相対）。build ではファイル名にハッシュが付く（scripts/textlint-worker.ts）。 */
declare module "virtual:textlint-worker" {
  export const textlintLoaderPath: string;
}
