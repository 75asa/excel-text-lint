/// <reference types="vite/client" />

interface ImportMetaEnv {
  /** kuromoji の辞書の配信先（`base.dat.gz` などが並ぶディレクトリの URL）。未設定なら jsdelivr から取得する。 */
  readonly VITE_TEXTLINT_DICT_BASE_URL?: string;
}
