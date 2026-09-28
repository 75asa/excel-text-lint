/**
 * テーマ（ライト / ダーク）の判定。
 *
 * 既定は CSS の `prefers-color-scheme` に任せる。Office のテーマ（`Office.context.officeTheme`）が取れるときは、
 * その背景色から判定して `<html data-theme="…">` で上書きする（Office のテーマと OS の設定が違うことがあるため）。
 */

export type ThemeName = "light" | "dark";

/** `#rgb` / `#rrggbb` の色を、相対輝度で暗い色か判定する。解釈できなければ null。 */
export function isDarkColor(color: string): boolean | null {
  const match = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(color.trim());
  let hex = match?.[1];
  if (!hex) return null;
  if (hex.length === 3) hex = [...hex].map((c) => c + c).join("");
  const [r, g, b] = [0, 2, 4].map((i) => Number.parseInt(hex.slice(i, i + 2), 16) / 255) as [
    number,
    number,
    number,
  ];
  const linear = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
  const luminance = 0.2126 * linear(r) + 0.7152 * linear(g) + 0.0722 * linear(b);
  // 白との対比と黒との対比が等しくなる輝度（約 0.179）を境にする
  return luminance < 0.179;
}

/** Office のテーマ（の一部）からライト / ダークを決める。判断できなければ null（CSS に任せる）。 */
export function themeFromOffice(
  theme: { bodyBackgroundColor?: string } | null | undefined,
): ThemeName | null {
  const background = theme?.bodyBackgroundColor;
  if (!background) return null;
  const dark = isDarkColor(background);
  return dark === null ? null : dark ? "dark" : "light";
}

/** `<html data-theme>` を設定する。null なら外して CSS（prefers-color-scheme）に任せる。 */
export function applyTheme(root: HTMLElement, theme: ThemeName | null): void {
  if (theme) root.dataset.theme = theme;
  else delete root.dataset.theme;
}
