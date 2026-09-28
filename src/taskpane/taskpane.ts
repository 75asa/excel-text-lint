import { LintEngine } from "../core/engine";
import { ExcelAdapter } from "../hosts/excel/adapter";
import { mountTaskpane } from "../ui/app";
import { applyTheme, themeFromOffice } from "../ui/theme";

// Office.js の初期化と、ホストのアダプタ・lint エンジンの用意だけを行う。UI は src/ui/ にある（#18）。

const root = document.getElementById("app") as HTMLElement;

Office.onReady(({ host }) => {
  followOfficeTheme();
  if (host !== Office.HostType.Excel) {
    root.textContent = `未対応のホストです: ${host ?? "（Office の外で開かれています）"}`;
    return;
  }
  mountTaskpane({
    root,
    adapter: new ExcelAdapter(),
    engine: new LintEngine({
      workerUrl: new URL(`${import.meta.env.BASE_URL}textlint/loader.js`, location.href),
      dictBaseUrl: dictBaseUrl(),
    }),
    title: "excel-text-lint",
  });
});

/**
 * Office のテーマ（背景色）に合わせてライト / ダークを切り替える。
 *
 * `Office.context.officeTheme` が取れない環境（Office on the web の一部など）では、
 * CSS の prefers-color-scheme に任せる。
 */
function followOfficeTheme(): void {
  const update = () => applyTheme(document.documentElement, themeFromOffice(officeTheme()));
  update();
  try {
    Office.context.document?.addHandlerAsync?.(Office.EventType.OfficeThemeChanged, update);
  } catch {
    // テーマの変更イベントに対応していないホスト
  }
}

function officeTheme(): { bodyBackgroundColor?: string } | undefined {
  try {
    return Office.context.officeTheme;
  } catch {
    return undefined;
  }
}

/**
 * kuromoji の辞書の配信先。
 *
 * - `VITE_TEXTLINT_DICT_BASE_URL` があればそれを使う
 * - build した成果物では `<base>/dict/`（build 時に kuromoji の辞書を dist/dict/ にコピーする。#7）
 * - dev サーバーでは /dict/ を配信していないので、jsdelivr（textlint の既定）から取得する
 *
 * Worker からの相対パスでは base を解決できないので、ここで絶対 URL にして渡す。
 */
function dictBaseUrl(): URL | undefined {
  const override = import.meta.env.VITE_TEXTLINT_DICT_BASE_URL;
  if (override) return new URL(override, location.href);
  if (import.meta.env.PROD) return new URL(`${import.meta.env.BASE_URL}dict/`, location.href);
  return undefined;
}
