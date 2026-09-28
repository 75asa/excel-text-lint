import { LintEngine } from "../core/engine";
import type { ExcelLocation } from "../core/locations";
import type { TextUnit, Violation } from "../core/types";
import { textlintWorkerUrl } from "../core/worker-url";
import { ExcelAdapter } from "../hosts/excel/adapter";

// 見た目は最低限。UI は #18 で作り込む。

const status = document.getElementById("status") as HTMLParagraphElement;
const lintButton = document.getElementById("lint-selection") as HTMLButtonElement;
const cancelButton = document.getElementById("cancel") as HTMLButtonElement;
const progress = document.getElementById("progress") as HTMLProgressElement;
const list = document.getElementById("violations") as HTMLOListElement;

const engine = new LintEngine({
  workerUrl: textlintWorkerUrl(),
  dictBaseUrl: dictBaseUrl(),
});
const adapter = new ExcelAdapter();
let controller: AbortController | undefined;

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel) {
    status.textContent = `未対応のホストです: ${host ?? "（Office の外で開かれています）"}`;
    return;
  }
  status.textContent = "セルを選択してボタンを押してください。";
  lintButton.disabled = false;
  lintButton.addEventListener("click", () => void lintSelection());
  cancelButton.addEventListener("click", () =>
    controller?.abort(new DOMException("中止しました", "AbortError")),
  );
  // Worker の起動を先に済ませておく（失敗しても lint 時にもう一度試す）
  engine.init().catch((error: unknown) => console.error(error));
});

async function lintSelection(): Promise<void> {
  controller = new AbortController();
  const { signal } = controller;
  lintButton.disabled = true;
  cancelButton.hidden = false;
  list.replaceChildren();
  try {
    status.textContent = "選択範囲を読み取っています…";
    const units = await adapter.collect("selection", { signal });
    if (units.length === 0) {
      status.textContent = "選択範囲に文字列のセルがありません。";
      return;
    }

    status.textContent = `${units.length} 件のセルを lint しています…（初回は辞書の読み込みに時間がかかります）`;
    progress.max = units.length;
    progress.value = 0;
    progress.hidden = false;
    const violations = await engine.lint(units, {
      signal,
      onProgress: ({ done }) => {
        progress.value = done;
      },
    });

    const byId = new Map(units.map((unit) => [unit.id, unit]));
    list.replaceChildren(
      ...violations.map((violation) => renderViolation(violation, byId.get(violation.unitId)!)),
    );
    status.textContent =
      violations.length === 0
        ? `${units.length} 件のセルに違反はありませんでした。`
        : `${units.length} 件のセルで ${violations.length} 件の違反が見つかりました。`;
  } catch (error) {
    if (signal.aborted) {
      status.textContent = "中止しました。";
    } else {
      console.error(error);
      status.textContent = `lint に失敗しました: ${error instanceof Error ? error.message : String(error)}`;
    }
  } finally {
    controller = undefined;
    lintButton.disabled = false;
    cancelButton.hidden = true;
    progress.hidden = true;
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

/** 違反の前後に付ける文脈の文字数。 */
const CONTEXT_CHARS = 12;

function renderViolation(violation: Violation, unit: TextUnit<ExcelLocation>): HTMLLIElement {
  const [start, end] = violation.displayRange;
  const text = unit.text;
  const before = text.slice(Math.max(0, start - CONTEXT_CHARS), start);
  const after = text.slice(end, end + CONTEXT_CHARS);

  const item = document.createElement("li");
  const address = element("span", "address", `${unit.location.sheet}!${unit.location.address}`);
  const context = element("span", "context");
  context.append(
    `${start > CONTEXT_CHARS ? "…" : ""}${before}`,
    element("mark", undefined, text.slice(start, end)),
    `${after}${end + CONTEXT_CHARS < text.length ? "…" : ""}`,
  );
  const message = element("span", "message", violation.message);
  const rule = element(
    "span",
    "rule",
    `${violation.ruleId}${violation.fix ? `（修正案: 「${violation.fix.text}」）` : ""}`,
  );
  item.append(address, context, message, rule);
  return item;
}

function element<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text !== undefined) el.textContent = text;
  return el;
}
