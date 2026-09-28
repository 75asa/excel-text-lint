/** DOM を組み立てる小さなヘルパー。 */

export type Child = Node | string | null | undefined | false;

type Handler = (event: Event) => void;

export type Attrs = Record<string, string | number | boolean | undefined | Handler>;

/**
 * 要素を作る。
 *
 * - `class` は className、`onclick` などの `on` で始まる関数はイベントリスナーにする
 * - 値が `false` / `undefined` の属性は付けない。`true` は空の値で付ける
 */
export function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Attrs = {},
  ...children: Child[]
): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  for (const [name, value] of Object.entries(attrs)) {
    if (value === undefined || value === false) continue;
    if (typeof value === "function") el.addEventListener(name.slice(2), value);
    else if (name === "class") el.className = String(value);
    else el.setAttribute(name, value === true ? "" : String(value));
  }
  append(el, children);
  return el;
}

export function append(parent: Element, children: readonly Child[]): void {
  for (const child of children) {
    if (child === null || child === undefined || child === false) continue;
    parent.append(child);
  }
}

/** Fluent System Icons 風の、線だけの小さなアイコン（20px のグリッド）。 */
export function icon(name: "chevron-left" | "chevron-right" | "settings" | "close"): SVGElement {
  const paths: Record<typeof name, string> = {
    "chevron-left": "M12.5 4.5 7 10l5.5 5.5",
    "chevron-right": "M7.5 4.5 13 10l-5.5 5.5",
    settings:
      "M10 7a3 3 0 1 0 0 6 3 3 0 0 0 0-6Zm6.5 3-1.6-.6a5 5 0 0 0-.5-1.2l.7-1.6-1.4-1.4-1.6.7a5 5 0 0 0-1.2-.5L10.3 3.5H9.7L9 5.1a5 5 0 0 0-1.2.5l-1.6-.7-1.4 1.4.7 1.6a5 5 0 0 0-.5 1.2L3.5 9.7v.6l1.6.7a5 5 0 0 0 .5 1.2l-.7 1.6 1.4 1.4 1.6-.7a5 5 0 0 0 1.2.5l.6 1.6h.6l.7-1.6a5 5 0 0 0 1.2-.5l1.6.7 1.4-1.4-.7-1.6a5 5 0 0 0 .5-1.2l1.6-.7Z",
    close: "M5 5l10 10M15 5 5 15",
  };
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 20 20");
  svg.setAttribute("width", "16");
  svg.setAttribute("height", "16");
  svg.setAttribute("aria-hidden", "true");
  svg.classList.add("icon");
  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", paths[name]);
  svg.append(path);
  return svg;
}

/** クリップボードにコピーする。Office のタスクペインでは Clipboard API が使えないことがあるので、execCommand で補う。 */
export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    const area = h("textarea", { class: "visually-hidden", readonly: true });
    area.value = text;
    document.body.append(area);
    area.select();
    try {
      return document.execCommand("copy");
    } catch {
      return false;
    } finally {
      area.remove();
    }
  }
}

/** 子要素をすべて置き換える（`null` などは飛ばす）。 */
export function replace(parent: Element, ...children: Child[]): void {
  parent.replaceChildren();
  append(parent, children);
}

export interface ConfirmOptions {
  title: string;
  /** 本文。1 要素が 1 段落になる。 */
  body: readonly string[];
  confirmLabel: string;
  cancelLabel?: string;
}

/**
 * 確認ダイアログを出し、確定なら true を返す。
 *
 * Office のアドインでは `window.confirm` が使えないので、ペインの中にモーダルを描く。
 * Escape とキャンセルで false。フォーカスはダイアログの中に閉じ込め、閉じたら元の要素に戻す。
 */
export function confirmDialog(host: HTMLElement, options: ConfirmOptions): Promise<boolean> {
  return new Promise((resolve) => {
    const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    // 後ろの要素は操作できないようにする（支援技術からも隠す）
    const background = [...host.children].filter(
      (el): el is HTMLElement => el instanceof HTMLElement && !el.inert,
    );
    for (const el of background) el.inert = true;
    const close = (result: boolean) => {
      for (const el of background) el.inert = false;
      overlay.remove();
      previous?.focus();
      resolve(result);
    };
    const cancel = h(
      "button",
      { type: "button", class: "button", onclick: () => close(false) },
      options.cancelLabel ?? "キャンセル",
    );
    const confirm = h(
      "button",
      { type: "button", class: "button primary", onclick: () => close(true) },
      options.confirmLabel,
    );
    const titleId = `dialog-title-${Math.random().toString(36).slice(2)}`;
    const dialog = h(
      "div",
      { class: "dialog", role: "alertdialog", "aria-modal": "true", "aria-labelledby": titleId },
      h("h2", { id: titleId }, options.title),
      ...options.body.map((text) => h("p", {}, text)),
      h("div", { class: "dialog-actions" }, cancel, confirm),
    );
    const overlay = h("div", { class: "dialog-overlay" }, dialog);
    overlay.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        close(false);
      } else if (event.key === "Tab") {
        // ボタンが 2 つだけなので、行き来させる
        event.preventDefault();
        (document.activeElement === cancel ? confirm : cancel).focus();
      }
    });
    host.append(overlay);
    cancel.focus();
  });
}
