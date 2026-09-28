/**
 * タスクペインの UI（実行 / 結果一覧 / 詳細）。
 *
 * ホストには `HostAdapter` のインターフェースだけで依存する。reveal / applyFix / highlight / 範囲の選択は
 * `capabilities` を見て出し分けるので、アダプタの実装が増えればそのまま使えるようになる。
 * 表示用のロジックは `results.ts` / `excerpt.ts` / `format.ts` / `status.ts` / `actions.ts` の純関数に置き、
 * ここは DOM の組み立てとイベントの配線だけにする。
 */

import "./ui.css";
import type { LintOptions } from "../core/engine";
import type { Fix, HostAdapter, Scope, Severity, TextUnit, Violation } from "../core/types";
import {
  applyAllMessage,
  applyFixMessage,
  applyFixNote,
  errorMessage,
  type FixStatus,
  type ItemState,
  itemActions,
  type Message,
  markFixed,
  markStale,
  planFixes,
  revealMessage,
} from "./actions";
import { append, type Child, confirmDialog, copyText, h, icon, replace } from "./dom";
import { buildExcerpt, buildFixPreview, type Excerpt } from "./excerpt";
import {
  type HostTerms,
  hostTerms,
  SEVERITIES,
  scopeLabel,
  severityLabel,
  splitRuleId,
} from "./format";
import {
  buildItems,
  emptyFilter,
  filterItems,
  flattenGroups,
  type GroupBy,
  groupItems,
  isFiltered,
  positionOf,
  type ResultFilter,
  type ResultItem,
  type ResultSummary,
  stepKey,
  summarize,
  toggleSeverity,
} from "./results";
import { describePhase, type EngineState, type RunPhase } from "./status";

/** lint を実行するもの（`LintEngine` がこの形を満たす）。デモやテストでは差し替える。 */
export interface LintClient {
  init(): Promise<void>;
  lint(units: readonly TextUnit[], options?: LintOptions): Promise<Violation[]>;
}

export interface TaskpaneOptions<L> {
  /** UI を描く要素。中身は置き換える。 */
  root: HTMLElement;
  adapter: HostAdapter<L>;
  engine: LintClient;
  /** アプリの名前（見出し）。 */
  title?: string;
  /** 設定画面を開く（#11）。省略すると「準備中」と案内する。 */
  onOpenSettings?: () => void;
}

export interface TaskpaneApp {
  /** 範囲を指定して実行する（デモ・テスト用）。 */
  run(scope?: Scope): Promise<void>;
  dispose(): void;
}

/** これより多くの TextUnit をハイライトするときは確認する（条件付き書式などが同じ数だけ増えるため）。 */
const HIGHLIGHT_CONFIRM_THRESHOLD = 200;

/** 一度に描く行の数。大きなブックで DOM が重くならないように、残りは「さらに表示」で出す。 */
const PAGE_SIZE = 200;

export function mountTaskpane<L>(options: TaskpaneOptions<L>): TaskpaneApp {
  return new Taskpane(options);
}

class Taskpane<L> implements TaskpaneApp {
  readonly #adapter: HostAdapter<L>;
  readonly #engine: LintClient;
  readonly #terms: HostTerms;
  readonly #onOpenSettings: (() => void) | undefined;

  // 状態
  #scope: Scope;
  #phase: RunPhase = { kind: "idle" };
  #engineState: EngineState = "starting";
  /** 最初の lint が終わったか（辞書の読み込みが済んだか）。 */
  #warm = false;
  #controller: AbortController | undefined;
  #items: ResultItem<L>[] = [];
  #summary: ResultSummary = summarize([]);
  #states = new Map<string, ItemState>();
  #filter: ResultFilter = emptyFilter();
  #groupBy: GroupBy = "container";
  #collapsed = new Set<string>();
  #selected: string | null = null;
  #view: "list" | "detail" = "list";
  #limit = PAGE_SIZE;
  #highlightOn = false;
  /** 「すべて適用」の実行中。 */
  #fixing = false;
  #message: Message | null = null;
  #disposers: (() => void)[] = [];

  // 表示の順（前へ / 次へ）。renderResults で更新する
  #order: string[] = [];

  // 要素
  readonly #el: {
    root: HTMLElement;
    scope: HTMLElement;
    run: HTMLButtonElement;
    cancel: HTMLButtonElement;
    statusText: HTMLElement;
    statusHint: HTMLElement;
    progress: HTMLElement;
    progressBar: HTMLElement;
    message: HTMLElement;
    results: HTMLElement;
    summary: HTMLElement;
    severityChips: HTMLElement;
    groupBy: HTMLSelectElement;
    containerFilter: HTMLSelectElement;
    ruleFilter: HTMLSelectElement;
    bulk: HTMLElement;
    toolbar: HTMLElement;
    list: HTMLElement;
    detail: HTMLElement;
    nav: HTMLElement;
    navPosition: HTMLElement;
  };

  constructor(options: TaskpaneOptions<L>) {
    this.#adapter = options.adapter;
    this.#engine = options.engine;
    this.#terms = hostTerms(options.adapter.host);
    this.#onOpenSettings = options.onOpenSettings;
    this.#scope = options.adapter.capabilities.scopes[0] ?? "selection";
    this.#el = this.#build(options.root, options.title ?? "textlint");

    this.#renderScope();
    this.#renderStatus();

    // Worker の起動を先に済ませておく（失敗しても実行のときにもう一度試す）
    this.#engine.init().then(
      () => this.#setEngineState("ready"),
      (error: unknown) => {
        console.error(error);
        this.#setEngineState("failed");
      },
    );

    if (this.#adapter.capabilities.selectionTracking && this.#adapter.onSelectionChanged) {
      const subscription = this.#adapter.onSelectionChanged((unitId) =>
        this.#followSelection(unitId),
      );
      this.#disposers.push(() => subscription.dispose());
    }

    // ペインを閉じるときにハイライトを消す（書式が残ったままファイルが保存されないように）。
    // 閉じる途中なので完了は保証されない。確実に消すには「ハイライトを解除」を使う
    const onHide = () => {
      if (this.#highlightOn) void this.#clearHighlight();
    };
    window.addEventListener("pagehide", onHide);
    this.#disposers.push(() => window.removeEventListener("pagehide", onHide));
  }

  dispose(): void {
    this.#controller?.abort(new DOMException("閉じました", "AbortError"));
    for (const dispose of this.#disposers) dispose();
    this.#disposers = [];
  }

  // -------------------------------------------------------------------------
  // 骨組み
  // -------------------------------------------------------------------------

  #build(root: HTMLElement, title: string) {
    const run = h("button", { type: "button", class: "button primary", onclick: () => this.run() });
    const cancel = h(
      "button",
      { type: "button", class: "button", hidden: true, onclick: () => this.#cancel() },
      "中止",
    );
    const scope = h("div", { class: "segmented", role: "radiogroup", "aria-label": "範囲" });
    const statusText = h("p", { class: "status-text" });
    const statusHint = h("p", { class: "status-hint", hidden: true });
    const progressBar = h("div", { class: "progress-bar" });
    const progress = h(
      "div",
      { class: "progress", role: "progressbar", "aria-label": "進捗", hidden: true },
      progressBar,
    );
    const message = h("div", { class: "message-bar", hidden: true });

    const summary = h("div", { class: "summary" });
    const severityChips = h("div", {
      class: "chips",
      role: "group",
      "aria-label": "重大度で絞り込む",
    });
    const groupBy = select("まとめ方", [
      ["container", `${this.#terms.container}別`],
      ["rule", "ルール別"],
      ["none", "まとめない"],
    ]);
    groupBy.value = this.#groupBy;
    groupBy.addEventListener("change", () => {
      this.#groupBy = groupBy.value as GroupBy;
      this.#collapsed.clear();
      this.#renderResults();
    });
    const containerFilter = select(`${this.#terms.container}で絞り込む`, []);
    containerFilter.addEventListener("change", () => {
      this.#filter = { ...this.#filter, containerKey: containerFilter.value || null };
      this.#onFilterChanged();
    });
    const ruleFilter = select("ルールで絞り込む", []);
    ruleFilter.addEventListener("change", () => {
      this.#filter = { ...this.#filter, ruleId: ruleFilter.value || null };
      this.#onFilterChanged();
    });

    const toolbar = h(
      "div",
      { class: "toolbar" },
      field("まとめ方", groupBy),
      field(this.#terms.container, containerFilter),
      field("ルール", ruleFilter),
    );
    const bulk = h("div", { class: "bulk" });
    const list = h("div", { class: "list" });
    list.addEventListener("keydown", (event) => this.#onListKeydown(event));
    const results = h(
      "section",
      { class: "results", "aria-label": "結果", hidden: true },
      summary,
      severityChips,
      toolbar,
      bulk,
      list,
    );
    const detail = h("section", { class: "detail", "aria-label": "違反の詳細", hidden: true });
    detail.addEventListener("keydown", (event) => {
      if (event.key === "Escape") this.#showList();
    });

    const navPosition = h("span", { class: "nav-position", "aria-live": "polite" });
    const nav = h(
      "nav",
      { class: "nav", "aria-label": "違反の移動", hidden: true },
      h(
        "button",
        { type: "button", class: "button subtle", onclick: () => this.#step(-1) },
        icon("chevron-left"),
        "前へ",
      ),
      navPosition,
      h(
        "button",
        { type: "button", class: "button subtle", onclick: () => this.#step(1) },
        "次へ",
        icon("chevron-right"),
      ),
    );

    const header = h(
      "header",
      { class: "app-header" },
      h("h1", {}, title),
      h(
        "button",
        {
          type: "button",
          class: "button subtle icon-button",
          title: "設定",
          onclick: () => this.#openSettings(),
        },
        icon("settings"),
        h("span", {}, "設定"),
      ),
    );
    const runSection = h(
      "section",
      { class: "run", "aria-label": "実行" },
      scope,
      h("div", { class: "run-actions" }, run, cancel),
      h(
        "div",
        { class: "status", role: "status", "aria-live": "polite" },
        statusText,
        statusHint,
        progress,
      ),
    );

    root.classList.add("taskpane");
    root.replaceChildren(
      header,
      h("main", { class: "content" }, runSection, message, results, detail),
      nav,
    );

    return {
      root,
      scope,
      run,
      cancel,
      statusText,
      statusHint,
      progress,
      progressBar,
      message,
      results,
      summary,
      severityChips,
      groupBy,
      containerFilter,
      ruleFilter,
      bulk,
      toolbar,
      list,
      detail,
      nav,
      navPosition,
    };
  }

  // -------------------------------------------------------------------------
  // 実行
  // -------------------------------------------------------------------------

  async run(scope: Scope = this.#scope): Promise<void> {
    if (this.#controller) return;
    this.#scope = scope;
    const controller = new AbortController();
    this.#controller = controller;
    const { signal } = controller;
    this.#message = null;
    this.#renderMessage();

    // 前回のハイライトは消す（条件付き書式が残ったまま保存されないように）。再実行後は自動では付けない
    if (this.#highlightOn) await this.#clearHighlight();
    this.#setResults([], []);

    try {
      this.#setPhase({ kind: "collecting", scope, collected: 0 });
      const units = await this.#adapter.collect(scope, {
        signal,
        onProgress: ({ collected, total }) =>
          this.#setPhase(
            total === undefined
              ? { kind: "collecting", scope, collected }
              : { kind: "collecting", scope, collected, total },
          ),
      });
      signal.throwIfAborted();
      if (units.length === 0) {
        this.#setPhase({ kind: "empty", scope });
        return;
      }

      if (this.#engineState !== "ready") {
        this.#setPhase({ kind: "starting" });
        await this.#engine.init();
        this.#setEngineState("ready");
      }
      signal.throwIfAborted();

      this.#setPhase({
        kind: "linting",
        done: 0,
        total: units.length,
        violations: 0,
        warm: this.#warm,
      });
      const violations = await this.#engine.lint(units, {
        signal,
        onProgress: ({ done, total, violations }) => {
          if (done > 0) this.#warm = true;
          this.#setPhase({ kind: "linting", done, total, violations, warm: this.#warm });
        },
      });
      this.#warm = true;
      this.#setResults(units, violations);
      this.#setPhase({
        kind: "done",
        scope,
        units: units.length,
        violations: this.#items.length,
      });
    } catch (error) {
      if (signal.aborted) {
        this.#setPhase({ kind: "cancelled" });
      } else {
        console.error(error);
        if (this.#engineState !== "ready") this.#setEngineState("failed");
        this.#setPhase({ kind: "error", message: errorMessage(error) });
      }
    } finally {
      this.#controller = undefined;
      this.#renderStatus();
    }
  }

  #cancel(): void {
    this.#controller?.abort(new DOMException("中止しました", "AbortError"));
  }

  #setPhase(phase: RunPhase): void {
    this.#phase = phase;
    this.#renderStatus();
  }

  #setEngineState(state: EngineState): void {
    this.#engineState = state;
    this.#renderStatus();
  }

  #setResults(units: readonly TextUnit<L>[], violations: readonly Violation[]): void {
    this.#items = buildItems(units, violations);
    this.#summary = summarize(this.#items);
    this.#states = new Map();
    this.#filter = emptyFilter();
    this.#collapsed.clear();
    this.#selected = null;
    this.#view = "list";
    this.#limit = PAGE_SIZE;
    this.#renderFilters();
    this.#renderResults();
  }

  // -------------------------------------------------------------------------
  // 選択・移動
  // -------------------------------------------------------------------------

  #item(key: string | null): ResultItem<L> | undefined {
    return key === null ? undefined : this.#items.find((item) => item.key === key);
  }

  #select(key: string, options: { reveal: boolean; focus?: boolean }): void {
    this.#selected = key;
    const index = this.#order.indexOf(key);
    if (index >= this.#limit) this.#limit = Math.ceil((index + 1) / PAGE_SIZE) * PAGE_SIZE;
    const item = this.#item(key);
    if (item) this.#collapsed.delete(this.#groupKey(item));
    this.#renderResults();
    if (this.#view === "detail") this.#renderDetail();
    else this.#scrollToSelected(options.focus ?? false);
    if (options.reveal && item) void this.#reveal(item);
  }

  #step(direction: 1 | -1): void {
    const key = stepKey(this.#order, this.#selected, direction);
    if (key !== null) this.#select(key, { reveal: true, focus: this.#view === "list" });
  }

  #openDetail(key: string): void {
    this.#view = "detail";
    this.#select(key, { reveal: true });
    this.#el.results.hidden = true;
    this.#el.detail.hidden = false;
    this.#el.detail.querySelector<HTMLElement>(".back")?.focus();
  }

  #showList(): void {
    this.#view = "list";
    this.#el.detail.hidden = true;
    this.#el.results.hidden = this.#phase.kind !== "done";
    this.#renderResults();
    this.#scrollToSelected(true);
  }

  #scrollToSelected(focus: boolean): void {
    if (this.#selected === null) return;
    const row = this.#el.list.querySelector<HTMLElement>(
      `[data-key="${CSS.escape(this.#selected)}"]`,
    );
    if (!row) return;
    row.scrollIntoView({ block: "nearest" });
    if (focus) row.focus({ preventScroll: true });
  }

  #followSelection(unitId: string | null): void {
    if (unitId === null || this.#controller) return;
    const key = this.#order.find((k) => this.#item(k)?.violation.unitId === unitId);
    if (key && key !== this.#selected) this.#select(key, { reveal: false });
  }

  #onListKeydown(event: KeyboardEvent): void {
    const target = event.target as HTMLElement;
    if (!target.classList.contains("row")) return;
    const key = target.dataset.key ?? null;
    switch (event.key) {
      case "ArrowDown":
      case "ArrowUp": {
        event.preventDefault();
        const next = stepKey(this.#order, key, event.key === "ArrowDown" ? 1 : -1);
        if (next !== null) this.#select(next, { reveal: false, focus: true });
        break;
      }
      case "Enter":
        if (key) this.#openDetail(key);
        break;
    }
  }

  // -------------------------------------------------------------------------
  // ホストの操作
  // -------------------------------------------------------------------------

  async #reveal(item: ResultItem<L>): Promise<void> {
    if (this.#adapter.capabilities.reveal === "none") return;
    try {
      const result = await this.#adapter.reveal(item.unit, item.violation.displayRange);
      this.#showMessage(
        revealMessage(result, this.#adapter.capabilities.reveal, item.location, this.#terms.unit),
      );
    } catch (error) {
      console.error(error);
      this.#showMessage({ tone: "error", text: `移動に失敗しました: ${errorMessage(error)}` });
    }
  }

  async #applyFix(item: ResultItem<L>): Promise<void> {
    const fix = item.violation.fix;
    if (!fix) return;
    let result: Awaited<ReturnType<HostAdapter<L>["applyFix"]>>;
    try {
      result = await this.#adapter.applyFix(item.unit, fix);
    } catch (error) {
      result = { status: "failed", error };
    }
    if (result.status === "applied")
      this.#states = markFixed(this.#items, this.#states, [item.key]);
    if (result.status === "stale")
      this.#states = markStale(this.#items, this.#states, item.violation.unitId);
    this.#showMessage(applyFixMessage(result, item.location));
    this.#renderResults();
    if (this.#view === "detail") this.#renderDetail();
  }

  async #setHighlight(on: boolean): Promise<void> {
    if (on && this.#summary.units > HIGHLIGHT_CONFIRM_THRESHOLD) {
      const ok = await confirmDialog(this.#el.root, {
        title: `${this.#summary.units} ${this.#terms.unit}をハイライトしますか？`,
        body: [
          `違反のある${this.#terms.unit}ごとに書式を追加します。数が多いと、ブックの動作が重くなることがあります。`,
          "ハイライトはファイルに保存されます。保存する前に「ハイライトを解除」を押してください。",
        ],
        confirmLabel: "ハイライト",
      });
      if (!ok) return;
    }
    if (on) await this.#applyHighlight();
    else await this.#clearHighlight();
    this.#renderResults();
  }

  async #applyHighlight(): Promise<void> {
    if (this.#items.length === 0) return;
    this.#highlightOn = true;
    try {
      const units = new Map(this.#items.map((item) => [item.unit.id, item.unit]));
      await this.#adapter.highlight(
        this.#items.map((item) => item.violation),
        units,
      );
    } catch (error) {
      console.error(error);
      this.#showMessage({ tone: "error", text: `強調表示に失敗しました: ${errorMessage(error)}` });
    }
  }

  async #clearHighlight(): Promise<void> {
    this.#highlightOn = false;
    try {
      await this.#adapter.clearHighlight();
    } catch (error) {
      console.error(error);
    }
  }

  async #applyAll(items: readonly ResultItem<L>[]): Promise<void> {
    const batches = planFixes(items, this.#states);
    const count = batches.reduce((sum, batch) => sum + batch.keys.length, 0);
    if (count === 0 || this.#fixing) return;
    const unitNoun = this.#terms.unit;
    const note = applyFixNote(this.#adapter.capabilities, unitNoun);
    const ok = await confirmDialog(this.#el.root, {
      title: "修正案をすべて適用しますか？",
      body: [
        `表示中の ${count} 件の修正を、${batches.length} ${unitNoun}に適用します。適用できない${unitNoun}（数式など）は飛ばします。`,
        ...(note ? [note] : []),
        "元に戻すには Ctrl+Z（Mac は ⌘+Z）を押します。古いバージョンの Office では元に戻せないことがあります。",
      ],
      confirmLabel: "すべて適用",
    });
    if (!ok) return;

    this.#fixing = true;
    this.#renderResults();
    const counts: Partial<Record<FixStatus, number>> = {};
    try {
      for (const batch of batches) {
        let result: Awaited<ReturnType<HostAdapter<L>["applyFix"]>>;
        try {
          result = await this.#adapter.applyFix(batch.item.unit, batch.fix);
        } catch (error) {
          result = { status: "failed", error };
        }
        const status = result.status as FixStatus;
        counts[status] = (counts[status] ?? 0) + 1;
        if (status === "applied") this.#states = markFixed(this.#items, this.#states, batch.keys);
        else if (status === "stale")
          this.#states = markStale(this.#items, this.#states, batch.item.violation.unitId);
      }
    } finally {
      this.#fixing = false;
    }
    this.#showMessage(applyAllMessage(counts, unitNoun));
    this.#renderResults();
    if (this.#view === "detail") this.#renderDetail();
  }

  #openSettings(): void {
    if (this.#onOpenSettings) this.#onOpenSettings();
    else this.#showMessage({ tone: "info", text: "設定画面は準備中です（#11）。" });
  }

  async #copy(text: string): Promise<void> {
    const ok = await copyText(text);
    this.#showMessage(
      ok
        ? { tone: "success", text: `「${text}」をコピーしました。` }
        : { tone: "error", text: "コピーできませんでした。" },
    );
  }

  #showMessage(message: Message | null): void {
    this.#message = message;
    this.#renderMessage();
  }

  #onFilterChanged(): void {
    this.#limit = PAGE_SIZE;
    this.#renderFilters();
    this.#renderResults();
  }

  // -------------------------------------------------------------------------
  // 描画
  // -------------------------------------------------------------------------

  #renderScope(): void {
    const scopes = this.#adapter.capabilities.scopes;
    const el = this.#el.scope;
    el.hidden = scopes.length <= 1;
    replace(
      el,
      ...scopes.map((scope) =>
        h(
          "button",
          {
            type: "button",
            role: "radio",
            class: "segment",
            "aria-checked": String(scope === this.#scope),
            onclick: () => {
              this.#scope = scope;
              this.#renderScope();
              this.#renderStatus();
            },
          },
          scopeLabel(scope),
        ),
      ),
    );
  }

  #renderStatus(): void {
    const view = describePhase(this.#phase, this.#engineState, this.#terms.unit);
    const busy = view.busy || this.#controller !== undefined;
    const { run, cancel, statusText, statusHint, progress, progressBar, root } = this.#el;

    run.textContent = `${scopeLabel(this.#scope)}をチェック`;
    run.disabled = busy;
    cancel.hidden = !busy;
    for (const segment of this.#el.scope.querySelectorAll("button")) segment.disabled = busy;
    root.setAttribute("aria-busy", String(busy));

    statusText.textContent = view.text;
    statusText.dataset.tone = view.tone;
    statusHint.hidden = !view.hint;
    statusHint.textContent = view.hint ?? "";

    progress.hidden = view.progress === null;
    if (view.progress && "max" in view.progress) {
      const ratio = view.progress.max > 0 ? view.progress.value / view.progress.max : 0;
      progress.classList.remove("indeterminate");
      progress.setAttribute("aria-valuemin", "0");
      progress.setAttribute("aria-valuemax", String(view.progress.max));
      progress.setAttribute("aria-valuenow", String(view.progress.value));
      progressBar.style.width = `${Math.round(ratio * 1000) / 10}%`;
    } else {
      progress.classList.add("indeterminate");
      progress.removeAttribute("aria-valuenow");
      progressBar.style.width = "";
    }

    // 結果の欄は、完了してから出す（0 件のときは「問題なし」の表示）
    const showResults = this.#phase.kind === "done" && this.#view === "list";
    this.#el.results.hidden = !showResults;
    if (this.#phase.kind !== "done") {
      this.#el.detail.hidden = true;
      this.#el.nav.hidden = true;
    } else {
      this.#el.nav.hidden = this.#order.length === 0;
    }
  }

  #renderMessage(): void {
    const el = this.#el.message;
    const message = this.#message;
    el.hidden = message === null;
    if (!message) {
      el.replaceChildren();
      return;
    }
    el.dataset.tone = message.tone;
    el.setAttribute("role", message.tone === "error" ? "alert" : "status");
    replace(
      el,
      h("span", { class: "message-text" }, message.text),
      h(
        "button",
        {
          type: "button",
          class: "button subtle icon-only",
          "aria-label": "閉じる",
          onclick: () => this.#showMessage(null),
        },
        icon("close"),
      ),
    );
  }

  #renderFilters(): void {
    const summary = this.#summary;
    const { severityChips, containerFilter, ruleFilter, groupBy } = this.#el;

    severityChips.replaceChildren(
      ...SEVERITIES.filter((severity) => summary.bySeverity[severity] > 0).map((severity) =>
        h(
          "button",
          {
            type: "button",
            class: `chip severity-${severity}`,
            "aria-pressed": String(this.#filter.severities.has(severity)),
            onclick: () => {
              this.#filter = {
                ...this.#filter,
                severities: toggleSeverity(this.#filter.severities, severity),
              };
              this.#onFilterChanged();
            },
          },
          h("span", { class: "dot", "aria-hidden": "true" }),
          `${severityLabel(severity)} ${summary.bySeverity[severity]}`,
        ),
      ),
    );

    setOptions(
      containerFilter,
      [["", `すべて（${summary.containers.length}）`]].concat(
        summary.containers.map((c) => [c.key, `${c.label}（${c.count}）`]),
      ) as [string, string][],
      this.#filter.containerKey ?? "",
    );
    // 入れ物が 1 つしかないときは絞り込みもまとめ方の「シート別」も意味がない
    containerFilter.closest(".field")?.toggleAttribute("hidden", summary.containers.length <= 1);
    setOptions(
      ruleFilter,
      [["", `すべて（${summary.rules.length}）`]].concat(
        summary.rules.map((r) => [r.key, `${r.label}（${r.count}）`]),
      ) as [string, string][],
      this.#filter.ruleId ?? "",
    );
    groupBy.value = this.#groupBy;
  }

  #renderResults(): void {
    const summary = this.#summary;
    const unitNoun = this.#terms.unit;
    const filtered = filterItems(this.#items, this.#filter);
    const groups = groupItems(filtered, this.#groupBy);
    this.#order = flattenGroups(groups).map((item) => item.key);

    // サマリ
    replace(
      this.#el.summary,
      summary.total === 0
        ? h("p", { class: "summary-main" }, "問題は見つかりませんでした")
        : h(
            "p",
            { class: "summary-main" },
            h("strong", {}, `${summary.total} 件`),
            ` の違反（${summary.units} ${unitNoun}）`,
            summary.fixable > 0
              ? h("span", { class: "muted" }, ` ・修正案 ${summary.fixable} 件`)
              : null,
          ),
      isFiltered(this.#filter)
        ? h(
            "p",
            { class: "summary-filter" },
            `${filtered.length} 件を表示中 `,
            h(
              "button",
              {
                type: "button",
                class: "link",
                onclick: () => {
                  this.#filter = emptyFilter();
                  this.#onFilterChanged();
                },
              },
              "絞り込みを解除",
            ),
          )
        : null,
    );
    const hasItems = summary.total > 0;
    this.#el.summary.hidden = !hasItems; // 0 件のときは一覧の欄に「問題なし」を出す
    this.#el.severityChips.hidden = !hasItems;
    this.#el.toolbar.hidden = !hasItems;
    this.#renderBulk(filtered, hasItems);

    // 一覧
    const focusedKey = (document.activeElement as HTMLElement | null)?.closest<HTMLElement>(
      ".list .row",
    )?.dataset.key;
    const list = this.#el.list;
    if (!hasItems) {
      list.replaceChildren(
        h(
          "div",
          { class: "empty" },
          h("div", { class: "empty-icon", "aria-hidden": "true" }, "✓"),
          h("p", {}, "問題は見つかりませんでした"),
          h(
            "p",
            { class: "muted" },
            this.#phase.kind === "done"
              ? `${this.#phase.units} ${unitNoun}をチェックしました。`
              : "",
          ),
        ),
      );
    } else if (filtered.length === 0) {
      list.replaceChildren(h("p", { class: "empty muted" }, "条件に合う違反はありません。"));
    } else {
      let budget = this.#limit;
      const nodes: Child[] = [];
      for (const group of groups) {
        if (budget <= 0) break;
        const shown = group.items.slice(0, budget);
        budget -= shown.length;
        const rows = shown.map((item) => this.#renderRow(item));
        if (this.#groupBy === "none") {
          nodes.push(h("ol", { class: "rows" }, ...rows));
          continue;
        }
        const collapsed = this.#collapsed.has(group.key);
        nodes.push(
          h(
            "section",
            { class: "group" },
            h(
              "h2",
              { class: "group-header" },
              h(
                "button",
                {
                  type: "button",
                  "aria-expanded": String(!collapsed),
                  onclick: () => {
                    if (collapsed) this.#collapsed.delete(group.key);
                    else this.#collapsed.add(group.key);
                    this.#renderResults();
                  },
                },
                h("span", { class: "chevron", "aria-hidden": "true" }, icon("chevron-right")),
                h("span", { class: "group-label" }, group.label),
                group.detail ? h("span", { class: "group-detail" }, group.detail) : null,
                h("span", { class: "count" }, String(group.items.length)),
              ),
            ),
            collapsed ? null : h("ol", { class: "rows" }, ...rows),
          ),
        );
      }
      const rest = filtered.length - this.#limit;
      if (rest > 0)
        nodes.push(
          h(
            "button",
            {
              type: "button",
              class: "button more",
              onclick: () => {
                this.#limit += PAGE_SIZE;
                this.#renderResults();
              },
            },
            `さらに表示（残り ${rest} 件）`,
          ),
        );
      list.replaceChildren();
      append(list, nodes);
    }
    if (focusedKey)
      list.querySelector<HTMLElement>(`[data-key="${CSS.escape(focusedKey)}"]`)?.focus();

    // 前へ / 次へ
    this.#el.nav.hidden = this.#order.length === 0 || this.#phase.kind !== "done";
    const position = positionOf(this.#order, this.#selected);
    this.#el.navPosition.textContent =
      position > 0 ? `${position} / ${this.#order.length}` : `${this.#order.length} 件`;
  }

  /** 一覧の上の一括操作（すべて適用・ハイライト）。 */
  #renderBulk(filtered: readonly ResultItem<L>[], hasItems: boolean): void {
    const caps = this.#adapter.capabilities;
    const buttons: Child[] = [];
    if (hasItems && caps.applyFix !== "none") {
      const count = planFixes(filtered, this.#states).reduce((n, b) => n + b.keys.length, 0);
      if (count > 0)
        buttons.push(
          h(
            "button",
            {
              type: "button",
              class: "button",
              disabled: this.#fixing,
              onclick: () => void this.#applyAll(filtered),
            },
            this.#fixing ? "適用しています…" : `修正案をすべて適用（${count}）`,
          ),
        );
    }
    if (caps.highlight && (hasItems || this.#highlightOn)) {
      if (hasItems && !this.#highlightOn)
        buttons.push(
          h(
            "button",
            { type: "button", class: "button", onclick: () => void this.#setHighlight(true) },
            `ハイライト（${this.#summary.units} ${this.#terms.unit}）`,
          ),
        );
      // ハイライトは書式としてファイルに残るので、結果があるあいだは解除のボタンを常に出しておく
      buttons.push(
        h(
          "button",
          { type: "button", class: "button", onclick: () => void this.#setHighlight(false) },
          "ハイライトを解除",
        ),
      );
    }
    replace(this.#el.bulk, ...buttons);
    this.#el.bulk.hidden = buttons.length === 0;
  }

  #groupKey(item: ResultItem<L>): string {
    if (this.#groupBy === "container") return item.container.key;
    if (this.#groupBy === "rule") return item.violation.ruleId;
    return "";
  }

  #renderRow(item: ResultItem<L>): HTMLLIElement {
    const { violation } = item;
    const state = this.#states.get(item.key);
    const actions = itemActions(this.#adapter.capabilities, violation, state);
    const selected = item.key === this.#selected;
    const rule = splitRuleId(violation.ruleId);

    const buttons: Child[] = [];
    if (actions.reveal)
      buttons.push(
        h(
          "button",
          {
            type: "button",
            class: "button small",
            onclick: (event) => {
              event.stopPropagation();
              this.#select(item.key, { reveal: true });
            },
          },
          actions.revealLabel,
        ),
      );
    if (actions.fix && violation.fix)
      buttons.push(
        h(
          "button",
          {
            type: "button",
            class: "button small",
            title: `「${violation.fix.text}」に置き換える`,
            onclick: (event) => {
              event.stopPropagation();
              void this.#applyFix(item);
            },
          },
          "修正",
        ),
      );

    return h(
      "li",
      {
        class: `row${selected ? " selected" : ""}${state ? ` ${state}` : ""}`,
        "data-key": item.key,
        tabindex: selected || (this.#selected === null && item.key === this.#order[0]) ? 0 : -1,
        "aria-current": selected ? "true" : undefined,
        onclick: () => this.#openDetail(item.key),
      },
      h(
        "div",
        { class: "row-head" },
        severityBadge(violation.severity),
        h("span", { class: "location" }, item.location),
        this.#groupBy === "rule"
          ? null
          : h("span", { class: "rule", title: violation.ruleId }, rule.name),
      ),
      renderExcerpt(buildExcerpt(item.unit.text, violation.displayRange, { singleLine: true }), {
        severity: violation.severity,
        className: "excerpt",
      }),
      h("p", { class: "row-message" }, violation.message),
      violation.fix ? renderFixInline(item.unit.text, violation.fix) : null,
      state ? stateBadge(state) : null,
      buttons.length > 0 ? h("div", { class: "row-actions" }, ...buttons) : null,
    );
  }

  #renderDetail(): void {
    const item = this.#item(this.#selected);
    const el = this.#el.detail;
    if (!item) {
      el.replaceChildren();
      return;
    }
    const { violation, unit } = item;
    const caps = this.#adapter.capabilities;
    const state = this.#states.get(item.key);
    const actions = itemActions(caps, violation, state);
    const rule = splitRuleId(violation.ruleId);
    const [start, end] = violation.displayRange;
    const matchText = unit.text.slice(start, end);
    const note = actions.fix ? applyFixNote(caps, this.#terms.unit) : null;

    replace(
      el,
      h(
        "button",
        { type: "button", class: "button subtle back", onclick: () => this.#showList() },
        icon("chevron-left"),
        "一覧に戻る",
      ),
      h(
        "div",
        { class: "detail-head" },
        severityBadge(violation.severity),
        h("h2", { class: "detail-location" }, item.location),
        state ? stateBadge(state) : null,
      ),
      h("p", { class: "detail-message" }, violation.message),
      h(
        "dl",
        { class: "props" },
        h("dt", {}, "ルール"),
        h(
          "dd",
          {},
          h("span", { class: "mono" }, rule.name),
          rule.preset ? h("span", { class: "muted" }, ` （${rule.preset}）`) : null,
        ),
        h("dt", {}, "位置"),
        h("dd", {}, `${violation.line} 行 ${violation.column} 文字目`),
      ),
      h("h3", {}, "該当箇所"),
      renderExcerpt(buildExcerpt(unit.text, violation.displayRange, { context: 40 }), {
        severity: violation.severity,
        className: "excerpt excerpt-block",
      }),
      violation.fix ? h("h3", {}, "修正案") : null,
      violation.fix ? renderFixPreview(unit.text, violation.fix) : null,
      h(
        "div",
        { class: "detail-actions" },
        actions.fix
          ? h(
              "button",
              {
                type: "button",
                class: "button primary",
                onclick: () => void this.#applyFix(item),
              },
              "修正を適用",
            )
          : null,
        actions.reveal
          ? h(
              "button",
              { type: "button", class: "button", onclick: () => void this.#reveal(item) },
              `${this.#terms.unit}へ${actions.revealLabel}`,
            )
          : null,
        h(
          "button",
          { type: "button", class: "button", onclick: () => void this.#copy(matchText) },
          "該当箇所をコピー",
        ),
      ),
      note ? h("p", { class: "note muted" }, note) : null,
      caps.reveal === "none"
        ? h(
            "p",
            { class: "note muted" },
            `このホストでは${this.#terms.unit}への移動に対応していません。該当箇所をコピーして検索してください。`,
          )
        : null,
      h(
        "details",
        { class: "fulltext" },
        h("summary", {}, `${this.#terms.unit}の全文`),
        h("p", { class: "excerpt-block" }, unit.text),
      ),
    );
  }
}

// ---------------------------------------------------------------------------
// 部品
// ---------------------------------------------------------------------------

function severityBadge(severity: Severity): HTMLElement {
  return h(
    "span",
    { class: `badge severity-${severity}` },
    h("span", { class: "dot", "aria-hidden": "true" }),
    severityLabel(severity),
  );
}

function stateBadge(state: ItemState): HTMLElement {
  return h(
    "span",
    { class: `badge state-${state}` },
    state === "fixed" ? "修正済み" : "要再チェック",
  );
}

function renderExcerpt(
  excerpt: Excerpt,
  options: { severity: Severity; className: string },
): HTMLElement {
  return h(
    "p",
    { class: options.className },
    `${excerpt.clippedStart ? "…" : ""}${excerpt.before}`,
    h("mark", { class: `severity-${options.severity}` }, excerpt.match || "␣"),
    `${excerpt.after}${excerpt.clippedEnd ? "…" : ""}`,
  );
}

function renderFixPreview(text: string, fix: { range: readonly [number, number]; text: string }) {
  const preview = buildFixPreview(text, fix, { context: 20 });
  return h(
    "p",
    { class: "excerpt excerpt-block fix-preview" },
    `${preview.clippedStart ? "…" : ""}${preview.before}`,
    preview.removed ? h("del", {}, preview.removed) : null,
    preview.inserted ? h("ins", {}, preview.inserted) : null,
    `${preview.after}${preview.clippedEnd ? "…" : ""}`,
  );
}

/** 一覧の行に出す「修正前 → 修正後」。 */
function renderFixInline(text: string, fix: Fix): HTMLElement {
  const preview = buildFixPreview(text, fix, { context: 0, singleLine: true });
  return h(
    "p",
    { class: "row-fix" },
    h("span", { class: "muted" }, "修正案 "),
    preview.removed ? h("del", {}, preview.removed) : h("span", { class: "muted" }, "挿入"),
    " → ",
    preview.inserted ? h("ins", {}, preview.inserted) : h("span", { class: "muted" }, "削除"),
  );
}

function select(label: string, options: [string, string][]): HTMLSelectElement {
  const el = h("select", { "aria-label": label });
  setOptions(el, options, options[0]?.[0] ?? "");
  return el;
}

function setOptions(el: HTMLSelectElement, options: [string, string][], value: string): void {
  el.replaceChildren(...options.map(([v, label]) => h("option", { value: v }, label)));
  el.value = value;
}

function field(label: string, control: HTMLElement): HTMLElement {
  return h("label", { class: "field" }, h("span", { class: "field-label" }, label), control);
}
