/**
 * textlint の Worker を包むクライアント（#8）。
 *
 * Worker は `@textlint/script-compiler` が生成したもの（`npm run build:worker`）を、
 * 辞書の URL を書き換えるローダー（`src/textlint/loader.js`）経由で起動する。
 * build ではどちらもファイル名にコンテンツハッシュが付くので、ローダーの URL は `textlintWorkerUrl()`
 * （`worker-url.ts`）で組み立てて `workerUrl` に渡す（#45）。
 *
 * Worker とのやり取り（script-compiler の仕様）:
 * - 送信: `{ id, command: "lint" | "fix", text, ext, ruleId? }`、`{ command: "merge-config", textlintrc }`
 * - 受信: `{ command: "init" }`（起動直後に 1 回）、`{ id, command: "lint:result" | "fix:result", result }`、
 *   `{ id?, command: "error", error }`（id がないのは Worker 全体のエラー）
 *
 * TextUnit は 1 件ずつ Worker に送る。連結して送ってはいけない（超線形に遅くなる。#2 の PoC を参照）。
 */

import type { TextUnit, Violation } from "./types";
import { type TextlintMessageLike, toViolation } from "./violation";

export interface LintEngineOptions {
  /**
   * ローダー（`src/textlint/loader.js`）の URL。相対 URL は `location.href` を基準に解決する。
   * build ではファイル名にハッシュが付くので、`textlintWorkerUrl()` の戻り値を渡す。
   */
  workerUrl: string | URL;
  /**
   * kuromoji の辞書を置いた場所（`base.dat.gz` などが並ぶディレクトリ）。
   * 省略すると jsdelivr（textlint の既定）から取得する。
   */
  dictBaseUrl?: string | URL;
  /** textlint に渡す拡張子。既定は `.txt`。 */
  ext?: string;
  /** 同時に Worker へ送っておくリクエストの数。Worker の中では順に処理される。既定は 8。 */
  maxInFlight?: number;
  /** Worker が `init` を返すまでの待ち時間（ミリ秒）。既定は 30 秒。 */
  initTimeoutMs?: number;
  /** Worker の生成方法（テスト用）。既定は `new Worker(url)`（classic worker）。 */
  createWorker?: (url: URL) => Worker;
}

export interface LintProgress {
  /** lint し終えた TextUnit の数。 */
  done: number;
  total: number;
  /** ここまでに見つかった違反の数。 */
  violations: number;
}

export interface LintOptions {
  signal?: AbortSignal;
  onProgress?: (progress: LintProgress) => void;
  /** TextUnit 1 件の lint が終わるたびに呼ばれる（結果を順次表示したいとき用）。完了の順は入力の順と限らない。 */
  onUnitResult?: (unit: TextUnit, violations: Violation[]) => void;
}

export interface FixOptions {
  signal?: AbortSignal;
  /** 指定したルールだけで修正する。 */
  ruleId?: string;
}

export interface FixResult {
  /** 修正後のテキスト。 */
  output: string;
  /** 修正後のテキストに残っている違反（`unitId` は元の TextUnit のまま）。 */
  remaining: Violation[];
}

interface TextlintLintResult {
  messages: TextlintMessageLike[];
}

interface TextlintFixResult {
  output: string;
  remainingMessages: TextlintMessageLike[];
}

type WorkerRequest = { command: "lint" | "fix"; text: string; ext: string; ruleId?: string };

interface Pending {
  resolve: (result: unknown) => void;
  reject: (error: unknown) => void;
}

export class LintEngine {
  readonly #options: Required<Pick<LintEngineOptions, "ext" | "maxInFlight" | "initTimeoutMs">> &
    LintEngineOptions;
  #worker: Worker | undefined;
  #ready: Promise<void> | undefined;
  #rejectReady: ((error: unknown) => void) | undefined;
  readonly #pending = new Map<number, Pending>();
  #seq = 0;

  constructor(options: LintEngineOptions) {
    this.#options = {
      ...options,
      ext: options.ext ?? ".txt",
      maxInFlight: options.maxInFlight ?? 8,
      initTimeoutMs: options.initTimeoutMs ?? 30_000,
    };
  }

  /** Worker を起動し、`init` を待つ。何度呼んでもよい（2 回目以降は同じ Promise を返す）。 */
  init(): Promise<void> {
    this.#ready ??= this.#start();
    return this.#ready;
  }

  /** textlint の設定（.textlintrc 相当の JSON）を上書きする。同梱していないルールは追加できない。 */
  async mergeConfig(textlintrc: object): Promise<void> {
    await this.init();
    // merge-config には応答がない。Worker はメッセージを順に処理するので、以降の lint には反映される。
    this.#worker?.postMessage({ command: "merge-config", textlintrc });
  }

  /**
   * TextUnit を 1 件ずつ lint して、違反を入力の順に並べて返す。
   *
   * `signal` で中断すると `signal.reason` で reject する。すでに Worker に送った分の結果は捨てる。
   */
  async lint(units: readonly TextUnit[], options: LintOptions = {}): Promise<Violation[]> {
    const { signal, onProgress, onUnitResult } = options;
    signal?.throwIfAborted();
    await this.init();
    signal?.throwIfAborted();

    // 1 件でも失敗したら、残りも止めるための内部の signal
    const controller = new AbortController();
    const abort = () => controller.abort(signal?.reason);
    signal?.addEventListener("abort", abort, { once: true });

    const results: Violation[][] = new Array(units.length);
    let next = 0;
    let done = 0;
    let violations = 0;
    const run = async () => {
      while (next < units.length) {
        controller.signal.throwIfAborted();
        const index = next++;
        const unit = units[index]!;
        const found = await this.lintUnit(unit, { signal: controller.signal });
        results[index] = found;
        done += 1;
        violations += found.length;
        onUnitResult?.(unit, found);
        onProgress?.({ done, total: units.length, violations });
      }
    };

    try {
      const lanes = Math.max(1, Math.min(this.#options.maxInFlight, units.length));
      await Promise.all(
        Array.from({ length: lanes }, () =>
          run().catch((error: unknown) => {
            if (!controller.signal.aborted) controller.abort(error);
            throw error;
          }),
        ),
      );
    } catch (error) {
      // 中断のときは、どの lane のエラーでもなく中断の理由を投げる
      throw controller.signal.aborted ? controller.signal.reason : error;
    } finally {
      signal?.removeEventListener("abort", abort);
    }
    return results.flat();
  }

  /** TextUnit 1 件を lint する。 */
  async lintUnit(unit: TextUnit, options: { signal?: AbortSignal } = {}): Promise<Violation[]> {
    if (unit.text.trim() === "") return [];
    const result = await this.#request<TextlintLintResult>(
      { command: "lint", text: unit.text, ext: this.#options.ext },
      options.signal,
    );
    return result.messages.map((message) => toViolation(unit.id, unit.text, message));
  }

  /** TextUnit 1 件に textlint の自動修正をかけた結果を返す。ホストへの書き込みはしない。 */
  async fix(unit: TextUnit, options: FixOptions = {}): Promise<FixResult> {
    const request: WorkerRequest = { command: "fix", text: unit.text, ext: this.#options.ext };
    if (options.ruleId !== undefined) request.ruleId = options.ruleId;
    const result = await this.#request<TextlintFixResult>(request, options.signal);
    return {
      output: result.output,
      remaining: result.remainingMessages.map((message) =>
        toViolation(unit.id, result.output, message),
      ),
    };
  }

  /** Worker を止める。処理中のリクエストは reject する。 */
  dispose(): void {
    this.#fail(new Error("LintEngine は破棄されました"));
    this.#worker?.terminate();
    this.#worker = undefined;
  }

  #start(): Promise<void> {
    const url = this.#resolveWorkerUrl();
    const worker = (this.#options.createWorker ?? ((u) => new Worker(u)))(url);
    this.#worker = worker;

    const ready = new Promise<void>((resolve, reject) => {
      const timer = setTimeout(
        () =>
          reject(
            new Error(
              `textlint の Worker が ${this.#options.initTimeoutMs} ms 以内に起動しませんでした`,
            ),
          ),
        this.#options.initTimeoutMs,
      );
      this.#rejectReady = (error) => {
        clearTimeout(timer);
        reject(error);
      };
      worker.addEventListener("message", (event: MessageEvent) => {
        if ((event.data as { command?: unknown } | null)?.command === "init") {
          clearTimeout(timer);
          this.#rejectReady = undefined;
          resolve();
        }
      });
    });

    worker.addEventListener("message", (event: MessageEvent) => this.#onMessage(event.data));
    worker.addEventListener("error", (event: ErrorEvent) => {
      this.#fail(
        new Error(
          `textlint の Worker でエラーが発生しました: ${event.message || "（詳細不明。Worker のスクリプトが読み込めない可能性があります）"}`,
        ),
      );
    });

    // 起動に失敗したら、次の init() でやり直せるようにする
    ready.catch(() => {
      if (this.#worker === worker) {
        worker.terminate();
        this.#worker = undefined;
        this.#ready = undefined;
      }
    });
    return ready;
  }

  #resolveWorkerUrl(): URL {
    const base = globalThis.location?.href;
    const url = new URL(this.#options.workerUrl, base);
    if (this.#options.dictBaseUrl !== undefined) {
      url.searchParams.set("dict", new URL(this.#options.dictBaseUrl, base).href);
    }
    return url;
  }

  #request<T>(message: WorkerRequest, signal?: AbortSignal): Promise<T> {
    return this.init().then(
      () =>
        new Promise<T>((resolve, reject) => {
          signal?.throwIfAborted();
          const worker = this.#worker;
          if (!worker) throw new Error("textlint の Worker が起動していません");
          const id = ++this.#seq;
          const onAbort = () => {
            this.#pending.delete(id);
            reject(signal?.reason);
          };
          this.#pending.set(id, {
            resolve: (result) => {
              signal?.removeEventListener("abort", onAbort);
              resolve(result as T);
            },
            reject: (error) => {
              signal?.removeEventListener("abort", onAbort);
              reject(error);
            },
          });
          signal?.addEventListener("abort", onAbort, { once: true });
          worker.postMessage({ id, ...message });
        }),
    );
  }

  #onMessage(data: unknown): void {
    if (typeof data !== "object" || data === null) return;
    const { id, command, result, error } = data as {
      id?: unknown;
      command?: unknown;
      result?: unknown;
      error?: unknown;
    };
    if (command === "init") return;

    if (typeof id !== "number") {
      // id のないエラーは Worker 全体のエラー（unhandledrejection など）。どのリクエストのものか分からないので、すべて失敗にする
      if (command === "error")
        this.#fail(toError(error, "textlint の Worker で予期しないエラーが発生しました"));
      return;
    }

    const pending = this.#pending.get(id);
    if (!pending) return; // 中断済み
    this.#pending.delete(id);
    if (command === "error") pending.reject(toError(error, "textlint の実行に失敗しました"));
    else pending.resolve(result);
  }

  #fail(error: Error): void {
    this.#rejectReady?.(error);
    this.#rejectReady = undefined;
    for (const pending of this.#pending.values()) pending.reject(error);
    this.#pending.clear();
  }
}

function toError(value: unknown, fallback: string): Error {
  if (value instanceof Error) return value;
  if (typeof value === "object" && value !== null && "message" in value)
    return new Error(String(value.message));
  return new Error(fallback, { cause: value });
}
