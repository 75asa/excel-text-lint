import { describe, expect, it } from "vitest";
import { LintEngine } from "./engine";
import type { TextUnit } from "./types";

type Message = {
  id?: number;
  command: string;
  text?: string;
  ruleId?: string;
  textlintrc?: unknown;
};

/**
 * script-compiler の Worker と同じメッセージをやり取りする偽物。
 * 「ダメ」という文字を 1 件の違反として返す。
 */
class FakeWorker extends EventTarget {
  readonly url: URL;
  readonly received: Message[] = [];
  terminated = false;
  /** false にすると応答を返さない（手動で respond する）。 */
  autoRespond = true;
  sendInit = true;

  constructor(url: URL) {
    super();
    this.url = url;
    queueMicrotask(() => {
      if (this.sendInit) this.emit({ command: "init" });
    });
  }

  postMessage(message: Message): void {
    this.received.push(message);
    if (this.autoRespond) queueMicrotask(() => this.respond(message));
  }

  respond(message: Message): void {
    const text = message.text ?? "";
    if (message.command === "lint") {
      if (text.includes("throw")) {
        this.emit({ id: message.id, command: "error", error: new Error("rule crashed") });
        return;
      }
      const index = text.indexOf("ダメ");
      const messages =
        index < 0
          ? []
          : [
              {
                ruleId: "fake/no-dame",
                message: "ダメ",
                severity: 2,
                index,
                line: 1,
                column: index + 1,
                range: [index, index + 1],
                fix: { range: [index, index + 2], text: "良い" },
              },
            ];
      this.emit({ id: message.id, command: "lint:result", result: { filePath: "", messages } });
    } else if (message.command === "fix") {
      this.emit({
        id: message.id,
        command: "fix:result",
        result: { output: text.replaceAll("ダメ", "良い"), remainingMessages: [] },
      });
    }
  }

  emit(data: unknown): void {
    this.dispatchEvent(new MessageEvent("message", { data }));
  }

  terminate(): void {
    this.terminated = true;
  }
}

function setup(
  options: {
    maxInFlight?: number;
    configure?: (worker: FakeWorker) => void;
    initTimeoutMs?: number;
  } = {},
) {
  const workers: FakeWorker[] = [];
  const engine = new LintEngine({
    workerUrl: "https://addin.example/textlint/loader.js",
    dictBaseUrl: "https://addin.example/dict/",
    maxInFlight: options.maxInFlight,
    initTimeoutMs: options.initTimeoutMs,
    createWorker: (url) => {
      const worker = new FakeWorker(url);
      options.configure?.(worker);
      workers.push(worker);
      return worker as unknown as Worker;
    },
  });
  return { engine, workers };
}

const units = (...texts: string[]): TextUnit[] =>
  texts.map((text, i) => ({ id: `u${i}`, text, location: null }));

describe("LintEngine", () => {
  it("辞書の配信先を Worker の URL のクエリで渡す", async () => {
    const { engine, workers } = setup();
    await engine.init();
    expect(workers[0]!.url.href).toBe(
      "https://addin.example/textlint/loader.js?dict=https%3A%2F%2Faddin.example%2Fdict%2F",
    );
  });

  it("init は Worker を 1 つしか作らない", async () => {
    const { engine, workers } = setup();
    await Promise.all([engine.init(), engine.init()]);
    await engine.init();
    expect(workers).toHaveLength(1);
  });

  it("TextUnit を 1 件ずつ送り、入力の順に Violation を返す", async () => {
    const { engine, workers } = setup({ maxInFlight: 2 });
    const progress: number[] = [];
    const violations = await engine.lint(units("これはダメ", "問題なし", "", "ダメです"), {
      onProgress: ({ done }) => progress.push(done),
    });
    expect(violations.map((v) => [v.unitId, v.range, v.displayRange])).toEqual([
      ["u0", [3, 4], [3, 5]],
      ["u3", [0, 1], [0, 2]],
    ]);
    expect(progress).toEqual([1, 2, 3, 4]);
    // 空のテキストは送らない。1 メッセージに 1 件ずつ送る
    const texts = workers[0]!.received.map((m) => m.text);
    expect(texts).toEqual(["これはダメ", "問題なし", "ダメです"]);
    // リクエスト ID は重複しない
    const ids = workers[0]!.received.map((m) => m.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("応答の順が入れ替わってもリクエスト ID で対応づける", async () => {
    const { engine, workers } = setup({
      maxInFlight: 3,
      configure: (w) => (w.autoRespond = false),
    });
    const promise = engine.lint(units("ダメ1", "OK", "xダメ"));
    await vi_waitFor(() => workers[0]?.received.length === 3);
    const [a, b, c] = workers[0]!.received;
    workers[0]!.respond(c!);
    workers[0]!.respond(a!);
    workers[0]!.respond(b!);
    const violations = await promise;
    expect(violations.map((v) => [v.unitId, v.range[0]])).toEqual([
      ["u0", 0],
      ["u2", 1],
    ]);
  });

  it("maxInFlight を超えて同時に送らない", async () => {
    const { engine, workers } = setup({
      maxInFlight: 2,
      configure: (w) => (w.autoRespond = false),
    });
    const promise = engine.lint(units("a", "b", "c", "d"));
    await vi_waitFor(() => workers[0]?.received.length === 2);
    await Promise.resolve();
    expect(workers[0]!.received).toHaveLength(2);
    workers[0]!.respond(workers[0]!.received[0]!);
    await vi_waitFor(() => workers[0]!.received.length === 3);
    for (const m of workers[0]!.received.slice(1)) workers[0]!.respond(m);
    await vi_waitFor(() => workers[0]!.received.length === 4);
    workers[0]!.respond(workers[0]!.received[3]!);
    await promise;
  });

  it("AbortSignal で中断すると signal.reason で reject し、残りは送らない", async () => {
    const { engine, workers } = setup({
      maxInFlight: 1,
      configure: (w) => (w.autoRespond = false),
    });
    const controller = new AbortController();
    const promise = engine.lint(units("a", "b", "c"), { signal: controller.signal });
    await vi_waitFor(() => workers[0]?.received.length === 1);
    const reason = new Error("stop");
    controller.abort(reason);
    await expect(promise).rejects.toBe(reason);
    // 中断した後に届いた応答は無視する
    workers[0]!.respond(workers[0]!.received[0]!);
    await Promise.resolve();
    expect(workers[0]!.received).toHaveLength(1);
  });

  it("中断済みの signal ではすぐに reject する", async () => {
    const { engine } = setup();
    const controller = new AbortController();
    controller.abort(new Error("already"));
    await expect(engine.lint(units("a"), { signal: controller.signal })).rejects.toThrow("already");
  });

  it("Worker が error を返したら reject する", async () => {
    const { engine } = setup();
    await expect(engine.lint(units("ok", "throw"))).rejects.toThrow("rule crashed");
    // エンジンはそのまま使い続けられる
    await expect(engine.lint(units("ダメ"))).resolves.toHaveLength(1);
  });

  it("id のない error（Worker 全体のエラー）は処理中のリクエストをすべて失敗にする", async () => {
    const { engine, workers } = setup({ configure: (w) => (w.autoRespond = false) });
    const promise = engine.lint(units("a", "b"));
    await vi_waitFor(() => workers[0]?.received.length === 2);
    workers[0]!.emit({ command: "error", error: new Error("unexpected") });
    await expect(promise).rejects.toThrow("unexpected");
  });

  it("init が返らなければタイムアウトし、次の init でやり直す", async () => {
    let first = true;
    const { engine, workers } = setup({
      initTimeoutMs: 10,
      configure: (w) => {
        w.sendInit = !first;
        first = false;
      },
    });
    await expect(engine.init()).rejects.toThrow("起動しませんでした");
    expect(workers[0]!.terminated).toBe(true);
    await engine.init();
    expect(workers).toHaveLength(2);
  });

  it("fix は修正後のテキストを返す", async () => {
    const { engine, workers } = setup();
    const result = await engine.fix(
      { id: "u", text: "これはダメ", location: null },
      { ruleId: "fake/no-dame" },
    );
    expect(result).toEqual({ output: "これは良い", remaining: [] });
    expect(workers[0]!.received.at(-1)).toMatchObject({ command: "fix", ruleId: "fake/no-dame" });
  });

  it("mergeConfig は merge-config を送る", async () => {
    const { engine, workers } = setup();
    await engine.mergeConfig({ rules: { "fake/no-dame": false } });
    expect(workers[0]!.received.at(-1)).toEqual({
      command: "merge-config",
      textlintrc: { rules: { "fake/no-dame": false } },
    });
  });

  it("dispose で処理中のリクエストを reject し、Worker を止める", async () => {
    const { engine, workers } = setup({ configure: (w) => (w.autoRespond = false) });
    const promise = engine.lint(units("a"));
    await vi_waitFor(() => workers[0]?.received.length === 1);
    engine.dispose();
    await expect(promise).rejects.toThrow("破棄");
    expect(workers[0]!.terminated).toBe(true);
  });
});

async function vi_waitFor(condition: () => boolean): Promise<void> {
  for (let i = 0; i < 100; i++) {
    if (condition()) return;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error("条件が満たされませんでした");
}
