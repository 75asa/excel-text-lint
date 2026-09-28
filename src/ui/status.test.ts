import { describe, expect, it } from "vitest";
import { describePhase, estimateRemainingMs, formatDuration, scopeNotice } from "./status";

describe("describePhase", () => {
  it("初回の lint で最初の結果が出るまでは、辞書の読み込み中として不確定の進捗を出す", () => {
    const view = describePhase(
      { kind: "linting", done: 0, total: 10, violations: 0, warm: false },
      "ready",
      "セル",
    );
    expect(view.text).toContain("辞書");
    expect(view.text).toContain("15MB");
    expect(view.progress).toEqual({ indeterminate: true });
    expect(view.busy).toBe(true);
  });

  it("辞書を読み込んだあとは件数の進捗を出す", () => {
    const view = describePhase(
      { kind: "linting", done: 3, total: 10, violations: 2, warm: true },
      "ready",
      "セル",
    );
    expect(view.text).toBe("チェックしています… 3 / 10 セル（違反 2 件）");
    expect(view.progress).toEqual({ value: 3, max: 10 });
  });

  it("2 回目以降は done が 0 でも辞書の読み込みとは言わない", () => {
    const view = describePhase(
      { kind: "linting", done: 0, total: 5, violations: 0, warm: true },
      "ready",
      "セル",
    );
    expect(view.text).not.toContain("辞書");
  });

  it("0 件のときは問題なしと伝える", () => {
    const view = describePhase(
      { kind: "done", scope: "selection", units: 4, violations: 0 },
      "ready",
      "セル",
    );
    expect(view).toMatchObject({ tone: "success", busy: false, progress: null });
    expect(view.text).toContain("問題は見つかりませんでした");
  });

  it("待機中はエンジンの状態を出す", () => {
    expect(describePhase({ kind: "idle" }, "starting", "セル").text).toContain("準備");
    expect(describePhase({ kind: "idle" }, "failed", "セル").tone).toBe("warning");
    expect(describePhase({ kind: "idle" }, "ready", "セル").tone).toBe("neutral");
  });

  it("収集・中止・エラー・空", () => {
    expect(
      describePhase(
        { kind: "collecting", scope: "sheet", collected: 5, total: 20 },
        "ready",
        "セル",
      ),
    ).toMatchObject({ progress: { value: 5, max: 20 }, busy: true });
    expect(describePhase({ kind: "cancelled" }, "ready", "セル").text).toBe("中止しました。");
    expect(describePhase({ kind: "error", message: "boom" }, "ready", "セル")).toMatchObject({
      tone: "error",
      text: "チェックに失敗しました: boom",
    });
    expect(describePhase({ kind: "empty", scope: "workbook" }, "ready", "セル").text).toBe(
      "ブックにチェックする文字列がありません。",
    );
  });
});

describe("estimateRemainingMs / formatDuration", () => {
  it("最初の結果からの速さで残りを見積もる", () => {
    // 10 件を 2 秒 → 残り 89 件は 17.8 秒
    expect(estimateRemainingMs({ done: 1, at: 0 }, { done: 11, at: 2000 }, 100)).toBe(17800);
  });

  it("件数か時間が少なすぎる、または終わっているなら null", () => {
    expect(estimateRemainingMs({ done: 1, at: 0 }, { done: 3, at: 5000 }, 100)).toBeNull();
    expect(estimateRemainingMs({ done: 1, at: 0 }, { done: 50, at: 500 }, 100)).toBeNull();
    expect(estimateRemainingMs({ done: 1, at: 0 }, { done: 100, at: 5000 }, 100)).toBeNull();
  });

  it("残り時間を丸める", () => {
    expect(formatDuration(3200)).toBe("約 3 秒");
    expect(formatDuration(23_000)).toBe("約 25 秒");
    expect(formatDuration(170_000)).toBe("約 3 分");
  });

  it("残り時間が分かれば補足に出す", () => {
    const view = describePhase(
      { kind: "linting", done: 50, total: 100, violations: 0, warm: true, remainingMs: 7000 },
      "ready",
      "セル",
    );
    expect(view.hint).toBe("残り 約 7 秒");
  });
});

describe("scopeNotice", () => {
  it("ブック・シートは時間がかかる旨を出し、選択範囲では出さない", () => {
    expect(scopeNotice("workbook", "セル")).toContain("1 万セルで約 7 秒");
    expect(scopeNotice("sheet", "セル")).toContain("時間がかかる");
    expect(scopeNotice("selection", "セル")).toBeNull();
  });
});
