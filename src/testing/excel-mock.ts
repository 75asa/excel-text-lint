/**
 * テスト用の Excel JavaScript API（Office.js）のフェイク。
 *
 * 方針:
 * - 本物の Office.js は読み込まない。Office の外（Node）では動かないため
 * - `globalThis.Excel` を `vi.stubGlobal` で差し替え、`Excel.run` に渡される
 *   `RequestContext` をフェイクにする。テスト対象のコードは変更せずに、グローバルの `Excel` を使ったまま動かせる
 * - フェイクは、テストで使う API だけを必要な分だけ実装する（型は `as unknown as` で合わせる）
 * - `load()` と `context.sync()` を呼ばずにプロパティを読むと、本物と同じように例外を投げる。
 *   load / sync の呼び忘れをテストで検出できるようにするため
 */
import { vi } from "vitest";

/** 選択範囲のフェイクの中身。 */
export interface FakeSelection {
  /** シート名。 */
  sheetName: string;
  /** 0 始まりの行番号（選択範囲の左上）。 */
  rowIndex: number;
  /** 0 始まりの列番号（選択範囲の左上）。 */
  columnIndex: number;
  /** `Range.text` の値（行 × 列）。 */
  text: string[][];
}

/** フェイクの `Excel.run` の呼び出し状況。テストでの検証に使う。 */
export interface ExcelMock {
  /** `Excel.run` が呼ばれた回数。 */
  runCount: number;
  /** `context.sync()` が呼ばれた回数。 */
  syncCount: number;
}

/**
 * `load()` と `sync()` を経たプロパティだけを読めるプロキシ・オブジェクトを作る。
 * `pending` は load 済みで sync 待ちのプロパティ、`loaded` は読めるプロパティ。
 */
class FakeClientObject<T extends object> {
  private readonly pending = new Set<string>();
  private readonly loaded = new Set<string>();

  constructor(private readonly values: T) {}

  load(names: string | string[]): void {
    for (const name of Array.isArray(names) ? names : names.split(",")) {
      this.pending.add(name.trim());
    }
  }

  /** `context.sync()` から呼ぶ。 */
  commit(): void {
    for (const name of this.pending) this.loaded.add(name);
    this.pending.clear();
  }

  get<K extends keyof T & string>(name: K): T[K] {
    if (!this.loaded.has(name)) {
      throw new Error(
        `PropertyNotLoaded: "${name}" を読む前に load() と context.sync() を呼んでください`,
      );
    }
    return this.values[name];
  }
}

/**
 * `globalThis.Excel` をフェイクに差し替える。
 * 元に戻すのは Vitest の `unstubGlobals: true`（vitest.config.ts）に任せる。
 */
export function stubExcel(selection: FakeSelection): ExcelMock {
  const mock: ExcelMock = { runCount: 0, syncCount: 0 };

  const sheet = new FakeClientObject({ name: selection.sheetName });
  const range = new FakeClientObject({
    text: selection.text,
    rowIndex: selection.rowIndex,
    columnIndex: selection.columnIndex,
  });

  const worksheet = {
    load: (names: string | string[]) => sheet.load(names),
    get name() {
      return sheet.get("name");
    },
  };
  const selectedRange = {
    load: (names: string | string[]) => range.load(names),
    worksheet,
    get text() {
      return range.get("text");
    },
    get rowIndex() {
      return range.get("rowIndex");
    },
    get columnIndex() {
      return range.get("columnIndex");
    },
  };
  const context = {
    workbook: { getSelectedRange: () => selectedRange },
    sync: async () => {
      mock.syncCount += 1;
      sheet.commit();
      range.commit();
    },
  };

  vi.stubGlobal("Excel", {
    run: async <R>(batch: (ctx: typeof context) => Promise<R>): Promise<R> => {
      mock.runCount += 1;
      return batch(context);
    },
  });

  return mock;
}
