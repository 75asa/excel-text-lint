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

// ---------------------------------------------------------------------------
// ブック全体のフェイク（reveal / highlight / applyFix のテスト用）
// ---------------------------------------------------------------------------

/** フェイクのセル 1 つ分。 */
export interface FakeCell {
  /** `Range.values` の値。 */
  value: string | number | boolean;
  /** `Range.formulas` の値。省略すると `value` と同じ（数式のないセル）。 */
  formula?: string;
  /** `Range.text`（表示形式を適用した文字列）。省略すると `String(value)`。 */
  text?: string;
}

/** フェイクの条件付き書式（Custom 型だけ）。 */
export interface FakeConditionalFormat {
  id: string;
  type: "Custom";
  /** シート名を含まない A1 形式のアドレス。 */
  address: string;
  formula: string;
  fillColor: string | null;
  priority: number;
}

/** フェイクのワークシート。テストから中身を直接読み書きしてよい。 */
export interface FakeSheet {
  name: string;
  /** キーはシート名を含まない A1 形式のアドレス（例: `B3`）。 */
  cells: Record<string, FakeCell>;
  conditionalFormats: FakeConditionalFormat[];
}

export interface ExcelWorkbookMock extends ExcelMock {
  sheets: FakeSheet[];
  /** `Worksheet.activate()` で最後にアクティブにしたシート名。 */
  activeSheet: string | null;
  /** `Range.select()` で最後に選択したセル（`Sheet1!B3` の形）。 */
  selected: string | null;
  /** `Range.values` への書き込みの記録（`Sheet1!B3` の形のアドレスと値）。 */
  writes: { address: string; values: unknown[][] }[];
  /** 次の `context.sync()` で投げる例外。1 回投げたら消える。 */
  failNextSync: Error | null;
  sheet(name: string): FakeSheet;
}

/**
 * `globalThis.Excel` を、複数のシートを持つブックのフェイクに差し替える。
 *
 * 1 セルのアドレス（`B3`）と、引数なしの `Worksheet.getRange()`（シート全体）だけを扱う。
 * 書き込み（`values` の代入・条件付き書式の追加や削除・activate / select）は呼んだ時点でフェイクに反映する。
 * 読み取りは `stubExcel` と同じく、`load()` と `context.sync()` を経てからでないと例外を投げる。
 */
export function stubExcelWorkbook(
  sheets: { name: string; cells?: Record<string, FakeCell> }[],
): ExcelWorkbookMock {
  const mock: ExcelWorkbookMock = {
    runCount: 0,
    syncCount: 0,
    sheets: sheets.map((s) => ({ name: s.name, cells: { ...s.cells }, conditionalFormats: [] })),
    activeSheet: null,
    selected: null,
    writes: [],
    failNextSync: null,
    sheet(name) {
      const found = this.sheets.find((s) => s.name === name);
      if (!found) throw new Error(`シート「${name}」はありません`);
      return found;
    },
  };
  let nextId = 1;

  /** load 済みで sync 待ちのものと、読めるもの。 */
  const pending = new Set<string>();
  const loaded = new Set<string>();
  let objectSeq = 0;
  const loadable = (): {
    load: (names: string | string[]) => void;
    read: <T>(name: string, value: () => T) => T;
  } => {
    const key = `o${objectSeq++}`;
    return {
      load: (names) => {
        for (const name of Array.isArray(names) ? names : names.split(",")) {
          pending.add(`${key}.${name.trim()}`);
        }
      },
      read: (name, value) => {
        if (!loaded.has(`${key}.${name}`)) {
          throw new Error(
            `PropertyNotLoaded: "${name}" を読む前に load() と context.sync() を呼んでください`,
          );
        }
        return value();
      },
    };
  };

  const findSheet = (name: string) => mock.sheets.find((s) => s.name === name);
  const liveSheet = (name: string) => {
    const sheet = findSheet(name);
    if (!sheet) throw new Error(`ItemNotFound: シート「${name}」は削除されています`);
    return sheet;
  };

  const makeConditionalFormat = (sheetName: string, cf: FakeConditionalFormat) => {
    const l = loadable();
    const rule = loadable();
    const custom = {
      rule: {
        load: rule.load,
        get formula() {
          return rule.read("formula", () => cf.formula);
        },
        set formula(value: string) {
          cf.formula = value;
        },
      },
      format: {
        fill: {
          set color(value: string) {
            cf.fillColor = value;
          },
        },
      },
    };
    return {
      load: l.load,
      get id() {
        return l.read("id", () => cf.id);
      },
      get type() {
        return l.read("type", () => cf.type);
      },
      get priority() {
        return l.read("priority", () => cf.priority);
      },
      set priority(value: number) {
        cf.priority = value;
      },
      set stopIfTrue(_value: boolean) {},
      get custom() {
        return custom;
      },
      delete: () => {
        const list = liveSheet(sheetName).conditionalFormats;
        list.splice(list.indexOf(cf), 1);
      },
    };
  };

  const makeConditionalFormats = (sheetName: string, address: string | null) => {
    const l = loadable();
    let items: ReturnType<typeof makeConditionalFormat>[] = [];
    return {
      load: (names: string | string[]) => {
        l.load("items");
        // `items/type` のような指定は、sync の時点で各要素にも load する
        const props = (Array.isArray(names) ? names : names.split(","))
          .map((n) => n.trim())
          .filter((n) => n.startsWith("items/"))
          .map((n) => n.slice("items/".length));
        pendingCollections.push(() => {
          items = liveSheet(sheetName)
            .conditionalFormats.filter((cf) => address === null || cf.address === address)
            .map((cf) => {
              const item = makeConditionalFormat(sheetName, cf);
              if (props.length > 0) item.load(props);
              return item;
            });
        });
      },
      get items() {
        return l.read("items", () => items);
      },
      add: (type: string) => {
        if (address === null) throw new Error("フェイクはシート全体への条件付き書式の追加に未対応");
        if (type !== "Custom") throw new Error(`フェイクは ${type} の条件付き書式に未対応`);
        const cf: FakeConditionalFormat = {
          id: `cf${nextId++}`,
          type: "Custom",
          address,
          formula: "",
          fillColor: null,
          priority: 0,
        };
        const list = liveSheet(sheetName).conditionalFormats;
        for (const other of list) other.priority += 1;
        list.unshift(cf);
        return makeConditionalFormat(sheetName, cf);
      },
    };
  };
  /** collection の load は、sync の時点の中身で items を作る。 */
  const pendingCollections: (() => void)[] = [];

  const makeRange = (sheetName: string, address: string | null) => {
    const l = loadable();
    const cell = (): FakeCell => {
      if (address === null) throw new Error("フェイクはシート全体の値の読み取りに未対応");
      return liveSheet(sheetName).cells[address] ?? { value: "" };
    };
    const valueType = (value: FakeCell["value"]) =>
      value === ""
        ? "Empty"
        : typeof value === "string"
          ? "String"
          : typeof value === "number"
            ? "Double"
            : "Boolean";
    return {
      load: l.load,
      get text() {
        return l.read("text", () => [[cell().text ?? String(cell().value)]]);
      },
      get values() {
        return l.read("values", () => [[cell().value]]);
      },
      set values(values: unknown[][]) {
        if (address === null) throw new Error("フェイクはシート全体への書き込みに未対応");
        const value = values[0]?.[0];
        if (typeof value !== "string") throw new Error("フェイクは文字列の書き込みだけに対応");
        liveSheet(sheetName).cells[address] = { value };
        mock.writes.push({ address: `${sheetName}!${address}`, values });
      },
      get formulas() {
        return l.read("formulas", () => [[cell().formula ?? cell().value]]);
      },
      get valueTypes() {
        return l.read("valueTypes", () => [[valueType(cell().value)]]);
      },
      select: () => {
        liveSheet(sheetName);
        mock.selected = `${sheetName}!${address}`;
      },
      get conditionalFormats() {
        return makeConditionalFormats(sheetName, address);
      },
    };
  };

  const makeWorksheet = (name: string, isNull: boolean) => {
    const l = loadable();
    return {
      load: l.load,
      get isNullObject() {
        return l.read("isNullObject", () => isNull);
      },
      get name() {
        return l.read("name", () => name);
      },
      activate: () => {
        liveSheet(name);
        mock.activeSheet = name;
      },
      getRange: (address?: string) => {
        if (address !== undefined && !/^[A-Z]+[1-9][0-9]*$/.test(address)) {
          throw new Error(`InvalidArgument: フェイクは 1 セルのアドレスだけに対応（${address}）`);
        }
        return makeRange(name, address ?? null);
      },
    };
  };

  const worksheetsLoad = loadable();
  let worksheetItems: ReturnType<typeof makeWorksheet>[] = [];
  const context = {
    workbook: {
      worksheets: {
        getItemOrNullObject: (name: string) => makeWorksheet(name, !findSheet(name)),
        getItem: (name: string) => makeWorksheet(name, false),
        load: (names: string | string[]) => {
          worksheetsLoad.load("items");
          pendingCollections.push(() => {
            worksheetItems = mock.sheets.map((s) => {
              const ws = makeWorksheet(s.name, false);
              if (names.includes("items/name")) ws.load("name");
              return ws;
            });
          });
        },
        get items() {
          return worksheetsLoad.read("items", () => worksheetItems);
        },
      },
    },
    sync: async () => {
      mock.syncCount += 1;
      const error = mock.failNextSync;
      mock.failNextSync = null;
      if (error) throw error;
      for (const run of pendingCollections.splice(0)) run();
      for (const key of pending) loaded.add(key);
      pending.clear();
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
