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
// ブック全体のフェイク（#13、#17）
// ---------------------------------------------------------------------------

/** セルの値。`null` と `""` は空のセル。数値は日付も含む（Excel では日付も数値）。 */
export type FakeCellValue = string | number | boolean | null | { error: string };

/** ワークシートのフェイクの中身。 */
export interface FakeSheet {
  name: string;
  /** `Worksheet.id`。省略すると `{sheet-<添字>}`。 */
  id?: string;
  visibility?: "Visible" | "Hidden" | "VeryHidden";
  /** A1 から始まる値（行 × 列）。数式のセルには計算結果を入れる。 */
  values?: FakeCellValue[][];
  /**
   * 大きなシート用。`values` の代わりに、大きさと値を返す関数で指定する（ベンチマーク用）。
   * 使用範囲は A1 から `rowCount` × `columnCount` とみなす。
   */
  cells?: { rowCount: number; columnCount: number; get(row: number, col: number): FakeCellValue };
  /** 数式（A1 形式の番地 → `=...`）。 */
  formulas?: Record<string, string>;
  /** 非表示の行（0 始まり）。フィルターで隠れた行もこれで表す。 */
  hiddenRows?: number[];
  /** 非表示の列（0 始まり）。 */
  hiddenColumns?: number[];
  /** 結合セル（例: `B2:C3`）。本物と同じく、左上以外のセルの値は空になる。 */
  merged?: string[];
}

export interface FakeWorkbook {
  sheets: FakeSheet[];
  /** アクティブシートの名前。省略すると最初のシート。 */
  activeSheet?: string;
  /** 選択範囲（A1 形式。`B2:C5`、`E:E`、`3:5` など）。シートはアクティブシート。 */
  selection?: string[];
  /** true のとき `getSpecialCells` を使った sync を失敗させる（ExcelApi 1.9 未満などの再現）。 */
  specialCellsUnsupported?: boolean;
  /**
   * 1 回の sync で読める `values` のセル数の上限。超えると `ResponsePayloadSizeLimitExceeded` で失敗する
   * （Excel on the web の応答サイズの上限の再現）。
   */
  maxCellsPerSync?: number;
}

export interface WorkbookMock extends ExcelMock {
  /** `values` を読んだセルの数の合計。 */
  loadedCells: number;
  /** sync のたびに呼ばれる（sync の回数を渡す）。キャンセルのテスト用。 */
  onSync?: (syncCount: number) => void;
}

interface FakeRect {
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
}

const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;

/** A1 形式の番地（`B2`、`B2:C5`、`E:E`、`3:5`）を長方形にする。 */
export function parseAddress(address: string): FakeRect {
  const [start, end = start] = address.split(":") as [string, string?];
  const a = parseCell(start);
  const b = parseCell(end);
  const row = a.row ?? 0;
  const col = a.col ?? 0;
  const rowEnd = b.row ?? MAX_ROWS - 1;
  const colEnd = b.col ?? MAX_COLUMNS - 1;
  return { row, col, rowCount: rowEnd - row + 1, colCount: colEnd - col + 1 };
}

function parseCell(ref: string): { row?: number; col?: number } {
  const match = /^([A-Z]*)(\d*)$/.exec(ref.replaceAll("$", "").toUpperCase());
  if (!match) throw new Error(`番地を読めません: ${ref}`);
  const [, letters = "", digits = ""] = match;
  let col = 0;
  for (const ch of letters) col = col * 26 + (ch.charCodeAt(0) - 64);
  return {
    ...(letters ? { col: col - 1 } : {}),
    ...(digits ? { row: Number(digits) - 1 } : {}),
  };
}

function cellAddress(row: number, col: number): string {
  let name = "";
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return `${name}${row + 1}`;
}

class SheetModel {
  readonly id: string;
  readonly hiddenRows: Set<number>;
  readonly hiddenColumns: Set<number>;
  /** 結合セルのうち、左上以外のセル（`行,列`）。 */
  readonly covered = new Set<string>();

  constructor(
    readonly spec: FakeSheet,
    readonly position: number,
  ) {
    this.id = spec.id ?? `{sheet-${position}}`;
    this.hiddenRows = new Set(spec.hiddenRows);
    this.hiddenColumns = new Set(spec.hiddenColumns);
    for (const address of spec.merged ?? []) {
      const rect = parseAddress(address);
      for (let r = 0; r < rect.rowCount; r++) {
        for (let c = 0; c < rect.colCount; c++) {
          if (r > 0 || c > 0) this.covered.add(`${rect.row + r},${rect.col + c}`);
        }
      }
    }
  }

  value(row: number, col: number): FakeCellValue {
    if (this.covered.size > 0 && this.covered.has(`${row},${col}`)) return null;
    const { cells, values } = this.spec;
    if (cells) {
      return row < cells.rowCount && col < cells.columnCount ? cells.get(row, col) : null;
    }
    return values?.[row]?.[col] ?? null;
  }

  formula(row: number, col: number): string | undefined {
    return this.spec.formulas?.[cellAddress(row, col)];
  }

  /** 値のあるセルを囲む長方形（`getUsedRange(true)`）。 */
  usedRect(): FakeRect | undefined {
    const { cells, values = [] } = this.spec;
    if (cells) return { row: 0, col: 0, rowCount: cells.rowCount, colCount: cells.columnCount };
    let top = Infinity;
    let left = Infinity;
    let bottom = -1;
    let right = -1;
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < (values[r]?.length ?? 0); c++) {
        if (isEmpty(this.value(r, c))) continue;
        top = Math.min(top, r);
        left = Math.min(left, c);
        bottom = Math.max(bottom, r);
        right = Math.max(right, c);
      }
    }
    if (bottom < 0) return undefined;
    return { row: top, col: left, rowCount: bottom - top + 1, colCount: right - left + 1 };
  }
}

function isEmpty(value: FakeCellValue): boolean {
  return value === null || value === "";
}

function valueType(value: FakeCellValue): string {
  if (isEmpty(value)) return "Empty";
  if (typeof value === "string") return "String";
  if (typeof value === "number") return "Double";
  if (typeof value === "boolean") return "Boolean";
  return "Error";
}

function plainValue(value: FakeCellValue): string | number | boolean {
  if (value === null) return "";
  if (typeof value === "object") return value.error;
  return value;
}

/** sync の状態（load 待ちのオブジェクトと、失敗させる理由）。 */
class SyncState {
  readonly dirty = new Set<FakeObject>();
  failure: Error | undefined;

  constructor(readonly specialCellsUnsupported: boolean) {}
}

/** `load()` と `sync()` を経たプロパティだけを読めるフェイクの基底クラス。 */
class FakeObject {
  readonly #pending = new Set<string>();
  readonly #loaded = new Set<string>();

  constructor(protected readonly state: SyncState) {}

  load(names?: string | string[]): this {
    const list = names === undefined ? [] : Array.isArray(names) ? names : names.split(",");
    for (const name of list) this.onLoad(name.trim());
    return this;
  }

  protected onLoad(name: string): void {
    this.#pending.add(name);
    this.state.dirty.add(this);
  }

  pendingNames(): ReadonlySet<string> {
    return this.#pending;
  }

  commit(): void {
    for (const name of this.#pending) this.#loaded.add(name);
    this.#pending.clear();
  }

  discard(): void {
    this.#pending.clear();
  }

  protected read<T>(name: string, value: () => T): T {
    if (!this.#loaded.has(name)) {
      throw new Error(
        `PropertyNotLoaded: "${name}" を読む前に load() と context.sync() を呼んでください`,
      );
    }
    return value();
  }
}

class FakeCollection<T extends FakeObject> extends FakeObject {
  constructor(
    state: SyncState,
    readonly all: T[],
  ) {
    super(state);
  }

  protected override onLoad(name: string): void {
    super.onLoad("items");
    if (name.startsWith("items/")) for (const item of this.all) item.load(name.slice(6));
  }

  get items(): T[] {
    return this.read("items", () => this.all);
  }
}

class FakeRange extends FakeObject {
  constructor(
    state: SyncState,
    readonly model: SheetModel,
    readonly rect: FakeRect,
    readonly isNullObject = false,
  ) {
    super(state);
  }

  get rowIndex(): number {
    return this.read("rowIndex", () => this.rect.row);
  }
  get columnIndex(): number {
    return this.read("columnIndex", () => this.rect.col);
  }
  get rowCount(): number {
    return this.read("rowCount", () => this.rect.rowCount);
  }
  get columnCount(): number {
    return this.read("columnCount", () => this.rect.colCount);
  }
  get values(): (string | number | boolean)[][] {
    return this.read("values", () => this.#map((v) => plainValue(v)));
  }
  get valueTypes(): string[][] {
    return this.read("valueTypes", () => this.#map((v) => valueType(v)));
  }
  get formulas(): (string | number | boolean)[][] {
    return this.read("formulas", () =>
      this.#map((v, row, col) => this.model.formula(row, col) ?? plainValue(v)),
    );
  }
  get rowHidden(): boolean | null {
    return this.read("rowHidden", () =>
      allOrNone(this.rect.row, this.rect.rowCount, this.model.hiddenRows),
    );
  }
  get columnHidden(): boolean | null {
    return this.read("columnHidden", () =>
      allOrNone(this.rect.col, this.rect.colCount, this.model.hiddenColumns),
    );
  }

  getRow(index: number): FakeRange {
    return new FakeRange(this.state, this.model, {
      ...this.rect,
      row: this.rect.row + index,
      rowCount: 1,
    });
  }

  getColumn(index: number): FakeRange {
    return new FakeRange(this.state, this.model, {
      ...this.rect,
      col: this.rect.col + index,
      colCount: 1,
    });
  }

  getSpecialCells(cellType: string): { areas: FakeCollection<FakeRange> } {
    if (cellType !== "Visible") throw new Error(`フェイクは ${cellType} に対応していません`);
    const rows = runs(this.rect.row, this.rect.rowCount, this.model.hiddenRows);
    const cols = runs(this.rect.col, this.rect.colCount, this.model.hiddenColumns);
    const areas = rows.flatMap(([row, rowCount]) =>
      cols.map(
        ([col, colCount]) =>
          new FakeRange(this.state, this.model, { row, col, rowCount, colCount }),
      ),
    );
    if (this.state.specialCellsUnsupported) {
      this.state.failure ??= Object.assign(new Error("getSpecialCells に対応していません"), {
        code: "ApiNotFound",
      });
    }
    return { areas: new FakeCollection(this.state, areas) };
  }

  #map<T>(fn: (value: FakeCellValue, row: number, col: number) => T): T[][] {
    const { row, col, rowCount, colCount } = this.rect;
    return Array.from({ length: rowCount }, (_, r) =>
      Array.from({ length: colCount }, (_, c) =>
        fn(this.model.value(row + r, col + c), row + r, col + c),
      ),
    );
  }
}

function allOrNone(start: number, count: number, hidden: ReadonlySet<number>): boolean | null {
  if (hidden.size === 0) return false;
  let n = 0;
  for (let i = start; i < start + count; i++) if (hidden.has(i)) n++;
  return n === 0 ? false : n === count ? true : null;
}

/** 非表示でない添字の連続（`[開始, 個数]`）。 */
function runs(start: number, count: number, hidden: ReadonlySet<number>): [number, number][] {
  const result: [number, number][] = [];
  for (let i = start; i < start + count; i++) {
    if (hidden.has(i)) continue;
    const last = result.at(-1);
    if (last && last[0] + last[1] === i) last[1] += 1;
    else result.push([i, 1]);
  }
  return result;
}

class FakeWorksheet extends FakeObject {
  constructor(
    state: SyncState,
    readonly model: SheetModel,
  ) {
    super(state);
  }

  get id(): string {
    return this.read("id", () => this.model.id);
  }
  get name(): string {
    return this.read("name", () => this.model.spec.name);
  }
  get visibility(): string {
    return this.read("visibility", () => this.model.spec.visibility ?? "Visible");
  }
  get position(): number {
    return this.read("position", () => this.model.position);
  }

  getUsedRangeOrNullObject(valuesOnly?: boolean): FakeRange {
    if (valuesOnly !== true) throw new Error("フェイクは getUsedRange(true) だけに対応しています");
    const rect = this.model.usedRect();
    return rect
      ? new FakeRange(this.state, this.model, rect)
      : new FakeRange(this.state, this.model, { row: 0, col: 0, rowCount: 1, colCount: 1 }, true);
  }

  getRangeByIndexes(row: number, col: number, rowCount: number, colCount: number): FakeRange {
    return new FakeRange(this.state, this.model, { row, col, rowCount, colCount });
  }
}

/**
 * `globalThis.Excel` を、ブック全体を持つフェイクに差し替える（collect のテストとベンチマーク用）。
 *
 * 対応している API: `workbook.worksheets`（`items` / `getActiveWorksheet`）、`workbook.getSelectedRanges()`、
 * `Worksheet.getUsedRangeOrNullObject(true)` / `getRangeByIndexes`、
 * `Range` の `values` / `valueTypes` / `formulas` / `rowHidden` / `columnHidden` / `getRow` / `getColumn` /
 * `getSpecialCells("Visible")`。
 */
export function stubWorkbook(workbook: FakeWorkbook): WorkbookMock {
  const mock: WorkbookMock = { runCount: 0, syncCount: 0, loadedCells: 0 };
  const models = workbook.sheets.map((sheet, i) => new SheetModel(sheet, i));

  const makeContext = () => {
    const state = new SyncState(workbook.specialCellsUnsupported ?? false);
    const sheets = models.map((model) => new FakeWorksheet(state, model));
    const active =
      sheets.find((sheet) => sheet.model.spec.name === workbook.activeSheet) ?? sheets[0];
    if (!active) throw new Error("シートがありません");

    return {
      workbook: {
        worksheets: Object.assign(new FakeCollection(state, sheets), {
          getActiveWorksheet: () => active,
        }),
        getSelectedRanges: () => ({
          worksheet: active,
          areas: new FakeCollection(
            state,
            (workbook.selection ?? ["A1"]).map(
              (address) => new FakeRange(state, active.model, parseAddress(address)),
            ),
          ),
        }),
      },
      sync: async () => {
        mock.syncCount += 1;
        mock.onSync?.(mock.syncCount);
        const objects = [...state.dirty];
        state.dirty.clear();
        let cells = 0;
        for (const object of objects) {
          if (object instanceof FakeRange && object.pendingNames().has("values")) {
            cells += object.rect.rowCount * object.rect.colCount;
          }
        }
        const failure =
          state.failure ??
          (workbook.maxCellsPerSync !== undefined && cells > workbook.maxCellsPerSync
            ? Object.assign(new Error("応答が大きすぎます"), {
                code: "ResponsePayloadSizeLimitExceeded",
              })
            : undefined);
        state.failure = undefined;
        if (failure) {
          for (const object of objects) object.discard();
          throw failure;
        }
        mock.loadedCells += cells;
        for (const object of objects) object.commit();
      },
    };
  };

  vi.stubGlobal("Excel", {
    run: async <R>(batch: (ctx: ReturnType<typeof makeContext>) => Promise<R>): Promise<R> => {
      mock.runCount += 1;
      return batch(makeContext());
    },
  });

  return mock;
}
