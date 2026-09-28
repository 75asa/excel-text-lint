/**
 * Excel のセルのテキストを集める（#13、#17）。
 *
 * 方針（詳しくは `src/hosts/excel/README.md`）:
 * - 範囲は `selection`（選択範囲。複数の領域も可）/ `sheet`（アクティブシート）/ `workbook`（ブック全体）
 * - 使用範囲（`getUsedRangeOrNullObject(true)`、値のあるセルだけ）の中だけを読む
 * - 値の種類が文字列（`valueTypes` が `String`）のセルだけを対象にする。数値・日付・真偽値・エラーは除く
 * - 数式のセルは、計算結果が文字列なら既定で対象にする（`location.isFormula` が true）
 * - 結合セルは左上のセルにだけ値があるので、自然に左上だけが対象になる
 * - 非表示の行・列・シートは既定で除く（`includeHidden: true` で含める）
 * - 範囲をチャンク（既定で 5,000 セル）に分け、チャンクごとに `values` / `valueTypes` / `formulas` をまとめて読む。
 *   `context.sync()` はチャンクごとに 1 回（非表示の行・列がまざるチャンクだけもう 1 回）
 */

import type { ExcelLocation } from "../../core/locations";
import type { CollectOptions, Scope, TextUnit } from "../../core/types";
import { columnName } from "./selection";

/** Excel で使える collect の範囲。 */
export const EXCEL_SCOPES = ["selection", "sheet", "workbook"] as const satisfies readonly Scope[];

export type ExcelScope = (typeof EXCEL_SCOPES)[number];

export function isExcelScope(scope: Scope): scope is ExcelScope {
  return (EXCEL_SCOPES as readonly Scope[]).includes(scope);
}

/** 1 回の `context.sync()` で読むセルの数の既定値。 */
export const DEFAULT_CHUNK_CELLS = 5_000;

/** collect の設定（アダプタに持たせて、UI から変えられるようにする）。 */
export interface ExcelCollectSettings {
  /** 非表示の行・列・シート（フィルターで隠れた行も含む）も対象にするか。既定は false（除く）。 */
  includeHidden?: boolean;
  /** 数式のセル（計算結果が文字列のもの）を対象にするか。既定は `include`。 */
  formulas?: "include" | "exclude";
  /** 1 回の `context.sync()` で読むセルの数の上限。既定は {@link DEFAULT_CHUNK_CELLS}。 */
  chunkCells?: number;
}

export interface ExcelCollectOptions extends CollectOptions, ExcelCollectSettings {
  /**
   * チャンクを読み終えるたびに、そのチャンクで見つかった TextUnit を渡す（空のチャンクでは呼ばない）。
   * collect の完了を待たずに lint を始めたいとき用。
   */
  onChunk?: (units: TextUnit<ExcelLocation>[]) => void;
}

/** TextUnit の id。シート ID と番地から作るので、シート名を変えても、collect をやり直しても同じになる。 */
export function excelUnitId(sheetId: string, address: string): string {
  return `${sheetId}!${address}`;
}

interface Rect {
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
}

interface SheetRef {
  id: string;
  name: string;
  proxy: Excel.Worksheet;
}

/** あるシートの中の、読む範囲。 */
interface Block extends Rect {
  sheet: SheetRef;
}

/**
 * 行・列ごとの表示状態（`undefined` の配列はすべて表示）。
 * `null` はチャンク全体が非表示。
 */
type Visibility = { rows?: boolean[]; cols?: boolean[] } | null;

const RECT_PROPS = ["rowIndex", "columnIndex", "rowCount", "columnCount"];
const CELL_PROPS = ["values", "valueTypes", "formulas"];

/**
 * 範囲の中の、文字列のセルを TextUnit にして返す。並びはシートの順、その中ではチャンクごとの行優先。
 *
 * `signal` で中断すると `signal.reason` で reject する（`context.sync()` の合間で確かめる）。
 */
export async function collectExcel(
  scope: ExcelScope,
  options: ExcelCollectOptions = {},
): Promise<TextUnit<ExcelLocation>[]> {
  const { signal } = options;
  signal?.throwIfAborted();
  const chunkCells = Math.max(1, Math.floor(options.chunkCells ?? DEFAULT_CHUNK_CELLS));
  const includeHidden = options.includeHidden ?? false;
  const includeFormulas = (options.formulas ?? "include") === "include";

  return Excel.run(async (context) => {
    const blocks = await planBlocks(context, scope, includeHidden);
    signal?.throwIfAborted();

    const units: TextUnit<ExcelLocation>[] = [];
    // 選択範囲の領域が重なっているときに、同じセルを 2 回返さないため
    const seen = new Set<string>();
    for (const batch of toBatches(splitBlocks(blocks, chunkCells), chunkCells)) {
      const found = await readBatch(context, batch, { includeHidden, includeFormulas, signal });
      signal?.throwIfAborted();
      const fresh: TextUnit<ExcelLocation>[] = [];
      for (const unit of found) {
        if (seen.has(unit.id)) continue;
        seen.add(unit.id);
        fresh.push(unit);
        units.push(unit);
      }
      if (fresh.length > 0) options.onChunk?.(fresh);
      options.onProgress?.({ collected: units.length });
    }
    options.onProgress?.({ collected: units.length, total: units.length });
    return units;
  });
}

// ---------------------------------------------------------------------------
// 読む範囲を決める
// ---------------------------------------------------------------------------

async function planBlocks(
  context: Excel.RequestContext,
  scope: ExcelScope,
  includeHidden: boolean,
): Promise<Block[]> {
  const { workbook } = context;

  if (scope === "workbook") {
    const sheets = workbook.worksheets;
    sheets.load("items/id,items/name,items/visibility,items/position");
    await context.sync();
    const targets = sheets.items
      .filter((sheet) => includeHidden || sheet.visibility === "Visible")
      .sort((a, b) => a.position - b.position);
    const used = targets.map((sheet) => loadRect(sheet.getUsedRangeOrNullObject(true)));
    await context.sync();
    return targets.flatMap((sheet, i) => {
      const range = used[i]!;
      return range.isNullObject ? [] : [{ sheet: toSheetRef(sheet), ...rectOf(range) }];
    });
  }

  if (scope === "sheet") {
    const sheet = workbook.worksheets.getActiveWorksheet();
    sheet.load(["id", "name"]);
    const used = loadRect(sheet.getUsedRangeOrNullObject(true));
    await context.sync();
    return used.isNullObject ? [] : [{ sheet: toSheetRef(sheet), ...rectOf(used) }];
  }

  // 選択範囲。Ctrl を押しながらの複数の領域にも対応するため getSelectedRanges を使う
  const selected = workbook.getSelectedRanges();
  selected.areas.load(RECT_PROPS.map((name) => `items/${name}`));
  const sheet = selected.worksheet;
  sheet.load(["id", "name"]);
  // 列全体などの大きな選択に備えて、使用範囲との重なりだけを読む
  const used = loadRect(sheet.getUsedRangeOrNullObject(true));
  await context.sync();
  if (used.isNullObject) return [];
  const usedRect = rectOf(used);
  const ref = toSheetRef(sheet);
  return selected.areas.items.flatMap((area) => {
    const rect = intersect(rectOf(area), usedRect);
    return rect ? [{ sheet: ref, ...rect }] : [];
  });
}

function loadRect(range: Excel.Range): Excel.Range {
  range.load(RECT_PROPS);
  return range;
}

function rectOf(range: Excel.Range): Rect {
  return {
    row: range.rowIndex,
    col: range.columnIndex,
    rowCount: range.rowCount,
    colCount: range.columnCount,
  };
}

function toSheetRef(sheet: Excel.Worksheet): SheetRef {
  return { id: sheet.id, name: sheet.name, proxy: sheet };
}

export function intersect(a: Rect, b: Rect): Rect | undefined {
  const row = Math.max(a.row, b.row);
  const col = Math.max(a.col, b.col);
  const rowEnd = Math.min(a.row + a.rowCount, b.row + b.rowCount);
  const colEnd = Math.min(a.col + a.colCount, b.col + b.colCount);
  if (rowEnd <= row || colEnd <= col) return undefined;
  return { row, col, rowCount: rowEnd - row, colCount: colEnd - col };
}

/**
 * 範囲を、`chunkCells` 以下のセル数の長方形（チャンク）に分ける。
 * 行で分けるのが基本で、1 行が `chunkCells` を超えるほど幅が広いときだけ列でも分ける。
 */
export function splitBlocks<B extends Rect>(blocks: readonly B[], chunkCells: number): B[] {
  const chunks: B[] = [];
  for (const block of blocks) {
    const cols = Math.min(block.colCount, chunkCells);
    const rows = Math.max(1, Math.floor(chunkCells / cols));
    for (let r = 0; r < block.rowCount; r += rows) {
      for (let c = 0; c < block.colCount; c += cols) {
        chunks.push({
          ...block,
          row: block.row + r,
          col: block.col + c,
          rowCount: Math.min(rows, block.rowCount - r),
          colCount: Math.min(cols, block.colCount - c),
        });
      }
    }
  }
  return chunks;
}

/** 小さなチャンク（小さなシートや選択範囲の領域）を、合計 `chunkCells` 以下になるようにまとめる。 */
function toBatches<B extends Rect>(chunks: readonly B[], chunkCells: number): B[][] {
  const batches: B[][] = [];
  let current: B[] = [];
  let cells = 0;
  for (const chunk of chunks) {
    const size = chunk.rowCount * chunk.colCount;
    if (current.length > 0 && cells + size > chunkCells) {
      batches.push(current);
      current = [];
      cells = 0;
    }
    current.push(chunk);
    cells += size;
  }
  if (current.length > 0) batches.push(current);
  return batches;
}

// ---------------------------------------------------------------------------
// チャンクを読む
// ---------------------------------------------------------------------------

interface ReadOptions {
  includeHidden: boolean;
  includeFormulas: boolean;
  signal: AbortSignal | undefined;
}

/**
 * チャンクの束を 1 回の `context.sync()` で読む。
 *
 * 応答が大きすぎて失敗したとき（Excel on the web の上限など）は、チャンクを半分にして読み直す。
 */
async function readBatch(
  context: Excel.RequestContext,
  batch: readonly Block[],
  options: ReadOptions,
): Promise<TextUnit<ExcelLocation>[]> {
  let ranges: Excel.Range[];
  try {
    ranges = batch.map((block) => {
      const range = block.sheet.proxy.getRangeByIndexes(
        block.row,
        block.col,
        block.rowCount,
        block.colCount,
      );
      range.load(options.includeHidden ? CELL_PROPS : [...CELL_PROPS, "rowHidden", "columnHidden"]);
      return range;
    });
    await context.sync();
  } catch (error) {
    const halves = batch.flatMap(halve);
    if (!isPayloadTooLarge(error) || halves.length === batch.length) throw error;
    options.signal?.throwIfAborted();
    const middle = Math.ceil(halves.length / 2);
    const first = await readBatch(context, halves.slice(0, middle), options);
    options.signal?.throwIfAborted();
    const second = await readBatch(context, halves.slice(middle), options);
    return [...first, ...second];
  }
  options.signal?.throwIfAborted();

  const visibility = options.includeHidden
    ? batch.map((): Visibility => ({}))
    : await resolveVisibility(context, batch, ranges);

  const units: TextUnit<ExcelLocation>[] = [];
  batch.forEach((block, i) => {
    const visible = visibility[i];
    if (visible === null || visible === undefined) return;
    const range = ranges[i]!;
    extractUnits(block, range, visible, options.includeFormulas, units);
  });
  return units;
}

function extractUnits(
  block: Block,
  range: Excel.Range,
  visible: NonNullable<Visibility>,
  includeFormulas: boolean,
  out: TextUnit<ExcelLocation>[],
): void {
  const { values, valueTypes, formulas } = range;
  const colNames: string[] = [];
  for (let c = 0; c < block.colCount; c++) colNames.push(columnName(block.col + c));

  for (let r = 0; r < block.rowCount; r++) {
    if (visible.rows && !visible.rows[r]) continue;
    const valueRow = values[r]!;
    const typeRow = valueTypes[r]!;
    const formulaRow = formulas[r]!;
    for (let c = 0; c < block.colCount; c++) {
      if ((typeRow[c] as string) !== "String") continue;
      if (visible.cols && !visible.cols[c]) continue;
      const text: unknown = valueRow[c];
      if (typeof text !== "string" || text.trim() === "") continue;
      const formula: unknown = formulaRow[c];
      // 定数のセルでは formulas は値と同じ文字列になる。「'=abc」のような「=」で始まる文字列の定数もこれで区別する
      const isFormula = typeof formula === "string" && formula.startsWith("=") && formula !== text;
      if (isFormula && !includeFormulas) continue;

      const row = block.row + r;
      const col = block.col + c;
      const address = `${colNames[c]}${row + 1}`;
      out.push({
        id: excelUnitId(block.sheet.id, address),
        text,
        location: {
          host: "excel",
          sheetId: block.sheet.id,
          sheet: block.sheet.name,
          address,
          row,
          col,
          isFormula,
        },
      });
    }
  }
}

/**
 * チャンクごとに、どの行・列が表示されているかを調べる。
 *
 * `rowHidden` / `columnHidden` は、すべて非表示なら true、すべて表示なら false、まざっていれば null。
 * まざっているチャンクだけ、表示されているセル（`getSpecialCells("Visible")`）の領域を読む。
 * それが使えない（ExcelApi 1.9 未満や、領域が多すぎる）ときは、行・列ごとに読む。
 */
async function resolveVisibility(
  context: Excel.RequestContext,
  batch: readonly Block[],
  ranges: readonly Excel.Range[],
): Promise<Visibility[]> {
  const result: Visibility[] = ranges.map((range) =>
    range.rowHidden === true || range.columnHidden === true ? null : {},
  );
  const mixed = ranges.flatMap((range, i) =>
    result[i] !== null && (range.rowHidden === null || range.columnHidden === null) ? [i] : [],
  );
  if (mixed.length === 0) return result;

  try {
    const visibleAreas = mixed.map((i) => {
      const areas = ranges[i]!.getSpecialCells("Visible").areas;
      areas.load(RECT_PROPS.map((name) => `items/${name}`));
      return areas;
    });
    await context.sync();
    mixed.forEach((i, k) => {
      const block = batch[i]!;
      const rows = new Array<boolean>(block.rowCount).fill(false);
      const cols = new Array<boolean>(block.colCount).fill(false);
      for (const area of visibleAreas[k]!.items) {
        // 非表示は行単位・列単位なので、表示されているセルは「表示の行 × 表示の列」になる
        fillRange(rows, area.rowIndex - block.row, area.rowCount);
        fillRange(cols, area.columnIndex - block.col, area.columnCount);
      }
      result[i] = { rows, cols };
    });
  } catch {
    const probes = mixed.map((i) => {
      const range = ranges[i]!;
      const block = batch[i]!;
      return {
        rows:
          range.rowHidden === null
            ? Array.from({ length: block.rowCount }, (_, r) => {
                const row = range.getRow(r);
                row.load("rowHidden");
                return row;
              })
            : undefined,
        cols:
          range.columnHidden === null
            ? Array.from({ length: block.colCount }, (_, c) => {
                const col = range.getColumn(c);
                col.load("columnHidden");
                return col;
              })
            : undefined,
      };
    });
    await context.sync();
    mixed.forEach((i, k) => {
      const probe = probes[k]!;
      const visible: NonNullable<Visibility> = {};
      if (probe.rows) visible.rows = probe.rows.map((row) => row.rowHidden !== true);
      if (probe.cols) visible.cols = probe.cols.map((col) => col.columnHidden !== true);
      result[i] = visible;
    });
  }
  return result;
}

function fillRange(flags: boolean[], start: number, count: number): void {
  const from = Math.max(0, start);
  const to = Math.min(flags.length, start + count);
  for (let i = from; i < to; i++) flags[i] = true;
}

function halve(block: Block): Block[] {
  if (block.rowCount > 1) {
    const top = Math.ceil(block.rowCount / 2);
    return [
      { ...block, rowCount: top },
      { ...block, row: block.row + top, rowCount: block.rowCount - top },
    ];
  }
  if (block.colCount > 1) {
    const left = Math.ceil(block.colCount / 2);
    return [
      { ...block, colCount: left },
      { ...block, col: block.col + left, colCount: block.colCount - left },
    ];
  }
  return [block];
}

function isPayloadTooLarge(error: unknown): boolean {
  const code = (error as { code?: unknown } | null)?.code;
  return typeof code === "string" && code.includes("PayloadSizeLimitExceeded");
}
