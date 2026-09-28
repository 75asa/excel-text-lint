import { describe, expect, it } from "vitest";
import type { ExcelLocation } from "../../core/locations";
import type { TextUnit } from "../../core/types";
import { stubWorkbook } from "../../testing/excel-mock";
import { ExcelAdapter } from "./adapter";
import { collectExcel, excelUnitId, splitBlocks } from "./collect";

/** テストで見やすいように `シート名!番地 → テキスト` にする。 */
function summarize(units: TextUnit<ExcelLocation>[]): string[] {
  return units.map((unit) => `${unit.location.sheet}!${unit.location.address}=${unit.text}`);
}

describe("collectExcel: 範囲ごとの走査", () => {
  const sheets = [
    {
      name: "表紙",
      id: "{A}",
      values: [
        ["タイトル", ""],
        ["", "説明文です。"],
      ],
    },
    {
      name: "明細",
      id: "{B}",
      values: [
        ["品名", "数量", "日付"],
        ["りんご", 3, 45_000],
        ["みかん", 5, 45_001],
      ],
    },
  ];

  it("sheet: アクティブシートの使用範囲の文字列のセルを、行優先で返す", async () => {
    stubWorkbook({ sheets, activeSheet: "明細" });

    const units = await collectExcel("sheet");

    // 数値と日付（Excel では数値）は除く
    expect(summarize(units)).toEqual([
      "明細!A1=品名",
      "明細!B1=数量",
      "明細!C1=日付",
      "明細!A2=りんご",
      "明細!A3=みかん",
    ]);
  });

  it("workbook: すべてのシートをシートの順に走査する", async () => {
    const mock = stubWorkbook({ sheets });

    const units = await collectExcel("workbook");

    expect(summarize(units)).toEqual([
      "表紙!A1=タイトル",
      "表紙!B2=説明文です。",
      "明細!A1=品名",
      "明細!B1=数量",
      "明細!C1=日付",
      "明細!A2=りんご",
      "明細!A3=みかん",
    ]);
    // シートの一覧・使用範囲・セルの値（小さいので 1 回にまとめる）の 3 回
    expect(mock.syncCount).toBe(3);
  });

  it("selection: 選択範囲と使用範囲の重なりだけを読む（列全体の選択でも使用範囲の分だけ）", async () => {
    const mock = stubWorkbook({ sheets, activeSheet: "明細", selection: ["A:A"] });

    const units = await collectExcel("selection");

    expect(summarize(units)).toEqual(["明細!A1=品名", "明細!A2=りんご", "明細!A3=みかん"]);
    expect(mock.loadedCells).toBe(3);
    expect(mock.syncCount).toBe(2);
  });

  it("selection: 複数の領域を順に読み、重なったセルは 1 回だけ返す", async () => {
    stubWorkbook({ sheets, activeSheet: "明細", selection: ["A2:A3", "A1:B2", "Z100"] });

    const units = await collectExcel("selection");

    expect(summarize(units)).toEqual([
      "明細!A2=りんご",
      "明細!A3=みかん",
      "明細!A1=品名",
      "明細!B1=数量",
    ]);
  });

  it("空のシートでは空の配列を返す", async () => {
    const mock = stubWorkbook({ sheets: [{ name: "空", values: [["", null]] }] });

    await expect(collectExcel("sheet")).resolves.toEqual([]);
    await expect(collectExcel("workbook")).resolves.toEqual([]);
    await expect(collectExcel("selection")).resolves.toEqual([]);
    expect(mock.loadedCells).toBe(0);
  });

  it("id はシート ID と番地から作る（シート名に依存しない）", async () => {
    stubWorkbook({ sheets, activeSheet: "表紙" });

    const [unit] = await collectExcel("sheet");

    expect(unit).toEqual({
      id: excelUnitId("{A}", "A1"),
      text: "タイトル",
      location: {
        host: "excel",
        sheetId: "{A}",
        sheet: "表紙",
        address: "A1",
        row: 0,
        col: 0,
        isFormula: false,
      },
    });
    expect(unit?.id).toBe("{A}!A1");
  });
});

describe("collectExcel: セルの種類", () => {
  it("真偽値・エラー・空白だけの文字列は除き、セル内改行はそのまま 1 件にする", async () => {
    stubWorkbook({
      sheets: [
        {
          name: "S",
          values: [[true, { error: "#N/A" }, "   ", "1 行目\n2 行目", "'123 は文字列"]],
        },
      ],
    });

    const units = await collectExcel("sheet");

    expect(summarize(units)).toEqual(["S!D1=1 行目\n2 行目", "S!E1='123 は文字列"]);
  });

  it("数式のセルは、計算結果が文字列なら isFormula: true で含める（formulas: exclude で除く）", async () => {
    const workbook = {
      sheets: [
        {
          name: "S",
          values: [["定数", "=で始まる定数", "結果の文字列", 42]],
          formulas: { B1: "=で始まる定数", C1: '=A1&"の結果"', D1: "=6*7" },
        },
      ],
    };
    stubWorkbook(workbook);

    const units = await collectExcel("sheet");

    expect(units.map((unit) => [unit.location.address, unit.location.isFormula])).toEqual([
      ["A1", false],
      ["B1", false], // 「'=…」と入力した文字列の定数（formulas が値と同じ）
      ["C1", true],
    ]);

    stubWorkbook(workbook);
    const withoutFormulas = await collectExcel("sheet", { formulas: "exclude" });
    expect(withoutFormulas.map((unit) => unit.location.address)).toEqual(["A1", "B1"]);
  });

  it("結合セルは左上のセルだけを対象にする", async () => {
    stubWorkbook({
      sheets: [
        {
          name: "S",
          values: [
            ["見出し", "左上以外", "C1"],
            ["左上以外", "左上以外", ""],
          ],
          merged: ["A1:B2"],
        },
      ],
    });

    const units = await collectExcel("sheet");

    expect(summarize(units)).toEqual(["S!A1=見出し", "S!C1=C1"]);
  });
});

describe("collectExcel: 非表示の行・列・シート", () => {
  const workbook = {
    sheets: [
      {
        name: "表示",
        values: [
          ["A1", "B1", "C1"],
          ["A2", "B2", "C2"],
          ["A3", "B3", "C3"],
        ],
        hiddenRows: [1],
        hiddenColumns: [2],
      },
      { name: "非表示", visibility: "Hidden" as const, values: [["隠し"]] },
      { name: "とても非表示", visibility: "VeryHidden" as const, values: [["隠し"]] },
    ],
  };

  it("既定では非表示の行・列・シートを除く", async () => {
    stubWorkbook(workbook);

    const units = await collectExcel("workbook");

    expect(summarize(units)).toEqual(["表示!A1=A1", "表示!B1=B1", "表示!A3=A3", "表示!B3=B3"]);
  });

  it("includeHidden: true ならすべて含める", async () => {
    stubWorkbook(workbook);

    const units = await collectExcel("workbook", { includeHidden: true });

    expect(units).toHaveLength(9 + 2);
  });

  it("getSpecialCells が使えないときは、行・列ごとに読んで判定する", async () => {
    const mock = stubWorkbook({ ...workbook, specialCellsUnsupported: true });

    const units = await collectExcel("sheet");

    expect(summarize(units)).toEqual(["表示!A1=A1", "表示!B1=B1", "表示!A3=A3", "表示!B3=B3"]);
    // 使用範囲・値・getSpecialCells（失敗）・行と列ごとの表示状態
    expect(mock.syncCount).toBe(4);
  });

  it("非表示の行・列がないチャンクでは、表示状態を読むための sync をしない", async () => {
    const mock = stubWorkbook({ sheets: [{ name: "S", values: [["a", "b"]] }] });

    await collectExcel("sheet");

    expect(mock.syncCount).toBe(2);
  });

  it("すべての行が非表示のチャンクは読み飛ばす", async () => {
    const mock = stubWorkbook({
      sheets: [{ name: "S", values: [["a"], ["b"], ["c"], ["d"]], hiddenRows: [2, 3] }],
      specialCellsUnsupported: true, // 使われないことを確かめる
    });

    const units = await collectExcel("sheet", { chunkCells: 2 });

    expect(summarize(units)).toEqual(["S!A1=a", "S!A2=b"]);
    expect(mock.syncCount).toBe(3);
  });
});

describe("collectExcel: チャンク分割", () => {
  const big = (rowCount: number, columnCount: number) => ({
    sheets: [
      {
        name: "大",
        cells: {
          rowCount,
          columnCount,
          get: (row: number, col: number) => (col === 1 ? row : `r${row}c${col}`),
        },
      },
    ],
  });

  it("chunkCells ごとに sync し、進捗とチャンクを順に知らせる", async () => {
    const mock = stubWorkbook(big(10, 3)); // 30 セルのうち文字列は 20
    const progress: number[] = [];
    const chunks: number[] = [];

    const units = await collectExcel("sheet", {
      chunkCells: 9, // 3 列なので 3 行ずつ → 4 チャンク
      onProgress: ({ collected, total }) => progress.push(total ?? -collected),
      onChunk: (found) => chunks.push(found.length),
    });

    expect(units).toHaveLength(20);
    expect(units.map((unit) => unit.location.address).slice(0, 4)).toEqual([
      "A1",
      "C1",
      "A2",
      "C2",
    ]);
    expect(chunks).toEqual([6, 6, 6, 2]);
    expect(progress).toEqual([-6, -12, -18, -20, 20]);
    expect(mock.syncCount).toBe(1 + 4);
    expect(mock.loadedCells).toBe(30);
  });

  it("1 行が chunkCells を超えるほど幅が広いときは列でも分ける", () => {
    expect(splitBlocks([{ row: 0, col: 0, rowCount: 2, colCount: 5 }], 2)).toEqual([
      { row: 0, col: 0, rowCount: 1, colCount: 2 },
      { row: 0, col: 2, rowCount: 1, colCount: 2 },
      { row: 0, col: 4, rowCount: 1, colCount: 1 },
      { row: 1, col: 0, rowCount: 1, colCount: 2 },
      { row: 1, col: 2, rowCount: 1, colCount: 2 },
      { row: 1, col: 4, rowCount: 1, colCount: 1 },
    ]);
  });

  it("小さなシートはまとめて 1 回の sync で読む", async () => {
    const mock = stubWorkbook({
      sheets: Array.from({ length: 5 }, (_, i) => ({ name: `S${i}`, values: [[`s${i}`]] })),
    });

    const units = await collectExcel("workbook");

    expect(units).toHaveLength(5);
    expect(mock.syncCount).toBe(3);
  });

  it("応答が大きすぎて失敗したら、チャンクを半分にして読み直す", async () => {
    const mock = stubWorkbook({ ...big(8, 3), maxCellsPerSync: 7 });

    const units = await collectExcel("sheet", { chunkCells: 24 });

    expect(units).toHaveLength(16);
    expect(units.map((unit) => unit.location.row)).toEqual(
      Array.from({ length: 8 }, (_, row) => [row, row]).flat(),
    );
    expect(mock.loadedCells).toBe(24);
  });

  it("1 セルでも大きすぎるときはエラーにする", async () => {
    stubWorkbook({ ...big(1, 1), maxCellsPerSync: 0 });

    await expect(collectExcel("sheet")).rejects.toMatchObject({
      code: "ResponsePayloadSizeLimitExceeded",
    });
  });
});

describe("collectExcel: キャンセル", () => {
  const workbook = {
    sheets: [
      {
        name: "S",
        cells: { rowCount: 100, columnCount: 1, get: (row: number) => `行 ${row}` },
      },
    ],
  };

  it("始める前に中断されていたら Excel.run を呼ばない", async () => {
    const mock = stubWorkbook(workbook);
    const controller = new AbortController();
    controller.abort(new DOMException("中止", "AbortError"));

    await expect(collectExcel("sheet", { signal: controller.signal })).rejects.toThrow("中止");
    expect(mock.runCount).toBe(0);
  });

  it("途中で中断されたら、次の sync をせずに signal.reason で reject する", async () => {
    const mock = stubWorkbook(workbook);
    const controller = new AbortController();
    const reason = new DOMException("中止", "AbortError");
    mock.onSync = (count) => {
      if (count === 3) controller.abort(reason);
    };
    const chunks: number[] = [];

    await expect(
      collectExcel("sheet", {
        signal: controller.signal,
        chunkCells: 10,
        onChunk: (units) => chunks.push(units.length),
      }),
    ).rejects.toBe(reason);
    // 使用範囲・1 チャンク目・2 チャンク目（この sync の間に中断）
    expect(mock.syncCount).toBe(3);
    expect(chunks).toEqual([10]);
  });
});

describe("ExcelAdapter.collect", () => {
  it("capabilities の scopes は selection / sheet / workbook", () => {
    expect(new ExcelAdapter().capabilities.scopes).toEqual(["selection", "sheet", "workbook"]);
  });

  it("コンストラクタの設定を collect に渡す", async () => {
    stubWorkbook({
      sheets: [{ name: "S", values: [["a", "b"]], hiddenColumns: [1] }],
    });
    const adapter = new ExcelAdapter({ includeHidden: true });

    await expect(adapter.collect("sheet")).resolves.toHaveLength(2);
    adapter.collectSettings.includeHidden = false;
    await expect(adapter.collect("sheet")).resolves.toHaveLength(1);
  });

  it("対応していない範囲はエラーにする", async () => {
    await expect(new ExcelAdapter().collect("document")).rejects.toThrow("document");
  });
});
