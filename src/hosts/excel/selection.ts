/** 選択範囲の 1 セル分のテキスト。 */
export interface CellText {
  /** A1 形式のアドレス（例: `Sheet1!B3`）。 */
  address: string;
  /** セルに表示されている文字列。 */
  text: string;
}

/**
 * 選択範囲のセルのうち、空でないものを行優先で返す。
 *
 * 表示形式を適用した文字列（`Range.text`）を読む。数式のセルは計算結果になる。
 */
export async function readSelectedCells(): Promise<CellText[]> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["text", "rowIndex", "columnIndex"]);
    const sheet = range.worksheet;
    sheet.load("name");
    await context.sync();

    const cells: CellText[] = [];
    range.text.forEach((row, r) => {
      row.forEach((text, c) => {
        if (text === "") return;
        cells.push({
          address: `${sheet.name}!${columnName(range.columnIndex + c)}${range.rowIndex + r + 1}`,
          text,
        });
      });
    });
    return cells;
  });
}

/** 0 始まりの列番号を列名（A, B, …, Z, AA, …）に変換する。 */
function columnName(index: number): string {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}
