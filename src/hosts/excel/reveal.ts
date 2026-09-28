import type { ExcelLocation } from "../../core/locations";
import type { RevealPrecision, TextUnit } from "../../core/types";

/**
 * TextUnit のセルのシートをアクティブにし、セルを選択する（#15）。
 *
 * Excel の API ではセル内の文字位置までは選択できないので、うまくいっても精度は `unit`。
 * 次のときは何もせずに `none` を返す。
 *
 * - シートが削除された・名前が変わった
 * - セルのテキストが collect したときと違う（行や列の挿入・削除でセルがずれた、書き換えられた）。
 *   別のセルを選んでしまうのを避けるため
 * - Excel の API がエラーを返した（セルの編集中など）
 */
export async function revealCell(unit: TextUnit<ExcelLocation>): Promise<RevealPrecision> {
  const { sheet: sheetName, address } = unit.location;
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      sheet.load("isNullObject");
      await context.sync();
      if (sheet.isNullObject) return "none";

      const range = sheet.getRange(address);
      range.load("text");
      await context.sync();
      if (range.text[0]?.[0] !== unit.text) return "none";

      sheet.activate();
      range.select();
      await context.sync();
      return "unit";
    });
  } catch (error) {
    console.warn("セルへの移動に失敗しました", error);
    return "none";
  }
}
