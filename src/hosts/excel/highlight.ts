import type { ExcelLocation } from "../../core/locations";
import type { TextUnit, Violation } from "../../core/types";

/**
 * ハイライトの条件付き書式を見分けるための印。
 *
 * ハイライトは「常に真になる数式」の条件付き書式（Custom）で付ける。数式の中にこの文字列を入れておき、
 * 解除のときは、ブックの中の条件付き書式のうち数式にこの文字列を含むものを消す。
 * 解除のための情報をどこにも保存しなくてよいので、タスクペインを閉じたあとや、ブックを開き直したあとでも解除できる。
 */
export const HIGHLIGHT_MARKER = "excel-text-lint";

/** ハイライトの条件付き書式の数式。文字列を与えた `ISTEXT` なので常に真になり、セルの値には依存しない。 */
export const HIGHLIGHT_FORMULA = `=ISTEXT("${HIGHLIGHT_MARKER}")`;

/** ハイライトの塗りつぶしの色（Excel の「どちらでもない」スタイルの黄色）。 */
export const HIGHLIGHT_FILL_COLOR = "#FFEB9C";

/**
 * 違反のあるセルをハイライトする（#16）。
 *
 * 先に、前回のハイライトをすべて解除する。セルの書式そのものには触れず、セルごとに条件付き書式を 1 つ足す
 * （優先順位はいちばん上）。同じセルに違反が複数あっても 1 つだけ足す。
 * 削除されたシートのセルは飛ばす。
 *
 * @returns ハイライトしたセルの数
 */
export async function highlightCells(
  violations: readonly Violation[],
  units: ReadonlyMap<string, TextUnit<ExcelLocation>>,
): Promise<number> {
  /** シート名 → アドレスの集合 */
  const targets = new Map<string, Set<string>>();
  for (const violation of violations) {
    const unit = units.get(violation.unitId);
    if (unit?.location.host !== "excel") continue;
    const { sheet, address } = unit.location;
    const addresses = targets.get(sheet) ?? new Set<string>();
    addresses.add(address);
    targets.set(sheet, addresses);
  }

  return Excel.run(async (context) => {
    await clearHighlightsIn(context);
    if (targets.size === 0) return 0;

    const sheets = [...targets].map(([name, addresses]) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(name);
      sheet.load("isNullObject");
      return { sheet, addresses };
    });
    await context.sync();

    let count = 0;
    for (const { sheet, addresses } of sheets) {
      if (sheet.isNullObject) continue;
      for (const address of addresses) {
        const conditionalFormat = sheet.getRange(address).conditionalFormats.add("Custom");
        conditionalFormat.custom.rule.formula = HIGHLIGHT_FORMULA;
        conditionalFormat.custom.format.fill.color = HIGHLIGHT_FILL_COLOR;
        conditionalFormat.priority = 0;
        conditionalFormat.stopIfTrue = false;
        count += 1;
      }
    }
    await context.sync();
    return count;
  });
}

/**
 * `highlightCells` で付けたハイライトをすべて解除する（#16）。
 *
 * ブックのすべてのシートから、印（`HIGHLIGHT_MARKER`）の入った条件付き書式を探して消す。
 * ユーザーが付けた条件付き書式やセルの書式には触れない。
 *
 * @returns 消した条件付き書式の数
 */
export async function clearHighlights(): Promise<number> {
  return Excel.run((context) => clearHighlightsIn(context));
}

async function clearHighlightsIn(context: Excel.RequestContext): Promise<number> {
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name");
  await context.sync();

  const collections = worksheets.items.map((sheet) => {
    const conditionalFormats = sheet.getRange().conditionalFormats;
    conditionalFormats.load("items/type");
    return conditionalFormats;
  });
  await context.sync();

  const rules = collections.flatMap((collection) =>
    collection.items
      .filter((cf) => cf.type === "Custom")
      .map((cf) => {
        const rule = cf.custom.rule;
        rule.load("formula");
        return { cf, rule };
      }),
  );
  if (rules.length === 0) return 0;
  await context.sync();

  let count = 0;
  for (const { cf, rule } of rules) {
    if (!rule.formula.includes(HIGHLIGHT_MARKER)) continue;
    cf.delete();
    count += 1;
  }
  if (count > 0) await context.sync();
  return count;
}
