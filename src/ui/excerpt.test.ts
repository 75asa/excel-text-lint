import { describe, expect, it } from "vitest";
import { buildExcerpt, buildFixPreview } from "./excerpt";

describe("buildExcerpt", () => {
  it("違反箇所と前後の文脈を取り出す", () => {
    expect(buildExcerpt("0123456789abcdefghij", [10, 12], { context: 3 })).toEqual({
      before: "789",
      match: "ab",
      after: "cde",
      clippedStart: true,
      clippedEnd: true,
    });
  });

  it("テキストの端では切り詰めない", () => {
    expect(buildExcerpt("することができる。", [2, 8], { context: 16 })).toEqual({
      before: "する",
      match: "ことができる",
      after: "。",
      clippedStart: false,
      clippedEnd: false,
    });
  });

  it("範囲をテキストの長さに収める", () => {
    const excerpt = buildExcerpt("abc", [2, 10], { context: 1 });
    expect(excerpt).toMatchObject({ before: "b", match: "c", after: "", clippedEnd: false });
  });

  it("文脈の切れ目でサロゲートペアを割らない", () => {
    // "😀" は 2 コード単位。context: 1 だと、前の切れ目がペアの途中に来る
    const text = "😀あいう";
    const excerpt = buildExcerpt(text, [2, 3], { context: 1 });
    expect(excerpt.before).toBe("");
    expect(excerpt.clippedStart).toBe(true);

    const tail = buildExcerpt("あ😀", [0, 1], { context: 2 });
    expect(tail.after).toBe("😀");
    const cut = buildExcerpt("あい😀", [0, 1], { context: 2 });
    expect(cut.after).toBe("い");
    expect(cut.clippedEnd).toBe(true);
  });

  it("singleLine なら改行を ↵ にする", () => {
    expect(buildExcerpt("一行目\n二行目", [4, 7], { singleLine: true })).toMatchObject({
      before: "一行目↵",
      match: "二行目",
    });
  });
});

describe("buildFixPreview", () => {
  it("消える部分と入る部分を分けて返す", () => {
    const preview = buildFixPreview("することができる。", { range: [2, 8], text: "できる" });
    expect(preview).toMatchObject({
      before: "する",
      removed: "ことができる",
      inserted: "できる",
      after: "。",
    });
  });
});
