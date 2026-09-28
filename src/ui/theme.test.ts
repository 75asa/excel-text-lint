import { describe, expect, it } from "vitest";
import { isDarkColor, themeFromOffice } from "./theme";

describe("isDarkColor", () => {
  it.each([
    ["#000000", true],
    ["#262626", true],
    ["#333", true],
    ["#ffffff", false],
    ["#F3F2F1", false],
    ["e6e6e6", false],
  ])("%s → %s", (color, dark) => {
    expect(isDarkColor(color)).toBe(dark);
  });

  it("解釈できない色は null", () => {
    expect(isDarkColor("black")).toBeNull();
    expect(isDarkColor("")).toBeNull();
  });
});

describe("themeFromOffice", () => {
  it("Office のテーマの背景色から決める", () => {
    expect(themeFromOffice({ bodyBackgroundColor: "#262626" })).toBe("dark");
    expect(themeFromOffice({ bodyBackgroundColor: "#FFFFFF" })).toBe("light");
  });

  it("取れないときは null（CSS の prefers-color-scheme に任せる）", () => {
    expect(themeFromOffice(undefined)).toBeNull();
    expect(themeFromOffice({})).toBeNull();
    expect(themeFromOffice({ bodyBackgroundColor: "transparent" })).toBeNull();
  });
});
