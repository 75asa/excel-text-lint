import { describe, expect, it } from "vitest";
import { textlintWorkerUrl } from "./worker-url";

describe("textlintWorkerUrl", () => {
  it("base と仮想モジュールのパスから絶対 URL を作る", () => {
    expect(
      textlintWorkerUrl("/excel-text-lint/", "https://example.com/excel-text-lint/taskpane/x.html")
        .href,
    ).toBe("https://example.com/excel-text-lint/textlint/loader-test.js");
  });

  it("base の末尾に / がなくても補う", () => {
    expect(textlintWorkerUrl("/app", "https://example.com/").href).toBe(
      "https://example.com/app/textlint/loader-test.js",
    );
  });

  it("絶対 URL の base はそのまま使う", () => {
    expect(textlintWorkerUrl("https://cdn.example.com/", "https://example.com/").href).toBe(
      "https://cdn.example.com/textlint/loader-test.js",
    );
  });
});
