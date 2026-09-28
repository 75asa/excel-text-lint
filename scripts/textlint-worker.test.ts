import { mkdir, mkdtemp, rm, utimes, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  checkKuromojiDictUrl,
  contentHash,
  hashTextlintWorker,
  isWorkerStale,
  virtualModuleCode,
  WORKER_FILE,
  WORKER_SRC_DIR,
} from "./textlint-worker.ts";

const CDN = "https://cdn.jsdelivr.net/npm/kuromoji@0.1.2/dict";
const loader = `const KUROMOJI_CDN = "${CDN}";\nimportScripts(new URL("./textlint-worker.js", self.location.href).href);\n`;

describe("hashTextlintWorker", () => {
  it("ファイル名にハッシュを付け、ローダーの参照を書き換える", () => {
    const { loader: l, worker: w } = hashTextlintWorker(loader, "worker v1");
    expect(w.fileName).toBe(`textlint/textlint-worker-${contentHash("worker v1")}.js`);
    expect(w.source).toBe("worker v1");
    expect(l.fileName).toMatch(/^textlint\/loader-[0-9a-f]{8}\.js$/);
    expect(l.source).toContain(`new URL("./textlint-worker-${contentHash("worker v1")}.js"`);
    expect(l.source).not.toContain('"./textlint-worker.js"');
    expect(l.fileName).toBe(`textlint/loader-${contentHash(l.source)}.js`);
  });

  it("Worker が変わればローダーの名前も変わる", () => {
    const a = hashTextlintWorker(loader, "worker v1");
    const b = hashTextlintWorker(loader, "worker v2");
    expect(b.worker.fileName).not.toBe(a.worker.fileName);
    expect(b.loader.fileName).not.toBe(a.loader.fileName);
  });

  it("同じ内容なら同じ名前になる", () => {
    expect(hashTextlintWorker(loader, "w")).toEqual(hashTextlintWorker(loader, "w"));
  });

  it("ローダーに Worker への参照が 1 つだけでなければ失敗する", () => {
    expect(() => hashTextlintWorker("importScripts('./other.js')", "w")).toThrow(/0 個/);
    expect(() => hashTextlintWorker(loader + loader, "w")).toThrow(/2 個/);
  });
});

describe("checkKuromojiDictUrl", () => {
  it("ローダー・Worker・kuromoji のバージョンが一致すれば URL を返す", () => {
    expect(checkKuromojiDictUrl(loader, `fetch("${CDN}/base.dat.gz")`, "0.1.2")).toBe(CDN);
  });

  it("kuromoji のバージョンが違えば失敗する", () => {
    expect(() => checkKuromojiDictUrl(loader, CDN, "0.1.3")).toThrow(/0\.1\.3/);
  });

  it("Worker が別の URL から辞書を取得していれば失敗する", () => {
    const other = "https://cdn.jsdelivr.net/npm/kuromoji@0.2.0/dict";
    expect(() => checkKuromojiDictUrl(loader, other, "0.1.2")).toThrow(/辞書を取得していません/);
  });

  it("ローダーに KUROMOJI_CDN がなければ失敗する", () => {
    expect(() => checkKuromojiDictUrl("", CDN, "0.1.2")).toThrow(/KUROMOJI_CDN/);
  });
});

describe("virtualModuleCode", () => {
  it("ローダーのパスを export する", () => {
    expect(virtualModuleCode("textlint/loader-abc.js")).toBe(
      'export const textlintLoaderPath = "textlint/loader-abc.js";\n',
    );
  });
});

describe("isWorkerStale", () => {
  let root: string;
  const worker = () => join(root, WORKER_SRC_DIR, WORKER_FILE);
  const touch = async (path: string, seconds: number) => {
    await writeFile(path, "x");
    await utimes(path, seconds, seconds);
  };

  beforeEach(async () => {
    root = await mkdtemp(join(tmpdir(), "textlint-worker-"));
    await mkdir(join(root, WORKER_SRC_DIR), { recursive: true });
    await mkdir(join(root, "prh"));
    await touch(join(root, ".textlintrc.json"), 1000);
    await touch(join(root, "package-lock.json"), 1000);
    await touch(join(root, "prh", "business-ja.yml"), 1000);
  });

  afterEach(async () => {
    await rm(root, { recursive: true, force: true });
  });

  it("Worker が無ければ true", async () => {
    expect(await isWorkerStale(root)).toBe(true);
  });

  it("Worker が入力より新しければ false", async () => {
    await touch(worker(), 2000);
    expect(await isWorkerStale(root)).toBe(false);
  });

  it(".textlintrc.json が Worker より新しければ true", async () => {
    await touch(worker(), 2000);
    await touch(join(root, ".textlintrc.json"), 3000);
    expect(await isWorkerStale(root)).toBe(true);
  });

  it("prh の辞書が Worker より新しければ true", async () => {
    await touch(worker(), 2000);
    await touch(join(root, "prh", "business-ja.yml"), 3000);
    expect(await isWorkerStale(root)).toBe(true);
  });

  it("入力が無くても失敗しない", async () => {
    await rm(join(root, "prh"), { recursive: true });
    await touch(worker(), 2000);
    expect(await isWorkerStale(root)).toBe(false);
  });
});
