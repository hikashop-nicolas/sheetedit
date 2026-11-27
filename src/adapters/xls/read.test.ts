import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { getCell } from "../../core/model";
import { readWorkbook } from "../../core/workbook";
import { md5, passwordHash16, Rc4, sha1, xorKeyAndArray } from "./crypto";
import { readXlsBook, xlsPasswordState } from "./read";
import { writeCfbStreams } from "./cfb";

const fixture = (): Uint8Array => new Uint8Array(readFileSync(join(__dirname, "../../fixtures/compound.xls")));
const hex = (b: Uint8Array): string => Array.from(b, (x) => x.toString(16).padStart(2, "0")).join("");

describe("crypto primitives", () => {
  it("md5 / sha1 / rc4 match their reference vectors", () => {
    const abc = new TextEncoder().encode("abc");
    expect(hex(md5(abc))).toBe("900150983cd24fb0d6963f7d28e17f72");
    expect(hex(sha1(abc))).toBe("a9993e364706816aba3e25717850c26c9cd0d89d");
    const data = new TextEncoder().encode("Plaintext");
    new Rc4(new TextEncoder().encode("Key")).apply(data);
    expect(hex(data)).toBe("bbf316e8d940af0ad3");
  });

  it("hashes protection passwords the way Excel does", () => {
    expect(passwordHash16("password").toString(16).toUpperCase()).toBe("83AF");
    expect(xorKeyAndArray("abc").verifier).toBe(passwordHash16("abc"));
  });
});

describe("xls read", () => {
  it("reads a LibreOffice-written .xls into the xlsx model", () => {
    const wb = readWorkbook(fixture());
    expect(wb.kind).toBe("xlsx");
    expect(wb.xls?.biff).toBe(8);
    const s = wb.sheets[0]!;
    expect(s.name).toBe("Data");
    expect(getCell(s, 1, 1)?.value).toBe("Region / item");
    expect(getCell(s, 6, 2)?.formula).toBe("SUM(B3:B5)");
    expect(getCell(s, 6, 2)?.value).toBe("1650");
    expect(s.rowOutline?.size).toBeGreaterThan(0);
  });

  it("says when no password is needed", () => {
    expect(xlsPasswordState(fixture())).toBe("none");
  });

  it("fails with a clear error on a stream that is not BIFF", () => {
    const junk = writeCfbStreams([{ name: "Workbook", data: new Uint8Array(600).fill(7) }]);
    expect(() => readXlsBook(junk)).toThrow(/damaged/);
  });
});
