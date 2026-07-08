import { describe, expect, it } from "vitest";
import {
  dateToSerial,
  durationToSerial,
  hasTimeFmt,
  isDateFmt,
  isTimeOnlyFmt,
  isoToSerial,
  parseDateInput,
  serialToDuration,
  serialToEditText,
  serialToIso,
  serialToParts,
} from "./dates";
import { formatNumber } from "./model";

describe("serial conversions", () => {
  it("uses the 1900 system including the fake leap day offset", () => {
    expect(dateToSerial(1900, 1, 1)).toBe(1);
    expect(dateToSerial(1900, 2, 28)).toBe(59);
    expect(dateToSerial(1900, 3, 1)).toBe(61); // serial 60 is the nonexistent Feb 29
    expect(dateToSerial(2020, 1, 1)).toBe(43831); // known Excel value
  });

  it("round-trips through serialToParts", () => {
    const s = dateToSerial(2026, 7, 8, 13, 30, 5);
    expect(serialToParts(s)).toEqual({ y: 2026, mo: 7, d: 8, hh: 13, mi: 30, ss: 5 });
    expect(serialToParts(dateToSerial(1900, 1, 1))).toMatchObject({ y: 1900, mo: 1, d: 1 });
  });

  it("agrees with SSF's rendering of the same serial", () => {
    const s = isoToSerial("2026-07-08")!;
    expect(formatNumber("yyyy-mm-dd", String(s))).toBe("2026-07-08");
  });

  it("serialToIso emits the time part only when present", () => {
    expect(serialToIso(isoToSerial("2026-07-08")!)).toBe("2026-07-08");
    expect(serialToIso(isoToSerial("2026-07-08T09:05:00")!)).toBe("2026-07-08T09:05:00");
  });

  it("serialToEditText uses a space-separated editable form", () => {
    expect(serialToEditText(isoToSerial("2026-07-08")!, false)).toBe("2026-07-08");
    expect(serialToEditText(isoToSerial("2026-07-08T09:05:00")!, true)).toBe("2026-07-08 09:05");
  });
});

describe("ODF durations", () => {
  it("parses PT#H#M#S to a day fraction and back", () => {
    expect(durationToSerial("PT13H30M0S")).toBeCloseTo(0.5625, 10);
    expect(serialToDuration(0.5625)).toBe("PT13H30M0S");
    expect(durationToSerial("not a duration")).toBeNull();
  });
});

describe("parseDateInput", () => {
  it("accepts ISO dates and datetimes", () => {
    expect(parseDateInput("2026-07-08", "mdy")).toMatchObject({ fmt: "yyyy-mm-dd" });
    expect(parseDateInput("2026/07/08 13:30", "mdy")?.fmt).toBe("yyyy-mm-dd hh:mm");
  });

  it("follows the locale day order for ambiguous d/m", () => {
    const fr = parseDateInput("8/7/2026", "dmy")!;
    expect(serialToIso(fr.serial)).toBe("2026-07-08");
    expect(fr.fmt).toBe("dd/mm/yyyy");
    const en = parseDateInput("8/7/2026", "mdy")!;
    expect(serialToIso(en.serial)).toBe("2026-08-07");
    expect(en.fmt).toBe("mm/dd/yyyy");
  });

  it("an out-of-range component disambiguates on its own", () => {
    expect(serialToIso(parseDateInput("25/12/2026", "mdy")!.serial)).toBe("2026-12-25");
    expect(serialToIso(parseDateInput("12/25/2026", "dmy")!.serial)).toBe("2026-12-25");
  });

  it("accepts bare times as day fractions", () => {
    const p = parseDateInput("13:30", "mdy")!;
    expect(p.timeOnly).toBe(true);
    expect(p.serial).toBeCloseTo(0.5625, 10);
    expect(p.fmt).toBe("hh:mm");
  });

  it("rejects invalid dates, invalid times, and plain text", () => {
    expect(parseDateInput("2026-02-30", "mdy")).toBeNull();
    expect(parseDateInput("13/13/2026", "mdy")).toBeNull();
    expect(parseDateInput("25:00", "mdy")).toBeNull();
    expect(parseDateInput("hello", "mdy")).toBeNull();
    expect(parseDateInput("3.5", "mdy")).toBeNull();
  });
});

describe("format classification", () => {
  it("detects date formats by code or built-in id", () => {
    expect(isDateFmt("yyyy-mm-dd")).toBe(true);
    expect(isDateFmt(14)).toBe(true);
    expect(isDateFmt("#,##0.00")).toBe(false);
    expect(isDateFmt('#,##0.00 "€"')).toBe(false);
    expect(isDateFmt('"dd" 0.00')).toBe(false); // quoted literals are not tokens
    expect(isDateFmt(undefined)).toBe(false);
  });

  it("distinguishes time-only from date formats", () => {
    expect(isTimeOnlyFmt("hh:mm:ss")).toBe(true);
    expect(isTimeOnlyFmt("yyyy-mm-dd hh:mm")).toBe(false);
    expect(hasTimeFmt("yyyy-mm-dd hh:mm")).toBe(true);
    expect(hasTimeFmt("dd/mm/yyyy")).toBe(false);
  });
});
