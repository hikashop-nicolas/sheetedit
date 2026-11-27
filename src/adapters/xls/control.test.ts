import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import type { SheetControl, Workbook } from "../../core/model";
import { readWorkbook } from "../../core/workbook";
import { capabilitiesOf } from "../../core/capabilities";
import { createXlsxControl } from "../xlsx/control-create";
import { readCfbStreams } from "./cfb";
import { records } from "./biff";
import { activeXObjBody, readCtlObj } from "./control";
import { writeXls, xlsSaveWarnings } from ".";

const hexBytes = (h: string): Uint8Array => Uint8Array.from(h.trim().split(/\s+/), (x) => parseInt(x, 16));

/** A .xls workbook with a list of fruit in E1:E3, read back as the editor would open it. */
function blankXls(): Workbook {
  const wb = readWorkbook(new Uint8Array(readFileSync(join(__dirname, "../../../cypress/fixtures/sample.xlsx"))));
  return readWorkbook(writeXls(wb));
}

const at = (r1: number, c1: number) => ({ r1, c1, r2: r1 + 1, c2: c1 + 2 });
const reopen = (wb: Workbook): Workbook => readWorkbook(writeXls(wb));
const byName = (wb: Workbook, name: string): SheetControl => wb.sheets[0]!.controls!.find((c) => c.label === name || c.name === name)!;

/** The OBJ records of the first sheet, as [ot, subrecord types]. */
function objRecords(bytes: Uint8Array): [number, number[]][] {
  const ws = readCfbStreams(bytes).streams.get("Workbook")!;
  const out: [number, number[]][] = [];
  for (const r of records(ws, 0)) {
    if (r.type !== 0x5d) continue;
    const o = readCtlObj(r.data, () => null);
    if (o) out.push([o.ot, o.subs.map((s) => s.ft)]);
  }
  return out;
}

describe("xls form controls", () => {
  it("are offered for a .xls", () => {
    expect(capabilitiesOf(blankXls()).formControls).toBe(true);
  });

  it("writes new controls of every kind and reads them back", () => {
    const wb = blankXls();
    const sheet = wb.sheets[0]!;
    const mk = (kind: SheetControl["kind"], r: number, extra: Partial<{ label: string; linkedCell: string; sourceRange: string }>) =>
      createXlsxControl(wb, sheet, { kind, at: at(r, 6), ...extra })!;
    mk("checkbox", 1, { label: "Paid", linkedCell: "$H$1" }).checked = true;
    mk("radio", 3, { label: "Small", linkedCell: "$H$2" });
    const drop = mk("dropdown", 5, { linkedCell: "$H$3", sourceRange: "$A$2:$A$3" });
    drop.selected = 2;
    mk("list", 7, { linkedCell: "$H$4", sourceRange: "$A$2:$A$3" });
    const spin = mk("spin", 10, { linkedCell: "$H$5" });
    Object.assign(spin, { value: 7, min: 1, max: 20, inc: 2 });
    mk("scroll", 12, { linkedCell: "$H$6" });
    mk("button", 14, { label: "Run" }).macro = "DoIt";
    mk("label", 16, { label: "Note" });
    mk("groupBox", 18, { label: "Size" });
    const bytes = writeXls(wb);
    expect(objRecords(bytes).map(([ot]) => ot)).toEqual([11, 12, 20, 18, 16, 17, 7, 14, 19]);
    const back = readWorkbook(bytes).sheets[0]!.controls!;
    expect(back.map((c) => c.kind)).toEqual(["checkbox", "radio", "dropdown", "list", "spin", "scroll", "button", "label", "groupBox"]);
    expect(back[0]).toMatchObject({ label: "Paid", linkedCell: "$H$1", checked: true });
    expect(back[1]).toMatchObject({ label: "Small", linkedCell: "$H$2", checked: false });
    expect(back[2]).toMatchObject({ linkedCell: "$H$3", sourceRange: "$A$2:$A$3", selected: 2 });
    expect(back[3]).toMatchObject({ linkedCell: "$H$4", sourceRange: "$A$2:$A$3", selected: 0 });
    expect(back[4]).toMatchObject({ linkedCell: "$H$5", value: 7, min: 1, max: 20, inc: 2 });
    expect(back[6]).toMatchObject({ label: "Run", macro: "DoIt" });
    expect(back[7]!.label).toBe("Note");
    expect(back[8]!.label).toBe("Size");
    expect(back[0]!.anchor).toMatchObject({ fromCol: 6, fromRow: 1, toCol: 8, toRow: 2 });
    expect(xlsSaveWarnings(readWorkbook(bytes)).map((w) => w.key)).not.toContain("xlsDropsControls");
  });

  it("writes back what the user changed and keeps the rest as read", () => {
    const wb = blankXls();
    const sheet = wb.sheets[0]!;
    createXlsxControl(wb, sheet, { kind: "checkbox", label: "Paid", linkedCell: "$H$1", at: at(1, 6) });
    createXlsxControl(wb, sheet, { kind: "dropdown", linkedCell: "$H$3", sourceRange: "$A$2:$A$3", at: at(5, 6) });
    createXlsxControl(wb, sheet, { kind: "spin", linkedCell: "$H$5", at: at(10, 6) });
    const read = reopen(wb);
    const [box, drop, spin] = read.sheets[0]!.controls!;
    expect(box!.xls).toBeDefined();
    box!.checked = true;
    box!.label = "Settled";
    box!.linkedCell = "$J$1";
    drop!.selected = 1;
    drop!.sourceRange = "$A$1:$A$4";
    spin!.value = 42;
    spin!.anchor = { ...spin!.anchor!, fromRow: 20, toRow: 21 };
    const again = reopen(read).sheets[0]!.controls!;
    expect(again[0]).toMatchObject({ label: "Settled", linkedCell: "$J$1", checked: true });
    expect(again[1]).toMatchObject({ selected: 1, sourceRange: "$A$1:$A$4" });
    expect(again[2]).toMatchObject({ value: 42, linkedCell: "$H$5" });
    expect(again[2]!.anchor).toMatchObject({ fromRow: 20, toRow: 21 });
    // A third save of untouched controls changes nothing.
    const third = reopen(reopen(read)).sheets[0]!.controls!;
    expect(third.map((c) => [c.kind, c.label, c.linkedCell, c.checked, c.selected, c.value])).toEqual(again.map((c) => [c.kind, c.label, c.linkedCell, c.checked, c.selected, c.value]));
  });

  it("drops a deleted control and unlinks a cleared link", () => {
    const wb = blankXls();
    const sheet = wb.sheets[0]!;
    createXlsxControl(wb, sheet, { kind: "checkbox", label: "A", linkedCell: "$H$1", at: at(1, 6) });
    createXlsxControl(wb, sheet, { kind: "checkbox", label: "B", linkedCell: "$H$2", at: at(3, 6) });
    const read = reopen(wb);
    read.sheets[0]!.controls = read.sheets[0]!.controls!.filter((c) => c.label !== "A");
    byName(read, "B").linkedCell = undefined;
    const back = reopen(read).sheets[0]!.controls!;
    expect(back).toHaveLength(1);
    expect(back[0]!.label).toBe("B");
    expect(back[0]!.linkedCell).toBeUndefined();
  });

  it("chains option buttons to their new ids", () => {
    const wb = blankXls();
    const sheet = wb.sheets[0]!;
    createXlsxControl(wb, sheet, { kind: "radio", label: "One", linkedCell: "$H$1", at: at(1, 6) });
    createXlsxControl(wb, sheet, { kind: "radio", label: "Two", at: at(3, 6) });
    const read = reopen(wb);
    const [one, two] = read.sheets[0]!.controls!;
    // Link the first to the second as Excel does: ftRboData's idRadNext.
    const rbo = one!.xls!.subs.find((s) => s.ft === 0x11)!;
    rbo.data = Uint8Array.of(two!.xls!.id & 255, two!.xls!.id >> 8, 1, 0);
    const bytes = writeXls(read);
    const ws = readCfbStreams(bytes).streams.get("Workbook")!;
    const objs = [...records(ws, 0)].filter((r) => r.type === 0x5d).map((r) => readCtlObj(r.data, () => null)!).filter((o) => o?.ot === 12);
    const next = objs[0]!.subs.find((s) => s.ft === 0x11)!.data;
    expect(next[0]! | (next[1]! << 8)).toBe(objs[1]!.id);
  });

  it("reads an ActiveX control's place in the Ctls stream and writes its new one", () => {
    // An OBJ for a Forms.CheckBox.1 control: ftCmo, ftCf, ftPioGrbit (fCtl | fPrstm), ftPictFmla.
    const obj = hexBytes(`15 00 12 00 08 00 01 00 15 20 00 00 00 00 3c 11 21 0a 00 00 00 00
      07 00 02 00 02 00 08 00 02 00 31 00
      09 00 32 00 20 00 05 00 34 22 2a 01 02 c0 e3 a2 00 03 10 00 00 46 6f 72 6d 73 2e 43 68 65 63 6b 42 6f 78 2e 31 00
      00 00 00 00 68 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00`);
    const o = readCtlObj(obj, () => null)!;
    expect(o.activeX).toEqual({ progId: "Forms.CheckBox.1", pos: 0, size: 0x68 });
    const ctl: SheetControl = { kind: "checkbox", name: "CheckBox1", activeX: true, xls: { raw: new Uint8Array(0), anchor: { fromCol: 1, fromRow: 1, fromColOff: 0, fromRowOff: 0, toCol: 2, toRow: 2, toColOff: 0, toRowOff: 0 }, ot: 8, id: 1, subs: o.subs, read: {}, progId: "Forms.CheckBox.1" } };
    const out = activeXObjBody(ctl, 5, 200, 0x68, { ref: () => null, macro: () => null, newId: () => undefined })!.out();
    const again = readCtlObj(out, () => null)!;
    expect(again.id).toBe(5);
    expect(again.activeX).toEqual({ progId: "Forms.CheckBox.1", pos: 200, size: 0x68 });
  });
});
