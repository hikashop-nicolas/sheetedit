import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { readWorkbook, writeWorkbook } from "../../core/workbook";
import { absoluteRange, absoluteRef, createXlsxControl, deleteXlsxControl, updateXlsxControlLinks } from "./control-create";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELNS = "http://schemas.openxmlformats.org/package/2006/relationships";
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

const anchor = (r1: number, r2: number): string =>
  `<controlPr defaultSize="0"><anchor moveWithCells="1">` +
  `<from><xdr:col xmlns:xdr="${XDR}">2</xdr:col><xdr:colOff xmlns:xdr="${XDR}">0</xdr:colOff><xdr:row xmlns:xdr="${XDR}">${r1}</xdr:row><xdr:rowOff xmlns:xdr="${XDR}">0</xdr:rowOff></from>` +
  `<to><xdr:col xmlns:xdr="${XDR}">4</xdr:col><xdr:colOff xmlns:xdr="${XDR}">0</xdr:colOff><xdr:row xmlns:xdr="${XDR}">${r2}</xdr:row><xdr:rowOff xmlns:xdr="${XDR}">0</xdr:rowOff></to>` +
  `</anchor></controlPr>`;

/** A workbook with a checkbox and a dropdown, both linked to cells. */
function book(opts: { withProps?: boolean; vmlOnly?: boolean } = {}): Uint8Array {
  const withProps = opts.withProps !== false && !opts.vmlOnly;
  const control = (sid: string, rid: string, name: string, r1: number, r2: number): string =>
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="x14">` +
    `<control shapeId="${sid}" r:id="${rid}" name="${name}">${anchor(r1, r2)}</control>` +
    `</mc:Choice></mc:AlternateContent>`;
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(
      `<worksheet xmlns="${MAIN}" xmlns:r="${R}" xmlns:mc="${MC}" mc:Ignorable="x14">` +
      `<sheetData><row r="1"><c r="D1" t="inlineStr"><is><t>Alpha</t></is></c></row>` +
      `<row r="2"><c r="D2" t="inlineStr"><is><t>Beta</t></is></c></row>` +
      `<row r="3"><c r="D3" t="inlineStr"><is><t>Gamma</t></is></c></row></sheetData>` +
      `<legacyDrawing r:id="rIdV"/><controls>${control("1025", "rId2", "Check Box 1", 0, 1)}${control("1026", "rId3", "Drop Down 2", 2, 3)}</controls></worksheet>`),
    "xl/worksheets/_rels/sheet1.xml.rels": strToU8(
      `<Relationships xmlns="${RELNS}"><Relationship Id="rIdV" Type="${R}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>` +
      (withProps ? `<Relationship Id="rId2" Type="${R}/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/><Relationship Id="rId3" Type="${R}/ctrlProp" Target="../ctrlProps/ctrlProp2.xml"/>` : "") +
      `</Relationships>`),
    "xl/drawings/vmlDrawing1.vml": strToU8(
      `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
      `<v:shape id="_x0000_s1025"><v:textbox><div>Enabled</div></v:textbox>` +
      `<x:ClientData ObjectType="Checkbox"><x:Anchor>2,0,0,0,4,0,1,0</x:Anchor><x:Checked>1</x:Checked><x:FmlaLink>$B$1</x:FmlaLink><x:FmlaMacro>[0]!Toggle_Click</x:FmlaMacro></x:ClientData></v:shape>` +
      `<v:shape id="_x0000_s1026">` +
      `<x:ClientData ObjectType="Drop"><x:Anchor>2,0,2,0,4,0,3,0</x:Anchor><x:Sel>2</x:Sel><x:FmlaLink>$B$2</x:FmlaLink><x:FmlaRange>$D$1:$D$3</x:FmlaRange></x:ClientData></v:shape></xml>`),
  };
  if (withProps) {
    files["xl/ctrlProps/ctrlProp1.xml"] = strToU8(`<formControlPr xmlns="${X14}" objectType="CheckBox" checked="Checked" fmlaLink="$B$1" lockText="1"/>`);
    files["xl/ctrlProps/ctrlProp2.xml"] = strToU8(`<formControlPr xmlns="${X14}" objectType="Drop" dropLines="3" fmlaLink="$B$2" fmlaRange="$D$1:$D$3" sel="2"/>`);
  }
  return zipSync(files);
}

const part = (b: Uint8Array, p: string): string => strFromU8(unzipSync(b)[p]!);

describe("xlsx form controls", () => {
  it("reads the kind, state and linked cell from ctrlProps", () => {
    const [check, drop] = readWorkbook(book()).sheets[0]!.controls!;
    expect(check!.kind).toBe("checkbox");
    expect(check!.checked).toBe(true);
    expect(check!.linkedCell).toBe("$B$1");
    expect(drop!.kind).toBe("dropdown");
    expect(drop!.selected).toBe(2);
    expect(drop!.sourceRange).toBe("$D$1:$D$3");
  });

  it("takes the label from the VML, which is the only place it lives", () => {
    expect(readWorkbook(book()).sheets[0]!.controls![0]!.label).toBe("Enabled");
  });

  it("positions from the worksheet anchor", () => {
    const a = readWorkbook(book()).sheets[0]!.controls![0]!.anchor!;
    expect(a.fromCol).toBe(3); // xdr col 2 is 0-based
    expect(a.fromRow).toBe(1);
    expect(a.toRow).toBe(2);
  });

  it("falls back to the VML for a file with no ctrlProps at all", () => {
    // Pre-2007 files carry the whole state in <x:ClientData>.
    const [check, drop] = readWorkbook(book({ vmlOnly: true })).sheets[0]!.controls!;
    expect(check!.kind).toBe("checkbox");
    expect(check!.checked).toBe(true);
    expect(check!.linkedCell).toBe("$B$1");
    expect(check!.anchor?.fromRow).toBe(1);
    expect(drop!.kind).toBe("dropdown");
    expect(drop!.selected).toBe(2);
  });

  it("a sheet with no controls has none", () => {
    const plain = zipSync({
      "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`),
      "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
      "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
      "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}"><sheetData/></worksheet>`),
    });
    expect(readWorkbook(plain).sheets[0]!.controls).toBeUndefined();
  });

  it("writes a toggled checkbox into both ctrlProps and the VML", () => {
    const wb = readWorkbook(book());
    const check = wb.sheets[0]!.controls![0]!;
    check.checked = false;
    check.dirty = true;
    const out = writeWorkbook(wb);
    expect(part(out, "xl/ctrlProps/ctrlProp1.xml")).toMatch(/checked="Unchecked"/);
    // The VML mirror matters: an older reader looks only there.
    expect(part(out, "xl/drawings/vmlDrawing1.vml")).toMatch(/<x:Checked>0<\/x:Checked>/);
  });

  it("writes a new dropdown selection to both places", () => {
    const wb = readWorkbook(book());
    const drop = wb.sheets[0]!.controls![1]!;
    drop.selected = 3;
    drop.dirty = true;
    const out = writeWorkbook(wb);
    expect(part(out, "xl/ctrlProps/ctrlProp2.xml")).toMatch(/sel="3"/);
    expect(part(out, "xl/drawings/vmlDrawing1.vml")).toMatch(/<x:Sel>3<\/x:Sel>/);
  });

  it("keeps the attributes it does not model", () => {
    const wb = readWorkbook(book());
    const check = wb.sheets[0]!.controls![0]!;
    check.checked = false;
    check.dirty = true;
    const xml = part(writeWorkbook(wb), "xl/ctrlProps/ctrlProp1.xml");
    expect(xml).toMatch(/lockText="1"/);
    expect(xml).toMatch(/fmlaLink="\$B\$1"/);
  });

  it("leaves the parts untouched when nothing was toggled", () => {
    const src = book();
    const out = writeWorkbook(readWorkbook(src));
    expect(part(out, "xl/ctrlProps/ctrlProp1.xml")).toBe(part(src, "xl/ctrlProps/ctrlProp1.xml"));
    expect(part(out, "xl/drawings/vmlDrawing1.vml")).toBe(part(src, "xl/drawings/vmlDrawing1.vml"));
  });

  it("round-trips a toggled state", () => {
    const wb = readWorkbook(book());
    Object.assign(wb.sheets[0]!.controls![0]!, { checked: false, dirty: true });
    Object.assign(wb.sheets[0]!.controls![1]!, { selected: 1, dirty: true });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!.controls!;
    expect(back[0]!.checked).toBe(false);
    expect(back[1]!.selected).toBe(1);
  });

  it("clears the dirty flag so a second save does not rewrite the parts", () => {
    const wb = readWorkbook(book());
    const check = wb.sheets[0]!.controls![0]!;
    check.checked = false;
    check.dirty = true;
    writeWorkbook(wb);
    expect(check.dirty).toBe(false);
  });
});

/** A workbook with no controls at all, to create them from scratch in. */
function bare(): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`),
    "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<worksheet xmlns="${MAIN}"><sheetData><row r="1"><c r="D1" t="inlineStr"><is><t>Alpha</t></is></c></row></sheetData></worksheet>`),
  });
}

describe("reference helpers", () => {
  it("normalises a reference to absolute form", () => {
    expect(absoluteRef("b2")).toBe("$B$2");
    expect(absoluteRef("$B$2")).toBe("$B$2");
    expect(absoluteRef("nonsense")).toBeUndefined();
  });
  it("normalises a range, and rejects a broken one", () => {
    expect(absoluteRange("d1:d3")).toBe("$D$1:$D$3");
    expect(absoluteRange("D1")).toBe("$D$1");
    expect(absoluteRange("D1:")).toBeUndefined();
  });
});

describe("creating form controls", () => {
  it("builds every part a control needs, so a reader can find it", () => {
    const wb = readWorkbook(bare());
    const sheet = wb.sheets[0]!;
    const ctl = createXlsxControl(wb, sheet, { kind: "checkbox", label: "Go", linkedCell: "$B$1", at: { r1: 1, c1: 3, r2: 2, c2: 5 } })!;
    expect(ctl).toBeTruthy();
    const out = unzipSync(writeWorkbook(wb));
    const dec = (p: string) => strFromU8(out[p]!);
    // The props part, its content type, its relationship, the VML shape, and the worksheet entry.
    expect(out[ctl.propsPath!]).toBeTruthy();
    expect(dec("[Content_Types].xml")).toContain("controlproperties");
    expect(dec("[Content_Types].xml")).toMatch(/Extension="vml"/);
    expect(dec("xl/worksheets/_rels/sheet1.xml.rels")).toContain("ctrlProp");
    expect(dec("xl/worksheets/_rels/sheet1.xml.rels")).toContain("vmlDrawing");
    expect(dec(ctl.vmlPath!)).toMatch(/ObjectType="Checkbox"/);
    const ws = dec("xl/worksheets/sheet1.xml");
    expect(ws).toContain("<controls>");
    expect(ws).toContain("legacyDrawing");
  });

  it("round-trips a created control through the reader", () => {
    const wb = readWorkbook(bare());
    createXlsxControl(wb, wb.sheets[0]!, { kind: "checkbox", label: "Go", linkedCell: "$B$1", at: { r1: 1, c1: 3, r2: 2, c2: 5 } });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!.controls!;
    expect(back.length).toBe(1);
    expect(back[0]!.kind).toBe("checkbox");
    expect(back[0]!.label).toBe("Go");
    expect(back[0]!.linkedCell).toBe("$B$1");
    expect(back[0]!.checked).toBe(false);
    expect(back[0]!.anchor?.fromCol).toBe(3);
  });

  it("a created dropdown carries its source range", () => {
    const wb = readWorkbook(bare());
    createXlsxControl(wb, wb.sheets[0]!, { kind: "dropdown", linkedCell: "$B$2", sourceRange: "$D$1:$D$3", at: { r1: 3, c1: 3, r2: 4, c2: 5 } });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!.controls![0]!;
    expect(back.kind).toBe("dropdown");
    expect(back.sourceRange).toBe("$D$1:$D$3");
    expect(back.selected).toBe(0);
  });

  it("a second control reuses the sheet's one VML drawing", () => {
    const wb = readWorkbook(bare());
    const sheet = wb.sheets[0]!;
    const a = createXlsxControl(wb, sheet, { kind: "checkbox", at: { r1: 1, c1: 1, r2: 2, c2: 3 } })!;
    const b = createXlsxControl(wb, sheet, { kind: "spin", at: { r1: 3, c1: 1, r2: 4, c2: 3 } })!;
    expect(b.vmlPath).toBe(a.vmlPath);
    expect(b.shapeId).not.toBe(a.shapeId); // distinct shapes, or they would collide
    expect(readWorkbook(writeWorkbook(wb)).sheets[0]!.controls!.length).toBe(2);
  });

  it("adds a control to a workbook that already had one", () => {
    const wb = readWorkbook(book());
    const sheet = wb.sheets[0]!;
    createXlsxControl(wb, sheet, { kind: "spin", linkedCell: "$B$5", at: { r1: 6, c1: 3, r2: 7, c2: 5 } });
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!.controls!;
    expect(back.length).toBe(3);
    expect(back.some((c) => c.kind === "spin" && c.linkedCell === "$B$5")).toBe(true);
  });

  it("deletes a control and everything belonging to it", () => {
    const wb = readWorkbook(book());
    const sheet = wb.sheets[0]!;
    const victim = sheet.controls![0]!;
    const propsPath = victim.propsPath!;
    deleteXlsxControl(wb, sheet, victim);
    expect(wb.files[propsPath]).toBeUndefined();
    const out = writeWorkbook(wb);
    const back = readWorkbook(out).sheets[0]!.controls!;
    expect(back.length).toBe(1);
    expect(back[0]!.name).toBe("Drop Down 2");
    // The shape must go too, or the VML keeps drawing a control nothing points at.
    expect(strFromU8(unzipSync(out)[victim.vmlPath!]!)).not.toContain("_x0000_s1025");
  });

  it("rewrites a linked cell in both parts", () => {
    const wb = readWorkbook(book());
    const ctl = wb.sheets[0]!.controls![0]!;
    ctl.linkedCell = "$Z$9";
    updateXlsxControlLinks(wb, ctl);
    const out = writeWorkbook(wb);
    expect(strFromU8(unzipSync(out)["xl/ctrlProps/ctrlProp1.xml"]!)).toMatch(/fmlaLink="\$Z\$9"/);
    expect(strFromU8(unzipSync(out)[ctl.vmlPath!]!)).toMatch(/<x:FmlaLink>\$Z\$9<\/x:FmlaLink>/);
    expect(readWorkbook(out).sheets[0]!.controls![0]!.linkedCell).toBe("$Z$9");
  });
});

describe("the macro a control runs", () => {
  it("reads FmlaMacro and strips the workbook prefix Excel writes", () => {
    const [check, drop] = readWorkbook(book()).sheets[0]!.controls!;
    expect(check!.macro).toBe("Toggle_Click");
    expect(drop!.macro).toBeUndefined();
  });

  it("reads it from a file that has no ctrlProps either", () => {
    expect(readWorkbook(book({ vmlOnly: true })).sheets[0]!.controls![0]!.macro).toBe("Toggle_Click");
  });

  it("writes an assignment back into the VML, qualified the way Excel writes it", () => {
    const wb = readWorkbook(book());
    const ctl = wb.sheets[0]!.controls![1]!; // the dropdown, which had none
    ctl.macro = "Refresh";
    updateXlsxControlLinks(wb, ctl);
    const vml = part(writeWorkbook(wb), "xl/drawings/vmlDrawing1.vml");
    expect(vml).toContain("<x:FmlaMacro>[0]!Refresh</x:FmlaMacro>");
  });

  it("removes the assignment when it is cleared", () => {
    const wb = readWorkbook(book());
    const ctl = wb.sheets[0]!.controls![0]!;
    ctl.macro = undefined;
    updateXlsxControlLinks(wb, ctl);
    expect(part(writeWorkbook(wb), "xl/drawings/vmlDrawing1.vml")).not.toContain("FmlaMacro");
  });

  it("survives a round trip through the reader", () => {
    const wb = readWorkbook(book());
    const ctl = wb.sheets[0]!.controls![1]!;
    ctl.macro = "Refresh";
    updateXlsxControlLinks(wb, ctl);
    expect(readWorkbook(writeWorkbook(wb)).sheets[0]!.controls![1]!.macro).toBe("Refresh");
  });

  it("reads the assignment out of the real macro fixture", async () => {
    const { readFileSync } = await import("node:fs");
    const wb = readWorkbook(new Uint8Array(readFileSync("src/fixtures/macros-cp950.xlsm")));
    // That file has no <controls> element at all: the button exists only in the VML, which is how
    // older Excel wrote them, and the whole control would otherwise be invisible.
    const [button] = wb.sheets.flatMap((s) => s.controls ?? []);
    expect(button?.kind).toBe("button");
    expect(button?.macro).toBe("Button1_Click");
    expect(button?.anchor).toBeTruthy();
  });

  it("does not mistake a cell comment's VML shape for a control", () => {
    const files = unzipSync(book());
    files["xl/drawings/vmlDrawing1.vml"] = strToU8(
      `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
      `<v:shape id="_x0000_s2000"><x:ClientData ObjectType="Note"><x:Anchor>1,0,1,0,2,0,2,0</x:Anchor></x:ClientData></v:shape></xml>`);
    // Both worksheet <control> entries lose their VML, and the only shape left is a comment.
    expect(readWorkbook(zipSync(files)).sheets[0]!.controls!.every((c) => c.kind !== "label" || c.name.startsWith("Check") || c.name.startsWith("Drop"))).toBe(true);
    expect(readWorkbook(zipSync(files)).sheets[0]!.controls!.length).toBe(2);
  });
});

describe("moving and resizing a control", () => {
  it("writes a new anchor into the worksheet and the VML", () => {
    const wb = readWorkbook(book());
    const ctl = wb.sheets[0]!.controls![0]!;
    ctl.anchor = { fromCol: 5, fromRow: 7, fromColOff: 0, fromRowOff: 0, toCol: 7, toRow: 9, toColOff: 0, toRowOff: 0 };
    updateXlsxControlLinks(wb, ctl, wb.sheets[0]!);
    const out = writeWorkbook(wb);
    // The VML anchor is 0-based and lists from-col, from-row, to-col, to-row with offsets between.
    expect(part(out, "xl/drawings/vmlDrawing1.vml")).toContain("<x:Anchor>4,0,6,0,6,0,8,0</x:Anchor>");
    // The worksheet's own <anchor> has to agree, or the two placements contradict each other.
    const ws = part(out, "xl/worksheets/sheet1.xml");
    expect(ws).toMatch(/<from>[\s\S]*?<xdr:col[^>]*>4<\/xdr:col>[\s\S]*?<xdr:row[^>]*>6<\/xdr:row>[\s\S]*?<\/from>/);
  });

  it("round-trips the new placement through the reader", () => {
    const wb = readWorkbook(book());
    const ctl = wb.sheets[0]!.controls![0]!;
    ctl.anchor = { fromCol: 3, fromRow: 4, fromColOff: 0, fromRowOff: 0, toCol: 5, toRow: 6, toColOff: 0, toRowOff: 0 };
    updateXlsxControlLinks(wb, ctl, wb.sheets[0]!);
    const back = readWorkbook(writeWorkbook(wb)).sheets[0]!.controls![0]!;
    expect([back.anchor?.fromCol, back.anchor?.fromRow, back.anchor?.toCol, back.anchor?.toRow]).toEqual([3, 4, 5, 6]);
  });

  it("still writes the VML anchor for a control the worksheet never listed", () => {
    // The older form: a VML shape with no <control> entry, which is what the real fixture has.
    const files = unzipSync(book({ vmlOnly: true }));
    files["xl/worksheets/sheet1.xml"] = strToU8(strFromU8(files["xl/worksheets/sheet1.xml"]!).replace(/<controls>[\s\S]*<\/controls>/, ""));
    const wb = readWorkbook(zipSync(files));
    const ctl = wb.sheets[0]!.controls![0]!;
    ctl.anchor = { fromCol: 2, fromRow: 3, fromColOff: 0, fromRowOff: 0, toCol: 4, toRow: 5, toColOff: 0, toRowOff: 0 };
    updateXlsxControlLinks(wb, ctl, wb.sheets[0]!);
    expect(part(writeWorkbook(wb), "xl/drawings/vmlDrawing1.vml")).toContain("<x:Anchor>1,0,2,0,3,0,4,0</x:Anchor>");
  });
});

describe("ActiveX controls", () => {
  const AX = "http://schemas.microsoft.com/office/2006/activeX";
  const AXREL = "http://schemas.microsoft.com/office/2006/relationships";
  const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff];
  const u32 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];
  const padded = (t: string): number[] => {
    const b = [...t].map((c) => c.charCodeAt(0));
    return [...b, ...Array((4 - (b.length % 4)) % 4).fill(0)];
  };
  const BUTTON_CLSID = [0x40, 0x32, 0x05, 0xd7, 0x69, 0xce, 0xcd, 0x11, 0xa7, 0x77, 0x00, 0xdd, 0x01, 0x14, 0x3c, 0x57];
  const CHECK_CLSID = [0x40, 0x1d, 0xd2, 0x8b, 0x42, 0xec, 0xce, 0x11, 0x9e, 0x0d, 0x00, 0xaa, 0x00, 0x60, 0x02, 0xf3];

  const buttonBin = (caption: string): Uint8Array => {
    const data = u32(0x80000000 | caption.length);
    const extra = [...padded(caption), ...u32(3731), ...u32(979)];
    return new Uint8Array([...BUTTON_CLSID, 0, 2, ...u16(4 + data.length + extra.length), ...u32(0x28), ...data, ...extra]);
  };
  const checkBin = (value: string, caption: string): Uint8Array => {
    const data = [0x04, 0, 0, 0, ...u32(0x80000000 | value.length), ...u32(0x80000000 | caption.length)];
    const extra = [...u32(2831), ...u32(767), ...padded(value), ...padded(caption)];
    return new Uint8Array([...CHECK_CLSID, 0, 2, ...u16(8 + data.length + extra.length), ...u32(0x80c00140 & ~1), ...u32(0), ...data, ...extra]);
  };

  /** A workbook whose controls are ActiveX, written the way Excel writes them. */
  function axBook(): Uint8Array {
    // Excel writes each control TWICE, under mc:Choice with its placement and mc:Fallback without.
    const control = (sid: string, rid: string, name: string, r1: number): string =>
      `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="x14">` +
      `<control shapeId="${sid}" r:id="${rid}" name="${name}">${anchor(r1, r1 + 1)}</control>` +
      `</mc:Choice><mc:Fallback><control shapeId="${sid}" r:id="${rid}" name="${name}"/></mc:Fallback></mc:AlternateContent>`;
    const ocx = (clsid: string): Uint8Array => strToU8(
      `<ax:ocx xmlns:ax="${AX}" xmlns:r="${R}" ax:classid="${clsid}" ax:persistence="persistStreamInit" r:id="rId1"/>`);
    const binRels = strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${AXREL}/activeXControlBinary" Target="activeX1.bin"/></Relationships>`);
    return zipSync({
      "[Content_Types].xml": strToU8(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/></Types>`),
      "_rels/.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      "xl/workbook.xml": strToU8(`<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
      "xl/_rels/workbook.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
      "xl/worksheets/sheet1.xml": strToU8(
        `<worksheet xmlns="${MAIN}" xmlns:r="${R}" xmlns:mc="${MC}" mc:Ignorable="x14"><sheetData/>` +
        `<controls>${control("1", "rIdA", "CommandButton1", 0)}${control("2", "rIdB", "CheckBox1", 4)}</controls></worksheet>`),
      "xl/worksheets/_rels/sheet1.xml.rels": strToU8(
        `<Relationships xmlns="${RELNS}">` +
        `<Relationship Id="rIdA" Type="${AXREL}/control" Target="../activeX/activeX1.xml"/>` +
        `<Relationship Id="rIdB" Type="${AXREL}/control" Target="../activeX/activeX2.xml"/></Relationships>`),
      "xl/activeX/activeX1.xml": ocx("{D7053240-CE69-11CD-A777-00DD01143C57}"),
      "xl/activeX/_rels/activeX1.xml.rels": binRels,
      "xl/activeX/activeX1.bin": buttonBin("CommandButton1"),
      "xl/activeX/activeX2.xml": ocx("{8BD21D40-EC42-11CE-9E0D-00AA006002F3}"),
      "xl/activeX/_rels/activeX2.xml.rels": strToU8(`<Relationships xmlns="${RELNS}"><Relationship Id="rId1" Type="${AXREL}/activeXControlBinary" Target="activeX2.bin"/></Relationships>`),
      "xl/activeX/activeX2.bin": checkBin("1", "CheckBox1"),
    });
  }

  it("reads each control once, not once per mc branch", () => {
    // Excel writes a Choice and a Fallback for every control; counting both drew each one twice.
    expect(readWorkbook(axBook()).sheets[0]!.controls).toHaveLength(2);
  });

  it("takes the kind from the class id, not from a formControlPr that is not there", () => {
    const [button, check] = readWorkbook(axBook()).sheets[0]!.controls!;
    expect(button!.kind).toBe("button");
    expect(check!.kind).toBe("checkbox");
    expect(button!.activeX).toBe(true);
  });

  it("reads the caption and the persisted value out of the binary", () => {
    const [button, check] = readWorkbook(axBook()).sheets[0]!.controls!;
    expect(button!.label).toBe("CommandButton1");
    expect(check!.label).toBe("CheckBox1");
    expect(check!.activeXValue).toBe("1");
    expect(check!.checked).toBe(true);
  });

  it("points a button at the handler its name implies", () => {
    // An ActiveX button runs <name>_Click from the sheet's own code module.
    expect(readWorkbook(axBook()).sheets[0]!.controls![0]!.macro).toBe("CommandButton1_Click");
  });

  it("places them from the worksheet anchor", () => {
    expect(readWorkbook(axBook()).sheets[0]!.controls![0]!.anchor?.fromRow).toBe(1);
  });
});
