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
      `<x:ClientData ObjectType="Checkbox"><x:Anchor>2,0,0,0,4,0,1,0</x:Anchor><x:Checked>1</x:Checked><x:FmlaLink>$B$1</x:FmlaLink></x:ClientData></v:shape>` +
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
