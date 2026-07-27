import { describe, expect, it } from "vitest";
import { isParentControlBin, readParentControl } from "./activex-form";

// A parent control persists as a STORAGE, so these fixtures are real compound files: the bytes are
// built here from [MS-OFORMS] section 2.1.2 (the Form stream, its site data, and the "o" stream
// holding each child's own control structure) rather than taken from someone's workbook.
//
// Every stream is padded past the 4096-byte mini-stream cutoff so each lives in ordinary sectors,
// which keeps the builder to a FAT and a directory. The padding is trailing bytes the format
// already ignores: DesignExData after the site data, and anything past a child's stated size.

const SECTOR = 512;
const FREE = 0xffffffff, ENDOFCHAIN = 0xfffffffe, FATSECT = 0xfffffffd;

/** A compound file holding the named streams at the root, plus optional child storages. */
function cfb(streams: Record<string, Uint8Array>, storages: Record<string, Record<string, Uint8Array>> = {}): Uint8Array {
  // Flatten to directory entries: root, then each root stream, then each storage and its streams.
  type Dir = { name: string; type: number; data?: Uint8Array; child: number; left: number; right: number; start: number; size: number };
  const dirs: Dir[] = [{ name: "Root Entry", type: 5, child: -1, left: -1, right: -1, start: ENDOFCHAIN, size: 0 }];
  const rootKids: number[] = [];
  for (const [name, data] of Object.entries(streams)) {
    dirs.push({ name, type: 2, data, child: -1, left: -1, right: -1, start: 0, size: data.length });
    rootKids.push(dirs.length - 1);
  }
  for (const [sname, inner] of Object.entries(storages)) {
    const kids: number[] = [];
    for (const [name, data] of Object.entries(inner)) {
      dirs.push({ name, type: 2, data, child: -1, left: -1, right: -1, start: 0, size: data.length });
      kids.push(dirs.length - 1);
    }
    dirs.push({ name: sname, type: 1, child: kids[0] ?? -1, left: -1, right: -1, start: ENDOFCHAIN, size: 0 });
    // Siblings hang off each other's right, which is all the reader's walk needs.
    for (let i = 0; i + 1 < kids.length; i++) dirs[kids[i]!]!.right = kids[i + 1]!;
    rootKids.push(dirs.length - 1);
  }
  dirs[0]!.child = rootKids[0] ?? -1;
  for (let i = 0; i + 1 < rootKids.length; i++) dirs[rootKids[i]!]!.right = rootKids[i + 1]!;

  // Lay the streams out in sectors after the FAT and the directory.
  const dirSectors = Math.ceil((dirs.length * 128) / SECTOR);
  const withData = dirs.filter((d) => d.data);
  const dataSectors = withData.map((d) => Math.ceil(d.data!.length / SECTOR));
  const totalData = dataSectors.reduce((a, b) => a + b, 0);
  const fatEntries = 1 + dirSectors + totalData;              // the FAT sector itself is entry 0
  const fatSectors = Math.max(1, Math.ceil(fatEntries / (SECTOR / 4)));
  const total = fatSectors + dirSectors + totalData;

  const out = new Uint8Array(SECTOR * (1 + total));
  const dv = new DataView(out.buffer);
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
  dv.setUint16(24, 0x003e, true); // minor version
  dv.setUint16(26, 3, true);      // major version 3, which means 512-byte sectors
  dv.setUint16(28, 0xfffe, true); // byte order
  dv.setUint16(30, 9, true);      // sector shift
  dv.setUint16(32, 6, true);      // mini sector shift
  dv.setUint32(44, fatSectors, true);
  dv.setUint32(48, fatSectors, true);       // the directory starts after the FAT
  dv.setUint32(56, 4096, true);             // mini-stream cutoff
  dv.setUint32(60, ENDOFCHAIN, true);       // no mini FAT
  dv.setUint32(64, 0, true);
  dv.setUint32(68, ENDOFCHAIN, true);       // no DIFAT
  dv.setUint32(72, 0, true);
  for (let i = 0; i < 109; i++) dv.setUint32(76 + i * 4, i < fatSectors ? i : FREE, true);

  const fatAt = SECTOR;                     // sector 0 of the file body
  const fat = (n: number, v: number): void => dv.setUint32(fatAt + n * 4, v, true);
  for (let i = 0; i < fatSectors * (SECTOR / 4); i++) fat(i, FREE);
  for (let i = 0; i < fatSectors; i++) fat(i, FATSECT);
  for (let i = 0; i < dirSectors; i++) fat(fatSectors + i, i + 1 === dirSectors ? ENDOFCHAIN : fatSectors + i + 1);

  let sector = fatSectors + dirSectors;
  withData.forEach((d, i) => {
    d.start = sector;
    const n = dataSectors[i]!;
    for (let k = 0; k < n; k++) fat(sector + k, k + 1 === n ? ENDOFCHAIN : sector + k + 1);
    out.set(d.data!, SECTOR * (1 + sector));
    sector += n;
  });

  // The directory: 128 bytes per entry, the name in UTF-16 with its byte length.
  const dirAt = SECTOR * (1 + fatSectors);
  dirs.forEach((d, i) => {
    const at = dirAt + i * 128;
    for (let k = 0; k < d.name.length; k++) dv.setUint16(at + k * 2, d.name.charCodeAt(k), true);
    dv.setUint16(at + 64, (d.name.length + 1) * 2, true);
    out[at + 66] = d.type;
    out[at + 67] = 1; // colour: black
    dv.setInt32(at + 68, d.left, true);
    dv.setInt32(at + 72, d.right, true);
    dv.setInt32(at + 76, d.child, true);
    dv.setUint32(at + 116, d.type === 2 ? d.start : ENDOFCHAIN, true);
    dv.setUint32(at + 120, d.size, true);
  });
  return out;
}

const u16b = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff];
const u32b = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];
/** An fmString stored one byte per character, which is the compressed form. */
const strBytes = (s: string): number[] => [...s].map((c) => c.charCodeAt(0));
const pad4 = (b: number[]): number[] => [...b, ...Array((4 - (b.length % 4)) % 4).fill(0)];
const padTo = (b: number[], n: number): Uint8Array => new Uint8Array([...b, ...Array(Math.max(0, n - b.length)).fill(0)]);

interface SiteSpec { name: string; cache: number; size: number; left?: number; top?: number; tip?: string }

/** A Form stream: the caption and displayed size, then the site data. */
function formStream(caption: string, sites: SiteSpec[], cx = 8000, cy = 4000): Uint8Array {
  // fDisplayedSize (10) and fCaption (19).
  const mask = (1 << 10) | (1 << 19);
  const data = pad4(u32b(0x80000000 | caption.length));      // the caption's length and flag
  const extra = [...u32b(cx), ...u32b(cy), ...pad4(strBytes(caption))];
  const cbForm = 4 + data.length + extra.length;

  // Site data: an empty class table, then the depths/types array and the sites.
  const depths: number[] = [];
  for (const _ of sites) depths.push(0, 1);                  // depth 0, SITE_TYPE 1, no run
  const depthPad = Array((4 - (depths.length % 4)) % 4).fill(0);

  const siteRecords: number[][] = sites.map((s) => {
    // fName(0), fID(2), fObjectStreamSize(5), fTabIndex(6), fClsidCacheIndex(7), fPosition(8),
    // and fControlTipText(11) when the site has one.
    let m = (1 << 0) | (1 << 2) | (1 << 5) | (1 << 6) | (1 << 7) | (1 << 8);
    if (s.tip) m |= 1 << 11;
    const d: number[] = [
      ...u32b(0x80000000 | s.name.length),                   // NameData
      ...u32b(1),                                            // ID
      ...u32b(s.size),                                       // ObjectStreamSize
      ...u16b(0),                                            // TabIndex
      ...u16b(s.cache),                                      // ClsidCacheIndex
    ];
    if (s.tip) d.push(...u32b(0x80000000 | s.tip.length));
    const block = pad4(d);
    const ex = [...pad4(strBytes(s.name)), ...u32b(s.left ?? 0), ...u32b(s.top ?? 0), ...(s.tip ? pad4(strBytes(s.tip)) : [])];
    const cbSite = 4 + block.length + ex.length;
    return [...u16b(0), ...u16b(cbSite), ...u32b(m), ...block, ...ex];
  });
  const sitesFlat = siteRecords.flat();
  const countOfBytes = depths.length + depthPad.length + sitesFlat.length;

  return padTo([
    0x00, 0x04, ...u16b(cbForm), ...u32b(mask), ...data, ...extra,
    ...u16b(0),                                              // CountOfSiteClassInfo: none
    ...u32b(sites.length), ...u32b(countOfBytes),
    ...depths, ...depthPad, ...sitesFlat,
  ], 5000);
}

/** A MorphData child (a text box), as it sits in the "o" stream: no class id in front. */
function morphChild(value: string, displayStyle = 1): number[] {
  const lo = (1 << 6) | (1 << 22); // fDisplayStyle, fValue
  const block = pad4([displayStyle, ...u32b(0x80000000 | value.length).slice(0, 0)]);
  // DisplayStyle is 1 byte and Value is a 4-byte length word, in bit order.
  const data = pad4([displayStyle, 0, 0, 0, ...u32b(0x80000000 | value.length)]);
  void block;
  const extra = pad4(strBytes(value));
  const cb = 8 + data.length + extra.length;
  return [0x00, 0x02, ...u16b(cb), ...u32b(lo), ...u32b(0), ...data, ...extra];
}

describe("ActiveX parent controls", () => {
  it("knows a parent control from a leaf one", () => {
    const parent = cfb({ f: formStream("Options", []), o: padTo([], 5000) });
    expect(isParentControlBin(parent)).toBe(true);
    // A leaf control's .bin is a flat stream, not a compound file.
    expect(isParentControlBin(new Uint8Array([0x40, 0x32, 0x05, 0xd7, 0x00, 0x02]))).toBe(false);
  });

  it("reads a frame's caption and size", () => {
    const wb = readParentControl(cfb({ f: formStream("Delivery options", []), o: padTo([], 5000) }))!;
    expect(wb).toBeTruthy();
    expect(wb.caption).toBe("Delivery options");
    expect(wb.size).toEqual({ cx: 8000, cy: 4000 });
    expect(wb.sites).toEqual([]);
  });

  it("lists the controls it contains, with their names, kinds and positions", () => {
    const sites: SiteSpec[] = [
      { name: "OptionA", cache: 27, size: 0, left: 120, top: 60 },
      { name: "OptionB", cache: 27, size: 0, left: 120, top: 300, tip: "the other one" },
    ];
    const parent = readParentControl(cfb({ f: formStream("Choose", sites), o: padTo([], 5000) }))!;
    expect(parent.sites.map((s) => s.name)).toEqual(["OptionA", "OptionB"]);
    expect(parent.sites.map((s) => s.kind)).toEqual(["radio", "radio"]);
    expect(parent.sites[0]!.position).toEqual({ left: 120, top: 60 });
    expect(parent.sites[1]!.position).toEqual({ left: 120, top: 300 });
    expect(parent.sites[1]!.tooltip).toBe("the other one");
  });

  it("reads each child's own properties out of the object stream", () => {
    const child = morphChild("hello");
    const sites: SiteSpec[] = [{ name: "TextBox1", cache: 15, size: child.length }];
    const parent = readParentControl(cfb({ f: formStream("Form", sites), o: padTo(child, 5000) }))!;
    expect(parent.sites[0]!.control?.value).toBe("hello");
  });

  it("reads a MultiPage's pages, each a form of its own", () => {
    const pageSites: SiteSpec[] = [
      { name: "Page1", cache: 7, size: 0 },
      { name: "Page2", cache: 7, size: 0 },
    ];
    const bytes = cfb(
      { f: formStream("Wizard", pageSites), o: padTo([], 5000) },
      {
        Page1: { f: formStream("Details", [{ name: "Name", cache: 15, size: 0 }]), o: padTo([], 5000) },
        Page2: { f: formStream("Payment", []), o: padTo([], 5000) },
      },
    );
    const parent = readParentControl(bytes)!;
    expect(parent.kind).toBe("multiPage");
    // The order is the one the parent's sites give, not the directory's traversal order.
    expect(parent.pages?.map((p) => p.name)).toEqual(["Page1", "Page2"]);
    expect(parent.pages?.map((p) => p.caption)).toEqual(["Details", "Payment"]);
    expect(parent.pages?.[0]!.sites.map((s) => s.name)).toEqual(["Name"]);
  });

  it("refuses bytes that are not a parent control at all", () => {
    expect(readParentControl(new Uint8Array([1, 2, 3, 4]))).toBeUndefined();
    // A compound file whose Form stream is nonsense is refused rather than half-read.
    expect(readParentControl(cfb({ f: padTo([9, 9, 9, 9], 5000) }))).toBeUndefined();
  });
});
