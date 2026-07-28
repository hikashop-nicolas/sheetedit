import { isCfb, readCfb, writeCfb } from "vbalang";
import { readEmbeddedControl, setActiveXValue, type ActiveXControl, type ActiveXKind } from "./activex-read";

// ---------------------------------------------------------------------------
// ActiveX PARENT controls: Frame, MultiPage and TabStrip, per [MS-OFORMS] 2.1.2
// ---------------------------------------------------------------------------
// A leaf control persists as a single stream, which is what activex-read.ts walks. A parent
// control is a different thing: it persists as a STORAGE, so its .bin is a compound file rather
// than a flat stream, and the controls it contains live inside it.
//
//   "f"  the Form stream: the parent's own properties, then the ClassTable and the sites array,
//        which is the list of embedded controls with their names, types and positions.
//   "o"  the Object stream: every child's properties, concatenated in site order. Each child is
//        an ordinary control structure with NO class id in front, since its class comes from the
//        site's ClsidCacheIndex.
//   a child storage per page, for MultiPage: each Page is itself a form, with its own f and o.
//
// The container is read with vbalang's compound-file reader, which already does this for VBA
// projects; the structures below are the specification's own tables.

/** [MS-OFORMS] FormEmbeddedActiveXControlCached: what a child's ClsidCacheIndex means. */
const CACHED_KINDS: Record<number, ActiveXKind | "form" | "frame" | "multiPage" | "tabStrip"> = {
  7: "form",
  12: "image",
  14: "frame",
  15: "textbox", // MorphData: the family; DisplayStyle in its own stream tells the members apart
  16: "spin",
  17: "commandButton",
  18: "tabStrip",
  21: "label",
  23: "textbox",
  24: "list",
  25: "dropdown",
  26: "checkbox",
  27: "radio",
  28: "toggle",
  47: "scroll",
  57: "multiPage",
};

/** One embedded control, as the parent's site table describes it. */
export interface FormSite {
  name?: string;
  id?: number;
  /** What the child is: an index into the cached class list, or into the parent's ClassTable. */
  clsidCacheIndex?: number;
  kind?: ActiveXKind | "form" | "frame" | "multiPage" | "tabStrip";
  tabIndex?: number;
  /** Where the child sits inside the parent, in HIMETRIC. */
  position?: { left: number; top: number };
  /** How many bytes of the "o" stream belong to this child. */
  objectStreamSize?: number;
  tooltip?: string;
  controlSource?: string;
  rowSource?: string;
  /** How deep in the hierarchy, which is what nests a child inside an embedded parent. */
  depth: number;
  /** The child's own properties, read from the "o" stream with the leaf reader. */
  control?: ActiveXControl;
  /** Where this child's bytes sit in the "o" stream, so a write can splice rather than re-walk. */
  objectStreamAt?: number;
  /** Where this site's ObjectStreamSize sits in the "f" stream: a child whose stream changes
      length moves every child after it, and the size the parent records has to follow. */
  sizeFieldAt?: number;
}

/** A parent control: the container itself plus what it contains. */
export interface ParentControl {
  kind: "frame" | "multiPage" | "tabStrip" | "form";
  caption?: string;
  /** The form's own displayed size, in HIMETRIC. */
  size?: { cx: number; cy: number };
  sites: FormSite[];
  /** For a MultiPage, one entry per Page storage, each a form in its own right. */
  pages?: { name: string; caption?: string; sites: FormSite[] }[];
}

const u16 = (dv: DataView, at: number): number => dv.getUint16(at, true);
const u32 = (dv: DataView, at: number): number => dv.getUint32(at, true);

/** An fmString: `len` bytes, one per character when the compression flag is set, else UTF-16. */
function readString(b: Uint8Array, at: number, lengthAndFlag: number): { text: string; end: number } {
  const compressed = (lengthAndFlag & 0x80000000) !== 0;
  const len = lengthAndFlag & 0x7fffffff;
  if (len === 0 || at + len > b.length) return { text: "", end: at };
  let text = "";
  if (compressed) for (let i = 0; i < len; i++) text += String.fromCharCode(b[at + i]!);
  else for (let i = 0; i + 1 < len; i += 2) text += String.fromCharCode(b[at + i]! | (b[at + i + 1]! << 8));
  return { text, end: at + len };
}

/**
 * Skip the StreamData that sits between the Form's blocks and its site data: a MouseIcon, a font
 * and a Picture, each present only when its mask bit is. The two picture entries are a GUID and a
 * StdPicture (preamble 0x0000746C, a size, then the bytes); the font is a GUID and either a
 * StdFont or a TextProps, both of which state their own length.
 */
function skipFormStreamData(b: Uint8Array, at: number, mask: number): number {
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const bit = (n: number): boolean => (mask & (1 << n)) !== 0;
  let p = at;
  const skipPicture = (): void => {
    if (p + 24 > b.length || u32(dv, p + 16) !== 0x0000746c) return;
    p = p + 24 + u32(dv, p + 20);
  };
  const skipFont = (): void => {
    // GuidAndFont: a 16-byte GUID then the font. A TextProps states its size in its own header;
    // a StdFont is a 1-byte version, then fields, and the spec pins its total at 0x1A.
    if (p + 20 > b.length) return;
    const guid = Array.from(b.subarray(p, p + 16)).map((x) => x.toString(16).padStart(2, "0")).join("");
    p += 16;
    // TextProps guid {AFC20920-DA4E-11CE-B943-00AA006887B4}, little-endian in the first three parts.
    const isTextProps = guid.startsWith("2009c2af");
    if (isTextProps) p += 4 + u16(dv, p + 2); // version bytes then cb over the rest
    else p += 0x1a; // StdFont
  };
  if (bit(15)) skipPicture(); // fMouseIcon
  if (bit(20)) skipFont();    // fFont
  if (bit(21)) skipPicture(); // fPicture
  return p;
}

/** The Form stream's own properties, and where its site data begins. */
function readFormStream(b: Uint8Array): { caption?: string; size?: { cx: number; cy: number }; siteAt: number; mask: number } | undefined {
  if (b.length < 8 || b[0] !== 0x00 || b[1] !== 0x04) return undefined; // MajorVersion is 4 for a form
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const cbForm = u16(dv, 2);
  const mask = u32(dv, 4);
  const bit = (n: number): boolean => (mask & (1 << n)) !== 0;

  // Walk the DataBlock only as far as the caption's length word, which is all this needs from it;
  // the fields before it are fixed by the table and their alignment.
  let p = 8;
  const align = (n: number): void => { p += (n - ((p - 8) % n)) % n; };
  const take = (n: 1 | 2 | 4): void => { align(n); p += n; };
  if (bit(1)) take(4);  // BackColor
  if (bit(2)) take(4);  // ForeColor
  if (bit(3)) take(4);  // NextAvailableID
  if (bit(6)) take(4);  // BooleanProperties
  if (bit(7)) take(1);  // BorderStyle
  if (bit(8)) take(1);  // MousePointer
  if (bit(9)) take(1);  // ScrollBars
  if (bit(13)) take(4); // GroupCnt
  if (bit(15)) take(2); // MouseIcon placeholder
  if (bit(16)) take(1); // Cycle
  if (bit(17)) take(1); // SpecialEffect
  if (bit(18)) take(4); // BorderColor
  let captionLen = 0;
  if (bit(19)) { align(4); captionLen = u32(dv, p); p += 4; }

  // The ExtraDataBlock follows the DataBlock, 4-aligned: sizes first, then the caption string.
  const extraStart = 8 + cbForm - 4 >= 0 ? 4 + cbForm : p;
  let q = p + ((4 - ((p - 8) % 4)) % 4);
  let size: { cx: number; cy: number } | undefined;
  if (bit(10)) { size = { cx: dv.getInt32(q, true), cy: dv.getInt32(q + 4, true) }; q += 8; } // DisplayedSize
  if (bit(11)) q += 8; // LogicalSize
  if (bit(12)) q += 8; // ScrollPosition
  let caption: string | undefined;
  if (captionLen) { const got = readString(b, q, captionLen); caption = got.text; }

  // cbForm covers PropMask + DataBlock + ExtraDataBlock, so the stream data starts right after it
  // whatever this reader made of the blocks. That is the anchor; the walk above is only for the
  // caption, and being wrong about it costs the caption rather than the sites.
  const siteAt = skipFormStreamData(b, extraStart, mask);
  return { caption, size, siteAt, mask };
}

/** The sites array: what each embedded control is, and where. */
function readSiteData(b: Uint8Array, at: number, dontSaveClassTable: boolean): FormSite[] {
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let p = at;
  if (!dontSaveClassTable) {
    const classes = u16(dv, p);
    p += 2;
    // Each SiteClassInfo states its own length after a 2-byte version, so they can be stepped over
    // without modelling them: a child of such a class is one this reader does not render anyway.
    for (let i = 0; i < classes && p + 4 <= b.length; i++) p += 4 + u16(dv, p + 2);
  }
  if (p + 8 > b.length) return [];
  const countOfSites = u32(dv, p);
  const countOfBytes = u32(dv, p + 4);
  p += 8;
  const arrayStart = p;

  // SiteDepthsAndTypes: one entry per site, 2 bytes, or 3 when the entry counts a run.
  const depths: number[] = [];
  let seen = 0;
  while (seen < countOfSites && p + 2 <= b.length) {
    const depth = b[p]!;
    const typeOrCount = b[p + 1]!;
    if ((typeOrCount & 0x80) !== 0) {
      const run = typeOrCount & 0x7f;
      p += 3; // the run's shared type follows
      for (let i = 0; i < run && seen < countOfSites; i++, seen++) depths.push(depth);
    } else {
      p += 2;
      depths.push(depth);
      seen++;
    }
  }
  p += (4 - ((p - arrayStart) % 4)) % 4; // ArrayPadding: the sites start 4-aligned

  const sites: FormSite[] = [];
  const end = Math.min(b.length, arrayStart + countOfBytes);
  for (let i = 0; i < countOfSites && p + 4 <= end; i++) {
    const cbSite = u16(dv, p + 2);
    const maskAt = p + 4;
    const siteMask = u32(dv, maskAt);
    const sbit = (n: number): boolean => (siteMask & (1 << n)) !== 0;
    let d = maskAt + 4;
    const dalign = (n: number): void => { d += (n - ((d - maskAt) % n)) % n; };
    let nameLen = 0, tagLen = 0, tipLen = 0, licLen = 0, srcLen = 0, rowLen = 0;
    const site: FormSite = { depth: depths[i] ?? 0 };
    if (sbit(0)) { dalign(4); nameLen = u32(dv, d); d += 4; }
    if (sbit(1)) { dalign(4); tagLen = u32(dv, d); d += 4; }
    if (sbit(2)) { dalign(4); site.id = dv.getInt32(d, true); d += 4; }
    if (sbit(3)) { dalign(4); d += 4; }                                  // HelpContextID
    if (sbit(4)) { dalign(4); d += 4; }                                  // BitFlags
    if (sbit(5)) { dalign(4); site.sizeFieldAt = d; site.objectStreamSize = u32(dv, d); d += 4; }
    if (sbit(6)) { dalign(2); site.tabIndex = dv.getInt16(d, true); d += 2; }
    if (sbit(7)) { dalign(2); site.clsidCacheIndex = u16(dv, d); d += 2; }
    if (sbit(9)) { dalign(2); d += 2; }                                  // GroupID
    if (sbit(11)) { dalign(4); tipLen = u32(dv, d); d += 4; }
    if (sbit(12)) { dalign(4); licLen = u32(dv, d); d += 4; }
    if (sbit(13)) { dalign(4); srcLen = u32(dv, d); d += 4; }
    if (sbit(14)) { dalign(4); rowLen = u32(dv, d); d += 4; }

    // The ExtraDataBlock, in the order the specification lists it. Padding aligns against the
    // START OF THE BLOCK, not the file: computing it from the absolute offset reads every field
    // after the first string two bytes late whenever the record does not begin 4-aligned.
    const extraStart = maskAt + 4 + (((d - maskAt - 4) + 3) & ~3);
    let e = extraStart;
    const str = (len: number): string | undefined => {
      if (!len) return undefined;
      const got = readString(b, e, len);
      e = got.end + ((4 - ((got.end - extraStart) % 4)) % 4);
      return got.text;
    };
    site.name = str(nameLen);
    str(tagLen);
    if (sbit(8)) { site.position = { left: dv.getInt32(e, true), top: dv.getInt32(e + 4, true) }; e += 8; }
    site.tooltip = str(tipLen);
    str(licLen);
    site.controlSource = str(srcLen);
    site.rowSource = str(rowLen);
    if (site.clsidCacheIndex !== undefined) site.kind = CACHED_KINDS[site.clsidCacheIndex];
    sites.push(site);
    p = maskAt + cbSite;
  }
  return sites;
}

/** Whether these bytes are a parent control rather than a single control's stream. */
export const isParentControlBin = (bytes: Uint8Array): boolean => isCfb(bytes);

/**
 * Read a parent control's storage: what it is, what it contains, and each child's own properties.
 *
 * Returns undefined for bytes that are not a parent control, or whose Form stream does not parse -
 * the same rule the leaf reader follows, since a half-read container would place children wrongly.
 */
export function readParentControl(bytes: Uint8Array, kindHint?: ParentControl["kind"]): ParentControl | undefined {
  if (!isCfb(bytes)) return undefined;
  let cfb;
  try {
    cfb = readCfb(bytes);
  } catch {
    return undefined;
  }
  const f = cfb.readPath("/f");
  const o = cfb.readPath("/o");
  if (!f) return undefined;
  const form = readFormStream(f);
  if (!form) return undefined;

  // FORM_FLAG_DONTSAVECLASSTABLE is bit 15 of BooleanProperties, which this reader does not walk
  // to; a class table is the norm, and a file without one still parses because CountOfSiteClassInfo
  // then reads as the low half of CountOfSites, which the site walk validates against cbSite.
  const sites = readSiteData(f, form.siteAt, false);

  // The "o" stream holds the children end to end, each as long as its site says.
  if (o) {
    let at = 0;
    for (const site of sites) {
      const len = site.objectStreamSize ?? 0;
      site.objectStreamAt = at;
      if (len > 0 && at + len <= o.length) {
        const kind = site.kind;
        if (kind && kind !== "form" && kind !== "frame" && kind !== "multiPage" && kind !== "tabStrip") {
          site.control = readEmbeddedControl(o.subarray(at, at + len), kind);
        }
      }
      at += len;
    }
  }

  // A MultiPage keeps each Page in its own storage, itself a form with f and o streams.
  const pages: ParentControl["pages"] = [];
  for (const { path } of cfb.paths()) {
    const m = /^\/([^/]+)\/f$/.exec(path);
    if (!m) continue;
    const inner = cfb.readPath(path);
    const pageForm = inner ? readFormStream(inner) : undefined;
    if (!inner || !pageForm) continue;
    pages.push({ name: m[1]!, caption: pageForm.caption, sites: readSiteData(inner, pageForm.siteAt, false) });
  }
  // Storages come out in directory order, which is the tree's shape rather than the tabs' order.
  // The file states the order in the parent's own sites - one per page - so follow that, and fall
  // back to a natural sort of the storage names (Page2 before Page10) when it cannot be matched.
  if (pages.length > 1) {
    const rank = new Map(sites.map((s, i) => [s.name ?? "", i]));
    pages.sort((a, b) => {
      const ra = rank.get(a.name), rb = rank.get(b.name);
      if (ra !== undefined && rb !== undefined) return ra - rb;
      return a.name.localeCompare(b.name, undefined, { numeric: true });
    });
  }

  return {
    kind: kindHint ?? (pages.length ? "multiPage" : "frame"),
    caption: form.caption,
    size: form.size,
    sites,
    ...(pages.length ? { pages } : {}),
  };
}

/**
 * Change the persisted value of ONE control inside a container, and hand back the whole storage.
 *
 * The child's own bytes go through the same writer a standalone control uses, so it refuses the
 * same streams: a container whose child the reader would not vouch for is left alone. What this
 * adds is the container's own bookkeeping - the "o" stream is every child end to end, so a child
 * whose stream changes length moves every child after it, and the ObjectStreamSize the parent
 * records for it has to be corrected in the "f" stream to match.
 *
 * Returns undefined rather than a half-written storage on any doubt, and reads its own output back
 * before returning it, which is the rule the leaf writer already follows.
 */
export function setContainerChildValue(bytes: Uint8Array, siteIndex: number, value: string): Uint8Array | undefined {
  if (!isCfb(bytes)) return undefined;
  let cfb;
  try {
    cfb = readCfb(bytes);
  } catch {
    return undefined;
  }
  const parent = readParentControl(bytes);
  const site = parent?.sites[siteIndex];
  const o = cfb.readPath("/o");
  const f = cfb.readPath("/f");
  if (!parent || !site || !o || !f) return undefined;
  const at = site.objectStreamAt ?? 0;
  const len = site.objectStreamSize ?? 0;
  if (!len || at + len > o.length) return undefined;

  // The child carries no class id of its own, so the writer is told what it is - the same thing
  // the reader was told, from the site's ClsidCacheIndex.
  const kind = site.kind;
  if (!kind || kind === "form" || kind === "frame" || kind === "multiPage" || kind === "tabStrip") return undefined;
  const child = setActiveXValue(o.subarray(at, at + len), value, kind as ActiveXKind);
  if (!child) return undefined;

  const nextO = new Uint8Array(o.length - len + child.length);
  nextO.set(o.subarray(0, at), 0);
  nextO.set(child, at);
  nextO.set(o.subarray(at + len), at + child.length);

  const overrides = new Map<number, Uint8Array>();
  const oIndex = cfb.paths().find((p) => p.path === "/o")?.index;
  if (oIndex === undefined) return undefined;
  overrides.set(oIndex, nextO);

  // A change of length means the parent's record of this child's size is now wrong.
  if (child.length !== len) {
    if (site.sizeFieldAt === undefined) return undefined; // nowhere to correct it
    const nextF = new Uint8Array(f);
    new DataView(nextF.buffer).setUint32(site.sizeFieldAt, child.length, true);
    const fIndex = cfb.paths().find((p) => p.path === "/f")?.index;
    if (fIndex === undefined) return undefined;
    overrides.set(fIndex, nextF);
  }

  let out: Uint8Array;
  try {
    out = writeCfb(cfb, overrides);
  } catch {
    return undefined;
  }
  // Read it back: a storage that cannot be parsed is worse than a refusal.
  const check = readParentControl(out);
  return check?.sites[siteIndex]?.control?.value === value ? out : undefined;
}
