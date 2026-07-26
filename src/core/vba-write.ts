import { isCfb, readCfb, type CfbFile } from "./cfb";
import { writeCfb } from "./cfb-write";
import { compressOvba } from "./ovba-compress";
import { decompressOvba, readVbaProject } from "./vba";

// ---------------------------------------------------------------------------
// Writing macro source back (Stage 5 of _plans/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// One module's source is replaced inside an existing vbaProject.bin. Everything else in the project
// is carried across untouched, because most of it is not modelled here and inventing values for it
// would be worse than leaving it alone.
//
// Two things are deliberately not preserved:
//
//   - The module stream's performance cache, the compiled p-code Excel keeps in front of the
//     source. Leaving a stale cache next to changed source is the one outcome that must not happen:
//     the workbook would then show one thing and run another. So the cache is dropped and its
//     MODULEOFFSET set to 0, which leaves the source as the only description of the macro.
//   - Nothing else. The dir stream is patched in place, record by record, not rebuilt.
//
// The result is verified by reading it back before the caller is allowed to use it, so a write that
// produced something sheetedit cannot itself parse never reaches a saved file.

/** Encode source with the project's code page. Only the encodings a browser actually carries. */
function encodeText(text: string, codePage: number): Uint8Array {
  if (codePage === 65001) return new TextEncoder().encode(text);
  // TextEncoder only speaks UTF-8, so a single-byte code page is written by code unit. Anything
  // outside it would be silently mangled, so it is refused instead.
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c > 0xff) throw new VbaWriteError(`"${text[i]}" cannot be written in this workbook's code page (${codePage})`);
    out[i] = c;
  }
  return out;
}

export class VbaWriteError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "VbaWriteError";
  }
}

/** Walk the decompressed dir stream, calling back with every record's id, body offset and size. */
function forEachDirRecord(dir: Uint8Array, visit: (id: number, body: number, size: number) => void): void {
  const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  let pos = 0;
  while (pos + 6 <= dir.length) {
    const id = dv.getUint16(pos, true);
    // PROJECTVERSION's 4-byte field is Reserved, not a size; reading it as one loses the framing.
    const size = id === 0x0009 ? 6 : dv.getUint32(pos + 2, true);
    if (size > dir.length) break;
    visit(id, pos + 6, size);
    pos = pos + 6 + size;
  }
}

/**
 * Replace one module's source inside a vbaProject.bin and return the new bytes.
 * Throws a VbaWriteError, having changed nothing, when the project cannot be rewritten faithfully.
 */
export function setModuleSource(bin: Uint8Array, moduleName: string, source: string): Uint8Array {
  if (!isCfb(bin)) throw new VbaWriteError("this workbook's macro project is not a compound file");
  let cfb: CfbFile;
  try {
    cfb = readCfb(bin);
  } catch (e) {
    throw new VbaWriteError(`this workbook's macro project could not be read: ${(e as Error).message}`);
  }
  const project = readVbaProject(bin);
  if (!project) throw new VbaWriteError("this workbook has no macro project to write to");
  const codePage = project.codePage;

  const streams = cfb.paths();
  const dirEntry = streams.find((s) => /\/vba\/dir$/i.test(s.path));
  if (!dirEntry) throw new VbaWriteError("this macro project has no dir stream");
  const dir = decompressOvba(cfb.read(dirEntry.index));
  if (!dir.length) throw new VbaWriteError("this macro project's dir stream could not be read");

  // --- find the module's records --------------------------------------------
  const nameBytes = encodeText(moduleName, codePage);
  const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  let inModule = false;
  let streamName = "";
  let offsetAt = -1;
  forEachDirRecord(dir, (id, body, size) => {
    if (id === 0x0019) {
      inModule = size === nameBytes.length && nameBytes.every((b, i) => dir[body + i] === b);
      if (inModule) streamName = moduleName;
      return;
    }
    if (!inModule) return;
    if (id === 0x001a) streamName = decodeAscii(dir.subarray(body, body + size));
    if (id === 0x0031 && size >= 4) offsetAt = body;
  });
  if (offsetAt < 0) throw new VbaWriteError(`there is no module called ${moduleName} in this project`);

  const streamHit = streams.find((s) => new RegExp(`/vba/${escapeRe(streamName)}$`, "i").test(s.path));
  if (!streamHit) throw new VbaWriteError(`the stream for ${moduleName} is missing from this project`);

  // --- rewrite ---------------------------------------------------------------
  // The module stream becomes the compressed source alone, with no performance cache in front:
  // a stale cache next to new source is the one thing that would make the workbook lie.
  const newStream = compressOvba(encodeText(source, codePage));
  dv.setUint32(offsetAt, 0, true);
  const newDir = compressOvba(dir);

  const overrides = new Map<number, Uint8Array>([
    [dirEntry.index, newDir],
    [streamHit.index, newStream],
  ]);
  let out: Uint8Array;
  try {
    out = writeCfb(cfb, overrides);
  } catch (e) {
    throw new VbaWriteError(`this workbook's macro project could not be rewritten: ${(e as Error).message}`);
  }

  // --- verify before handing it back ----------------------------------------
  // The plan's rule for this stage: read it back, and refuse the result rather than let a file
  // sheetedit cannot itself parse reach the user's disk.
  const check = readVbaProject(out);
  const written = check?.modules.find((m) => m.name === moduleName);
  if (!written) throw new VbaWriteError("the rewritten macro project did not read back correctly");
  if (written.source !== source) throw new VbaWriteError("the rewritten macro source did not read back unchanged");
  if (check!.modules.length !== project.modules.length) {
    throw new VbaWriteError("the rewritten macro project lost a module");
  }
  return out;
}

const escapeRe = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

function decodeAscii(bytes: Uint8Array): string {
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return s;
}
