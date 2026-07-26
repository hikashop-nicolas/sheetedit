# VBA macros in sheetedit

## Where this came from

sheetedit preserves `vbaProject.bin` byte-for-byte but does nothing with it. The question was whether
running macros is possible in a browser. It is: JavaScript cannot leave the tab, so a VBA engine
would be as sandboxed as the rest of the editor. The obstacle is scale, not safety.

An earlier claim of mine, that no ground exists to build on, was wrong and made without searching.
What actually exists:

| Piece | Status |
|---|---|
| **[MS-OVBA]** container format (CFB, `/VBA/dir`, module streams, compression) | Public Microsoft spec |
| **[MS-VBAL]** language spec, full ABNF grammar + runtime semantics | Public Microsoft spec, ~200pp |
| VB6/VBA ANTLR grammars with a JavaScript/TypeScript target (`antlr4-vb6-js`, `vb6-antlr4-typescript`, Rubberduck's `VBAParser.g4`) | Exist, reusable |
| MS-OVBA decompression reference implementations (`MS-OVBA-Compression`, `pyOpenVBA`) | Exist, Python; algorithm is small and documented |
| A VBA **interpreter** in JavaScript | **Does not exist** |

So the parser and the container format are solved problems with public specs. What has to be built
is the runtime and the Excel object model, which is where real macros actually live.

## Clean-room rule

Same discipline as [[mlang]]: build from the published specs only. MS-OVBA and MS-VBAL are the
sources of truth. Never decompile Office, never copy code out of a GPL/AGPL project. Reference
implementations may be read for *understanding the spec*, not transcribed.

## Guiding decision: refuse rather than approximate

A macro that half-runs is worse than one that does not run. If a macro calls something unmodelled,
the run **stops with a clear error naming what is missing**, and any changes it made are rolled back
as a single undo step. Silently no-oping `Shell`, `FileSystemObject` or an unimplemented method
would leave a workbook in a state its author never intended, and the user would then save it.

This is the one real risk, and it is a fidelity risk, not a sandbox one.

## Stages

Each stage is useful on its own and ships independently.

### Stage 0 — Read the macros (foundation, useful alone)

Extract and display the VBA source. No execution.

- CFB (compound file) reader: header, FAT/miniFAT, directory entries, stream extraction.
- MS-OVBA decompression: `0x01` signature, chunk headers (size = `hdr & 0x0FFF` + 3, signature bits
  `0b011`, bit 15 = compressed), flag bytes, literal vs copy tokens with the position-dependent
  offset/length split.
- `/VBA/dir` parsing: `PROJECTCODEPAGE`, `PROJECTMODULES`, per module `MODULESTREAMNAME` and
  `MODULEOFFSET`.
- Module source: read `/VBA/<stream>`, skip `MODULEOFFSET`, decompress, decode by code page.
- UI: a read-only macro viewer, so a user can see what a workbook does before trusting it.

Value on its own: auditing. Today a `.xlsm` is opaque.

### Stage 1 — Parse to an AST

What stage 0 taught, worth carrying forward: the `dir` stream is NOT uniformly (id, size, body).
`PROJECTVERSION` (0x0009) has a 4-byte *Reserved* field where a size would be, and its body is 6
bytes. Reading it as a size desynchronises the whole walk and yields a project with no modules,
silently. Expect more of these; validate against real files, never only against hand-built ones.


MS-VBAL's ABNF, restricted to the constructs that appear in real macros.

Decision to make when we get there: hand-written recursive descent (matches how the rest of this
codebase parses, no dependency) versus an ANTLR grammar (correct by construction, ~200KB runtime).
Lean hand-written for the subset, since sheetedit ships no parser generators today.

- Declarations (`Dim`, `Const`, `Sub`, `Function`, `Property`), types, arrays.
- Expressions with VBA's precedence, `&` concatenation, `Like`, `Is`.
- Statements: assignment, `Set`, `If/ElseIf/Else`, `For`/`For Each`, `Do/While`, `Select Case`,
  `With`, `Exit`, `Call`, `On Error`.

### Stage 2 — Interpret the language

Values and coercion (Variant, the numeric tower, `Empty`/`Null`/`Nothing`), scoping, calling
convention (ByRef default, which matters), `On Error Resume Next`, and the built-in function library
(`Left`, `Mid`, `InStr`, `Format`, `CStr`, `IsEmpty`, `UBound`, ...).

No file, network or shell surface exists to implement, so those calls hit the refusal path.

### Stage 3 — The Excel object model

The bulk of the work, and where 80% of real macros live.

- `Range` (the big one): `Value`, `Value2`, `Formula`, `Text`, `Offset`, `Resize`, `Cells`, `Rows`,
  `Columns`, `Count`, `Address`, `Interior`, `Font`, `Copy`, `ClearContents`, `Sort`, `AutoFilter`.
- `Worksheet` / `Worksheets`, `Workbook` / `ActiveWorkbook`, `Selection`, `ActiveCell`.
- `Application`: only the harmless parts (`ScreenUpdating`, `Calculation`, `WorksheetFunction`).
  `Application.Quit`, `Shell`, `CreateObject` refuse by design.
- Everything maps onto sheetedit's existing model, so a macro edit is an ordinary edit: it
  recalculates, renders and undoes like any other.

### Stage 4 — Run it

- Explicit per-macro run from a menu listing the `Sub`s that take no arguments.
- The whole run is one undo step.
- A step budget, so a runaway loop cannot hang the tab.
- Errors surface with the module, line and what was missing.

### Stage 5 — Write VBA back

Editing macro source, which means rebuilding `vbaProject.bin` rather than preserving it.

The inverse of Stage 0, and each half is independently checkable against the half we already have:

- **MS-OVBA compression.** The decompressor is the oracle: `decompress(compress(x)) === x` for
  arbitrary input, and our compressor's output must also decompress correctly in Excel. Emitting
  only raw (uncompressed) chunks is legal per the spec and is the safe first implementation, since
  bit 15 of a chunk header says whether tokens are used at all.
- **CFB writer.** Header, FAT, miniFAT, directory tree, then the streams. The existing reader is the
  round-trip oracle.
- **`dir` stream rebuild**, module streams, and keeping the parts we do not model
  (`_VBA_PROJECT`, `PROJECT`, `PROJECTwm`) byte-for-byte from the original.

Risk to respect: a malformed `vbaProject.bin` can make Excel refuse the whole workbook, not just the
macros. So the first version should modify an existing project in place (replace one module stream,
patch its `dir` record) rather than build one from nothing, and verify by reading it back before the
save is allowed to complete.

Only worth doing once Stage 1 can parse what the user typed, so an edit cannot save syntactic
nonsense into the file.

### Not planned

- **ActiveX controls.** Windows COM objects in a `.bin`. `ActiveX.js` is an IE9 shim for web pages
  calling `new ActiveXObject(...)`; it has nothing to do with the OLE objects embedded in a
  workbook. Preserve, never run.

## Where it should live

Probably its own repo, as [[mlang]] is, once past Stage 1: a VBA engine is useful beyond sheetedit
and keeping it separate keeps the clean-room boundary visible. Stage 0 (container reading) can start
inside sheetedit and move later.

## Status

- [x] Research, spec sourcing, this plan
- [x] **Stage 0** - CFB reader, MS-OVBA decompression, dir walk, macro viewer. Verified end to end
  on real `vbaProject.bin` files (Apache POI test data, see `src/fixtures/README.md`): a module's
  source comes out whole, in two different code pages.
- [ ] Stages 1-4
- [ ] Stage 5 (write back), after 4
