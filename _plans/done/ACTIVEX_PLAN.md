# sheetedit: ActiveX controls, end to end

The longest single thread in the project, and the one with the most ways to be quietly wrong. It
began as "the .bin is untouchable", which was false: [MS-OFORMS] is a public normative
specification, so a Forms 2.0 control is parseable on the same clean-room basis as the VBA
container. Only a THIRD-PARTY COM control is genuinely beyond reach, because OOXML says its content
"shall be solely determined by the corresponding object".

Finished 2026-07-28: leaf controls read and written, per-column widths, captions authored, and the
container controls (Frame / MultiPage / TabStrip) read, drawn and written back.

## What each piece cost

- **Preserved-only**: ActiveX (carried through a save, never rendered). The macro dialog now says so,
  since the control otherwise leaves an unexplained gap on the grid.
  - RESEARCHED 2026-07-26, and it corrects an earlier claim here that the `.bin` is untouchable.
    **[MS-OFORMS]** is a public, normative Microsoft Open Specification for exactly this binary
    persistence, so a Forms 2.0 control (which is what Excel's ActiveX toolbox inserts) is
    parseable on the same clean-room basis as MS-OVBA. A THIRD-PARTY COM control still is not: the
    OOXML spec says the content "shall be solely determined by the corresponding object", so its
    format belongs to whoever wrote the control.
  - The surprise is where the data sits. A real `activeX1.xml` (checked against one in the wild)
    carries ONLY `ax:classid` plus a relationship to the `.bin` when persistence is
    `persistStreamInit`. Caption, value, size, colours and the linked cell are all in the binary.
    So there is no cheap "read the XML and render it" half: the work IS the MS-OFORMS parser.
  - DONE for reading: CommandButton and every MorphData kind (checkbox, combo, text, list, option,
    toggle, label) give up their kind, caption and value. The parse self-checks against the stream's
    own `cb` field and returns the kind alone rather than half-right values when it does not land
    exactly. ScrollBar and SpinButton stay kind-only: their mask's bit order is unconfirmed.
  - Two bugs it uncovered, both worse than the missing feature. Every ActiveX control was read as a
    blank "label" because its part was fed to the formControlPr reader; and each was read TWICE,
    since Excel writes a control under both mc:Choice and mc:Fallback. A workbook with six controls
    drew twelve phantom labels.
  - Samples: found via a plain web search, not GitHub code search, which does not index binaries.
    Contextures' combo-box tutorial workbook has six. It is their copyright, so it is used locally
    and never committed; the committed fixtures are synthetic and were verified byte-identical to
    the real streams before being written down.
  - `linkedCell`, `listFillRange` and `macro` are read off `<controlPr>`, which is where Excel keeps
    the properties that are its own rather than the control's. A combo with both is fully live: it
    lists from the named range and writes the chosen TEXT to the linked cell, where a form control
    writes the item's position. Verified against the real file, whose three combos carry
    RegionList / MonthList / DayList and linked cells H9 / H5 / C7.
  - DONE: an MS-OFORMS writer for the Value. Same length patches IN PLACE so the stream stays
    byte-identical including padding; a length change rebuilds the ExtraDataBlock and carries the
    unmodelled trailing blocks (StreamData, TextProps, rgColumnInfo) across. It refuses on any
    control the reader would not vouch for, so a write never proceeds where a read would not.
    Identity confirmed on four real Excel streams, and a change survives editor -> save -> reread.
  - VERIFICATION CAVEAT, and it is a real one. LibreOffice does not surface ActiveX from xlsx, so
    unlike the VBA writer (which an independent engine was made to RUN) there is no outside judge.
    It reopens a rewritten workbook without complaint, and that only proves the package is not
    corrupt. Excel remains the untested case.
  - ScrollBar, SpinButton and Label are read too, from their own spec tables fetched rather than
    guessed. Three findings worth keeping: a LABEL IS NOT A MORPHDATA CONTROL, it has its own
    LabelPropMask, which is an easy and wrong assumption; ScrollBar and SpinButton do NOT share a
    mask (the spin has no LargeChange or ProportionalThumb, and fMousePointer moves to the end);
    and fPrevEnabled / fNextEnabled are mask bits with NO field behind them, so consuming bytes for
    them would push every later read out of place. The layouts are table-driven now, one table per
    family, each ending on the same cb check.
  - Writing covers every string a control carries (Value, Caption, GroupName), not just Value. The
    layout records each string's own place, so rewriting the middle of three moves the last
    correctly. Identity holds on all nine real streams for every property they carry.
  - SPEC AUDIT (2026-07-27), against [MS-OFORMS] section 2 property by property. What used to be
    read and thrown away is now kept and used:
    - **VariousPropertyBits**, the 32-bit field a dozen booleans share. Enabled, Locked, BackStyle,
      ColumnHeads, MatchRequired, Alignment, Editable, WordWrap, AutoSize, MultiLine. The bit
      positions are pinned by the spec's own file-format defaults (0x2C80081B for the MorphData
      family, 0x0080001B for a label), which is a real check on having read the table straight.
    - **DisplayStyle**, which is the ONLY thing separating an editable combo (3) from a drop-list
      one (7): they share a class id. Also the numeric properties around it - MaxLength,
      PasswordChar, BorderStyle, BorderColor, SpecialEffect, ScrollBars, ListRows, ListWidth,
      ColumnCount, BoundColumn, TextColumn, MultiSelect, MatchEntry, ListStyle, ShowDropButtonWhen,
      DropButtonStyle, MousePointer, Accelerator, PicturePosition - and SmallChange / LargeChange /
      Orientation / Delay / ProportionalThumb on the range controls.
    - **TextProps**, the font, which is its own versioned structure sitting after the control's cb.
      Name, size (twips), bold / italic / underline / strikeout, weight, paragraph alignment.
    - **The Image control**, which had no layout at all. Its mask puts fAutoSize and fPictureTiling
      in the BIT rather than in a field, like the scroll bar's fPrevEnabled.
    - **StreamData pictures**, sniffed to a MIME type and rendered from a data: URI. A metafile
      (WMF/EMF) is a drawing program rather than an image, so it is skipped rather than mislabelled.
  - What that buys on the page: a TextBox is an editor (a textarea when MultiLine, a password box
    when it names a PasswordChar, with its MaxLength), a ToggleButton is a button that stays down,
    an Image shows its picture, a list honours ListRows and MultiSelect, every control wears the
    file's own font and colours, a disabled control looks and behaves disabled, MousePointer is a
    CSS cursor and Accelerator is an access key. An OLE_COLOR naming a Windows SYSTEM palette entry
    is left unset on purpose: its colour is the desktop theme's, not the document's.
  - The writer can now ADD a property the control does not yet carry, which is what an empty text
    box needs: its Value bit is clear, so there is nothing to patch. That path re-emits the whole
    DataBlock from the fields the read recorded (splicing would not do, since a length word has to
    land 4-aligned and inserting one shifts everything after it) and reads its own output back
    before returning it. Same-length changes still patch in place, byte for byte.
  - **Multi-column lists** render as a grid, since that is what Excel draws and a `<select>` cannot
    be one. The columns come from the source range (which is where the items come from anyway) and
    BoundColumn decides which one the control reports - 0 means the row number, as in Excel.
- **Windows metafiles (WMF / EMF)** render now, which they never did: an `xl/media/*.emf` came
  through as a `data:image/emf` URI and a browser drew nothing at all. A metafile is not an image -
  it is a recorded list of GDI drawing calls - so showing one means replaying them onto a canvas.
  That is emf-converter (Apache-2.0, no dependencies), lazy-loaded so a workbook without one never
  pays for the code, and failing soft: no picture rather than a broken one. It covers sheet images,
  an ActiveX control's Picture, and anything else that arrives as a data URI.
  - GOTCHA, and it is ours to handle: the converter ignores a placeable WMF's own frame and renders
    into a square canvas at its 8192px cap - a 1.7MB PNG of mostly white for a picture two inches
    across. `metafileSize` reads the frame first (a WMF's placeable bounding box and units-per-inch,
    an EMF's rclFrame in hundredths of a millimetre) and passes it as an explicit cap, which puts
    both formats on the same sane size.
  - Text inside a metafile used to draw at the file's LOGICAL height, so a label came out many times
    too large and read as a white block punched through the drawing. Diagnosed against LibreOffice's
    own rasterisation of the same files, fixed upstream (emf-converter#9, in 2.0.2): the font height
    now goes through the same window/viewport mapping as every coordinate.
- **Worksheet.OLEObjects**: the ActiveX controls, reachable from a macro. `.Count`, by name or by
  1-based index, `For Each`, and per object `.Object`, `.Name`, `.Index`, `.LinkedCell`,
  `.ListFillRange`, `.TopLeftCell`. The control itself gives `.Value` / `.Text` / `.Caption` /
  `.ListCount` / `.ListIndex` / `.List(i)`, and Value and ListIndex are settable: the write goes to
  the persisted binary AND to the linked cell, since Excel keeps the two in step. Everything else
  refuses by name. The list comes from the host, because a listFillRange is usually a DEFINED NAME.

## The named-for-later list, as it ended

## Named for later (nothing here is forgotten, and none of it is a mystery)

Each of these is understood; what stops it is stated, so picking one up starts from a known place.

| | |
|---|---|
| **rgColumnInfo: per-column widths** | DONE (2026-07-28). The record IS in the specification, just not in the HTML index that was searched: the downloadable .docx defines MorphDataColumnInfo (2.2.5.6), its PropMask (2.2.5.7) and its DataBlock (2.2.5.8). Each column is its own versioned record - `0x00 0x02`, `cbColumnInfo` over the 4-byte PropMask plus the DataBlock, one mask bit saying whether a width follows - and the width is a signed 4-byte HIMETRIC, defaulting to -1 meaning the application chooses. The array holds exactly `cColumnInfo` entries, which is the LAST column with a non-default width and so can be shorter than the column count. It follows TextProps, which the spec makes mandatory while marking rgColumnInfo optional, so the reader anchors on it. |
| **Frame / MultiPage / TabStrip** | READ + RENDERED (2026-07-28). The structure is what made these different: a parent control persists as a STORAGE, so its .bin is a compound file. "f" holds its own properties then the sites array (each child's name, type and position), "o" holds every child's properties end to end, and a MultiPage keeps a storage per page. The container comes from vbalang's compound-file reader and the children from the existing leaf reader, since a child is an ordinary control structure with no class id in front. A TabStrip turned out NOT to be a container at all - it has tabs, not children - so it is a leaf layout of its own, read down to its tab captions and selected index. Not written back: the whole storage is preserved byte-for-byte. |
| **Caption on a CommandButton / Label that has none** | DONE (2026-07-28). The flat layouts now record their fields the way the MorphData family already did, so their blocks can be rebuilt to ADD a caption rather than only patch one. The command button moved from a hand-written branch to the same table the other flat families use, which also fixed its MousePointer being read as 2 bytes where the spec says 1 (masked until now: every field before it is 4 bytes wide, so the following alignment landed in the same place). Double-clicking a control opens a caption dialog, offered only where a no-op write proves the stream writable. |
| **Third-party ActiveX** | Irreducible. OOXML says the content of such a control "shall be solely determined by the corresponding object", so the format belongs to whoever wrote the control. |

## The standing caveat

There is no Excel on the machine this was built on. LibreOffice does not surface ActiveX from
`.xlsx` at all, so unlike the VBA writer - which an independent engine could be made to RUN - there
is no outside judge. Every writer here reads its own output back before returning it, refuses any
stream the reader would not vouch for, and is verified by round-trip and by identity on real
streams. That is what it proves, and no more.
