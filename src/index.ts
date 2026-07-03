// sheetedit: a standalone, framework-agnostic, client-side spreadsheet editor for
// .xlsx (OOXML SpreadsheetML) and .ods (ODF spreadsheet). Both are zips of XML.
//
// Philosophy (same as the docx/odt siblings): edit in place and preserve everything
// untouched. See model.ts (types + helpers), xlsx.ts / ods.ts (format adapters),
// recalc.ts (the formula engine), workbook.ts (public read/write) and editor.ts
// (the grid UI). This entry point re-exports the public surface.
export * from "./model";
export * from "./xlsx";
export * from "./ods";
export * from "./recalc";
export * from "./workbook";
export * from "./editor";
