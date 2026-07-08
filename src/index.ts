// sheetedit: a standalone, framework-agnostic, client-side spreadsheet editor for
// .xlsx (OOXML SpreadsheetML) and .ods (ODF spreadsheet). Both are zips of XML.
//
// Philosophy (same as the docx/odt siblings): edit in place and preserve everything
// untouched. See model.ts (types + helpers), xlsx.ts / ods.ts (format adapters),
// recalc.ts (the formula engine), workbook.ts (public read/write) and editor.ts
// (the grid UI). This entry point re-exports the public surface.
export * from "./core/model";
export * from "./core/dates";
export * from "./core/structure";
export * from "./adapters/xlsx";
export * from "./adapters/ods";
export * from "./core/recalc";
export * from "./core/workbook";
export * from "./core/editor";
