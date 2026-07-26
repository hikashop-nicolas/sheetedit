// The VBA engine lives in its own library now (vbalang): the container, the language, the runtime
// and the writer, none of which know anything about spreadsheets. What stays here is the part that
// does: where an xlsx keeps its macro project, and the Excel object model in vba-excel.ts.

export { decompressOvba, isSigned, readVbaProject, subNames, type VbaModule, type VbaProject } from "vbalang";

/** The macro-bearing part of a workbook, if it has one. */
export const vbaPartOf = (files: Record<string, Uint8Array>): Uint8Array | undefined =>
  files["xl/vbaProject.bin"] ?? files["xl/vbaproject.bin"];
