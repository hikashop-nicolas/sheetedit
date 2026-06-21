// Minimal ambient types for fast-formula-parser (the package ships no .d.ts).
// It is CommonJS: `module.exports = FormulaParser`, with DepParser / FormulaError
// attached as static properties. We import the default and read the statics.
declare module "fast-formula-parser" {
  export interface CellRef {
    row: number;
    col: number;
    sheet?: string;
  }
  export interface RangeRef {
    from: { row: number; col: number };
    to: { row: number; col: number };
    sheet?: string;
  }
  export type Ref = CellRef | RangeRef;

  export interface ParserConfig {
    onCell?: (ref: CellRef) => unknown;
    onRange?: (ref: RangeRef) => unknown[][];
    onVariable?: (name: string, sheet?: string) => unknown;
    functions?: Record<string, (...args: unknown[]) => unknown>;
  }

  export class DepParser {
    constructor(config?: { onVariable?: (name: string, sheet?: string) => unknown });
    parse(formula: string, position?: CellRef): Ref[];
  }

  export interface SSFModule {
    format(fmt: string | number, value: number): string;
  }

  export default class FormulaParser {
    constructor(config?: ParserConfig);
    parse(formula: string, position?: CellRef, allowReturnArray?: boolean): unknown;
    static DepParser: typeof DepParser;
    static FormulaError: unknown;
    static SSF: SSFModule;
  }
}
