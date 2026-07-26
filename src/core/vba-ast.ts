// The AST the VBA parser produces (Stage 1 of _plans/VBA_PLAN.md). Deliberately small: every node
// carries only what the interpreter will need, plus the line, so a runtime error can say where.

export interface Node {
  line: number;
}

// --- expressions -------------------------------------------------------------

export type Expr =
  | { type: "lit"; value: number | string | boolean | null; line: number }
  | { type: "nothing"; line: number }
  | { type: "empty"; line: number }
  | { type: "date"; text: string; line: number }
  /** A bare name: a variable, a constant, or a zero-argument call. Which one is a runtime question. */
  | { type: "name"; name: string; line: number }
  /** `a.b`, or `.b` inside a With block (target null). */
  | { type: "member"; target: Expr | null; name: string; line: number }
  /** `f(x)` and `a(i)` are the same syntax; the runtime decides which it is. */
  | { type: "call"; target: Expr; args: Arg[]; line: number }
  | { type: "unary"; op: "-" | "+" | "Not"; operand: Expr; line: number }
  | { type: "binary"; op: BinOp; left: Expr; right: Expr; line: number }
  | { type: "new"; className: string; line: number };

export type BinOp =
  | "^" | "*" | "/" | "\\" | "Mod" | "+" | "-" | "&"
  | "=" | "<>" | "<" | ">" | "<=" | ">=" | "Is" | "Like"
  | "And" | "Or" | "Xor" | "Eqv" | "Imp";

/** An argument, which may be named (`Foo bar:=1`) or omitted (`Foo 1, , 3`). */
export interface Arg {
  name?: string;
  value: Expr | null;
}

// --- statements --------------------------------------------------------------

export type Stmt =
  | { type: "dim"; vars: VarDecl[]; line: number }
  | { type: "const"; vars: { name: string; value: Expr }[]; line: number }
  | { type: "redim"; preserve: boolean; vars: VarDecl[]; line: number }
  | { type: "assign"; target: Expr; value: Expr; line: number }
  /** `Set x = ...`: assigns a reference rather than a value, which matters for objects. */
  | { type: "set"; target: Expr; value: Expr; line: number }
  /** A call written without parentheses, or with Call. */
  | { type: "callStmt"; expr: Expr; line: number }
  | { type: "if"; branches: { cond: Expr; body: Stmt[] }[]; else?: Stmt[]; line: number }
  | { type: "for"; varName: string; from: Expr; to: Expr; step?: Expr; body: Stmt[]; line: number }
  | { type: "forEach"; varName: string; collection: Expr; body: Stmt[]; line: number }
  | { type: "doLoop"; test?: { kind: "While" | "Until"; cond: Expr; atEnd: boolean }; body: Stmt[]; line: number }
  | { type: "while"; cond: Expr; body: Stmt[]; line: number }
  | { type: "select"; subject: Expr; cases: CaseClause[]; elseBody?: Stmt[]; line: number }
  | { type: "with"; subject: Expr; body: Stmt[]; line: number }
  | { type: "exit"; what: "Sub" | "Function" | "Property" | "For" | "Do"; line: number }
  | { type: "onError"; mode: "resumeNext" | "goto"; label: string; line: number }
  | { type: "goto"; label: string; line: number }
  | { type: "label"; name: string; line: number }
  /** Anything recognised but not modelled, kept so the interpreter can refuse it by name. */
  | { type: "unsupported"; text: string; line: number };

export interface VarDecl {
  name: string;
  /** `As Long`, `As Range`, ... when stated. */
  asType?: string;
  /** Present for an array: the bounds as written, empty for `Dim a()`. */
  dims?: { lower?: Expr; upper: Expr }[];
  isArray?: boolean;
}

export interface CaseClause {
  /** `Case 1, 2`, `Case Is > 5`, `Case 1 To 9`. */
  tests: CaseTest[];
  body: Stmt[];
}

export type CaseTest =
  | { kind: "value"; value: Expr }
  | { kind: "range"; from: Expr; to: Expr }
  | { kind: "compare"; op: BinOp; value: Expr };

// --- procedures and modules --------------------------------------------------

export interface Param {
  name: string;
  byVal: boolean;
  optional: boolean;
  paramArray: boolean;
  asType?: string;
  default?: Expr;
}

export interface Procedure extends Node {
  kind: "sub" | "function" | "propertyGet" | "propertyLet" | "propertySet";
  name: string;
  params: Param[];
  asType?: string;
  isPublic: boolean;
  body: Stmt[];
}

export interface Module {
  name: string;
  /** Module-level Dim/Const, run before any procedure. */
  declarations: Stmt[];
  procedures: Procedure[];
  /** `Option Explicit` and friends, kept for the interpreter to honour. */
  options: string[];
}
