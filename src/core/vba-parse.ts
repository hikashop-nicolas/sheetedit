import type { Arg, BinOp, CaseClause, CaseTest, Expr, Module, Param, Procedure, Stmt, VarDecl } from "./vba-ast";
import { lex, VbaSyntaxError, type Token } from "./vba-lex";

// ---------------------------------------------------------------------------
// VBA parser (Stage 1 of _plans/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// Recursive descent over the token stream, following [MS-VBAL]'s grammar for the constructs that
// appear in real macros. Hand-written rather than generated: the rest of this codebase parses by
// hand, and a parser generator's runtime would be the largest dependency we ship.
//
// The two things that make VBA awkward to parse:
//   - statements end at a line break OR a colon, so the separator is a token and every statement
//     parser has to stop at one
//   - `f(x)` is a call, an array index, or a parenthesised argument to a call written without
//     parentheses, and which one it is cannot be known until runtime. The AST keeps them the same
//     shape and lets the interpreter decide.

/** Precedence, lowest binding first. Within a level, operators are left-associative. */
const PRECEDENCE: BinOp[][] = [
  ["Imp"],
  ["Eqv"],
  ["Xor"],
  ["Or"],
  ["And"],
  // Not is unary and binds tighter than And, handled in parseUnary.
  ["=", "<>", "<", ">", "<=", ">=", "Is", "Like"],
  ["&"],
  ["+", "-"],
  ["Mod"],
  ["\\"],
  ["*", "/"],
  // ^ binds tightest and is right-associative, handled in parsePower.
];

class Parser {
  private i = 0;
  /** Inside a single-line If, where a line break ends the whole statement rather than one part. */
  private inline = false;
  constructor(private readonly tokens: Token[]) {}

  private get tok(): Token { return this.tokens[this.i]!; }
  private get line(): number { return this.tok.line; }
  private at(upper: string): boolean { return this.tok.upper === upper; }
  private atAny(...uppers: string[]): boolean { return uppers.includes(this.tok.upper); }
  private next(): Token { return this.tokens[this.i++]!; }
  private eat(upper: string): boolean {
    if (this.tok.upper === upper) { this.i++; return true; }
    return false;
  }
  private expect(upper: string): Token {
    if (this.tok.upper !== upper) this.fail(`expected ${upper}, found ${JSON.stringify(this.tok.text)}`);
    return this.next();
  }
  private fail(message: string): never { throw new VbaSyntaxError(message, this.line); }

  /** A statement ends at a line break, a colon, or the end of input. */
  private atStatementEnd(): boolean { return this.tok.kind === "eol" || this.at(":") || this.tok.kind === "eof"; }
  private skipSeparators(): void {
    while (this.tok.kind === "eol" || this.at(":")) this.i++;
  }
  private endStatement(): void {
    if (this.tok.kind === "eof") return;
    // ELSE can only follow a statement inside a single-line If, where it ends the body without
    // being consumed here: the If parser needs to see it.
    if (this.at("ELSE")) return;
    // In a single-line If the line break belongs to the If, not to the statement inside it, or the
    // body would run on and swallow whatever follows.
    if (this.inline && this.tok.kind === "eol") return;
    if (this.tok.kind === "eol" || this.at(":")) { this.i++; return; }
    this.fail(`unexpected ${JSON.stringify(this.tok.text)} after statement`);
  }

  // --- expressions -----------------------------------------------------------

  parseExpr(level = 0): Expr {
    if (level >= PRECEDENCE.length) return this.parseUnary();
    let left = this.parseExpr(level + 1);
    for (;;) {
      const ops = PRECEDENCE[level]!;
      // Operator words (Mod, And, Is, Like) arrive as keywords, symbols as ops.
      const hit = ops.find((o) => this.tok.upper === o.toUpperCase());
      if (!hit) return left;
      const line = this.line;
      this.next();
      const right = this.parseExpr(level + 1);
      left = { type: "binary", op: hit, left, right, line };
    }
  }

  private parseUnary(): Expr {
    const line = this.line;
    if (this.at("NOT")) { this.next(); return { type: "unary", op: "Not", operand: this.parseUnary(), line }; }
    if (this.at("-")) { this.next(); return { type: "unary", op: "-", operand: this.parseUnary(), line }; }
    if (this.at("+")) { this.next(); return { type: "unary", op: "+", operand: this.parseUnary(), line }; }
    return this.parsePower();
  }

  /** `^` binds tightest and associates to the right: 2^3^2 is 2^(3^2). */
  private parsePower(): Expr {
    const line = this.line;
    const base = this.parsePostfix();
    if (this.at("^")) {
      this.next();
      // The exponent may itself be signed, as in 2^-1.
      return { type: "binary", op: "^", left: base, right: this.parseUnary(), line };
    }
    return base;
  }

  /** Member access and call/index, which chain: `a.b(1).c`. */
  private parsePostfix(): Expr {
    let expr = this.parsePrimary();
    for (;;) {
      const line = this.line;
      if (this.at(".")) {
        this.next();
        expr = { type: "member", target: expr, name: this.identName(), line };
        continue;
      }
      if (this.at("!")) {
        // `rs!Field` is shorthand for a default-member lookup by name.
        this.next();
        expr = { type: "member", target: expr, name: this.identName(), line };
        continue;
      }
      if (this.at("(")) {
        this.next();
        const args = this.parseArgs(")");
        this.expect(")");
        expr = { type: "call", target: expr, args, line };
        continue;
      }
      return expr;
    }
  }

  private identName(): string {
    if (this.tok.kind !== "ident" && this.tok.kind !== "keyword") this.fail(`expected a name, found ${JSON.stringify(this.tok.text)}`);
    return this.next().text;
  }

  private parseArgs(closer: string): Arg[] {
    const args: Arg[] = [];
    if (this.tok.upper === closer) return args;
    for (;;) {
      // An omitted argument leaves its slot empty: Foo(1, , 3).
      if (this.at(",")) { args.push({ value: null }); this.next(); continue; }
      if (this.tok.upper === closer) { args.push({ value: null }); break; }
      let name: string | undefined;
      // `name:=value` is a named argument; the ":=" lexes as ":" then "=".
      if ((this.tok.kind === "ident") && this.tokens[this.i + 1]?.upper === ":" && this.tokens[this.i + 2]?.upper === "=") {
        name = this.next().text;
        this.i += 2;
      }
      args.push({ name, value: this.parseExpr() });
      if (!this.eat(",")) break;
    }
    return args;
  }

  private parsePrimary(): Expr {
    const t = this.tok;
    const line = t.line;
    switch (t.kind) {
      case "number": this.next(); return { type: "lit", value: t.value as number, line };
      case "string": this.next(); return { type: "lit", value: t.value as string, line };
      case "date": this.next(); return { type: "date", text: String(t.value), line };
      case "ident": this.next(); return { type: "name", name: t.text, line };
    }
    if (this.at("TRUE")) { this.next(); return { type: "lit", value: true, line }; }
    if (this.at("FALSE")) { this.next(); return { type: "lit", value: false, line }; }
    if (this.at("NOTHING")) { this.next(); return { type: "nothing", line }; }
    if (this.at("EMPTY")) { this.next(); return { type: "empty", line }; }
    if (this.at("NULL")) { this.next(); return { type: "lit", value: null, line }; }
    if (this.at("NEW")) { this.next(); return { type: "new", className: this.identName(), line }; }
    if (this.at("(")) {
      this.next();
      const inner = this.parseExpr();
      this.expect(")");
      return inner;
    }
    // A leading dot is a member of the enclosing With block.
    if (this.at(".")) {
      this.next();
      return { type: "member", target: null, name: this.identName(), line };
    }
    // Keywords are legal identifiers in many positions (a property called Value, Count, Error...).
    if (t.kind === "keyword") { this.next(); return { type: "name", name: t.text, line }; }
    this.fail(`unexpected ${JSON.stringify(t.text)}`);
  }

  // --- declarations ----------------------------------------------------------

  private parseTypeName(): string | undefined {
    if (!this.eat("AS")) return undefined;
    this.eat("NEW");
    let name = this.identName();
    // A qualified type: Excel.Range.
    while (this.at(".")) { this.next(); name += `.${this.identName()}`; }
    return name;
  }

  private parseVarDecl(): VarDecl {
    const name = this.identName();
    const decl: VarDecl = { name };
    if (this.at("(")) {
      this.next();
      decl.isArray = true;
      decl.dims = [];
      if (!this.at(")")) {
        for (;;) {
          const first = this.parseExpr();
          if (this.eat("TO")) decl.dims.push({ lower: first, upper: this.parseExpr() });
          else decl.dims.push({ upper: first });
          if (!this.eat(",")) break;
        }
      }
      this.expect(")");
    }
    decl.asType = this.parseTypeName();
    return decl;
  }

  // --- statements ------------------------------------------------------------

  /** Parse statements until one of `enders` is the next keyword. Does not consume the ender. */
  private parseBlock(enders: string[]): Stmt[] {
    const body: Stmt[] = [];
    for (;;) {
      this.skipSeparators();
      if (this.tok.kind === "eof") return body;
      if (enders.includes(this.tok.upper)) return body;
      // "End If"/"End Sub" and friends: the ender is two words.
      if (this.at("END") && enders.includes(`END ${this.tokens[this.i + 1]?.upper ?? ""}`)) return body;
      body.push(this.parseStatement());
    }
  }

  private parseStatement(): Stmt {
    const line = this.line;

    // A label is a name at the start of a line followed by a colon.
    if (this.tok.kind === "ident" && this.tokens[this.i + 1]?.upper === ":" && this.tokens[this.i + 2]?.kind !== "op") {
      const name = this.next().text;
      this.next();
      return { type: "label", name, line };
    }

    if (this.atAny("DIM", "PUBLIC", "PRIVATE", "STATIC") && this.isDeclaration()) return this.parseDim(line);
    if (this.at("CONST")) return this.parseConst(line);
    if (this.at("REDIM")) return this.parseRedim(line);
    if (this.at("IF")) return this.parseIf(line);
    if (this.at("FOR")) return this.parseFor(line);
    if (this.at("DO")) return this.parseDo(line);
    if (this.at("WHILE")) return this.parseWhile(line);
    if (this.at("SELECT")) return this.parseSelect(line);
    if (this.at("WITH")) return this.parseWith(line);
    if (this.at("EXIT")) return this.parseExit(line);
    if (this.at("ON")) return this.parseOnError(line);
    if (this.at("GOTO")) { this.next(); const label = this.identName(); this.endStatement(); return { type: "goto", label, line }; }
    if (this.at("SET")) {
      this.next();
      // Not parseExpr: that would swallow "r = Range(..)" whole as an equality comparison.
      const target = this.parsePostfix();
      this.expect("=");
      const value = this.parseExpr();
      this.endStatement();
      return { type: "set", target, value, line };
    }
    if (this.at("CALL")) {
      this.next();
      const expr = this.parseExpr();
      this.endStatement();
      return { type: "callStmt", expr, line };
    }
    if (this.at("LET")) this.next(); // `Let x = 1` is the old spelling of a plain assignment
    // Anything the runtime cannot do is still parsed, so it can be refused by name rather than
    // silently skipped. Attribute lines come from the extracted source and carry no behaviour.
    if (this.at("ATTRIBUTE") || this.at("OPTION")) {
      const parts: string[] = [];
      while (!this.atStatementEnd()) parts.push(this.next().text);
      this.endStatement();
      return { type: "unsupported", text: parts.join(" "), line };
    }

    // Otherwise: an assignment, or a call written without parentheses. The target is parsed as an
    // lvalue rather than a full expression, because at statement level "=" assigns; inside an
    // expression the same token compares, and parseExpr would take "x = 1" as a comparison.
    const expr = this.parsePostfix();
    if (this.at("=")) {
      this.next();
      const value = this.parseExpr();
      this.endStatement();
      return { type: "assign", target: expr, value, line };
    }
    // `Foo 1, 2` is a call whose arguments follow with no parentheses.
    if (!this.atStatementEnd() && !this.atAny("THEN", "ELSE")) {
      const args = this.parseArgs("\n");
      this.endStatement();
      return { type: "callStmt", expr: { type: "call", target: expr, args, line }, line };
    }
    if (!this.atAny("THEN", "ELSE")) this.endStatement();
    return { type: "callStmt", expr, line };
  }

  /** Whether a Public/Private/Static at the top of a statement introduces a variable, not a proc. */
  private isDeclaration(): boolean {
    if (this.at("DIM")) return true;
    for (let k = this.i; k < this.tokens.length; k++) {
      const u = this.tokens[k]!.upper;
      if (u === "DIM" || u === "CONST") return true;
      if (u === "SUB" || u === "FUNCTION" || u === "PROPERTY" || u === "DECLARE" || u === "TYPE" || u === "ENUM") return false;
      if (this.tokens[k]!.kind === "eol") return false;
      if (u !== "PUBLIC" && u !== "PRIVATE" && u !== "STATIC" && u !== "FRIEND") return false;
    }
    return false;
  }

  private parseDim(line: number): Stmt {
    while (this.atAny("DIM", "PUBLIC", "PRIVATE", "STATIC", "FRIEND")) this.next();
    if (this.at("CONST")) return this.parseConst(line);
    const vars: VarDecl[] = [];
    do {
      this.eat("WITHEVENTS");
      vars.push(this.parseVarDecl());
    } while (this.eat(","));
    this.endStatement();
    return { type: "dim", vars, line };
  }

  private parseConst(line: number): Stmt {
    this.expect("CONST");
    const vars: { name: string; value: Expr }[] = [];
    do {
      const name = this.identName();
      this.parseTypeName();
      this.expect("=");
      vars.push({ name, value: this.parseExpr() });
    } while (this.eat(","));
    this.endStatement();
    return { type: "const", vars, line };
  }

  private parseRedim(line: number): Stmt {
    this.expect("REDIM");
    const preserve = this.eat("PRESERVE");
    const vars: VarDecl[] = [];
    do vars.push(this.parseVarDecl());
    while (this.eat(","));
    this.endStatement();
    return { type: "redim", preserve, vars, line };
  }

  private parseIf(line: number): Stmt {
    this.expect("IF");
    const cond = this.parseExpr();
    this.expect("THEN");
    const branches: { cond: Expr; body: Stmt[] }[] = [];
    // A single-line If has its body on the same line and no End If.
    if (!this.atStatementEnd()) {
      const wasInline = this.inline;
      this.inline = true;
      try {
        const body: Stmt[] = [this.parseStatement()];
        while (!this.atStatementEnd() && !this.at("ELSE")) body.push(this.parseStatement());
        branches.push({ cond, body });
        let elseBody: Stmt[] | undefined;
        if (this.eat("ELSE")) {
          elseBody = [];
          while (!this.atStatementEnd()) elseBody.push(this.parseStatement());
        }
        this.inline = wasInline;
        this.endStatement();
        return { type: "if", branches, else: elseBody, line };
      } finally {
        this.inline = wasInline;
      }
    }
    branches.push({ cond, body: this.parseBlock(["ELSEIF", "ELSE", "END IF"]) });
    while (this.at("ELSEIF")) {
      this.next();
      const c = this.parseExpr();
      this.expect("THEN");
      branches.push({ cond: c, body: this.parseBlock(["ELSEIF", "ELSE", "END IF"]) });
    }
    let elseBody: Stmt[] | undefined;
    if (this.eat("ELSE")) elseBody = this.parseBlock(["END IF"]);
    this.expect("END");
    this.expect("IF");
    this.endStatement();
    return { type: "if", branches, else: elseBody, line };
  }

  private parseFor(line: number): Stmt {
    this.expect("FOR");
    if (this.eat("EACH")) {
      const varName = this.identName();
      this.expect("IN");
      const collection = this.parseExpr();
      const body = this.parseBlock(["NEXT"]);
      this.expect("NEXT");
      if (this.tok.kind === "ident") this.next(); // `Next x` names the loop variable
      this.endStatement();
      return { type: "forEach", varName, collection, body, line };
    }
    const varName = this.identName();
    this.expect("=");
    const from = this.parseExpr();
    this.expect("TO");
    const to = this.parseExpr();
    const step = this.eat("STEP") ? this.parseExpr() : undefined;
    const body = this.parseBlock(["NEXT"]);
    this.expect("NEXT");
    if (this.tok.kind === "ident") this.next();
    this.endStatement();
    return { type: "for", varName, from, to, step, body, line };
  }

  private parseDo(line: number): Stmt {
    this.expect("DO");
    let test: { kind: "While" | "Until"; cond: Expr; atEnd: boolean } | undefined;
    if (this.atAny("WHILE", "UNTIL")) {
      const kind = this.next().upper === "WHILE" ? "While" : "Until";
      test = { kind, cond: this.parseExpr(), atEnd: false };
    }
    const body = this.parseBlock(["LOOP"]);
    this.expect("LOOP");
    if (this.atAny("WHILE", "UNTIL")) {
      const kind = this.next().upper === "WHILE" ? "While" : "Until";
      // A test at the end runs the body once before checking, which is the point of Do...Loop While.
      test = { kind, cond: this.parseExpr(), atEnd: true };
    }
    this.endStatement();
    return { type: "doLoop", test, body, line };
  }

  private parseWhile(line: number): Stmt {
    this.expect("WHILE");
    const cond = this.parseExpr();
    const body = this.parseBlock(["WEND"]);
    this.expect("WEND");
    this.endStatement();
    return { type: "while", cond, body, line };
  }

  private parseSelect(line: number): Stmt {
    this.expect("SELECT");
    this.expect("CASE");
    const subject = this.parseExpr();
    this.skipSeparators();
    const cases: CaseClause[] = [];
    let elseBody: Stmt[] | undefined;
    while (this.at("CASE")) {
      this.next();
      if (this.eat("ELSE")) {
        elseBody = this.parseBlock(["CASE", "END SELECT"]);
        break;
      }
      const tests: CaseTest[] = [];
      do {
        if (this.eat("IS")) {
          const op = this.next().upper as BinOp;
          tests.push({ kind: "compare", op: (op === "<>" ? "<>" : op) as BinOp, value: this.parseExpr() });
          continue;
        }
        const first = this.parseExpr();
        if (this.eat("TO")) tests.push({ kind: "range", from: first, to: this.parseExpr() });
        else tests.push({ kind: "value", value: first });
      } while (this.eat(","));
      cases.push({ tests, body: this.parseBlock(["CASE", "END SELECT"]) });
    }
    this.expect("END");
    this.expect("SELECT");
    this.endStatement();
    return { type: "select", subject, cases, elseBody, line };
  }

  private parseWith(line: number): Stmt {
    this.expect("WITH");
    const subject = this.parseExpr();
    const body = this.parseBlock(["END WITH"]);
    this.expect("END");
    this.expect("WITH");
    this.endStatement();
    return { type: "with", subject, body, line };
  }

  private parseExit(line: number): Stmt {
    this.expect("EXIT");
    const w = this.next().upper;
    const what = w === "SUB" ? "Sub" : w === "FUNCTION" ? "Function" : w === "PROPERTY" ? "Property" : w === "FOR" ? "For" : "Do";
    this.endStatement();
    return { type: "exit", what, line };
  }

  private parseOnError(line: number): Stmt {
    this.expect("ON");
    if (!this.eat("ERROR")) {
      // `On x GoTo ...` (computed goto) is not modelled; keep it so it can be refused by name.
      const parts: string[] = ["On"];
      while (!this.atStatementEnd()) parts.push(this.next().text);
      this.endStatement();
      return { type: "unsupported", text: parts.join(" "), line };
    }
    if (this.eat("RESUME")) {
      this.expect("NEXT");
      this.endStatement();
      return { type: "onError", mode: "resumeNext", label: "", line };
    }
    this.expect("GOTO");
    const label = this.tok.kind === "number" ? String(this.next().value) : this.identName();
    this.endStatement();
    return { type: "onError", mode: "goto", label, line };
  }

  // --- procedures and modules ------------------------------------------------

  private parseParams(): Param[] {
    const params: Param[] = [];
    this.expect("(");
    if (this.at(")")) { this.next(); return params; }
    for (;;) {
      const optional = this.eat("OPTIONAL");
      const paramArray = this.eat("PARAMARRAY");
      let byVal = false;
      if (this.eat("BYVAL")) byVal = true;
      else this.eat("BYREF"); // ByRef is the default, so saying it changes nothing
      const name = this.identName();
      // A parameter can be declared as an array with empty parentheses.
      if (this.at("(") && this.tokens[this.i + 1]?.upper === ")") { this.i += 2; }
      const asType = this.parseTypeName();
      const def = this.eat("=") ? this.parseExpr() : undefined;
      params.push({ name, byVal, optional, paramArray, asType, default: def });
      if (!this.eat(",")) break;
    }
    this.expect(")");
    return params;
  }

  private parseProcedure(): Procedure {
    const line = this.line;
    let isPublic = true;
    while (this.atAny("PUBLIC", "PRIVATE", "FRIEND", "STATIC")) {
      if (this.at("PRIVATE")) isPublic = false;
      this.next();
    }
    let kind: Procedure["kind"];
    let ender: string;
    if (this.eat("SUB")) { kind = "sub"; ender = "SUB"; }
    else if (this.eat("FUNCTION")) { kind = "function"; ender = "FUNCTION"; }
    else {
      this.expect("PROPERTY");
      const which = this.next().upper;
      kind = which === "GET" ? "propertyGet" : which === "SET" ? "propertySet" : "propertyLet";
      ender = "PROPERTY";
    }
    const name = this.identName();
    const params = this.at("(") ? this.parseParams() : [];
    const asType = this.parseTypeName();
    const body = this.parseBlock([`END ${ender}`]);
    this.expect("END");
    this.next();
    return { kind, name, params, asType, isPublic, body, line };
  }

  parseModule(name: string): Module {
    const declarations: Stmt[] = [];
    const procedures: Procedure[] = [];
    const options: string[] = [];
    for (;;) {
      this.skipSeparators();
      if (this.tok.kind === "eof") break;
      if (this.at("OPTION")) {
        const parts: string[] = [];
        this.next();
        while (!this.atStatementEnd()) parts.push(this.next().text);
        options.push(parts.join(" "));
        this.endStatement();
        continue;
      }
      // A procedure is the only thing whose header runs to an End; everything else is a declaration.
      const startsProc = this.startsProcedure();
      if (startsProc) { procedures.push(this.parseProcedure()); continue; }
      declarations.push(this.parseStatement());
    }
    return { name, declarations, procedures, options };
  }

  private startsProcedure(): boolean {
    for (let k = this.i; k < this.tokens.length; k++) {
      const t = this.tokens[k]!;
      if (t.kind === "eol") return false;
      if (t.upper === "SUB" || t.upper === "FUNCTION" || t.upper === "PROPERTY") return true;
      if (!["PUBLIC", "PRIVATE", "FRIEND", "STATIC"].includes(t.upper)) return false;
    }
    return false;
  }
}

/** Parse one module's source. Throws VbaSyntaxError with the line on bad input. */
export function parseModule(source: string, name = "Module"): Module {
  return new Parser(lex(source)).parseModule(name);
}

/** Parse a single expression, for tests and for evaluating a watch. */
export function parseExpression(source: string): Expr {
  const p = new Parser(lex(source));
  return (p as unknown as { parseExpr(level?: number): Expr }).parseExpr(0);
}
