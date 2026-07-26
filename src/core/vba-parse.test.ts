import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { unzipSync } from "fflate";
import { lex, VbaSyntaxError } from "./vba-lex";
import { parseExpression, parseModule } from "./vba-parse";
import { readVbaProject } from "./vba";
import type { Expr, Stmt } from "./vba-ast";

// Stage 1 of _plans/VBA_PLAN.md. The parser is checked on constructed source for the awkward
// corners, and then on the real macros in the fixtures, which is the test that matters: source
// written by Excel rather than by me.

/** A compact s-expression of an AST node, so a test reads as the shape it expects. */
function sexp(e: Expr): string {
  switch (e.type) {
    case "lit": return typeof e.value === "string" ? JSON.stringify(e.value) : String(e.value);
    case "name": return e.name;
    case "nothing": return "Nothing";
    case "empty": return "Empty";
    case "date": return `#${e.text}#`;
    case "new": return `(new ${e.className})`;
    case "unary": return `(${e.op} ${sexp(e.operand)})`;
    case "binary": return `(${e.op} ${sexp(e.left)} ${sexp(e.right)})`;
    case "member": return `(. ${e.target ? sexp(e.target) : "<with>"} ${e.name})`;
    case "call": return `(call ${sexp(e.target)}${e.args.map((a) => ` ${a.name ? `${a.name}:=` : ""}${a.value ? sexp(a.value) : "_"}`).join("")})`;
  }
}
const expr = (src: string): string => sexp(parseExpression(src));

describe("lexer", () => {
  it("joins a continued line and drops the underscore", () => {
    const toks = lex("x = 1 + _\n  2");
    expect(toks.filter((t) => t.kind === "eol").length).toBe(0);
    expect(toks.map((t) => t.text).join(" ")).toBe("x = 1 + 2 ");
  });

  it("treats a comment as running to the end of the line", () => {
    const toks = lex("x = 1 ' this = 2\ny = 3");
    expect(toks.filter((t) => t.kind === "number").map((t) => t.value)).toEqual([1, 3]);
  });

  it("does not let a comment continue over an underscore", () => {
    // The underscore is inside the comment, so the next line is real code.
    const toks = lex("' comment _\ny = 3");
    expect(toks.filter((t) => t.kind === "number").map((t) => t.value)).toEqual([3]);
  });

  it("reads a doubled quote as one character", () => {
    expect(lex('s = "say ""hi"""')[2]!.value).toBe('say "hi"');
  });

  it("reads hex and octal, which start with the concatenation operator", () => {
    expect(lex("&HFF")[0]!.value).toBe(255);
    expect(lex("&O17")[0]!.value).toBe(15);
    // A bare & is still concatenation.
    expect(lex('a & b')[1]!.text).toBe("&");
  });

  it("drops a literal's type character", () => {
    expect(lex("1000&")[0]!.value).toBe(1000);
    expect(lex("1.5!")[0]!.value).toBe(1.5);
  });

  it("keeps line breaks, since they end statements", () => {
    expect(lex("a\nb").filter((t) => t.kind === "eol").length).toBe(1);
  });

  it("reports the line of an unterminated string", () => {
    expect(() => lex('a = 1\nb = "oops\n')).toThrow(VbaSyntaxError);
    try { lex('a = 1\nb = "oops\n'); } catch (e) { expect((e as VbaSyntaxError).line).toBe(2); }
  });
});

describe("expressions", () => {
  it("applies VBA's precedence", () => {
    expect(expr("1 + 2 * 3")).toBe("(+ 1 (* 2 3))");
    expect(expr("1 & 2 + 3")).toBe("(& 1 (+ 2 3))"); // & binds looser than +
    expect(expr("a Mod b * c")).toBe("(Mod a (* b c))"); // Mod binds looser than *
    expect(expr("a = b And c = d")).toBe("(And (= a b) (= c d))");
    expect(expr("Not a And b")).toBe("(And (Not a) b)");
  });

  it("makes ^ bind tightest and associate rightwards", () => {
    expect(expr("2 ^ 3 ^ 2")).toBe("(^ 2 (^ 3 2))");
    expect(expr("-2 ^ 2")).toBe("(- (^ 2 2))"); // the sign applies to the power, as in Excel
  });

  it("chains member access and calls", () => {
    expect(expr("a.b.c")).toBe("(. (. a b) c)");
    expect(expr("Range(\"A1\").Value")).toBe('(. (call Range "A1") Value)');
    expect(expr("a.b(1).c")).toBe("(. (call (. a b) 1) c)");
  });

  it("parses a member with no target, which belongs to the enclosing With", () => {
    expect(expr(".Value")).toBe("(. <with> Value)");
  });

  it("parses named and omitted arguments", () => {
    expect(expr("f(a:=1, b:=2)")).toBe("(call f a:=1 b:=2)");
    expect(expr("f(1, , 3)")).toBe("(call f 1 _ 3)");
  });

  it("parses Is and Like, which are words rather than symbols", () => {
    expect(expr("a Is Nothing")).toBe("(Is a Nothing)");
    expect(expr('s Like "a*"')).toBe('(Like s "a*")');
  });

  it("allows a keyword as a property name", () => {
    // Value, Count and Error are all real Excel members and all lex as keywords or idents.
    expect(expr("a.Error")).toBe("(. a Error)");
  });
});

/** Parse one procedure body, for statement-shape tests. */
const body = (src: string): Stmt[] => parseModule(`Sub T()\n${src}\nEnd Sub`).procedures[0]!.body;

describe("statements", () => {
  it("parses a multi-line If with ElseIf and Else", () => {
    const s = body("If a Then\nx = 1\nElseIf b Then\nx = 2\nElse\nx = 3\nEnd If")[0]!;
    expect(s.type).toBe("if");
    if (s.type !== "if") return;
    expect(s.branches.length).toBe(2);
    expect(s.else?.length).toBe(1);
  });

  it("parses a single-line If, which has no End If", () => {
    const stmts = body('If a Then x = 1 Else x = 2\ny = 9');
    const s = stmts[0]!;
    expect(s.type).toBe("if");
    if (s.type !== "if") return;
    expect(s.branches[0]!.body.length).toBe(1);
    expect(s.else?.length).toBe(1);
    // The line break ends the If: what follows is a separate statement, not part of the Else.
    expect(stmts.length).toBe(2);
    expect(stmts[1]!.type).toBe("assign");
  });

  it("parses For with and without Step, and For Each", () => {
    const forStep = body("For i = 1 To 10 Step 2\nx = i\nNext i")[0]!;
    expect(forStep.type).toBe("for");
    if (forStep.type === "for") expect(forStep.step).toBeTruthy();
    const each = body("For Each c In rng\nx = c\nNext c")[0]!;
    expect(each.type).toBe("forEach");
    if (each.type === "forEach") expect(each.varName).toBe("c");
  });

  it("tells a Do...Loop While from a Do While...Loop", () => {
    const atEnd = body("Do\nx = 1\nLoop While a")[0]!;
    const atStart = body("Do While a\nx = 1\nLoop")[0]!;
    if (atEnd.type === "doLoop") expect(atEnd.test?.atEnd).toBe(true);
    if (atStart.type === "doLoop") expect(atStart.test?.atEnd).toBe(false);
  });

  it("parses Select Case with values, ranges and Is comparisons", () => {
    const s = body("Select Case n\nCase 1, 2\nx = 1\nCase 3 To 5\nx = 2\nCase Is > 9\nx = 3\nCase Else\nx = 4\nEnd Select")[0]!;
    expect(s.type).toBe("select");
    if (s.type !== "select") return;
    expect(s.cases.length).toBe(3);
    expect(s.cases[0]!.tests.length).toBe(2);
    expect(s.cases[1]!.tests[0]!.kind).toBe("range");
    expect(s.cases[2]!.tests[0]!.kind).toBe("compare");
    expect(s.elseBody?.length).toBe(1);
  });

  it("parses With, and the dotted members inside it", () => {
    const s = body('With Range("A1")\n.Value = 1\n.Font.Bold = True\nEnd With')[0]!;
    expect(s.type).toBe("with");
    if (s.type !== "with") return;
    expect(s.body.length).toBe(2);
    const first = s.body[0]!;
    expect(first.type).toBe("assign");
    if (first.type !== "assign") return;
    expect(sexp(first.target)).toBe("(. <with> Value)");
  });

  it("separates statements written on one line with colons", () => {
    expect(body("x = 1: y = 2: z = 3").length).toBe(3);
  });

  it("tells Set from a plain assignment", () => {
    expect(body('Set r = Range("A1")')[0]!.type).toBe("set");
    expect(body("x = 1")[0]!.type).toBe("assign");
  });

  it("reads = as assignment at statement level and as comparison inside an expression", () => {
    // The same token means two things, and parsing the target with the full expression parser
    // turned every assignment into a discarded comparison.
    const assign = body("x = 1")[0]!;
    expect(assign.type).toBe("assign");
    if (assign.type === "assign") {
      expect(sexp(assign.target)).toBe("x");
      expect(sexp(assign.value)).toBe("1");
    }
    const cond = body("If a = b Then\nx = 1\nEnd If")[0]!;
    if (cond.type === "if") expect(sexp(cond.branches[0]!.cond)).toBe("(= a b)");
    // An assignment to a member, which is the common case in Excel macros.
    const member = body('Range("A1").Value = 42')[0]!;
    expect(member.type).toBe("assign");
    if (member.type === "assign") expect(sexp(member.target)).toBe('(. (call Range "A1") Value)');
  });

  it("parses a call written without parentheses", () => {
    const s = body('MsgBox "hi", 1')[0]!;
    expect(s.type).toBe("callStmt");
    if (s.type === "callStmt") expect(sexp(s.expr)).toBe('(call MsgBox "hi" 1)');
  });

  it("parses Dim, including arrays and types", () => {
    const s = body("Dim a As Long, b(1 To 10) As String, c()")[0]!;
    expect(s.type).toBe("dim");
    if (s.type !== "dim") return;
    expect(s.vars.map((v) => v.name)).toEqual(["a", "b", "c"]);
    expect(s.vars[0]!.asType).toBe("Long");
    expect(s.vars[1]!.dims?.[0]?.lower).toBeTruthy();
    expect(s.vars[2]!.isArray).toBe(true);
  });

  it("parses error handling and labels", () => {
    const stmts = body("On Error Resume Next\nOn Error GoTo Oops\nGoTo Done\nOops:\nDone:");
    expect(stmts[0]!.type).toBe("onError");
    if (stmts[0]!.type === "onError") expect(stmts[0]!.mode).toBe("resumeNext");
    if (stmts[1]!.type === "onError") expect(stmts[1]!.label).toBe("Oops");
    expect(stmts[2]!.type).toBe("goto");
    expect(stmts[3]!.type).toBe("label");
  });

  it("keeps what it cannot model, so it can be refused by name later", () => {
    const s = body('Attribute VB_Name = "Module1"')[0]!;
    expect(s.type).toBe("unsupported");
  });
});

describe("procedures and modules", () => {
  it("parses a Sub with parameters", () => {
    const m = parseModule("Public Sub Go(ByVal n As Long, Optional s As String = \"x\")\nEnd Sub");
    const p = m.procedures[0]!;
    expect(p.kind).toBe("sub");
    expect(p.isPublic).toBe(true);
    expect(p.params.map((x) => x.name)).toEqual(["n", "s"]);
    expect(p.params[0]!.byVal).toBe(true);
    expect(p.params[1]!.optional).toBe(true);
    expect(p.params[1]!.default).toBeTruthy();
  });

  it("parses a Function with a return type, and Private", () => {
    const p = parseModule("Private Function F(a) As Double\nF = a * 2\nEnd Function").procedures[0]!;
    expect(p.kind).toBe("function");
    expect(p.asType).toBe("Double");
    expect(p.isPublic).toBe(false);
  });

  it("parses the three Property forms", () => {
    const m = parseModule("Property Get X() As Long\nEnd Property\nProperty Let X(v As Long)\nEnd Property\nProperty Set Y(v As Object)\nEnd Property");
    expect(m.procedures.map((p) => p.kind)).toEqual(["propertyGet", "propertyLet", "propertySet"]);
  });

  it("separates module-level declarations from procedures", () => {
    const m = parseModule("Option Explicit\nDim total As Long\nConst MAX = 10\nSub A()\nEnd Sub");
    expect(m.options).toEqual(["Explicit"]);
    expect(m.declarations.map((d) => d.type)).toEqual(["dim", "const"]);
    expect(m.procedures.map((p) => p.name)).toEqual(["A"]);
  });

  it("reports the line of a syntax error", () => {
    try {
      parseModule("Sub A()\nx = = 1\nEnd Sub");
      expect.unreachable("should have thrown");
    } catch (e) {
      expect(e).toBeInstanceOf(VbaSyntaxError);
      expect((e as VbaSyntaxError).line).toBe(2);
    }
  });
});

describe("parsing the real macros in the fixtures", () => {
  const modulesOf = (file: string) =>
    readVbaProject(unzipSync(new Uint8Array(readFileSync(`src/fixtures/${file}`)))["xl/vbaProject.bin"]!)!.modules;

  it("parses every module of both fixtures", () => {
    for (const file of ["macros-cp950.xlsm", "macros-cp1252.xlsm"]) {
      for (const mod of modulesOf(file)) {
        // Source written by Excel, complete with its Attribute preamble.
        expect(() => parseModule(mod.source, mod.name), `${file} / ${mod.name}`).not.toThrow();
      }
    }
  });

  it("finds the procedures Excel wrote", () => {
    const mod = modulesOf("macros-cp950.xlsm").find((m) => m.name === "Module1")!;
    const parsed = parseModule(mod.source, mod.name);
    expect(parsed.procedures.map((p) => p.name)).toEqual(["Button1_Click"]);
    expect(parsed.procedures[0]!.body.length).toBeGreaterThan(0);
  });

  it("parses the German module's two handlers", () => {
    const mod = modulesOf("macros-cp1252.xlsm").find((m) => m.name === "Modul1")!;
    expect(parseModule(mod.source, mod.name).procedures.map((p) => p.name)).toEqual(["Plus1_Klicken", "Minus2_Klicken"]);
  });
});
