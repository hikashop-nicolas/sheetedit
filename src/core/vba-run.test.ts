import { readFileSync } from "node:fs";
import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { parseModule } from "./vba-parse";
import { readVbaProject } from "./vba";
import { runMacro, VbaInterpreter } from "./vba-run";
import { EMPTY, NOTHING, NULL, VbaArray, VbaError, type VbaObject, type VbaValue } from "./vba-value";

// Stage 2 of _plans/VBA_PLAN.md. These pin the coercion corners that make VBA VBA: True is -1,
// Empty and Null are different things, + is not &, and the operators that look arithmetic are not.
// Getting any of them subtly wrong produces answers that look right, which is worse than stopping.

/** Run a Sub body and read back what it left in `out`. */
function evalExpr(src: string, globals?: Map<string, VbaValue>): VbaValue {
  const m = parseModule(`Function T()\nT = ${src}\nEnd Function`);
  return new VbaInterpreter(m, { globals }).run("T").value ?? EMPTY;
}

/** Run statements and return everything Debug.Print / MsgBox emitted. */
function messages(src: string, globals?: Map<string, VbaValue>): string[] {
  const m = parseModule(`Sub T()\n${src}\nEnd Sub`);
  return new VbaInterpreter(m, { globals }).run("T").messages;
}

/** Run statements and read one variable back out of the procedure. */
function valueOf(src: string, varName = "r"): VbaValue {
  const m = parseModule(`Function T()\n${src}\nT = ${varName}\nEnd Function`);
  return new VbaInterpreter(m).run("T").value ?? EMPTY;
}

describe("VBA coercion", () => {
  it("makes True -1 in arithmetic but keeps it Boolean in logic", () => {
    expect(evalExpr("True + 0")).toBe(-1);
    expect(evalExpr("True * 5")).toBe(-5);
    expect(evalExpr("True And False")).toBe(false);
    expect(evalExpr("CStr(True)")).toBe("True");
  });

  it("treats Empty as 0 and \"\" at once", () => {
    expect(valueOf("Dim x\nDim r\nr = x + 1")).toBe(1);
    expect(valueOf('Dim x\nDim r\nr = (x = 0)')).toBe(true);
    expect(valueOf('Dim x\nDim r\nr = (x = "")')).toBe(true);
    expect(valueOf("Dim x\nDim r\nr = IsEmpty(x)")).toBe(true);
  });

  it("propagates Null through arithmetic but not through &", () => {
    const g = new Map<string, VbaValue>([["N", NULL]]);
    expect(evalExpr("N + 1", g)).toBe(NULL);
    expect(evalExpr("N * 0", g)).toBe(NULL);
    expect(evalExpr("N = 1", g)).toBe(NULL);
    // Concatenation is the exception: Null contributes an empty string.
    expect(evalExpr('"a" & N & "b"', g)).toBe("ab");
    expect(evalExpr("IsNull(N)", g)).toBe(true);
    expect(evalExpr("IsEmpty(N)", g)).toBe(false);
  });

  it("adds with + only when a side is not a string, and always concatenates with &", () => {
    expect(evalExpr('"1" + "2"')).toBe("12");
    expect(evalExpr('"1" + 2')).toBe(3);
    expect(evalExpr('1 & 2')).toBe("12");
    expect(evalExpr('"a" & 1')).toBe("a1");
  });

  it("compares numbers numerically and anything with a string as text", () => {
    expect(evalExpr("10 > 9")).toBe(true);
    expect(evalExpr('"10" > "9"')).toBe(false); // text order, as in Excel
    expect(evalExpr('"abc" < "abd"')).toBe(true);
  });

  it("truncates both operands for integer division and Mod", () => {
    expect(evalExpr("7 \\ 2")).toBe(3);
    expect(evalExpr("-7 \\ 2")).toBe(-3);
    expect(evalExpr("7.9 \\ 2.9")).toBe(3);
    expect(evalExpr("7 Mod 3")).toBe(1);
    expect(evalExpr("-7 Mod 3")).toBe(-1); // keeps the dividend's sign
  });

  it("rounds half to even, which is what VBA's Round does", () => {
    expect(evalExpr("Round(2.5)")).toBe(2);
    expect(evalExpr("Round(3.5)")).toBe(4);
    expect(evalExpr("Round(-2.5)")).toBe(-2);
    // 0.125 is exactly representable, so this is a real half and rounds to the even neighbour.
    expect(evalExpr("Round(0.125, 2)")).toBe(0.12);
  });

  it("floors with Int and truncates with Fix, which differ only for negatives", () => {
    expect(evalExpr("Int(-2.5)")).toBe(-3);
    expect(evalExpr("Fix(-2.5)")).toBe(-2);
  });

  it("stops on division by zero rather than producing Infinity", () => {
    expect(() => evalExpr("1 / 0")).toThrow(/division by zero/);
    expect(() => evalExpr("1 \\ 0")).toThrow(/division by zero/);
    expect(() => evalExpr("1 Mod 0")).toThrow(/division by zero/);
  });

  it("implements Like with VBA's wildcards", () => {
    expect(evalExpr('"file.txt" Like "*.txt"')).toBe(true);
    expect(evalExpr('"a1" Like "a#"')).toBe(true);
    expect(evalExpr('"ab" Like "a?"')).toBe(true);
    expect(evalExpr('"ax" Like "a[!bc]"')).toBe(true);
    expect(evalExpr('"ab" Like "a[!bc]"')).toBe(false);
  });
});

describe("VBA control flow", () => {
  it("runs If / ElseIf / Else", () => {
    expect(valueOf("Dim r\nIf 1 > 2 Then\nr = 1\nElseIf 2 > 1 Then\nr = 2\nElse\nr = 3\nEnd If")).toBe(2);
    expect(valueOf("Dim r\nIf 1 > 2 Then r = 1 Else r = 9")).toBe(9);
  });

  it("counts a For loop inclusively and honours Step", () => {
    expect(valueOf("Dim r, i\nr = 0\nFor i = 1 To 5\nr = r + i\nNext i")).toBe(15);
    expect(valueOf("Dim r, i\nr = 0\nFor i = 10 To 1 Step -2\nr = r + 1\nNext i")).toBe(5);
    // A For whose start is already past its end never runs.
    expect(valueOf("Dim r, i\nr = 0\nFor i = 5 To 1\nr = r + 1\nNext i")).toBe(0);
  });

  it("refuses a For with Step 0 instead of hanging", () => {
    expect(() => valueOf("Dim i\nFor i = 1 To 5 Step 0\nNext i")).toThrow(/never end/);
  });

  it("leaves the loop with Exit For and the procedure with Exit Sub", () => {
    expect(valueOf("Dim r, i\nr = 0\nFor i = 1 To 10\nIf i = 4 Then Exit For\nr = i\nNext i")).toBe(3);
    expect(messages("Debug.Print \"a\"\nExit Sub\nDebug.Print \"b\"")).toEqual(["a"]);
  });

  it("runs Do Until at the end at least once", () => {
    expect(valueOf("Dim r\nr = 0\nDo\nr = r + 1\nLoop Until r >= 3")).toBe(3);
    // The same test at the top never runs when it is already satisfied.
    expect(valueOf("Dim r\nr = 5\nDo Until r >= 3\nr = r + 1\nLoop")).toBe(5);
  });

  it("runs While / Wend", () => {
    expect(valueOf("Dim r\nr = 0\nWhile r < 4\nr = r + 1\nWend")).toBe(4);
  });

  it("matches Select Case values, ranges and Is comparisons", () => {
    const sel = (n: number): VbaValue => valueOf(
      `Dim r\nSelect Case ${n}\nCase 1, 2\nr = "low"\nCase 3 To 5\nr = "mid"\nCase Is > 5\nr = "high"\nCase Else\nr = "other"\nEnd Select`,
    );
    expect(sel(2)).toBe("low");
    expect(sel(4)).toBe("mid");
    expect(sel(9)).toBe("high");
    expect(sel(0)).toBe("other");
  });

  it("jumps with GoTo to a label in the same block", () => {
    expect(valueOf("Dim r\nr = 0\nGoTo Skip\nr = 1\nSkip:\nr = r + 10")).toBe(10);
  });
});

describe("VBA procedures", () => {
  it("returns a Function's value through its own name", () => {
    const m = parseModule("Function Double1(n)\nDouble1 = n * 2\nEnd Function\nFunction T()\nT = Double1(21)\nEnd Function");
    expect(new VbaInterpreter(m).run("T").value).toBe(42);
  });

  it("passes ByRef by default and ByVal when asked", () => {
    const byRef = "Sub Bump(n)\nn = n + 1\nEnd Sub\nFunction T()\nDim x\nx = 1\nBump x\nT = x\nEnd Function";
    expect(new VbaInterpreter(parseModule(byRef)).run("T").value).toBe(2);
    const byVal = "Sub Bump(ByVal n)\nn = n + 1\nEnd Sub\nFunction T()\nDim x\nx = 1\nBump x\nT = x\nEnd Function";
    expect(new VbaInterpreter(parseModule(byVal)).run("T").value).toBe(1);
  });

  it("fills an omitted Optional parameter from its default", () => {
    const m = parseModule('Function Greet(Optional who As String = "world")\nGreet = "hi " & who\nEnd Function\n'
      + 'Function T()\nT = Greet()\nEnd Function');
    expect(new VbaInterpreter(m).run("T").value).toBe("hi world");
  });

  it("lists the Subs a user can run and leaves out the rest", () => {
    const m = parseModule([
      "Public Sub Runnable()", "End Sub",
      "Private Sub Hidden()", "End Sub",
      "Sub NeedsArgs(n As Long)", "End Sub",
      "Function NotASub()", "End Function",
    ].join("\n"));
    expect(new VbaInterpreter(m).runnableSubs()).toEqual(["Runnable"]);
  });
});

describe("VBA arrays", () => {
  it("indexes a Dim'd array and defaults its lower bound to 0", () => {
    expect(valueOf("Dim a(3)\nDim r\na(2) = 7\nr = a(2)")).toBe(7);
    expect(valueOf("Dim a(3)\nDim r\nr = LBound(a) & \":\" & UBound(a)")).toBe("0:3");
  });

  it("honours an explicit lower bound", () => {
    expect(valueOf("Dim a(1 To 3)\nDim r\nr = LBound(a)")).toBe(1);
  });

  it("stops on an out-of-bounds index rather than growing the array", () => {
    expect(() => valueOf("Dim a(2)\na(5) = 1")).toThrow(/subscript out of range/);
  });

  it("keeps what fits on ReDim Preserve and drops the rest", () => {
    expect(valueOf("Dim a()\nDim r\nReDim a(2)\na(1) = 5\nReDim Preserve a(4)\nr = a(1)")).toBe(5);
    expect(valueOf("Dim a()\nDim r\nReDim a(4)\na(4) = 5\nReDim Preserve a(2)\nr = a(2)")).toBe(EMPTY);
  });

  it("walks an array with For Each", () => {
    expect(valueOf('Dim r\nDim v\nr = ""\nFor Each v In Array("a", "b", "c")\nr = r & v\nNext v')).toBe("abc");
  });

  it("splits and joins", () => {
    expect(evalExpr('Join(Split("a,b,c", ","), "-")')).toBe("a-b-c");
  });
});

describe("VBA string builtins", () => {
  it("uses 1-based positions, as VBA does throughout", () => {
    expect(evalExpr('Mid("abcdef", 2, 3)')).toBe("bcd");
    expect(evalExpr('Mid("abcdef", 4)')).toBe("def");
    expect(evalExpr('InStr("hello", "l")')).toBe(3);
    expect(evalExpr('InStr(4, "hello", "l")')).toBe(4);
    expect(evalExpr('InStr("hello", "z")')).toBe(0);
  });

  it("gives Left and Right an empty string rather than an error at 0", () => {
    expect(evalExpr('Left("abc", 0)')).toBe("");
    expect(evalExpr('Right("abc", 0)')).toBe("");
    expect(evalExpr('Right("abc", 2)')).toBe("bc");
  });

  it("puts a leading space on a positive Str, which trips up naive concatenation", () => {
    expect(evalExpr("Str(5)")).toBe(" 5");
    expect(evalExpr("Str(-5)")).toBe("-5");
    expect(evalExpr("CStr(5)")).toBe("5");
  });

  it("reads the numeric prefix with Val and gives 0 when there is none", () => {
    expect(evalExpr('Val("12abc")')).toBe(12);
    expect(evalExpr('Val("abc")')).toBe(0);
    expect(evalExpr('Val("3.5e2")')).toBe(350);
  });

  it("names types the way TypeName does", () => {
    expect(evalExpr("TypeName(1)")).toBe("Long");
    expect(evalExpr("TypeName(1.5)")).toBe("Double");
    expect(evalExpr('TypeName("a")')).toBe("String");
    expect(evalExpr("TypeName(True)")).toBe("Boolean");
    expect(valueOf("Dim x\nDim r\nr = TypeName(x)")).toBe("Empty");
  });
});

describe("On Error", () => {
  it("carries on past a failing statement with Resume Next", () => {
    expect(valueOf("Dim r\nOn Error Resume Next\nr = 1 / 0\nr = 9")).toBe(9);
  });

  it("jumps to the handler with On Error GoTo", () => {
    expect(valueOf('Dim r\nOn Error GoTo Oops\nr = 1 / 0\nr = "no"\nExit Function\nOops:\nr = "caught"')).toBe("caught");
  });

  it("stops again after On Error GoTo 0", () => {
    expect(() => valueOf("Dim r\nOn Error Resume Next\nOn Error GoTo 0\nr = 1 / 0")).toThrow(/division by zero/);
  });

  it("names the line the error happened on", () => {
    let caught: VbaError | undefined;
    try {
      valueOf("Dim r\nr = 1\nr = 1 / 0");
    } catch (e) { caught = e as VbaError; }
    // Line 1 is the Function header the helper wraps the body in.
    expect(caught?.line).toBe(4);
  });
});

describe("Err", () => {
  it("reports the number and description of the last error", () => {
    expect(valueOf("Dim r\nOn Error Resume Next\nr = 1 / 0\nr = Err.Number")).toBe(11);
    expect(valueOf("Dim r\nOn Error Resume Next\nr = 1 / 0\nr = Err.Description")).toMatch(/division by zero/);
  });

  it("resets on Err.Clear", () => {
    expect(valueOf("Dim r\nOn Error Resume Next\nr = 1 / 0\nErr.Clear\nr = Err.Number")).toBe(0);
  });

  it("raises an error the handler can then see", () => {
    expect(valueOf('Dim r\nOn Error Resume Next\nErr.Raise 513, , "custom"\nr = Err.Number')).toBe(513);
  });
});

describe("Debug.Print", () => {
  it("collects one line per call", () => {
    expect(messages('Debug.Print "a"\nDebug.Print 1 + 1')).toEqual(["a", "2"]);
  });

  it("prints Null as the word rather than failing on it", () => {
    const g = new Map<string, VbaValue>([["N", NULL]]);
    expect(messages("Debug.Print N", g)).toEqual(["Null"]);
  });
});

describe("the step budget", () => {
  it("stops a runaway loop instead of hanging the tab", () => {
    // An empty body executes no statement, so the budget has to tick per iteration, not per
    // statement, or there is nothing to stop this at all.
    const m = parseModule("Sub T()\nDo\nLoop\nEnd Sub");
    expect(() => new VbaInterpreter(m, { maxSteps: 1000 }).run("T")).toThrow(/ran too long/);
  });

  it("stops an empty For and an empty While too", () => {
    const empty = (src: string): void => {
      const m = parseModule(`Sub T()\n${src}\nEnd Sub`);
      new VbaInterpreter(m, { maxSteps: 500 }).run("T");
    };
    expect(() => empty("Dim i\nFor i = 1 To 100000\nNext i")).toThrow(/ran too long/);
    expect(() => empty("Dim i\ni = 0\nWhile i = 0\nWend")).toThrow(/ran too long/);
  });
});

describe("refusing rather than approximating", () => {
  it("stops on a call to something it does not model, naming it", () => {
    expect(() => messages("Shell \"rm -rf /\"")).toThrow(/Shell.*does not provide/);
  });

  it("stops on an undefined name in an expression", () => {
    expect(() => evalExpr("SomeUnknownThing + 1")).toThrow(/SomeUnknownThing is not defined/);
  });

  it("stops on New, since there are no classes to build", () => {
    expect(() => valueOf("Dim r\nSet r = New Collection")).toThrow(/New Collection is not supported/);
  });

  it("stops on InputBox, since nobody is there to answer it", () => {
    expect(() => valueOf('Dim r\nr = InputBox("name?")')).toThrow(/needs someone to answer/);
  });

  it("collects MsgBox text instead of blocking on a dialog", () => {
    expect(messages('MsgBox "done"')).toEqual(["done"]);
    // It still answers vbOK so `If MsgBox(...) = vbOK` takes the branch its author meant.
    expect(valueOf('Dim r\nIf MsgBox("go?", vbYesNo) = vbOK Then r = "yes"')).toBe("yes");
  });

  it("passes Attribute and Option lines over without complaint", () => {
    const m = parseModule('Option Explicit\nSub T()\nAttribute T.VB_Description = "x"\nEnd Sub');
    expect(() => new VbaInterpreter(m).run("T")).not.toThrow();
  });
});

// --- the host object model seam ------------------------------------------------
// Stage 3 supplies the real Excel objects. The interpreter only ever talks to this interface, so a
// stub here proves the language and the object model are genuinely separable.

class StubCell implements VbaObject {
  readonly typeName = "Range";
  constructor(private readonly store: Map<string, VbaValue>, private readonly addr: string) {}
  get(name: string): VbaValue {
    if (name === "" || name.toLowerCase() === "value") return this.store.get(this.addr) ?? EMPTY;
    if (name.toLowerCase() === "address") return this.addr;
    throw new VbaError(`Range has no ${name}`, 438);
  }
  set(name: string, _args: VbaValue[], value: VbaValue): void {
    if (name === "" || name.toLowerCase() === "value") this.store.set(this.addr, value);
    else throw new VbaError(`Range has no ${name}`, 438);
  }
  defaultValue(): VbaValue { return this.store.get(this.addr) ?? EMPTY; }
}

class StubSheet implements VbaObject {
  readonly typeName = "Worksheet";
  readonly store = new Map<string, VbaValue>();
  get(name: string, args: VbaValue[]): VbaValue {
    if (name.toLowerCase() === "range") return new StubCell(this.store, String(args[0]));
    if (name.toLowerCase() === "name") return "Sheet1";
    throw new VbaError(`Worksheet has no ${name}`, 438);
  }
}

function withSheet(src: string): { sheet: StubSheet; result: VbaValue } {
  const sheet = new StubSheet();
  const globals = new Map<string, VbaValue>([["ACTIVESHEET", sheet]]);
  const m = parseModule(`Function T()\n${src}\nEnd Function`);
  return { sheet, result: new VbaInterpreter(m, { globals }).run("T").value ?? EMPTY };
}

describe("the host object seam", () => {
  it("reads and writes a host property", () => {
    const { sheet } = withSheet('ActiveSheet.Range("A1").Value = 42');
    expect(sheet.store.get("A1")).toBe(42);
  });

  it("takes an object's default value on a plain assignment", () => {
    const { result } = withSheet('ActiveSheet.Range("A1").Value = 7\nT = ActiveSheet.Range("A1")');
    expect(result).toBe(7);
  });

  it("resolves a leading dot through With", () => {
    const { sheet } = withSheet('With ActiveSheet\n.Range("B2").Value = "x"\nEnd With');
    expect(sheet.store.get("B2")).toBe("x");
  });

  it("nests With blocks and pops back to the outer subject", () => {
    const { sheet } = withSheet([
      "With ActiveSheet",
      '  With .Range("C3")',
      '    .Value = 1',
      "  End With",
      '  .Range("C4").Value = 2',
      "End With",
    ].join("\n"));
    expect(sheet.store.get("C3")).toBe(1);
    expect(sheet.store.get("C4")).toBe(2);
  });

  it("refuses a leading dot with no enclosing With", () => {
    expect(() => withSheet('.Range("A1").Value = 1')).toThrow(/needs an enclosing With/);
  });

  it("lets a host error surface with its own message", () => {
    expect(() => withSheet("T = ActiveSheet.NoSuchThing")).toThrow(/Worksheet has no NoSuchThing/);
  });
});

describe("Nothing and Is", () => {
  it("compares object references with Is", () => {
    const sheet = new StubSheet();
    const g = new Map<string, VbaValue>([["S", sheet], ["S2", sheet], ["OTHER", new StubSheet()]]);
    expect(evalExpr("S Is S2", g)).toBe(true);
    expect(evalExpr("S Is Other", g)).toBe(false);
    expect(evalExpr("Nothing Is Nothing")).toBe(true);
  });

  it("stops when Nothing is used as a value", () => {
    const g = new Map<string, VbaValue>([["N", NOTHING]]);
    expect(() => evalExpr("N + 1", g)).toThrow(/object variable not set/);
  });
});

describe("runMacro", () => {
  it("runs a whole module end to end", () => {
    const src = [
      "Option Explicit",
      "",
      "Sub Report()",
      "  Dim i As Long, total As Long",
      "  For i = 1 To 10",
      "    If i Mod 2 = 0 Then total = total + i",
      "  Next i",
      '  Debug.Print "even total: " & total',
      "End Sub",
    ].join("\n");
    expect(runMacro(parseModule(src), "Report").messages).toEqual(["even total: 30"]);
  });

  it("says so when the procedure does not exist", () => {
    expect(() => runMacro(parseModule("Sub A()\nEnd Sub"), "B")).toThrow(/no procedure called B/);
  });
});

describe("named arguments", () => {
  it("sends Name:=value to the parameter it names, wherever it sits", () => {
    const m = parseModule('Function F(a, b, c)\nF = a & "/" & b & "/" & c\nEnd Function\n'
      + 'Function T()\nT = F(c:="z", a:="x", b:="y")\nEnd Function');
    expect(new VbaInterpreter(m).run("T").value).toBe("x/y/z");
  });

  it("stops on a name no parameter answers to", () => {
    const m = parseModule("Sub F(a)\nEnd Sub\nSub T()\nF nope:=1\nEnd Sub");
    expect(() => new VbaInterpreter(m).run("T")).toThrow(/no argument called nope/);
  });

  it("hands the names to a host method rather than dropping them", () => {
    const seen: (string | undefined)[][] = [];
    const obj: VbaObject = {
      typeName: "Sheet",
      get(_n, _a, names) { seen.push(names ?? []); return EMPTY; },
    };
    const m = parseModule('Sub T()\nS.Protect Password:="p"\nEnd Sub');
    new VbaInterpreter(m, { globals: new Map<string, VbaValue>([["S", obj]]) }).run("T");
    expect(seen).toEqual([["Password"]]);
  });
});

describe("real macros from the fixtures", () => {
  const modulesOf = (file: string): { name: string; source: string }[] =>
    readVbaProject(unzipSync(new Uint8Array(readFileSync(`src/fixtures/${file}`)))["xl/vbaProject.bin"]!)!.modules;

  it("stops on the Excel object it has no model for, naming it", () => {
    // Stage 3 supplies Worksheets and Excel's xl* constants. Until then this must refuse by name
    // rather than do nothing and report success, which would leave the user thinking it ran.
    const mod = modulesOf("macros-cp1252.xlsm").find((m) => m.name === "Modul1")!;
    expect(() => runMacro(parseModule(mod.source, mod.name), "Plus1_Klicken"))
      .toThrow(/xlCellTypeBlanks is not defined/);
  });

  it("runs one once the host provides what it asks for", () => {
    // A stub Worksheets is enough to prove the language half is done: the macro's own control flow,
    // its named Password argument and its chained members all resolve.
    const calls: string[] = [];
    const range: VbaObject = {
      typeName: "Range",
      get: (n, a) => { calls.push(`Range.${n}(${a.map(String).join(",")})`); return range; },
      set: (n, _a, v) => { calls.push(`Range.${n}=${String(v)}`); },
    };
    const sheet: VbaObject = {
      typeName: "Worksheet",
      get: (n, a, names) => {
        calls.push(`${n}(${a.map(String).join(",")}${names ? ` names=${names.join()}` : ""})`);
        return n.toLowerCase() === "range" ? range : EMPTY;
      },
    };
    const worksheets: VbaObject = { typeName: "Sheets", get: () => sheet };
    const mod = modulesOf("macros-cp1252.xlsm").find((m) => m.name === "Modul1")!;
    runMacro(parseModule(mod.source, mod.name), "Plus1_Klicken", {
      globals: new Map<string, VbaValue>([["WORKSHEETS", worksheets], ["XLCELLTYPEBLANKS", 4]]),
    });
    expect(calls).toContain("Range(A5:A44)");
    expect(calls).toContain("Range.SpecialCells(4)");   // the xlCellTypeBlanks the host supplied
    expect(calls).toContain("Range.EntireRow()");       // a property read, reaching get with no args
    expect(calls).toContain("Range.Hidden=false");
    // Its trailing `Protect Password:=...` arrives as a named argument, value left out of the test.
    expect(calls.some((c) => c.startsWith("Protect(") && c.includes("names=Password"))).toBe(true);
  });
});

describe("VbaArray bounds", () => {
  it("reports Empty for a slot never written", () => {
    const a = VbaArray.ofDims([{ lower: 0, upper: 2 }]);
    expect(a.get([1])).toBe(EMPTY);
    expect(a.inBounds([3])).toBe(false);
    expect(a.values()).toEqual([EMPTY, EMPTY, EMPTY]);
  });
});
