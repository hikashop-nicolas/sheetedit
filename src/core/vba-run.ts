import type { Arg, Expr, Module, Procedure, Stmt } from "./vba-ast";
import {
  compare, EMPTY, isObject, like, NOTHING, NULL, numToVbaString, scalarOf, toBool, toNumber, toStr,
  typeName, VbaArray, VbaError, type VbaObject, type VbaValue,
} from "./vba-value";
import { builtins } from "./vba-builtins";

// ---------------------------------------------------------------------------
// VBA interpreter (Stage 2 of _plans/VBA_PLAN.md)
// ---------------------------------------------------------------------------
// Walks the AST. The host supplies the object model (Stage 3) through `globals`, so the language
// and Excel stay separable and the interpreter can be tested with no spreadsheet at all.
//
// The plan's rule holds here: anything unmodelled STOPS the run with a message naming it. A macro
// that half-runs leaves a workbook in a state its author never intended, and the user then saves it.

/** How a statement finished, so loops and procedures know what to do next. */
type Flow =
  | { kind: "normal" }
  | { kind: "exit"; what: "Sub" | "Function" | "Property" | "For" | "Do" }
  | { kind: "goto"; label: string };

const NORMAL: Flow = { kind: "normal" };

/** One variable slot. Boxed so ByRef arguments can share it. */
interface Slot { value: VbaValue }

/** The `Name:=` labels of a call's arguments, positionally, so a host method can honour them. */
const argNamesOf = (args: Arg[]): (string | undefined)[] | undefined =>
  args.some((a) => a.name) ? args.map((a) => a.name) : undefined;

/** Debug.Print, which real macros use constantly. Its output is collected, not written anywhere. */
class DebugObject implements VbaObject {
  readonly typeName = "Debug";
  constructor(private readonly emit: (text: string) => void) {}
  get(name: string, args: VbaValue[]): VbaValue {
    if (name.toLowerCase() !== "print") throw new VbaError(`Debug has no ${name}`, 438);
    this.emit(args.map((a) => (a === NULL ? "Null" : toStr(a))).join(" "));
    return EMPTY;
  }
}

/** The Err object an On Error handler reads. */
class ErrObject implements VbaObject {
  readonly typeName = "ErrObject";
  constructor(private readonly current: () => VbaError | null, private readonly clear: () => void) {}
  get(name: string, args: VbaValue[]): VbaValue {
    const err = this.current();
    switch (name.toLowerCase()) {
      case "number": return err?.number ?? 0;
      case "description": return err?.message ?? "";
      case "source": return err ? "sheetedit" : "";
      case "clear": this.clear(); return EMPTY;
      case "raise": throw new VbaError(args[1] !== undefined ? toStr(args[1]) : "an error was raised", toNumber(args[0] ?? 5));
      default: throw new VbaError(`Err has no ${name}`, 438);
    }
  }
}

export interface RunOptions {
  /** Names the host provides: Range, Cells, Worksheets, Application, ... */
  globals?: Map<string, VbaValue>;
  /** Called for MsgBox and Debug.Print, since a browser must not block on a dialog. */
  output?: (text: string) => void;
  /** Stops a runaway loop from hanging the tab. */
  maxSteps?: number;
}

export interface RunResult {
  /** What MsgBox and Debug.Print produced, in order. */
  messages: string[];
  /** A Function's return value, when a function was called. */
  value?: VbaValue;
}

class Scope {
  readonly vars = new Map<string, Slot>();
  constructor(readonly parent?: Scope) {}
  find(name: string): Slot | undefined {
    return this.vars.get(name.toUpperCase()) ?? this.parent?.find(name);
  }
  declare(name: string, value: VbaValue = EMPTY): Slot {
    const slot: Slot = { value };
    this.vars.set(name.toUpperCase(), slot);
    return slot;
  }
}

export class VbaInterpreter {
  private readonly moduleScope = new Scope();
  private readonly procs = new Map<string, Procedure>();
  private readonly globals: Map<string, VbaValue>;
  private readonly messages: string[] = [];
  private readonly maxSteps: number;
  private steps = 0;
  /** The stack of With subjects, so a leading dot knows what it belongs to. */
  private withStack: VbaValue[] = [];
  private onErrorMode: "stop" | "resumeNext" | "goto" = "stop";
  private onErrorLabel = "";
  private lastError: VbaError | null = null;
  /** Names VBA itself provides, before the host's own. */
  private readonly intrinsics: Map<string, VbaValue>;

  constructor(private readonly module: Module, private readonly options: RunOptions = {}) {
    this.globals = options.globals ?? new Map();
    this.maxSteps = options.maxSteps ?? 2_000_000;
    for (const p of module.procedures) this.procs.set(p.name.toUpperCase(), p);
    this.intrinsics = new Map<string, VbaValue>([
      ["DEBUG", new DebugObject((t) => { this.messages.push(t); this.options.output?.(t); })],
      ["ERR", new ErrObject(() => this.lastError, () => { this.lastError = null; })],
    ]);
  }

  /** The Subs a user can run: public, and taking no required arguments. */
  runnableSubs(): string[] {
    return this.module.procedures
      .filter((p) => p.kind === "sub" && p.isPublic && p.params.every((x) => x.optional || x.paramArray))
      .map((p) => p.name);
  }

  /** Run the module's declarations, then one procedure by name. */
  run(procName: string, args: VbaValue[] = []): RunResult {
    for (const d of this.module.declarations) this.exec(d, this.moduleScope);
    const proc = this.procs.get(procName.toUpperCase());
    if (!proc) throw new VbaError(`there is no procedure called ${procName}`, 35);
    const value = this.callProcedure(proc, args.map((v) => ({ value: v })));
    return { messages: this.messages, value };
  }

  private step(line: number): void {
    if (++this.steps > this.maxSteps) throw new VbaError("the macro ran too long and was stopped", 28, line);
  }

  // --- procedures ------------------------------------------------------------

  private callProcedure(proc: Procedure, args: Slot[]): VbaValue {
    const scope = new Scope(this.moduleScope);
    proc.params.forEach((p, i) => {
      const supplied = args[i];
      if (supplied) {
        // ByRef shares the caller's slot, so an assignment inside is visible outside. That is the
        // default in VBA and macros rely on it.
        scope.vars.set(p.name.toUpperCase(), p.byVal ? { value: supplied.value } : supplied);
      } else if (p.default) {
        scope.declare(p.name, this.eval(p.default, scope));
      } else {
        scope.declare(p.name);
      }
    });
    // A Function returns by assigning to its own name.
    if (proc.kind !== "sub") scope.declare(proc.name);
    const flow = this.execBlock(proc.body, scope);
    if (flow.kind === "goto") throw new VbaError(`there is no label called ${flow.label}`, 35, proc.line);
    return proc.kind === "sub" ? EMPTY : scope.find(proc.name)?.value ?? EMPTY;
  }

  // --- statements ------------------------------------------------------------

  private execBlock(body: Stmt[], scope: Scope): Flow {
    let i = 0;
    while (i < body.length) {
      const flow = this.execGuarded(body[i]!, scope);
      if (flow.kind === "goto") {
        // A GoTo inside this block jumps to a label in it; otherwise it propagates outwards.
        const target = body.findIndex((s) => s.type === "label" && s.name.toUpperCase() === flow.label.toUpperCase());
        if (target >= 0) { i = target + 1; continue; }
        return flow;
      }
      if (flow.kind !== "normal") return flow;
      i++;
    }
    return NORMAL;
  }

  /** Run one statement, applying whatever On Error is in force. */
  private execGuarded(stmt: Stmt, scope: Scope): Flow {
    try {
      return this.exec(stmt, scope);
    } catch (e) {
      if (!(e instanceof VbaError)) throw e;
      this.lastError = e;
      if (this.onErrorMode === "resumeNext") return NORMAL;
      if (this.onErrorMode === "goto") return { kind: "goto", label: this.onErrorLabel };
      if (e.line === undefined) e.line = stmt.line;
      throw e;
    }
  }

  private exec(stmt: Stmt, scope: Scope): Flow {
    this.step(stmt.line);
    switch (stmt.type) {
      case "label":
        return NORMAL;
      case "dim":
        for (const v of stmt.vars) {
          if (v.isArray && v.dims?.length) scope.declare(v.name, this.makeArray(v.dims, scope));
          else scope.declare(v.name);
        }
        return NORMAL;
      case "const":
        for (const v of stmt.vars) scope.declare(v.name, this.eval(v.value, scope));
        return NORMAL;
      case "redim":
        for (const v of stmt.vars) {
          const slot = scope.find(v.name) ?? scope.declare(v.name);
          const next = this.makeArray(v.dims ?? [], scope);
          // Preserve keeps whatever still fits in the new bounds.
          if (stmt.preserve && slot.value instanceof VbaArray) {
            for (const [k, val] of slot.value.data) {
              const idx = k.split(",").map(Number);
              if (next.inBounds(idx)) next.set(idx, val);
            }
          }
          slot.value = next;
        }
        return NORMAL;
      case "assign":
        this.assign(stmt.target, this.eval(stmt.value, scope), scope, false);
        return NORMAL;
      case "set":
        this.assign(stmt.target, this.eval(stmt.value, scope), scope, true);
        return NORMAL;
      case "callStmt":
        this.eval(stmt.expr, scope, true);
        return NORMAL;
      case "if": {
        for (const b of stmt.branches) {
          if (toBool(this.eval(b.cond, scope))) return this.execBlock(b.body, scope);
        }
        return stmt.else ? this.execBlock(stmt.else, scope) : NORMAL;
      }
      case "for": {
        const slot = scope.find(stmt.varName) ?? scope.declare(stmt.varName);
        const from = toNumber(this.eval(stmt.from, scope));
        const to = toNumber(this.eval(stmt.to, scope));
        const step = stmt.step ? toNumber(this.eval(stmt.step, scope)) : 1;
        if (step === 0) throw new VbaError("a For loop with Step 0 would never end", 5, stmt.line);
        for (let i = from; step > 0 ? i <= to : i >= to; i += step) {
          this.step(stmt.line);
          slot.value = i;
          const flow = this.execBlock(stmt.body, scope);
          if (flow.kind === "exit" && flow.what === "For") break;
          if (flow.kind !== "normal") return flow;
        }
        return NORMAL;
      }
      case "forEach": {
        const slot = scope.find(stmt.varName) ?? scope.declare(stmt.varName);
        const coll = this.eval(stmt.collection, scope);
        const items = coll instanceof VbaArray ? coll.values()
          : isObject(coll) && coll.enumerate ? coll.enumerate()
          : (() => { throw new VbaError(`${typeName(coll)} cannot be used with For Each`, 438, stmt.line); })();
        for (const item of items) {
          this.step(stmt.line);
          slot.value = item;
          const flow = this.execBlock(stmt.body, scope);
          if (flow.kind === "exit" && flow.what === "For") break;
          if (flow.kind !== "normal") return flow;
        }
        return NORMAL;
      }
      case "doLoop": {
        const test = stmt.test;
        const passes = (): boolean => {
          if (!test) return true;
          const v = toBool(this.eval(test.cond, scope));
          return test.kind === "While" ? v : !v;
        };
        // A test at the end runs the body once before it is checked.
        if (test?.atEnd) {
          for (;;) {
            this.step(stmt.line);
            const flow = this.execBlock(stmt.body, scope);
            if (flow.kind === "exit" && flow.what === "Do") break;
            if (flow.kind !== "normal") return flow;
            if (!passes()) break;
          }
          return NORMAL;
        }
        while (passes()) {
          this.step(stmt.line);
          const flow = this.execBlock(stmt.body, scope);
          if (flow.kind === "exit" && flow.what === "Do") break;
          if (flow.kind !== "normal") return flow;
        }
        return NORMAL;
      }
      case "while": {
        while (toBool(this.eval(stmt.cond, scope))) {
          this.step(stmt.line);
          const flow = this.execBlock(stmt.body, scope);
          if (flow.kind === "exit" && flow.what === "Do") break;
          if (flow.kind !== "normal") return flow;
        }
        return NORMAL;
      }
      case "select": {
        const subject = this.eval(stmt.subject, scope);
        for (const clause of stmt.cases) {
          const hit = clause.tests.some((t) => {
            if (t.kind === "value") return compare(subject, this.eval(t.value, scope)) === 0;
            if (t.kind === "range") return compare(subject, this.eval(t.from, scope)) >= 0 && compare(subject, this.eval(t.to, scope)) <= 0;
            const c = compare(subject, this.eval(t.value, scope));
            switch (t.op) {
              case "=": return c === 0;
              case "<>": return c !== 0;
              case "<": return c < 0;
              case "<=": return c <= 0;
              case ">": return c > 0;
              case ">=": return c >= 0;
              default: throw new VbaError(`Case Is ${t.op} is not supported`, 5, stmt.line);
            }
          });
          if (hit) return this.execBlock(clause.body, scope);
        }
        return stmt.elseBody ? this.execBlock(stmt.elseBody, scope) : NORMAL;
      }
      case "with": {
        this.withStack.push(this.eval(stmt.subject, scope));
        try {
          return this.execBlock(stmt.body, scope);
        } finally {
          this.withStack.pop();
        }
      }
      case "exit":
        return { kind: "exit", what: stmt.what };
      case "goto":
        return { kind: "goto", label: stmt.label };
      case "onError":
        if (stmt.mode === "resumeNext") { this.onErrorMode = "resumeNext"; this.onErrorLabel = ""; }
        else if (stmt.label === "0") { this.onErrorMode = "stop"; this.onErrorLabel = ""; }
        else { this.onErrorMode = "goto"; this.onErrorLabel = stmt.label; }
        return NORMAL;
      case "unsupported":
        // Attribute and Option lines carry no behaviour; anything else is refused when reached.
        if (/^(Attribute|Option)\b/i.test(stmt.text)) return NORMAL;
        throw new VbaError(`this macro uses "${stmt.text}", which sheetedit does not support`, 5, stmt.line);
    }
  }

  private makeArray(dims: { lower?: Expr; upper: Expr }[], scope: Scope): VbaArray {
    return VbaArray.ofDims(dims.map((d) => ({
      lower: d.lower ? toNumber(this.eval(d.lower, scope)) : 0,
      upper: toNumber(this.eval(d.upper, scope)),
    })));
  }

  // --- assignment ------------------------------------------------------------

  private assign(target: Expr, value: VbaValue, scope: Scope, isSet: boolean): void {
    // `x = obj` without Set assigns the object's default value, which is the usual source of
    // "object variable not set" confusion in real VBA.
    const v = isSet ? value : isObject(value) ? scalarOf(value) : value;
    if (target.type === "name") {
      const slot = scope.find(target.name) ?? scope.declare(target.name);
      slot.value = v;
      return;
    }
    if (target.type === "member") {
      const obj = target.target ? this.eval(target.target, scope) : this.currentWith(target.line);
      if (!isObject(obj)) throw new VbaError(`cannot set ${target.name} on ${typeName(obj)}`, 424, target.line);
      if (!obj.set) throw new VbaError(`${obj.typeName} is read-only`, 383, target.line);
      obj.set(target.name, [], v);
      return;
    }
    if (target.type === "call") {
      // Either an array element or a parameterised property (Cells(1,1) = 5).
      const base = target.target;
      const args = target.args.map((a) => (a.value ? this.eval(a.value, scope) : EMPTY));
      const names = argNamesOf(target.args);
      if (base.type === "name") {
        const slot = scope.find(base.name);
        if (slot?.value instanceof VbaArray) {
          const idx = args.map(toNumber);
          if (!slot.value.inBounds(idx)) throw new VbaError("subscript out of range", 9, target.line);
          slot.value.set(idx, v);
          return;
        }
      }
      const objTarget = base.type === "member" && base.target === null ? this.currentWith(target.line) : undefined;
      const obj = objTarget ?? this.eval(base.type === "member" ? { ...base, target: base.target ?? undefined } as Expr : base, scope);
      if (isObject(obj) && obj.set) {
        obj.set(base.type === "member" ? base.name : "", args, v, names);
        return;
      }
      throw new VbaError("cannot assign to that", 424, target.line);
    }
    throw new VbaError("cannot assign to that", 424, target.line);
  }

  private currentWith(line: number): VbaValue {
    const top = this.withStack[this.withStack.length - 1];
    if (top === undefined) throw new VbaError("a leading dot needs an enclosing With", 91, line);
    return top;
  }

  // --- expressions -----------------------------------------------------------

  private eval(e: Expr, scope: Scope, asStatement = false): VbaValue {
    this.step(e.line);
    switch (e.type) {
      case "lit": return e.value === null ? NULL : e.value;
      case "nothing": return NOTHING;
      case "empty": return EMPTY;
      case "date": return this.parseDateLiteral(e.text, e.line);
      case "new": throw new VbaError(`New ${e.className} is not supported`, 5, e.line);
      case "name": return this.lookup(e.name, [], scope, e.line, asStatement);
      case "member": {
        const obj = e.target ? this.eval(e.target, scope) : this.currentWith(e.line);
        if (!isObject(obj)) throw new VbaError(`${typeName(obj)} has no member called ${e.name}`, 424, e.line);
        return obj.get(e.name, []);
      }
      case "unary": {
        if (e.op === "Not") {
          const v = this.eval(e.operand, scope);
          if (v === NULL) return NULL;
          // Not is bitwise on numbers, which for a Boolean gives the other Boolean.
          const n = toNumber(v);
          return typeof scalarOf(v) === "boolean" ? !toBool(v) : ~n;
        }
        const n = toNumber(this.eval(e.operand, scope));
        return e.op === "-" ? -n : n;
      }
      case "binary": return this.binary(e.op, e.left, e.right, scope, e.line);
      case "call": {
        // `a(1)` is an array index, a call, or a parameterised property, decided here.
        const args = e.args.map((a) => (a.value ? this.eval(a.value, scope) : EMPTY));
        const names = argNamesOf(e.args);
        if (e.target.type === "name") {
          const slot = scope.find(e.target.name);
          if (slot?.value instanceof VbaArray) {
            const idx = args.map(toNumber);
            if (!slot.value.inBounds(idx)) throw new VbaError("subscript out of range", 9, e.line);
            return slot.value.get(idx);
          }
          return this.lookup(e.target.name, args, scope, e.line, asStatement, e.args);
        }
        if (e.target.type === "member") {
          const obj = e.target.target ? this.eval(e.target.target, scope) : this.currentWith(e.line);
          if (!isObject(obj)) throw new VbaError(`${typeName(obj)} has no member called ${e.target.name}`, 424, e.line);
          return obj.get(e.target.name, args, names);
        }
        const callee = this.eval(e.target, scope);
        if (isObject(callee)) return callee.get("", args, names);
        throw new VbaError("that is not something that can be called", 424, e.line);
      }
    }
  }

  /** Resolve a bare name: a variable, a procedure in this module, a host global, or a built-in. */
  private lookup(name: string, args: VbaValue[], scope: Scope, line: number, asStatement: boolean, rawArgs?: Arg[]): VbaValue {
    const slot = scope.find(name);
    if (slot && args.length === 0) return slot.value;
    const proc = this.procs.get(name.toUpperCase());
    if (proc) {
      // ByRef needs the caller's slot, so a bare-name argument passes its own box.
      const slots: Slot[] = [];
      args.forEach((v, i) => {
        const raw = rawArgs?.[i];
        let slot: Slot = { value: v };
        if (raw?.value?.type === "name") {
          const found = scope.find(raw.value.name);
          if (found) slot = found;
        }
        // `Name:=value` goes to the parameter it names, wherever that sits.
        if (raw?.name) {
          const at = proc.params.findIndex((p) => p.name.toUpperCase() === raw.name!.toUpperCase());
          if (at < 0) throw new VbaError(`${proc.name} has no argument called ${raw.name}`, 449, line);
          slots[at] = slot;
        } else slots[i] = slot;
      });
      return this.callProcedure(proc, slots);
    }
    // The host may shadow an intrinsic, so its globals are consulted first.
    const global = this.globals.get(name.toUpperCase()) ?? this.intrinsics.get(name.toUpperCase());
    if (global !== undefined) {
      if (args.length === 0) return global;
      if (isObject(global)) return global.get("", args, argNamesOf(rawArgs ?? []));
      throw new VbaError(`${name} does not take arguments`, 5, line);
    }
    const fn = builtins[name.toUpperCase()];
    if (fn) {
      if (rawArgs?.some((a) => a.name)) throw new VbaError(`${name} does not take named arguments`, 449, line);
      return fn(args, { output: (t) => { this.messages.push(t); this.options.output?.(t); }, lastError: () => this.lastError });
    }
    if (slot) return slot.value;
    // The plan's rule: name what is missing and stop, rather than carry on with a wrong answer.
    throw new VbaError(
      asStatement
        ? `this macro calls ${name}, which sheetedit does not provide`
        : `${name} is not defined`,
      asStatement ? 5 : 2020,
      line,
    );
  }

  private parseDateLiteral(text: string, line: number): number {
    const ms = Date.parse(text);
    if (Number.isNaN(ms)) throw new VbaError(`${text} is not a date sheetedit can read`, 13, line);
    // Excel's serial: days since 1899-12-30.
    return ms / 86400000 + 25569;
  }

  private binary(op: string, leftExpr: Expr, rightExpr: Expr, scope: Scope, line: number): VbaValue {
    // And/Or short-circuit in no VBA version, but evaluating both is still correct.
    const a = this.eval(leftExpr, scope);
    const b = this.eval(rightExpr, scope);
    switch (op) {
      case "&": {
        // Concatenation is the one operator that treats Null as an empty string.
        const sa = a === NULL ? "" : toStr(a);
        const sb = b === NULL ? "" : toStr(b);
        return sa + sb;
      }
      case "+": {
        if (a === NULL || b === NULL) return NULL;
        const sa = scalarOf(a), sb = scalarOf(b);
        // + concatenates only when BOTH sides are strings; otherwise it adds.
        if (typeof sa === "string" && typeof sb === "string") return sa + sb;
        return toNumber(a) + toNumber(b);
      }
      case "-": case "*": case "/": case "\\": case "Mod": case "^": {
        if (a === NULL || b === NULL) return NULL;
        const x = toNumber(a), y = toNumber(b);
        switch (op) {
          case "-": return x - y;
          case "*": return x * y;
          case "/":
            if (y === 0) throw new VbaError("division by zero", 11, line);
            return x / y;
          case "\\": {
            if (Math.trunc(y) === 0) throw new VbaError("division by zero", 11, line);
            // Integer division truncates both operands first, then the result.
            return Math.trunc(Math.trunc(x) / Math.trunc(y));
          }
          case "Mod": {
            if (Math.trunc(y) === 0) throw new VbaError("division by zero", 11, line);
            // VBA's Mod works on integers and keeps the dividend's sign.
            return Math.trunc(x) % Math.trunc(y);
          }
          default: return Math.pow(x, y);
        }
      }
      case "=": case "<>": case "<": case ">": case "<=": case ">=": {
        if (a === NULL || b === NULL) return NULL;
        const c = compare(a, b);
        switch (op) {
          case "=": return c === 0;
          case "<>": return c !== 0;
          case "<": return c < 0;
          case ">": return c > 0;
          case "<=": return c <= 0;
          default: return c >= 0;
        }
      }
      case "Is": return a === b || (a === NOTHING && b === NOTHING);
      case "Like": return like(toStr(a), toStr(b));
      case "And": {
        if (a === NULL || b === NULL) return NULL;
        // And is bitwise on numbers; with Booleans that comes to the logical answer.
        if (typeof scalarOf(a) === "boolean" && typeof scalarOf(b) === "boolean") return toBool(a) && toBool(b);
        return toNumber(a) & toNumber(b);
      }
      case "Or": {
        if (a === NULL || b === NULL) return NULL;
        if (typeof scalarOf(a) === "boolean" && typeof scalarOf(b) === "boolean") return toBool(a) || toBool(b);
        return toNumber(a) | toNumber(b);
      }
      case "Xor": return toBool(a) !== toBool(b);
      case "Eqv": return toBool(a) === toBool(b);
      case "Imp": return !toBool(a) || toBool(b);
      default: throw new VbaError(`the operator ${op} is not supported`, 5, line);
    }
  }
}

/** Convenience: run one Sub of a parsed module. */
export function runMacro(module: Module, procName: string, options: RunOptions = {}): RunResult {
  return new VbaInterpreter(module, options).run(procName);
}

export { numToVbaString };
