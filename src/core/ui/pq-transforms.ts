// The transform catalogue for the Power Query editor's ribbon. Each entry is a pure M-code
// generator plus a field spec for its dialog; the editor renders the dialog, collects values,
// and appends `Table.Fn(#"prev step", ...)` as a new Applied Step. Kept pure (no DOM) so the M
// generation is unit-tested directly. Every function used here is shipped by mlang's stdlib.

const M_KEYWORDS = new Set(["and", "as", "each", "else", "error", "false", "if", "in", "is", "let", "meta", "not", "null", "otherwise", "or", "section", "shared", "then", "true", "try", "type"]);

/** Quote a name as an M identifier (for field access `[#"..."]`). */
export function quoteName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !M_KEYWORDS.has(name)) return name;
  return `#"${name.replace(/"/g, '""')}"`;
}
/** An M text literal. */
export const strLit = (s: string): string => `"${s.replace(/"/g, '""')}"`;
/** A field-access reference: `[Col]` or `[#"Col Name"]`. */
export const fieldRef = (col: string): string => `[${quoteName(col)}]`;
/** `{"A", "B"}` from a list of names. */
export const nameList = (names: string[]): string => `{${names.map(strLit).join(", ")}}`;
const isNumeric = (s: string): boolean => /^-?\d+(\.\d+)?$/.test(s.trim());
/** A comparison operand: bare when numeric, else a text literal. */
const operand = (raw: string): string => (isNumeric(raw) ? raw.trim() : strLit(raw));

const asArray = (v: string | string[] | undefined): string[] => (Array.isArray(v) ? v : v ? [v] : []);
const str = (v: string | string[] | undefined): string => (Array.isArray(v) ? (v[0] ?? "") : (v ?? ""));

export interface TfField {
  key: string;
  label: string;
  type: "text" | "number" | "select" | "columns";
  options?: { value: string; label: string }[];
  multi?: boolean; // for type "columns"
  default?: string;
  placeholder?: string;
}

export interface TransformSpec {
  id: string;
  label: string;
  group: string;
  /** Base name for the appended step (the editor makes it unique). */
  stepName: string;
  /** Fields to collect, given the columns available at the insertion point. */
  fields: (cols: string[]) => TfField[];
  /** Build the M expression from the previous step's raw name and the collected values. */
  buildM: (prev: string, v: Record<string, string | string[]>) => string;
}

const TYPE_OPTIONS = [
  { value: "type text", label: "Text" },
  { value: "Int64.Type", label: "Whole number" },
  { value: "type number", label: "Decimal number" },
  { value: "type logical", label: "True/False" },
  { value: "type date", label: "Date" },
  { value: "type datetime", label: "Date/Time" },
  { value: "type time", label: "Time" },
  { value: "type duration", label: "Duration" },
];

const FILTER_OPS = [
  { value: "eq", label: "equals" },
  { value: "ne", label: "does not equal" },
  { value: "gt", label: "greater than" },
  { value: "ge", label: "greater than or equal" },
  { value: "lt", label: "less than" },
  { value: "le", label: "less than or equal" },
  { value: "contains", label: "contains" },
  { value: "starts", label: "begins with" },
  { value: "ends", label: "ends with" },
];

function filterPredicate(col: string, op: string, value: string): string {
  const f = fieldRef(col);
  switch (op) {
    case "ne": return `${f} <> ${operand(value)}`;
    case "gt": return `${f} > ${operand(value)}`;
    case "ge": return `${f} >= ${operand(value)}`;
    case "lt": return `${f} < ${operand(value)}`;
    case "le": return `${f} <= ${operand(value)}`;
    case "contains": return `Text.Contains(${f}, ${strLit(value)})`;
    case "starts": return `Text.StartsWith(${f}, ${strLit(value)})`;
    case "ends": return `Text.EndsWith(${f}, ${strLit(value)})`;
    default: return `${f} = ${operand(value)}`;
  }
}

function aggregation(agg: string, valueCol: string): { name: string; m: string } {
  const f = fieldRef(valueCol);
  switch (agg) {
    case "sum": return { name: `Sum of ${valueCol}`, m: `each List.Sum(${f}), type number` };
    case "avg": return { name: `Average of ${valueCol}`, m: `each List.Average(${f}), type number` };
    case "max": return { name: `Max of ${valueCol}`, m: `each List.Max(${f})` };
    case "min": return { name: `Min of ${valueCol}`, m: `each List.Min(${f})` };
    default: return { name: "Count", m: `each Table.RowCount(_), Int64.Type` };
  }
}


export const TRANSFORMS: TransformSpec[] = [
  {
    id: "removeColumns", label: "Remove columns", group: "Columns", stepName: "Removed Columns",
    fields: () => [{ key: "cols", label: "Columns to remove", type: "columns", multi: true }],
    buildM: (prev, v) => `Table.RemoveColumns(${prev}, ${nameList(asArray(v.cols))})`,
  },
  {
    id: "chooseColumns", label: "Choose columns", group: "Columns", stepName: "Chosen Columns",
    fields: () => [{ key: "cols", label: "Columns to keep", type: "columns", multi: true }],
    buildM: (prev, v) => `Table.SelectColumns(${prev}, ${nameList(asArray(v.cols))})`,
  },
  {
    id: "renameColumn", label: "Rename column", group: "Columns", stepName: "Renamed Columns",
    fields: () => [
      { key: "column", label: "Column", type: "columns", multi: false },
      { key: "name", label: "New name", type: "text" },
    ],
    buildM: (prev, v) => `Table.RenameColumns(${prev}, {{${strLit(str(v.column))}, ${strLit(str(v.name))}}})`,
  },
  {
    id: "filterRows", label: "Filter rows", group: "Rows", stepName: "Filtered Rows",
    fields: () => [
      { key: "column", label: "Column", type: "columns", multi: false },
      { key: "op", label: "Condition", type: "select", options: FILTER_OPS },
      { key: "value", label: "Value", type: "text" },
    ],
    buildM: (prev, v) => `Table.SelectRows(${prev}, each ${filterPredicate(str(v.column), str(v.op), str(v.value))})`,
  },
  {
    id: "sort", label: "Sort", group: "Rows", stepName: "Sorted Rows",
    fields: () => [
      { key: "column", label: "Column", type: "columns", multi: false },
      { key: "dir", label: "Direction", type: "select", options: [{ value: "Order.Ascending", label: "Ascending" }, { value: "Order.Descending", label: "Descending" }] },
    ],
    buildM: (prev, v) => `Table.Sort(${prev}, {{${strLit(str(v.column))}, ${str(v.dir)}}})`,
  },
  {
    id: "keepTop", label: "Keep top rows", group: "Rows", stepName: "Kept First Rows",
    fields: () => [{ key: "n", label: "Number of rows", type: "number", default: "10" }],
    buildM: (prev, v) => `Table.FirstN(${prev}, ${Number(str(v.n)) || 0})`,
  },
  {
    id: "keepBottom", label: "Keep bottom rows", group: "Rows", stepName: "Kept Last Rows",
    fields: () => [{ key: "n", label: "Number of rows", type: "number", default: "10" }],
    buildM: (prev, v) => `Table.LastN(${prev}, ${Number(str(v.n)) || 0})`,
  },
  {
    id: "removeTop", label: "Remove top rows", group: "Rows", stepName: "Removed Top Rows",
    fields: () => [{ key: "n", label: "Number of rows", type: "number", default: "1" }],
    buildM: (prev, v) => `Table.Skip(${prev}, ${Number(str(v.n)) || 0})`,
  },
  {
    id: "removeDuplicates", label: "Remove duplicates", group: "Rows", stepName: "Removed Duplicates",
    fields: () => [],
    buildM: (prev) => `Table.Distinct(${prev})`,
  },
  {
    id: "reverse", label: "Reverse rows", group: "Rows", stepName: "Reversed Rows",
    fields: () => [],
    buildM: (prev) => `Table.ReverseRows(${prev})`,
  },
  {
    id: "changeType", label: "Change type", group: "Transform", stepName: "Changed Type",
    fields: () => [
      { key: "column", label: "Column", type: "columns", multi: false },
      { key: "type", label: "Type", type: "select", options: TYPE_OPTIONS },
    ],
    buildM: (prev, v) => `Table.TransformColumnTypes(${prev}, {{${strLit(str(v.column))}, ${str(v.type)}}})`,
  },
  {
    id: "replaceValues", label: "Replace values", group: "Transform", stepName: "Replaced Value",
    fields: () => [
      { key: "column", label: "Column", type: "columns", multi: false },
      { key: "find", label: "Value to find", type: "text" },
      { key: "replace", label: "Replace with", type: "text" },
    ],
    buildM: (prev, v) => `Table.ReplaceValue(${prev}, ${operand(str(v.find))}, ${operand(str(v.replace))}, Replacer.ReplaceValue, ${nameList([str(v.column)])})`,
  },
  {
    id: "splitColumn", label: "Split column by delimiter", group: "Transform", stepName: "Split Column",
    fields: () => [
      { key: "column", label: "Column", type: "columns", multi: false },
      { key: "delimiter", label: "Delimiter", type: "text", placeholder: "e.g. , or -" },
      { key: "parts", label: "Number of parts", type: "number", default: "2" },
    ],
    buildM: (prev, v) => {
      const col = str(v.column);
      const n = Math.max(2, Number(str(v.parts)) || 2);
      const names = Array.from({ length: n }, (_, i) => `${col}.${i + 1}`);
      return `Table.SplitColumn(${prev}, ${strLit(col)}, Splitter.SplitTextByDelimiter(${strLit(str(v.delimiter))}, QuoteStyle.Csv), ${nameList(names)})`;
    },
  },
  {
    id: "transpose", label: "Transpose", group: "Transform", stepName: "Transposed Table",
    fields: () => [],
    buildM: (prev) => `Table.Transpose(${prev})`,
  },
  {
    id: "promoteHeaders", label: "Use first row as headers", group: "Transform", stepName: "Promoted Headers",
    fields: () => [],
    buildM: (prev) => `Table.PromoteHeaders(${prev}, [PromoteAllScalars=true])`,
  },
  {
    id: "unpivotOthers", label: "Unpivot other columns", group: "Transform", stepName: "Unpivoted Columns",
    fields: () => [{ key: "keep", label: "Columns to keep", type: "columns", multi: true }],
    buildM: (prev, v) => `Table.UnpivotOtherColumns(${prev}, ${nameList(asArray(v.keep))}, "Attribute", "Value")`,
  },
  {
    id: "groupBy", label: "Group by", group: "Transform", stepName: "Grouped Rows",
    fields: () => [
      { key: "column", label: "Group by column", type: "columns", multi: false },
      { key: "agg", label: "Aggregation", type: "select", options: [
        { value: "count", label: "Count rows" }, { value: "sum", label: "Sum" }, { value: "avg", label: "Average" }, { value: "max", label: "Max" }, { value: "min", label: "Min" },
      ] },
      { key: "valueColumn", label: "Column to aggregate (not for Count)", type: "columns", multi: false },
    ],
    buildM: (prev, v) => {
      const { name, m } = aggregation(str(v.agg), str(v.valueColumn));
      return `Table.Group(${prev}, ${nameList([str(v.column)])}, {{${strLit(name)}, ${m}}})`;
    },
  },
  {
    id: "customColumn", label: "Custom column", group: "Add column", stepName: "Added Custom",
    fields: () => [
      { key: "name", label: "New column name", type: "text", default: "Custom" },
      { key: "expr", label: "Formula (M, per row)", type: "text", placeholder: "e.g. [Qty] * [Price]" },
    ],
    buildM: (prev, v) => `Table.AddColumn(${prev}, ${strLit(str(v.name))}, each ${str(v.expr)})`,
  },
  {
    id: "indexColumn", label: "Index column", group: "Add column", stepName: "Added Index",
    fields: () => [{ key: "start", label: "Start from", type: "select", options: [{ value: "0", label: "0" }, { value: "1", label: "1" }] }],
    buildM: (prev, v) => `Table.AddIndexColumn(${prev}, "Index", ${Number(str(v.start)) || 0}, 1, Int64.Type)`,
  },
];
