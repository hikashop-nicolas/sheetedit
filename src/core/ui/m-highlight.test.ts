import { describe, expect, it } from "vitest";
import { highlightM } from "./m-highlight";

// Strip the trailing "\n" the highlighter appends for scroll alignment.
const hl = (s: string): string => highlightM(s).replace(/\n$/, "");

describe("highlightM", () => {
  it("classifies keywords, library calls, strings, numbers and comments", () => {
    const out = hl(`let x = Table.SelectRows(t, each [Qty] > 5) // note`);
    expect(out).toContain('<span class="mtok-kw">let</span>');
    expect(out).toContain('<span class="mtok-kw">each</span>');
    expect(out).toContain('<span class="mtok-fn">Table.SelectRows</span>');
    expect(out).toContain('<span class="mtok-num">5</span>');
    expect(out).toContain('<span class="mtok-com">// note</span>');
    expect(out).toContain('<span class="mtok-id">x</span>');
  });

  it("handles text literals with escaped quotes and #-keywords", () => {
    expect(hl(`"a ""b"" c"`)).toBe('<span class="mtok-str">"a ""b"" c"</span>');
    expect(hl(`#date(2021,1,1)`)).toContain('<span class="mtok-kw">#date</span>');
    expect(hl(`#"quoted id"`)).toBe('<span class="mtok-id">#"quoted id"</span>');
  });

  it("escapes HTML metacharacters in all content", () => {
    expect(hl(`a < b & "c>d"`)).not.toContain("<b");
    expect(hl(`"x<y>&z"`)).toContain("&lt;y&gt;&amp;z");
    // The operator '<' is emitted escaped inside its span.
    expect(hl(`a<b`)).toContain('<span class="mtok-op">&lt;</span>');
  });

  it("tolerates unterminated strings and block comments without throwing", () => {
    expect(() => highlightM(`"unclosed`)).not.toThrow();
    expect(hl(`"unclosed`)).toBe('<span class="mtok-str">"unclosed</span>');
    expect(hl(`/* open`)).toBe('<span class="mtok-com">/* open</span>');
  });

  it("recognizes numbers (decimal, exponent, hex)", () => {
    expect(hl(`1.5`)).toBe('<span class="mtok-num">1.5</span>');
    expect(hl(`2e10`)).toBe('<span class="mtok-num">2e10</span>');
    expect(hl(`0xFF`)).toBe('<span class="mtok-num">0xFF</span>');
  });
});
