// A minimal element tree both readers walk. Built from DOMParser when there is one, otherwise from
// a small tokenizer that knows HTML's void elements and implied table-cell closings.

export interface HNode {
  /** Lower-case local name ("td", "cell", "data"). */
  tag: string;
  /** Attributes by lower-case name; XML attributes by lower-case local name. */
  attrs: Record<string, string>;
  kids: (HNode | string)[];
  parent?: HNode;
}

export const isNode = (k: HNode | string | undefined): k is HNode => typeof k === "object" && k != null;
export const childEls = (n: HNode, tag?: string): HNode[] => n.kids.filter((k): k is HNode => isNode(k) && (!tag || k.tag === tag));
export const firstEl = (n: HNode, tag: string): HNode | undefined => n.kids.find((k): k is HNode => isNode(k) && k.tag === tag);

/** Every descendant with this tag, in document order. */
export function findAll(n: HNode, tag: string, out: HNode[] = []): HNode[] {
  for (const k of n.kids) if (isNode(k)) {
    if (k.tag === tag) out.push(k);
    findAll(k, tag, out);
  }
  return out;
}

export function textOf(n: HNode): string {
  let s = "";
  for (const k of n.kids) s += typeof k === "string" ? k : textOf(k);
  return s;
}

const SS_NS = "urn:schemas-microsoft-com:office:spreadsheet";

function fromDom(el: Element, parent: HNode | undefined, xml: boolean): HNode {
  const node: HNode = { tag: (el.localName || el.nodeName).toLowerCase(), attrs: {}, kids: [], parent };
  for (const a of Array.from(el.attributes)) {
    if (xml) {
      const key = (a.localName || a.name).toLowerCase();
      // ss: attributes win over a same-named one from another namespace.
      if (a.namespaceURI === SS_NS || node.attrs[key] == null) node.attrs[key] = a.value;
    } else node.attrs[a.name.toLowerCase()] = a.value;
  }
  for (let c = el.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1) node.kids.push(fromDom(c as Element, node, xml));
    else if (c.nodeType === 3 || c.nodeType === 4) node.kids.push(c.nodeValue ?? "");
  }
  return node;
}

const VOID = new Set(["area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr", "frame"]);
const RAW = new Set(["script", "style", "textarea", "title", "xmp"]);
// HTML named entities: the Latin-1 range in code point order, then the common others.
const LATIN1 = "nbsp iexcl cent pound curren yen brvbar sect uml copy ordf laquo not shy reg macr deg plusmn sup2 sup3 acute micro para middot cedil sup1 ordm raquo frac14 frac12 frac34 iquest Agrave Aacute Acirc Atilde Auml Aring AElig Ccedil Egrave Eacute Ecirc Euml Igrave Iacute Icirc Iuml ETH Ntilde Ograve Oacute Ocirc Otilde Ouml times Oslash Ugrave Uacute Ucirc Uuml Yacute THORN szlig agrave aacute acirc atilde auml aring aelig ccedil egrave eacute ecirc euml igrave iacute icirc iuml eth ntilde ograve oacute ocirc otilde ouml divide oslash ugrave uacute ucirc uuml yacute thorn yuml";
const OTHER: Record<string, number> = { amp: 38, lt: 60, gt: 62, quot: 34, apos: 39, OElig: 0x152, oelig: 0x153, Scaron: 0x160, scaron: 0x161, Yuml: 0x178, fnof: 0x192, circ: 0x2c6, tilde: 0x2dc, ensp: 0x2002, emsp: 0x2003, thinsp: 0x2009, zwnj: 0x200c, zwj: 0x200d, lrm: 0x200e, rlm: 0x200f, ndash: 0x2013, mdash: 0x2014, lsquo: 0x2018, rsquo: 0x2019, sbquo: 0x201a, ldquo: 0x201c, rdquo: 0x201d, bdquo: 0x201e, dagger: 0x2020, Dagger: 0x2021, bull: 0x2022, hellip: 0x2026, permil: 0x2030, prime: 0x2032, Prime: 0x2033, lsaquo: 0x2039, rsaquo: 0x203a, oline: 0x203e, euro: 0x20ac, trade: 0x2122, larr: 0x2190, uarr: 0x2191, rarr: 0x2192, darr: 0x2193, harr: 0x2194, minus: 0x2212, hearts: 0x2665 };
let entities: Map<string, string> | undefined;
const entity = (name: string): string | undefined => {
  if (!entities) {
    entities = new Map(LATIN1.split(" ").map((n, i) => [n, String.fromCharCode(0xa0 + i)]));
    for (const [n, cp] of Object.entries(OTHER)) entities.set(n, String.fromCharCode(cp));
  }
  return entities.get(name) ?? entities.get(name.toLowerCase());
};

export function decodeEntities(s: string): string {
  if (!s.includes("&")) return s;
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+\d*);?/gi, (m, e: string) => {
    if (e[0] === "#") {
      const cp = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(cp) && cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : m;
    }
    return entity(e) ?? m;
  });
}

// Opening one of these closes the open element named on the right (HTML's implied end tags).
const CLOSES: Record<string, string[]> = {
  td: ["td", "th"], th: ["td", "th"], tr: ["td", "th", "tr"],
  tbody: ["td", "th", "tr", "tbody", "thead", "tfoot"], thead: ["td", "th", "tr", "tbody", "thead", "tfoot"], tfoot: ["td", "th", "tr", "tbody", "thead", "tfoot"],
  li: ["li"], option: ["option"], p: ["p"],
};
// End tags that close everything up to their element (a table end closes its open cells).
const SCOPE = new Set(["table", "tr", "tbody", "thead", "tfoot", "td", "th", "ul", "ol", "select"]);

/** Parse markup without a DOM. `xml` keeps names as written and skips HTML's implied closings. */
export function parseMarkup(text: string, xml: boolean): HNode {
  const root: HNode = { tag: "#root", attrs: {}, kids: [] };
  const stack: HNode[] = [root];
  const top = (): HNode => stack[stack.length - 1]!;
  const tagRe = /<!--[\s\S]*?(?:-->|$)|<!\[CDATA\[([\s\S]*?)\]\]>|<![^>]*>|<\?[\s\S]*?\?>|<(\/?)([A-Za-z][\w:.-]*)((?:\s+[^\s=>/]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?|\s*\/(?!>))*)\s*(\/?)>/g;
  let pos = 0;
  const text_ = (s: string): void => { if (s) top().kids.push(decodeEntities(s)); };
  for (let m = tagRe.exec(text); m; m = tagRe.exec(text)) {
    text_(text.slice(pos, m.index));
    pos = tagRe.lastIndex;
    if (m[1] != null) { top().kids.push(m[1]); continue; }
    if (!m[3]) continue; // comment, doctype, processing instruction
    const raw = m[3];
    const name = (xml ? raw.replace(/^[\w.-]+:/, "") : raw).toLowerCase();
    if (m[2]) {
      for (let i = stack.length - 1; i > 0; i--) {
        if (stack[i]!.tag === name) { stack.length = i; break; }
        if (!xml && SCOPE.has(stack[i]!.tag) && !SCOPE.has(name)) break;
      }
      continue;
    }
    if (!xml) {
      const closes = CLOSES[name];
      if (closes) {
        let cut = -1;
        for (let i = stack.length - 1; i > 0; i--) {
          const t = stack[i]!.tag;
          if (t === "table" || ((name === "td" || name === "th") && t === "tr") || (name === "li" && (t === "ul" || t === "ol"))) break;
          if (closes.includes(t)) cut = i;
        }
        if (cut > 0) stack.length = cut;
      }
    }
    const node: HNode = { tag: name, attrs: {}, kids: [], parent: top() };
    const attrRe = /([^\s=>/]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g;
    for (let a = attrRe.exec(m[4]!); a; a = attrRe.exec(m[4]!)) {
      let key = a[1]!.toLowerCase();
      if (xml) key = key.replace(/^[\w.-]+:/, "");
      if (xml && /^xmlns/.test(a[1]!)) continue;
      const v = decodeEntities(a[2] ?? a[3] ?? a[4] ?? "");
      if (!xml || /^ss:/i.test(a[1]!) || node.attrs[key] == null) node.attrs[key] = v;
    }
    top().kids.push(node);
    if (m[5] || (!xml && VOID.has(name))) continue;
    if (!xml && RAW.has(name)) {
      const end = text.toLowerCase().indexOf(`</${name}`, pos);
      const stop = end < 0 ? text.length : end;
      node.kids.push(text.slice(pos, stop));
      const close = text.indexOf(">", stop);
      pos = close < 0 ? text.length : close + 1;
      tagRe.lastIndex = pos;
      continue;
    }
    stack.push(node);
  }
  text_(text.slice(pos));
  return root;
}

/** Parse an HTML document into the tree (DOMParser when available). */
export function parseHtmlTree(text: string, useDom = true): HNode {
  if (useDom && typeof DOMParser !== "undefined") {
    const doc = new DOMParser().parseFromString(text, "text/html");
    const root: HNode = { tag: "#root", attrs: {}, kids: [] };
    root.kids.push(fromDom(doc.documentElement, root, false));
    return root;
  }
  return parseMarkup(text, false);
}

/** Parse an XML document into the tree (DOMParser when available, tolerant fallback otherwise). */
export function parseXmlTree(text: string, useDom = true): HNode {
  if (useDom && typeof DOMParser !== "undefined") {
    const doc = new DOMParser().parseFromString(text, "application/xml");
    if (!doc.getElementsByTagName("parsererror").length) {
      const root: HNode = { tag: "#root", attrs: {}, kids: [] };
      root.kids.push(fromDom(doc.documentElement, root, true));
      return root;
    }
  }
  return parseMarkup(text, true);
}
