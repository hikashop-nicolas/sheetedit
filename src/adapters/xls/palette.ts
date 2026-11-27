// The BIFF8 colour palette: 56 entries at indexes 8..63, overridable by a PALETTE record.
export const DEFAULT_PALETTE = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
].map((h) => "#" + h.toLowerCase());

/** A palette index to CSS, or undefined for the automatic / system colours. */
export function icvCss(palette: string[], icv: number): string | undefined {
  if (icv < 8) return DEFAULT_PALETTE[icv];
  if (icv >= 8 && icv < 64) return palette[icv - 8];
  return undefined;
}

const rgb = (css: string): [number, number, number] => [parseInt(css.slice(1, 3), 16), parseInt(css.slice(3, 5), 16), parseInt(css.slice(5, 7), 16)];

/** Map CSS colours onto a 56-entry palette: exact matches first, then free slots, then nearest. */
export class PaletteBuilder {
  readonly colors: string[];
  private used = new Set<number>();
  private custom = false;
  constructor(base: string[] = DEFAULT_PALETTE) {
    this.colors = base.slice(0, 56);
  }
  get changed(): boolean { return this.custom; }
  /** Mark the indexes the file already uses so they are not repurposed. */
  reserve(icv: number): void { if (icv >= 8 && icv < 64) this.used.add(icv - 8); }
  index(css: string | undefined, fallback: number): number {
    if (!css || !/^#[0-9a-f]{6}$/i.test(css)) return fallback;
    const want = css.toLowerCase();
    const exact = this.colors.indexOf(want);
    if (exact >= 0) { this.used.add(exact); return exact + 8; }
    // Repurpose an entry nothing references yet, from the end (rarely used greys / dark tones).
    for (let i = 55; i >= 16; i--) {
      if (!this.used.has(i)) {
        this.colors[i] = want;
        this.used.add(i);
        this.custom = true;
        return i + 8;
      }
    }
    const [r, g, b] = rgb(want);
    let best = 0, bestD = Infinity;
    this.colors.forEach((c, i) => {
      const [r2, g2, b2] = rgb(c);
      const d = 2 * (r - r2) ** 2 + 4 * (g - g2) ** 2 + 3 * (b - b2) ** 2;
      if (d < bestD) { bestD = d; best = i; }
    });
    return best + 8;
  }
}
