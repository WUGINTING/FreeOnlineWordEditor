// Fonts of Taiwanese government documents on computers that don't have them (persona-300 B-8).
// Word files name Windows fonts (標楷體 = DFKai-SB, 新細明體 = PMingLiU …); a Mac, Linux or a
// tablet has none of them and the text silently fell back to the browser's sans-serif. Each
// such font gets a list of the same kind of font under its other names and on other systems
// (Mac: BiauKai, Kaiti TC, Apple LiSung …; Linux: TW-Kai, AR PL, Noto CJK), ending in the
// right generic family (楷書 and 明體 are serif), and the editor can say which of the
// document's fonts this computer doesn't have (missingFonts in editor/core.ts).

interface FontFamily {
  /** The names Word files use for it (the first is how the ribbon names it). */
  names: string[];
  /** What a person calls it. */
  label: string;
  /** The same kind of font on other systems, best first. */
  similar: string[];
  generic: 'serif' | 'sans-serif' | 'monospace';
}

const FAMILIES: FontFamily[] = [
  {
    names: ['DFKai-SB', '標楷體', 'BiauKai'],
    label: '標楷體',
    similar: ['BiauKai', 'TW-Kai', 'TW-Kai-98_1', 'Kaiti TC', 'STKaiti', 'KaiTi', 'AR PL UKai TW', 'AR PL UKai CN'],
    generic: 'serif',
  },
  {
    names: ['PMingLiU', '新細明體', 'PMingLiU-ExtB'],
    label: '新細明體',
    similar: ['MingLiU', '細明體', 'Apple LiSung', 'LiSong Pro', 'Songti TC', 'Noto Serif CJK TC', 'Noto Serif TC', 'AR PL UMing TW'],
    generic: 'serif',
  },
  {
    names: ['MingLiU', '細明體', 'MingLiU-ExtB'],
    label: '細明體',
    similar: ['PMingLiU', '新細明體', 'Apple LiSung', 'LiSong Pro', 'Songti TC', 'Noto Serif CJK TC', 'Noto Serif TC', 'AR PL UMing TW'],
    generic: 'serif',
  },
  {
    names: ['Microsoft JhengHei', '微軟正黑體', 'Microsoft JhengHei UI'],
    label: '微軟正黑體',
    similar: ['PingFang TC', 'Heiti TC', 'Noto Sans CJK TC', 'Noto Sans TC', 'Source Han Sans TC', 'WenQuanYi Zen Hei'],
    generic: 'sans-serif',
  },
  { names: ['SimSun', '宋体', '新宋体', 'NSimSun'], label: '新宋體', similar: ['Songti SC', 'STSong', 'Noto Serif CJK SC', 'AR PL UMing CN'], generic: 'serif' },
  { names: ['KaiTi', '楷体', 'KaiTi_GB2312'], label: '楷體', similar: ['STKaiti', 'Kaiti SC', 'Kaiti TC', 'AR PL UKai CN'], generic: 'serif' },
  { names: ['Times New Roman'], label: 'Times New Roman', similar: ['Times', 'Liberation Serif', 'Tinos'], generic: 'serif' },
  { names: ['Arial'], label: 'Arial', similar: ['Helvetica', 'Liberation Sans', 'Arimo'], generic: 'sans-serif' },
  { names: ['Calibri'], label: 'Calibri', similar: ['Carlito'], generic: 'sans-serif' },
  { names: ['Cambria'], label: 'Cambria', similar: ['Caladea'], generic: 'serif' },
  { names: ['Courier New'], label: 'Courier New', similar: ['Courier', 'Liberation Mono', 'Cousine'], generic: 'monospace' },
];

const byName = new Map<string, FontFamily>();
for (const f of FAMILIES) for (const n of f.names) byName.set(n.toLowerCase(), f);

/** East Asian fonts by name (the Latin letters of these are wide): names with CJK letters, and these. */
const EAST_ASIAN = /^(PMingLiU|MingLiU|DFKai-SB|BiauKai|Microsoft JhengHei|Microsoft YaHei|SimSun|NSimSun|SimHei|KaiTi|FangSong|MS Mincho|MS Gothic|MS PMincho|MS PGothic|Meiryo|Yu Gothic|Yu Mincho|Malgun Gothic|Batang|Gulim|Dotum|PingFang|Heiti|Kaiti|Songti|STKaiti|STSong|Apple LiSung|LiSong|Noto (Sans|Serif) CJK|Noto (Sans|Serif) (TC|SC|JP|KR)|Source Han|TW-Kai|TW-Sung|AR PL|WenQuanYi)/i;

/**
 * Whether choosing this font in the 字型 box applies to Chinese text too. In Word an East Asian
 * font (標楷體, 新細明體 …) is set for Chinese and Latin text alike, a Latin font (Times New Roman,
 * Arial …) only for Latin text: the Chinese text keeps its font.
 */
export function isEastAsianFont(name: string | null | undefined): boolean {
  const n = (name ?? '').trim();
  return !!n && (/[　-鿿가-힯豈-﫿]/.test(n) || EAST_ASIAN.test(n));
}

const quote = (n: string) => `"${n}"`;

/**
 * The CSS font list for a font a document names (already safe for CSS, see cssFontName): the
 * font, its other names and similar fonts. `generic` adds the generic family (only for the
 * last font of a font-family list: a generic family covers every letter).
 */
export function fontStack(name: string, generic = false): string {
  const f = byName.get(name.toLowerCase());
  if (!f) return quote(name);
  const names = [name, ...f.names, ...f.similar].filter((n, i, all) => all.findIndex((m) => m.toLowerCase() === n.toLowerCase()) === i);
  return names.map(quote).join(', ') + (generic ? `, ${f.generic}` : '');
}

/** What a person calls a font (標楷體 for DFKai-SB). */
export function fontLabel(name: string): string {
  return byName.get(name.toLowerCase())?.label ?? name;
}

/** Fonts that look alike on other systems (for the notice: which one is shown instead). */
export function similarFonts(name: string): string[] {
  const f = byName.get(name.toLowerCase());
  return f ? [...f.names.filter((n) => n.toLowerCase() !== name.toLowerCase()), ...f.similar] : [];
}

let probe: CanvasRenderingContext2D | null | undefined;
/** A test stubbed the canvas (HTMLCanvasElement.prototype.getContext is a mock). */
const vitestCanvas = () => 'mock' in (HTMLCanvasElement.prototype.getContext as object);
const known = new Map<string, boolean | null>();

/**
 * Whether this computer has a font: text in the font measures differently from the same text in
 * each generic family it would fall back to. Null when that can't be told (no canvas).
 */
export function fontInstalled(name: string): boolean | null {
  const key = name.toLowerCase();
  if (known.has(key)) return known.get(key)!;
  let result: boolean | null = null;
  try {
    // jsdom (tests) has no canvas and says so on the console each time: nothing to measure there.
    if (probe === undefined) probe = /jsdom/i.test(navigator.userAgent) && !vitestCanvas() ? null : document.createElement('canvas').getContext('2d');
    if (probe && typeof probe.measureText === 'function') {
      // Latin letters and digits: CJK letters are one em in almost every font.
      const sample = 'mmmmmmmmmmlli1WQ@#%&0123456789中文';
      result = false;
      for (const g of ['monospace', 'serif', 'sans-serif']) {
        probe.font = `72px ${g}`;
        const base = probe.measureText(sample).width;
        probe.font = `72px ${quote(name)}, ${g}`;
        if (Math.abs(probe.measureText(sample).width - base) > 0.5) {
          result = true;
          break;
        }
      }
    }
  } catch {
    result = null;
  }
  known.set(key, result);
  return result;
}

/** Forgets what fontInstalled found (tests; fonts installed meanwhile). */
export function resetFontChecks(): void {
  known.clear();
  probe = undefined;
}

export interface MissingFont {
  /** The font the document names. */
  font: string;
  /** What a person calls it (標楷體). */
  label: string;
  /** The similar font shown instead on this computer, if any (BiauKai …). */
  substitute: string | null;
}

/** Of `fonts` (names a document uses), the ones this computer doesn't have. */
export function missingFonts(fonts: Iterable<string>): MissingFont[] {
  const out: MissingFont[] = [];
  const seen = new Set<string>();
  for (const font of fonts) {
    const f = byName.get(font.toLowerCase());
    // One entry per font whatever name the file uses (DFKai-SB and 標楷體 are one font).
    const key = f ? f.label : font.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const names = f ? f.names : [font];
    const has = names.map(fontInstalled);
    if (has.some((h) => h !== false)) continue; // installed, or unknown
    out.push({ font, label: fontLabel(font), substitute: similarFonts(font).find((n) => !names.includes(n) && fontInstalled(n)) ?? null });
  }
  return out;
}
