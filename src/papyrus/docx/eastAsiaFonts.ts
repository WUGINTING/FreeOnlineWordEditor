// w:rFonts/@w:hint="eastAsia" on screen and on paper. Word picks a font per character: w:ascii for
// Latin text, w:eastAsia for CJK text. The characters both kinds of font have (①, ※, ○, ±, ×, →,
// ℃, “ ”, Greek …; HINT_RANGES below) are drawn in the Latin font, unless the run
// carries the eastAsia hint, which zh-TW Word writes on symbols entered in Chinese text (and
// 插入符號 writes too: commands.ts insertSymbol). A 公文 run in Times New Roman + 標楷體 then shows
// ① in 標楷體.
//
// The editor's font list puts the Latin font first (styles.ts FONT_FAMILY), and CSS can't pick a
// font per character from a list. It can with unicode-range: for each East Asian font the
// document uses, an @font-face alias of that local font covering only HINT_RANGES. Wherever
// --dx-font-e names a font, --dx-font-h names its alias (fontNames.ts eastAsiaAlias; set in
// styles.ts and by the font mark in schema.ts);
// a hinted run (toDOM marks it data-ea-hint) puts --dx-font-h first. The browser then takes the
// East Asian font for exactly those characters and the run's usual list for everything else: no
// element around characters, nothing in the way of typing (IME) or of the text's layout, and the
// same in the page copies of headers and footers, the printout and the PDF, which carry the same
// DOM and CSS (the rules are in the editor's own <style>, which print.ts copies). The server PDF
// has the same local fonts as the product.
//
// Only a hint on the run itself is read (w:rPr of the w:r); a hint in a paragraph or character
// style's w:rPr is not applied (a limitation: Word writes the hint on runs).
import { cssFontName, eastAsiaAlias } from './fontNames';

/**
 * The characters Word draws in a run's Latin or East Asian font depending on w:rFonts/@w:hint
 * (ECMA-376 Part 1, §17.3.2.26 rFonts; with hint="eastAsia" these take the eastAsia font). The
 * Latin-1 code points are the ones listed for the hint there; the other blocks are those Word
 * treats as shared between Latin and East Asian text:
 * - Latin-1 signs, only these: ¡ ¤ § ¨ ª soft hyphen ¯ ° ± ² ³ ´ ¶ · ¸ ¹ º
 *   ¼ ½ ¾ ¿ × ÷ (not ¢ £ ¥ ¦ © « ¬ ® µ », which keep the Latin font);
 * - modifier letters (the 注音 tone marks ˊ ˇ ˋ ˙), Greek and Cyrillic;
 * - general punctuation (— … ‧ “ ” ※ ′ ″), super/subscripts and currency signs;
 * - the symbol blocks up to the CJK ones: letterlike (℃), number forms (Ⅰ), arrows, maths,
 *   technical, enclosed alphanumerics (①), box drawing, shapes, dingbats.
 * Not Latin letters, spaces, combining marks (U+0300–036F, U+20D0–20FF) or other scripts, which
 * keep the Latin font; not the CJK blocks (U+2E80 and up), which always take the East Asian one.
 */
export const HINT_RANGES: readonly [number, number][] = [
  [0xa1, 0xa1], [0xa4, 0xa4], [0xa7, 0xa8], [0xaa, 0xaa], [0xad, 0xad], [0xaf, 0xb4], [0xb6, 0xba], [0xbc, 0xbf],
  [0xd7, 0xd7], [0xf7, 0xf7], [0x2b0, 0x2ff], [0x370, 0x4ff], [0x1f00, 0x1fff], [0x2010, 0x2027], [0x2030, 0x205e],
  [0x2070, 0x20cf], [0x2100, 0x2e7f],
];
/** Whether code point `cp` is in HINT_RANGES. */
export const inHintRanges = (cp: number) => HINT_RANGES.some(([a, b]) => cp >= a && cp <= b);

/** HINT_RANGES as a CSS unicode-range ("U+A1, U+A4, U+A7-A8, …"). */
export const HINT_UNICODE_RANGE = HINT_RANGES.map(([a, b]) => (a === b ? `U+${a.toString(16).toUpperCase()}` : `U+${a.toString(16).toUpperCase()}-${b.toString(16).toUpperCase()}`)).join(', ');

/**
 * Other names a local font is known by (Windows registers a CJK font under its Chinese and its
 * English name; documents use either). Every name of a font is given as a local() source.
 */
const FONT_NAMES: readonly string[][] = [
  ['PMingLiU', '新細明體', 'PMingLiU-ExtB', '新細明體-ExtB'],
  ['MingLiU', '細明體', 'MingLiU-ExtB', '細明體-ExtB'],
  ['DFKai-SB', '標楷體'],
  ['Microsoft JhengHei', '微軟正黑體', 'Microsoft JhengHei UI'],
  ['Microsoft YaHei', '微软雅黑'],
  ['SimSun', '宋体'],
  ['SimHei', '黑体'],
  ['KaiTi', '楷体'],
  ['MS Mincho', 'ＭＳ 明朝'],
  ['MS Gothic', 'ＭＳ ゴシック'],
];

/** The names `name` is known by, `name` first (標楷體 is also DFKai-SB). */
export function localNames(name: string): string[] {
  const same = FONT_NAMES.find((names) => names.some((n) => n.toLowerCase() === name.toLowerCase()));
  return [name, ...(same ?? []).filter((n) => n.toLowerCase() !== name.toLowerCase())];
}

/**
 * The rule that puts the alias first in a hinted run: on the run's element and on any element in
 * it that sets its own font list (the font mark's inline style), whose variables are then the
 * ones that count. !important only for that inline font list.
 */
export const HINT_RULE =
  '[data-ea-hint],[data-ea-hint] [style*="font-family"]{font-family:var(--dx-font-h),var(--dx-font-l),var(--dx-font-e),sans-serif!important}';

/** The @font-face alias of East Asian font `name`, limited to the hinted characters; '' for no usable name. */
export function eastAsiaFontFace(name: string): string {
  const font = cssFontName(name);
  if (!font) return '';
  const src = localNames(font)
    .map(cssFontName)
    .filter(Boolean)
    .map((n) => `local("${n}")`)
    .join(',');
  return `@font-face{font-family:"${eastAsiaAlias(font)}";src:${src};unicode-range:${HINT_UNICODE_RANGE}}`;
}

/** The East Asian fonts named in generated CSS (--dx-font-e:"…"): those of styles.xml and the defaults. */
export function eastAsiaFontsInCss(css: string): string[] {
  return Array.from(css.matchAll(/--dx-font-e:\s*"([^"]+)"/g), (m) => m[1]);
}

// ----- The fonts in use: one list for the page (every editor shows the same aliases) -----

const fonts = new Set<string>();
/**
 * At most this many aliases on the page. Real documents use a handful of East Asian fonts; a
 * crafted file or paste naming thousands would otherwise put a @font-face each into every editor's
 * CSS. A font past the cap keeps the run's own font list, as before hinted runs had aliases.
 */
const MAX_FONTS = 64;
const listeners = new Set<() => void>();
let notifying = false;

/**
 * An East Asian font is in use (a style's, a run's): its alias is added to the CSS of every
 * editor on the page, once. Called while text is drawn, so the editors hear of it a little later.
 */
export function useEastAsiaFont(name: string | null | undefined): void {
  const font = cssFontName(name);
  if (!font || fonts.has(font) || fonts.size >= MAX_FONTS) return;
  fonts.add(font);
  if (notifying) return;
  notifying = true;
  queueMicrotask(() => {
    notifying = false;
    for (const l of listeners) l();
  });
}

/** Calls `listener` when a font is added; returns the function that stops it. */
export function onEastAsiaFonts(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** The CSS for the fonts in use: their aliases and the rule for hinted runs. */
export function eastAsiaFontCss(): string {
  return [...Array.from(fonts, eastAsiaFontFace).filter(Boolean), HINT_RULE].join('\n');
}
