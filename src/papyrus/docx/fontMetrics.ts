// Word's "single" line height per font, as a multiple of the font size. Word takes it from the
// font's own metrics (East Asian fonts get extra room), so it differs a lot between fonts:
// 微軟正黑體 lines are 1.75 × the size, Calibri's 1.22. Measured in Word (Windows, 12 pt,
// no document grid). A line holding East Asian text is as tall as the taller of its Latin and
// East Asian fonts; a Latin-only or empty line uses the Latin font.

const RATIOS: Record<string, number> = {
  calibri: 1.217,
  'calibri light': 1.217,
  cambria: 1.179,
  'times new roman': 1.146,
  arial: 1.146,
  consolas: 1.179,
  'courier new': 1.133,
  'microsoft jhenghei': 1.75,
  微軟正黑體: 1.75,
  'microsoft yahei': 1.717,
  微软雅黑: 1.717,
  pmingliu: 1.321,
  新細明體: 1.321,
  mingliu: 1.321,
  細明體: 1.321,
  'mingliu_hkscs': 1.321,
  'dfkai-sb': 1.321,
  標楷體: 1.321,
  simsun: 1.321,
  宋体: 1.321,
  nsimsun: 1.321,
  新宋体: 1.321,
};

/** Fallbacks for fonts not in the table. */
export const LATIN_LINE = 1.15;
export const EAST_ASIAN_LINE = 1.321;

const EAST_ASIAN_CHAR = /[⺀-鿿가-힯豈-﫿＀-￯]/;

/** Word's single line height for a font, as a multiple of its size (null when no font is named). */
export function wordLineRatio(font: string | null | undefined): number | null {
  if (!font) return null;
  const name = font.trim().replace(/^["']|["']$/g, '').toLowerCase();
  return RATIOS[name] ?? (EAST_ASIAN_CHAR.test(name) ? EAST_ASIAN_LINE : LATIN_LINE);
}

/** Whether text holds East Asian characters (its lines use the East Asian font's height). */
export function hasEastAsianText(text: string): boolean {
  return EAST_ASIAN_CHAR.test(text);
}
