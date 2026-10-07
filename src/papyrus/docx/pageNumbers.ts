// Page numbers as Word prints them: a section's w:pgNumType/@w:fmt (and a field's \* switch)
// decide how PAGE fields and page references show the number.
//
// The results were read from Word itself (PAGE fields in sections restarted at 0..32767 with each
// format, Word 2016+ zh-TW), not from ECMA-376's examples, which differ: e.g. Word shows
// taiwaneseCountingThousand digit by digit (10 → 一○) and taiwaneseCounting digit by digit from
// 100 on (100 → 一○○). These are page numbers only: list numbers are docx/numbering.ts.

/** Page number formats the editor shows (w:pgNumType/@w:fmt values). */
export const PAGE_NUMBER_FORMATS = [
  'decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter',
  'taiwaneseCounting', 'chineseCounting', 'ideographTraditional', 'taiwaneseCountingThousand',
  'chineseCountingThousand', 'taiwaneseDigital', 'ideographDigital', 'ideographLegalTraditional',
  'ideographZodiac', 'decimalZero', 'decimalFullWidth', 'numberInDash',
] as const;

const SHOWN = new Set<string>(PAGE_NUMBER_FORMATS);

/** Whether the editor shows page numbers in this format like Word (else as 1, 2, 3). */
export function isShownPageFormat(fmt: string | null | undefined): boolean {
  return !fmt || SHOWN.has(fmt);
}

/** What Word shows for a letter page number past ZZZ…Z (30 letters). */
export const LETTER_OVERFLOW = '錯誤! 數字無法以指定格式顯示。';

const DIGITS = '○一二三四五六七八九';
const MAX = 2147483647;

/** Page number `n` in a w:fmt format, as Word shows it in a PAGE field. Unknown formats: 1, 2, 3. */
export function formatPageNumber(n: number, fmt: string | null | undefined): string {
  n = Number.isFinite(n) ? Math.min(MAX, Math.max(0, Math.trunc(n))) : 1;
  switch (fmt) {
    case 'upperRoman':
      return roman(n);
    case 'lowerRoman':
      return roman(n).toLowerCase();
    case 'upperLetter':
      return letters(n);
    case 'lowerLetter':
      return letters(n).toLowerCase();
    case 'taiwaneseCounting':
    case 'chineseCounting':
      return n < 100 ? counting(n) : digitwise(n, '○');
    case 'taiwaneseCountingThousand':
    case 'taiwaneseDigital':
      return digitwise(n, '○');
    case 'ideographDigital':
      return digitwise(n, '〇');
    case 'chineseCountingThousand':
      return thousands(n, '〇一二三四五六七八九', ['', '十', '百', '千'], ['', '万', '亿'], false);
    case 'ideographLegalTraditional':
      return thousands(n, '零壹貳參肆伍陸柒捌玖', ['', '拾', '佰', '仟'], ['', '萬', '億'], true);
    case 'ideographTraditional':
      return n >= 1 && n <= 10 ? '甲乙丙丁戊己庚辛壬癸'[n - 1] : String(n);
    case 'ideographZodiac':
      // Word shows 戍 (not 戌) for the eleventh.
      return n >= 1 && n <= 12 ? '子丑寅卯辰巳午未申酉戍亥'[n - 1] : String(n);
    case 'decimalZero':
      return n < 10 ? '0' + n : String(n);
    case 'decimalFullWidth':
      return String(n).replace(/\d/g, (d) => String.fromCharCode(0xff10 + Number(d)));
    case 'numberInDash':
      return `- ${n} -`;
    default:
      return String(n);
  }
}

/** Roman numerals; thousands repeat M (Word: 4000 → MMMM); 0 shows nothing. */
function roman(n: number): string {
  const table: [number, string][] = [
    [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  // Word gives up far beyond any real page number; keep the string bounded.
  if (n > 100000) return String(n);
  let out = 'M'.repeat(Math.floor(n / 1000));
  n %= 1000;
  for (const [v, s] of table) while (n >= v) { out += s; n -= v; }
  return out;
}

/** A, B … Z, AA, BB … ZZ, AAA …; 0 shows nothing, past 780 Word shows an error text. */
function letters(n: number): string {
  if (n < 1) return '';
  if (n > 780) return LETTER_OVERFLOW;
  return String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
}

/** 一 … 九, 十, 十一 … 九十九 (0: ○). */
function counting(n: number): string {
  if (n < 10) return DIGITS[n];
  const tens = Math.floor(n / 10);
  const ones = n % 10;
  return (tens > 1 ? DIGITS[tens] : '') + '十' + (ones ? DIGITS[ones] : '');
}

/** Each digit as a character: 2024 → 二○二四. */
function digitwise(n: number, zero: string): string {
  return String(n).replace(/\d/g, (d) => (d === '0' ? zero : DIGITS[Number(d)]));
}

/**
 * Full counting with units: 2024 → 二千〇二十四, 10001 → 一万〇一. A run of zeros reads as one
 * zero; `legal` keeps the leading 壹 of 10..19 (壹拾), the plain form drops it (十一).
 */
function thousands(n: number, digits: string, units: string[], groups: string[], legal: boolean): string {
  if (n === 0) return digits[0];
  const parts: string[] = [];
  let zeroPending = false;
  let out = '';
  // Groups of four digits, from the lowest.
  while (n > 0) {
    parts.push(String(n % 10000));
    n = Math.floor(n / 10000);
  }
  for (let g = parts.length - 1; g >= 0; g--) {
    const value = Number(parts[g]);
    if (value === 0) {
      zeroPending = out !== '';
      continue;
    }
    const s = String(value).padStart(4, '0');
    let text = '';
    for (let i = 0; i < 4; i++) {
      const d = Number(s[i]);
      const unit = units[3 - i];
      if (d === 0) {
        if (text || out) zeroPending = true;
        continue;
      }
      if (zeroPending && (text || out)) text += digits[0];
      zeroPending = false;
      text += digits[d] + unit;
    }
    out += text + (groups[g] ?? '');
  }
  // 十一 rather than 一十一 when the number starts with its tens.
  if (!legal && out.startsWith(digits[1] + units[1])) out = out.slice(1);
  return out;
}

/**
 * The number format a field's \* switch asks for (w:fmt names), or null when the field follows
 * its section. `PAGE \* roman` → lowerRoman; \* MERGEFORMAT and \* CHARFORMAT don't format.
 */
export function fieldNumberFormat(instr: string): string | null {
  const re = /\\\*\s*("?)([A-Za-z]+)\1/g;
  let fmt: string | null = null;
  for (const m of instr.matchAll(re)) {
    const s = m[2];
    const lower = s === s.toLowerCase();
    switch (s.toLowerCase()) {
      case 'roman':
        fmt = lower ? 'lowerRoman' : 'upperRoman';
        break;
      case 'alphabetic':
        fmt = lower ? 'lowerLetter' : 'upperLetter';
        break;
      case 'arabic':
        fmt = 'decimal';
        break;
      case 'arabicdash':
        fmt = 'numberInDash';
        break;
      default:
        break;
    }
  }
  return fmt;
}
