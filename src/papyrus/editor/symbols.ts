// Word's 插入 › 符號 (Insert › Symbol): the symbols offered, grouped the way Word's 符號 dialog
// groups them in its 子集 (subset) list, chosen for Taiwan government documents (公文). The
// 帶圈及括號字元 subset is also how Word users write 圍繞字元 (enclosed characters): ①, ⑴, ㈠, ㊣ …
// are single Unicode characters, so they need no special formatting and look the same in Word.
// Pure data and checks; the insertion is `insertSymbol` in commands.ts.
import { inHintRanges } from '../docx/eastAsiaFonts';

/** A 子集 of the 符號 dialog. */
export interface SymbolCategory {
  /** Stable id (the dialog's select value). */
  id: string;
  /** The name shown, as in Word's zh-TW 子集 list. */
  label: string;
  /** The symbols, one character each, in the order shown. */
  chars: readonly string[];
}

/** The characters of `text`, one per code point. */
const chars = (text: string): string[] => Array.from(text);

/** The characters from code point `from` to `to`, both included. */
function span(from: number, to: number): string[] {
  const out: string[] = [];
  for (let cp = from; cp <= to; cp++) out.push(String.fromCodePoint(cp));
  return out;
}

/** The 符號 dialog's subsets, in Word's order where Word has them. */
export const SYMBOL_CATEGORIES: readonly SymbolCategory[] = [
  {
    id: 'punctuation',
    label: '標點符號',
    chars: chars('、。，；：？！「」『』（）〔〕【】《》〈〉…—～．‧'),
  },
  {
    id: 'enclosed',
    label: '帶圈及括號字元',
    chars: [
      ...span(0x2460, 0x2473), // ①–⑳
      ...span(0x3251, 0x325f), // ㉑–㉟
      ...span(0x32b1, 0x32bf), // ㊱–㊿
      ...span(0x2474, 0x2487), // ⑴–⒇
      ...span(0x2488, 0x249b), // ⒈–⒛
      ...span(0x3220, 0x3229), // ㈠–㈩
      ...span(0x3280, 0x3289), // ㊀–㊉
      ...span(0x24b6, 0x24cf), // Ⓐ–Ⓩ
      ...span(0x24d0, 0x24e9), // ⓐ–ⓩ
      // Circled ideographs: 正 上 中 下 左 右, 印 注, 特 財 祝 労 秘 男 女 適 優.
      ...chars('㊣㊤㊥㊦㊧㊨㊞㊟㊕㊖㊗㊘㊙㊚㊛㊜㊝'),
    ],
  },
  {
    id: 'math',
    label: '數學符號',
    chars: chars('±×÷=≠≒≈≦≧<>∞√∑∫∵∴∠⊥∩∪∈⊂⊃°‰％′″'),
  },
  {
    id: 'units',
    label: '單位',
    chars: chars('℃℉㎎㎏㎜㎝㎞㎡㏄㏎㏑㏒㏕㎥㎖㎗㎘'),
  },
  {
    id: 'currency',
    label: '貨幣',
    chars: chars('＄￥€£¢₩'),
  },
  {
    id: 'arrows',
    label: '箭頭',
    chars: chars('→←↑↓↗↖↘↙⇒⇔↔↕'),
  },
  {
    id: 'shapes',
    label: '圖形與方塊',
    chars: chars('※○●◎□■△▲▽▼◇◆☆★♀♂☐☑☒✓✔✕✗♪♠♥♦♣☎☏'),
  },
  {
    id: 'roman',
    label: '羅馬數字',
    chars: [...span(0x2160, 0x216b), ...span(0x2170, 0x2179)], // Ⅰ–Ⅻ, ⅰ–ⅹ
  },
  {
    id: 'greek',
    label: '希臘字母',
    // Α–Ω and α–ω: the 24 letters each (U+03A2 is unassigned; the final ς is left out so the
    // two cases line up, as in the 希臘字母 table of Word's 符號 dialog).
    chars: [...span(0x0391, 0x03a1), ...span(0x03a3, 0x03a9), ...span(0x03b1, 0x03c1), ...span(0x03c3, 0x03c9)],
  },
  {
    id: 'bopomofo',
    label: '注音符號',
    chars: [...span(0x3105, 0x3129), ...chars('ˊˇˋ˙')], // ㄅ–ㄩ and the tone marks
  },
  {
    id: 'fullwidth',
    label: '全形英數字',
    chars: [...span(0xff10, 0xff19), ...span(0xff21, 0xff3a), ...span(0xff41, 0xff5a)], // ０–９, Ａ–Ｚ, ａ–ｚ
  },
];

/**
 * The ribbon gallery's symbols when none has been used yet (Word shows its own defaults there):
 * the reference mark, circles and check boxes, circled numbers, the enumeration comma, ellipsis
 * and dash, and the signs 公文 tables use most.
 */
export const RIBBON_SYMBOLS: readonly string[] = chars('※○●◎☐☑✓①②③、…—℃±×÷≦≧→');

/** How many symbols the ribbon gallery and the recently used list hold (Word: 20). */
export const MAX_RECENT_SYMBOLS = 20;

/**
 * Why code point `cp` can't be inserted as text, in Chinese; null when it can. Surrogates are
 * halves of a character, control characters are not text (and most are not allowed in XML, so
 * Word could not open the file), and noncharacters are reserved by Unicode.
 */
export function codePointError(cp: number): string | null {
  if (!Number.isInteger(cp) || cp < 0 || cp > 0x10ffff) return '字元代碼超出 Unicode 範圍（0 到 10FFFF）。';
  if (cp >= 0xd800 && cp <= 0xdfff) return '這是 UTF-16 代理字元（D800–DFFF），只是字元的一半，無法單獨插入。';
  if (cp <= 0x1f || (cp >= 0x7f && cp <= 0x9f)) return '這是控制字元，不是可以顯示的文字，無法插入。';
  if ((cp >= 0xfdd0 && cp <= 0xfdef) || (cp & 0xfffe) === 0xfffe) return '這是 Unicode 保留的非字元，無法插入。';
  return null;
}

/** Whether `value` is exactly one character that can be inserted as text. */
export function isSymbol(value: unknown): value is string {
  if (typeof value !== 'string' || !value) return false;
  const cp = value.codePointAt(0)!;
  return String.fromCodePoint(cp) === value && codePointError(cp) == null;
}

/** The result of reading a typed 字元代碼. */
export type SymbolCodeResult = { ok: true; char: string } | { ok: false; error: string };

/**
 * Read a 字元代碼 typed in the 符號 dialog: hexadecimal, as Word's 「Unicode (十六進位)」, with an
 * optional "U+" in front (2605, U+2605). Returns the character, or why it is refused.
 */
export function parseSymbolCode(text: string): SymbolCodeResult {
  const code = text.trim().replace(/^u\+/i, '');
  if (!code) return { ok: false, error: '請輸入字元代碼（16 進位，例如 2605）。' };
  if (!/^[0-9a-f]+$/i.test(code)) return { ok: false, error: '字元代碼只能是 16 進位數字（0–9、A–F），例如 2605。' };
  const digits = code.replace(/^0+(?=.)/, ''); // 002605 is 2605
  const cp = digits.length > 6 ? Infinity : parseInt(digits, 16);
  const error = codePointError(cp);
  return error ? { ok: false, error } : { ok: true, char: String.fromCodePoint(cp) };
}

/**
 * Whether `ch` is drawn in the East Asian font only under w:hint="eastAsia" (docx/eastAsiaFonts.ts HINT_RANGES).
 * A symbol entered in Chinese text gets that hint (commands.ts insertSymbol), as Word gives it.
 */
export function needsEastAsiaHint(ch: string): boolean {
  const cp = ch.codePointAt(0);
  return cp != null && String.fromCodePoint(cp) === ch && inHintRanges(cp);
}

/** A character's code in hexadecimal, at least four digits: "2460", "1F600". */
export const symbolHex = (ch: string): string => (ch.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0');

/** A character's code as Unicode writes it: "U+2460". */
export const symbolCode = (ch: string): string => `U+${symbolHex(ch)}`;

/** The subset a character is listed in (the first), or null. */
export function symbolCategory(ch: string): SymbolCategory | null {
  return SYMBOL_CATEGORIES.find((c) => c.chars.includes(ch)) ?? null;
}

/** The ribbon gallery: the recently used symbols first, then the defaults, 20 in all. */
export function menuSymbols(recent: readonly string[]): string[] {
  const out: string[] = [];
  for (const ch of [...recent, ...RIBBON_SYMBOLS]) {
    if (out.length >= MAX_RECENT_SYMBOLS) break;
    if (isSymbol(ch) && !out.includes(ch)) out.push(ch);
  }
  return out;
}
