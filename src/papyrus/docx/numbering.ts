import { NS, attr, child, children, numAttr, serializeXml, escapeXml } from './xml';
import type { Node as PMNode } from 'prosemirror-model';
import type { AbstractNum, NumDef, NumLevel, Numbering } from './model';
import { emptyNumbering } from './model';

/** Highest list level (w:ilvl): Word has nine, 0..8. */
export const MAX_LEVEL = 8;

/** A list level as used for layout and counting: a whole number from 0 to 8. */
export function clampLevel(ilvl: number): number {
  return Number.isFinite(ilvl) ? Math.min(MAX_LEVEL, Math.max(0, Math.trunc(ilvl))) : 0;
}

/** Largest list number Word shows (w:start is limited to 32767 in its dialog). */
const MAX_NUMBER = 32767;

/** A list number as used for display: a whole number from 0 to 32767 (a list may start at 0). */
function clampNumber(n: number): number {
  return Number.isFinite(n) ? Math.min(MAX_NUMBER, Math.max(0, Math.trunc(n))) : 1;
}

export function parseNumbering(doc: Document | null): Numbering {
  const out = emptyNumbering();
  if (!doc) return out;
  const root = doc.documentElement;
  for (const an of children(root, 'abstractNum')) {
    const id = attr(an, 'abstractNumId');
    if (id == null) continue;
    const levels: NumLevel[] = [];
    for (const lvl of children(an, 'lvl')) {
      const i = numAttr(lvl, 'ilvl') ?? levels.length;
      // Levels outside 0..8 are never shown; indexing an array by them could make it huge.
      // (The definition is saved as written: see raw.)
      if (!Number.isInteger(i) || i < 0 || i > MAX_LEVEL) continue;
      const ind = child(child(lvl, 'pPr'), 'ind');
      const suff = attr(child(lvl, 'suff'), 'val');
      levels[i] = {
        fmt: attr(child(lvl, 'numFmt'), 'val') ?? 'decimal',
        text: attr(child(lvl, 'lvlText'), 'val') ?? '',
        start: numAttr(child(lvl, 'start'), 'val') ?? 1,
        indLeft: numAttr(ind, 'left') ?? numAttr(ind, 'start'),
        hanging: numAttr(ind, 'hanging'),
        ...(suff === 'space' || suff === 'nothing' ? { suff } : {}),
      };
    }
    out.abstracts[id] = { id, levels, raw: serializeXml(an) };
  }
  out.head = children(root, 'numPicBullet').map(serializeXml);
  out.tail = children(root).filter((c) => !['numPicBullet', 'abstractNum', 'num'].includes(c.localName)).map(serializeXml);
  for (const n of children(root, 'num')) {
    const id = attr(n, 'numId');
    const abstractId = attr(child(n, 'abstractNumId'), 'val');
    if (id == null || abstractId == null) continue;
    const startOverrides: Record<number, number> = {};
    for (const o of children(n, 'lvlOverride')) {
      const s = numAttr(child(o, 'startOverride'), 'val');
      const ilvl = numAttr(o, 'ilvl') ?? 0;
      if (s != null && Number.isInteger(ilvl) && ilvl >= 0 && ilvl <= MAX_LEVEL) startOverrides[ilvl] = s;
    }
    out.nums[id] = { id, abstractId, startOverrides, raw: serializeXml(n) };
  }
  return out;
}

export function levelOf(numbering: Numbering, numId: string | null, ilvl: number): NumLevel | null {
  if (!numId) return null;
  const num = numbering.nums[numId];
  if (!num) return null;
  return numbering.abstracts[num.abstractId]?.levels[ilvl] ?? null;
}

const BULLETS = ['•', '◦', '▪'];

/** The kinds of list the ribbon makes: bullets, 1. a. i., and the 公文 list 一、（一）1.（1）甲、（甲）. */
export type ListKind = 'bullet' | 'decimal' | 'gongwen';

/**
 * 公文 numbering (文書處理手冊's 一、（一）、1、（1）、甲、（甲）、子、（子）, with 1. as Word
 * writes it): Word's own formats, so Word shows the same numbers — taiwaneseCountingThousand
 * (一、二 … 十一), decimal, ideographTraditional (甲乙丙) and ideographZodiac (子丑寅), in full-width
 * brackets. Each level's number starts where the text of the level above starts; the hanging
 * indent is the number's width in a 16 pt line (320 twips a character), so the text lines up
 * after it. Nothing follows the number (w:suff "nothing", Word's 「編號之後：無」), the text comes
 * right after it: after a tab, a wider number (「十一、」「（十一）」) sent the text on to the next
 * tab stop in Word (checked in Word 16: 「十一、」's text 2½ characters further than 「十、」's,
 * now one character, as its number is).
 */
const GONGWEN: { fmt: string; text: string; chars: number }[] = [
  { fmt: 'taiwaneseCountingThousand', text: '%1、', chars: 2 },
  { fmt: 'taiwaneseCountingThousand', text: '（%2）', chars: 3 },
  { fmt: 'decimal', text: '%3.', chars: 1.5 },
  { fmt: 'decimal', text: '（%4）', chars: 2.5 },
  { fmt: 'ideographTraditional', text: '%5、', chars: 2 },
  { fmt: 'ideographTraditional', text: '（%6）', chars: 3 },
  { fmt: 'ideographZodiac', text: '%7、', chars: 2 },
  { fmt: 'ideographZodiac', text: '（%8）', chars: 3 },
  { fmt: 'decimal', text: '%9)', chars: 1.5 },
];

/**
 * Whether a list definition's first level is of this kind: a numbered list is any list of
 * numbers other than the 公文 list (so 編號清單 inside a 公文 list turns it into 1. 2. 3., as Word does).
 */
export function isListKind(abs: AbstractNum | undefined, kind: ListKind): boolean {
  const fmt = abs?.levels[0]?.fmt;
  if (kind === 'bullet') return fmt === 'bullet';
  if (kind === 'gongwen') return fmt === 'taiwaneseCountingThousand' && abs?.levels[0]?.text === '%1、';
  return fmt != null && fmt !== 'bullet' && !isListKind(abs, 'gongwen');
}

/** Create (or reuse) a list definition made by this editor. Returns its numId. */
export function ensureList(numbering: Numbering, kind: ListKind): string {
  // Reuse a list this editor created earlier (those have no original XML), not one made to restart a list.
  for (const n of Object.values(numbering.nums)) {
    const abs = numbering.abstracts[n.abstractId];
    if (n.raw || !abs || abs.raw || abs.pasted || Object.keys(n.startOverrides).length) continue;
    if (kind === 'decimal' ? abs.levels[0]?.fmt === 'decimal' : isListKind(abs, kind)) return n.id;
  }
  const absId = newId(numbering, 'abstracts');
  const levels: NumLevel[] = [];
  let at = 0;
  for (let i = 0; i < 9; i++) {
    const indLeft = 720 * (i + 1);
    if (kind === 'bullet') {
      levels.push({ fmt: 'bullet', text: BULLETS[i % 3], start: 1, indLeft, hanging: 360 });
    } else if (kind === 'gongwen') {
      const g = GONGWEN[i];
      const hanging = Math.round(g.chars * 320);
      at += hanging;
      levels.push({ fmt: g.fmt, text: g.text, start: 1, indLeft: at, hanging, suff: 'nothing' });
    } else {
      const fmts = ['decimal', 'lowerLetter', 'lowerRoman'];
      levels.push({ fmt: fmts[i % 3], text: `%${i + 1}.`, start: 1, indLeft, hanging: 360 });
    }
  }
  numbering.abstracts[absId] = { id: absId, levels };
  const numId = newId(numbering, 'nums');
  numbering.nums[numId] = { id: numId, abstractId: absId, startOverrides: {} };
  return numId;
}

/**
 * A new list for pasted content: `formats[i]` is the numbering format of level i
 * (bullet, decimal, lowerLetter ...), `start` the first number of the top level.
 * Each pasted list gets its own definition so it counts on its own. Returns its numId.
 */
/** A new list; `texts` are the numbers' texts per level (Word's lvlText, e.g. "(%1)"), "%N." by default. */
export function createList(numbering: Numbering, formats: string[], start = 1, texts: (string | undefined)[] = []): string {
  const absId = newId(numbering, 'abstracts');
  const topBullet = (formats[0] ?? 'decimal') === 'bullet';
  const fallback = topBullet ? ['bullet'] : ['decimal', 'lowerLetter', 'lowerRoman'];
  const levels: NumLevel[] = [];
  for (let i = 0; i < 9; i++) {
    // Written as w:numFmt: only a plain format name (pasted HTML can say anything).
    const given = formats[i];
    const fmt = given && /^[A-Za-z]+$/.test(given) ? given : fallback[i % fallback.length];
    const indLeft = 720 * (i + 1);
    levels.push(
      fmt === 'bullet'
        ? { fmt, text: BULLETS[i % 3], start: 1, indLeft, hanging: 360 }
        : { fmt, text: texts[i] ?? `%${i + 1}.`, start: i === 0 ? Math.max(0, Math.floor(start)) : 1, indLeft, hanging: 360 },
    );
  }
  numbering.abstracts[absId] = { id: absId, levels, pasted: true };
  const numId = newId(numbering, 'nums');
  numbering.nums[numId] = { id: numId, abstractId: absId, startOverrides: {} };
  return numId;
}

/**
 * A new list instance of `numId`'s list definition that starts again at `level`'s first
 * number (Word's 「重新從 1 開始編號」: a w:num with a w:lvlOverride/w:startOverride). Returns its numId.
 */
export function restartedList(numbering: Numbering, numId: string, level: number): string | null {
  const num = numbering.nums[numId];
  const abs = num && numbering.abstracts[num.abstractId];
  if (!num || !abs) return null;
  const ilvl = clampLevel(level);
  const id = newId(numbering, 'nums');
  numbering.nums[id] = { id, abstractId: num.abstractId, startOverrides: { [ilvl]: clampNumber(abs.levels[ilvl]?.start ?? 1) } };
  return id;
}

function nextId(ids: string[]): number {
  let max = 0;
  for (const id of ids) max = Math.max(max, Number(id) || 0);
  return max + 1;
}

/** A new id, also past the set-aside lists' (redo gives those back with their own ids). */
function newId(numbering: Numbering, kind: 'abstracts' | 'nums'): string {
  return String(nextId([...Object.keys(numbering[kind]), ...Object.keys(numbering.parked?.[kind] ?? {})]));
}

// ---------- lists made in the editor that nothing uses ----------

const NUM_ID_IN_TEXT = /\bnumId\b[^>]*?\bw:val="([^"]+)"|"numId":"([^"]+)"/g;

/**
 * The numIds the documents use: list paragraphs, and ids named in kept XML or JSON attributes
 * (a tracked paragraph change's old w:numPr, a page break's paragraph settings ...).
 */
export function usedNumIds(docs: Iterable<PMNode>): Set<string> {
  const out = new Set<string>();
  const scan = (value: unknown) => {
    const text = typeof value === 'string' ? value : value && typeof value === 'object' ? JSON.stringify(value) : '';
    if (!text.includes('numId')) return;
    for (const m of text.matchAll(NUM_ID_IN_TEXT)) out.add(m[1] ?? m[2]);
  };
  const visit = (node: PMNode) => {
    for (const [name, value] of Object.entries(node.attrs)) {
      if (name === 'numId' && typeof value === 'string') out.add(value);
      else scan(value);
    }
  };
  for (const doc of docs) {
    visit(doc);
    doc.descendants((node) => {
      if (!node.isText) visit(node);
    });
  }
  return out;
}

/** Whether there are lists made in the editor (in use or set aside), which syncMadeLists looks after. */
export function hasMadeLists(numbering: Numbering): boolean {
  if (numbering.parked && Object.keys(numbering.parked.nums).length) return true;
  return Object.values(numbering.nums).some((n) => !n.raw);
}

/**
 * Keeps the lists made in the editor in step with the documents (undo / redo of a paste or a
 * list button): one that no document uses is set aside, one set aside that is used again
 * comes back. Lists from the opened file (those with their original XML) are never touched,
 * used or not. Returns whether anything moved.
 */
export function syncMadeLists(numbering: Numbering, used: Set<string>): boolean {
  const parked = (numbering.parked ??= { abstracts: {}, nums: {} });
  let changed = false;
  for (const [id, n] of Object.entries(numbering.nums)) {
    if (n.raw || used.has(id)) continue;
    parked.nums[id] = n;
    delete numbering.nums[id];
    changed = true;
  }
  for (const [id, n] of Object.entries(parked.nums)) {
    if (!used.has(id) || numbering.nums[id]) continue;
    numbering.nums[id] = n;
    delete parked.nums[id];
    changed = true;
  }
  const wanted = new Set(Object.values(numbering.nums).map((n) => n.abstractId));
  for (const [id, a] of Object.entries(numbering.abstracts)) {
    if (a.raw || wanted.has(id)) continue;
    parked.abstracts[id] = a;
    delete numbering.abstracts[id];
    changed = true;
  }
  for (const [id, a] of Object.entries(parked.abstracts)) {
    if (!wanted.has(id) || numbering.abstracts[id]) continue;
    numbering.abstracts[id] = a;
    delete parked.abstracts[id];
    changed = true;
  }
  return changed;
}

/**
 * The numbering to save: everything from the opened file (unused definitions included, so they
 * round-trip unchanged) and only those lists made in the editor that the documents use.
 */
export function numberingToWrite(numbering: Numbering, used: Set<string>): Numbering {
  if (!hasMadeLists(numbering)) return numbering;
  const all = { ...numbering.parked?.nums, ...numbering.nums };
  const nums: Record<string, NumDef> = {};
  for (const [id, n] of Object.entries(all)) if (n.raw || used.has(id)) nums[id] = n;
  const wanted = new Set(Object.values(nums).map((n) => n.abstractId));
  const allAbstracts = { ...numbering.parked?.abstracts, ...numbering.abstracts };
  const abstracts: Record<string, AbstractNum> = {};
  for (const [id, a] of Object.entries(allAbstracts)) if (a.raw || wanted.has(id)) abstracts[id] = a;
  return { ...numbering, abstracts, nums, parked: undefined };
}

export function serializeNumbering(n: Numbering, rootAttrs = `xmlns:w="${NS.w}" xmlns:r="${NS.r}"`): string {
  const parts: string[] = [...(n.head ?? [])];
  for (const a of Object.values(n.abstracts)) parts.push(a.raw ?? abstractXml(a));
  for (const d of Object.values(n.nums)) parts.push(d.raw ?? numXml(d));
  parts.push(...(n.tail ?? []));
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:numbering ${rootAttrs}>` +
    parts.join('') +
    '</w:numbering>'
  );
}

function abstractXml(a: AbstractNum): string {
  const lvls = a.levels
    .map((l, i) => {
      const ind = l.indLeft != null ? `<w:pPr><w:ind w:left="${l.indLeft}" w:hanging="${l.hanging ?? 360}"/></w:pPr>` : '';
      const rpr = l.fmt === 'bullet' ? '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/></w:rPr>' : '';
      const suff = l.suff && l.suff !== 'tab' ? `<w:suff w:val="${l.suff}"/>` : '';
      return (
        `<w:lvl w:ilvl="${i}"><w:start w:val="${l.start}"/><w:numFmt w:val="${l.fmt}"/>${suff}` +
        `<w:lvlText w:val="${escapeXml(l.text)}"/><w:lvlJc w:val="left"/>${ind}${rpr}</w:lvl>`
      );
    })
    .join('');
  return `<w:abstractNum w:abstractNumId="${a.id}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls}</w:abstractNum>`;
}

function numXml(d: NumDef): string {
  const ov = Object.entries(d.startOverrides)
    .map(([lvl, s]) => `<w:lvlOverride w:ilvl="${lvl}"><w:startOverride w:val="${s}"/></w:lvlOverride>`)
    .join('');
  return `<w:num w:numId="${d.id}"><w:abstractNumId w:val="${d.abstractId}"/>${ov}</w:num>`;
}

// ---------- marker text ----------

/** Marker text of list number `n`; never throws (numbers from the file can be anything). */
export function formatNumber(n: number, fmt: string): string {
  n = clampNumber(n);
  switch (fmt) {
    case 'lowerLetter':
      return letters(n).toLowerCase();
    case 'upperLetter':
      return letters(n);
    case 'lowerRoman':
      return roman(n).toLowerCase();
    case 'upperRoman':
      return roman(n);
    case 'decimalZero':
      return n < 10 ? '0' + n : String(n);
    case 'none':
      return '';
    // Heavenly stems 甲乙丙… and earthly branches 子丑寅…, repeating after the last one.
    case 'ideographTraditional':
      return '甲乙丙丁戊己庚辛壬癸'[(n - 1 + 10) % 10] ?? String(n);
    case 'ideographZodiac':
      return '子丑寅卯辰巳午未申酉戌亥'[(n - 1 + 12) % 12] ?? String(n);
    case 'taiwaneseCountingThousand':
    case 'chineseCounting':
    case 'chineseCountingThousand':
      return chinese(n);
    default:
      return String(n);
  }
}

function letters(n: number): string {
  if (n < 1) return String(n);
  // Word repeats the letter: 27 -> AA, 28 -> BB
  const ch = String.fromCharCode(65 + ((n - 1) % 26));
  return ch.repeat(Math.floor((n - 1) / 26) + 1);
}

function roman(n: number): string {
  const table: [number, string][] = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  let out = '';
  for (const [v, s] of table) while (n >= v) { out += s; n -= v; }
  return out;
}

function chinese(n: number): string {
  const d = '零一二三四五六七八九';
  if (n < 10) return d[n];
  if (n < 20) return '十' + (n % 10 ? d[n % 10] : '');
  if (n < 100) return d[Math.floor(n / 10)] + '十' + (n % 10 ? d[n % 10] : '');
  if (n < 10000) {
    // 一百、一百零一、一百一十、一千零五 …, as Word counts.
    const units = ['千', '百', '十', ''];
    const digits = String(n).padStart(4, '0').split('').map(Number);
    let out = '';
    let zero = false;
    digits.forEach((v, i) => {
      if (!v) {
        zero = out !== '';
        return;
      }
      if (zero) out += '零';
      zero = false;
      out += d[v] + units[i];
    });
    return out;
  }
  return String(n);
}

/**
 * Tracks list counters while walking paragraphs in order and produces marker text.
 * Counters are per abstract list, as Word does.
 */
export class ListCounter {
  private counters = new Map<string, number[]>();
  /** The list instances (numIds) met so far. */
  private seen = new Set<string>();
  constructor(private numbering: Numbering) {}

  next(numId: string, ilvl: number): string | null {
    ilvl = clampLevel(ilvl);
    const num = this.numbering.nums[numId];
    if (!num) return null;
    const abs = this.numbering.abstracts[num.abstractId];
    const lvl = abs?.levels[ilvl];
    if (!abs || !lvl) return null;
    const key = num.abstractId;
    let c = this.counters.get(key);
    if (!c) {
      c = [];
      this.counters.set(key, c);
    }
    const startFor = (i: number) => clampNumber(num.startOverrides[i] ?? abs.levels[i]?.start ?? 1);
    // A list instance with a start override starts again there (Word's 重新從 1 開始編號): its first
    // paragraph restarts the overridden levels; later paragraphs of the definition count on.
    if (!this.seen.has(numId)) {
      this.seen.add(numId);
      const restarted = Object.keys(num.startOverrides).map(Number).filter((i) => Number.isInteger(i));
      if (restarted.length) for (let i = Math.min(...restarted); i < c.length; i++) delete c[i];
    }
    c[ilvl] = c[ilvl] == null ? startFor(ilvl) : c[ilvl] + 1;
    for (let i = ilvl + 1; i < c.length; i++) delete c[i];
    if (lvl.fmt === 'bullet') return lvl.text;
    return lvl.text.replace(/%(\d)/g, (_, d) => {
      const i = Number(d) - 1;
      const v = c![i] ?? startFor(i);
      return formatNumber(v, abs.levels[i]?.fmt ?? 'decimal');
    });
  }
}
