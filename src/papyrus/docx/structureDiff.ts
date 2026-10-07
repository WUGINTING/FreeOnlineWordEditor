import type { Node as PMNode } from 'prosemirror-model';
import { readDocx, type ReadResult } from './reader';
import { documentSections } from './sections';
import type { PageSetup } from './model';
import { tl } from '../i18n';

/**
 * What changed between two versions of a document besides the words: tables, pictures,
 * headers and footers, page setup, paragraph and text formatting, the document's styles.
 * A text comparison alone says nothing about these ("the same words" is not "the same
 * document"), so a reviewer reads this summary next to it.
 */
export interface StructureChange {
  area: 'table' | 'image' | 'headerFooter' | 'page' | 'format' | 'style';
  /** One line for the reviewer, e.g. "表格 2：列數 3 → 4". */
  text: string;
  /** Examples (paragraph starts) when many places changed the same way. */
  examples?: string[];
}

/** Compares two .docx files (older first). */
export async function compareStructure(older: ArrayBuffer | Uint8Array | Blob, newer: ArrayBuffer | Uint8Array | Blob): Promise<StructureChange[]> {
  const [a, b] = await Promise.all([readDocx(older), readDocx(newer)]);
  return structureChanges(a, b);
}

export async function structureChanges(a: ReadResult, b: ReadResult): Promise<StructureChange[]> {
  return [
    ...pageChanges(a, b),
    ...headerFooterChanges(a, b),
    ...tableChanges(a.doc, b.doc),
    ...imageChanges(a.doc, b.doc),
    ...formatChanges(a.doc, b.doc),
    ...(await styleChanges(a, b)),
  ];
}

// ----- page setup -----

const TWIPS_PER_CM = 1440 / 2.54;
const cm = (twips: number) => tl('{0} 公分', (twips / TWIPS_PER_CM).toFixed(2).replace(/\.?0+$/, ''));
const PAPER: Record<string, string> = { '11906x16838': 'A4', '16838x23811': 'A3', '8391x11906': 'A5', '12240x15840': 'Letter', '12240x20160': 'Legal' };
function paper(p: PageSetup): string {
  const short = Math.min(p.width, p.height);
  const long = Math.max(p.width, p.height);
  const name = PAPER[`${short}x${long}`] ?? `${cm(short)}×${cm(long)}`;
  return p.width > p.height ? tl('{0}橫向', name) : tl('{0}直向', name);
}

function pageChanges(a: ReadResult, b: ReadResult): StructureChange[] {
  const sa = documentSections(a.doc, a.model);
  const sb = documentSections(b.doc, b.model);
  const out: StructureChange[] = [];
  if (sa.length !== sb.length) out.push({ area: 'page', text: tl('分節數 {0} → {1}', sa.length, sb.length) });
  for (let i = 0; i < Math.min(sa.length, sb.length); i++) {
    const pa = sa[i].page;
    const pb = sb[i].page;
    const where = sb.length > 1 ? tl('第 {0} 節', i + 1) : tl('版面');
    if (paper(pa) !== paper(pb)) out.push({ area: 'page', text: tl('{0}紙張：{1} → {2}', where, paper(pa), paper(pb)) });
    const margins = (['marginTop', 'marginBottom', 'marginLeft', 'marginRight'] as const).filter((k) => pa[k] !== pb[k]);
    if (margins.length) {
      const name = { marginTop: tl('上'), marginBottom: tl('下'), marginLeft: tl('左'), marginRight: tl('右') };
      out.push({ area: 'page', text: tl('{0}邊界：{1}', where, margins.map((k) => `${name[k]} ${cm(pa[k])} → ${cm(pb[k])}`).join(tl('、'))) });
    }
    if (sa[i].titlePage !== sb[i].titlePage) out.push({ area: 'page', text: sb[i].titlePage ? tl('{0}「首頁不同」開啟', where) : tl('{0}「首頁不同」關閉', where) });
  }
  return out;
}

// ----- headers and footers -----

const KIND = { header: '頁首', footer: '頁尾' } as const;
const TYPE = { default: '', first: '（首頁）', even: '（偶數頁）' } as const;

function headerFooterChanges(a: ReadResult, b: ReadResult): StructureChange[] {
  const texts = (r: ReadResult) => {
    const m = new Map<string, string>();
    for (const hf of r.model.headerFooters) {
      const key = `${hf.kind}|${hf.type}`;
      // Several sections may have their own; list them all under the same kind.
      m.set(key, [m.get(key), hf.doc.textContent.trim()].filter((t) => t != null).join(tl(' ／ ')));
    }
    return m;
  };
  const ta = texts(a);
  const tb = texts(b);
  const out: StructureChange[] = [];
  for (const key of new Set([...ta.keys(), ...tb.keys()])) {
    const [kind, type] = key.split('|') as [keyof typeof KIND, keyof typeof TYPE];
    const name = `${tl(KIND[kind])}${tl(TYPE[type] ?? '')}`;
    const before = ta.get(key);
    const after = tb.get(key);
    if (before === after) continue;
    if (before == null) out.push({ area: 'headerFooter', text: tl('新增{0}：「{1}」', name, short(after!)) });
    else if (after == null) out.push({ area: 'headerFooter', text: tl('移除{0}（原為「{1}」）', name, short(before)) });
    else out.push({ area: 'headerFooter', text: tl('{0}內容：「{1}」→「{2}」', name, short(before), short(after)) });
  }
  return out;
}

// ----- tables -----

interface TableInfo {
  rows: number;
  cols: number;
  cells: string[];
  look: string;
  first: string;
}

function tables(doc: PMNode): TableInfo[] {
  const out: TableInfo[] = [];
  doc.descendants((node) => {
    if (node.type.name !== 'table') return true;
    const cells: string[] = [];
    const looks: unknown[] = [node.attrs.styleId, node.attrs.tblPr];
    let cols = 0;
    node.forEach((row) => {
      let n = 0;
      looks.push(row.attrs.height, row.attrs.heightRule, row.attrs.header, row.attrs.cantSplit);
      row.forEach((cell) => {
        n += cell.attrs.colspan ?? 1;
        cells.push(cell.textContent);
        looks.push(cell.attrs.background, cell.attrs.vAlign, cell.attrs.borders, cell.attrs.colwidth, cell.attrs.colspan, cell.attrs.rowspan);
      });
      cols = Math.max(cols, n);
    });
    out.push({ rows: node.childCount, cols, cells, look: JSON.stringify(looks), first: cells.find((c) => c.trim()) ?? '' });
    return false; // nested tables count with their outer table
  });
  return out;
}

function tableChanges(a: PMNode, b: PMNode): StructureChange[] {
  const ta = tables(a);
  const tb = tables(b);
  const out: StructureChange[] = [];
  if (ta.length !== tb.length) out.push({ area: 'table', text: tl('表格數 {0} → {1}', ta.length, tb.length) });
  for (let i = 0; i < Math.min(ta.length, tb.length); i++) {
    const x = ta[i];
    const y = tb[i];
    const name = y.first ? tl('表格 {0}（「{1}」）', i + 1, short(y.first, 12)) : tl('表格 {0}', i + 1);
    const parts: string[] = [];
    if (x.rows !== y.rows) parts.push(tl('列數 {0} → {1}', x.rows, y.rows));
    if (x.cols !== y.cols) parts.push(tl('欄數 {0} → {1}', x.cols, y.cols));
    if (x.rows === y.rows && x.cols === y.cols) {
      const changed = x.cells.filter((c, k) => c !== y.cells[k]).length;
      if (changed) parts.push(tl('{0} 格內容改變', changed));
    }
    if (x.look !== y.look) parts.push(tl('外觀改變（框線、底色、對齊、欄寬或列高）'));
    if (parts.length) out.push({ area: 'table', text: tl('{0}：{1}', name, parts.join(tl('；'))) });
  }
  return out;
}

// ----- pictures -----

interface ImageInfo {
  key: string;
  width: number | null;
  height: number | null;
  alt: string;
}

function images(doc: PMNode): ImageInfo[] {
  const out: ImageInfo[] = [];
  doc.descendants((node) => {
    if (node.type.name === 'image') {
      out.push({ key: hash(String(node.attrs.src ?? '')), width: node.attrs.width, height: node.attrs.height, alt: node.attrs.alt ?? '' });
    }
    return true;
  });
  return out;
}

function imageChanges(a: PMNode, b: PMNode): StructureChange[] {
  const ia = images(a);
  const ib = images(b);
  const out: StructureChange[] = [];
  // Same picture (same bytes) in the same order: match them; the rest were added or removed.
  const left = [...ia];
  const matched: [ImageInfo, ImageInfo][] = [];
  let added = 0;
  for (const y of ib) {
    const k = left.findIndex((x) => x.key === y.key);
    if (k < 0) added++;
    else matched.push([left.splice(k, 1)[0], y]);
  }
  if (added) out.push({ area: 'image', text: tl('新增或替換 {0} 張圖片', added) });
  if (left.length) out.push({ area: 'image', text: tl('移除或替換 {0} 張圖片', left.length) });
  const resized = matched.filter(([x, y]) => size(x) !== size(y));
  if (resized.length) {
    out.push({ area: 'image', text: tl('{0} 張圖片改變大小', resized.length), examples: resized.slice(0, 5).map(([x, y]) => `${size(x)} → ${size(y)}`) });
  }
  const alt = matched.filter(([x, y]) => x.alt !== y.alt).length;
  if (alt) out.push({ area: 'image', text: tl('{0} 張圖片的替代文字改變', alt) });
  return out;
}

const size = (i: ImageInfo) => (i.width && i.height ? tl('{0}×{1} 像素', Math.round(i.width), Math.round(i.height)) : tl('原始大小'));

// ----- formatting -----

const PARA_ASPECTS: [string, string[]][] = [
  ['段落樣式', ['styleId']],
  ['對齊', ['align']],
  ['縮排', ['indLeft', 'indRight', 'indFirst', 'indLeftChars', 'indRightChars', 'indFirstChars']],
  ['段落間距或行距', ['spaceBefore', 'spaceAfter', 'line', 'lineRule']],
  ['清單編號', ['numId', 'ilvl']],
  ['分頁設定', ['pageBreakBefore']],
];
/** Marks that are formatting (not links, revisions, comments or wrappers, which are not "format"). */
const FORMAT_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'color', 'highlight', 'font', 'fontSize', 'charStyle', 'run']);

interface Para {
  text: string;
  attrs: Record<string, unknown>;
  runs: string;
}

function paragraphs(doc: PMNode): Para[] {
  const out: Para[] = [];
  doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    const runs: string[] = [];
    node.forEach((child) => {
      const marks = child.marks.filter((m) => FORMAT_MARKS.has(m.type.name)).map((m) => `${m.type.name}${JSON.stringify(m.attrs)}`).sort();
      runs.push(`${child.isText ? child.text!.length : child.type.name}:${marks.join(',')}`);
    });
    out.push({ text: node.textContent, attrs: node.attrs, runs: runs.join('|') });
    return false;
  });
  return out;
}

function formatChanges(a: PMNode, b: PMNode): StructureChange[] {
  const pa = paragraphs(a);
  const pb = paragraphs(b);
  // Pair paragraphs with the same text, in order (a paragraph whose words changed is the
  // text comparison's business).
  const byText = new Map<string, number[]>();
  pa.forEach((p, i) => byText.set(p.text, [...(byText.get(p.text) ?? []), i]));
  const counts = new Map<string, string[]>();
  let last = -1;
  for (const p of pb) {
    if (!p.text.trim()) continue;
    const list = byText.get(p.text);
    const k = list?.findIndex((i) => i > last) ?? -1;
    if (!list || k < 0) continue;
    const i = list.splice(k, 1)[0];
    last = i;
    const q = pa[i];
    const aspects = PARA_ASPECTS.filter(([, keys]) => keys.some((key) => (q.attrs[key] ?? null) !== (p.attrs[key] ?? null))).map(([name]) => tl(name));
    if (q.runs !== p.runs) aspects.push(tl('文字格式（粗體、字型、字級、顏色等）'));
    for (const name of aspects) counts.set(name, [...(counts.get(name) ?? []), short(p.text, 20)]);
  }
  return [...counts].map(([name, where]) => ({
    area: 'format' as const,
    text: tl('{0} 段的{1}改變', where.length, name),
    examples: where.slice(0, 5),
  }));
}

// ----- document styles -----

// Compared as read (the look the styles give, the list definitions), not as XML text: saving
// writes the same definitions with attributes in another order.
async function styleChanges(a: ReadResult, b: ReadResult): Promise<StructureChange[]> {
  const out: StructureChange[] = [];
  if (a.model.css !== b.model.css) {
    out.push({ area: 'style', text: tl('文件的樣式定義改變（套用同一樣式的文字外觀可能不同）') });
  }
  if (JSON.stringify(a.model.numbering) !== JSON.stringify(b.model.numbering)) {
    out.push({ area: 'style', text: tl('清單編號的定義改變') });
  }
  return out;
}

// ----- helpers -----

function short(text: string, max = 30): string {
  const t = text.replace(/\s+/g, ' ').trim();
  return t.length > max ? t.slice(0, max) + '…' : t;
}

function hash(text: string): string {
  let h = 5381;
  for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0;
  return `${text.length}:${(h >>> 0).toString(36)}`;
}
