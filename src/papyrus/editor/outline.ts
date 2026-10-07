import type { Node as PMNode } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import type { StyleInheritance } from '../docx/inheritance';

/**
 * The document's headings for the navigation pane (Word's 導覽窗格): paragraphs with an
 * outline level, set on the paragraph (w:outlineLvl) or by its style (標題 1 ... 標題 9).
 */
export interface Heading {
  /** 0 = level 1. */
  level: number;
  text: string;
  /** Position of the heading paragraph. */
  pos: number;
}

const OUTLINE = /<(?:\w+:)?outlineLvl\b[^>]*\bw:val="(\d)"/;

function outlineLevel(node: PMNode, styles?: StyleInheritance): number | null {
  const own = OUTLINE.exec((node.attrs.pPr as string | null) ?? '');
  const styleId = node.attrs.styleId as string | null;
  const level = own ? Number(own[1]) : styleId ? styles?.paragraph[styleId]?.outlineLevel ?? builtIn(styleId) : null;
  return level != null && level >= 0 && level < 9 ? level : null;
}

/** The editor's own new documents use "Heading1" ... without an outline level in styles.xml. */
function builtIn(styleId: string): number | null {
  const m = /^heading(\d)$/i.exec(styleId);
  return m ? Number(m[1]) - 1 : null;
}

const cache = new WeakMap<PMNode, { styles?: StyleInheritance; headings: Heading[] }>();

export function documentOutline(doc: PMNode, styles?: StyleInheritance): Heading[] {
  const hit = cache.get(doc);
  if (hit && hit.styles === styles) return hit.headings;
  const headings: Heading[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph') return true;
    const level = outlineLevel(node, styles);
    const text = node.textContent.replace(/\s+/g, ' ').trim();
    if (level != null && text) headings.push({ level, text, pos });
    return false;
  });
  cache.set(doc, { styles, headings });
  return headings;
}

// ----- page references (the table of contents' page numbers, 「見第 N 頁」) -----

const BOOKMARK = /<(?:\w+:)?bookmarkStart\b[^>]*\bw:name="([^"]+)"/;

/** Where each bookmark starts. */
export function bookmarkPositions(doc: PMNode): Map<string, number> {
  const out = new Map<string, number>();
  doc.descendants((node, pos) => {
    if (node.type.name === 'raw_inline') {
      const m = BOOKMARK.exec(node.attrs.xml as string);
      if (m && !out.has(m[1])) out.set(m[1], pos);
    }
    return true;
  });
  return out;
}

/** The displayed results of PAGEREF fields: the text range and the bookmark it points to. */
export function pageReferences(doc: PMNode): { from: number; to: number; bookmark: string; instr: string }[] {
  const out: { from: number; to: number; bookmark: string; instr: string }[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText) return true;
    const mark = node.marks.find((m) => m.type.name === 'fieldResult' && /^\s*PAGEREF\s/i.test(m.attrs.instr));
    if (!mark) return false;
    const bookmark = String(mark.attrs.instr).trim().split(/\s+/)[1] ?? '';
    const last = out[out.length - 1];
    // One field's result may be several text nodes (different formatting): one range.
    if (last && last.to === pos && last.bookmark === bookmark) last.to = pos + node.nodeSize;
    else out.push({ from: pos, to: pos + node.nodeSize, bookmark, instr: String(mark.attrs.instr) });
    return false;
  });
  return out;
}

/**
 * Updates the page numbers shown by PAGEREF fields (a table of contents, 「見第 N 頁」) from
 * the laid-out pages, as Word's 更新功能變數 does. Returns how many changed; the text keeps its
 * formatting. `pageOf` gives the page number at a document position as Word prints it there (the
 * section's page number format, or the field's own \* switch: `instr`).
 */
export function updatePageReferences(tr: Transaction, pageOf: (pos: number, instr: string) => number | string | null): number {
  const marks = bookmarkPositions(tr.doc);
  // Every page is looked up first, on the unchanged document: a new number of another length
  // moves the positions after it (the headings come after the table of contents).
  const updates: { from: number; to: number; text: string }[] = [];
  for (const ref of pageReferences(tr.doc)) {
    const at = marks.get(ref.bookmark);
    const page = at == null ? null : pageOf(at, ref.instr);
    if (page == null) continue;
    const text = String(page);
    if (tr.doc.textBetween(ref.from, ref.to) !== text) updates.push({ from: ref.from, to: ref.to, text });
  }
  // From the end, so the earlier positions stay valid.
  for (const u of updates.reverse()) {
    const marksHere = tr.doc.nodeAt(u.from)?.marks ?? [];
    tr.replaceWith(u.from, u.to, tr.doc.type.schema.text(u.text, marksHere));
  }
  return updates.length;
}
