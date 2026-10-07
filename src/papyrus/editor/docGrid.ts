// Word's document grid (版面設定 › 文件格線, w:docGrid) for line snapping (persona-300): in a
// section whose grid has lines (w:type "lines" or "linesAndChars", Word zh-TW's default for a new
// document), every line of a body paragraph is a whole number of grid lines (w:linePitch) tall:
//   - "auto" spacing m: max(k, m) grid lines, k the fewest grid lines the line's tallest text fits in;
//   - "at least" X: the larger of k grid lines and X;
//   - "exact" X: X, as without a grid.
// Measured in Word (COM, compatibility modes 12 and 15): space before / after is not snapped, text
// in tables is not snapped, a paragraph with w:snapToGrid w:val="0" is not snapped. A grid without
// w:type (Word's "no grid", what most generators write) changes nothing. Character spacing
// (w:charSpace, "linesAndChars") is not followed.
//
// The paragraphs get a class and the pitch; the line height itself is CSS (docGrid.css), from the
// spacing variables every paragraph and style carries (lineGridVars in schema.ts).
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { twipsToPx } from '../units';
import './docGrid.css';

/** A section's line grid: the pitch in px (null: no line grid). */
export function linePitchOf(sectPr: string | null): number | null {
  const el = sectPr ? /<(?:[\w.-]+:)?docGrid\b[^>]*>/.exec(sectPr)?.[0] : null;
  if (!el) return null;
  const type = /\bw:type="([^"]*)"/.exec(el)?.[1];
  if (type !== 'lines' && type !== 'linesAndChars') return null;
  const pitch = Number(/\bw:linePitch="(\d+)"/.exec(el)?.[1]);
  return pitch > 0 ? twipsToPx(pitch) : null;
}

const SNAP_OFF = /<(?:[\w.-]+:)?snapToGrid\b[^>]*\bw:val="(0|false|off)"/;

/** Top-level blocks of each section and the section's w:sectPr. */
export interface GridSection {
  firstBlock: number;
  lastBlock: number;
  xml: string | null;
}

function build(doc: PMNode, sections: GridSection[]): { set: DecorationSet; key: string } {
  const pitches = sections.map((s) => linePitchOf(s.xml));
  const key = sections.map((s, i) => `${s.lastBlock}:${pitches[i] ?? ''}`).join('|');
  if (pitches.every((p) => p == null)) return { set: DecorationSet.empty, key };
  const decos: Decoration[] = [];
  let s = 0;
  doc.forEach((node, offset, index) => {
    while (s < sections.length - 1 && index > sections[s].lastBlock) s++;
    const pitch = pitches[s];
    if (pitch == null || node.type.name !== 'paragraph' || SNAP_OFF.test((node.attrs.pPr as string | null) ?? '')) return;
    decos.push(Decoration.node(offset, offset + node.nodeSize, { class: 'dx-grid', style: `--dx-pitch:${+pitch.toFixed(3)}px` }));
  });
  return { set: DecorationSet.create(doc, decos), key };
}

/**
 * Whether an edit may change which paragraphs snap: blocks added or removed, or a changed block
 * with another type, section break (w:sectPr) or paragraph properties (w:snapToGrid).
 */
function gridChanged(before: PMNode, after: PMNode): boolean {
  // (The last section's w:sectPr is a document attribute.)
  if (before.childCount !== after.childCount || before.attrs.sectPr !== after.attrs.sectPr) return true;
  for (let i = 0; i < after.childCount; i++) {
    const a = before.child(i);
    const b = after.child(i);
    if (a === b) continue;
    if (a.type !== b.type || a.attrs.sectPr !== b.attrs.sectPr || a.attrs.pPr !== b.attrs.pPr) return true;
  }
  return false;
}

const gridKey = new PluginKey<{ set: DecorationSet; key: string }>('dx-doc-grid');

/** The grid classes on the body's paragraphs (see above); rebuilt when a paragraph or section changes. */
export function docGrid(getSections: (doc: PMNode) => GridSection[]): Plugin<{ set: DecorationSet; key: string }> {
  return new Plugin({
    key: gridKey,
    state: {
      init: (_, state) => build(state.doc, getSections(state.doc)),
      apply(tr, value, old, state) {
        if (!tr.docChanged) return value;
        // Typing: the same blocks, none with another section break, grid switch or type.
        if (!gridChanged(old.doc, state.doc)) return { set: value.set.map(tr.mapping, tr.doc), key: value.key };
        const sections = getSections(state.doc);
        const key = sections.map((sec) => `${sec.lastBlock}:${linePitchOf(sec.xml) ?? ''}`).join('|');
        // Typing inside a paragraph moves the decorations along; new paragraphs or sections rebuild them.
        if (key === value.key && state.doc.childCount === old.doc.childCount) {
          return { set: value.set.map(tr.mapping, tr.doc), key };
        }
        return build(state.doc, sections);
      },
    },
    props: { decorations: (state) => gridKey.getState(state)!.set },
  });
}
