import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { SECTION_LABEL } from './schema';

const key = new PluginKey<DecorationSet>('dx-section-marks');

/** The w:type of a w:sectPr (null: next page, Word's default). */
function startOf(sectPr: string | null | undefined): string {
  return (sectPr && /<w:type\b[^>]*w:val="(\w+)"/.exec(sectPr)?.[1]) || 'nextPage';
}

/**
 * Word names a section break mark after how the section AFTER it starts (a continuous break is
 * the next section's w:type="continuous", not the ending section's): each paragraph ending a
 * section gets that label (see .dx-sect-end in editor.css).
 */
function labels(doc: PMNode): DecorationSet {
  const ends: { pos: number; node: PMNode }[] = [];
  doc.forEach((node, pos) => {
    if (node.type.name === 'paragraph' && node.attrs.sectPr) ends.push({ pos, node });
  });
  if (!ends.length) return DecorationSet.empty;
  const decorations = ends.map(({ pos, node }, i) => {
    const next = i + 1 < ends.length ? ends[i + 1].node.attrs.sectPr : doc.attrs.sectPr;
    const label = SECTION_LABEL[startOf(next)] ?? SECTION_LABEL.nextPage;
    return Decoration.node(pos, pos + node.nodeSize, { 'data-sect-label': label });
  });
  return DecorationSet.create(doc, decorations);
}

/** Section break marks labelled as Word labels them (by the next section's start). */
export function sectionMarks(): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key,
    state: {
      init: (_config, state) => labels(state.doc),
      apply: (tr, set) => (tr.docChanged ? labels(tr.doc) : set),
    },
    props: { decorations: (state) => key.getState(state) },
  });
}
