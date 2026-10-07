import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import { isHistoryTransaction } from 'prosemirror-history';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { ListCounter, hasMadeLists, levelOf, syncMadeLists, usedNumIds } from '../docx/numbering';
import type { Numbering } from '../docx/model';
import { twipsToPx } from '../units';
import { touches } from './steps';

export const listMarkersKey = new PluginKey<DecorationSet>('dx-list-markers');

/** Meta flag to force a recount when the numbering definitions change. */
export const NUMBERING_CHANGED = 'dx-numbering-changed';

/**
 * Whether a node names a list: a list paragraph, or an attribute with a list id in kept XML or
 * JSON (what usedNumIds in docx/numbering.ts looks at).
 */
function namesList(node: PMNode): boolean {
  if (node.isText) return false;
  for (const [name, value] of Object.entries(node.attrs)) {
    if (value == null || value === '') continue;
    if (name === 'numId') return true;
    if (typeof value === 'string' ? value.includes('numId') : typeof value === 'object' && JSON.stringify(value).includes('numId')) return true;
  }
  return false;
}

/**
 * Whether a transaction may change which lists the document uses or how they count: it adds,
 * removes or changes a list paragraph (or an attribute naming a list). Typing inside a list
 * item, or anywhere else, does not.
 */
export function changesLists(tr: Transaction): boolean {
  return tr.docChanged && touches(tr, { node: namesList });
}

function build(doc: PMNode, numbering: Numbering): DecorationSet {
  const counter = new ListCounter(numbering);
  const decos: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph') return true;
    const { numId, ilvl } = node.attrs;
    if (!numId) return false;
    const text = counter.next(numId, ilvl);
    const lvl = levelOf(numbering, numId, ilvl);
    if (text == null || !lvl) return false;

    // Use the list level's indentation unless the paragraph sets its own.
    const style: string[] = [];
    const left = node.attrs.indLeft ?? lvl.indLeft;
    const hanging = node.attrs.indFirst != null ? -node.attrs.indFirst : lvl.hanging ?? 0;
    if (node.attrs.indLeft == null && lvl.indLeft != null) style.push(`margin-left:${twipsToPx(lvl.indLeft)}px`);
    if (node.attrs.indFirst == null && lvl.hanging != null) style.push(`text-indent:${-twipsToPx(lvl.hanging)}px`);
    if (style.length) decos.push(Decoration.node(pos, pos + node.nodeSize, { style: style.join(';') }));

    // After the number (w:suff): a tab to the hanging indent (Word's default), a space, or
    // nothing, the text right after the number (the 公文 list, docx/numbering.ts GONGWEN).
    const suff = lvl.suff ?? 'tab';
    const width = suff !== 'tab' ? 0 : hanging > 0 ? twipsToPx(hanging) : left ? 0 : 24;
    decos.push(
      Decoration.widget(
        pos + 1,
        () => {
          const span = document.createElement('span');
          span.className = suff === 'tab' ? 'dx-marker' : `dx-marker dx-marker-${suff}`;
          span.textContent = suff === 'space' ? text + ' ' : text;
          if (width) span.style.minWidth = `${width}px`;
          return span;
        },
        { side: -1, ignoreSelection: true, key: `m${text}|${width}|${suff}` },
      ),
    );
    return false;
  });
  return DecorationSet.create(doc, decos);
}

/**
 * Keeps the lists made in the editor (pasted lists, the list buttons) in step with the document,
 * so undo takes a pasted list's definition back out of the numbering and redo brings it back
 * (syncMadeLists). `otherDocs` are the documents sharing the numbering (body, headers, footers);
 * this editor's own is skipped by identity. Put it before listMarkers.
 */
export function madeListsSync(getNumbering: () => Numbering, otherDocs: () => Iterable<PMNode> = () => []): Plugin {
  return new Plugin({
    state: {
      init: () => null,
      apply: (tr, value, oldState, newState) => {
        const numbering = getNumbering();
        if (!tr.docChanged || !hasMadeLists(numbering)) return value;
        // Only an edit that adds or removes list paragraphs (or undo / redo) can change what is used.
        if (!isHistoryTransaction(tr) && !changesLists(tr)) return value;
        const docs = [newState.doc];
        for (const d of otherDocs()) if (d !== oldState.doc && d !== newState.doc) docs.push(d);
        syncMadeLists(numbering, usedNumIds(docs));
        return value;
      },
    },
  });
}

export function listMarkers(getNumbering: () => Numbering): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: listMarkersKey,
    state: {
      init: (_, state) => build(state.doc, getNumbering()),
      apply: (tr, old, _oldState, newState) => {
        if (tr.getMeta(NUMBERING_CHANGED)) return build(newState.doc, getNumbering());
        if (!tr.docChanged) return old;
        // The numbers only depend on the list paragraphs: other edits move the markers along.
        return changesLists(tr) ? build(newState.doc, getNumbering()) : old.map(tr.mapping, tr.doc);
      },
    },
    props: {
      decorations: (state) => listMarkersKey.getState(state),
    },
  });
}
