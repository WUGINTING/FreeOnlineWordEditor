import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { cssIdent } from './schema';
import { tblMarginVars } from '../docx/cellMargins';
import { touches } from './steps';

/**
 * Tags every table with its table style (class dx-ts-<id>, or dx-ts--none for the default
 * table style), so the style's paragraph and run properties apply to the text in it
 * (rules made in docx/styles.ts). Tables are drawn by prosemirror-tables' own view, so the
 * class goes on through a node decoration.
 */
export function tableStyleClasses(): Plugin<DecorationSet> {
  const build = (doc: PMNode): DecorationSet => {
    const decos: Decoration[] = [];
    doc.descendants((node, pos) => {
      if (node.type.name === 'table') {
        const id = node.attrs.styleId as string | null;
        const attrs: Record<string, string> = { class: `dx-ts-${id ? cssIdent(id) : '-none'}` };
        // The table's own cell margins (w:tblCellMar) win over its style's.
        const margins = tblMarginVars(node.attrs.tblPr as string | null);
        if (margins) attrs.style = margins;
        decos.push(Decoration.node(pos, pos + node.nodeSize, attrs));
        return true; // nested tables
      }
      // Only blocks can hold tables.
      return !node.isTextblock;
    });
    return DecorationSet.create(doc, decos);
  };
  const key = new PluginKey<DecorationSet>('dx-table-styles');
  return new Plugin<DecorationSet>({
    key,
    state: {
      init: (_, state) => build(state.doc),
      // Edits that add, remove or restyle no table move the classes along with the tables.
      apply: (tr, old, _oldState, newState) =>
        !tr.docChanged ? old : touches(tr, { node: isTable }) ? build(newState.doc) : old.map(tr.mapping, tr.doc),
    },
    props: {
      decorations: (state) => key.getState(state),
    },
  });
}

const isTable = (node: PMNode) => node.type.name === 'table';
