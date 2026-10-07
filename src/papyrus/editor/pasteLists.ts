import { DOMSerializer, Fragment, Slice, type DOMOutputSpec, type Node as PMNode } from 'prosemirror-model';
import { COPIED_LIST, PASTED_LIST, copiedLists, pastedLists, schema, type PastedList } from './schema';
import { createList } from '../docx/numbering';
import type { Numbering } from '../docx/model';

// A list paragraph copied from an editor carries its numId, the id of the document it came
// from (data-doc) and how its list looks (data-num-def). Pasted into the same document it
// continues its list; pasted into another one, where the same numId may be a different list,
// it gets a new list that looks the same.

const documentIds = new WeakMap<Numbering, string>();

/** The id of the document owning `numbering` (one per opened document), for data-doc. */
export function documentId(numbering: Numbering): string {
  let id = documentIds.get(numbering);
  if (!id) {
    const bytes = new Uint8Array(12);
    if (globalThis.crypto?.getRandomValues) globalThis.crypto.getRandomValues(bytes);
    else for (let i = 0; i < bytes.length; i++) bytes[i] = Math.floor(Math.random() * 256);
    id = Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
    documentIds.set(numbering, id);
  }
  return id;
}

/** How list `numId` looks (formats and number texts per level, first number), for data-num-def. */
function listDef(numbering: Numbering, numId: string): PastedList | null {
  const num = numbering.nums[numId];
  const abs = num && numbering.abstracts[num.abstractId];
  if (!abs || !abs.levels.length) return null;
  const levels = Array.from(abs.levels.slice(0, 9), (l) => l ?? { fmt: 'decimal', text: '', start: 1 });
  return {
    formats: levels.map((l) => l.fmt),
    texts: levels.map((l) => (l.fmt === 'bullet' ? undefined : l.text)),
    start: num.startOverrides[0] ?? levels[0].start ?? 1,
  };
}

/** The editor's clipboard serializer: list paragraphs also carry data-doc and data-num-def. */
export function listClipboardSerializer(getNumbering: () => Numbering): DOMSerializer {
  const base = DOMSerializer.fromSchema(schema);
  const paragraph = base.nodes.paragraph;
  const rawInline = base.nodes.raw_inline;
  return new DOMSerializer(
    {
      ...base.nodes,
      // A tracked deletion's text is in the page for screen readers, never on the clipboard: an
      // empty span of its class (which pasting ignores), without its text, data-text or title.
      raw_inline: (node: PMNode): DOMOutputSpec => {
        // A shape travels as its run, pasted back as the same shape (see schema.ts).
        if (node.attrs.shape && node.attrs.xml) return ['span', { 'data-dx-shape-run': node.attrs.xml }];
        const spec = rawInline(node);
        const del = node.attrs.label === 'del' || node.attrs.label === 'moveFrom';
        if (!del || !Array.isArray(spec)) return spec;
        const attrs = spec[1] && typeof spec[1] === 'object' && !Array.isArray(spec[1]) ? (spec[1] as Record<string, string>) : {};
        return [spec[0], attrs.class ? { class: attrs.class } : {}] as DOMOutputSpec;
      },
      paragraph: (node: PMNode): DOMOutputSpec => {
        const spec = paragraph(node);
        const numId = node.attrs.numId as string | null;
        if (!numId || !Array.isArray(spec) || typeof spec[1] !== 'object' || spec[1] == null || Array.isArray(spec[1])) return spec;
        const numbering = getNumbering();
        const def = listDef(numbering, numId);
        const extra: Record<string, string> = { 'data-doc': documentId(numbering) };
        if (def) extra['data-num-def'] = JSON.stringify(def);
        return [spec[0], { ...(spec[1] as Record<string, string>), ...extra }, ...spec.slice(2)] as DOMOutputSpec;
      },
    },
    base.marks,
  );
}

/**
 * Pasted list items carry placeholder list ids (see schema.ts). This gives them real lists
 * of the document: an HTML list becomes a new Word list with its levels and start number;
 * a list paragraph copied inside this document keeps its own list, and one copied from
 * another document gets a new list looking like the one it came from.
 */
export function adoptPastedLists(slice: Slice, numbering: Numbering): Slice {
  const made = new Map<string, string | null>();
  const resolve = (placeholder: string): string | null => {
    if (made.has(placeholder)) return made.get(placeholder)!;
    let numId: string | null = null;
    if (placeholder.startsWith(COPIED_LIST)) {
      const rest = placeholder.slice(COPIED_LIST.length);
      const at = rest.lastIndexOf('@');
      const own = at < 0 ? rest : rest.slice(0, at);
      const from = at < 0 ? '' : rest.slice(at + 1);
      const def = copiedLists.get(placeholder);
      copiedLists.delete(placeholder);
      // A list of this document (possibly set aside by an undo: pasting it brings it back).
      // Without data-doc (HTML from the schema's own serializer, not the editor's clipboard)
      // the numId is taken to be this document's, as before.
      const here = (!from || from === documentId(numbering)) && !!(numbering.nums[own] ?? numbering.parked?.nums[own]);
      numId = here ? own : createList(numbering, def?.formats ?? ['decimal'], def?.start ?? 1, def?.texts ?? []);
    } else {
      const list = pastedLists.get(placeholder);
      pastedLists.delete(placeholder);
      numId = list ? createList(numbering, list.formats, list.start, list.texts) : null;
    }
    made.set(placeholder, numId);
    return numId;
  };
  let changed = false;
  const map = (node: PMNode): PMNode => {
    if (node.type === schema.nodes.paragraph) {
      const id = node.attrs.numId as string | null;
      if (!id || !(id.startsWith(PASTED_LIST) || id.startsWith(COPIED_LIST))) return node;
      changed = true;
      const numId = resolve(id);
      return node.type.create({ ...node.attrs, numId, ilvl: numId ? node.attrs.ilvl : 0 }, node.content, node.marks);
    }
    if (node.isLeaf || !node.childCount) return node;
    const kids: PMNode[] = [];
    node.forEach((c) => kids.push(map(c)));
    return node.copy(Fragment.from(kids));
  };
  const kids: PMNode[] = [];
  slice.content.forEach((c) => kids.push(map(c)));
  if (!changed) return slice;
  // An open start would merge the first list item into the paragraph at the cursor, which
  // keeps that paragraph's (non-list) settings: the first item would lose its number. An
  // open end would pull the text after the cursor into the last item. So list items at
  // either edge are pasted as paragraphs of their own.
  const isItem = (n: PMNode | undefined) => n?.type === schema.nodes.paragraph && !!n.attrs.numId;
  const openStart = isItem(kids[0]) ? 0 : slice.openStart;
  const openEnd = isItem(kids[kids.length - 1]) ? 0 : slice.openEnd;
  return new Slice(Fragment.from(kids), openStart, openEnd);
}
