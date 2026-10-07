// A section's header/footer references (w:headerReference / w:footerReference in its w:sectPr)
// change only by edits that are not undo steps: a header/footer made in the editor once it has
// content (DocxEditor.claimHeaderFooter), the headers 浮水印 makes (DocxEditor.setWatermark) and
// 連結到前一節 (setHeaderFooterLinked). Page setup, 不同的首頁, columns ... are undo steps that
// put back the whole w:sectPr as it was, references included: undone, they would drop a header
// made since (the writer leaves out a part no section refers to, so it would be lost on save) or
// bring back one unlinked since.

import { Plugin } from 'prosemirror-state';
import { isHistoryTransaction } from 'prosemirror-history';
import { sectionRefs, withReference, withoutReference } from '../docx/sections';

const KINDS = ['header', 'footer'] as const;
const TYPES = ['default', 'first', 'even'] as const;

/** `now` with the header/footer references `was` has; null when they are the same already. */
function withReferencesOf(was: string | null, now: string | null): string | null {
  if (was === now || now == null) return null;
  const a = sectionRefs(was);
  const b = sectionRefs(now);
  let xml = now;
  for (const kind of KINDS) {
    for (const type of TYPES) {
      const id = a[kind][type];
      if (id === b[kind][type]) continue;
      xml = id ? withReference(xml, kind, type, id) : withoutReference(xml, kind, type);
    }
  }
  return xml === now ? null : xml;
}

/** `xml` without the references to relationships `known` doesn't know; null when it has none of those. */
function withoutUnknown(xml: string | null, known: (relId: string) => boolean): string | null {
  if (!xml) return null;
  const refs = sectionRefs(xml);
  let out = xml;
  for (const kind of KINDS) {
    for (const type of TYPES) {
      const id = refs[kind][type];
      if (id && !known(id)) out = withoutReference(out, kind, type);
    }
  }
  return out === xml ? null : out;
}

/**
 * Ctrl+Z / Ctrl+Y keep every section's header/footer references as they are: after an undo or
 * redo, a section that is still there refers to the same headers and footers as before it. A
 * section it brings back (an undone section break removal) loses any reference to a
 * header/footer that is gone since (`known` says which relationship ids still name one), so the
 * file never refers to a part it doesn't have.
 */
export function keepSectionReferences(known: (relId: string) => boolean = () => true): Plugin {
  return new Plugin({
    appendTransaction(trs, oldState, newState) {
      if (!trs.some((tr) => tr.docChanged && isHistoryTransaction(tr))) return null;
      const tr = newState.tr;
      const last = withReferencesOf(oldState.doc.attrs.sectPr ?? null, newState.doc.attrs.sectPr ?? null);
      if (last) tr.setDocAttribute('sectPr', last);
      oldState.doc.forEach((node, offset) => {
        const was = node.attrs.sectPr as string | null | undefined;
        if (!was) return;
        // The same paragraph after the undo (none when the undo took it out). A position inside it
        // is followed: undoing a change of its attributes replaces the paragraph's start and end
        // tokens (a ReplaceAroundStep), which "deletes" its start position but not its content.
        let inside = offset + 1;
        for (const t of trs) {
          const r = t.mapping.mapResult(inside, 1);
          if (r.deleted) return;
          inside = r.pos;
        }
        if (inside > newState.doc.content.size) return;
        const $inside = newState.doc.resolve(inside);
        if ($inside.depth < 1) return;
        const pos = $inside.before(1);
        const now = newState.doc.nodeAt(pos);
        if (!now || now.type !== node.type) return;
        const xml = withReferencesOf(was, now.attrs.sectPr ?? null);
        if (xml) tr.setNodeMarkup(pos, undefined, { ...now.attrs, sectPr: xml });
      });
      // Then no reference to a header/footer that is gone, wherever the undo put one.
      const lastNow = tr.doc.attrs.sectPr as string | null;
      const lastKnown = withoutUnknown(lastNow, known);
      if (lastKnown) tr.setDocAttribute('sectPr', lastKnown);
      tr.doc.forEach((node, offset) => {
        const xml = withoutUnknown(node.attrs.sectPr ?? null, known);
        if (xml) tr.setNodeMarkup(offset, undefined, { ...node.attrs, sectPr: xml });
      });
      return tr.docChanged ? tr.setMeta('addToHistory', false) : null;
    },
  });
}
