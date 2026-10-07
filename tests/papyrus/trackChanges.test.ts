// 追蹤修訂 (QA round 8): edits made with tracking on are recorded as Word records them, with the
// author and time; everything recorded can be accepted / rejected with review.ts; each edit is
// one undo step; what can't be recorded is refused with a notice; the switch is saved in
// settings.xml like Word does, and nothing changes in a file opened and saved without edits.
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, NodeSelection, TextSelection, type Command, type Transaction } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { deleteSelection, joinBackward } from 'prosemirror-commands';
import { Fragment, Slice, type Node as PMNode } from 'prosemirror-model';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { deletedInfo, insertionOfLayer, paragraphRevisions } from '../../src/papyrus/docx/revisions';
import { parseLayers } from '../../src/papyrus/docx/wrappers';
import { schema } from '../../src/papyrus/editor/schema';
import {
  acceptAllRevisions, collectRevisions, insertComment, rejectAllRevisions, review,
} from '../../src/papyrus/editor/review';
import { trackChanges, trackTransaction } from '../../src/papyrus/editor/trackChanges';
import {
  clearFormatting, insertHardBreak, insertImage, insertPageBreak, insertTab, insertTable, setAlign, splitParagraph, toggleFormat, toggleList,
} from '../../src/papyrus/editor/commands';
import { CellSelection } from 'prosemirror-tables';
import { documentFields, fieldPlaceholders, lockedFields } from '../../src/papyrus/editor/fields';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { createApp, h, nextTick, ref } from 'vue';
import { loadSfc } from './sfc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const NS_ALL = `${W} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const P = (inner: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
const R = (t: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const NOW = new Date(2026, 8, 24, 10, 30, 12);
const DATE = '2026-09-24T10:30:00Z';
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class {
    observe() {}
    unobserve() {}
    disconnect() {}
  };
  const empty = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= empty;
  Range.prototype.getBoundingClientRect ??= zero;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

async function pack(body: string, opts: { settings?: string; image?: boolean } = {}): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS_ALL}><w:body>${body}${SECT}</w:body></w:document>`);
  if (opts.settings != null) zip.file('word/settings.xml', opts.settings);
  if (opts.image) {
    zip.file('word/media/image1.png', PNG, { base64: true });
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>'));
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    zip.file('[Content_Types].xml', ct.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'));
  }
  return zip.generateAsync({ type: 'uint8array' });
}

const IMAGE_RUN =
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Picture 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg1"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

/** An editor state with the plugins of the real editor that matter here, dispatching like DocxEditor. */
class Ed {
  state: EditorState;
  notices: string[] = [];
  on = true;
  author = 'Tester';
  model: Awaited<ReturnType<typeof readDocx>>['model'];

  constructor(read: Awaited<ReturnType<typeof readDocx>>) {
    this.model = read.model;
    this.state = EditorState.create({
      schema,
      doc: read.doc,
      plugins: [
        history(),
        trackChanges({ enabled: () => this.on, author: () => this.author, onNotice: (m) => this.notices.push(m), now: () => NOW }),
        fieldPlaceholders(),
        lockedFields((f) => this.notices.push(`locked:${f.title}`)),
        review(),
      ],
    });
  }

  static async open(body: string, opts?: { image?: boolean }): Promise<Ed> {
    return new Ed(await readDocx(await pack(body, opts)));
  }

  dispatch = (tr: Transaction) => {
    const t = trackTransaction(this.state, tr);
    if (t) this.state = this.state.apply(t);
  };

  run(cmd: Command): boolean {
    return cmd(this.state, this.dispatch);
  }

  /** Position of `needle` in the text (+ offset). */
  at(needle: string, offset = 0): number {
    let found = -1;
    this.state.doc.descendants((n, pos) => {
      if (found >= 0) return false;
      if (n.isText && n.text!.includes(needle)) found = pos + n.text!.indexOf(needle) + offset;
      return true;
    });
    if (found < 0) throw new Error('not found: ' + needle);
    return found;
  }

  cursor(pos: number): void {
    this.dispatch(this.state.tr.setSelection(TextSelection.create(this.state.doc, pos)));
  }

  select(from: number, to: number): void {
    this.dispatch(this.state.tr.setSelection(TextSelection.create(this.state.doc, from, to)));
  }

  /** Typing, as ProseMirror does it (the stored marks at the cursor). */
  type(text: string): void {
    for (const ch of text) this.dispatch(this.state.tr.insertText(ch));
  }

  /** Backspace: the browser's deletion of the character before the cursor, or the keymap's commands. */
  backspace(): void {
    const { empty, from, $from } = this.state.selection;
    if (!empty) this.run(deleteSelection);
    else if ($from.parentOffset === 0) this.run(joinBackward);
    else this.dispatch(this.state.tr.delete(from - 1, from));
  }

  /** Delete: the character after the cursor. */
  del(): void {
    const { from } = this.state.selection;
    this.dispatch(this.state.tr.delete(from, from + 1));
  }

  show(): string {
    return show(this.state.doc);
  }

  revisions() {
    return collectRevisions(this.state.doc);
  }

  async documentXml(): Promise<string> {
    const bytes = await writeDocx(this.state.doc, this.model);
    return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
  }
}

/** Paragraphs joined by |: inserted text [+..], deleted [-..], paragraph marks ¶+ / ¶-, other revisions {fmt}. */
function show(doc: PMNode): string {
  const out: string[] = [];
  doc.descendants((n) => {
    if (n.type.name !== 'paragraph') return true;
    let s = '';
    n.forEach((c) => {
      const wrap = c.marks.find((m) => m.type.name === 'inlineWrap');
      const ins = wrap && parseLayers(wrap.attrs.layers).some((l) => insertionOfLayer(l.open));
      if (c.type.name === 'raw_inline' && c.attrs.label === 'del') s += `[-${deletedInfo(c.attrs.xml)!.text}]`;
      else if (c.isText) s += ins ? `[+${c.text}]` : c.text;
      else if (c.type.name === 'image') s += ins ? '[+img]' : 'img';
    });
    const mark = paragraphRevisions(n.attrs.pPr).mark;
    out.push(s + (mark ? (mark.kind === 'ins' ? '¶+' : '¶-') : ''));
    return false;
  });
  return out.join('|');
}

describe('追蹤修訂: typing', () => {
  it('records typed text as one insertion with the author and Word\'s date', async () => {
    const ed = await Ed.open(P(R('Hello world')));
    ed.cursor(ed.at('Hello', 5));
    ed.type(' big');
    expect(ed.show()).toBe('Hello[+ big] world');
    const revs = ed.revisions();
    expect(revs.map((r) => [r.kind, r.author, r.date])).toEqual([['ins', 'Tester', DATE]]);
    // Text inserted without the stored marks (another input path) joins the same insertion.
    const end = ed.state.selection.from;
    ed.dispatch(ed.state.tr.insert(end, schema.text('!')));
    expect(ed.show()).toBe('Hello[+ big!] world');
    expect(ed.revisions()).toHaveLength(1);
    const xml = await ed.documentXml();
    expect(xml.match(/<w:ins /g)).toHaveLength(1);
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Tester" w:date="2026-09-24T10:30:00Z"><w:r><w:t xml:space="preserve"> big!<\/w:t><\/w:r><\/w:ins>/);
  });

  it('typed text already carries the insertion (no rewrap, safe for IME composition)', async () => {
    const ed = await Ed.open(P(R('Hello')));
    ed.cursor(ed.at('Hello', 5));
    const tr = ed.state.tr.insertText('字');
    expect(trackTransaction(ed.state, tr)).toBe(tr);
    ed.dispatch(tr);
    expect(ed.show()).toBe('Hello[+字]');
  });

  it('replaces a selection with a deletion followed by the insertion', async () => {
    const ed = await Ed.open(P(R('Hello world')));
    ed.select(ed.at('world'), ed.at('world', 5));
    ed.type('X');
    expect(ed.show()).toBe('Hello [-world][+X]');
    expect(ed.state.selection.from).toBe(ed.at('X', 1));
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('Hello world');
  });

  it('with tracking off, typing is not tracked', async () => {
    const ed = await Ed.open(P(R('Hello')));
    ed.on = false;
    ed.cursor(ed.at('Hello', 5));
    ed.type('!');
    ed.backspace();
    ed.backspace();
    expect(ed.show()).toBe('Hell');
    expect(ed.revisions()).toHaveLength(0);
  });

  it('every edit is one undo step', async () => {
    const ed = await Ed.open(P(R('Hello world')));
    const original = ed.state.doc;
    ed.cursor(ed.at('world', 5));
    ed.backspace();
    expect(ed.show()).toBe('Hello worl[-d]');
    ed.run(undo);
    expect(ed.state.doc.eq(original)).toBe(true);
    ed.run(splitParagraph);
    ed.run(undo);
    expect(ed.state.doc.eq(original)).toBe(true);
  });
});

describe('追蹤修訂: deleting', () => {
  it('Backspace keeps the text as a deletion and moves left over it; deletions join', async () => {
    const ed = await Ed.open(P(R('Hello world')));
    ed.cursor(ed.at('world', 5));
    ed.backspace();
    ed.backspace();
    expect(ed.show()).toBe('Hello wor[-ld]');
    const revs = ed.revisions();
    expect(revs.map((r) => [r.kind, r.author, r.date])).toEqual([['del', 'Tester', DATE]]);
    // The cursor is before the deleted text.
    expect(ed.state.selection.from).toBe(ed.at('wor', 3));
    ed.type('k');
    expect(ed.show()).toBe('Hello wor[+k][-ld]');
    const xml = await ed.documentXml();
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="Tester" w:date="2026-09-24T10:30:00Z"><w:r><w:delText>l<\/w:delText><\/w:r><w:r><w:delText>d<\/w:delText><\/w:r><\/w:del>/);
  });

  it('Delete keeps the text as a deletion and moves past it', async () => {
    const ed = await Ed.open(P(R('Hello world')));
    ed.cursor(ed.at('world'));
    ed.del();
    ed.del();
    expect(ed.show()).toBe('Hello [-wo]rld');
    expect(ed.state.selection.from).toBe(ed.at('rld'));
    ed.run(acceptAllRevisions);
    expect(ed.show()).toBe('Hello rld');
  });

  it('deleting the author\'s own insertion removes it', async () => {
    const ed = await Ed.open(P(R('Hello')));
    ed.cursor(ed.at('Hello', 5));
    ed.type('abc');
    ed.backspace();
    expect(ed.show()).toBe('Hello[+ab]');
    // Another author's insertion is deleted as a tracked deletion inside it.
    ed.author = 'Other';
    ed.backspace();
    expect(ed.show()).toBe('Hello[+a][-b]');
    expect(ed.revisions().map((r) => [r.kind, r.author])).toEqual([['ins', 'Tester'], ['del', 'Other']]);
    expect(await ed.documentXml()).toMatch(/<w:ins [^>]*w:author="Tester"[^>]*><w:r><w:t>a<\/w:t><\/w:r><w:del [^>]*w:author="Other"[^>]*><w:r><w:delText>b<\/w:delText><\/w:r><\/w:del><\/w:ins>/);
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('Hello');
  });

  it('deleting across paragraphs marks the paragraph mark deleted', async () => {
    const ed = await Ed.open(P(R('One')) + P(R('Two')));
    ed.select(ed.at('One', 2), ed.at('Two', 1));
    ed.backspace();
    expect(ed.show()).toBe('On[-e]¶-|[-T]wo');
    expect(ed.revisions().map((r) => r.kind)).toEqual(['del', 'paraMark', 'del']);
    const tracked = ed.state;
    ed.run(acceptAllRevisions);
    expect(ed.show()).toBe('Onwo');
    ed.state = tracked;
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('One|Two');
    const xml = await new Ed({ doc: tracked.doc, model: ed.model }).documentXml();
    expect(xml).toMatch(/<w:pPr><w:rPr><w:del w:id="\d+" w:author="Tester" w:date="2026-09-24T10:30:00Z"\/><\/w:rPr><\/w:pPr>/);
  });

  it('Backspace at the start of a paragraph deletes the mark before it', async () => {
    const ed = await Ed.open(P(R('One')) + P(R('Two')));
    ed.cursor(ed.at('Two'));
    ed.backspace();
    expect(ed.show()).toBe('One¶-|Two');
    expect(ed.state.selection.from).toBe(ed.at('One', 3));
    ed.run(acceptAllRevisions);
    expect(ed.show()).toBe('OneTwo');
  });

  it('pictures from the file are kept as a deletion and come back when rejected', async () => {
    const ed = await Ed.open(P(R('a') + IMAGE_RUN + R('b')), { image: true });
    const pos = ed.at('a', 1);
    ed.dispatch(ed.state.tr.setSelection(NodeSelection.create(ed.state.doc, pos)));
    ed.run(deleteSelection);
    expect(ed.revisions().map((r) => r.kind)).toEqual(['del']);
    expect(await ed.documentXml()).toMatch(/<w:del [^>]*><w:r><w:drawing>[\s\S]*r:embed="rIdImg1"[\s\S]*<\/w:drawing><\/w:r><\/w:del>/);
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('aimgb');
  });
});

describe('追蹤修訂: paragraphs, paste, formatting', () => {
  it('Enter inserts a paragraph mark; Backspace right after removes it again', async () => {
    const ed = await Ed.open(P(R('Hello'), '<w:jc w:val="center"/>'));
    const original = ed.state.doc;
    ed.cursor(ed.at('Hello', 3));
    ed.run(splitParagraph);
    expect(ed.show()).toBe('Hel¶+|lo');
    ed.backspace();
    expect(ed.show()).toBe('Hello');
    expect(ed.state.doc.eq(original)).toBe(true);
    ed.run(splitParagraph);
    const tracked = ed.state;
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('Hello');
    ed.state = tracked;
    ed.run(acceptAllRevisions);
    expect(ed.show()).toBe('Hel|lo');
  });

  it('pasted paragraphs are inserted with their marks', async () => {
    const ed = await Ed.open(P(R('XY')));
    ed.cursor(ed.at('XY', 1));
    const slice = new Slice(Fragment.from([schema.nodes.paragraph.create(null, schema.text('A')), schema.nodes.paragraph.create(null, schema.text('B'))]), 1, 1);
    ed.dispatch(ed.state.tr.replaceSelection(slice).setMeta('uiEvent', 'paste'));
    expect(ed.show()).toBe('X[+A]¶+|[+B]Y');
    const tracked = ed.state;
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('XY');
    ed.state = tracked;
    ed.run(acceptAllRevisions);
    expect(ed.show()).toBe('XA|BY');
  });

  it('bold records the run properties before (w:rPrChange); reject restores them', async () => {
    const ed = await Ed.open(P(R('Hello world', '<w:i/>')));
    ed.select(ed.at('world'), ed.at('world', 5));
    ed.run(toggleFormat('bold'));
    expect(ed.revisions().map((r) => [r.kind, r.author])).toEqual([['format', 'Tester']]);
    const xml = await ed.documentXml();
    expect(xml).toMatch(/<w:rPr><w:b\/><w:i\/><w:rPrChange w:id="\d+" w:author="Tester" w:date="2026-09-24T10:30:00Z"><w:rPr><w:i\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t>world<\/w:t>/);
    const tracked = ed.state;
    ed.run(rejectAllRevisions);
    expect(ed.state.doc.rangeHasMark(ed.at('world'), ed.at('world', 5), schema.marks.bold)).toBe(false);
    ed.state = tracked;
    ed.run(acceptAllRevisions);
    expect(ed.revisions()).toHaveLength(0);
    expect(ed.state.doc.rangeHasMark(ed.at('world'), ed.at('world', 5), schema.marks.bold)).toBe(true);
  });

  it('text formatted differently before gets a record for each part', async () => {
    const ed = await Ed.open(P(R('ab') + R('CD', '<w:b/>')));
    ed.select(ed.at('ab'), ed.at('CD', 2));
    ed.run(toggleFormat('bold'));
    expect(ed.state.doc.rangeHasMark(ed.at('ab'), ed.at('ab', 2), schema.marks.bold)).toBe(true);
    // Only "ab" changed: "CD" was already bold and gets no record.
    const xml = await ed.documentXml();
    expect(xml.match(/<w:rPrChange /g)).toHaveLength(1);
    expect(xml).toMatch(/<w:rPr><w:b\/><w:rPrChange [^>]*><w:rPr\/><\/w:rPrChange><\/w:rPr><w:t>ab<\/w:t>/);
    ed.run(rejectAllRevisions);
    expect(ed.state.doc.rangeHasMark(ed.at('ab'), ed.at('ab', 2), schema.marks.bold)).toBe(false);
    expect(ed.state.doc.rangeHasMark(ed.at('CD'), ed.at('CD', 2), schema.marks.bold)).toBe(true);
  });

  it('paragraph alignment records w:pPrChange; reject restores it', async () => {
    const ed = await Ed.open(P(R('Hello')));
    ed.cursor(ed.at('Hello', 2));
    ed.run(setAlign('center'));
    expect(ed.revisions().map((r) => r.kind)).toEqual(['paraFormat']);
    expect(await ed.documentXml()).toMatch(/<w:pPr><w:jc w:val="center"\/><w:pPrChange w:id="\d+" w:author="Tester" w:date="[^"]+"><w:pPr\/><\/w:pPrChange><\/w:pPr>/);
    ed.run(rejectAllRevisions);
    expect(ed.state.doc.firstChild!.attrs.align).toBe(null);
    expect(ed.revisions()).toHaveLength(0);
  });
});

describe('追蹤修訂: other ways of editing', () => {
  it('Backspace next to a deletion moves over it, and the next one joins it', async () => {
    const ed = await Ed.open(P(R('Hello world')));
    ed.cursor(ed.at('world', 5));
    ed.backspace();
    // The cursor after the deletion again: Backspace only moves before it.
    ed.cursor(ed.at('worl', 4) + 1);
    const doc = ed.state.doc;
    ed.backspace();
    expect(ed.state.doc.eq(doc)).toBe(true);
    expect(ed.state.selection.from).toBe(ed.at('worl', 4));
    ed.backspace();
    expect(ed.show()).toBe('Hello wor[-ld]');
    expect(ed.revisions()).toHaveLength(1);
  });

  it('moving text by drag and drop is a deletion and an insertion', async () => {
    const ed = await Ed.open(P(R('Hello world here')));
    const from = ed.at('world');
    const tr = ed.state.tr;
    const slice = ed.state.doc.slice(from, from + 5);
    tr.delete(from, from + 5);
    tr.replace(1, 1, slice);
    ed.dispatch(tr.setMeta('uiEvent', 'drop'));
    expect(ed.show()).toBe('[+world]Hello [-world] here');
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('Hello world here');
  });

  it('replacing in several places at once (replace all) records each', async () => {
    const ed = await Ed.open(P(R('a cat and a cat')));
    const tr = ed.state.tr;
    const first = ed.at('cat');
    const second = first + 'cat and a '.length;
    tr.replaceWith(second, second + 3, schema.text('dog'));
    tr.replaceWith(first, first + 3, schema.text('dog'));
    ed.dispatch(tr);
    expect(ed.show()).toBe('a [-cat][+dog] and a [-cat][+dog]');
    ed.run(acceptAllRevisions);
    expect(ed.show()).toBe('a dog and a dog');
  });

  it('a page break is inserted as a tracked insertion', async () => {
    const ed = await Ed.open(P(R('Hello')));
    ed.cursor(ed.at('Hello', 2));
    ed.run(insertPageBreak);
    expect(ed.revisions().map((r) => r.kind)).toEqual(['ins']);
    expect(await ed.documentXml()).toMatch(/<w:ins [^>]*><w:r><w:br w:type="page"\/><\/w:r><\/w:ins>/);
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('Hello');
  });

  it('an inserted picture is a tracked insertion', async () => {
    const ed = await Ed.open(P(R('ab')));
    ed.cursor(ed.at('ab', 1));
    ed.run(insertImage('data:image/png;base64,' + PNG, 1, 1, 100));
    expect(ed.show()).toBe('a[+img]b');
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('ab');
  });

  it('clearing formatting records one change with the properties before', async () => {
    const ed = await Ed.open(P(R('Hello', '<w:b/><w:caps/><w:color w:val="FF0000"/>')));
    ed.select(ed.at('Hello'), ed.at('Hello', 5));
    ed.run(clearFormatting);
    expect(ed.revisions().map((r) => r.kind)).toEqual(['format']);
    const xml = await ed.documentXml();
    expect(xml).toMatch(/<w:rPr><w:rPrChange [^>]*><w:rPr><w:b\/><w:caps\/><w:color w:val="FF0000"\/><\/w:rPr><\/w:rPrChange><\/w:rPr><w:t>Hello<\/w:t>/);
    ed.run(rejectAllRevisions);
    expect(await ed.documentXml()).toContain('<w:rPr><w:b/><w:caps/><w:color w:val="FF0000"/></w:rPr><w:t>Hello</w:t>');
  });

  it('a list change records the paragraph properties before', async () => {
    const ed = await Ed.open(P(R('Item')));
    ed.cursor(ed.at('Item', 1));
    ed.run(toggleList('decimal', ed.model.numbering));
    expect(ed.revisions().map((r) => r.kind)).toEqual(['paraFormat']);
    ed.run(rejectAllRevisions);
    expect(ed.state.doc.firstChild!.attrs.numId).toBe(null);
  });

  it('deleting the text of table cells keeps their last paragraph marks', async () => {
    const TBL2 = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>' +
      `<w:tc>${P(R('c1'))}${P(R('c2'))}</w:tc><w:tc>${P(R('d1'))}</w:tc></w:tr></w:tbl>`;
    const ed = await Ed.open(P(R('x')) + TBL2 + P(R('y')));
    const $a = ed.state.doc.resolve(ed.at('c1'));
    const $b = ed.state.doc.resolve(ed.at('d1'));
    ed.dispatch(ed.state.tr.setSelection(new CellSelection(ed.state.doc.resolve($a.before(3)), ed.state.doc.resolve($b.before(3)))));
    ed.run(deleteSelection);
    expect(ed.show()).toBe('x|[-c1]¶-|[-c2]|[-d1]|y');
    ed.run(acceptAllRevisions);
    expect(ed.show()).toBe('x|||y');
  });

  it('tabs and line breaks are tracked insertions', async () => {
    const ed = await Ed.open(P(R('Body')));
    ed.cursor(ed.at('Body', 4));
    ed.run(insertHardBreak);
    ed.run(insertTab);
    expect(ed.revisions().map((r) => r.kind)).toEqual(['ins']);
    ed.run(rejectAllRevisions);
    expect(ed.show()).toBe('Body');
  });
});

describe('追蹤修訂: what it can\'t record is refused, and it works with the other plugins', () => {
  const TBL = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' + P(R('cell')) + '</w:tc></w:tr></w:tbl>';

  it('refuses deleting a table or inserting one, with a notice', async () => {
    const ed = await Ed.open(P(R('before')) + TBL + P(R('after')));
    const doc = ed.state.doc;
    ed.select(ed.at('before', 3), ed.at('after', 2));
    ed.backspace();
    expect(ed.state.doc.eq(doc)).toBe(true);
    expect(ed.notices).toHaveLength(1);
    ed.cursor(ed.at('before', 6));
    ed.run(insertTable(2, 2, 400));
    expect(ed.state.doc.eq(doc)).toBe(true);
    expect(ed.notices).toHaveLength(2);
    // Text in a cell is tracked as usual.
    ed.cursor(ed.at('cell', 4));
    ed.type('s');
    expect(ed.show()).toBe('before|cell[+s]|after');
  });

  it('typing over a field\'s placeholder records only the new text', async () => {
    const sdt = '<w:sdt><w:sdtPr><w:alias w:val="客戶名稱"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>' + R('按一下輸入文字') + '</w:sdtContent></w:sdt>';
    const ed = await Ed.open(P(R('名稱：') + sdt));
    const f = documentFields(ed.state.doc)[0];
    ed.select(f.from, f.to);
    ed.type('王');
    expect(ed.show()).toBe('名稱：[+王]');
    expect(documentFields(ed.state.doc)[0].placeholder).toBe(false);
    expect(ed.revisions().map((r) => r.kind)).toEqual(['ins']);
  });

  it('a locked field stays locked', async () => {
    const sdt = '<w:sdt><w:sdtPr><w:alias w:val="金額"/><w:lock w:val="contentLocked"/><w:text/></w:sdtPr><w:sdtContent>' + R('100') + '</w:sdtContent></w:sdt>';
    const ed = await Ed.open(P(R('A') + sdt));
    const doc = ed.state.doc;
    ed.cursor(ed.at('100', 1));
    ed.type('9');
    ed.backspace();
    expect(ed.state.doc.eq(doc)).toBe(true);
    expect(ed.notices.every((n) => n === 'locked:金額')).toBe(true);
  });

  it('comments are not tracked', async () => {
    const ed = await Ed.open(P(R('Hello')));
    const id = insertComment(ed.state, ed.dispatch, { from: ed.at('Hello'), to: ed.at('Hello', 5) }, {
      text: 'x', author: 'Tester', initials: 'T', date: DATE, dateUtc: null,
    });
    expect(id).not.toBe(null);
    expect(ed.revisions()).toHaveLength(0);
  });

  it('saved and opened again, the revisions are the same; ids are new and unique', async () => {
    const ed = await Ed.open(P(R('One') + `<w:ins w:id="41" w:author="Old" w:date="2026-01-01T00:00:00Z">${R('x')}</w:ins>`) + P(R('Two')));
    ed.cursor(ed.at('One', 3));
    ed.type('!');
    ed.cursor(ed.at('Two', 1));
    ed.backspace();
    ed.backspace();
    ed.select(ed.at('wo'), ed.at('wo', 2));
    ed.run(toggleFormat('bold'));
    const before = ed.revisions().map((r) => `${r.kind}:${r.author}`);
    const bytes = await writeDocx(ed.state.doc, ed.model);
    const again = new Ed(await readDocx(bytes));
    expect(again.revisions().map((r) => `${r.kind}:${r.author}`)).toEqual(before);
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    const ids = [...xml.matchAll(/<w:(?:ins|del|rPrChange|pPrChange) w:id="(\d+)"/g)].map((m) => Number(m[1]));
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids.filter((i) => i !== 41).every((i) => i > 41)).toBe(true);
  });
});

describe('追蹤修訂 in DocxEditor', () => {
  const settings = (inner: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:settings ${W}><w:zoom w:percent="100"/>${inner}<w:defaultTabStop w:val="480"/><w:compat/></w:settings>`;

  async function open(bytes: Uint8Array, opts: Record<string, unknown> = {}) {
    const host = document.createElement('div');
    document.body.append(host);
    const notices: string[] = [];
    const editor = new DocxEditor(host, { author: { name: '王小明' }, onNotice: (m) => notices.push(m), ...opts });
    await editor.open(bytes);
    return { editor, notices };
  }
  const settingsOf = async (blob: Blob) => (await JSZip.loadAsync(blob)).file('word/settings.xml')!.async('string');

  it('starts from w:trackRevisions, and settings.xml is untouched while the switch is', async () => {
    const original = settings('<w:trackRevisions/>');
    const bytes = await pack(P(R('Hello')), { settings: original });
    const { editor } = await open(bytes);
    expect(editor.trackChanges()).toBe(true);
    expect(editor.snapshot()!.trackChanges).toBe(true);
    expect(await settingsOf(await editor.save())).toBe(original);
    // Opened and saved without edits: document.xml is what an untouched save gives.
    const plain = await readDocx(bytes);
    const untouched = await (await JSZip.loadAsync(await writeDocx(plain.doc, plain.model))).file('word/document.xml')!.async('string');
    const saved = await (await JSZip.loadAsync(await editor.save())).file('word/document.xml')!.async('string');
    expect(saved).toBe(untouched);
    expect(editor.isModified()).toBe(false);
    editor.destroy();
  });

  it('turning it off / on writes or removes w:trackRevisions where Word puts it', async () => {
    const { editor } = await open(await pack(P(R('Hello')), { settings: settings('<w:trackRevisions/>') }));
    expect(editor.setTrackChanges(false)).toBe(true);
    expect(editor.isModified()).toBe(true);
    expect(await settingsOf(await editor.save())).not.toContain('trackRevisions');
    editor.setTrackChanges(true);
    expect(editor.isModified()).toBe(false);
    editor.destroy();

    const other = await open(await pack(P(R('Hello')), { settings: settings('') }));
    other.editor.setTrackChanges(true);
    expect(await settingsOf(await other.editor.save())).toContain('<w:zoom w:percent="100"/><w:trackRevisions/><w:defaultTabStop');
    other.editor.destroy();
  });

  it('records edits with the author given to the editor; Ctrl+Shift+E toggles', async () => {
    const { editor } = await open(await pack(P(R('Hello'))));
    expect(editor.trackChanges()).toBe(false);
    const view = editor.view!;
    view.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'E', code: 'KeyE', ctrlKey: true, shiftKey: true, bubbles: true }));
    expect(editor.trackChanges()).toBe(true);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 6)));
    view.dispatch(view.state.tr.insertText('!'));
    const revs = collectRevisions(view.state.doc);
    expect(revs.map((r) => [r.kind, r.author])).toEqual([['ins', '王小明']]);
    expect(revs[0].date).toMatch(/^\d{4}-\d\d-\d\dT\d\d:\d\d:00Z$/);
    editor.destroy();
  });

  it('edits in a header are tracked too', async () => {
    const { editor } = await open(await pack(P(R('Body'))));
    editor.setTrackChanges(true);
    editor.editHeaderFooter('header', 0);
    const hv = editor.activeView!;
    expect(hv).not.toBe(editor.view);
    hv.dispatch(hv.state.tr.insertText('頁首'));
    expect(collectRevisions(hv.state.doc).map((r) => [r.kind, r.author])).toEqual([['ins', '王小明']]);
    expect(collectRevisions(editor.view!.state.doc)).toHaveLength(0);
    const zip = await JSZip.loadAsync(await editor.save());
    const header = Object.keys(zip.files).find((n) => /^word\/header\d*\.xml$/.test(n))!;
    expect(await zip.file(header)!.async('string')).toMatch(/<w:ins w:id="\d+" w:author="王小明" w:date="[^"]+"><w:r><w:t>頁首<\/w:t><\/w:r><\/w:ins>/);
    editor.destroy();
  });

  it('the toolbar button shows and switches 追蹤修訂 (aria-pressed)', async () => {
    const snapshot = ref<unknown>(null);
    const { editor } = await open(await pack(P(R('Hello'))), { onUpdate: (s: unknown) => (snapshot.value = s) });
    const Toolbar = await loadSfc('src/papyrus/vue/DocxToolbar.vue');
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp({ render: () => h(Toolbar, { editor, snapshot: snapshot.value, styles: [] }) });
    app.mount(host);
    await nextTick();
    const button = () => Array.from(host.querySelectorAll('button')).find((b) => b.textContent?.trim() === '追蹤修訂')!;
    expect(button().getAttribute('aria-pressed')).toBe('false');
    expect(button().title).toContain('Ctrl+Shift+E');
    button().click();
    await nextTick();
    expect(editor.trackChanges()).toBe(true);
    expect(button().getAttribute('aria-pressed')).toBe('true');
    button().click();
    await nextTick();
    expect(editor.trackChanges()).toBe(false);
    app.unmount();
    editor.destroy();
  });

  it('read-only: the switch can\'t change and nothing is tracked', async () => {
    const { editor } = await open(await pack(P(R('Hello'))), { editable: false });
    expect(editor.setTrackChanges(true)).toBe(false);
    expect(editor.trackChanges()).toBe(false);
    editor.destroy();
  });

  it('section and page setup changes are refused while tracking', async () => {
    const { editor, notices } = await open(await pack(P(R('Hello'))));
    editor.setTrackChanges(true);
    const doc = editor.view!.state.doc;
    expect(editor.insertSectionBreak()).toBe(false);
    editor.setPageSetup({ ...editor.model.page, marginLeft: 1000 });
    expect(editor.view!.state.doc.eq(doc)).toBe(true);
    expect(notices.length).toBe(2);
    editor.destroy();
  });
});
