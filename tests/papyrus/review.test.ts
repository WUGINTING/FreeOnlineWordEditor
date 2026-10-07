// Tracked changes: accepting / rejecting each kind must give what Word gives (checked against
// Word's own Accept All / Reject All), change nothing else, and be one undo step.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { DOMSerializer } from 'prosemirror-model';
import { history, undo } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { parseXml } from '../../src/papyrus/docx/xml';
import { schema } from '../../src/papyrus/editor/schema';
import {
  acceptAllRevisions, acceptRevision, collectRevisions, commentAtSelection, commentRanges, commentedText,
  goToRevision, rejectAllRevisions, rejectRevision, review, revisionsAtSelection, toggleRevisionMarks, reviewKey,
} from '../../src/papyrus/editor/review';
import { compareElements } from './canonical';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const A = (id: number, author = 'Alice') => `w:id="${id}" w:author="${author}" w:date="2026-01-02T03:04:00Z"`;

export async function openBody(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ schema, doc, plugins: [history(), review()] });
  const run = (cmd: Command) => cmd(state, (tr) => (state = state.apply(tr)));
  const select = (from: number, to = from) => (state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, from, to))));
  const bytes = () => writeDocx(state.doc, model);
  const xml = async () => (await JSZip.loadAsync(await bytes())).file('word/document.xml')!.async('string');
  /** The saved body, one line per paragraph, without the section properties. */
  const paras = async () => {
    const b = /<w:body>([\s\S]*)<w:sectPr/.exec(await xml())![1];
    return b.split(/(?=<w:p[ >])/);
  };
  return { model, run, select, bytes, xml, paras, get state() { return state; } };
}

/** Where a piece of text starts in the document. */
function posOf(state: EditorState, text: string): number {
  let found = -1;
  state.doc.descendants((n, pos) => {
    if (found < 0 && n.isText && n.text!.includes(text)) found = pos + n.text!.indexOf(text);
    return found < 0;
  });
  if (found < 0) throw new Error('not found: ' + text);
  return found;
}

describe('showing tracked changes', () => {
  it('insertions are underlined with author and date, deletions visible but not text', async () => {
    const d = await openBody(`<w:p><w:r><w:t xml:space="preserve">Keep </w:t></w:r><w:ins ${A(1)}><w:r><w:t>added</w:t></w:r></w:ins><w:del ${A(2, 'Bob')}><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>`);
    const dom = DOMSerializer.fromSchema(schema).serializeFragment(d.state.doc.content);
    const div = document.createElement('div');
    div.append(dom);
    const ins = div.querySelector('.dx-rev-ins')!;
    expect(ins.textContent).toBe('added');
    expect(ins.getAttribute('title')).toContain('插入：Alice');
    const del = div.querySelector('.dx-rev-del')!;
    expect(del.getAttribute('data-text')).toBe('gone');
    // In the page for screen readers (persona-300), between hidden 「刪除：」 and 「（刪除結束）」; not editable
    // (an atom) nor copied (the clipboard is written from the document, see p300-revisions).
    expect(del.textContent).toBe('刪除：gone（刪除結束）');
    expect(del.querySelector('.dx-rev-del-text')!.textContent).toBe('gone');
    // Said once, by the hidden text: no role="deletion" as well.
    expect(del.getAttribute('role')).toBeNull();
    expect(del.getAttribute('title')).toContain('刪除：Bob');
    // The deletion is not document text.
    expect(d.state.doc.textContent).toBe('Keep added');
    expect(collectRevisions(d.state.doc).map((r) => [r.kind, r.author])).toEqual([['ins', 'Alice'], ['del', 'Bob']]);
  });

  it('hiding markup is a view setting, not an edit', async () => {
    const d = await openBody(`<w:p><w:ins ${A(1)}><w:r><w:t>x</w:t></w:r></w:ins></w:p>`);
    const before = d.state.doc;
    d.run(toggleRevisionMarks);
    expect(reviewKey.getState(d.state)!.show).toBe(false);
    expect(d.state.doc).toBe(before);
    d.run(toggleRevisionMarks);
    expect(reviewKey.getState(d.state)!.show).toBe(true);
  });

  it('opening and saving a document with revisions changes nothing', async () => {
    const body =
      `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:del ${A(1)}/></w:rPr><w:pPrChange ${A(2)}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr>` +
      `<w:r><w:t xml:space="preserve">a </w:t></w:r><w:ins ${A(3)}><w:r><w:rPr><w:b/><w:rPrChange ${A(4)}><w:rPr/></w:rPrChange></w:rPr><w:t>b</w:t></w:r></w:ins>` +
      `<w:del ${A(5)}><w:r w:rsidDel="00AA"><w:rPr><w:i/></w:rPr><w:delText xml:space="preserve"> c </w:delText></w:r></w:del></w:p><w:p><w:r><w:t>d</w:t></w:r></w:p>`;
    const d = await openBody(body);
    const orig = parseXml(`<w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`).documentElement;
    const saved = parseXml(await d.xml()).documentElement;
    expect(compareElements(orig, saved).join('\n')).toBe('');
  });
});

describe('accept / reject, as Word does it', () => {
  const INS = `<w:p><w:r><w:t xml:space="preserve">Para two </w:t></w:r><w:ins ${A(2)}><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">INSERTED </w:t></w:r></w:ins><w:r><w:t>plain</w:t></w:r></w:p>`;

  it('accepting an insertion drops the w:ins and keeps the text and its formatting', async () => {
    const d = await openBody(INS);
    d.select(posOf(d.state, 'INSERTED') + 2);
    d.run(acceptRevision);
    const xml = await d.xml();
    expect(xml).not.toContain('<w:ins');
    expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">INSERTED </w:t></w:r>');
    expect(d.state.doc.textContent).toBe('Para two INSERTED plain');
  });

  it('rejecting an insertion removes the text', async () => {
    const d = await openBody(INS);
    d.select(posOf(d.state, 'INSERTED') + 2);
    d.run(rejectRevision);
    expect(d.state.doc.textContent).toBe('Para two plain');
    expect(await d.xml()).not.toContain('<w:ins');
  });

  const DEL = `<w:p><w:r><w:t>X</w:t></w:r><w:del ${A(1)}><w:r w:rsidDel="00112233"><w:rPr><w:i/></w:rPr><w:delText xml:space="preserve"> gone </w:delText></w:r><w:r><w:tab/></w:r></w:del><w:r><w:t>Y</w:t></w:r></w:p>`;

  it('accepting a deletion removes the deleted content', async () => {
    const d = await openBody(DEL);
    d.run(goToRevision(1));
    d.run(acceptRevision);
    expect((await d.paras())[0]).toBe('<w:p><w:r><w:t>XY</w:t></w:r></w:p>');
  });

  it('rejecting a deletion turns w:delText back into text with its run properties', async () => {
    const d = await openBody(DEL);
    d.run(goToRevision(1));
    d.run(rejectRevision);
    const xml = await d.xml();
    expect(xml).not.toContain('<w:del ');
    expect(xml).not.toContain('delText');
    expect(xml).not.toContain('rsidDel');
    expect(xml).toContain('<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> gone </w:t></w:r><w:r><w:tab/></w:r>');
    expect(d.state.doc.textContent).toBe('X gone Y');
    // Restored text is ordinary, editable text with its formatting.
    const gone = d.state.doc.nodeAt(posOf(d.state, ' gone '))!;
    expect(gone.marks.map((m) => m.type.name)).toContain('italic');
  });

  it('an insertion that someone else deleted: reject removes it, accept keeps the deletion', async () => {
    const body = `<w:p><w:r><w:t>X</w:t></w:r><w:ins ${A(2, 'U')}><w:del ${A(3, 'V')}><w:r><w:delText>both</w:delText></w:r></w:del></w:ins></w:p>`;
    const r = await openBody(body);
    r.run(rejectAllRevisions);
    expect((await r.paras())[0]).toBe('<w:p><w:r><w:t>X</w:t></w:r></w:p>');

    const a = await openBody(body);
    a.run(acceptAllRevisions);
    expect((await a.paras())[0]).toBe('<w:p><w:r><w:t>X</w:t></w:r></w:p>');

    // Rejecting only the deletion leaves the text as an insertion by U.
    const k = await openBody(body);
    k.select(2, 3);
    expect(revisionsAtSelection(k.state).map((x) => x.kind)).toEqual(['ins', 'del']);
    k.run(goToRevision(1));
    k.run(goToRevision(1));
    expect(revisionsAtSelection(k.state).map((x) => x.kind)).toEqual(['del']);
    k.run(rejectRevision);
    expect(await k.xml()).toContain(`<w:ins ${A(2, 'U')}><w:r><w:t>both</w:t></w:r></w:ins>`);
  });

  const FMT = `<w:p><w:r><w:t xml:space="preserve">Para three </w:t></w:r><w:r><w:rPr><w:b/><w:color w:val="FF0000"/><w:rPrChange ${A(5)}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>bold</w:t></w:r></w:p>`;

  it('accepting a formatting change drops the record and keeps the new formatting', async () => {
    const d = await openBody(FMT);
    d.select(posOf(d.state, 'bold') + 1);
    d.run(acceptRevision);
    expect(await d.xml()).toContain('<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>bold</w:t></w:r>');
  });

  it('rejecting a formatting change restores the old formatting (in the editor too)', async () => {
    const d = await openBody(FMT);
    d.select(posOf(d.state, 'bold') + 1);
    d.run(rejectRevision);
    expect(await d.xml()).toContain('<w:r><w:rPr><w:i/></w:rPr><w:t>bold</w:t></w:r>');
    const marks = d.state.doc.nodeAt(posOf(d.state, 'bold'))!.marks.map((m) => m.type.name);
    expect(marks).toContain('italic');
    expect(marks).not.toContain('bold');
    expect(marks).not.toContain('color');
  });

  const PFMT = `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr><w:pPrChange ${A(1)}><w:pPr><w:jc w:val="right"/><w:ind w:left="720"/></w:pPr></w:pPrChange></w:pPr><w:r><w:t>P</w:t></w:r></w:p>`;

  it('paragraph formatting change: accept keeps, reject restores (and the editor shows it)', async () => {
    const a = await openBody(PFMT);
    a.select(2);
    a.run(acceptRevision);
    expect((await a.paras())[0]).toBe('<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>P</w:t></w:r></w:p>');

    const r = await openBody(PFMT);
    r.select(2);
    r.run(rejectRevision);
    expect((await r.paras())[0]).toBe('<w:p><w:pPr><w:jc w:val="right"/><w:ind w:left="720"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>P</w:t></w:r></w:p>');
    expect(r.state.doc.child(0).attrs.align).toBe('right');
    expect(r.state.doc.child(0).attrs.indLeft).toBe(720);
  });

  const MARK = (kind: 'ins' | 'del', content = '<w:r><w:t>A</w:t></w:r>') =>
    `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:${kind} ${A(1)}/></w:rPr></w:pPr>${content}</w:p>` +
    '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>C</w:t></w:r></w:p>';

  it('accepting a deleted paragraph mark joins with the next paragraph, which keeps its own properties', async () => {
    const d = await openBody(MARK('del'));
    d.run(acceptAllRevisions);
    const p = await d.paras();
    expect(p[0]).toBe('<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>AB</w:t></w:r></w:p>');
    expect(p.length).toBe(2);
  });

  it('rejecting a deleted paragraph mark just drops the record', async () => {
    const d = await openBody(MARK('del'));
    d.run(rejectAllRevisions);
    const p = await d.paras();
    expect(p[0]).toBe('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
    expect(p.length).toBe(3);
  });

  it('inserted paragraph mark: accept drops the record, reject joins', async () => {
    const a = await openBody(MARK('ins'));
    a.run(acceptAllRevisions);
    expect((await a.paras())[0]).toBe('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>');
    const r = await openBody(MARK('ins'));
    r.run(rejectAllRevisions);
    expect((await r.paras())[0]).toBe('<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>AB</w:t></w:r></w:p>');
  });

  it('a whole deleted paragraph (content and mark) disappears on accept', async () => {
    const d = await openBody(MARK('del', `<w:del ${A(2)}><w:r><w:delText>A</w:delText></w:r></w:del>`));
    d.run(acceptAllRevisions);
    const p = await d.paras();
    expect(p).toEqual(['<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>', '<w:p><w:r><w:t>C</w:t></w:r></w:p>']);
  });

  it('the paragraph mark revision is reached with next and acted on alone', async () => {
    const d = await openBody(MARK('del', `<w:ins ${A(2)}><w:r><w:t>A</w:t></w:r></w:ins>`));
    d.run(goToRevision(1));
    expect(revisionsAtSelection(d.state).map((r) => r.kind)).toEqual(['ins']);
    d.run(goToRevision(1));
    expect(revisionsAtSelection(d.state).map((r) => r.kind)).toEqual(['paraMark']);
    d.run(rejectRevision);
    expect(await d.xml()).toContain(`<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:ins ${A(2)}><w:r><w:t>A</w:t></w:r></w:ins></w:p>`);
  });

  it('moves: accept keeps the destination, reject the source, and the move markers go', async () => {
    const body =
      `<w:p><w:moveFromRangeStart w:id="1" w:author="M" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:moveFrom ${A(2, 'M')}><w:r><w:t>moved</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="1"/><w:r><w:t>stay</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>end </w:t></w:r><w:moveToRangeStart w:id="3" w:author="M" w:date="2026-01-01T00:00:00Z" w:name="move1"/><w:moveTo ${A(4, 'M')}><w:r><w:t>moved</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="3"/></w:p>`;
    const a = await openBody(body);
    a.run(acceptAllRevisions);
    expect(await a.paras()).toEqual(['<w:p><w:r><w:t>stay</w:t></w:r></w:p>', '<w:p><w:r><w:t>end moved</w:t></w:r></w:p>']);
    const r = await openBody(body);
    r.run(rejectAllRevisions);
    expect(await r.paras()).toEqual(['<w:p><w:r><w:t>movedstay</w:t></w:r></w:p>', '<w:p><w:r><w:t xml:space="preserve">end </w:t></w:r></w:p>']);
  });

  it('rejecting an insertion keeps comment and bookmark markers inside it', async () => {
    const d = await openBody(`<w:p><w:r><w:t>a</w:t></w:r><w:ins ${A(1)}><w:bookmarkStart w:id="0" w:name="bm"/><w:r><w:t>new</w:t></w:r><w:bookmarkEnd w:id="0"/></w:ins></w:p>`);
    d.run(rejectAllRevisions);
    expect((await d.paras())[0]).toBe('<w:p><w:r><w:t>a</w:t></w:r><w:bookmarkStart w:id="0" w:name="bm"/><w:bookmarkEnd w:id="0"/></w:p>');
  });

  it('revisions written by Word 365 (w16du:dateUtc, declared only on the part root) save and resolve', async () => {
    const U = `w:id="1" w:author="灌庭 吳" w:date="2026-09-24T12:01:00Z" w16du:dateUtc="2026-09-24T04:01:00Z"`;
    const body =
      `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:del ${U}/></w:rPr></w:pPr><w:r><w:t>A</w:t></w:r><w:del ${U}><w:r><w:delText>x</w:delText></w:r></w:del></w:p>` +
      '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>';
    const d = await openBody(body);
    expect((await d.paras())[0]).toContain(`<w:rPr><w:del ${U}/></w:rPr>`);
    expect(collectRevisions(d.state.doc).map((r) => [r.kind, r.author])).toEqual([['del', '灌庭 吳'], ['paraMark', '灌庭 吳']]);
    d.run(acceptAllRevisions);
    expect(await d.paras()).toEqual(['<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>AB</w:t></w:r></w:p>']);
  });

  it('each accept / reject is one undo step, and undo restores the file exactly', async () => {
    const body = INS + DEL + FMT + PFMT + MARK('del');
    const d = await openBody(body);
    const before = await d.xml();
    d.run(acceptAllRevisions);
    expect(collectRevisions(d.state.doc)).toEqual([]);
    d.run(undo);
    expect(await d.xml()).toBe(before);
    d.run(rejectAllRevisions);
    expect(collectRevisions(d.state.doc)).toEqual([]);
    d.run(undo);
    expect(await d.xml()).toBe(before);
  });

  it('next / previous wrap around and select each revision', async () => {
    const d = await openBody(INS + DEL);
    d.run(goToRevision(1));
    expect(d.state.doc.textBetween(d.state.selection.from, d.state.selection.to)).toBe('INSERTED ');
    d.run(goToRevision(1));
    expect(d.state.selection.from).toBe(collectRevisions(d.state.doc)[1].from);
    d.run(goToRevision(1));
    expect(d.state.doc.textBetween(d.state.selection.from, d.state.selection.to)).toBe('INSERTED ');
    d.run(goToRevision(-1));
    expect(d.state.selection.from).toBe(collectRevisions(d.state.doc)[1].from);
  });

  it('accept at a cursor outside any revision does nothing', async () => {
    const d = await openBody(INS);
    d.select(2);
    expect(d.run(acceptRevision)).toBe(false);
  });
});

describe('comment ranges', () => {
  it('finds each comment’s range, text and reference mark', async () => {
    const d = await openBody(
      '<w:p><w:r><w:t xml:space="preserve">Para five </w:t></w:r><w:commentRangeStart w:id="9"/><w:r><w:t>end</w:t></w:r><w:commentRangeEnd w:id="9"/>' +
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="9"/></w:r><w:r><w:t>.</w:t></w:r></w:p>',
    );
    const r = commentRanges(d.state.doc).get('9')!;
    expect(d.state.doc.textBetween(r.from, r.to)).toBe('end');
    expect(commentedText(d.state.doc, '9')).toBe('end');
    expect(d.state.doc.nodeAt(r.ref!)!.attrs.label).toBe('註解');
    d.select(r.from + 1);
    expect(commentAtSelection(d.state)).toBe('9');
    d.select(2);
    expect(commentAtSelection(d.state)).toBe(null);
  });
});
