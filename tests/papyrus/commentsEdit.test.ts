// Editing comments in the web editor: add (on the selection, or the word at the cursor), reply,
// resolve / reopen, edit and delete (own comments only), with undo, "modified" tracking and
// read-only mode. What is saved must be what Word writes: comments.xml, commentsExtended.xml
// (replies via w15:paraIdParent, w15:done), commentsIds.xml / commentsExtensible.xml, their
// relationships and content types; comments nobody touched stay byte for byte as they were.
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { blankPackage } from '../../src/papyrus/docx/template';
import { commentThreads, readComments } from '../../src/papyrus/docx/comments';
import { commentRanges, commentedText } from '../../src/papyrus/editor/review';
import { readDocx } from '../../src/papyrus/docx/reader';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class {
    observe() {}
    unobserve() {}
    disconnect() {}
  };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const NSX =
  W +
  ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"' +
  ' xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function body(inner: string): string {
  return `${HEAD}<w:document ${W}><w:body>${inner}${SECT}</w:body></w:document>`;
}

/** A document without comments. */
async function plain(text = 'Hello world here. Second part.'): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', body(`<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`));
  return zip.generateAsync({ type: 'uint8array' });
}

// A document as Word 365 writes it: two threads, the first with a reply, all four parts.
const COMMENTS =
  `${HEAD}<w:comments ${NSX}>` +
  '<w:comment w:id="0" w:author="Ann" w:date="2026-01-02T10:00:00Z" w:initials="A"><w:p w14:paraId="11111111" w14:textId="77777777"><w:r><w:annotationRef/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>First</w:t></w:r></w:p></w:comment>' +
  '<w:comment w:id="1" w:author="Bob" w:date="2026-01-02T11:00:00Z" w:initials="B"><w:p w14:paraId="22222222" w14:textId="77777777"><w:r><w:annotationRef/></w:r><w:r><w:t>Reply by Bob</w:t></w:r></w:p></w:comment>' +
  '<w:comment w:id="2" w:author="我" w:date="2026-01-02T12:00:00Z" w:initials="我"><w:p w14:paraId="33333333" w14:textId="77777777"><w:r><w:annotationRef/></w:r><w:r><w:t>Mine</w:t></w:r></w:p></w:comment>' +
  '</w:comments>';
const EXTENDED =
  `${HEAD}<w15:commentsEx ${NSX}><w15:commentEx w15:paraId="11111111" w15:done="0"/><w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="0"/><w15:commentEx w15:paraId="33333333" w15:done="0"/></w15:commentsEx>`;
const IDS =
  `${HEAD}<w16cid:commentsIds ${NSX}><w16cid:commentId w16cid:paraId="11111111" w16cid:durableId="0AAAAAAA"/><w16cid:commentId w16cid:paraId="22222222" w16cid:durableId="0BBBBBBB"/><w16cid:commentId w16cid:paraId="33333333" w16cid:durableId="0CCCCCCC"/></w16cid:commentsIds>`;
const EXTENSIBLE =
  `${HEAD}<w16cex:commentsExtensible ${NSX}><w16cex:commentExtensible w16cex:durableId="0AAAAAAA" w16cex:dateUtc="2026-01-02T02:00:00Z"/><w16cex:commentExtensible w16cex:durableId="0BBBBBBB" w16cex:dateUtc="2026-01-02T03:00:00Z"/><w16cex:commentExtensible w16cex:durableId="0CCCCCCC" w16cex:dateUtc="2026-01-02T04:00:00Z"/></w16cex:commentsExtensible>`;
const REF = (id: number) => `<w:r><w:commentReference w:id="${id}"/></w:r>`;

async function withComments(): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file(
    'word/document.xml',
    body(
      '<w:p><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:r><w:t>Alpha</w:t></w:r><w:commentRangeEnd w:id="0"/>' + REF(0) +
        '<w:commentRangeEnd w:id="1"/>' + REF(1) + '<w:r><w:t xml:space="preserve"> beta </w:t></w:r>' +
        '<w:commentRangeStart w:id="2"/><w:r><w:t>gamma</w:t></w:r><w:commentRangeEnd w:id="2"/>' + REF(2) + '</w:p>',
    ),
  );
  zip.file('word/comments.xml', COMMENTS);
  zip.file('word/commentsExtended.xml', EXTENDED);
  zip.file('word/commentsIds.xml', IDS);
  zip.file('word/commentsExtensible.xml', EXTENSIBLE);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
        '<Relationship Id="rIdC2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>' +
        '<Relationship Id="rIdC3" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/>' +
        '<Relationship Id="rIdC4" Type="http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible" Target="commentsExtensible.xml"/></Relationships>',
    ),
  );
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  const o = (p: string, t: string) => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${t}+xml"/>`;
  zip.file('[Content_Types].xml', ct.replace('</Types>', o('comments', 'comments') + o('commentsExtended', 'commentsExtended') + o('commentsIds', 'commentsIds') + o('commentsExtensible', 'commentsExtensible') + '</Types>'));
  return zip.generateAsync({ type: 'uint8array' });
}

async function open(bytes: Uint8Array, options: Record<string, unknown> = {}) {
  const host = document.createElement('div');
  document.body.append(host);
  let changes = 0;
  const editor = new DocxEditor(host, { onChange: () => changes++, ...options });
  await editor.open(bytes);
  return { editor, view: editor.view!, changes: () => changes, done: () => (editor.destroy(), host.remove()) };
}

async function saved(editor: DocxEditor): Promise<JSZip> {
  return JSZip.loadAsync(await editor.save());
}
const part = async (zip: JSZip, name: string) => (await zip.file(name)?.async('string')) ?? null;

/** Select the first occurrence of `text` in the body. */
function select(editor: DocxEditor, text: string, len = text.length) {
  const view = editor.view!;
  let at = -1;
  view.state.doc.descendants((n, pos) => {
    if (at < 0 && n.isText && n.text!.includes(text)) at = pos + n.text!.indexOf(text);
    return at < 0;
  });
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at, at + len)));
}

describe('comments: adding', () => {
  it('adds a comment on the selection, creating every part Word needs', async () => {
    const { editor, view, changes, done } = await open(await plain(), { author: { name: '王小明' } });
    expect(editor.comments()).toEqual([]);
    select(editor, 'world');
    const id = editor.addComment('請確認這個詞');
    expect(id).toBe('0');
    expect(editor.comments()).toMatchObject([{ id: '0', author: '王小明', initials: '王', text: '請確認這個詞', parentId: null, done: false }]);
    expect(commentedText(view.state.doc, '0')).toBe('world');
    expect(editor.isModified()).toBe(true);
    expect(changes()).toBeGreaterThan(0);
    // Word's date: local time digits with "Z"; the real UTC time goes to commentsExtensible.
    const c = editor.comments()[0];
    expect(c.date).toMatch(/^\d{4}-\d\d-\d\dT\d\d:\d\d:00Z$/);
    expect(c.dateUtc).toMatch(/^\d{4}-\d\d-\d\dT\d\d:\d\d:00Z$/);

    const zip = await saved(editor);
    const doc = (await part(zip, 'word/document.xml'))!;
    expect(doc).toMatch(/<w:commentRangeStart w:id="0"\/><w:r><w:t(?: xml:space="preserve")?>world<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:commentReference w:id="0"\/><\/w:r>/);
    const comments = (await part(zip, 'word/comments.xml'))!;
    expect(comments).toContain('w:author="王小明"');
    expect(comments).toContain('<w:annotationRef/>');
    expect(await part(zip, 'word/commentsExtended.xml')).toMatch(/<w15:commentEx w15:paraId="[0-7][0-9A-F]{7}" w15:done="0"\/>/);
    expect(await part(zip, 'word/commentsIds.xml')).toMatch(/<w16cid:commentId w16cid:paraId="[0-7][0-9A-F]{7}" w16cid:durableId="[0-7][0-9A-F]{7}"\/>/);
    expect(await part(zip, 'word/commentsExtensible.xml')).toContain(`w16cex:dateUtc="${c.dateUtc}"`);
    const rels = (await part(zip, 'word/_rels/document.xml.rels'))!;
    for (const t of ['relationships/comments"', '2011/relationships/commentsExtended"', '2016/09/relationships/commentsIds"', '2018/08/relationships/commentsExtensible"']) {
      expect(rels).toContain(t);
    }
    const ct = (await part(zip, '[Content_Types].xml'))!;
    for (const p of ['comments', 'commentsExtended', 'commentsIds', 'commentsExtensible']) {
      expect(ct).toContain(`PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p}+xml"`);
    }
    // Read back as a document.
    expect((await readComments(zip)).map((x) => [x.id, x.author, x.text, x.dateUtc])).toEqual([['0', '王小明', '請確認這個詞', c.dateUtc]]);
    const again = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(commentedText(again.doc, '0')).toBe('world');
    done();
  });

  it('comments on the word at the cursor when nothing is selected (as Word does)', async () => {
    const { editor, view, done } = await open(await plain());
    select(editor, 'orld', 0); // cursor inside "world"
    const id = editor.addComment('x')!;
    expect(commentedText(view.state.doc, id)).toBe('world');
    expect(editor.comments()[0].author).toBe('使用者');
    done();
  });

  it('undo removes the comment in one step, redo brings it back; undoing everything is "unmodified"', async () => {
    const { editor, view, done } = await open(await plain());
    const before = view.state.doc;
    select(editor, 'Second');
    editor.addComment('note');
    expect(view.state.doc.textContent).toBe(before.textContent);
    editor.undo();
    expect(editor.comments()).toEqual([]);
    expect(commentRanges(view.state.doc).size).toBe(0);
    expect(editor.isModified()).toBe(false);
    editor.redo();
    expect(editor.comments().map((c) => c.text)).toEqual(['note']);
    expect(commentedText(view.state.doc, '0')).toBe('Second');
    done();
  });

  it('refuses everything in read-only mode, and keeps showing the comments', async () => {
    const { editor, done } = await open(await withComments(), { editable: false, author: { name: '我' } });
    expect(editor.comments().map((c) => c.id)).toEqual(['0', '1', '2']);
    select(editor, 'beta');
    expect(editor.addComment('x')).toBeNull();
    expect(editor.replyComment('0', 'x')).toBeNull();
    expect(editor.resolveComment('0', true)).toBe(false);
    expect(editor.editComment('2', 'x')).toBe(false);
    expect(editor.deleteComment('2')).toBe(false);
    expect(editor.canChangeComment('2')).toBe(false);
    expect(editor.isModified()).toBe(false);
    done();
  });
});

describe('comments: saving what was not touched', () => {
  it('keeps every comment part byte for byte when no comment changed', async () => {
    const src = await withComments();
    const { editor, view, done } = await open(src);
    // An ordinary edit elsewhere.
    view.dispatch(view.state.tr.insertText('!', view.state.doc.content.size - 2));
    const zip = await saved(editor);
    for (const [name, xml] of [['comments', COMMENTS], ['commentsExtended', EXTENDED], ['commentsIds', IDS], ['commentsExtensible', EXTENSIBLE]]) {
      expect(await part(zip, `word/${name}.xml`)).toBe(xml);
    }
    done();
  });
});

describe('comments: threads', () => {
  it('replies with w15:paraIdParent, and keeps the other comments byte for byte', async () => {
    const { editor, view, done } = await open(await withComments(), { author: { name: 'Carol', initials: 'CC' } });
    const id = editor.replyComment('2', 'Agreed')!;
    expect(id).toBe('3');
    expect(commentThreads(editor.comments()).map((t) => [t.comment.id, t.replies.map((r) => r.id)])).toEqual([['0', ['1']], ['2', ['3']]]);
    // The reply covers the same text, like Word's.
    expect(commentedText(view.state.doc, '3')).toBe('gamma');
    const zip = await saved(editor);
    const comments = (await part(zip, 'word/comments.xml'))!;
    // Untouched comments are exactly as they were.
    for (const m of COMMENTS.match(/<w:comment [\s\S]*?<\/w:comment>/g)!) expect(comments).toContain(m);
    expect(comments).toMatch(/<w:comment w:id="3" w:author="Carol" w:date="[^"]+" w:initials="CC">/);
    const ext = (await part(zip, 'word/commentsExtended.xml'))!;
    expect(ext.startsWith(EXTENDED.slice(0, EXTENDED.indexOf('</w15:commentsEx>')))).toBe(true);
    expect(ext).toMatch(/<w15:commentEx w15:paraId="[0-7][0-9A-F]{7}" w15:paraIdParent="33333333" w15:done="0"\/><\/w15:commentsEx>$/);
    const ids = (await part(zip, 'word/commentsIds.xml'))!;
    expect(ids.startsWith(IDS.slice(0, IDS.indexOf('</w16cid:commentsIds>')))).toBe(true);
    const cex = (await part(zip, 'word/commentsExtensible.xml'))!;
    expect(cex.startsWith(EXTENSIBLE.slice(0, EXTENSIBLE.indexOf('</w16cex:commentsExtensible>')))).toBe(true);
    // The reply's paraId, durableId and time are linked across the parts.
    const para = /w15:paraId="([0-9A-F]{8})" w15:paraIdParent="33333333"/.exec(ext)![1];
    expect(comments).toContain(`w14:paraId="${para}"`);
    const durable = new RegExp(`w16cid:paraId="${para}" w16cid:durableId="([0-9A-F]{8})"`).exec(ids)![1];
    expect(cex).toContain(`w16cex:durableId="${durable}"`);
    expect((await readComments(zip)).find((c) => c.id === '3')).toMatchObject({ parentId: '2', author: 'Carol', text: 'Agreed' });
    done();
  });

  it('resolves and reopens a thread with w15:done, undoably', async () => {
    const { editor, done } = await open(await withComments());
    expect(editor.resolveComment('1', true)).toBe(true); // a reply resolves its thread
    expect(editor.comments().find((c) => c.id === '0')!.done).toBe(true);
    expect(editor.isModified()).toBe(true);
    let zip = await saved(editor);
    expect(await part(zip, 'word/commentsExtended.xml')).toBe(EXTENDED.replace('w15:paraId="11111111" w15:done="0"', 'w15:paraId="11111111" w15:done="1"'));
    expect(await part(zip, 'word/comments.xml')).toBe(COMMENTS);
    expect((await readComments(zip)).map((c) => c.done)).toEqual([true, false, false]);
    editor.undo();
    expect(editor.comments()[0].done).toBe(false);
    expect(editor.isModified()).toBe(false);
    editor.redo();
    expect(editor.resolveComment('0', false)).toBe(true);
    zip = await saved(editor);
    expect(await part(zip, 'word/commentsExtended.xml')).toBe(EXTENDED);
    done();
  });
});

describe('comments: editing and deleting your own', () => {
  it('only the author may edit or delete a comment', async () => {
    const { editor, done } = await open(await withComments(), { author: { name: '我' } });
    expect(editor.canChangeComment('2')).toBe(true);
    expect(editor.canChangeComment('0')).toBe(false);
    expect(editor.editComment('0', 'hijack')).toBe(false);
    expect(editor.deleteComment('1')).toBe(false);
    expect(editor.isModified()).toBe(false);
    done();
  });

  it('edits the text; the comment keeps its attributes, id and paraId, the others stay byte for byte', async () => {
    const { editor, done } = await open(await withComments(), { author: { name: '我' } });
    expect(editor.editComment('2', 'Mine, edited\nsecond line')).toBe(true);
    const zip = await saved(editor);
    const comments = (await part(zip, 'word/comments.xml'))!;
    const [c0, c1] = COMMENTS.match(/<w:comment [\s\S]*?<\/w:comment>/g)!;
    expect(comments).toContain(c0);
    expect(comments).toContain(c1);
    expect(comments).toContain('<w:comment w:id="2" w:author="我" w:date="2026-01-02T12:00:00Z" w:initials="我">');
    expect(comments).toMatch(/<w:t xml:space="preserve">second line<\/w:t><\/w:r><\/w:p><\/w:comment>/);
    // The last paragraph keeps the paraId the other parts refer to.
    expect(comments).toMatch(/<w:p w14:paraId="33333333"[^>]*>(?:(?!<w:p[ >]).)*second line/);
    expect(await part(zip, 'word/commentsExtended.xml')).toBe(EXTENDED);
    expect(await part(zip, 'word/commentsIds.xml')).toBe(IDS);
    expect((await readComments(zip)).find((c) => c.id === '2')!.text).toBe('Mine, edited\nsecond line');
    done();
  });

  it('deleting the root of a thread deletes its replies, their ranges and every part entry', async () => {
    const { editor, view, done } = await open(await withComments(), { author: { name: 'Ann' } });
    expect(editor.deleteComment('0')).toBe(true);
    expect(editor.comments().map((c) => c.id)).toEqual(['2']);
    expect([...commentRanges(view.state.doc).keys()]).toEqual(['2']);
    expect(view.state.doc.textContent).toBe('Alpha beta gamma');
    const zip = await saved(editor);
    const doc = (await part(zip, 'word/document.xml'))!;
    expect(doc).not.toMatch(/w:id="[01]"/);
    const comments = (await part(zip, 'word/comments.xml'))!;
    expect(comments).toBe(COMMENTS.replace(/<w:comment w:id="0"[\s\S]*?<\/w:comment><w:comment w:id="1"[\s\S]*?<\/w:comment>/, ''));
    expect(await part(zip, 'word/commentsExtended.xml')).toBe(EXTENDED.replace(/<w15:commentEx w15:paraId="11111111"[^>]*\/><w15:commentEx w15:paraId="22222222"[^>]*\/>/, ''));
    expect(await part(zip, 'word/commentsIds.xml')).toBe(IDS.replace(/<w16cid:commentId w16cid:paraId="11111111"[^>]*\/><w16cid:commentId w16cid:paraId="22222222"[^>]*\/>/, ''));
    expect(await part(zip, 'word/commentsExtensible.xml')).toBe(EXTENSIBLE.replace(/<w16cex:commentExtensible w16cex:durableId="0AAAAAAA"[^>]*\/><w16cex:commentExtensible w16cex:durableId="0BBBBBBB"[^>]*\/>/, ''));
    editor.undo();
    expect(editor.comments().map((c) => c.id)).toEqual(['0', '1', '2']);
    expect(commentedText(view.state.doc, '1')).toBe('Alpha');
    expect(editor.isModified()).toBe(false);
    done();
  });

  it('deleting a comment also removes its markers from inside a tracked deletion', async () => {
    const zip = blankPackage();
    const A = 'w:author="X" w:date="2026-01-02T03:04:00Z"';
    zip.file(
      'word/document.xml',
      body(
        `<w:p><w:r><w:t>a</w:t></w:r><w:del w:id="1" ${A}><w:commentRangeStart w:id="7"/><w:r><w:delText>gone</w:delText></w:r></w:del>` +
          '<w:r><w:t>kept</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>',
      ),
    );
    zip.file('word/comments.xml', `${HEAD}<w:comments ${NSX}><w:comment w:id="7" w:author="我"><w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment></w:comments>`);
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
    const { editor, view, done } = await open(await zip.generateAsync({ type: 'uint8array' }), { author: { name: '我' } });
    expect(editor.deleteComment('7')).toBe(true);
    const doc = (await part(await saved(editor), 'word/document.xml'))!;
    expect(doc).not.toMatch(/comment(RangeStart|RangeEnd|Reference)/);
    expect(doc).toContain(`<w:del w:id="1" ${A}><w:r><w:delText>gone</w:delText></w:r></w:del>`);
    editor.undo();
    expect(editor.isModified()).toBe(false);
    done();
  });

  it('deleting a reply keeps the rest of its thread', async () => {
    const { editor, done } = await open(await withComments(), { author: { name: 'Bob' } });
    expect(editor.deleteComment('1')).toBe(true);
    expect(editor.comments().map((c) => c.id)).toEqual(['0', '2']);
    const zip = await saved(editor);
    expect((await readComments(zip)).map((c) => c.id)).toEqual(['0', '2']);
    done();
  });

  it('a new comment after deleting the highest id is written as new, not as the old one', async () => {
    const { editor, done } = await open(await withComments(), { author: { name: '我' } });
    editor.deleteComment('2');
    select(editor, 'beta');
    const id = editor.addComment('fresh')!;
    const zip = await saved(editor);
    const list = await readComments(zip);
    expect(list.find((c) => c.id === id)).toMatchObject({ author: '我', text: 'fresh', parentId: null, done: false });
    expect(list.map((c) => c.text)).not.toContain('Mine');
    done();
  });
});
