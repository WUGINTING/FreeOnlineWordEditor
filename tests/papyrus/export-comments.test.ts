// Writes documents whose comments were added, replied to, resolved, edited or deleted in the
// editor, each with what Word must show (<case>.expected.json), so Word itself can check them
// (scripts/inspect-comments-in-word.ps1). Opt-in:
//   DOCX_EXPORT=<out> [WORD_COMMENTS_DOCX=<a .docx with comments saved by Word>] npx vitest run tests/papyrus/export-comments.test.ts
// WORD_COMMENTS_DOCX: the same edits are also made on that file ("word-*" cases): its first
// comment gets a reply, its second is resolved, and a new comment is added.
import { beforeAll, describe, it } from 'vitest';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { blankPackage } from '../../src/papyrus/docx/template';
import { threadRoot } from '../../src/papyrus/docx/comments';
import { commentedText } from '../../src/papyrus/editor/review';

const out = process.env.DOCX_EXPORT;
const wordFile = process.env.WORD_COMMENTS_DOCX;

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
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const REF = (id: number) => `<w:r><w:commentReference w:id="${id}"/></w:r>`;

/** No comments at all. */
async function plain(): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `${HEAD}<w:document ${W}><w:body><w:p><w:r><w:t xml:space="preserve">第一段文字 first paragraph.</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Second paragraph with more words.</w:t></w:r></w:p>${SECT}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** Two threads (Ann's with Bob's reply, and 我's), with every part Word 365 writes. */
async function withComments(): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file(
    'word/document.xml',
    `${HEAD}<w:document ${W}><w:body><w:p><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:r><w:t>Alpha</w:t></w:r><w:commentRangeEnd w:id="0"/>${REF(0)}` +
      `<w:commentRangeEnd w:id="1"/>${REF(1)}<w:r><w:t xml:space="preserve"> beta </w:t></w:r><w:commentRangeStart w:id="2"/><w:r><w:t>gamma</w:t></w:r><w:commentRangeEnd w:id="2"/>${REF(2)}` +
      `<w:r><w:t xml:space="preserve"> delta.</w:t></w:r></w:p>${SECT}</w:body></w:document>`,
  );
  const c = (id: number, who: string, para: string, text: string) =>
    `<w:comment w:id="${id}" w:author="${who}" w:date="2026-01-02T10:00:00Z" w:initials="${who[0]}"><w:p w14:paraId="${para}" w14:textId="77777777"><w:r><w:annotationRef/></w:r><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
  zip.file('word/comments.xml', `${HEAD}<w:comments ${NSX}>${c(0, 'Ann', '11111111', 'First')}${c(1, 'Bob', '22222222', 'Reply by Bob')}${c(2, '我', '33333333', 'Mine')}</w:comments>`);
  zip.file('word/commentsExtended.xml', `${HEAD}<w15:commentsEx ${NSX}><w15:commentEx w15:paraId="11111111" w15:done="0"/><w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="0"/><w15:commentEx w15:paraId="33333333" w15:done="0"/></w15:commentsEx>`);
  zip.file('word/commentsIds.xml', `${HEAD}<w16cid:commentsIds ${NSX}><w16cid:commentId w16cid:paraId="11111111" w16cid:durableId="0AAAAAAA"/><w16cid:commentId w16cid:paraId="22222222" w16cid:durableId="0BBBBBBB"/><w16cid:commentId w16cid:paraId="33333333" w16cid:durableId="0CCCCCCC"/></w16cid:commentsIds>`);
  zip.file('word/commentsExtensible.xml', `${HEAD}<w16cex:commentsExtensible ${NSX}><w16cex:commentExtensible w16cex:durableId="0AAAAAAA" w16cex:dateUtc="2026-01-02T02:00:00Z"/><w16cex:commentExtensible w16cex:durableId="0BBBBBBB" w16cex:dateUtc="2026-01-02T02:00:00Z"/><w16cex:commentExtensible w16cex:durableId="0CCCCCCC" w16cex:dateUtc="2026-01-02T02:00:00Z"/></w16cex:commentsExtensible>`);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      rel('rIdC1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'comments.xml') +
        rel('rIdC2', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'commentsExtended.xml') +
        rel('rIdC3', 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds', 'commentsIds.xml') +
        rel('rIdC4', 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible', 'commentsExtensible.xml') +
        '</Relationships>',
    ),
  );
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  const o = (p: string) => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p}+xml"/>`;
  zip.file('[Content_Types].xml', ct.replace('</Types>', o('comments') + o('commentsExtended') + o('commentsIds') + o('commentsExtensible') + '</Types>'));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

function select(editor: DocxEditor, text: string, len = text.length) {
  const view = editor.view!;
  let at = -1;
  view.state.doc.descendants((n, pos) => {
    if (at < 0 && n.isText && n.text!.includes(text)) at = pos + n.text!.indexOf(text);
    return at < 0;
  });
  if (at < 0) throw new Error(`"${text}" not found`);
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at, at + len)));
}

/** What Word must list for the editor's comments (compared without regard to order). */
function expected(editor: DocxEditor) {
  const list = editor.comments();
  const doc = editor.view!.state.doc;
  const text = (id: string) => list.find((c) => c.id === id)!.text.replace(/\n/g, '\r');
  return list.map((c) => {
    const root = threadRoot(list, c.id)!;
    return {
      author: c.author,
      text: text(c.id),
      scope: commentedText(doc, c.id).replace(/\n/g, '\r'),
      done: root.done,
      parent: root === c ? null : text(root.id),
    };
  });
}

type Edit = (editor: DocxEditor) => void;

const CASES: Record<string, { src: () => Promise<Uint8Array>; author: string; edit: Edit }> = {
  // A document that had no comments: two comments (one on the word at the cursor) and a reply.
  'add-to-none': {
    src: plain,
    author: '王小明',
    edit: (e) => {
      select(e, 'first paragraph');
      e.addComment('請確認這一段');
      select(e, 'more', 0); // cursor in "more": the word
      const id = e.addComment('第二則\n兩行')!;
      e.replyComment(id, '好的');
    },
  },
  'add-to-existing': {
    src: withComments,
    author: '王小明',
    edit: (e) => {
      select(e, 'delta');
      e.addComment('New one');
    },
  },
  reply: { src: withComments, author: 'Carol', edit: (e) => e.replyComment('2', 'Agreed') },
  'reply-to-reply-thread': { src: withComments, author: 'Carol', edit: (e) => e.replyComment('1', 'Me too') },
  resolve: { src: withComments, author: 'Carol', edit: (e) => e.resolveComment('0', true) },
  reopen: {
    src: withComments,
    author: 'Carol',
    edit: (e) => {
      e.resolveComment('2', true);
      e.resolveComment('2', false);
    },
  },
  edit: { src: withComments, author: '我', edit: (e) => e.editComment('2', 'Mine, edited\nwith a second line') },
  'delete-root': { src: withComments, author: 'Ann', edit: (e) => e.deleteComment('0') },
  'delete-reply': { src: withComments, author: 'Bob', edit: (e) => e.deleteComment('1') },
  'delete-all': {
    src: withComments,
    author: 'Ann',
    edit: (e) => {
      e.deleteComment('0');
      e.setAuthor({ name: '我' });
      e.deleteComment('2');
    },
  },
  untouched: { src: withComments, author: 'Carol', edit: () => {} },
};

async function run(name: string, src: Uint8Array, author: string, edit: Edit): Promise<void> {
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host, { author: { name: author } });
  await editor.open(src);
  edit(editor);
  // jsdom's Blob has no arrayBuffer(): read it with JSZip.
  const bytes = await (await JSZip.loadAsync(await editor.save())).generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  writeFileSync(join(out!, `${name}-src.docx`), src);
  writeFileSync(join(out!, `${name}.docx`), bytes);
  writeFileSync(join(out!, `${name}.expected.json`), JSON.stringify(expected(editor), null, 1));
  editor.destroy();
  host.remove();
}

describe.skipIf(!out)('export documents with comments edited in the editor', () => {
  it('writes them', async () => {
    mkdirSync(out!, { recursive: true });
    for (const [name, c] of Object.entries(CASES)) await run(name, await c.src(), c.author, c.edit);
    if (wordFile) {
      const src = new Uint8Array(readFileSync(wordFile));
      const first = (e: DocxEditor) => e.comments().find((c) => !c.parentId)!.id;
      const second = (e: DocxEditor) => e.comments().filter((c) => !c.parentId)[1]?.id ?? first(e);
      await run('word-reply', src, 'Carol', (e) => e.replyComment(first(e), 'Reply from the web'));
      await run('word-resolve', src, 'Carol', (e) => e.resolveComment(second(e), true));
      await run('word-add', src, 'Carol', (e) => {
        const doc = e.view!.state.doc;
        e.view!.dispatch(e.view!.state.tr.setSelection(TextSelection.create(doc, doc.content.size - 3)));
        e.addComment('Added on the web');
      });
      await run('word-delete', src, 'Carol', (e) => {
        const c = e.comments().find((x) => !x.parentId)!;
        e.setAuthor({ name: c.author });
        e.deleteComment(c.id);
      });
      await run('word-untouched', src, 'Carol', () => {});
    }
  }, 120_000);
});
