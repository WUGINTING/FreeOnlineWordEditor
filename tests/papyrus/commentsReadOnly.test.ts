// persona-300 A-2d: a document opened read-only can still take comments when the page allows it (option
// `commenting`): add, reply and resolve work, the text cannot be typed into, and what is saved differs from
// the file only in its comments (the server checks the same, DocxCommentCheck).
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { blankPackage } from '../../src/papyrus/docx/template';

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

async function plain(): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello world here.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

async function open(bytes: Uint8Array, options: Record<string, unknown>) {
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host, { author: { name: '王小明' }, ...options });
  await editor.open(bytes);
  return { editor, done: () => (editor.destroy(), host.remove()) };
}

function select(editor: DocxEditor, text: string) {
  const view = editor.view!;
  let at = -1;
  view.state.doc.descendants((n, pos) => {
    if (at < 0 && n.isText && n.text!.includes(text)) at = pos + n.text!.indexOf(text);
    return at < 0;
  });
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at, at + text.length)));
}

describe('comments while viewing (commenting)', () => {
  it('adds, replies to and resolves comments in read-only mode', async () => {
    const { editor, done } = await open(await plain(), { editable: false, commenting: true });
    expect(editor.editable).toBe(false);
    expect(editor.view!.editable).toBe(false);
    select(editor, 'world');
    const id = editor.addComment('請確認')!;
    expect(id).not.toBeNull();
    expect(editor.replyComment(id, '好')).not.toBeNull();
    expect(editor.resolveComment(id, true)).toBe(true);
    expect(editor.comments().map((c) => [c.author, c.text])).toEqual([['王小明', '請確認'], ['王小明', '好']]);
    expect(editor.isModified()).toBe(true);
    // The text is the same; only comment marks were added.
    expect(editor.view!.state.doc.textContent).toBe('Hello world here.');
    const zip = await JSZip.loadAsync(await editor.save());
    expect(await zip.file('word/comments.xml')!.async('string')).toContain('請確認');
    const doc = await zip.file('word/document.xml')!.async('string');
    expect(doc).toContain('<w:commentRangeStart w:id="0"/>');
    expect(doc.replace(/<[^>]+>/g, '')).toBe('Hello world here.');
    done();
  });

  it('is off by default, and can be turned on and off', async () => {
    const { editor, done } = await open(await plain(), { editable: false });
    select(editor, 'world');
    expect(editor.commentTarget()).toBeNull();
    expect(editor.addComment('x')).toBeNull();
    editor.setCommenting(true);
    expect(editor.addComment('x')).not.toBeNull();
    editor.setCommenting(false);
    expect(editor.replyComment('0', 'y')).toBeNull();
    done();
  });
});
