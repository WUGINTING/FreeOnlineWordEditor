// The comments panel in DocxEditor.vue: 「新增留言」 (toolbar, panel, Ctrl+Alt+M) opens a box for
// the selection; replies, resolve / reopen, edit and delete of your own comments; nothing to
// change in read-only mode; clicking a comment still selects its text.
import { beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import type { DocxEditor } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

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
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';

async function sample(): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file(
    'word/document.xml',
    `<w:document ${W}><w:body><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Alpha</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> beta gamma</w:t></w:r></w:p>${SECT}</w:body></w:document>`,
  );
  zip.file('word/comments.xml', `<w:comments ${W}><w:comment w:id="1" w:author="Ann" w:date="2026-01-02T10:00:00Z"><w:p><w:r><w:t>Old note</w:t></w:r></w:p></w:comment></w:comments>`);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
  return zip.generateAsync({ type: 'uint8array' });
}

const flush = async () => {
  for (let i = 0; i < 10; i++) {
    await new Promise((r) => setTimeout(r, 0));
    await nextTick();
  }
};

async function mount(props: Record<string, unknown>) {
  const host = document.createElement('div');
  document.body.append(host);
  let editor: DocxEditor | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { ...props, onReady: (e: DocxEditor) => (editor = e) }) });
  app.mount(host);
  for (let i = 0; i < 50 && !editor; i++) await flush();
  await flush();
  const panel = () => host.querySelector('.dx-comments') as HTMLElement | null;
  const buttons = (label: string) => Array.from(host.querySelectorAll('.dx-comments button')).filter((b) => b.textContent!.trim() === label) as HTMLButtonElement[];
  const type = async (box: HTMLTextAreaElement, text: string) => {
    box.value = text;
    box.dispatchEvent(new Event('input'));
    await nextTick();
  };
  return { host, editor: editor!, panel, buttons, type, done: () => (app.unmount(), host.remove()) };
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

describe('comments panel', () => {
  it('adds, replies, resolves, edits and deletes through the panel', async () => {
    const { host, editor, panel, buttons, type, done } = await mount({ src: await sample(), author: { name: '我' } });
    expect(panel()).not.toBeNull(); // opened: the document has a comment
    expect(panel()!.textContent).not.toContain('唯讀');

    // Ctrl+Alt+M on the selection opens a box quoting it.
    select(editor, 'gamma');
    host.querySelector('.dx-vue')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'm', code: 'KeyM', ctrlKey: true, altKey: true, bubbles: true }));
    await flush();
    const draft = panel()!.querySelector('.dx-comment-draft') as HTMLFormElement;
    expect(draft.querySelector('blockquote')!.textContent).toBe('gamma');
    const box = draft.querySelector('textarea')!;
    expect(buttons('留言')[0].disabled).toBe(true);
    await type(box, '  新的留言  ');
    box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }));
    await flush();
    expect(panel()!.querySelector('.dx-comment-draft')).toBeNull();
    expect(editor.comments().map((c) => [c.author, c.text])).toEqual([['Ann', 'Old note'], ['我', '新的留言']]);
    expect(editor.isModified()).toBe(true);

    // Only my comment has 編輯 / 刪除.
    expect(buttons('編輯').length).toBe(1);
    expect(buttons('刪除').length).toBe(1);

    // Reply to Ann's thread.
    buttons('回覆')[0].click();
    await flush();
    const replyBox = panel()!.querySelector('.dx-comment-reply-box') as HTMLTextAreaElement;
    await type(replyBox, '已處理');
    (replyBox.closest('form') as HTMLFormElement).dispatchEvent(new Event('submit'));
    await flush();
    expect(editor.comments().find((c) => c.text === '已處理')).toMatchObject({ parentId: '1', author: '我' });
    expect(panel()!.querySelectorAll('li')[0].querySelectorAll('.dx-comment-reply').length).toBe(1);

    // Resolve, then reopen.
    buttons('標示為已解決')[0].click();
    await flush();
    expect(editor.comments()[0].done).toBe(true);
    expect(panel()!.querySelector('li.done .dx-comment-done')!.textContent).toBe('已解決');
    buttons('重新開啟')[0].click();
    await flush();
    expect(editor.comments()[0].done).toBe(false);

    // Edit my comment.
    const mine = () => Array.from(panel()!.querySelectorAll('li')).find((li) => li.textContent!.includes('gamma'))!;
    (Array.from(mine().querySelectorAll('button')).find((b) => b.textContent === '編輯') as HTMLButtonElement).click();
    await flush();
    const editBox = panel()!.querySelector('.dx-comment-edit-box') as HTMLTextAreaElement;
    expect(editBox.value).toBe('新的留言');
    await type(editBox, '改過的留言');
    (editBox.closest('form') as HTMLFormElement).dispatchEvent(new Event('submit'));
    await flush();
    expect(editor.comments().find((c) => c.author === '我' && !c.parentId)!.text).toBe('改過的留言');

    // Delete it; undo brings it back.
    (Array.from(mine().querySelectorAll('button')).find((b) => b.textContent === '刪除') as HTMLButtonElement).click();
    await flush();
    expect(editor.comments().map((c) => c.text)).toEqual(['Old note', '已處理']);
    editor.undo();
    await flush();
    expect(editor.comments().map((c) => c.text)).toEqual(['Old note', '改過的留言', '已處理']);

    // Clicking a comment selects its text.
    (panel()!.querySelector('.dx-comment-main') as HTMLButtonElement).click();
    const sel = editor.view!.state.selection;
    expect(editor.view!.state.doc.textBetween(sel.from, sel.to)).toBe('Alpha');
    done();
  });

  it('keeps the keyboard in the panel when a comment, reply or edit box closes (GOV-093)', async () => {
    const { host, editor, panel, buttons, type, done } = await mount({ src: await sample(), author: { name: '我' } });
    const focused = () => document.activeElement as HTMLElement;

    // Posting a new comment: the keyboard lands on that comment.
    select(editor, 'gamma');
    host.querySelector('.dx-vue')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'm', code: 'KeyM', ctrlKey: true, altKey: true, bubbles: true }));
    await flush();
    const box = panel()!.querySelector('.dx-comment-draft textarea') as HTMLTextAreaElement;
    await type(box, '新的留言');
    box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }));
    await flush();
    const added = editor.comments().find((c) => c.text === '新的留言')!;
    expect(focused().classList.contains('dx-comment-main')).toBe(true);
    expect(focused().dataset.id).toBe(added.id);

    // Sending a reply: back on that thread's 回覆 button.
    buttons('回覆')[0].click();
    await flush();
    const replyBox = panel()!.querySelector('.dx-comment-reply-box') as HTMLTextAreaElement;
    expect(focused()).toBe(replyBox);
    await type(replyBox, '已處理');
    replyBox.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }));
    await flush();
    expect(focused()).toBe(buttons('回覆')[0]);
    expect(focused().closest('.dx-comment-thread')!.getAttribute('data-id')).toBe('1');

    // Cancelling a reply with Esc: the same.
    buttons('回覆')[1].click();
    await flush();
    (panel()!.querySelector('.dx-comment-reply-box') as HTMLTextAreaElement).dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await flush();
    expect(focused()).toBe(buttons('回覆')[1]);

    // Saving an edit: back on its 編輯 button.
    buttons('編輯')[0].click();
    await flush();
    const editBox = panel()!.querySelector('.dx-comment-edit-box') as HTMLTextAreaElement;
    await type(editBox, '改過的留言');
    (editBox.closest('form') as HTMLFormElement).dispatchEvent(new Event('submit'));
    await flush();
    expect(focused()).toBe(buttons('編輯')[0]);
    done();
  });

  it('goes to the previous / next comment and lists only the open ones (GOV-051, GOV-060)', async () => {
    const { host, editor, panel, buttons, type, done } = await mount({ src: await sample(), author: { name: '我' } });
    const focused = () => document.activeElement as HTMLElement;
    const selected = () => {
      const sel = editor.view!.state.selection;
      return editor.view!.state.doc.textBetween(sel.from, sel.to);
    };
    const nav = (cls: string) => panel()!.querySelector(`.dx-comments-${cls}`) as HTMLButtonElement;
    // A second comment, on 「gamma」.
    select(editor, 'gamma');
    host.querySelector('.dx-vue')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'm', code: 'KeyM', ctrlKey: true, altKey: true, bubbles: true }));
    await flush();
    const box = panel()!.querySelector('.dx-comment-draft textarea') as HTMLTextAreaElement;
    await type(box, '第二則');
    box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }));
    await flush();

    // The cursor on text without a comment: 下一則 goes to the first one, then on.
    select(editor, 'beta');
    await flush();
    nav('next').click();
    await flush();
    expect(selected()).toBe('Alpha');
    expect(focused()).toBe(nav('next')); // the keyboard can press it again
    expect(panel()!.querySelector('li.active')!.textContent).toContain('Old note');
    nav('next').click();
    await flush();
    expect(selected()).toBe('gamma');
    // The last one: 下一則 is off, the keyboard goes to 上一則.
    expect(nav('next').disabled).toBe(true);
    expect(focused()).toBe(nav('prev'));
    nav('prev').click();
    await flush();
    expect(selected()).toBe('Alpha');
    expect(nav('prev').disabled).toBe(true);
    expect(focused()).toBe(nav('next'));

    // 只看未解決: Ann's comment resolved, only mine is listed.
    buttons('標示為已解決')[0].click();
    await flush();
    expect(nav('filter').textContent).toBe('只看未解決（1）');
    nav('filter').click();
    await flush();
    expect(nav('filter').getAttribute('aria-pressed')).toBe('true');
    expect(Array.from(panel()!.querySelectorAll('li')).map((li) => li.textContent!.includes('第二則'))).toEqual([true]);
    // Resolving the last open one: nothing listed, said so; the keyboard stays on the filter.
    buttons('標示為已解決')[0].click();
    await flush();
    expect(panel()!.querySelectorAll('li').length).toBe(0);
    expect(panel()!.textContent).toContain('沒有未解決的留言');
    expect(focused()).toBe(nav('filter'));
    expect(nav('next').disabled).toBe(true);
    // Off again: both listed.
    nav('filter').click();
    await flush();
    expect(panel()!.querySelectorAll('li').length).toBe(2);
    done();
  });

  it('opens the box from the toolbar in a document without comments', async () => {
    const { host, editor, panel, buttons, type, done } = await mount({ src: null });
    expect(panel()).toBeNull();
    editor.view!.dispatch(editor.view!.state.tr.insertText('Hello there'));
    select(editor, 'there');
    const tool = Array.from(host.querySelectorAll('.dx-toolbar button')).find((b) => b.textContent === '新增留言') as HTMLButtonElement;
    tool.click();
    await flush();
    await type(panel()!.querySelector('.dx-comment-draft textarea') as HTMLTextAreaElement, 'hi');
    buttons('留言')[0].click();
    await flush();
    expect(editor.comments()).toMatchObject([{ author: '使用者', text: 'hi' }]);
    // Saved as a package Word opens with the comment.
    const zip = await JSZip.loadAsync(await editor.save());
    expect(await zip.file('word/comments.xml')!.async('string')).toContain('w:author="使用者"');
    done();
  });

  it('shows the comments read-only when editing is not allowed', async () => {
    const { host, editor, panel, buttons, done } = await mount({ src: await sample(), editable: false, author: { name: 'Ann' } });
    expect(panel()!.textContent).toContain('唯讀');
    for (const label of ['＋ 新增留言', '回覆', '標示為已解決', '編輯', '刪除']) expect(buttons(label)).toEqual([]);
    select(editor, 'beta');
    host.querySelector('.dx-vue')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'm', code: 'KeyM', ctrlKey: true, altKey: true, bubbles: true }));
    await flush();
    expect(panel()!.querySelector('.dx-comment-draft')).toBeNull();
    expect(editor.comments().length).toBe(1);
    // Reading still works.
    (panel()!.querySelector('.dx-comment-main') as HTMLButtonElement).click();
    const sel = editor.view!.state.selection;
    expect(editor.view!.state.doc.textBetween(sel.from, sel.to)).toBe('Alpha');
    done();
  });
});
