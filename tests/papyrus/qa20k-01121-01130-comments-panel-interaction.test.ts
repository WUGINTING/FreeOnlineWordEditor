import { beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { commentedText } from '../../src/papyrus/editor/review';
import type { DocxEditor } from '../../src/papyrus/editor/core';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

async function fixture(options: { comment?: string; author?: string; field?: boolean; quote?: string } = {}): Promise<Uint8Array> {
  const zip = blankPackage();
  const quote = options.quote ?? 'Alpha';
  const content = `<w:p><w:commentRangeStart w:id="1"/><w:r><w:t xml:space="preserve">${esc(quote)}</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> rest of memo</w:t></w:r></w:p>`;
  const field = options.field
    ? '<w:p><w:sdt><w:sdtPr><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:alias w:val="客戶名稱*"/><w:tag w:val="customer"/><w:id w:val="2"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>輸入客戶名稱</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    : '';
  zip.file('word/document.xml', `${HEAD}<w:document ${W}><w:body>${content}${field}${SECT}</w:body></w:document>`);
  zip.file('word/comments.xml', `${HEAD}<w:comments ${W}><w:comment w:id="1" w:author="${esc(options.author ?? 'Ann')}" w:date="2026-01-02T10:00:00Z"><w:p><w:r><w:t>${esc(options.comment ?? 'Old note')}</w:t></w:r></w:p></w:comment></w:comments>`);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
  return zip.generateAsync({ type: 'uint8array' });
}

const flush = async () => { for (let i = 0; i < 12; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };

async function mount(src: Uint8Array | null = null, author = 'Ann') {
  const host = document.createElement('div'); document.body.append(host);
  let editor: DocxEditor | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, author: { name: author }, onReady: (e: DocxEditor) => { editor = e; } }) });
  app.mount(host);
  for (let i = 0; i < 80 && !editor; i++) await flush();
  if (!editor) throw new Error('DocxEditor did not become ready');
  await flush();
  const status = (prefix: string) => Array.from(host.querySelectorAll<HTMLButtonElement>('.dx-status-btn')).find((b) => b.textContent?.trim().startsWith(prefix));
  const panel = () => host.querySelector<HTMLElement>('.dx-comments');
  const buttons = (label: string) => Array.from(host.querySelectorAll<HTMLButtonElement>('.dx-comments button')).filter((b) => b.textContent?.trim() === label);
  const type = async (box: HTMLTextAreaElement, value: string) => { box.value = value; box.dispatchEvent(new Event('input', { bubbles: true })); await nextTick(); };
  const select = (text: string) => {
    let pos = -1;
    editor!.view!.state.doc.descendants((node, at) => { if (pos < 0 && node.isText && node.text!.includes(text)) pos = at + node.text!.indexOf(text); });
    if (pos < 0) throw new Error(`Text not found: ${text}`);
    editor!.view!.dispatch(editor!.view!.state.tr.setSelection(TextSelection.create(editor!.view!.state.doc, pos, pos + text.length)));
  };
  return { host, editor, app, status, panel, buttons, type, select, done: () => { app.unmount(); host.remove(); } };
}

describe('QA20K comments panel user workflows 01121–01130', () => {
  it('QA20K-01121 switches exclusively between the template fields panel and comments panel', async () => {
    const x = await mount(await fixture({ field: true }));
    try {
      expect(x.host.querySelector('.dx-fields')).not.toBeNull();
      expect(x.panel()).toBeNull();
      x.status('留言')!.click(); await flush();
      expect(x.host.querySelector('.dx-fields')).toBeNull();
      expect(x.panel()?.textContent).toContain('Old note');
      x.status('欄位')!.click(); await flush();
      expect(x.panel()).toBeNull();
      expect(x.host.querySelector('.dx-fields')?.textContent).toContain('客戶名稱*');
    } finally { x.done(); }
  });

  it('QA20K-01122 reports message count including replies while keeping replies nested in two discussions', async () => {
    const x = await mount(null, '我');
    try {
      x.editor.view!.dispatch(x.editor.view!.state.tr.insertText('Budget Friday. Delivery Monday.'));
      x.select('Budget'); const rootA = x.editor.addComment('Check budget')!;
      x.editor.replyComment(rootA, 'Finance confirmed');
      x.select('Delivery'); x.editor.addComment('Confirm delivery');
      await flush();
      expect(x.status('留言')?.textContent?.trim()).toBe('留言 3');
      x.status('留言')!.click(); await flush();
      expect(x.panel()?.querySelector('.dx-comments-title')?.textContent).toBe('留言（3）');
      expect(x.panel()?.querySelectorAll(':scope > ol > li')).toHaveLength(2);
      expect(x.panel()?.querySelector('.dx-comment-reply')?.textContent).toContain('Finance confirmed');
    } finally { x.done(); }
  });

  it('QA20K-01123 closes and reopens the comments panel without dirtying or losing the document comments', async () => {
    const x = await mount(await fixture());
    try {
      expect(x.panel()).not.toBeNull();
      expect(x.editor.isModified()).toBe(false);
      x.panel()!.querySelector<HTMLButtonElement>('.dx-comments-close')!.click(); await flush();
      expect(x.panel()).toBeNull();
      expect(x.status('留言')?.getAttribute('aria-pressed')).toBe('false');
      expect(x.editor.isModified()).toBe(false);
      x.status('留言')!.click(); await flush();
      expect(x.panel()?.textContent).toContain('Old note');
      expect(x.editor.view!.state.doc.textContent).toContain('Alpha rest of memo');
      expect(x.editor.isModified()).toBe(false);
    } finally { x.done(); }
  });

  it('QA20K-01124 hides the panel after deleting the only comment and restores it on undo', async () => {
    const x = await mount(await fixture({ author: '我' }), '我');
    try {
      expect(x.panel()).not.toBeNull();
      x.buttons('刪除')[0].click(); await flush();
      expect(x.editor.comments()).toHaveLength(0);
      expect(x.panel()).toBeNull();
      expect(x.status('留言')).toBeUndefined();
      expect(x.editor.isModified()).toBe(true);
      x.editor.undo(); await flush();
      expect(x.editor.comments().map((c) => c.text)).toEqual(['Old note']);
      expect(x.panel()?.textContent).toContain('Old note');
      expect(x.status('留言')?.textContent?.trim()).toBe('留言 1');
      expect(x.editor.isModified()).toBe(false);
    } finally { x.done(); }
  });

  it('QA20K-01125 asks the writer to confirm a refreshed anchor when the body changes during a draft', async () => {
    const x = await mount(null, '我');
    try {
      x.editor.view!.dispatch(x.editor.view!.state.tr.insertText('Payment is due Friday'));
      x.select('Friday');
      Array.from(x.host.querySelectorAll<HTMLButtonElement>('.dx-toolbar button')).find((b) => b.textContent?.trim() === '新增留言')!.click();
      await flush();
      const box = x.panel()!.querySelector<HTMLTextAreaElement>('.dx-comment-draft textarea')!;
      await x.type(box, 'Please verify');
      x.editor.view!.dispatch(x.editor.view!.state.tr.insertText('approved ', 1));
      await flush();
      (box.closest('form') as HTMLFormElement).dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
      await flush();
      expect(x.editor.comments()).toHaveLength(0);
      expect(x.panel()?.querySelector('blockquote')?.textContent).toBe('Friday');
      expect(x.host.textContent).toContain('文件已變更，請確認留言的位置後再送出。');
    } finally { x.done(); }
  });

  it('QA20K-01126 focuses and labels the comment editor when a selected-text draft opens from the toolbar', async () => {
    const x = await mount(null, '我');
    try {
      x.editor.view!.dispatch(x.editor.view!.state.tr.insertText('Please review the total'));
      x.select('the total');
      Array.from(x.host.querySelectorAll<HTMLButtonElement>('.dx-toolbar button')).find((b) => b.textContent?.trim() === '新增留言')!.click();
      await flush();
      const box = x.panel()!.querySelector<HTMLTextAreaElement>('.dx-comment-draft textarea')!;
      expect(document.activeElement).toBe(box);
      expect(box.getAttribute('aria-label')).toBe('新留言');
      expect(x.panel()!.querySelector('.dx-comment-draft blockquote')?.textContent).toBe('the total');
    } finally { x.done(); }
  });

  it('QA20K-01127 sends a reply with Ctrl+Enter to the open thread and closes its reply editor', async () => {
    const x = await mount(await fixture(), '我');
    try {
      x.buttons('回覆')[0].click(); await flush();
      const box = x.panel()!.querySelector<HTMLTextAreaElement>('.dx-comment-reply-box')!;
      await x.type(box, 'Approved for Friday');
      box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }));
      await flush();
      const reply = x.editor.comments().find((c) => c.text === 'Approved for Friday');
      expect(reply).toMatchObject({ parentId: '1', author: '我' });
      expect(x.panel()!.querySelector('.dx-comment-reply-box')).toBeNull();
      expect(x.panel()!.querySelector('.dx-comment-reply')?.textContent).toContain('Approved for Friday');
    } finally { x.done(); }
  });

  it('QA20K-01128 cancels an in-progress edit with Escape and keeps the saved comment unchanged', async () => {
    const x = await mount(await fixture({ author: '我' }), '我');
    try {
      expect(x.editor.isModified()).toBe(false);
      x.buttons('編輯')[0].click(); await flush();
      const box = x.panel()!.querySelector<HTMLTextAreaElement>('.dx-comment-edit-box')!;
      await x.type(box, 'A replacement');
      box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      await flush();
      expect(x.panel()!.querySelector('.dx-comment-edit-box')).toBeNull();
      expect(x.panel()!.querySelector('.dx-comment-text')?.textContent).toBe('Old note');
      expect(x.editor.comments()[0].text).toBe('Old note');
      expect(x.editor.isModified()).toBe(false);
    } finally { x.done(); }
  });

  it('QA20K-01129 shortens a long anchor preview without changing the actual commented range', async () => {
    const x = await mount(null, '我');
    const quote = 'Quarterly   maintenance    handover checklist for the north region with vendor approval before Friday';
    try {
      x.editor.view!.dispatch(x.editor.view!.state.tr.insertText(quote));
      x.select(quote);
      const id = x.editor.addComment('Review the handover')!;
      await flush();
      x.status('留言')!.click(); await flush();
      const normalized = quote.replace(/\s+/g, ' ').trim();
      expect(x.panel()!.querySelector('.dx-comment-main blockquote')?.textContent).toBe(normalized.slice(0, 60) + '…');
      expect(commentedText(x.editor.view!.state.doc, id)).toBe(quote);
    } finally { x.done(); }
  });

  it('QA20K-01130 labels an imported empty legacy comment instead of rendering an indistinguishable blank card', async () => {
    const x = await mount(await fixture({ comment: '' }));
    try {
      expect(x.editor.comments()).toHaveLength(1);
      expect(x.panel()?.querySelector('.dx-comments-title')?.textContent).toBe('留言（1）');
      expect(x.panel()?.querySelector('.dx-comment-text')?.textContent).toBe('（空白留言）');
      expect(x.panel()?.querySelectorAll(':scope > ol > li')).toHaveLength(1);
    } finally { x.done(); }
  });
});
