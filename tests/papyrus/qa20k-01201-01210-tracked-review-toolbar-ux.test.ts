import { beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import type { DocxEditor } from '../../src/papyrus/editor/core';
import { blankPackage } from '../../src/papyrus/docx/template';
import { collectRevisions, reviewSummary } from '../../src/papyrus/editor/review';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const REV = (id: number, author: string) => `w:id="${id}" w:author="${author}" w:date="2026-09-25T09:00:00Z"`;

async function fixture(): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `${HEAD}<w:document ${W}><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">Memo </w:t></w:r><w:ins ${REV(1, 'Kai')}><w:r><w:t>approved</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> today; </w:t></w:r><w:del ${REV(2, 'Mina')}><w:r><w:delText>draft</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> copy.</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>${SECT}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const flush = async () => { for (let i = 0; i < 10; i++) { await new Promise((resolve) => setTimeout(resolve, 0)); await nextTick(); } };

async function mount(src: Uint8Array | null = null) {
  const host = document.createElement('div'); document.body.append(host);
  let editor: DocxEditor | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (value: DocxEditor) => { editor = value; } }) });
  app.mount(host);
  for (let i = 0; i < 100 && !editor; i++) await flush();
  if (!editor) throw new Error('DocxEditor did not become ready');
  await flush();
  const ed = editor;
  const button = (title: string) => {
    const found = [...host.querySelectorAll<HTMLButtonElement>('.dx-toolbar button')].find((item) => item.title.startsWith(title));
    if (!found) throw new Error(`Toolbar button not found: ${title}`);
    return found;
  };
  const reviewGroup = () => host.querySelector<HTMLElement>('.dx-toolbar .dx-review');
  const reviewCount = () => reviewGroup()?.querySelector('.dx-review-label')?.textContent?.trim() ?? null;
  const summary = () => reviewSummary(ed.view!.state);
  return { host, app, ed, button, reviewGroup, reviewCount, summary, done: () => { app.unmount(); host.remove(); } };
}

describe('QA20K tracked review toolbar workflows 01201–01210', () => {
  it('QA20K-01201 starts tracking from the toolbar and marks the next typed text', async () => {
    const x = await mount();
    try {
      const tracking = x.button('追蹤修訂');
      expect(tracking.getAttribute('aria-pressed')).toBe('false');
      tracking.click(); await flush();
      expect(tracking.getAttribute('aria-pressed')).toBe('true');
      x.ed.view!.dispatch(x.ed.view!.state.tr.insertText('Approved: ')); await flush();
      expect(x.ed.snapshot()?.trackChanges).toBe(true);
      expect(collectRevisions(x.ed.view!.state.doc).map((item) => item.kind)).toContain('ins');
      expect(x.reviewCount()).toBe('修訂 1');
    } finally { x.done(); }
  });

  // As in Word's 校閱 › 變更, the review actions are always there, and unavailable (disabled,
  // no count) until the document has a revision.
  const reviewActions = (x: Awaited<ReturnType<typeof mount>>) =>
    [...x.reviewGroup()!.querySelectorAll<HTMLButtonElement>('button')];

  it('QA20K-01202 keeps review actions unavailable in a clean document until a revision exists', async () => {
    const x = await mount();
    try {
      expect(x.reviewCount()).toBeNull();
      expect(reviewActions(x).every((b) => b.disabled)).toBe(true);
      expect(x.host.querySelector('.dx-toolbar [aria-label="追蹤修訂"]')).not.toBeNull();
      x.ed.view!.dispatch(x.ed.view!.state.tr.insertText('Ready')); await flush();
      expect(x.reviewCount()).toBeNull();
      expect(reviewActions(x).every((b) => b.disabled)).toBe(true);
      expect(x.ed.isModified()).toBe(true);
    } finally { x.done(); }
  });

  it('QA20K-01203 shows the two pending changes and enables navigation and bulk decisions', async () => {
    const x = await mount(await fixture());
    try {
      expect(x.reviewCount()).toBe('修訂 2');
      expect(x.summary().count).toBe(2);
      expect(x.button('上一個修訂').disabled).toBe(false);
      expect(x.button('下一個修訂').disabled).toBe(false);
      expect(x.button('接受這部分的所有修訂').disabled).toBe(false);
      expect(x.button('拒絕這部分的所有修訂').disabled).toBe(false);
    } finally { x.done(); }
  });

  it('QA20K-01204 prevents a single-change decision when the caret is outside all revisions', async () => {
    const x = await mount(await fixture());
    try {
      let position = -1;
      x.ed.view!.state.doc.descendants((node, at) => { if (position < 0 && node.isText && node.text!.includes('Second paragraph')) position = at + node.text!.indexOf('Second paragraph'); });
      if (position < 0) throw new Error('Outside-revision paragraph not found');
      x.ed.view!.dispatch(x.ed.view!.state.tr.setSelection(TextSelection.create(x.ed.view!.state.doc, position))); await flush();
      expect(x.reviewCount()).toBe('修訂 2');
      expect(x.button('接受游標處').disabled).toBe(true);
      expect(x.button('拒絕游標處').disabled).toBe(true);
      expect(x.button('下一個修訂').disabled).toBe(false);
    } finally { x.done(); }
  });

  it('QA20K-01205 moves the review cursor to the next change and updates current-item controls', async () => {
    const x = await mount(await fixture());
    try {
      x.button('下一個修訂').click(); await flush();
      expect(x.summary().current?.kind).toBe('ins');
      expect(x.ed.view!.state.doc.textBetween(x.ed.view!.state.selection.from, x.ed.view!.state.selection.to)).toBe('approved');
      expect(x.button('接受游標處').disabled).toBe(false);
      expect(x.reviewCount()).toBe('修訂 2');
    } finally { x.done(); }
  });

  it('QA20K-01206 moves backward from the deletion to the earlier insertion', async () => {
    const x = await mount(await fixture());
    try {
      x.button('下一個修訂').click(); await flush();
      x.button('下一個修訂').click(); await flush();
      expect(x.summary().current?.kind).toBe('del');
      x.button('上一個修訂').click(); await flush();
      expect(x.summary().current?.kind).toBe('ins');
      expect(x.ed.view!.state.doc.textBetween(x.ed.view!.state.selection.from, x.ed.view!.state.selection.to)).toBe('approved');
      expect(x.button('拒絕游標處').disabled).toBe(false);
    } finally { x.done(); }
  });

  it('QA20K-01207 accepts the selected insertion while leaving the unrelated deletion pending', async () => {
    const x = await mount(await fixture());
    try {
      x.button('下一個修訂').click(); await flush();
      x.button('接受游標處').click(); await flush();
      expect(x.ed.view!.state.doc.textContent).toContain('approved');
      expect(collectRevisions(x.ed.view!.state.doc).map((item) => item.kind)).toEqual(['del']);
      expect(x.reviewCount()).toBe('修訂 1');
      expect(x.ed.isModified()).toBe(true);
    } finally { x.done(); }
  });

  it('QA20K-01208 rejects the selected deletion without discarding the accepted insertion', async () => {
    const x = await mount(await fixture());
    try {
      x.button('下一個修訂').click(); await flush();
      x.button('下一個修訂').click(); await flush();
      expect(x.summary().current?.kind).toBe('del');
      x.button('拒絕游標處').click(); await flush();
      expect(x.ed.view!.state.doc.textContent).toContain('approved');
      expect(x.ed.view!.state.doc.textContent).toContain('draft');
      expect(collectRevisions(x.ed.view!.state.doc).map((item) => item.kind)).toEqual(['ins']);
      expect(x.reviewCount()).toBe('修訂 1');
    } finally { x.done(); }
  });

  it('QA20K-01209 accepts all changes, preserving inserted wording and removing deleted wording', async () => {
    const x = await mount(await fixture());
    try {
      x.button('接受這部分的所有修訂').click(); await flush();
      expect(x.ed.view!.state.doc.textContent).toContain('approved');
      expect(x.ed.view!.state.doc.textContent).not.toContain('draft');
      expect(collectRevisions(x.ed.view!.state.doc)).toEqual([]);
      expect(x.reviewCount()).toBeNull();
      expect(reviewActions(x).every((b) => b.disabled)).toBe(true);
    } finally { x.done(); }
  });

  it('QA20K-01210 hides and restores revision marks without changing document content', async () => {
      const x = await mount(await fixture());
      try {
        const before = x.ed.view!.state.doc.toJSON();
        const revisionsBefore = collectRevisions(x.ed.view!.state.doc);
        const revisionSummaryBefore = (({ show, ...summary }) => summary)(x.summary());
        const marks = x.button('隱藏修訂標記');
      expect(marks.getAttribute('aria-pressed')).toBe('true');
      marks.click(); await flush();
      expect(x.ed.view!.dom.classList.contains('dx-rev-off')).toBe(true);
      expect(marks.getAttribute('aria-pressed')).toBe('false');
        expect(x.ed.view!.state.doc.toJSON()).toEqual(before);
        expect(collectRevisions(x.ed.view!.state.doc)).toEqual(revisionsBefore);
        expect((( { show, ...summary }) => summary)(x.summary())).toEqual(revisionSummaryBefore);
        expect(x.ed.isModified()).toBe(false);
      x.button('顯示修訂標記').click(); await flush();
      expect(x.ed.view!.dom.classList.contains('dx-rev-off')).toBe(false);
        expect(x.button('隱藏修訂標記').getAttribute('aria-pressed')).toBe('true');
        expect(x.ed.view!.state.doc.toJSON()).toEqual(before);
        expect(collectRevisions(x.ed.view!.state.doc)).toEqual(revisionsBefore);
        expect((( { show, ...summary }) => summary)(x.summary())).toEqual(revisionSummaryBefore);
        expect(x.ed.isModified()).toBe(false);
    } finally { x.done(); }
  });
});
