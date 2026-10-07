import { beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { TextSelection } from 'prosemirror-state';
import { blankPackage } from '../src/papyrus/docx/template';
import { DocxEditor } from '../src/papyrus/editor/core';
import type { DocxEditor as DocxEditorType } from '../src/papyrus/editor/core';
import DocxEditorVue from '../src/papyrus/vue/DocxEditor.vue';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const para = (text: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const heading = (text: string, style = 'Heading1') => para(text, `<w:pStyle w:val="${style}"/>`);
const field = (bookmark: string, value: string) => `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF ${bookmark} \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${value}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
const styles = (extra = '') => `<w:styles ${W}><w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>${extra}</w:styles>`;

async function docx(body: string, styleXml = styles()) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  zip.file('word/styles.xml', styleXml);
  return zip.generateAsync({ type: 'uint8array' });
}
async function mount(body: string, options: { editable?: boolean; styleXml?: string } = {}) {
  const host = document.createElement('div'); document.body.append(host);
  let editor: DocxEditorType | null = null;
  const src = await docx(body, options.styleXml ?? styles());
  const mounted = createApp({ render: () => h(DocxEditorVue, { src, editable: options.editable, onReady: (ed: DocxEditorType) => (editor = ed) }) });
  mounted.mount(host);
  const flush = async () => { for (let i = 0; i < 7; i++) { await new Promise(r => setTimeout(r, 0)); await nextTick(); } };
  for (let i = 0; i < 80 && !editor; i++) await flush();
  await flush();
  if (!editor) throw new Error('DocxEditorVue did not emit ready');
  const nav = () => host.querySelector('.dx-outline') as HTMLElement | null;
  const openNav = async () => {
    (host.querySelector('.dx-status-btn[title*="導覽窗格"]') as HTMLButtonElement).click();
    await flush();
  };
  const done = () => { mounted.unmount(); host.remove(); };
  return { host, editor, nav, openNav, flush, done };
}
function positionOf(editor: DocxEditorType, text: string) {
  let pos = -1;
  editor.view!.state.doc.descendants((n, p) => { if (pos < 0 && n.isText && n.text!.includes(text)) pos = p + n.text!.indexOf(text); });
  if (pos < 0) throw new Error(`missing text: ${text}`);
  return pos;
}
function activeIndex(nav: HTMLElement) {
  return Array.from(nav.querySelectorAll('ol > li')).findIndex((li) => li.classList.contains('active'));
}

describe('QA20K navigation UI scenarios 00811–00820', () => {
  it('QA20K-00811 clicking the second duplicate heading moves the active highlight to that exact row', async () => {
    const ui = await mount(heading('Summary') + para('middle') + heading('Summary'));
    await ui.openNav();
    const rows = ui.nav()!.querySelectorAll('ol > li');
    expect(rows).toHaveLength(2);
    (rows[1].querySelector('button') as HTMLButtonElement).click(); await ui.flush();
    const second = ui.editor.outline()[1];
    expect(ui.editor.view!.state.selection.$from.before()).toBe(second.pos);
    expect(activeIndex(ui.nav()!)).toBe(1);
    ui.done();
  });

  it('QA20K-00812 keeps the preceding heading active while the cursor is in its body paragraph', async () => {
    const ui = await mount(heading('Expense policy') + para('Employees submit receipts within 30 days.') + heading('Exceptions'));
    await ui.openNav();
    const pos = positionOf(ui.editor, 'Employees');
    ui.editor.view!.dispatch(ui.editor.view!.state.tr.setSelection(TextSelection.create(ui.editor.view!.state.doc, pos)));
    await ui.flush();
    expect(activeIndex(ui.nav()!)).toBe(0);
    expect(ui.nav()!.querySelector('li.active button')!.textContent).toBe('Expense policy');
    ui.done();
  });

  it('QA20K-00813 shows no active heading when the cursor precedes the first heading', async () => {
    const ui = await mount(para('Cover note') + heading('Approval workflow'));
    await ui.openNav();
    ui.editor.view!.dispatch(ui.editor.view!.state.tr.setSelection(TextSelection.create(ui.editor.view!.state.doc, positionOf(ui.editor, 'Cover note'))));
    await ui.flush();
    expect(activeIndex(ui.nav()!)).toBe(-1);
    expect(Array.from(ui.nav()!.querySelectorAll('ol > li button')).map(b => b.textContent?.trim())).toEqual(['Approval workflow']);
    ui.done();
  });

  it('QA20K-00814 gives a useful empty-state message for a document without headings', async () => {
    const ui = await mount(para('Ordinary body text only.'));
    await ui.openNav();
    expect(ui.nav()!.querySelector('ol')).toBeNull();
    expect(ui.nav()!.querySelector('.dx-outline-empty')?.textContent).toContain('這份文件沒有標題');
    ui.done();
  });

  it('QA20K-00815 indents rendered heading levels in consistent 14-pixel steps', async () => {
    const ui = await mount(heading('Policy', 'Heading1') + heading('Travel', 'Heading2') + heading('Rail', 'Heading3'));
    await ui.openNav();
    expect(Array.from(ui.nav()!.querySelectorAll('ol > li')).map(li => (li as HTMLElement).style.paddingLeft)).toEqual(['0px', '14px', '28px']);
    ui.done();
  });

  it('QA20K-00816 closing and reopening the navigation pane restores its headings', async () => {
    const ui = await mount(heading('Monthly close') + heading('Reconciliation', 'Heading2'));
    await ui.openNav();
    expect(ui.nav()!.querySelectorAll('ol > li')).toHaveLength(2);
    (ui.nav()!.querySelector('.dx-outline-close') as HTMLButtonElement).click(); await ui.flush();
    expect(ui.nav()).toBeNull();
    await ui.openNav();
    expect(Array.from(ui.nav()!.querySelectorAll('ol > li button')).map(b => b.textContent?.trim())).toEqual(['Monthly close', 'Reconciliation']);
    ui.done();
  });

  it('QA20K-00817 keeps heading navigation available in read-only mode but hides page-number editing', async () => {
    const body = heading('Quarterly report') + field('Q4', '4') + `<w:p><w:bookmarkStart w:id="1" w:name="Q4"/><w:r><w:t>Details</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;
    const ui = await mount(body, { editable: false });
    await ui.openNav();
    expect(ui.nav()!.querySelector('li button')?.textContent?.trim()).toBe('Quarterly report');
    expect(ui.nav()!.querySelector('.dx-outline-update')).toBeNull();
    (ui.nav()!.querySelector('li button') as HTMLButtonElement).click(); await ui.flush();
    expect(ui.editor.view!.state.selection.$from.parent.textContent).toBe('Quarterly report');
    ui.done();
  });

  it('QA20K-00818 hides the update-pages control when the document has no PAGEREF fields', async () => {
    const ui = await mount(heading('Runbook') + para('Steps'));
    await ui.openNav();
    expect(ui.nav()!.querySelectorAll('li')).toHaveLength(1);
    expect(ui.nav()!.querySelector('.dx-outline-update')).toBeNull();
    expect(ui.nav()!.textContent).not.toContain('更新目錄頁碼');
    ui.done();
  });

  it('QA20K-00819 normalizes repeated whitespace in the navigation label and tooltip', async () => {
    const ui = await mount(heading('Quarterly    revenue\tforecast'));
    await ui.openNav();
    const button = ui.nav()!.querySelector('li button') as HTMLButtonElement;
    expect(button.textContent?.trim()).toBe('Quarterly revenue forecast');
    expect(button.title).toBe('Quarterly revenue forecast');
    ui.done();
  });

  it('QA20K-00820 refreshes a stale page number from the UI and makes that change undoable', async () => {
    const body = field('Current', '12') + `<w:p><w:bookmarkStart w:id="1" w:name="Current"/><w:r><w:t>Current chapter</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;
    const ui = await mount(body);
    await ui.openNav();
    (ui.editor as any).pageNumberAt = () => 7;
    (ui.nav()!.querySelector('.dx-outline-update') as HTMLButtonElement).click(); await ui.flush();
    expect(ui.editor.view!.state.doc.textContent).toContain('7');
    expect(ui.host.querySelector('.dx-notice')?.textContent).toContain('已更新 1 個頁碼');
    expect(ui.editor.undo()).toBe(true); await ui.flush();
    expect(ui.editor.view!.state.doc.textContent).toContain('12');
    expect(ui.editor.view!.state.doc.textContent).not.toContain('7');
    ui.done();
  });
});
