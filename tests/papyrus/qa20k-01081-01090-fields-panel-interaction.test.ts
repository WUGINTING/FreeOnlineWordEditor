import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick, ref } from 'vue';
import JSZip from 'jszip';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';
import { blankPackage } from '../../src/papyrus/docx/template';
import type { DocxEditor } from '../../src/papyrus/editor/core';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const text = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;
const field = (id: number, alias: string, tag: string, value: string, opts: { required?: boolean; placeholder?: boolean; kind?: string; lock?: string } = {}) => {
  const requiredTag = opts.required ? `${tag}-required` : tag;
  const placeholder = opts.placeholder ? '<w:showingPlcHdr/>' : '';
  const placeholderText = opts.placeholder ? '<w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>' : '';
  const lock = opts.lock ? `<w:lock w:val="${opts.lock}"/>` : '';
  return `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:tag w:val="${requiredTag}"/><w:id w:val="${id}"/>${placeholder}${lock}${opts.kind ?? '<w:text/>'}</w:sdtPr><w:sdtContent>${placeholderText}<w:r><w:t>${value}</w:t></w:r></w:sdtContent></w:sdt>`;
};
const para = (content: string) => `<w:p>${content}</w:p>`;
async function source(body: string): Promise<Blob> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return new Blob([await zip.generateAsync({ type: 'uint8array' })], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

let app: ReturnType<typeof createApp> | null = null;
let host: HTMLElement | null = null;
let editor: DocxEditor | null = null;
afterEach(() => { app?.unmount(); app = null; host?.remove(); host = null; editor = null; });

async function open(body: string, editable = true) {
  host = document.createElement('div'); document.body.append(host);
  const src = await source(body);
  const vm = ref<any>(null);
  let ready!: (value: DocxEditor) => void;
  const readyPromise = new Promise<DocxEditor>((resolve) => { ready = resolve; });
  app = createApp({ render: () => h(DocxEditorVue, { ref: vm, src, editable, toolbar: false, onReady: ready }) });
  app.mount(host);
  const ed = await readyPromise;
  await nextTick();
  await new Promise((resolve) => setTimeout(resolve, 0));
  editor = ed;
  return { ed, root: host };
}

function card(root: HTMLElement, label: string): HTMLElement {
  const found = [...root.querySelectorAll<HTMLElement>('.dx-field-card')].find((el) => el.textContent?.includes(label));
  if (!found) throw new Error(`Missing field card: ${label}`);
  return found;
}
function panel(root: HTMLElement): HTMLElement | null { return root.querySelector<HTMLElement>('[aria-label="填寫欄位"]'); }

describe('QA20K FieldsPanel mounted user workflows 01081–01090', () => {
  it('QA20K-01081 opens the fields panel with completion and required-field status for a template', async () => {
    const { root } = await open(para(field(1, '客戶名稱', 'customer', '按一下輸入客戶名稱', { required: true, placeholder: true }) + field(2, '案號', 'case-id', 'A-102')));
    expect(panel(root)?.querySelector('.dx-fields-title')?.textContent?.trim()).toBe('欄位（已填 1 / 2）');
    expect(panel(root)?.querySelector('[role="status"]')?.textContent?.trim()).toBe('還有 1 個必填欄位沒有填寫');
    expect(card(root, '客戶名稱').closest('li')?.classList.contains('required')).toBe(true);
  });

  it('QA20K-01082 keeps the side panel out of an ordinary document without fillable controls', async () => {
    const { root } = await open(para(text('一般會議紀錄，沒有範本欄位。')));
    expect(panel(root)).toBeNull();
    expect(root.textContent).toContain('一般會議紀錄，沒有範本欄位。');
  });

  it('QA20K-01083 selecting a field card moves the editor selection to that field and marks it active', async () => {
    const { ed, root } = await open(para(field(1, '部門', 'department', '營運') + text('，') + field(2, '收件人', 'recipient', '林怡君')));
    card(root, '收件人').click();
    await nextTick();
    expect(ed.view!.state.doc.textBetween(ed.view!.state.selection.from, ed.view!.state.selection.to)).toBe('林怡君');
    expect(card(root, '收件人').closest('li')?.classList.contains('active')).toBe(true);
    expect(card(root, '部門').closest('li')?.classList.contains('active')).toBe(false);
  });

  it('QA20K-01084 the next-field control skips completed entries and selects the next placeholder', async () => {
    const { ed, root } = await open(para(field(1, '部門', 'department', '營運') + field(2, '客戶', 'customer', '輸入客戶名稱', { required: true, placeholder: true }) + field(3, '金額', 'amount', '輸入金額', { placeholder: true })));
    const next = root.querySelector<HTMLButtonElement>('.dx-fields-next');
    expect(next).not.toBeNull();
    next!.click();
    await nextTick();
    expect(ed.view!.state.doc.textBetween(ed.view!.state.selection.from, ed.view!.state.selection.to)).toBe('輸入客戶名稱');
    expect(card(root, '客戶').closest('li')?.classList.contains('active')).toBe(true);
  });

  it('QA20K-01085 updates the visible completion count and required status after a placeholder is filled', async () => {
    const { ed, root } = await open(para(field(1, '客戶名稱', 'customer', '按一下輸入客戶名稱', { required: true, placeholder: true }) + field(2, '日期', 'date', '選擇日期', { placeholder: true, kind: '<w:date/>' })));
    card(root, '客戶名稱').click();
    ed.run((state, dispatch) => (dispatch?.(state.tr.insertText('北辰股份有限公司')), true));
    await nextTick();
    expect(panel(root)?.querySelector('.dx-fields-title')?.textContent?.trim()).toBe('欄位（已填 1 / 2）');
    expect(panel(root)?.querySelector('[role="status"]')?.textContent?.trim()).toBe('必填欄位都已填寫');
    expect(card(root, '客戶名稱').textContent).toContain('北辰股份有限公司');
    expect(ed.fields()[0].filled).toBe(true);
  });

  it('QA20K-01086 distinguishes an optional blank field from a required missing field', async () => {
    const { root } = await open(para(field(1, '必填地址', 'address', '按一下輸入地址', { required: true, placeholder: true }) + field(2, '備註', 'note', '   ', { placeholder: false })));
    expect(panel(root)?.querySelector('.dx-fields-title')?.textContent?.trim()).toBe('欄位（已填 0 / 2）');
    expect(card(root, '必填地址').querySelector('.dx-field-state')?.textContent?.trim()).toBe('!');
    expect(card(root, '備註').querySelector('.dx-field-state')?.textContent?.trim()).toBe('○');
    expect(panel(root)?.querySelector('[role="status"]')?.textContent?.trim()).toBe('還有 1 個必填欄位沒有填寫');
  });

  it('QA20K-01087 keeps field information available in read-only mode but hides the fill-next action', async () => {
    const { root } = await open(para(field(1, '簽署日期', 'sign-date', '選擇日期', { required: true, placeholder: true, kind: '<w:date/>' })), false);
    expect(panel(root)?.textContent).toContain('簽署日期');
    expect(panel(root)?.querySelector('[role="status"]')?.textContent?.trim()).toBe('還有 1 個必填欄位沒有填寫');
    expect(panel(root)?.querySelector('.dx-fields-next')).toBeNull();
    expect(editor?.view?.editable).toBe(false);
  });

  it('QA20K-01088 identifies a locked required field and routes next-field navigation to an editable field', async () => {
    const { ed, root } = await open(para(field(1, '核准條款', 'terms', '不可變更條款', { required: true, lock: 'contentLocked' }) + field(2, '聯絡人', 'contact', '按一下輸入聯絡人', { required: true, placeholder: true })));
    expect(card(root, '核准條款').textContent).toContain('鎖定');
    root.querySelector<HTMLButtonElement>('.dx-fields-next')!.click();
    await nextTick();
    expect(ed.fields().find((f) => f.title === '聯絡人')?.id).toBeTruthy();
    expect(card(root, '聯絡人').closest('li')?.classList.contains('active')).toBe(true);
    expect(card(root, '核准條款').closest('li')?.classList.contains('active')).toBe(false);
  });

  it('QA20K-01089 presents the control kind and placeholder cue before the user opens a field', async () => {
    const { root } = await open(para(text('欄位尚待確認：') + field(1, '預計到職日', 'start-date', '選擇日期', { placeholder: true, kind: '<w:date/>' })));
    expect(card(root, '預計到職日').textContent).toContain('日期');
    expect(card(root, '預計到職日').textContent).toContain('（尚未填寫）');
    expect(card(root, '預計到職日').closest('li')?.classList.contains('active')).toBe(false);
  });

  it('QA20K-01090 closing the fields panel removes it while leaving the opened document intact', async () => {
    const { ed, root } = await open(para(field(1, '客戶', 'customer', '星河科技')));
    root.querySelector<HTMLButtonElement>('.dx-fields-close')!.click();
    await nextTick();
    expect(panel(root)).toBeNull();
    expect(ed.view!.state.doc.textContent).toContain('星河科技');
  });
});
