// @vitest-environment jsdom
import { createApp, h, reactive, type App, nextTick } from 'vue';
import { afterEach, describe, expect, it } from 'vitest';
import type { CompatReport } from '@/papyrus/docx/compat';
import CompatNotice from '@/papyrus/vue/CompatNotice.vue';

let app: App | null = null;
let root: HTMLElement;
function mount(report: CompatReport | null, keepOriginal?: string) {
  const props = reactive({ report, keepOriginal });
  root = document.createElement('div'); document.body.append(root);
  app = createApp({ render: () => h(CompatNotice, { report: props.report, keepOriginal: props.keepOriginal }) });
  app.mount(root);
  return { props };
}
const item = (id: string, title: string, effect: string, where = '') => ({ id, title, count: 1, effect, where });
const click = (selector: string) => (root.querySelector(selector) as HTMLButtonElement).click();
afterEach(() => { app?.unmount(); app = null; root?.remove(); });

describe('QA20K compatibility notice user workflows 01151–01160', () => {
  it('QA20K-01151 keeps the editor clear before a compatibility scan has produced a report', () => {
    mount(null);
    expect(root.querySelector('.dx-compat')).toBeNull();
    expect(root.querySelector('[role="status"]')).toBeNull();
  });

  it('QA20K-01152 avoids warning an employee when the opened file has no limited features', () => {
    mount({ items: [] });
    expect(root.querySelector('.dx-compat')).toBeNull();
    expect(root.textContent).not.toContain('這份文件有');
  });

  it('QA20K-01153 gives a compact accessible count before the employee opens details', () => {
    mount({ items: [item('shape', '圖案', '圖案不會顯示。'), item('chart', '圖表', '圖表不會顯示。')] });
    const status = root.querySelector('.dx-compat[role="status"]');
    expect(status?.textContent).toContain('這份文件有 2 項內容在網頁上無法完整顯示或編輯');
    expect(status?.querySelector('.dx-compat-panel')).toBeNull();
    expect(status?.querySelector('.dx-compat-link')?.getAttribute('aria-expanded')).toBe('false');
  });

  it('QA20K-01154 reveals the unsupported feature names in the order supplied by the report', async () => {
    mount({ items: [item('textbox', '文字方塊', '文字方塊會保存在檔案中。'), item('equation', '公式', '公式可在原檔中保留。')] });
    click('.dx-compat-link'); await nextTick();
    expect([...root.querySelectorAll('.dx-compat-title')].map((el) => el.textContent?.trim())).toEqual(['文字方塊', '公式']);
  });

  it('QA20K-01155 lets an employee read the concrete editing impact before changing a report', async () => {
    mount({ items: [item('chart', '內嵌圖表', '網頁版不會顯示圖表資料；存檔時仍保留原始圖表內容。')] });
    click('.dx-compat-link'); await nextTick();
    expect(root.querySelector('.dx-compat-effect')?.textContent).toBe('網頁版不會顯示圖表資料；存檔時仍保留原始圖表內容。');
    expect(root.querySelector('.dx-compat-effect')?.textContent).not.toContain('可完整編輯');
  });

  it('QA20K-01156 identifies the affected document area when the report has a location', async () => {
    mount({ items: [item('columns', '分欄設定', '分欄版面在網頁版有限制。', '第 2 節')] });
    click('.dx-compat-link'); await nextTick();
    expect(root.querySelector('.dx-compat-where')?.textContent).toBe('第 2 節');
    expect(root.querySelector('.dx-compat-title')?.textContent).toContain('分欄設定');
  });

  it('QA20K-01157 does not invent a section location for a compatibility item with no location', async () => {
    mount({ items: [item('ole', '嵌入式物件', '網頁版不支援編輯此物件。', ''), item('fields', '欄位', '顯示快取結果。', '頁首頁尾 1 處')] });
    click('.dx-compat-link'); await nextTick();
    const entries = [...root.querySelectorAll('.dx-compat-title')];
    expect(entries[0].querySelector('.dx-compat-where')).toBeNull();
    expect(entries[0].textContent).not.toContain('頁首頁尾');
    expect(entries[1].querySelector('.dx-compat-where')?.textContent).toBe('頁首頁尾 1 處');
  });

  it('QA20K-01158 tells the employee how to retrieve the untouched uploaded original', async () => {
    mount({ items: [item('ole', '嵌入式物件', '網頁版保留但不支援編輯。')] });
    click('.dx-compat-link'); await nextTick();
    const guidance = root.querySelector('.dx-compat-keep')?.textContent ?? '';
    expect(guidance).toContain('保留原檔');
    expect(guidance).toContain('版本紀錄');
    expect(guidance).toContain('下載');
    expect(guidance).toContain('存檔也不會刪除上列內容');
  });

  it('QA20K-01159 dismisses the current compatibility banner so it no longer obscures the workspace', async () => {
    mount({ items: [item('shape', 'SmartArt', '僅保留原始內容。')] });
    click('.dx-compat-close'); await nextTick();
    expect(root.querySelector('.dx-compat')).toBeNull();
  });

  it('QA20K-01160 shows only the new document warning after a dismissed report is replaced', async () => {
    const ui = mount({ items: [item('chart-old', '舊報告圖表', '舊圖表影響。')] });
    click('.dx-compat-close'); await nextTick();
    expect(root.querySelector('.dx-compat')).toBeNull();
    ui.props.report = { items: [item('textbox-new', '新合約文字方塊', '新文字方塊影響。')] };
    await nextTick();
    expect(root.querySelector('.dx-compat-text')?.textContent).toContain('1 項內容');
    expect(root.querySelector('.dx-compat-panel')).toBeNull();
    click('.dx-compat-link'); await nextTick();
    expect(root.querySelector('.dx-compat-panel')?.textContent).toContain('新合約文字方塊');
    expect(root.querySelector('.dx-compat-panel')?.textContent).not.toContain('舊報告圖表');
  });
});
