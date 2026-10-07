import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { compileScript, parse } from '@vue/compiler-sfc';
import ts from 'typescript';
import { createApp, h, nextTick, reactive } from 'vue';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';

const SFC = resolve(__dirname, '../../src/papyrus/vue/PageSetupDialog.vue');
const OUT = resolve(__dirname, `.qa-page-setup-dialog-${process.pid}.tmp.js`);
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const page = (width = 11906, height = 16838, left = 1440) =>
  `<w:pgSz w:w="${width}" w:h="${height}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="${left}" w:header="720" w:footer="720"/>`;

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
});
afterAll(() => rmSync(OUT, { force: true }));

async function compiledDialog(): Promise<any> {
  const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
  const script = compileScript(descriptor, { id: 'qa-page-setup-dialog', inlineTemplate: true });
  const code = ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText;
  writeFileSync(OUT, code);
  return (await import(/* @vite-ignore */ OUT.replace(/\\/g, '/'))).default;
}

async function openEditor() {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t>Memo</w:t></w:r></w:p><w:sectPr>${page()}</w:sectPr></w:body></w:document>`);
  const host = document.createElement('div'); document.body.append(host);
  const editor = new DocxEditor(host);
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  return { editor, host, done: () => { editor.destroy(); host.remove(); } };
}

async function mountDialog(editor: DocxEditor) {
  const component = await compiledDialog();
  const host = document.createElement('div'); document.body.append(host);
  const state = reactive({ open: true });
  const app = createApp({ setup: () => () => h(component, { editor, open: state.open, onClose: () => { state.open = false; } }) });
  app.mount(host);
  await nextTick();
  await nextTick();
  return {
    host,
    async setOpen(open: boolean) { state.open = open; await nextTick(); await nextTick(); },
    done: () => { app.unmount(); host.remove(); },
  };
}

function typeNumber(host: HTMLElement, index: number, value: number) {
  const input = host.querySelectorAll<HTMLInputElement>('input[type="number"]')[index];
  input.value = String(value);
  input.dispatchEvent(new Event('input', { bubbles: true }));
}
async function apply(host: HTMLElement) {
  host.querySelector<HTMLButtonElement>('.dx-primary')!.click();
  await nextTick();
  await nextTick();
}
async function savedSection(editor: DocxEditor) {
  const zip = await JSZip.loadAsync(await editor.save());
  const xml = await zip.file('word/document.xml')!.async('string');
  return xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0] ?? '';
}
function changedOrientation(host: HTMLElement, landscape: boolean) {
  const input = host.querySelectorAll<HTMLInputElement>('input[name="dx-ps-orient"]')[landscape ? 1 : 0];
  input.checked = true;
  input.dispatchEvent(new Event('change', { bubbles: true }));
}

describe('QA20K section page setup dialog interaction workflows 01031–01040', () => {
  it('QA20K-01031 discards uncommitted page changes through each supported dismissal path', async () => {
    for (const dismissal of ['cancel', 'escape', 'backdrop'] as const) {
      const x = await openEditor();
      const before = await savedSection(x.editor);
      const d = await mountDialog(x.editor);
      try {
        typeNumber(d.host, 0, 20);
        const dialog = d.host.querySelector<HTMLElement>('[role="dialog"]')!;
        if (dismissal === 'cancel') {
          [...d.host.querySelectorAll('button')].find(button => button.textContent?.includes('取消'))!.click();
        } else if (dismissal === 'escape') {
          dialog.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
        } else {
          d.host.querySelector<HTMLElement>('.dx-dialog-backdrop')!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        }
        await nextTick(); await nextTick();
        expect(d.host.querySelector('[role="dialog"]')).toBeNull();
        expect(x.editor.isModified()).toBe(false);
        expect(await savedSection(x.editor)).toBe(before);
      } finally { d.done(); x.done(); }
    }
  });

  it('QA20K-01032 keeps the setup dialog open while the user interacts inside it', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      const dialog = d.host.querySelector<HTMLElement>('[role="dialog"]')!;
      typeNumber(d.host, 0, 20);
      dialog.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      await nextTick();
      expect(d.host.querySelector('[role="dialog"]')).not.toBeNull();
      expect(d.host.querySelectorAll<HTMLInputElement>('input[type="number"]')[0].value).toBe('20');
      expect(x.editor.isModified()).toBe(false);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-01033 moves focus into the dialog and restores it to the control that opened it', async () => {
    const x = await openEditor();
    const trigger = document.createElement('button'); trigger.textContent = 'Page setup'; document.body.append(trigger); trigger.focus();
    const d = await mountDialog(x.editor);
    try {
      const firstControl = d.host.querySelector<HTMLElement>('[role="dialog"] select, [role="dialog"] input, [role="dialog"] button')!;
      expect(document.activeElement).toBe(firstControl);
      [...d.host.querySelectorAll('button')].find(button => button.textContent?.includes('取消'))!.click();
      await nextTick(); await nextTick();
      expect(document.activeElement).toBe(trigger);
    } finally { d.done(); trigger.remove(); x.done(); }
  });

  it('QA20K-01034 cycles keyboard focus from the last to first control and back', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      const dialog = d.host.querySelector<HTMLElement>('[role="dialog"]')!;
      const controls = [...dialog.querySelectorAll<HTMLElement>('select, input, button')];
      const first = controls[0]; const last = controls[controls.length - 1];
      last.focus();
      const forward = new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true });
      last.dispatchEvent(forward);
      expect(forward.defaultPrevented).toBe(true);
      expect(document.activeElement).toBe(first);
      first.focus();
      const backward = new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true });
      first.dispatchEvent(backward);
      expect(backward.defaultPrevented).toBe(true);
      expect(document.activeElement).toBe(last);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-01035 exposes a named modal and an accessible page-width control', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      const dialog = d.host.querySelector<HTMLElement>('[role="dialog"]')!;
      const labelledBy = dialog.getAttribute('aria-labelledby');
      expect(dialog.getAttribute('aria-modal')).toBe('true');
      expect(labelledBy).toBeTruthy();
      expect(d.host.querySelector(`#${labelledBy}`)?.textContent).toContain('版面設定');
      const width = d.host.querySelectorAll<HTMLInputElement>('input[type="number"]')[0];
      expect(width.closest('label')?.textContent).toContain('寬');
    } finally { d.done(); x.done(); }
  });

  it('QA20K-01036 rejects a negative margin without changing the page setup', async () => {
    const x = await openEditor(); const before = await savedSection(x.editor); const d = await mountDialog(x.editor);
    try {
      typeNumber(d.host, 4, -1);
      await apply(d.host);
      expect(d.host.querySelector('[role="alert"]')?.textContent).toContain('0 以上');
      expect(d.host.querySelector('[role="dialog"]')).not.toBeNull();
      expect(d.host.querySelectorAll<HTMLInputElement>('input[type="number"]')[4].value).toBe('-1');
      expect(x.editor.isModified()).toBe(false);
      expect(await savedSection(x.editor)).toBe(before);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-01037 lets the user correct invalid input and commit the repaired margin', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      typeNumber(d.host, 4, -1);
      await apply(d.host);
      expect(d.host.querySelector('[role="alert"]')).not.toBeNull();
      typeNumber(d.host, 4, 2);
      await apply(d.host);
      expect(d.host.querySelector('[role="dialog"]')).toBeNull();
      expect(x.editor.isModified()).toBe(true);
      expect(await savedSection(x.editor)).toContain('w:left="1134"');
    } finally { d.done(); x.done(); }
  });

  it('QA20K-01038 reloads current document values and clears stale errors when reopened', async () => {
    const x = await openEditor(); const d = await mountDialog(x.editor);
    try {
      typeNumber(d.host, 4, -1);
      await apply(d.host);
      expect(d.host.querySelector('[role="alert"]')).not.toBeNull();
      [...d.host.querySelectorAll('button')].find(button => button.textContent?.includes('取消'))!.click();
      await nextTick(); await nextTick();
      x.editor.setPageSetup({ ...x.editor.cursorSection().page, width: 16838, height: 11906, marginLeft: 2000 }, 'section');
      await d.setOpen(true);
      const fields = [...d.host.querySelectorAll<HTMLInputElement>('input[type="number"]')].map(input => Number(input.value));
      expect(fields.slice(0, 2)).toEqual([29.7, 21]);
      expect(fields[4]).toBe(3.53);
      expect(d.host.querySelector('[role="alert"]')).toBeNull();
      expect(d.host.querySelectorAll<HTMLInputElement>('input[name="dx-ps-orient"]')[1].checked).toBe(true);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-01039 preserves exact original dimensions when orientation is toggled away and back', async () => {
    const x = await openEditor(); const before = await savedSection(x.editor); const d = await mountDialog(x.editor);
    try {
      changedOrientation(d.host, true);
      await nextTick();
      expect([...d.host.querySelectorAll<HTMLInputElement>('input[type="number"]')].slice(0, 2).map(input => Number(input.value))).toEqual([29.7, 21]);
      changedOrientation(d.host, false);
      await nextTick();
      expect([...d.host.querySelectorAll<HTMLInputElement>('input[type="number"]')].slice(0, 2).map(input => Number(input.value))).toEqual([21, 29.7]);
      await apply(d.host);
      expect(x.editor.isModified()).toBe(false);
      expect(await savedSection(x.editor)).toBe(before);
    } finally { d.done(); x.done(); }
  });

  it('QA20K-01040 marks manual dimensions as custom without committing them before Apply', async () => {
    const x = await openEditor(); const before = await savedSection(x.editor); const d = await mountDialog(x.editor);
    try {
      typeNumber(d.host, 0, 20);
      await nextTick();
      expect(d.host.querySelector('select')?.value).toBe('custom');
      expect(x.editor.isModified()).toBe(false);
      expect(await savedSection(x.editor)).toBe(before);
      await apply(d.host);
      expect(x.editor.isModified()).toBe(true);
      expect(await savedSection(x.editor)).toContain('<w:pgSz w:w="11339" w:h="16838"/>');
    } finally { d.done(); x.done(); }
  });
});
