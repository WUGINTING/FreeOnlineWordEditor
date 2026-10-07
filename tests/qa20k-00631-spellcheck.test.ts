import { afterAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { compileScript, parse } from '@vue/compiler-sfc';
import ts from 'typescript';
import { createApp, h, nextTick } from 'vue';
import { instances } from './qa20kSpellcheckStubs';

const SFC = resolve(__dirname, '../src/papyrus/vue/DocxEditor.vue');
const OUT = resolve(__dirname, `.qa-spellcheck-${process.pid}.tmp.js`);
const slash = (p: string) => p.replace(/\\/g, '/');
afterAll(() => rmSync(OUT, { force: true }));
beforeEach(() => { localStorage.clear(); instances.length = 0; });

async function compiled(): Promise<any> {
  const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
  const script = compileScript(descriptor, { id: 'dx-spellcheck-test', inlineTemplate: true });
  let code = ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText;
  const stubs = slash(resolve(__dirname, 'qa20kSpellcheckStubs.ts'));
  code = code.replace(/from\s+(['"])(\.{1,2}\/[^'"]+)\1/g, (_m, _q, spec: string) =>
    spec.endsWith('.vue') || spec === '../editor/core' ? `from '${stubs}'` : `from '${slash(resolve(dirname(SFC), spec))}'`,
  );
  writeFileSync(OUT, code);
  return (await import(/* @vite-ignore */ slash(OUT))).default;
}

const flush = async () => {
  for (let i = 0; i < 10; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); }
};

async function mount(editable = true, onChange = () => {}) {
  const Editor = await compiled();
  const host = document.createElement('div');
  document.body.append(host);
  const app = createApp({ render: () => h(Editor, { src: null, editable, toolbar: false, onChange }) });
  app.mount(host);
  await flush();
  return { host, app, button: () => host.querySelector('.dx-status-btn[title*="拼字檢查"]') as HTMLButtonElement | null };
}

describe('QA20K browser spell-check preference workflow 00631–00640', () => {
  it('QA20K-00631 starts with browser spellcheck off when no preference exists', async () => {
    const { host, app, button } = await mount();
    expect(button()?.getAttribute('aria-pressed')).toBe('false');
    expect(button()?.textContent).toContain('關');
    expect(instances[0].view!.dom.getAttribute('spellcheck')).toBe('false');
    app.unmount(); host.remove();
  });

  it('QA20K-00632 applies a previously saved on preference to the new editor', async () => {
    localStorage.setItem('papyrus.spellcheck', '1');
    const { host, app, button } = await mount();
    expect(button()?.getAttribute('aria-pressed')).toBe('true');
    expect(instances[0].view!.dom.getAttribute('spellcheck')).toBe('true');
    app.unmount(); host.remove();
  });

  it('QA20K-00633 honors an explicitly saved off preference', async () => {
    localStorage.setItem('papyrus.spellcheck', '0');
    const { host, app, button } = await mount();
    expect(button()?.textContent).toContain('關');
    expect(instances[0].spellcheckOn).toBe(false);
    expect(instances[0].view!.dom.getAttribute('spellcheck')).toBe('false');
    app.unmount(); host.remove();
  });

  it('QA20K-00634 enables spellcheck immediately and remembers the choice', async () => {
    const { host, app, button } = await mount();
    button()!.click(); await flush();
    expect(button()?.getAttribute('aria-pressed')).toBe('true');
    expect(instances[0].view!.dom.getAttribute('spellcheck')).toBe('true');
    expect(localStorage.getItem('papyrus.spellcheck')).toBe('1');
    app.unmount(); host.remove();
  });

  it('QA20K-00635 disabling spellcheck updates editor, button and stored preference', async () => {
    localStorage.setItem('papyrus.spellcheck', '1');
    const { host, app, button } = await mount();
    button()!.click(); await flush();
    expect(button()?.getAttribute('aria-pressed')).toBe('false');
    expect(instances[0].view!.dom.getAttribute('spellcheck')).toBe('false');
    expect(localStorage.getItem('papyrus.spellcheck')).toBe('0');
    app.unmount(); host.remove();
  });

  it('QA20K-00636 repeated toggles leave the final on state and preference aligned', async () => {
    const { host, app, button } = await mount();
    button()!.click(); await flush(); button()!.click(); await flush(); button()!.click(); await flush();
    expect(button()?.getAttribute('aria-pressed')).toBe('true');
    expect(instances[0].spellcheckOn).toBe(true);
    expect(localStorage.getItem('papyrus.spellcheck')).toBe('1');
    app.unmount(); host.remove();
  });

  it('QA20K-00637 preference survives closing and reopening the editor component', async () => {
    const first = await mount();
    first.button()!.click(); await flush(); first.app.unmount(); first.host.remove();
    const second = await mount();
    expect(second.button()?.getAttribute('aria-pressed')).toBe('true');
    expect(instances.at(-1)!.view!.dom.getAttribute('spellcheck')).toBe('true');
    second.app.unmount(); second.host.remove();
  });

  it('QA20K-00638 a read-only collaborator gets no spellcheck mutation control', async () => {
    localStorage.setItem('papyrus.spellcheck', '1');
    const { host, app, button } = await mount(false);
    expect(button()).toBeNull();
    expect(localStorage.getItem('papyrus.spellcheck')).toBe('1');
    app.unmount(); host.remove();
  });

  it('QA20K-00639 changing spellcheck preference does not emit a document-change event', async () => {
    const changed = vi.fn();
    const { host, app, button } = await mount(true, changed);
    button()!.click(); await flush();
    expect(changed).not.toHaveBeenCalled();
    expect(instances[0].view!.dom.getAttribute('spellcheck')).toBe('true');
    app.unmount(); host.remove();
  });

  it('QA20K-00640 storage denial keeps the current toggle usable for this session', async () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new Error('storage denied'); });
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('storage denied'); });
    const { host, app, button } = await mount();
    expect(button()?.getAttribute('aria-pressed')).toBe('false');
    button()!.click(); await flush();
    expect(button()?.getAttribute('aria-pressed')).toBe('true');
    expect(instances[0].view!.dom.getAttribute('spellcheck')).toBe('true');
    app.unmount(); host.remove();
    vi.restoreAllMocks();
  });
});
