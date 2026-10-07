import { afterAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { compileScript, parse } from '@vue/compiler-sfc';
import ts from 'typescript';
import { createApp, h, nextTick } from 'vue';
import { instances } from './qa20kZoomStubs';

const SFC = resolve(__dirname, '../src/papyrus/vue/DocxEditor.vue');
const OUT = resolve(__dirname, `.qa-zoom-${process.pid}.tmp.js`);
const slash = (p: string) => p.replace(/\\/g, '/');
afterAll(() => rmSync(OUT, { force: true }));
// The zoom picked is remembered on this computer (persona-300 B-10): each test starts without one.
beforeEach(() => { instances.length = 0; localStorage.clear(); });

async function compiled(): Promise<any> {
  const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
  const script = compileScript(descriptor, { id: 'dx-zoom-test', inlineTemplate: true });
  let code = ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText;
  const stubs = slash(resolve(__dirname, 'qa20kZoomStubs.ts'));
  code = code.replace(/from\s+(['"])(\.{1,2}\/[^'"]+)\1/g, (_m, _q, spec: string) =>
    spec.endsWith('.vue') || spec === '../editor/core' ? `from '${stubs}'` : `from '${slash(resolve(dirname(SFC), spec))}'`,
  );
  writeFileSync(OUT, code);
  return (await import(/* @vite-ignore */ slash(OUT))).default;
}
const flush = async () => { for (let i = 0; i < 8; i++) { await new Promise(r => setTimeout(r, 0)); await nextTick(); } };
async function mount(editable = true) {
  const Editor = await compiled();
  const host = document.createElement('div'); document.body.append(host);
  const app = createApp({ render: () => h(Editor, { src: null, editable, toolbar: false }) });
  app.mount(host); await flush();
  return { host, app, select: host.querySelector('select[aria-label="縮放"]') as HTMLSelectElement };
}
async function choose(value: string, editable = true) {
  const mounted = await mount(editable);
  mounted.select.value = value;
  mounted.select.dispatchEvent(new Event('change', { bubbles: true }));
  await flush();
  return mounted;
}

describe('QA20K zoom and page-fit workflow 00641–00650', () => {
  it('QA20K-00641 starts the status zoom selector at 100 percent', async () => {
    const x = await mount(); expect(x.select.value).toBe('1'); expect(x.select.options[3].textContent).toContain('100%'); x.app.unmount(); x.host.remove();
  });
  it('QA20K-00642 selecting 50 percent sends numeric 0.5 zoom to the editor', async () => {
    const x = await choose('0.5'); expect(x.select.value).toBe('0.5'); expect(instances[0].zoomCalls).toEqual([0.5]); x.app.unmount(); x.host.remove();
  });
  it('QA20K-00643 selecting 75 percent sends numeric 0.75 zoom to the editor', async () => {
    const x = await choose('0.75'); expect(x.select.value).toBe('0.75'); expect(instances[0].zoomCalls).toEqual([0.75]); x.app.unmount(); x.host.remove();
  });
  it('QA20K-00644 selecting 90 percent sends numeric 0.9 zoom to the editor', async () => {
    const x = await choose('0.9'); expect(x.select.value).toBe('0.9'); expect(instances[0].zoomCalls).toEqual([0.9]); x.app.unmount(); x.host.remove();
  });
  it('QA20K-00645 selecting 125 percent sends numeric 1.25 zoom to the editor', async () => {
    const x = await choose('1.25'); expect(x.select.value).toBe('1.25'); expect(instances[0].zoomCalls).toEqual([1.25]); x.app.unmount(); x.host.remove();
  });
  it('QA20K-00646 selecting 150 percent sends numeric 1.5 zoom to the editor', async () => {
    const x = await choose('1.5'); expect(x.select.value).toBe('1.5'); expect(instances[0].zoomCalls).toEqual([1.5]); x.app.unmount(); x.host.remove();
  });
  it('QA20K-00647 selecting 200 percent sends numeric 2 zoom to the editor', async () => {
    const x = await choose('2'); expect(x.select.value).toBe('2'); expect(instances[0].zoomCalls).toEqual([2]); x.app.unmount(); x.host.remove();
  });
  it('QA20K-00648 choosing fit invokes the editor page-width zoom mode', async () => {
    const x = await choose('fit'); expect(x.select.value).toBe('fit'); expect(instances[0].zoomCalls).toEqual(['fit']); x.app.unmount(); x.host.remove();
  });
  it('QA20K-00649 a read-only user can still change viewing zoom', async () => {
    const x = await choose('0.75', false); expect(x.select.value).toBe('0.75'); expect(instances[0].zoomCalls).toEqual([0.75]); x.app.unmount(); x.host.remove();
  });
  it('QA20K-00650 changing zoom does not emit a document content change', async () => {
    const changed = vi.fn();
    const Editor = await compiled(); const host = document.createElement('div'); document.body.append(host);
    const app = createApp({ render: () => h(Editor, { src: null, editable: true, toolbar: false, onChange: changed }) });
    app.mount(host); await flush();
    const select = host.querySelector('select[aria-label="縮放"]') as HTMLSelectElement;
    select.value = '1.5'; select.dispatchEvent(new Event('change', { bubbles: true })); await flush();
    expect(instances.at(-1)!.zoomCalls).toEqual([1.5]); expect(changed).not.toHaveBeenCalled(); app.unmount(); host.remove();
  });
});
