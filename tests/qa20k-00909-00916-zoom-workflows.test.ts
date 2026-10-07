import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { compileScript, parse } from '@vue/compiler-sfc';
import ts from 'typescript';
import { createApp, h, nextTick } from 'vue';
import { blankPackage } from '../src/papyrus/docx/template';
import { DocxEditor } from '../src/papyrus/editor/core';
import { instances } from './qa20kZoomStubs';

const SFC = resolve(__dirname, '../src/papyrus/vue/DocxEditor.vue');
const OUT = resolve(__dirname, `.qa-zoom-usecases-${process.pid}.tmp.js`);
const slash = (p: string) => p.replace(/\\/g, '/');
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';

afterAll(() => rmSync(OUT, { force: true }));
beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
});

async function compiledUi(): Promise<any> {
  const { descriptor } = parse(readFileSync(SFC, 'utf8'), { filename: SFC });
  const script = compileScript(descriptor, { id: 'dx-zoom-usecase-test', inlineTemplate: true });
  let code = ts.transpileModule(script.content, { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2020 } }).outputText;
  const stubs = slash(resolve(__dirname, 'qa20kZoomStubs.ts'));
  code = code.replace(/from\s+(['"])(\.{1,2}\/[^'"]+)\1/g, (_m, _q, spec: string) =>
    spec.endsWith('.vue') || spec === '../editor/core' ? `from '${stubs}'` : `from '${slash(resolve(dirname(SFC), spec))}'`,
  );
  writeFileSync(OUT, code);
  return (await import(/* @vite-ignore */ slash(OUT))).default;
}

const flush = async () => { for (let i = 0; i < 8; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };

function docx(text: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p>${SECT}</w:body></w:document>`);
  return zip;
}

async function mountCore(text: string) {
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host);
  await editor.open(await docx(text).generateAsync({ type: 'uint8array' }));
  const root = host.querySelector('.dx-root') as HTMLElement;
  const canvas = host.querySelector('.dx-canvas') as HTMLElement;
  const stage = host.querySelector('.dx-stage') as HTMLElement;
  return { host, editor, root, canvas, stage, done: () => { editor.destroy(); host.remove(); } };
}

describe('QA20K-00909–00916 zoom workflows', () => {
  it('QA20K-00909 fit-to-width caps scale when a very wide screen exceeds supported zoom', async () => {
    const x=await mountCore('Wide monitor memo');
    try { Object.defineProperty(x.root,'clientWidth',{configurable:true,value:50000}); x.editor.setZoom('fit'); expect(x.editor.zoom).toBe(4); expect(x.stage.style.width).toBe(`${Number.parseFloat(x.canvas.style.width)*4}px`); }
    finally { x.done(); }
  });
  it('QA20K-00910 fixed percentage remains fixed after viewport resize', async () => {
    const x=await mountCore('Fixed reading scale');
    try { Object.defineProperty(x.root,'clientWidth',{configurable:true,value:650}); x.editor.setZoom(1.25); Object.defineProperty(x.root,'clientWidth',{configurable:true,value:1200}); window.dispatchEvent(new Event('resize')); expect(x.editor.zoom).toBe(1.25); expect(x.stage.style.width).toBe(`${Number.parseFloat(x.canvas.style.width)*1.25}px`); }
    finally { x.done(); }
  });
  it('QA20K-00911 selecting a fixed preset exits fit mode for subsequent resize events', async () => {
    const x=await mountCore('Fit then fixed');
    try { let width=640; Object.defineProperty(x.root,'clientWidth',{configurable:true,get:()=>width}); x.editor.setZoom('fit'); const first=x.editor.zoom; expect(first).not.toBe(1); x.editor.setZoom(1.5); width=900; window.dispatchEvent(new Event('resize')); expect(x.editor.zoom).toBe(1.5); }
    finally { x.done(); }
  });
  it('QA20K-00912 switching from preset to fit recalculates against the current viewport immediately', async () => {
    const x=await mountCore('Fixed then fit');
    try { let width=640; Object.defineProperty(x.root,'clientWidth',{configurable:true,get:()=>width}); x.editor.setZoom(2); width=820; x.editor.setZoom('fit'); const canvasWidth=Number.parseFloat(x.canvas.style.width); const expected=(820-32)/canvasWidth; expect(x.editor.zoom).toBeCloseTo(expected,6); expect(x.stage.style.width).toBe(`${canvasWidth*expected}px`); }
    finally { x.done(); }
  });
  it('QA20K-00913 returning to 100 percent clears transform and restores natural stage dimensions', async () => {
    const x=await mountCore('Reset view');
    try { const w=Number.parseFloat(x.canvas.style.width),h=Number.parseFloat(x.canvas.style.minHeight); x.editor.setZoom(1.75); x.editor.setZoom(1); expect(x.canvas.style.transform).toBe(''); expect(x.stage.style.width).toBe(`${w}px`); expect(x.stage.style.height).toBe(`${h}px`); }
    finally { x.done(); }
  });
  it('QA20K-00914 zoom selection emits the chosen value in the editor snapshot', async () => {
    const x=await mountCore('Snapshot scale');
    try { x.editor.setZoom(1.25); expect(x.editor.snapshot()?.zoom).toBe(1.25); x.editor.setZoom('fit'); expect(x.editor.snapshot()?.zoom).toBe(x.editor.zoom); }
    finally { x.done(); }
  });
  it('QA20K-00915 changing text after zoom keeps the selected scale on the reflowed canvas', async () => {
    const x=await mountCore('Short note');
    try { x.editor.setZoom(1.5); const view=x.editor.view!; view.dispatch(view.state.tr.insertText(' a longer appended sentence')); expect(x.editor.zoom).toBe(1.5); expect(x.stage.style.width).toBe(`${Number.parseFloat(x.canvas.style.width)*1.5}px`); expect(view.state.doc.textContent).toContain('longer appended sentence'); }
    finally { x.done(); }
  });
  it('QA20K-00916 fit mode recalculates down as the viewport narrows but stops at the minimum', async () => {
    const x=await mountCore('Narrowing preview');
    try { let width=1100; Object.defineProperty(x.root,'clientWidth',{configurable:true,get:()=>width}); x.editor.setZoom('fit'); const wide=x.editor.zoom; width=100; window.dispatchEvent(new Event('resize')); expect(x.editor.zoom).toBe(0.25); expect(x.editor.zoom).toBeLessThan(wide); expect(x.stage.style.width).toBe(`${Number.parseFloat(x.canvas.style.width)*0.25}px`); }
    finally { x.done(); }
  });
});
