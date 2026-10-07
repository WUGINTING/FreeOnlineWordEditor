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

describe('QA20K zoom view use cases 00642–00647', () => {
  it('QA20K-00642 exposes labeled preset zoom choices on a keyboard-focusable native selector', async () => {
    instances.length = 0;
    const Editor = await compiledUi();
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp({ render: () => h(Editor, { src: null, editable: true, toolbar: false }) });
    app.mount(host);
    await flush();
    const select = host.querySelector('select[aria-label="縮放"]') as HTMLSelectElement;
    expect(select).toBeTruthy();
    expect(select.tabIndex).toBe(0);
    expect(Array.from(select.options).map((o) => o.value)).toEqual(['0.5', '0.75', '0.9', '1', '1.25', '1.5', '2', 'fit']);
    select.focus();
    expect(document.activeElement).toBe(select);
    const arrowDown = new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true, cancelable: true });
    select.dispatchEvent(arrowDown);
    expect(arrowDown.defaultPrevented).toBe(false);
    expect(select.getAttribute('aria-label')).toBe('縮放');
    app.unmount();
    host.remove();
  });

  it('QA20K-00643 fit-to-width recalculates page scale when the viewport is resized', async () => {
    const x = await mountCore('A4 portrait memo');
    try {
      let viewportWidth = 632;
      Object.defineProperty(x.root, 'clientWidth', { configurable: true, get: () => viewportWidth });
      const pageCanvasWidth = Number.parseFloat(x.canvas.style.width);
      expect(pageCanvasWidth).toBeGreaterThan(0);
      x.editor.setZoom('fit');
      expect(x.editor.zoom).toBeCloseTo((viewportWidth - 32) / pageCanvasWidth, 6);
      viewportWidth = 832;
      window.dispatchEvent(new Event('resize'));
      expect(x.editor.zoom).toBeCloseTo((viewportWidth - 32) / pageCanvasWidth, 6);
      expect(x.stage.style.width).toBe(`${pageCanvasWidth * x.editor.zoom}px`);
    } finally { x.done(); }
  });

  it('QA20K-00644 magnification scales the page canvas and viewing stage without changing paper dimensions', async () => {
    const x = await mountCore('Quarterly report');
    try {
      const page = x.host.querySelector('.dx-page') as HTMLElement;
      const paperSize = [page.style.width, page.style.height];
      const width = Number.parseFloat(x.canvas.style.width);
      const height = Number.parseFloat(x.canvas.style.minHeight);
      x.editor.setZoom(1.5);
      expect(x.editor.snapshot()?.zoom).toBe(1.5);
      expect(x.canvas.style.transform).toBe('scale(1.5)');
      expect(x.stage.style.width).toBe(`${width * 1.5}px`);
      expect(x.stage.style.height).toBe(`${height * 1.5}px`);
      expect([page.style.width, page.style.height]).toEqual(paperSize);
    } finally { x.done(); }
  });

  it('QA20K-00645 screen magnification leaves printable page geometry unchanged', async () => {
    const x = await mountCore('Invoice total 275');
    try {
      const geometry = (html: string | null) => {
        expect(html).not.toBeNull();
        const parsed = new DOMParser().parseFromString(html!, 'text/html');
        return Array.from(parsed.querySelectorAll('.dx-print-page')).map((page) => {
          const box = page as HTMLElement;
          return [box.style.width, box.style.height, box.querySelector('.dx-page')?.getAttribute('style') ?? ''];
        });
      };
      const before = geometry(x.editor.printHtml('Invoice'));
      x.editor.setZoom(1.5);
      const afterHtml = x.editor.printHtml('Invoice');
      expect(geometry(afterHtml)).toEqual(before);
      expect(afterHtml).not.toContain('scale(1.5)');
    } finally { x.done(); }
  });

  it('QA20K-00646 fit-to-width respects the minimum zoom in a narrow preview pane', async () => {
    const x = await mountCore('Narrow preview pane');
    try {
      Object.defineProperty(x.root, 'clientWidth', { configurable: true, value: 100 });
      const width = Number.parseFloat(x.canvas.style.width);
      const paperWidth = (x.host.querySelector('.dx-page') as HTMLElement).style.width;
      expect((100 - 32) / width).toBeLessThan(0.25);
      x.editor.setZoom('fit');
      expect(x.editor.zoom).toBe(0.25);
      expect(x.canvas.style.transform).toBe('scale(0.25)');
      expect(x.stage.style.width).toBe(`${width * 0.25}px`);
      expect((x.host.querySelector('.dx-page') as HTMLElement).style.width).toBe(paperWidth);
    } finally { x.done(); }
  });

  it('QA20K-00647 keeps zoom states isolated between independent document editor instances', async () => {
    const first = await mountCore('Document alpha');
    const second = await mountCore('Document beta');
    try {
      let secondWidth = 632;
      Object.defineProperty(second.root, 'clientWidth', { configurable: true, get: () => secondWidth });
      expect(first.editor.zoom).toBe(1);
      expect(second.editor.zoom).toBe(1);
      first.editor.setZoom(1.5);
      expect(first.editor.zoom).toBe(1.5);
      expect(second.editor.zoom).toBe(1);
      second.editor.setZoom('fit');
      expect(second.editor.zoom).toBeCloseTo((secondWidth - 32) / Number.parseFloat(second.canvas.style.width), 6);
      expect(first.editor.zoom).toBe(1.5);
      secondWidth = 432;
      window.dispatchEvent(new Event('resize'));
      expect(second.editor.zoom).toBeCloseTo((secondWidth - 32) / Number.parseFloat(second.canvas.style.width), 6);
      expect(first.editor.zoom).toBe(1.5);
      expect(first.editor.view!.state.doc.textContent).toBe('Document alpha');
      expect(second.editor.view!.state.doc.textContent).toBe('Document beta');
    } finally { first.done(); second.done(); }
  });
});
