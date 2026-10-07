// @vitest-environment jsdom
// GOV-171: an old browser (IE, Edge in IE mode, a Chrome from before 2021) is told what to do
// instead of getting a blank page. GOV-186: the browser's page translation leaves the document alone.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { blankPackage } from '../src/papyrus/docx/template';
import { DocxEditor } from '../src/papyrus/editor/core';

const script = readFileSync(resolve(__dirname, '../public/browser-check.js'), 'utf8');
const html = readFileSync(resolve(__dirname, '../index.html'), 'utf8');

/** Runs public/browser-check.js as the browser would, the page already loaded (jsdom). */
async function runCheck() {
  new Function(script)();
  await new Promise((r) => setTimeout(r, 0));
  return document.getElementById('dx-browser-notice');
}

afterEach(() => {
  document.getElementById('dx-browser-notice')?.remove();
  delete (window as any).__appStarted;
  delete (document as any).documentMode;
});

describe('old-browser notice (public/browser-check.js)', () => {
  beforeAll(() => {
    // jsdom does not implement noModule; every browser that runs modules does.
    if (!('noModule' in HTMLScriptElement.prototype)) {
      Object.defineProperty(HTMLScriptElement.prototype, 'noModule', { value: false, configurable: true });
    }
    // Nor CSS.supports or inert, which every browser the application supports has (Chrome/Edge 105).
    if (!(window as any).CSS?.supports) {
      (window as any).CSS = { ...((window as any).CSS ?? {}), supports: (q: string) => q === 'selector(:has(*))' };
    }
    if (!('inert' in HTMLElement.prototype)) {
      Object.defineProperty(HTMLElement.prototype, 'inert', { value: false, configurable: true, writable: true });
    }
  });

  /** Runs the check with `feature` missing only while the script checks (as in an older browser). */
  async function runWithout(feature: 'has' | 'inert') {
    (window as any).__appStarted = true;
    const supports = (window as any).CSS.supports;
    const inert = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'inert');
    if (feature === 'has') (window as any).CSS.supports = () => false;
    else delete (HTMLElement.prototype as any).inert;
    try {
      new Function(script)();
    } finally {
      (window as any).CSS.supports = supports;
      if (inert) Object.defineProperty(HTMLElement.prototype, 'inert', inert);
    }
    await new Promise((r) => setTimeout(r, 0));
    return document.getElementById('dx-browser-notice');
  }

  it('warns in a browser without the CSS selector :has() (Chrome/Edge before 105): the layout depends on it', async () => {
    const notice = await runWithout('has');
    expect(notice).not.toBeNull();
    expect(notice!.textContent).toContain('您的瀏覽器版本較舊');
    expect(notice!.textContent).toContain('Microsoft Edge 或 Google Chrome');
  });

  it('warns in a browser without the inert attribute (Chrome/Edge before 102)', async () => {
    const notice = await runWithout('inert');
    expect(notice!.textContent).toContain('您的瀏覽器版本較舊');
  });

  it('names the minimum versions in its description (Chrome/Edge 105)', () => {
    expect(script).toMatch(/Chrome\/Edge 105/);
    expect(script).toContain("CSS.supports('selector(:has(*))')");
  });

  it('is loaded by index.html as a classic script (it must run where modules cannot), before the app', () => {
    // At the site's base path (Vite's %BASE_URL%, for a deployment in a folder of a site).
    const tag = html.match(/<script src="(?:%BASE_URL%|\/)browser-check\.js"><\/script>/);
    expect(tag).not.toBeNull();
    expect(html.indexOf('/demo/main.ts')).toBeGreaterThan(html.indexOf('browser-check.js'));
  });

  it('says nothing in a current browser where the application started', async () => {
    (window as any).__appStarted = true;
    expect(await runCheck()).toBeNull();
  });

  it('tells the user when the application could not start, with the address to open elsewhere', async () => {
    const notice = await runCheck();
    expect(notice).not.toBeNull();
    expect(notice!.getAttribute('role')).toBe('alert');
    expect(notice!.textContent).toContain('這個瀏覽器無法開啟本系統');
    expect(notice!.textContent).toContain('Microsoft Edge 或 Google Chrome');
    expect(notice!.textContent).toContain('資訊人員');
    expect(notice!.querySelector('input')!.value).toBe(location.href);
    // Nothing else works here: no "continue" button.
    expect(notice!.querySelector('button')).toBeNull();
  });

  it('warns about an old browser that still started the application, and can be put away', async () => {
    const hasOwn = Object.hasOwn;
    (window as any).__appStarted = true;
    // Missing only while the script checks (jsdom itself needs it to build the notice later).
    delete (Object as any).hasOwn;
    try {
      new Function(script)();
    } finally {
      (Object as any).hasOwn = hasOwn;
    }
    await new Promise((r) => setTimeout(r, 0));
    const notice = document.getElementById('dx-browser-notice');
    expect(notice!.textContent).toContain('您的瀏覽器版本較舊');
    notice!.querySelector('button')!.click();
    expect(document.getElementById('dx-browser-notice')).toBeNull();
  });

  it('names Internet Explorer mode when the page is opened in it', async () => {
    (document as any).documentMode = 11;
    const notice = await runCheck();
    expect(notice!.textContent).toContain('IE 模式');
  });

  it('shows a single notice when the check runs twice', async () => {
    await runCheck();
    await runCheck();
    expect(document.querySelectorAll('#dx-browser-notice').length).toBe(1);
  });

  it('is written for old browsers: no arrow functions, let/const, classes or template strings', () => {
    const code = script.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '');
    expect(code).not.toMatch(/=>|\blet\b|\bconst\b|\bclass\b|`|\?\.|\?\?/);
  });
});

describe('page translation (GOV-186)', () => {
  beforeAll(() => {
    (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
    Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
    Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  });

  it('the page says it is in Traditional Chinese', () => {
    expect(html).toMatch(/<html lang="zh-Hant-TW"/);
  });

  it('the document and the header/footer being edited are marked translate="no"', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const editor = new DocxEditor(host);
    await editor.open(await blankPackage().generateAsync({ type: 'uint8array' }));
    expect(editor.view!.dom.getAttribute('translate')).toBe('no');
    editor.editHeaderFooter('header');
    const hf = host.querySelector('.dx-hf-doc');
    expect(hf).not.toBeNull();
    expect(hf!.getAttribute('translate')).toBe('no');
    editor.destroy();
    host.remove();
  });
});
