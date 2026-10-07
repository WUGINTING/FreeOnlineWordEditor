// @vitest-environment jsdom
// The interface's languages: the lookup (src/papyrus/i18n.ts), that every interface text of the
// source is translated (tests/papyrus/i18n/scan.mjs), and the editor shown in another language.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { nextTick } from 'vue';
import { addMessages, DEFAULT_LOCALE, getLocale, locales, matchLocale, onLocaleChange, setLocale, tl } from '../../src/papyrus/i18n';
import { createDocxEditor, type DocxEditorHandle } from '../../src/papyrus';
// @ts-expect-error a plain script (also run from the command line), without declarations
import { check } from './i18n/scan.mjs';

afterEach(() => {
  setLocale(DEFAULT_LOCALE);
});

describe('tl: the lookup', () => {
  it('gives the source text in the language the interface is written in', () => {
    expect(getLocale()).toBe('zh-TW');
    expect(tl('檔案')).toBe('檔案');
    expect(tl('第 {0} 頁，共 {1} 頁', 2, 5)).toBe('第 2 頁，共 5 頁');
  });

  it('gives the language’s text after setLocale, with the values in their places', () => {
    expect(setLocale('en')).toBe('en');
    expect(tl('檔案')).toBe('File');
    expect(tl('尋找')).toBe('Find');
    setLocale('zh-CN');
    expect(tl('檔案')).toBe('文件');
    expect(tl('列印')).toBe('打印');
  });

  it('shows a text the language does not have as it is written', () => {
    setLocale('en');
    expect(tl('這句話沒有翻譯 {0}', 3)).toBe('這句話沒有翻譯 3');
  });

  it('leaves a place without a value as it is, and repeats a value used twice', () => {
    expect(tl('{0} 與 {1}', '甲')).toBe('甲 與 {1}');
    expect(tl('{0}、{0}', '甲')).toBe('甲、甲');
  });

  it('matches language tags: the tag itself, the script of Chinese, a language’s region, English otherwise', () => {
    expect(locales()).toEqual(expect.arrayContaining(['zh-TW', 'zh-CN', 'en']));
    expect(matchLocale('zh-TW')).toBe('zh-TW');
    expect(matchLocale('zh-tw')).toBe('zh-TW');
    expect(matchLocale('zh-HK')).toBe('zh-TW');
    expect(matchLocale('zh-Hant')).toBe('zh-TW');
    expect(matchLocale('zh')).toBe('zh-TW');
    expect(matchLocale('zh-CN')).toBe('zh-CN');
    expect(matchLocale('zh-Hans-CN')).toBe('zh-CN');
    expect(matchLocale('zh-SG')).toBe('zh-CN');
    expect(matchLocale('en-US')).toBe('en');
    expect(matchLocale('en-GB')).toBe('en');
    expect(matchLocale('fr-FR')).toBe('en');
    expect(matchLocale('')).toBe('zh-TW');
    expect(matchLocale(null)).toBe('zh-TW');
  });

  it('takes texts of your own: for a language it has, and a language it has not', () => {
    addMessages('de', { 檔案: 'Datei' });
    expect(locales()).toContain('de');
    expect(matchLocale('de-AT')).toBe('de');
    setLocale('de');
    expect(tl('檔案')).toBe('Datei');
    expect(tl('尋找')).toBe('尋找');
    setLocale('en');
    addMessages('en', { 這是我的字: 'My own text' });
    expect(tl('這是我的字')).toBe('My own text');
    expect(tl('檔案')).toBe('File');
  });

  it('tells listeners when the language changes, not when it is set to what it is', () => {
    let calls = 0;
    const stop = onLocaleChange(() => calls++);
    setLocale('en');
    setLocale('en');
    expect(calls).toBe(1);
    stop();
    setLocale('zh-CN');
    expect(calls).toBe(1);
  });
});

describe('every interface text is translated', () => {
  it('the source has no Chinese text that is neither translatable nor listed as not interface text', async () => {
    const result = await check();
    expect(result.problems).toEqual([]);
    expect(result.unusedExempt).toEqual([]);
    expect(result.keys.length).toBeGreaterThan(500);
  });

  it.each(['en', 'zh-CN'])('%s has every text, nothing the source no longer uses, and the same {n} places', async (locale) => {
    const { languages } = await check();
    expect(languages[locale]).toBeDefined();
    expect(languages[locale].missing).toEqual([]);
    expect(languages[locale].unused).toEqual([]);
    expect(languages[locale].places).toEqual([]);
  });

  it('English has no Chinese left in it, apart from quoted document text', async () => {
    const en = (await import('../../src/papyrus/locales/en')).default;
    // Characters a translation may keep: none of the interface's words.
    const left = Object.entries(en).filter(([, text]) => /[㐀-鿿]/.test(text));
    expect(left.map(([key]) => key)).toEqual([]);
  });
});

describe('the editor in another language', () => {
  beforeAll(() => {
    (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
    Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
    Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
    (Element.prototype as any).scrollIntoView ??= () => {};
  });

  let host: HTMLElement;
  let handle: DocxEditorHandle | null = null;
  async function create(locale?: string): Promise<DocxEditorHandle> {
    host = document.createElement('div');
    document.body.append(host);
    let ready = false;
    handle = createDocxEditor(host, { locale, onReady: () => (ready = true) });
    for (let i = 0; i < 200 && !ready; i++) await new Promise((r) => setTimeout(r, 10));
    return handle;
  }
  afterEach(() => {
    handle?.destroy();
    handle = null;
    host?.remove();
  });

  const tabs = () => [...host.querySelectorAll('[role="tab"]')].map((el) => el.textContent!.trim());
  /** Chinese characters in what the ribbon shows and says (texts, tooltips, labels). */
  const chineseLeft = () => {
    const found = new Set<string>();
    const ribbon = host.querySelector('.dx-vue')!;
    const walker = document.createTreeWalker(ribbon, NodeFilter.SHOW_TEXT);
    // The document itself is not interface.
    const inDocument = (node: Node) => !!(node.parentElement?.closest('.ProseMirror, .dx-doc'));
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (!inDocument(node) && /[㐀-鿿]/.test(node.textContent ?? '')) found.add(node.textContent!.trim());
    }
    for (const el of ribbon.querySelectorAll('[title], [aria-label], [placeholder]')) {
      if (el.closest('.ProseMirror, .dx-doc')) continue;
      for (const name of ['title', 'aria-label', 'placeholder']) {
        const value = el.getAttribute(name);
        if (value && /[㐀-鿿]/.test(value)) found.add(`${name}=${value}`);
      }
    }
    return [...found];
  };

  it('is in Traditional Chinese unless told otherwise', async () => {
    await create();
    expect(tabs()).toEqual(expect.arrayContaining(['常用', '插入', '校閱']));
  });

  it('createDocxEditor({ locale: "en" }) shows the ribbon in English, with no Chinese left on it', async () => {
    await create('en');
    expect(getLocale()).toBe('en');
    expect(tabs()).toEqual(expect.arrayContaining(['Home', 'Insert', 'Review']));
    // Font and style names are the document's and the computer's own: not interface words.
    const left = chineseLeft().filter((text) => !/體|黑|明|宋|楷/.test(text));
    expect(left).toEqual([]);
  });

  it('a browser’s language tag is enough: en-US, zh-Hans-CN', async () => {
    await create('zh-Hans-CN');
    expect(getLocale()).toBe('zh-CN');
    expect(tabs()).toEqual(expect.arrayContaining(['开始', '插入', '审阅']));
  });

  it('what is on screen follows setLocale', async () => {
    await create();
    expect(tabs()).toContain('常用');
    setLocale('en');
    await nextTick();
    expect(tabs()).toContain('Home');
    expect(tabs()).not.toContain('常用');
  });
});
