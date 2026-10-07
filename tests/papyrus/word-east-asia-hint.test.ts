// @vitest-environment jsdom
// w:rFonts/@w:hint="eastAsia" on screen and on paper. In zh-TW Word a run whose Latin font differs
// from its East Asian font (Times New Roman + 標楷體) draws the characters both kinds of font have
// (①, ※, ○, ±, ×, →, ℃, “ ”, Greek …) in the Latin font, unless the run carries the eastAsia
// hint: then they are drawn in the East Asian font. Latin letters stay in the Latin font either way.
// The editor does it with CSS only: for each East Asian font in use an @font-face alias of that
// local font limited to those characters (unicode-range), put first in hinted runs (data-ea-hint).
// No element goes around characters, so typing (and an input method) sees the same text as before.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { insertSymbol, setMark } from '../../src/papyrus/editor/commands';
import { schema } from '../../src/papyrus/editor/schema';
import { HINT_RULE, HINT_UNICODE_RANGE, eastAsiaFontFace, inHintRanges } from '../../src/papyrus/docx/eastAsiaFonts';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});

const editors: DocxEditor[] = [];
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy();
});

const FONTS = 'w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="DFKai-SB"';
const hinted = (text: string) => `<w:r><w:rPr><w:rFonts ${FONTS} w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const plain = (text: string) => `<w:r><w:rPr><w:rFonts ${FONTS}/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

async function zipOf(body: string, header?: string): Promise<Uint8Array> {
  const zip = blankPackage();
  const sect = header ? '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/></w:sectPr>' : '<w:sectPr/>';
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${body}${sect}</w:body></w:document>`);
  if (header) {
    zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}>${header}</w:hdr>`);
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
      '<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
    const types = await zip.file('[Content_Types].xml')!.async('string');
    zip.file('[Content_Types].xml', types.replace('</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
  }
  return zip.generateAsync({ type: 'uint8array' });
}
async function open(body: string, header?: string) {
  const host = document.createElement('div');
  document.body.append(host);
  const editor = new DocxEditor(host);
  await editor.open(await zipOf(body, header));
  editors.push(editor);
  await settle();
  return editor;
}
/** Fonts that come into use while text is drawn reach the editors' CSS a microtask later. */
const settle = () => new Promise((r) => setTimeout(r, 0));
/** The editor's own <style> (the document's styles and the East Asian font aliases). */
const css = (editor: DocxEditor) => (editor as any).root.querySelector(':scope > style').textContent as string;
function posOf(view: EditorView, text: string, n = 0): number {
  let found = -1;
  view.state.doc.descendants((node, pos) => {
    if (found < 0 && node.isText && node.text!.includes(text)) found = pos + node.text!.indexOf(text) + n;
    return found < 0;
  });
  if (found < 0) throw new Error(`no ${text}`);
  return found;
}
const hintedRuns = (root: ParentNode) => Array.from(root.querySelectorAll('[data-ea-hint]')).map((e) => e.textContent);

describe('w:hint="eastAsia": the East Asian font for shared characters, by CSS', () => {
  it('a hinted run is marked data-ea-hint, its text in one piece; an unhinted run is not marked', async () => {
    const editor = await open(`<w:p>${hinted('公文①AB')}</w:p><w:p>${plain('公文①')}</w:p>`);
    const dom = editor.view!.dom;
    expect(hintedRuns(dom)).toEqual(['公文①AB']);
    // Nothing around characters: the run's text is one text node inside its font element.
    const run = dom.querySelector('[data-ea-hint]')!;
    const texts: Node[] = [];
    const walker = document.createTreeWalker(run, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) texts.push(n);
    expect(texts.map((t) => t.textContent)).toEqual(['公文①AB']);
    // The font element names the East Asian font's alias next to the font.
    expect(run.querySelector('[style*="font-family"]')!.getAttribute('style')).toMatch(/--dx-font-h:\s*"dx-ea DFKai-SB"/);
  });

  it('an @font-face alias per East Asian font in use (both names as local sources), and the rule for hinted runs', async () => {
    const editor = await open(`<w:p>${hinted('①')}</w:p>`);
    const style = css(editor);
    // The document default (標楷體: a new document is a 公文 page, persona-300) and the run's own font.
    expect(style).toContain('@font-face{font-family:"dx-ea 標楷體";src:local("標楷體"),local("DFKai-SB")');
    expect(style).toContain('@font-face{font-family:"dx-ea DFKai-SB";src:local("DFKai-SB"),local("標楷體");');
    expect(style).toContain(`unicode-range:${HINT_UNICODE_RANGE}}`);
    expect(style).toContain(HINT_RULE);
    expect(HINT_RULE).toMatch(/^\[data-ea-hint\],\[data-ea-hint\] \[style\*="font-family"\]\{font-family:var\(--dx-font-h\),var\(--dx-font-l\)/);
    // Every font list names its alias too (the document defaults here).
    // (the East Asian list with the font's other names and similar fonts, persona-300 B-8).
    expect(style).toMatch(/--dx-font-e:\s*"標楷體", [^;]*;\s*--dx-font-h:\s*"dx-ea 標楷體"/);
  });

  it('the alias covers the shared characters and not Latin letters, spaces or combining marks', () => {
    for (const ch of ['①', '※', '○', '±', '×', '→', '℃', '“', '”', '—', '…', '§', '°', 'α', 'Ⅰ', 'ˇ']) expect(inHintRanges(ch.codePointAt(0)!), ch).toBe(true);
    for (const ch of ['A', 'z', 'é', ' ', '\u00a0', '¢', '£', '¥', '©', '«', '®', 'µ', '»', '\u0301', '\u20d0', '\u20ff', '、', '公'])
      expect(inHintRanges(ch.codePointAt(0)!), JSON.stringify(ch)).toBe(false);
    expect(HINT_UNICODE_RANGE).toContain('U+A7-A8');
    expect(HINT_UNICODE_RANGE).toContain('U+2070-20CF, U+2100-2E7F');
  });

  it('font names reaching the CSS are sanitised', () => {
    const face = eastAsiaFontFace('Evil"}body{color:red');
    expect(face).toContain('font-family:"dx-ea Evilbodycolorred"');
    expect(face).not.toMatch(/[{]body|"}/);
    expect(eastAsiaFontFace('"";{}')).toBe('');
  });

  it('follows edits: an inserted symbol gets a hinted run; a new East Asian font gets its alias', async () => {
    const editor = await open(`<w:p>${plain('乙')}</w:p>`);
    const view = editor.view!;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, posOf(view, '乙', 1))));
    editor.run(insertSymbol('→', editor.model.styles));
    expect(hintedRuns(view.dom)).toEqual(['→']);
    expect(css(editor)).not.toContain('dx-ea Microsoft JhengHei');
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, posOf(view, '乙'), posOf(view, '乙') + 2)));
    editor.run(setMark(schema.marks.font, { family: 'Microsoft JhengHei', eastAsia: 'Microsoft JhengHei' }));
    await settle();
    expect(css(editor)).toContain('@font-face{font-family:"dx-ea Microsoft JhengHei";src:local("Microsoft JhengHei"),local("微軟正黑體")');
    expect(hintedRuns(view.dom)).toEqual(['→']);
  });

  it('the static header copy on each page and the printout (printHtml) keep the mark and the CSS', async () => {
    const editor = await open(`<w:p>${hinted('本文※')}</w:p>`, `<w:p>${hinted('頁首①A')}</w:p>`);
    // jsdom lays nothing out: one page holding every block (all measured at 0).
    (editor as any).pages = [{ ...(editor as any).firstPage(), section: 0, inSection: 0, number: 1, top: -1 }];
    (editor as any).renderPages((editor as any).pages);
    expect(hintedRuns(document.querySelector('.dx-page .dx-header')!)).toEqual(['頁首①A']);
    const html = editor.printHtml('t')!;
    expect(html).toContain('@font-face{font-family:"dx-ea DFKai-SB"');
    expect(html).toContain(HINT_RULE);
    const body = html.slice(html.indexOf('<body'));
    expect(body).toMatch(/data-ea-hint="">(?:<[^>]+>)*頁首①A/);
    expect(body).toMatch(/data-ea-hint="">(?:<[^>]+>)*本文※/);
  });
});
