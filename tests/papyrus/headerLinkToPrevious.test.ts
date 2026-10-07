// @vitest-environment jsdom
// GOV-088, Word's 「連結到前一節」 while editing a header/footer: a later section shows the previous
// section's header until it is unlinked; unlinked, it gets its own (starting as a copy) that Word
// saves as its own part; linked again, its own goes (after asking when it has content).
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { writeDocx } from '../../src/papyrus/docx/writer';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const SECT = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
});

afterEach(() => vi.unstubAllGlobals());

/** Two sections: the first has header 「甲」, the second none of its own (it shows 「甲」). */
async function twoSections() {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>`
    + `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/>${SECT}</w:sectPr></w:pPr><w:r><w:t>第一節</w:t></w:r></w:p>`
    + `<w:p><w:r><w:t>第二節</w:t></w:r></w:p><w:sectPr>${SECT}</w:sectPr></w:body></w:document>`);
  zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NS}><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:hdr>`);
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>',
    '<Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>',
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
  const host = document.createElement('div');
  document.body.append(host);
  const ed = new DocxEditor(host);
  await ed.open(await zip.generateAsync({ type: 'uint8array' }));
  // One page per section (jsdom lays nothing out).
  const page = (ed as any).firstPage();
  (ed as any).pages = [
    { ...page, section: 0, inSection: 0, number: 1, top: 0 },
    { ...page, section: 1, inSection: 0, number: 2, top: page.height + 20 },
  ];
  return { ed, host };
}

const linkBox = (host: HTMLElement) => host.querySelector('.dx-hf-link input') as HTMLInputElement | null;
const editing = (ed: DocxEditor) => (ed as any).hfSession.view.state.doc.textContent as string;
const type = (ed: DocxEditor, text: string) => ed.run((s, d) => (d?.(s.tr.insertText(text, s.doc.content.size - 1)), true));

async function saved(ed: DocxEditor) {
  const zip = await JSZip.loadAsync(await writeDocx(ed.view!.state.doc, ed.model));
  const document = await zip.file('word/document.xml')!.async('string');
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  const headers: Record<string, string> = {};
  for (const [, id, target] of rels.matchAll(/Id="([^"]+)"[^>]*relationships\/header" Target="([^"]+)"|Id="([^"]+)"/g)) {
    if (id && target) headers[id] = (await zip.file('word/' + target)!.async('string')).replace(/<[^>]+>/g, '');
  }
  // Each section's header text, in order.
  const sections = [...document.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)].map((m) => {
    const id = /<w:headerReference[^>]*w:type="default"[^>]*r:id="([^"]+)"|<w:headerReference[^>]*r:id="([^"]+)"[^>]*w:type="default"/.exec(m[0]);
    return id ? headers[id[1] ?? id[2]] ?? '?' : null;
  });
  return { sections, headerParts: Object.keys(zip.files).filter((f) => /^word\/header\d+\.xml$/.test(f)).length };
}

describe('連結到前一節', () => {
  it('is not offered for the first section', async () => {
    const { ed, host } = await twoSections();
    ed.editHeaderFooter('header', 0);
    expect(editing(ed)).toBe('甲');
    expect(linkBox(host)).toBeNull();
    expect(ed.headerFooterLinked()).toBeNull();
  });

  it('a later section without its own header is linked, and shows the previous one', async () => {
    const { ed, host } = await twoSections();
    ed.editHeaderFooter('header', 1);
    expect(editing(ed)).toBe('甲');
    expect(linkBox(host)!.checked).toBe(true);
    expect(ed.headerFooterLinked()).toBe(true);
    expect(host.querySelector('.dx-hf-label')!.textContent).toContain('第 2 節');
  });

  it('unlinked: its own header, starting as a copy; the first section keeps 「甲」', async () => {
    const { ed, host } = await twoSections();
    ed.editHeaderFooter('header', 1);
    const box = linkBox(host)!;
    box.checked = false;
    box.dispatchEvent(new Event('change'));
    expect(ed.headerFooterLinked()).toBe(false);
    expect(linkBox(host)!.checked).toBe(false);
    expect(editing(ed)).toBe('甲');
    type(ed, '乙');
    expect(editing(ed)).toBe('甲乙');
    ed.closeHeaderFooter();
    expect(ed.isModified()).toBe(true);
    // Section 1 still 「甲」.
    ed.editHeaderFooter('header', 0);
    expect(editing(ed)).toBe('甲');
    ed.closeHeaderFooter();
    expect(await saved(ed)).toEqual({ sections: ['甲', '甲乙'], headerParts: 2 });
  });

  it('linked again: its own header goes (after asking, as it has content) and the previous one shows', async () => {
    const { ed, host } = await twoSections();
    ed.editHeaderFooter('header', 1);
    ed.setHeaderFooterLinked(false);
    type(ed, '乙');
    const ask = vi.fn(() => true);
    vi.stubGlobal('confirm', ask);
    const box = linkBox(host)!;
    box.checked = true;
    box.dispatchEvent(new Event('change'));
    expect(ask).toHaveBeenCalledWith('要刪除這一節自己的頁首，改為和前一節相同嗎？');
    expect(ed.headerFooterLinked()).toBe(true);
    expect(editing(ed)).toBe('甲');
    ed.closeHeaderFooter();
    // Section 2 has no header of its own: Word shows section 1's there.
    expect(await saved(ed)).toEqual({ sections: ['甲', null], headerParts: 1 });
  });

  it("answering no keeps the section's own header", async () => {
    const { ed, host } = await twoSections();
    ed.editHeaderFooter('header', 1);
    ed.setHeaderFooterLinked(false);
    type(ed, '乙');
    vi.stubGlobal('confirm', () => false);
    ed.setHeaderFooterLinked(true);
    expect(ed.headerFooterLinked()).toBe(false);
    expect(linkBox(host)!.checked).toBe(false);
    expect(editing(ed)).toBe('甲乙');
  });

  it('an unlinked header with nothing in it is linked again without asking', async () => {
    const { ed } = await twoSections();
    ed.editHeaderFooter('header', 1);
    ed.setHeaderFooterLinked(false);
    ed.run((s, d) => (d?.(s.tr.delete(1, s.doc.content.size - 1)), true));
    const ask = vi.fn(() => false);
    vi.stubGlobal('confirm', ask);
    ed.setHeaderFooterLinked(true);
    expect(ask).not.toHaveBeenCalled();
    expect(ed.headerFooterLinked()).toBe(true);
  });

  it('is refused while tracking changes, like other section settings', async () => {
    const notices: string[] = [];
    const { ed } = await twoSections();
    (ed as any).options.onNotice = (m: string) => notices.push(m);
    ed.setTrackChanges(true);
    ed.editHeaderFooter('header', 1);
    ed.setHeaderFooterLinked(false);
    expect(ed.headerFooterLinked()).toBe(true);
    expect(notices[0]).toContain('追蹤修訂');
  });
});

// Opt-in: DOCX_EXPORT=<out> writes the results for Word to confirm (section 2's header text and LinkToPrevious).
describe.skipIf(!process.env.DOCX_EXPORT)('連結到前一節 in Word', () => {
  it('writes unlinked.docx and relinked.docx', async () => {
    const { mkdirSync, writeFileSync } = await import('node:fs');
    const out = process.env.DOCX_EXPORT!;
    mkdirSync(out, { recursive: true });
    let { ed } = await twoSections();
    ed.editHeaderFooter('header', 1);
    ed.setHeaderFooterLinked(false);
    type(ed, '乙');
    ed.closeHeaderFooter();
    writeFileSync(`${out}/unlinked.docx`, await writeDocx(ed.view!.state.doc, ed.model));
    vi.stubGlobal('confirm', () => true);
    ed.editHeaderFooter('header', 1);
    ed.setHeaderFooterLinked(true);
    ed.closeHeaderFooter();
    writeFileSync(`${out}/relinked.docx`, await writeDocx(ed.view!.state.doc, ed.model));
    ({ ed } = await twoSections());
    writeFileSync(`${out}/original.docx`, await writeDocx(ed.view!.state.doc, ed.model));
  });
});
