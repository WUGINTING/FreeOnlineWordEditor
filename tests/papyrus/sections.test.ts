// Documents with several sections: each keeps its own page setup and headers/footers,
// and editing them changes only the section concerned.
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import {
  documentSections, sectionHeaderFooter, variantFor, withPageSetup,
} from '../../src/papyrus/docx/sections';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { schema } from '../../src/papyrus/editor/schema';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class {
    observe() {}
    unobserve() {}
    disconnect() {}
  };
  const empty = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= empty;
  Range.prototype.getBoundingClientRect ??= zero;
});

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const HDR = (text: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`;
const PORTRAIT = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';
const LANDSCAPE = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';

/** The QA layout: portrait section with header A, then a landscape section with header B, and a third section inheriting B. */
async function threeSections(extraFirstRefs = '') {
  const zip = blankPackage();
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>` +
      `<w:p><w:r><w:t>A1</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="hA"/>${extraFirstRefs}${PORTRAIT}</w:sectPr></w:pPr><w:r><w:t>A2</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="hB"/>${LANDSCAPE}<w:titlePg/></w:sectPr></w:pPr><w:r><w:t>B1</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>C1</w:t></w:r></w:p>` +
      `<w:sectPr>${PORTRAIT}</w:sectPr></w:body></w:document>`,
  );
  zip.file('word/headerA.xml', HDR('HEADER_A'));
  zip.file('word/headerB.xml', HDR('HEADER_B'));
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="hA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerA.xml"/>' +
        '<Relationship Id="hB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerB.xml"/></Relationships>',
    ),
  );
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file(
    '[Content_Types].xml',
    ct.replace(
      '</Types>',
      '<Override PartName="/word/headerA.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '<Override PartName="/word/headerB.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
    ),
  );
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  return { bytes, ...(await readDocx(bytes)) };
}

const text = (hf: { doc: { textContent: string } } | undefined) => hf?.doc.textContent;

describe('sections', () => {
  it('each section has its own page size and orientation, and blocks belong to their section', async () => {
    const { doc, model } = await threeSections();
    const s = documentSections(doc, model);
    expect(s.map((x) => [x.page.width, x.page.height])).toEqual([[11906, 16838], [16838, 11906], [11906, 16838]]);
    expect(s.map((x) => [x.firstBlock, x.lastBlock])).toEqual([[0, 1], [2, 2], [3, 3]]);
    expect(s.map((x) => x.titlePage)).toEqual([false, true, false]);
  });

  it('headers are read for every section, and a section without one shows the previous section’s', async () => {
    const { doc, model } = await threeSections();
    const s = documentSections(doc, model);
    expect(model.headerFooters.map((h) => h.relId).sort()).toEqual(['hA', 'hB']);
    expect(text(sectionHeaderFooter(model, s, 0, 'header', 'default'))).toBe('HEADER_A');
    expect(text(sectionHeaderFooter(model, s, 1, 'header', 'default'))).toBe('HEADER_B');
    // The third section has no reference: linked to the previous section.
    expect(text(sectionHeaderFooter(model, s, 2, 'header', 'default'))).toBe('HEADER_B');
    // "Different first page" in section 2 without a first-page header: its first page has none.
    expect(variantFor(s[1], 0, 3, false)).toBe('first');
    expect(sectionHeaderFooter(model, s, 1, 'header', 'first')).toBeUndefined();
  });

  it('saving without edits keeps every section break byte for byte', async () => {
    const { bytes, doc, model } = await threeSections();
    const before = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    const after = await (await JSZip.loadAsync(await writeDocx(doc, model))).file('word/document.xml')!.async('string');
    const breaks = (xml: string) => xml.match(/<w:sectPr>.*?<\/w:sectPr>/g);
    expect(breaks(after)).toEqual(breaks(before));
  });

  it('page setup changes only the chosen section, keeping everything else in its w:sectPr', () => {
    const xml = `<w:sectPr ${W}><w:headerReference w:type="default" r:id="hA"/>${PORTRAIT}<w:cols w:space="425"/><w:titlePg/></w:sectPr>`;
    const out = withPageSetup(xml, {
      width: 16838, height: 11906, marginTop: 720, marginBottom: 720, marginLeft: 1000, marginRight: 1000, header: 500, footer: 500,
    });
    expect(out).toContain('<w:headerReference w:type="default" r:id="hA"/>');
    expect(out).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
    expect(out).toContain('w:top="720"');
    expect(out).toContain('<w:cols w:space="425"/><w:titlePg/>');
  });
});

describe('sections in the editor', () => {
  async function editor() {
    const { bytes } = await threeSections();
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host);
    await ed.open(bytes);
    return { ed, host };
  }
  const at = (ed: DocxEditor, pos: number) => {
    const v = ed.view!;
    v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, pos)));
  };

  it('page setup of the cursor’s section is written into that section only', async () => {
    const { ed, host } = await editor();
    at(ed, 2); // "A1": first section
    expect(ed.cursorSection().index).toBe(0);
    const page = { ...ed.cursorSection().page, marginLeft: 2000 };
    ed.setPageSetup(page, 'section');
    const xml = await (await JSZip.loadAsync(await ed.save())).file('word/document.xml')!.async('string');
    const sects = xml.match(/<w:sectPr>.*?<\/w:sectPr>/g)!;
    expect(sects[0]).toContain('w:left="2000"');
    expect(sects[1]).toContain('w:left="1440"');
    expect(sects[2]).toContain('w:left="1440"');
    expect(ed.isModified()).toBe(true);
    ed.destroy();
    host.remove();
  });

  it('inserting a section break keeps both halves in the same page setup and headers; removing it merges them', async () => {
    const { ed, host } = await editor();
    const c1 = ed.view!.state.doc.child(0).nodeSize + ed.view!.state.doc.child(1).nodeSize + ed.view!.state.doc.child(2).nodeSize;
    at(ed, c1 + 2); // inside "C1", the last section
    expect(ed.sections()).toHaveLength(3);
    expect(ed.insertSectionBreak()).toBe(true);
    const s = ed.sections();
    expect(s).toHaveLength(4);
    expect(s[2].page).toEqual(s[3].page);
    expect(s[2].lastBlock).toBe(3);
    expect(ed.view!.state.doc.child(3).textContent).toBe('C');
    expect(ed.view!.state.doc.child(4).textContent).toBe('1');
    // Undo brings the document back exactly.
    ed.undo();
    expect(ed.sections()).toHaveLength(3);
    expect(ed.isModified()).toBe(false);
    // Removing the break between section 1 and 2: the text of section 1 joins section 2.
    at(ed, 6); // inside "A2" (positions 5–7), which ends section 1
    expect(ed.snapshot()!.atSectionBreak).toBe(true);
    expect(ed.removeSectionBreak()).toBe(true);
    expect(ed.sections()).toHaveLength(2);
    expect(ed.sections()[0].page.width).toBe(16838);
    expect(ed.view!.state.doc.textContent).toBe('A1A2B1C1');
    ed.destroy();
    host.remove();
  });

  it('a header made for an earlier section is referenced by that section once it has content', async () => {
    const { ed, host } = await editor();
    // Section 2 has "different first page" but no first-page header: open it on page 2 (section 2's first page).
    const sections = ed.sections();
    (ed as any).pages = [
      { ...(ed as any).firstPage(), section: 0, inSection: 0, number: 1, top: 0 },
      { ...(ed as any).firstPage(), section: 1, inSection: 0, number: 2, top: 1200 },
    ];
    ed.editHeaderFooter('header', 1);
    const hf = ed.model.headerFooters.find((h) => h.pending)!;
    expect(hf).toBeTruthy();
    // Opening alone changes nothing.
    expect(ed.isModified()).toBe(false);
    ed.run((state, dispatch) => (dispatch?.(state.tr.insertText('FIRST_B')), true));
    expect(hf.pending).toBe(false);
    const s2 = ed.sections()[1];
    expect(s2.refs.header.first).toBe(hf.relId);
    expect(sections[1].refs.header.default).toBe('hB');
    ed.closeHeaderFooter();

    const out = await JSZip.loadAsync(await ed.save());
    const xml = await out.file('word/document.xml')!.async('string');
    const rels = await out.file('word/_rels/document.xml.rels')!.async('string');
    const sect2 = xml.match(/<w:sectPr>.*?<\/w:sectPr>/g)![1];
    expect(sect2).toContain(`<w:headerReference w:type="first" r:id="${hf.relId}"/>`);
    const target = new RegExp(`Id="${hf.relId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] ?? new RegExp(`Target="([^"]+)"[^>]*Id="${hf.relId}"`).exec(rels)?.[1];
    expect(target).toBeTruthy();
    expect(await out.file('word/' + target)!.async('string')).toContain('FIRST_B');
    // The other sections' headers are untouched.
    expect(await out.file('word/headerA.xml')!.async('string')).toBe(HDR('HEADER_A'));
    expect(await out.file('word/headerB.xml')!.async('string')).toBe(HDR('HEADER_B'));
    expect(ed.sections()[1].refs.header.default).toBe('hB');
    expect(text(sectionHeaderFooter(ed.model, ed.sections(), 2, 'header', 'default'))).toBe('HEADER_B');
    ed.destroy();
    host.remove();
  });

  it('opening an empty footer of an earlier section and leaving it empty writes nothing', async () => {
    const { ed, host } = await editor();
    (ed as any).pages = [{ ...(ed as any).firstPage(), section: 1, inSection: 1, number: 2, top: 0 }];
    ed.editHeaderFooter('footer', 0);
    ed.closeHeaderFooter();
    expect(ed.model.headerFooters.some((h) => h.pending)).toBe(false);
    expect(ed.isModified()).toBe(false);
    const saved = await JSZip.loadAsync(await ed.save());
    const docXml = await saved.file('word/document.xml')!.async('string');
    const rels = await saved.file('word/_rels/document.xml.rels')!.async('string');
    const contentTypes = await saved.file('[Content_Types].xml')!.async('string');
    expect(docXml).not.toMatch(/<w:footerReference\b/);
    expect(rels).not.toMatch(/relationships\/footer/);
    expect(contentTypes).not.toMatch(/footer\+xml/);
    expect(Object.keys(saved.files).filter((name) => /^word\/footer\d+\.xml$/.test(name))).toEqual([]);
    ed.destroy();
    host.remove();
  });
});

describe('the QA document with portrait and landscape sections', () => {
  const qa = process.env.QA_SECTIONS_DOCX;
  it.skipIf(!qa)('reads both sections with their own headers', async () => {
    const { readFileSync } = await import('node:fs');
    const { doc, model } = await readDocx(readFileSync(qa!));
    const s = documentSections(doc, model);
    expect(s.map((x) => x.page.width > x.page.height)).toEqual([false, true]);
    expect(text(sectionHeaderFooter(model, s, 0, 'header', 'default'))).toContain('QA_HEADER_A');
    expect(text(sectionHeaderFooter(model, s, 1, 'header', 'default'))).toContain('QA_HEADER_B');
    void schema;
  });
});
