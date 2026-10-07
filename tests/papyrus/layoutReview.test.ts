// Regressions from the review of sections, layout and printing.
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { TextSelection } from 'prosemirror-state';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { parseStyles } from '../../src/papyrus/docx/styles';
import { parseXml } from '../../src/papyrus/docx/xml';
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
const PG = (w: number, h: number, orient = '') =>
  `<w:pgSz w:w="${w}" w:h="${h}"${orient}/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`;

async function editorFor(body: string, extra?: (zip: JSZip) => Promise<void>) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  if (extra) await extra(zip);
  const host = document.createElement('div');
  document.body.append(host);
  const ed = new DocxEditor(host);
  await ed.open(await zip.generateAsync({ type: 'uint8array' }));
  const at = (pos: number) => ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, pos)));
  const type = (text: string) => ed.run((s, d) => (d?.(s.tr.insertText(text)), true));
  const save = async () => JSZip.loadAsync(await writeDocx(ed.view!.state.doc, ed.model));
  const done = () => {
    ed.destroy();
    host.remove();
  };
  return { ed, at, type, save, done };
}

const TWO = `<w:p><w:r><w:t>A1</w:t></w:r></w:p><w:p><w:pPr><w:sectPr>${PG(11906, 16838)}</w:sectPr></w:pPr><w:r><w:t>A2</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>B1</w:t></w:r></w:p><w:sectPr>${PG(16838, 11906, ' w:orient="landscape"')}</w:sectPr>`;

describe('last section page setup is part of the undo history', () => {
  it('undo reverts a page setup change of the last section, not the typing before it', async () => {
    const { ed, at, type, done } = await editorFor(TWO);
    at(ed.view!.state.doc.content.size - 2); // in "B1", the last section
    type('X');
    const page = ed.cursorSection().page;
    ed.setPageSetup({ ...page, marginLeft: 3000 }, 'section');
    expect(ed.model.page.marginLeft).toBe(3000);
    ed.undo();
    expect(ed.model.page.marginLeft).toBe(1440);
    expect(ed.view!.state.doc.textContent).toContain('X');
    expect(ed.isModified()).toBe(true);
    ed.undo();
    expect(ed.view!.state.doc.textContent).not.toContain('X');
    expect(ed.isModified()).toBe(false);
    done();
  });

  it('"whole document" margin setup is one undo step for every section', async () => {
    const { ed, done } = await editorFor(TWO);
    const page = ed.cursorSection().page;
    ed.setPageSetup({ ...page, marginLeft: 2500 }, 'all');
    expect(ed.sections().map((s) => s.page.marginLeft)).toEqual([2500, 2500]);
    ed.undo();
    expect(ed.sections().map((s) => s.page.marginLeft)).toEqual([1440, 1440]);
    done();
  });

  it('"different first page" of the last section can be undone', async () => {
    const { ed, at, done } = await editorFor(TWO);
    at(ed.view!.state.doc.content.size - 2);
    ed.setTitlePage(true);
    expect(ed.model.titlePage).toBe(true);
    ed.undo();
    expect(ed.model.titlePage).toBe(false);
    done();
  });
});

describe('section breaks and headers', () => {
  it('a header made in this session is kept by both halves when a section break is inserted', async () => {
    const { ed, at, type, save, done } = await editorFor(`<w:p><w:r><w:t>one two</w:t></w:r></w:p><w:sectPr>${PG(11906, 16838)}</w:sectPr>`);
    ed.editHeaderFooter('header', 0);
    type('MY HEADER');
    ed.closeHeaderFooter();
    at(4);
    ed.insertSectionBreak();
    const xml = await (await save()).file('word/document.xml')!.async('string');
    const sects = xml.match(/<w:sectPr\b.*?<\/w:sectPr>/g)!;
    expect(sects).toHaveLength(2);
    for (const s of sects) expect(s).toMatch(/<w:headerReference w:type="default" r:id="[^"]+"\/>/);
    done();
  });

  it('the section after an inserted break starts on a new page and continues the numbering', async () => {
    const body = `<w:p><w:r><w:t>one two</w:t></w:r></w:p><w:sectPr><w:type w:val="continuous"/>${PG(11906, 16838)}<w:pgNumType w:start="5"/></w:sectPr>`;
    const { ed, at, done } = await editorFor(body);
    at(4);
    ed.insertSectionBreak();
    const [first, second] = ed.sections();
    expect(first.start).toBe('continuous');
    expect(first.pageNumberStart).toBe(5);
    expect(second.start).toBe('nextPage');
    expect(second.pageNumberStart).toBeNull();
    done();
  });

  it('a header part two references share is edited in place and stays in step', async () => {
    const hdr = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:r><w:t>SHARED</w:t></w:r></w:p></w:hdr>`;
    const body = `<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="h1"/><w:headerReference w:type="first" r:id="h2"/>${PG(11906, 16838)}<w:titlePg/></w:sectPr>`;
    const { ed, type, save, done } = await editorFor(body, async (zip) => {
      zip.file('word/header1.xml', hdr);
      const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
      const rel = (id: string) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`;
      zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', rel('h1') + rel('h2') + '</Relationships>'));
    });
    ed.editHeaderFooter('header', 0); // page 1: the first-page header
    ed.run((s, d) => (d?.(s.tr.insertText('EDITED ', 1)), true));
    ed.closeHeaderFooter();
    const both = ed.model.headerFooters.map((h) => h.doc.textContent);
    expect(both).toEqual(['EDITED SHARED', 'EDITED SHARED']);
    const out = await save();
    expect(await out.file('word/header1.xml')!.async('string')).toContain('EDITED ');
    expect(Object.keys(out.files).filter((n) => /header\d+\.xml$/.test(n))).toEqual(['word/header1.xml']);
    done();
    void type;
  });

  it('a new header whose section reference was undone is not written', async () => {
    const { ed, at, type, save, done } = await editorFor(`<w:p><w:r><w:t>one two</w:t></w:r></w:p><w:sectPr>${PG(11906, 16838)}</w:sectPr>`);
    at(4);
    ed.insertSectionBreak();
    // A first-page header for section 1 ... then the break is undone.
    (ed as any).pages = [{ ...(ed as any).firstPage(), section: 0, inSection: 0, number: 1, top: 0 }];
    ed.editHeaderFooter('footer', 0);
    type('GONE');
    ed.closeHeaderFooter();
    ed.undo();
    expect(ed.sections()).toHaveLength(1);
    const out = await save();
    expect(Object.keys(out.files).some((n) => /footer\d+\.xml$/.test(n))).toBe(false);
    done();
  });
});

describe('document styles', () => {
  it('font names and colours from styles.xml cannot add CSS of their own', () => {
    const xml = `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr>` +
      '<w:rFonts w:ascii="Evil&quot;;}body{background:url(https://evil.example/x)}.x{" w:eastAsia="新細明體"/>' +
      '<w:color w:val="000;background:url(https://evil.example/y)"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:jc w:val="center;background:red"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>';
    const { css } = parseStyles(parseXml(xml));
    expect(css).not.toContain('url(');
    expect(css).not.toContain('body{');
    expect(css).toContain('新細明體');
  });

  it('paragraph alignment from the file only becomes a known CSS value', () => {
    const p = schema.nodes.paragraph.create({ align: 'center;background:url(x)' });
    const dom = (schema.nodes.paragraph.spec.toDOM!(p) as any)[1];
    expect(dom.style ?? '').not.toContain('url(');
  });

  it('text in a table gets the table style’s paragraph spacing (Word: under paragraph styles)', () => {
    const xml = `<w:styles ${W}><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>' +
      '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Big"><w:name w:val="Big"/><w:pPr><w:spacing w:after="400"/></w:pPr></w:style></w:styles>';
    const rules = parseStyles(parseXml(xml)).css.split('\n');
    const table = rules.findIndex((r) => r.startsWith('.dx-doc :where(.dx-ts-TableGrid) .dx-p{'));
    const named = rules.findIndex((r) => r.startsWith('.dx-doc .dx-ps-Big{'));
    expect(table).toBeGreaterThan(-1);
    expect(rules[table]).toContain('margin-bottom:0px');
    expect(rules[table]).toContain('line-height:calc(var(--dx-lh, 1.15) * 1)');
    // Paragraph styles come later, so they win over the table style.
    expect(named).toBeGreaterThan(table);
  });

  it('print keeps the editor’s style scope and prints headers in full colour', async () => {
    const { ed, done } = await editorFor(`<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr>${PG(11906, 16838)}</w:sectPr>`);
    (ed as any).pages = [{ ...(ed as any).firstPage(), section: 0, inSection: 0, number: 1, top: 0 }];
    let frame: HTMLIFrameElement | null = null;
    const obs = new MutationObserver((m) => m.forEach((r) => r.addedNodes.forEach((n) => (n as HTMLElement).tagName === 'IFRAME' && (frame = n as HTMLIFrameElement))));
    obs.observe(document.body, { childList: true });
    ed.print();
    await new Promise((r) => setTimeout(r, 0));
    obs.disconnect();
    const doc = (frame as HTMLIFrameElement | null)?.contentDocument;
    const scope = [...(ed.view!.dom.closest('.dx-root')!.classList)].find((c) => /^dx-e\d+$/.test(c));
    expect(doc?.querySelector('.dx-print-root')?.classList.contains(scope!)).toBe(true);
    expect([...doc!.querySelectorAll('style')].map((s) => s.textContent).join('')).toContain('.dx-footer{opacity:1}');
    (frame as HTMLIFrameElement | null)?.remove();
    done();
  });

  it('a row breaking across pages: its cover and repeated header rows are drawn and printed', async () => {
    const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
    const body = `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
      `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('HEAD A')}${cell('HEAD B')}</w:tr><w:tr>${cell('long')}${cell('text')}</w:tr></w:tbl><w:p/><w:sectPr>${PG(11906, 16838)}</w:sectPr>`;
    const { ed, done } = await editorFor(body);
    const first = (ed as any).firstPage();
    const header = { table: 0, rows: [1] }; // the table starts the document; its first row is at 1
    const pages = [
      { ...first, cuts: [{ top: 900, bottom: first.height + 24 + first.textTop, left: 90, right: 700, border: '1px solid rgb(0, 0, 0)', header }] },
      { ...first, inSection: 1, number: 2, top: first.height + 24 },
    ];
    (ed as any).renderPages(pages);
    const root = ed.view!.dom.closest('.dx-root')!;
    const band = root.querySelector<HTMLElement>('.dx-cuts > .dx-cut')!;
    expect(band.style.top).toBe('900px');
    expect(band.style.borderTop).toContain('solid');
    const copy = root.querySelector<HTMLElement>('.dx-cuts > .dx-cut-header')!;
    expect(copy.textContent).toContain('HEAD A');
    expect(copy.textContent).not.toContain('long');
    expect(copy.classList.contains('dx-doc')).toBe(true);

    let frame: HTMLIFrameElement | null = null;
    const obs = new MutationObserver((m) => m.forEach((r) => r.addedNodes.forEach((n) => (n as HTMLElement).tagName === 'IFRAME' && (frame = n as HTMLIFrameElement))));
    obs.observe(document.body, { childList: true });
    ed.print();
    await new Promise((r) => setTimeout(r, 0));
    obs.disconnect();
    const sheets = [...(frame as HTMLIFrameElement | null)!.contentDocument!.querySelectorAll('.dx-print-page')];
    expect(sheets).toHaveLength(2);
    expect(sheets[0].querySelector<HTMLElement>('.dx-cut')?.style.top).toBe('900px');
    expect(sheets[1].querySelector('.dx-cut-header')?.textContent).toContain('HEAD B');
    (frame as HTMLIFrameElement | null)?.remove();
    done();
  });
});

describe('fonts per script (Word: w:ascii for Latin text, w:eastAsia for CJK text)', () => {
  it('a run naming only a CJK font keeps the inherited Latin font for its Latin text', () => {
    const font = schema.marks.font.create({ family: null, eastAsia: '標楷體' });
    const dom = schema.marks.font.spec.toDOM!(font, true) as [string, Record<string, string>];
    const style = dom[1].style;
    expect(style).toContain('--dx-font-e:"標楷體"');
    expect(style).not.toContain('--dx-font-l:');
    expect(style).toContain('font-family:var(--dx-font-l),"標楷體"');
    // Pasting our own output back gives the same mark.
    const p = document.createElement('p');
    p.innerHTML = `<span style='${style}'>GPT 模型</span>`;
    const span = p.firstElementChild as HTMLElement;
    const rule = schema.marks.font.spec.parseDOM![0];
    expect((rule.getAttrs as (v: string) => unknown)(span.style.fontFamily)).toEqual({ family: null, eastAsia: '標楷體' });
  });

  it('styles set the two fonts separately, and paragraphs rebuild the font list from them', () => {
    const xml = `<w:styles ${W}>
      <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="新細明體"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Kai"><w:name w:val="Kai"/><w:rPr><w:rFonts w:eastAsia="標楷體"/></w:rPr></w:style>
    </w:styles>`;
    const { css } = parseStyles(parseXml(xml));
    expect(css).toContain('.dx-doc .dx-ps-Kai{--dx-font-e:"標楷體"');
    expect(css).not.toMatch(/\.dx-ps-Kai\{[^}]*--dx-font-l/);
    expect(css).toMatch(/\.dx-doc\{[^}]*--dx-font-l:"Times New Roman"/);
    expect(css).toMatch(/\.dx-doc \.dx-p\{[^}]*font-family:var\(--dx-font-l\), var\(--dx-font-e\), sans-serif/);
  });
});

describe('cell margins (w:tblCellMar, w:tcMar)', () => {
  it('a table style sets them for its tables, over Word\'s defaults', () => {
    const xml = `<w:styles ${W}>
      <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Table Normal"/>
        <w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
      <w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/><w:basedOn w:val="TableNormal"/>
        <w:tblPr><w:tblCellMar><w:left w:w="150" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
    </w:styles>`;
    const { css } = parseStyles(parseXml(xml));
    const defaults = css.indexOf('.dx-doc :where([class*="dx-ts-"]){--dx-cm-t:0px;--dx-cm-r:7.2px');
    expect(defaults).toBeGreaterThan(-1);
    const normal = css.indexOf('.dx-doc :where(.dx-ts-TableNormal),.dx-doc :where(.dx-ts--none){--dx-cm-t:0px;--dx-cm-l:0px;--dx-cm-b:0px;--dx-cm-r:0px}');
    expect(normal).toBeGreaterThan(defaults);
    // basedOn: the left margin is the style's own, the rest comes from TableNormal.
    expect(css).toContain('.dx-doc :where(.dx-ts-Grid){--dx-cm-t:0px;--dx-cm-l:10px;--dx-cm-b:0px;--dx-cm-r:0px}');
  });

  it('a table\'s own margins and a cell\'s own margins are shown', async () => {
    const { ed, done } = await editorFor(`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>
      <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:tcMar><w:left w:w="300" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`);
    const table = ed.view!.dom.querySelector('table')!;
    const tagged = table.closest('[class*="dx-ts-"]') as HTMLElement;
    expect(tagged.style.getPropertyValue('--dx-cm-l')).toBe('0px');
    expect(tagged.style.getPropertyValue('--dx-cm-r')).toBe('0px');
    const cells = table.querySelectorAll('td');
    expect((cells[0] as HTMLElement).style.paddingLeft).toBe('');
    expect((cells[1] as HTMLElement).style.paddingLeft).toBe('20px');
    done();
  });
});


describe('page and section breaks keep one paragraph mark', () => {
  const MARK = '<w:rPr><w:ins w:id="7" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr>';
  const paras = (xml: string) => xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? [];

  it('a page break goes inside the paragraph, like Word: one w:p, its mark written once', async () => {
    const { commands } = { commands: await import('../../src/papyrus/editor/commands') };
    const { ed, at, save, done } = await editorFor(`<w:p><w:pPr>${MARK}</w:pPr><w:r><w:t>one two</w:t></w:r></w:p><w:sectPr>${PG(11906, 16838)}</w:sectPr>`);
    at(4);
    ed.run(commands.insertPageBreak);
    const xml = await (await save()).file('word/document.xml')!.async('string');
    const ps = paras(xml);
    expect(ps).toHaveLength(1);
    expect(ps[0]).toContain('w:type="page"');
    expect(ps[0].match(/<w:ins w:id=/g)).toHaveLength(1);
    expect(ps[0]).toMatch(/one[\s\S]*w:type="page"[\s\S]*two/);
    done();
  });

  it('a section break leaves the original mark (and its revision) on the second half', async () => {
    const { ed, at, save, done } = await editorFor(`<w:p><w:pPr>${MARK}</w:pPr><w:r><w:t>one two</w:t></w:r></w:p><w:sectPr>${PG(11906, 16838)}</w:sectPr>`);
    at(4);
    ed.insertSectionBreak();
    const xml = await (await save()).file('word/document.xml')!.async('string');
    const ps = paras(xml);
    expect(ps).toHaveLength(2);
    expect(ps[0]).toContain('<w:sectPr');
    expect(ps[0]).not.toContain('<w:ins ');
    expect(ps[1]).toContain('<w:ins w:id=');
    done();
  });
});

describe('Tab in tables (R3-02)', () => {
  const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
  const TABLE = `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
    `<w:tr>${cell('A')}${cell('B')}</w:tr><w:tr>${cell('C')}${cell('200 元')}</w:tr></w:tbl><w:p/><w:sectPr>${PG(11906, 16838)}</w:sectPr>`;
  const press = (ed: DocxEditor, key: string) =>
    ed.view!.someProp('handleKeyDown', (f) => f(ed.view!, new KeyboardEvent('keydown', { key })));

  it('Tab in the last cell adds a row and moves into it, like Word; undo removes it', async () => {
    const { ed, done } = await editorFor(TABLE);
    let end = 0;
    ed.view!.state.doc.descendants((n, pos) => {
      if (n.isText && n.text === '200 元') end = pos + n.nodeSize;
    });
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, end)));
    expect(press(ed, 'Tab')).toBe(true);
    const table = ed.view!.state.doc.firstChild!;
    expect(table.childCount).toBe(3);
    let tabs = 0;
    ed.view!.state.doc.descendants((n) => void (n.type.name === 'tab' && tabs++));
    expect(tabs).toBe(0);
    // The cursor is in the new row's first cell.
    const $sel = ed.view!.state.selection.$from;
    expect($sel.node($sel.depth - 2)).toBe(table.child(2));
    expect($sel.index($sel.depth - 2)).toBe(0);
    ed.undo();
    expect(ed.view!.state.doc.firstChild!.childCount).toBe(2);
    expect(ed.isModified()).toBe(false);
    done();
  });

  it('Tab in any other cell still moves to the next cell', async () => {
    const { ed, done } = await editorFor(TABLE);
    ed.view!.dispatch(ed.view!.state.tr.setSelection(TextSelection.create(ed.view!.state.doc, 4)));
    press(ed, 'Tab');
    expect(ed.view!.state.doc.firstChild!.childCount).toBe(2);
    expect(ed.view!.state.selection.$from.parent.textContent).toBe('B');
    done();
  });
});

describe('printout as one HTML page (for a server-made PDF)', () => {
  it('holds every sheet, its paper size and the styles as text, with nothing to load from the site', async () => {
    const { ed, done } = await editorFor(`<w:p><w:r><w:t>PDF 內容</w:t></w:r></w:p><w:sectPr>${PG(11906, 16838)}</w:sectPr>`);
    (ed as any).pages = [{ ...(ed as any).firstPage(), section: 0, inSection: 0, number: 1, top: 0 }];
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = '/assets/app.css';
    document.head.append(link);
    const html = ed.printHtml('合約')!;
    link.remove();
    expect(html.startsWith('<!doctype html>')).toBe(true);
    expect(html).toContain('<title>合約</title>');
    expect(html).toContain('class="dx-print-page dxp');
    expect(html).toMatch(/@page dxp\d+x\d+\{size:/);
    expect(html).not.toContain('<link');
    // (Which blocks land on which sheet needs a real layout; jsdom measures everything as 0.)
    // The editor's own styles (editor.css) come along as text.
    expect(html).toContain('.dx-print-root');
    done();
  });
});
