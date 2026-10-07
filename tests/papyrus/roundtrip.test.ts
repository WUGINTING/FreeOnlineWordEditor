import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { schema } from '../../src/papyrus/editor/schema';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { ensureList, ListCounter } from '../../src/papyrus/docx/numbering';
import { content, stableIds } from './helpers';

const PNG_1PX =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

async function blank() {
  return readDocx(await blankPackage().generateAsync({ type: 'uint8array' }));
}

const n = schema.nodes;
const m = schema.marks;

describe('docx round trip', () => {
  it('reads the blank template', async () => {
    const { doc, model } = await blank();
    expect(doc.childCount).toBe(1);
    expect(doc.firstChild!.type.name).toBe('paragraph');
    expect(model.page.width).toBe(11906);
    expect(model.paragraphStyles.map((s) => s.id)).toContain('Heading1');
    expect(model.css).toContain('.dx-ps-Heading1');
  });

  it('keeps text, formatting, lists, tables, images and page breaks', async () => {
    const { model } = await blank();
    const bullet = ensureList(model.numbering, 'bullet');
    const decimal = ensureList(model.numbering, 'decimal');
    expect(bullet).not.toBe(decimal);

    const doc = n.doc.create(null, [
      n.paragraph.create({ styleId: 'Heading1' }, schema.text('標題 Title')),
      n.paragraph.create({ align: 'center', indFirst: 480, spaceAfter: 200, line: 360, lineRule: 'auto' }, [
        schema.text('bold ', [m.bold.create()]),
        schema.text('italic', [m.italic.create(), m.color.create({ color: '#FF0000' })]),
        n.tab.create(),
        schema.text('link', [m.link.create({ href: 'https://example.com/' })]),
        n.hard_break.create(),
        schema.text('中文', [m.font.create({ family: 'Times New Roman', eastAsia: '標楷體' }), m.fontSize.create({ pt: 14 })]),
        schema.text('hl', [m.highlight.create({ color: '#ffff00' })]),
        schema.text('x2', [m.superscript.create()]),
      ]),
      n.paragraph.create({ numId: bullet, ilvl: 0 }, schema.text('item one')),
      n.paragraph.create({ numId: bullet, ilvl: 1 }, schema.text('item two')),
      n.page_break.create(),
      n.table.create(null, [
        n.table_row.create(null, [
          n.table_cell.create({ rowspan: 2, colwidth: [100] }, n.paragraph.create(null, schema.text('A'))),
          n.table_cell.create({ colspan: 2, colwidth: [80, 120], background: '#DDDDDD' }, n.paragraph.create(null, schema.text('B'))),
        ]),
        n.table_row.create(null, [
          n.table_cell.create({ colwidth: [80] }, n.paragraph.create(null, schema.text('C'))),
          n.table_cell.create({ colwidth: [120] }, n.paragraph.create(null, schema.text('D'))),
        ]),
      ]),
      n.paragraph.create(null, [n.image.create({ src: PNG_1PX, width: 40, height: 30 }), n.field.create({ instr: 'PAGE', text: '1' })]),
    ]);

    const bytes = await writeDocx(doc, model);
    const again = await readDocx(bytes);

    expect(content(again.doc.toJSON())).toEqual(content(doc.toJSON()));

    // Package sanity
    const zip = await JSZip.loadAsync(bytes);
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).toContain('Extension="png"');
    expect(ct).toContain('/word/numbering.xml');
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).toContain('https://example.com/');
    expect(Object.keys(zip.files).some((f) => f.startsWith('word/media/'))).toBe(true);

    // Saving twice does not accumulate media or relationships.
    const bytes2 = await writeDocx(again.doc, again.model);
    const zip2 = await JSZip.loadAsync(bytes2);
    expect(Object.values(zip2.files).filter((f) => !f.dir && f.name.startsWith("word/media/")).length).toBe(1);
    const rels2 = await zip2.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels2.match(/<Relationship /g)!.length).toBe(rels.match(/<Relationship /g)!.length);
  });

  it('parses a hand-written document with complex fields, vMerge and page breaks', async () => {
    const zip = blankPackage();
    zip.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
  <w:p><w:r><w:t>before</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>after</w:t></w:r></w:p>
  <w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
  <w:p><w:del><w:r><w:delText>gone</w:delText></w:r></w:del><w:ins><w:r><w:t>kept</w:t></w:r></w:ins></w:p>
  <w:tbl><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>M</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>X</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0"/></w:sectPr>
</w:body></w:document>`,
    );
    const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(model.page).toMatchObject({ width: 12240, height: 15840, marginTop: 720 });
    const types = [] as string[];
    doc.forEach((c) => types.push(c.type.name));
    expect(types).toEqual(['paragraph', 'page_break', 'paragraph', 'paragraph', 'paragraph', 'table']);
    expect(doc.child(0).textContent).toBe('before');
    expect(doc.child(2).textContent).toBe('after');
    // A complex field keeps its codes as invisible markers; its result stays editable text.
    const para = doc.child(3);
    expect(para.textContent).toBe('Page 3');
    let result: any = null;
    para.forEach((n) => {
      if (n.isText && n.text === '3') result = n;
    });
    expect(result.marks.find((m: any) => m.type.name === 'fieldResult').attrs.instr).toBe('PAGE');
    expect(doc.child(4).textContent).toBe('kept');
    const table = doc.child(5);
    expect(table.child(0).child(0).attrs.rowspan).toBe(2);
    expect(table.child(1).childCount).toBe(1);
  });
});

describe('list numbering', () => {
  it('counts nested levels like Word', async () => {
    const { model } = await blank();
    const id = ensureList(model.numbering, 'decimal');
    const c = new ListCounter(model.numbering);
    expect([c.next(id, 0), c.next(id, 1), c.next(id, 1), c.next(id, 0), c.next(id, 1)]).toEqual([
      '1.', 'a.', 'b.', '2.', 'a.',
    ]);
  });
});

describe('safety', () => {
  it('never renders script links as clickable', async () => {
    const { safeHref } = await import('../../src/papyrus/editor/schema');
    expect(safeHref('javascript:alert(1)')).toBe('#');
    expect(safeHref(' JavaScript:alert(1)')).toBe('#');
    expect(safeHref('https://example.com')).toBe('https://example.com');
    expect(safeHref('#bookmark')).toBe('#bookmark');
  });

  it('keeps formatting on tabs and fields', async () => {
    const { model } = await blank();
    const bold = [m.bold.create()];
    const doc = n.doc.create(null, [
      n.paragraph.create(null, [n.tab.create(null, null, bold), n.field.create({ instr: 'PAGE', text: '1' }, null, bold)]),
    ]);
    const again = await readDocx(await writeDocx(doc, model));
    expect(content(again.doc.toJSON())).toEqual(content(doc.toJSON()));
  });
});

describe('content the editor cannot edit', () => {
  it('keeps footnote marks, floating pictures and charts (with their files)', async () => {
    const zip = blankPackage();
    zip.file('word/media/float.png', PNG_1PX.split(',')[1], { base64: true });
    zip.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/float.png"/>
<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>`,
    );
    zip.file('word/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>');
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const D = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
    zip.file(
      'word/document.xml',
      `<w:document ${W} ${D}><w:body>
<w:p><w:r><w:t>note</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>
<w:p><w:r><w:drawing><wp:anchor><wp:extent cx="952500" cy="952500"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>
<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><a:graphic><a:graphicData><c:chart r:id="rId8"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
</w:body></w:document>`,
    );
    const first = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    const raws: any[] = [];
    first.doc.descendants((node) => {
      if (node.type.name === 'raw_inline') raws.push(node.attrs);
      return true;
    });
    expect(raws.map((r) => r.label)).toEqual(['註腳', '圖形', '圖表']);
    expect(raws[1].src).toMatch(/^data:image\/png/); // floating picture still shows

    const out = await JSZip.loadAsync(await writeDocx(first.doc, first.model));
    const xml = await out.file('word/document.xml')!.async('string');
    expect(xml).toContain('footnoteReference');
    expect(xml).toContain('wp:anchor');
    expect(xml).toContain('c:chart');
    expect(xml).toContain('<a:blip r:embed="rId7"/>');
    expect(xml).toContain('<c:chart r:id="rId8"/>');
    const rels = await out.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).toMatch(/<Relationship Id="rId7" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image" Target="media\/float\.png"\/>/);
    expect(rels).toMatch(/<Relationship Id="rId8" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/chart" Target="charts\/chart1\.xml"\/>/);
    expect(out.file('word/media/float.png')).not.toBeNull();
    expect(out.file('word/charts/chart1.xml')).not.toBeNull();

    const second = await readDocx(await writeDocx(first.doc, first.model));
    expect(stableIds(second.doc.toJSON())).toEqual(stableIds(first.doc.toJSON()));
  });
});

describe('headers and footers', () => {
  const para = (text: string) => n.paragraph.create(null, text ? schema.text(text) : undefined);
  const hdoc = (...blocks: any[]) => n.doc.create(null, blocks);

  it('creates new header/footer parts and wires them into the section', async () => {
    const { doc, model } = await blank();
    model.headerFooters.push(
      { kind: 'header', type: 'default', relId: null, part: null, doc: hdoc(para('公司名稱')), dirty: true },
      { kind: 'footer', type: 'default', relId: null, part: null, doc: hdoc(n.paragraph.create({ align: 'center' }, [schema.text('第 '), n.field.create({ instr: 'PAGE', text: '1' }), schema.text(' 頁')])), dirty: true },
      { kind: 'header', type: 'first', relId: null, part: null, doc: hdoc(para('')), dirty: true }, // empty: not created
    );
    model.titlePage = true;
    const bytes = await writeDocx(doc, model);
    const zip = await JSZip.loadAsync(bytes);
    const docXml = await zip.file('word/document.xml')!.async('string');
    expect(docXml).toMatch(/headerReference[^>]*w:type="default"/);
    expect(docXml).toMatch(/footerReference[^>]*w:type="default"/);
    expect(docXml).not.toMatch(/headerReference[^>]*w:type="first"/);
    expect(docXml).toContain('titlePg');
    expect(docXml.indexOf('headerReference')).toBeLessThan(docXml.indexOf('pgSz'));
    const ct = await zip.file('[Content_Types].xml')!.async('string');
    expect(ct).toContain('/word/header1.xml');
    expect(ct).toContain('/word/footer1.xml');

    const again = await readDocx(bytes);
    expect(again.model.titlePage).toBe(true);
    const header = again.model.headerFooters.find((h) => h.kind === 'header' && h.type === 'default')!;
    const footer = again.model.headerFooters.find((h) => h.kind === 'footer')!;
    expect(header.doc.textContent).toBe('公司名稱');
    expect(content(footer.doc.toJSON())).toEqual(content(model.headerFooters[1].doc.toJSON()));
    expect(header.part).toBe('word/header1.xml');
  });

  it('edits an existing header in place and keeps its images apart from body images', async () => {
    const { doc, model } = await blank();
    model.headerFooters.push({ kind: 'header', type: 'default', relId: null, part: null, doc: hdoc(n.paragraph.create(null, [n.image.create({ src: PNG_1PX, width: 10, height: 10 }), schema.text('v1')])), dirty: true });
    const bodyImageSrc = PNG_1PX.replace('AAAAAElFTkSuQmCC', 'AAAAAElFTkSuQmCD');
    const body = n.doc.create(null, [n.paragraph.create(null, [n.image.create({ src: bodyImageSrc, width: 5, height: 5 })])]);
    const first = await readDocx(await writeDocx(body, model));

    // Edit the header text only; save twice.
    const h = first.model.headerFooters[0];
    const img = h.doc.firstChild!.firstChild!;
    h.doc = hdoc(n.paragraph.create(null, [img, schema.text('v2')]));
    h.dirty = true;
    const second = await readDocx(await writeDocx(first.doc, first.model));
    const zip = await JSZip.loadAsync(await writeDocx(second.doc, second.model));

    const header = second.model.headerFooters[0];
    expect(header.part).toBe('word/header1.xml'); // same part, not a new one
    expect(header.doc.textContent).toBe('v2');
    expect(header.doc.firstChild!.firstChild!.attrs.src).toBe(PNG_1PX);
    expect(second.doc.firstChild!.firstChild!.type.name).toBe('image');
    expect(second.doc.firstChild!.firstChild!.attrs.src).toBe(bodyImageSrc);
    const docXml = await zip.file('word/document.xml')!.async('string');
    const bodyEmbed = /<a:blip r:embed="([^"]+)"/.exec(docXml)?.[1];
    expect(bodyEmbed).toBeTruthy();
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    const bodyRelationship = new RegExp(`<Relationship Id="${bodyEmbed}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/([^"]+)"/>`).exec(rels);
    expect(bodyRelationship, `body embed ${bodyEmbed} should target its media through an image relationship: ${rels}`).toBeTruthy();
    const bodyMedia = await zip.file(`word/media/${bodyRelationship![1]}`)!.async('base64');
    expect(bodyMedia).toBe(bodyImageSrc.split(',')[1]);
    const media = Object.values(zip.files).filter((f) => !f.dir && f.name.startsWith('word/media/'));
    expect(media.length).toBe(2);
    expect(Object.keys(zip.files).filter((f) => /word\/header\d+\.xml$/.test(f))).toEqual(['word/header1.xml']);
  });
});
