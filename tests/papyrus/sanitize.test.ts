// Values from the file (or a paste) never reach a style attribute unchecked, and pasted
// cell styles map to valid Word values.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { DOMParser as PMDOMParser, DOMSerializer } from 'prosemirror-model';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const EVIL = 'FF0000;background-image:url(https://evil.example/x)';

async function open(body: string) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  return readDocx(await zip.generateAsync({ type: 'uint8array' }));
}

/** Every style attribute the document renders to. */
function styles(doc: ReturnType<typeof schema.nodes.doc.create>): string {
  const box = document.createElement('div');
  box.append(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
  return Array.from(box.querySelectorAll('[style]')).map((e) => e.getAttribute('style')).join('\n');
}

describe('style attributes built from file content', () => {
  const tcPr =
    `<w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:tcBorders><w:top w:val="single;color:red" w:sz="4;x:y" w:space="0" w:color="${EVIL}"/>` +
    `<w:left w:val="dashed" w:sz="8" w:color="00FF00"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="${EVIL}"/><w:vAlign w:val="top;background:url(x)"/></w:tcPr>`;
  const body =
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>${tcPr}<w:p/></w:tc></w:tr></w:tbl>` +
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial&quot;;background:url(x);&quot;" w:eastAsia="新細明體"/><w:color w:val="${EVIL}"/><w:shd w:val="clear" w:color="auto" w:fill="${EVIL}"/></w:rPr><w:t>x</w:t></w:r></w:p>`;

  it('cell shading, vertical alignment, borders, run colour and font cannot inject CSS', async () => {
    const { doc } = await open(body);
    const css = styles(doc);
    // Outside quoted font names nothing but the intended declarations.
    expect(css.replace(/"[^"]*"/g, '""')).not.toMatch(/url\(|evil|color:red|x:y|background/);
    // What is valid still shows; a font name cannot close its quotes.
    // (jsdom normalizes style attributes.)
    expect(css).toContain('border-top: 1px solid rgb(0, 0, 0); border-left: 1px dashed rgb(0, 255, 0);');
    // The Chinese font comes with its similar fonts (persona-300 B-8); nothing else gets in.
    expect(css).toContain('font-family: "Arialbackground:url(x)", "新細明體", "PMingLiU"');
    expect(css).not.toMatch(/url\(x\)[^"]/);
  });

  it('marks from other sources are checked too', () => {
    const text = schema.text('x', [
      schema.marks.color.create({ color: 'red;background:url(x)' }),
      schema.marks.highlight.create({ color: '#ff0;background:url(x)' }),
      schema.marks.font.create({ family: 'A";background:url(x);"', eastAsia: null }),
    ]);
    const css = styles(schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, text)));
    expect(css.replace(/"[^"]*"/g, '""')).not.toMatch(/url\(|background:/);
    const ok = schema.text('y', [schema.marks.color.create({ color: '#1F4E79' }), schema.marks.highlight.create({ color: 'rgb(255, 255, 0)' })]);
    const okCss = styles(schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, ok)));
    expect(okCss).toContain('color: rgb(31, 78, 121);');
    expect(okCss).toContain('background-color: rgb(255, 255, 0);');
  });

  it('odd values are still written back unchanged', async () => {
    const { doc, model } = await open(body);
    const xml = await (await JSZip.loadAsync(await writeDocx(doc, model))).file('word/document.xml')!.async('string');
    expect(xml).toContain(tcPr);
  });
});

describe('pasted table cells', () => {
  it('vertical-align baseline and a transparent background map to nothing', async () => {
    const html = document.createElement('div');
    html.innerHTML =
      '<table><tr><td style="vertical-align: baseline; background-color: rgba(0, 0, 0, 0)">a</td>' +
      '<td style="vertical-align: middle; background-color: rgb(255, 0, 0)">b</td>' +
      '<td style="vertical-align: bottom; background-color: transparent">c</td>' +
      '<td valign="top" style="background-color: rgba(0, 128, 0, 0.5)">d</td></tr></table>';
    const doc = PMDOMParser.fromSchema(schema).parse(html);
    const cells: Record<string, unknown>[] = [];
    doc.descendants((n) => void (n.type.name === 'table_cell' && cells.push(n.attrs)));
    expect(cells.map((a) => [a.vAlign, a.background])).toEqual([
      [null, null],
      ['middle', '#FF0000'],
      ['bottom', null],
      ['top', '#008000'],
    ]);
    const { model } = await open('<w:p/>');
    const xml = await (await JSZip.loadAsync(await writeDocx(doc, model))).file('word/document.xml')!.async('string');
    expect(xml).not.toMatch(/w:vAlign w:val="baseline"|w:fill="000000"/);
    expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/><w:vAlign w:val="center"/>');
  });
});
