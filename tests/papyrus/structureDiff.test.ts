// "What changed besides the words" between two versions (docx/structureDiff.ts).
import { describe, expect, it } from 'vitest';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { compareStructure } from '../../src/papyrus/docx/structureDiff';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const PG = (w = 11906, h = 16838, left = 1800) =>
  `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="${left}" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>`;
const cell = (t: string, shade = '') => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${shade}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
const table = (rows: string[][], shade = '') =>
  `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
  rows.map((r) => `<w:tr>${r.map((t) => cell(t, shade)).join('')}</w:tr>`).join('') + '</w:tbl>';

async function docx(body: string): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const BASE = `<w:p><w:r><w:t>金額為新台幣十萬元整。</w:t></w:r></w:p>` + table([['項目', '金額'], ['服務費', '100']]) + `<w:p/>`;

describe('structure changes between two versions', () => {
  it('reports nothing for the same document', async () => {
    const a = await docx(BASE + PG());
    expect(await compareStructure(a, a)).toEqual([]);
  });

  it('reports page setup: paper, orientation and margins', async () => {
    const changes = await compareStructure(await docx(BASE + PG()), await docx(BASE + PG(16838, 11906, 1440)));
    const text = changes.map((c) => c.text).join('\n');
    expect(text).toContain('A4直向 → A4橫向');
    expect(text).toMatch(/邊界：左 3\.17 公分 → 2\.54 公分/);
  });

  it('reports a table row added, a cell changed and a new cell shading, even when the words elsewhere are the same', async () => {
    const shaded = '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>';
    const b = `<w:p><w:r><w:t>金額為新台幣十萬元整。</w:t></w:r></w:p>` + table([['項目', '金額'], ['服務費', '100'], ['稅', '5']]) + `<w:p/>`;
    const rows = await compareStructure(await docx(BASE + PG()), await docx(b + PG()));
    expect(rows.find((c) => c.area === 'table')?.text).toContain('列數 2 → 3');
    const look = await compareStructure(await docx(BASE + PG()), await docx(BASE.replace(/<w:tcW w:w="3000" w:type="dxa"\/>/g, `<w:tcW w:w="3000" w:type="dxa"/>${shaded}`) + PG()));
    expect(look.map((c) => c.text).join()).toContain('外觀改變');
  });

  it('reports formatting changed on unchanged words (bold, alignment)', async () => {
    const b = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>金額為新台幣十萬元整。</w:t></w:r></w:p>` + table([['項目', '金額'], ['服務費', '100']]) + `<w:p/>`;
    const changes = await compareStructure(await docx(BASE + PG()), await docx(b + PG()));
    const format = changes.filter((c) => c.area === 'format');
    expect(format.map((c) => c.text)).toEqual(expect.arrayContaining(['1 段的對齊改變', '1 段的文字格式（粗體、字型、字級、顏色等）改變']));
    expect(format[0].examples).toEqual(['金額為新台幣十萬元整。']);
  });

  it('reports a header whose words changed', async () => {
    const withHeader = async (words: string) => {
      const zip = blankPackage();
      zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:r><w:t>${words}</w:t></w:r></w:p></w:hdr>`);
      const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
      zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
      zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${BASE}${PG().replace('<w:sectPr>', '<w:sectPr><w:headerReference w:type="default" r:id="rH"/>')}</w:body></w:document>`);
      return zip.generateAsync({ type: 'uint8array' });
    };
    const changes = await compareStructure(await withHeader('機密 V1'), await withHeader('機密 V2'));
    expect(changes).toEqual([{ area: 'headerFooter', text: '頁首內容：「機密 V1」→「機密 V2」' }]);
  });
});

// With DOCX_SAMPLES: a document saved by the editor without edits has no structure changes.
const samples = process.env.DOCX_SAMPLES;
const files = (dir: string, out: string[] = []): string[] => {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) files(p, out);
    else if (/\.docx$/i.test(name) && !name.startsWith('~$')) out.push(p);
  }
  return out;
};
describe.skipIf(!samples)('real documents: no false changes after a save without edits', () => {
  it('every sample', async () => {
    const noisy: string[] = [];
    for (const file of files(samples!)) {
      const data = readFileSync(file);
      const { doc, model } = await readDocx(data);
      const saved = await writeDocx(doc, model);
      const changes = await compareStructure(data, saved);
      if (changes.length) noisy.push(`${file}: ${changes.map((c) => c.text).join(' | ')}`);
    }
    expect(noisy).toEqual([]);
  }, 300000);
});
