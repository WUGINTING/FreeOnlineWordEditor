// Tables whose w:tblGrid is a placeholder (generated files) are shown with the cells' own
// widths, as Word lays them out, and are still written back unchanged.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const cell = (w: string, text: string) => `<w:tc><w:tcPr>${w}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

async function table(tbl: string) {
  const zip = blankPackage();
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${tbl}<w:p/>` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
  );
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const { doc, model } = await readDocx(bytes);
  const widths = (row = 0) => {
    const out: number[] = [];
    doc.child(0).child(row).forEach((c) => out.push(...(c.attrs.colwidth ?? [])));
    return out;
  };
  const before = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
  const after = await (await JSZip.loadAsync(await writeDocx(doc, model))).file('word/document.xml')!.async('string');
  return { widths, before, after };
}

const tbl = (tblPr: string, grid: number[], rows: string[]) =>
  `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.map((r) => `<w:tr>${r}</w:tr>`).join('')}</w:tbl>`;

describe('table column widths', () => {
  it('a placeholder grid (100 twips per column) is shown with the cells’ widths and saved unchanged', async () => {
    const t = await table(
      tbl('<w:tblW w:type="dxa" w:w="9386"/>', [100, 100, 100, 100], [
        [2346, 2346, 2347, 2347].map((w, i) => cell(`<w:tcW w:type="dxa" w:w="${w}"/>`, `c${i}`)).join(''),
      ]),
    );
    // 2346 twips ≈ 156 px (was 7 px).
    expect(t.widths()).toEqual([156, 156, 156, 156]);
    const tblXml = (xml: string) => xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'));
    expect(tblXml(t.after)).toBe(tblXml(t.before));
  });

  it('QA20K-00720 percentage cell widths are calculated from the containing table width', async () => {
    const t = await table(
      tbl('<w:tblW w:type="dxa" w:w="9000"/>', [100, 100], [
        cell('<w:tcW w:type="pct" w:w="1250"/>', 'a') + cell('<w:tcW w:type="pct" w:w="3750"/>', 'b'),
      ]),
    );
    // 25 % and 75 % of 9000 twips.
    expect(t.widths()).toEqual([150, 450]);
  });

  it('without usable cell widths a stale grid is stretched to the table width', async () => {
    const t = await table(
      tbl('<w:tblW w:type="dxa" w:w="9000"/>', [100, 200], [cell('', 'a') + cell('', 'b')]),
    );
    expect(t.widths()).toEqual([200, 400]);
  });

  it('a normal grid and fixed-layout tables keep their grid', async () => {
    const normal = await table(
      tbl('<w:tblW w:type="dxa" w:w="6000"/>', [3000, 3000], [
        cell('<w:tcW w:type="dxa" w:w="2900"/>', 'a') + cell('<w:tcW w:type="dxa" w:w="3100"/>', 'b'),
      ]),
    );
    expect(normal.widths()).toEqual([200, 200]);
    const fixed = await table(
      tbl('<w:tblW w:type="dxa" w:w="9386"/><w:tblLayout w:type="fixed"/>', [100, 100], [
        cell('<w:tcW w:type="dxa" w:w="4693"/>', 'a') + cell('<w:tcW w:type="dxa" w:w="4693"/>', 'b'),
      ]),
    );
    expect(fixed.widths()).toEqual([7, 7]);
  });
});
