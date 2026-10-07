import { describe, expect, it } from 'vitest';
import { blankPackage } from '../src/papyrus/docx/template';
import { readDocx } from '../src/papyrus/docx/reader';
import { writeDocx } from '../src/papyrus/docx/writer';

/** 200 synthetic table geometries, isolated from customer documents. */
const geometries = Array.from({ length: 200 }, (_, i) => ({
  id: `L${String(i + 1).padStart(3, '0')}`,
  rows: Math.floor(i / 20) + 1,
  cols: (i % 20) + 1,
}));

describe('QA600 layout stability: table geometry matrix', () => {
  it.each(geometries)('$id preserves $rows rows × $cols columns through DOCX round trip', async ({ id, rows, cols }) => {
    const zip = blankPackage();
    const expected: string[] = [];
    const tableRows = Array.from({ length: rows }, (_, r) => {
      const cells = Array.from({ length: cols }, (_, c) => {
        const text = `${id}-r${r + 1}-c${c + 1}-中文`;
        expected.push(text);
        return `<w:tc><w:tcPr/><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
      }).join('');
      return `<w:tr><w:trPr/>${cells}</w:tr>`;
    }).join('');
    // Keep the generated table grid aligned with each test's declared column count.
    const grid = Array.from({ length: cols }, () => '<w:gridCol w:w="1200"/>').join('');
    const table = `<w:tbl><w:tblPr/><w:tblGrid>${grid}</w:tblGrid>${tableRows}</w:tbl>`;
    const original = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', original.replace('</w:body>', `${table}</w:body>`));
    const parsed = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    const tables: any[] = [];
    parsed.doc.descendants((node) => {
      if (node.type.name === 'table') tables.push(node);
      return true;
    });
    expect(tables, `${id}: parsed tables`).toHaveLength(1);
    expect(tables[0].childCount, `${id}: row count`).toBe(rows);
    expect(tables[0].textContent, `${id}: cell text`).toBe(expected.join(''));

    const saved = await readDocx(await writeDocx(parsed.doc, parsed.model));
    const savedTables: any[] = [];
    saved.doc.descendants((node) => {
      if (node.type.name === 'table') savedTables.push(node);
      return true;
    });
    expect(savedTables, `${id}: saved tables`).toHaveLength(1);
    expect(savedTables[0].childCount, `${id}: saved row count`).toBe(rows);
    expect(savedTables[0].textContent, `${id}: saved cell text`).toBe(expected.join(''));
    for (let r = 0; r < rows; r++) {
      expect(savedTables[0].child(r).childCount, `${id}: row ${r + 1} cell count`).toBe(cols);
    }
  });
});
