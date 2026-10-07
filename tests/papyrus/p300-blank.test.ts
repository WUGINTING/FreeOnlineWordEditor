// persona-300 12: a new blank document is a 公文 page, as the built-in templates: A4, 標楷體 for
// Chinese and Times New Roman for Latin text, 16 pt, single spacing, margins 2.5 / 2.5 / 2.5 /
// 2.0 cm, header and footer 1.2 cm.
import { describe, expect, it } from 'vitest';
import { readDocx } from '../../src/papyrus/docx/reader';
import { GONGWEN_DEFAULTS, blankPackage } from '../../src/papyrus/docx/template';
import { cmToTwips } from './p300Units';

describe('空白文件', () => {
  it('writes the 公文 defaults into styles.xml and the section', async () => {
    const zip = blankPackage();
    const styles = await zip.file('word/styles.xml')!.async('string');
    expect(styles).toContain('<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="標楷體" w:cs="Times New Roman"/>');
    expect(styles).toContain('<w:sz w:val="32"/><w:szCs w:val="32"/>');
    expect(styles).toContain('<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>');
    const doc = await zip.file('word/document.xml')!.async('string');
    expect(doc).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(doc).toContain('<w:pgMar w:top="1417" w:right="1134" w:bottom="1417" w:left="1417" w:header="680" w:footer="680" w:gutter="0"/>');
    expect([GONGWEN_DEFAULTS.marginTop, GONGWEN_DEFAULTS.marginRight, GONGWEN_DEFAULTS.header]).toEqual([cmToTwips(2.5), cmToTwips(2), cmToTwips(1.2)]);
  });

  it('opens with those settings', async () => {
    const { model } = await readDocx(await blankPackage().generateAsync({ type: 'uint8array' }));
    expect(model.page).toMatchObject({ width: 11906, height: 16838, marginTop: 1417, marginBottom: 1417, marginLeft: 1417, marginRight: 1134, header: 680, footer: 680 });
    expect(model.css).toContain('--dx-font-l:"Times New Roman"');
    expect(model.css).toContain('--dx-font-e:"標楷體"');
    expect(model.css).toMatch(/font-size:16pt/);
  });
});
