// persona-300: the file title follows a rename, but only a title the editor wrote itself. The
// editor records the title it writes in a custom document property (docx/titleMarker.ts).
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { docxOf, domStubs } from './p300Helpers';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { TITLE_PROPERTY } from '../../src/papyrus/docx/titleMarker';
import { DocxEditor } from '../../src/papyrus/editor/core';

domStubs();

const BODY = '<w:p><w:r><w:t>內容</w:t></w:r></w:p>';
const CUSTOM_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';

async function save(bytes: Uint8Array, title?: string): Promise<Uint8Array> {
  const { doc, model } = await readDocx(bytes);
  return writeDocx(doc, model, title ? { title } : {});
}
const part = async (bytes: Uint8Array | Blob, path: string) => (await JSZip.loadAsync(bytes)).file(path)?.async('string');
const titleOf = async (bytes: Uint8Array | Blob) => /<dc:title>([^<]*)<\/dc:title>/.exec((await part(bytes, 'docProps/core.xml')) ?? '')?.[1] ?? null;

/** What Word does when the author types a title in 檔案 › 資訊 (custom properties kept). */
async function authorSetsTitle(bytes: Uint8Array, title: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('docProps/core.xml')!.async('string');
  zip.file('docProps/core.xml', xml.replace(/<dc:title>[^<]*<\/dc:title>/, `<dc:title>${title}</dc:title>`));
  return zip.generateAsync({ type: 'uint8array' });
}

describe('the file title after a rename', () => {
  it('a title the editor wrote follows the new name, save after save', async () => {
    const first = await save(await docxOf(BODY), '會議紀錄');
    expect(await titleOf(first)).toBe('會議紀錄');
    const custom = (await part(first, 'docProps/custom.xml'))!;
    expect(custom).toContain(`name="${TITLE_PROPERTY}"><vt:lpwstr>會議紀錄</vt:lpwstr>`);
    expect(await part(first, '_rels/.rels')).toContain(`Type="${CUSTOM_REL}" Target="docProps/custom.xml"`);
    expect(await part(first, '[Content_Types].xml')).toContain(
      '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>',
    );

    // Opened again (the saved version), the document renamed: the title follows.
    const second = await save(first, '會議紀錄（修正）');
    expect(await titleOf(second)).toBe('會議紀錄（修正）');
    expect(await part(second, 'docProps/custom.xml')).toContain('<vt:lpwstr>會議紀錄（修正）</vt:lpwstr>');
    // One property, one relationship: nothing piles up.
    expect((await part(second, 'docProps/custom.xml'))!.match(/<property /g)).toHaveLength(1);
    expect((await part(second, '_rels/.rels'))!.match(new RegExp(CUSTOM_REL, 'g'))).toHaveLength(1);

    // Same name again: core.xml and custom.xml stay byte for byte.
    const third = await save(second, '會議紀錄（修正）');
    expect(await part(third, 'docProps/core.xml')).toBe(await part(second, 'docProps/core.xml'));
    expect(await part(third, 'docProps/custom.xml')).toBe(await part(second, 'docProps/custom.xml'));
  });

  it('a title the author typed in Word is never changed', async () => {
    const saved = await save(await docxOf(BODY), '會議紀錄');
    const edited = await authorSetsTitle(saved, '第 3 次工作會議');
    const again = await save(edited, '會議紀錄（修正）');
    expect(await titleOf(again)).toBe('第 3 次工作會議');
    expect(await part(again, 'docProps/core.xml')).toBe(await part(edited, 'docProps/core.xml'));
    expect(await part(again, 'docProps/custom.xml')).toBe(await part(edited, 'docProps/custom.xml'));
  });

  it('a file with its own title and no record keeps it, and gets no record', async () => {
    const zip = await JSZip.loadAsync(await docxOf(BODY));
    const core = await zip.file('docProps/core.xml')!.async('string');
    zip.file('docProps/core.xml', core.replace('</cp:coreProperties>', '<dc:title>局長交辦事項</dc:title></cp:coreProperties>'));
    const own = await zip.generateAsync({ type: 'uint8array' });
    const out = await save(own, '局長交辦事項');
    expect(await titleOf(out)).toBe('局長交辦事項');
    expect(await part(out, 'docProps/custom.xml')).toBeUndefined();
    expect(await titleOf(await save(own, '別的名稱'))).toBe('局長交辦事項');
  });

  it('other custom properties stay; the record is added beside them', async () => {
    const custom =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="文號"><vt:lpwstr>北市文字第 1 號</vt:lpwstr></property></Properties>';
    const zip = await JSZip.loadAsync(await docxOf(BODY));
    zip.file('docProps/custom.xml', custom);
    const rels = await zip.file('_rels/.rels')!.async('string');
    zip.file('_rels/.rels', rels.replace('</Relationships>', `<Relationship Id="rId9" Type="${CUSTOM_REL}" Target="docProps/custom.xml"/></Relationships>`));
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const out = await save(bytes, '公告');
    const xml = (await part(out, 'docProps/custom.xml'))!;
    expect(xml).toContain('name="文號"><vt:lpwstr>北市文字第 1 號</vt:lpwstr>');
    expect(xml).toContain(`pid="3" name="${TITLE_PROPERTY}"><vt:lpwstr>公告</vt:lpwstr>`);
    expect(await titleOf(await save(out, '公告（二）'))).toBe('公告（二）');
  });

  it('the record is found and changed however its name is written (single quotes, character references)', async () => {
    const first = await save(await docxOf(BODY), '會議紀錄');
    for (const name of [`name='${TITLE_PROPERTY}'`, `name="${[...TITLE_PROPERTY].map((c) => `&#${c.codePointAt(0)};`).join('')}"`]) {
      // Another program wrote the part again, its own way.
      const zip = await JSZip.loadAsync(first);
      const xml = await zip.file('docProps/custom.xml')!.async('string');
      zip.file('docProps/custom.xml', xml.replace(`name="${TITLE_PROPERTY}"`, name));
      const renamed = await save(await zip.generateAsync({ type: 'uint8array' }), '會議紀錄（修正）');
      expect(await titleOf(renamed)).toBe('會議紀錄（修正）');
      const custom = (await part(renamed, 'docProps/custom.xml'))!;
      expect(custom).toContain(`${name}><vt:lpwstr>會議紀錄（修正）</vt:lpwstr></property>`);
      expect(custom.match(/<property /g)).toHaveLength(1);
    }
  });

  it('without a name nothing about the title is written', async () => {
    const out = await save(await docxOf(BODY));
    expect(await titleOf(out)).toBeNull();
    expect(await part(out, 'docProps/custom.xml')).toBeUndefined();
  });

  it('the editor: renamed while open (setDocumentTitle), the next save carries the new name', async () => {
    const host = document.createElement('div');
    const editor = new DocxEditor(host, { documentTitle: '簽稿' });
    await editor.open(await save(await docxOf(BODY), '簽稿'));
    editor.setDocumentTitle('簽稿（核定）');
    const saved = await editor.save();
    expect(await titleOf(saved)).toBe('簽稿（核定）');
    editor.destroy();
  });
});
