import { describe, expect, it } from 'vitest';
import { EditorState } from 'prosemirror-state';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx, readRels } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { schema } from '../../src/papyrus/editor/schema';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const rel = (id: string, type: string, target: string, external = false) => `<Relationship Id="${id}" Type="${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

async function fixture(options: {
  body: string;
  mainRels?: string[];
  rootRels?: string[];
  settingsRels?: string[];
  parts?: Record<string, string | Uint8Array>;
  overrides?: Array<[string, string]>;
  settingsXml?: string;
}) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body>${options.body}<w:sectPr/></w:body></w:document>`);
  if (options.parts) for (const [path, bytes] of Object.entries(options.parts)) zip.file(path, bytes);
  const append = async (path: string, items: string[]) => {
    if (!items.length) return;
    const existing = zip.file(path);
    const old = existing
      ? await existing.async('string')
      : `<Relationships xmlns="${PKG_REL}"></Relationships>`;
    zip.file(path, old.replace('</Relationships>', `${items.join('')}</Relationships>`));
  };
  await append('word/_rels/document.xml.rels', options.mainRels ?? []);
  await append('_rels/.rels', options.rootRels ?? []);
  await append('word/_rels/settings.xml.rels', options.settingsRels ?? []);
  if (options.settingsXml) zip.file('word/settings.xml', options.settingsXml);
  if (options.overrides?.length) {
    const old = await zip.file('[Content_Types].xml')!.async('string');
    const entries = options.overrides.map(([path, type]) => `<Override PartName="/${path}" ContentType="${type}"/>`).join('');
    zip.file('[Content_Types].xml', old.replace('</Types>', `${entries}</Types>`));
  }
  return zip.generateAsync({ type: 'uint8array' });
}

async function editAndSave(input: Uint8Array) {
  const { doc, model } = await readDocx(input);
  const state = EditorState.create({ schema, doc });
  const edited = state.apply(state.tr.insertText('Reviewed: ', 1)).doc;
  expect(edited.textContent).toContain('Reviewed: ');
  const output = await JSZip.loadAsync(await writeDocx(edited, model));
  return output;
}

async function mainText(zip: JSZip) { return zip.file('word/document.xml')!.async('string'); }
async function packageText(zip: JSZip, path: string) { return zip.file(path)!.async('string'); }
async function assertRel(zip: JSZip, owner: string, id: string, type: string, target: string, external = false) {
  const item = (await readRels(zip, owner)).get(id);
  expect(item).toMatchObject({ id, type, target, external });
}
async function assertOverride(zip: JSZip, path: string, type: string) {
  const xml = await packageText(zip, '[Content_Types].xml');
  expect(xml).toContain(`<Override PartName="/${path}" ContentType="${type}"/>`);
}
async function tinyWorkbook() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="${PKG_REL}">${rel('rIdWorkbook', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'xl/workbook.xml')}</Relationships>`);
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Forecast" sheetId="1" r:id="rIdSheet"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="${PKG_REL}">${rel('rIdSheet', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', 'worksheets/sheet1.xml')}</Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Budget</t></is></c></row></sheetData></worksheet>');
  return zip.generateAsync({ type: 'uint8array' });
}

describe('QA20K DOCX package extension retention 01021–01030', () => {
  it('QA20K-01021 retains the footnote part and reference when editing the adjacent sentence', async () => {
    const notes = `<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p/></w:footnote><w:footnote w:id="4"><w:p><w:r><w:t>Receipt policy source</w:t></w:r></w:p></w:footnote></w:footnotes>`;
    const input = await fixture({ body: `<w:p><w:r><w:t>Expense claim</w:t></w:r><w:r><w:footnoteReference w:id="4"/></w:r></w:p>`, mainRels: [rel('rIdFoot', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes', 'footnotes.xml')], parts: { 'word/footnotes.xml': notes }, overrides: [['word/footnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml']] });
    const output = await editAndSave(input);
    expect(await mainText(output)).toContain('<w:footnoteReference w:id="4"/>');
    expect(await packageText(output, 'word/footnotes.xml')).toBe(notes);
    await assertRel(output, 'word/document.xml', 'rIdFoot', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes', 'word/footnotes.xml');
    await assertOverride(output, 'word/footnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml');
  });

  it('QA20K-01022 retains the endnote part and its note text during a body edit', async () => {
    const notes = `<w:endnotes ${W}><w:endnote w:type="separator" w:id="-1"><w:p/></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p/></w:endnote><w:endnote w:id="9"><w:p><w:r><w:t>Employment policy source</w:t></w:r></w:p></w:endnote></w:endnotes>`;
    const input = await fixture({ body: `<w:p><w:r><w:t>Policy summary</w:t></w:r><w:r><w:endnoteReference w:id="9"/></w:r></w:p>`, mainRels: [rel('rIdEnd', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes', 'endnotes.xml')], parts: { 'word/endnotes.xml': notes }, overrides: [['word/endnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml']] });
    const output = await editAndSave(input);
    expect(await mainText(output)).toContain('<w:endnoteReference w:id="9"/>');
    expect(await packageText(output, 'word/endnotes.xml')).toBe(notes);
    await assertRel(output, 'word/document.xml', 'rIdEnd', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes', 'word/endnotes.xml');
    await assertOverride(output, 'word/endnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml');
  });

  it('QA20K-01023 keeps an embedded quarterly workbook object and its package target', async () => {
    const workbook = await tinyWorkbook();
    const input = await fixture({ body: `<w:p><w:r><w:t>Budget source </w:t></w:r><w:r><w:object><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rIdOle"/></w:object></w:r></w:p>`, mainRels: [rel('rIdOle', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject', 'embeddings/Quarterly.xlsx')], parts: { 'word/embeddings/Quarterly.xlsx': workbook }, overrides: [['word/embeddings/Quarterly.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']] });
    const output = await editAndSave(input);
    expect(await mainText(output)).toContain('<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rIdOle"/>');
    expect(Array.from(await output.file('word/embeddings/Quarterly.xlsx')!.async('uint8array'))).toEqual(Array.from(workbook));
    await assertRel(output, 'word/document.xml', 'rIdOle', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject', 'word/embeddings/Quarterly.xlsx');
    await assertOverride(output, 'word/embeddings/Quarterly.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  });

  it('QA20K-01024 retains a linked workbook URI as an external OLE relationship', async () => {
    const uri = 'https://finance.example.test/shared/forecast.xlsx';
    const input = await fixture({ body: `<w:p><w:r><w:t>Linked forecast </w:t></w:r><w:r><w:object><o:OLEObject Type="Link" ProgID="Excel.Sheet.12" r:id="rIdLinked"/></w:object></w:r></w:p>`, mainRels: [rel('rIdLinked', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject', uri, true)] });
    const output = await editAndSave(input);
    expect(await mainText(output)).toContain('<o:OLEObject Type="Link" ProgID="Excel.Sheet.12" r:id="rIdLinked"/>');
    await assertRel(output, 'word/document.xml', 'rIdLinked', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject', uri, true);
  });

  it('QA20K-01025 preserves a custom XML expense data store and its properties relationship', async () => {
    const item = `<expense xmlns="urn:contoso:expense"><claim>EXP-8301</claim><amount currency="TWD">2450</amount></expense>`;
    const props = `<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:itemID>{A8C67D91-1AB5-4D4F-8B7E-12C177A1AC10}</ds:itemID></ds:datastoreItem>`;
    const input = await fixture({ body: `${p('Review sentence')}<w:customXml w:uri="urn:contoso:expense" w:element="claim"><w:p><w:r><w:t>Claim EXP-8301</w:t></w:r></w:p></w:customXml>`, mainRels: [rel('rIdCustom', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml', '../customXml/item1.xml')], parts: { 'customXml/item1.xml': item, 'customXml/itemProps1.xml': props, 'customXml/_rels/item1.xml.rels': `<Relationships xmlns="${PKG_REL}">${rel('rIdProps', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps', 'itemProps1.xml')}</Relationships>` }, overrides: [['customXml/item1.xml', 'application/xml'], ['customXml/itemProps1.xml', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml']] });
    const output = await editAndSave(input);
    const docXml = await mainText(output);
    expect(docXml).toContain('<w:customXml w:uri="urn:contoso:expense" w:element="claim">');
    expect(docXml).toContain('<w:t>Claim EXP-8301</w:t>');
    expect(await packageText(output, 'customXml/item1.xml')).toBe(item);
    expect(await packageText(output, 'customXml/itemProps1.xml')).toBe(props);
    await assertRel(output, 'word/document.xml', 'rIdCustom', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml', 'customXml/item1.xml');
    await assertRel(output, 'customXml/item1.xml', 'rIdProps', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps', 'customXml/itemProps1.xml');
    await assertOverride(output, 'customXml/itemProps1.xml', 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml');
  });

  it('QA20K-01026 preserves custom document properties used for routing a handoff', async () => {
    const custom = `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="DepartmentCode"><vt:lpwstr>FIN-OPS</vt:lpwstr></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="RetentionClass"><vt:lpwstr>R7</vt:lpwstr></property></Properties>`;
    const input = await fixture({ body: p('Handoff draft'), rootRels: [rel('rIdCustomProps', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties', 'docProps/custom.xml')], parts: { 'docProps/custom.xml': custom }, overrides: [['docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml']] });
    const output = await editAndSave(input);
    expect(await packageText(output, 'docProps/custom.xml')).toBe(custom);
    await assertRel(output, '', 'rIdCustomProps', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties', 'docProps/custom.xml');
    await assertOverride(output, 'docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml');
  });

  it('QA20K-01027 retains company metadata in extended app properties', async () => {
    const app = `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office Word</Application><Company>Contoso Finance</Company><Manager>Operations</Manager><Pages>8</Pages><Words>1240</Words></Properties>`;
    const input = await fixture({ body: p('Monthly report'), rootRels: [rel('rIdApp', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'docProps/app.xml')], parts: { 'docProps/app.xml': app }, overrides: [['docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml']] });
    const output = await editAndSave(input);
    expect(await packageText(output, 'docProps/app.xml')).toBe(app);
    await assertRel(output, '', 'rIdApp', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'docProps/app.xml');
    await assertOverride(output, 'docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');
  });

  it('QA20K-01028 retains the corporate theme part and its package relationship', async () => {
    const theme = `<a:theme ${W} name="Contoso Finance"><a:themeElements><a:clrScheme name="Corporate"><a:dk1><a:srgbClr val="17365D"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme></a:themeElements></a:theme>`;
    const input = await fixture({ body: p('Quarterly summary'), mainRels: [rel('rIdTheme', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml')], parts: { 'word/theme/theme1.xml': theme }, overrides: [['word/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml']] });
    const output = await editAndSave(input);
    expect(await packageText(output, 'word/theme/theme1.xml')).toBe(theme);
    await assertRel(output, 'word/document.xml', 'rIdTheme', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'word/theme/theme1.xml');
    await assertOverride(output, 'word/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
  });

  it('QA20K-01029 preserves the embedded company font file and font table link', async () => {
    const fontTable = `<w:fonts ${W}><w:font w:name="Contoso Sans"><w:embedRegular r:id="rIdEmbeddedFont"/></w:font></w:fonts>`;
    const fontRels = `<Relationships xmlns="${PKG_REL}">${rel('rIdEmbeddedFont', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font', 'fonts/contoso.odttf')}</Relationships>`;
    const font = new Uint8Array([0x4f, 0x54, 0x54, 0x4f, 0x2d, 0x46, 0x4f, 0x4e, 0x54]);
    const input = await fixture({ body: p('Policy approval'), mainRels: [rel('rIdFontTable', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable', 'fontTable.xml')], parts: { 'word/fontTable.xml': fontTable, 'word/_rels/fontTable.xml.rels': fontRels, 'word/fonts/contoso.odttf': font }, overrides: [['word/fontTable.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'], ['word/fonts/contoso.odttf', 'application/vnd.openxmlformats-officedocument.obfuscatedFont']] });
    const output = await editAndSave(input);
    expect(await packageText(output, 'word/fontTable.xml')).toBe(fontTable);
    expect(await packageText(output, 'word/_rels/fontTable.xml.rels')).toBe(fontRels);
    expect(Array.from(await output.file('word/fonts/contoso.odttf')!.async('uint8array'))).toEqual(Array.from(font));
    await assertRel(output, 'word/document.xml', 'rIdFontTable', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable', 'word/fontTable.xml');
    await assertRel(output, 'word/fontTable.xml', 'rIdEmbeddedFont', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font', 'word/fonts/contoso.odttf');
    await assertOverride(output, 'word/fontTable.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml');
    await assertOverride(output, 'word/fonts/contoso.odttf', 'application/vnd.openxmlformats-officedocument.obfuscatedFont');
  });

  it('QA20K-01030 preserves the attached corporate template URI from settings', async () => {
    const settings = `<w:settings ${W}><w:zoom w:percent="100"/><w:attachedTemplate r:id="rIdTemplate"/><w:compat/></w:settings>`;
    const templateUrl = 'https://templates.example.test/Finance.dotx';
    const input = await fixture({ body: p('Finance handoff'), mainRels: [rel('rIdSettings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml')], settingsRels: [rel('rIdTemplate', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate', templateUrl, true)], settingsXml: settings });
    const output = await editAndSave(input);
    expect(await packageText(output, 'word/settings.xml')).toBe(settings);
    await assertRel(output, 'word/document.xml', 'rIdSettings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'word/settings.xml');
    await assertRel(output, 'word/settings.xml', 'rIdTemplate', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate', templateUrl, true);
    await assertOverride(output, 'word/settings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  });
});
