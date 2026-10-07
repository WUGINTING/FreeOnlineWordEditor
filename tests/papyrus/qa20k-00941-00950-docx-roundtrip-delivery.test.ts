import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { Transform } from 'prosemirror-transform';
import { blankPackage } from '../../src/papyrus/docx/template';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { documentSections } from '../../src/papyrus/docx/sections';
import { ListCounter } from '../../src/papyrus/docx/numbering';
import { schema } from '../../src/papyrus/editor/schema';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const FINAL = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

async function importDoc(body: string, setup?: (zip: JSZip) => Promise<void>) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${W} ${R}><w:body>${body}${FINAL}</w:body></w:document>`);
  await setup?.(zip);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  return { bytes, ...(await readDocx(bytes)) };
}

function replaceText(doc: Awaited<ReturnType<typeof readDocx>>['doc'], target: string, replacement: string) {
  let from = -1;
  doc.descendants((node, pos) => {
    if (from < 0 && node.isText && node.text) {
      const at = node.text.indexOf(target);
      if (at >= 0) from = pos + at;
    }
  });
  if (from < 0) throw new Error(`Missing fixture text ${target}`);
  const tr = new Transform(doc);
  tr.replaceWith(from, from + target.length, schema.text(replacement));
  return tr.doc;
}

async function saveAndOpen(doc: Awaited<ReturnType<typeof readDocx>>['doc'], model: Awaited<ReturnType<typeof readDocx>>['model']) {
  const bytes = await writeDocx(doc, model);
  const zip = await JSZip.loadAsync(bytes);
  return { bytes, zip, ...(await readDocx(bytes)) };
}

describe('QA20K-00941–00950 DOCX roundtrip delivery workflows', () => {
  it('QA20K-00941 edits text in a multi-section report while retaining each section page size', async () => {
    const { doc, model } = await importDoc(
      `<w:p><w:r><w:t>Cover</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr><w:r><w:t>Approval copy</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Wide appendix</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
    );
    const edited = replaceText(doc, 'Approval copy', 'Approved copy');
    const reopened = await saveAndOpen(edited, model);
    expect(documentSections(reopened.doc, reopened.model).map(s => [s.page.width, s.page.height])).toEqual([[11906, 16838], [16838, 11906]]);
    expect(reopened.doc.textContent).toContain('Approved copy');
    const xml = await reopened.zip.file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<w:t>Approved copy<\/w:t>[\s\S]*?<w:sectPr>[\s\S]*?<w:pgSz w:w="16838" w:h="11906"/);
  });

  it('QA20K-00942 edits a page-break paragraph without losing Word compatibility settings', async () => {
    const { doc, model } = await importDoc('<w:p><w:r><w:t>Before handoff</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>After handoff</w:t></w:r></w:p>', async zip => {
      zip.file('word/settings.xml', `<?xml version="1.0"?><w:settings ${W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:suppressSpBfAfterPgBrk/></w:compat></w:settings>`);
    });
    const reopened = await saveAndOpen(replaceText(doc, 'Before handoff', 'Before approval'), model);
    expect(reopened.model.compat).toEqual({ mode: 15, suppressSpaceAfterPageBreak: true });
    const settings = await reopened.zip.file('word/settings.xml')!.async('string');
    expect(settings).toContain('w:name="compatibilityMode"');
    expect(settings).toContain('w:val="15"');
    expect(settings).toContain('<w:suppressSpBfAfterPgBrk');
    const modelSequence: string[] = [];
    reopened.doc.descendants(node => {
      if (node.type.name === 'page_break' || (node.type.name === 'raw_inline' && /<w:br\b[^>]*w:type="page"/.test(node.attrs.xml ?? ''))) modelSequence.push('<PAGE_BREAK>');
      else if (node.isText && node.text) modelSequence.push(node.text);
    });
    expect(modelSequence).toEqual(['Before approval', '<PAGE_BREAK>', 'After handoff']);
    const documentXml = await reopened.zip.file('word/document.xml')!.async('string');
    const xmlSequence = [...documentXml.matchAll(/<w:t(?: [^>]*)?>(.*?)<\/w:t>|<w:br\b(?=[^>]*w:type="page")[^>]*\/>/g)]
      .map(match => match[0].startsWith('<w:br') ? '<PAGE_BREAK>' : match[1]);
    expect(xmlSequence).toEqual(['Before approval', '<PAGE_BREAK>', 'After handoff']);
    expect(reopened.doc.textContent).toBe('Before approvalAfter handoff');
  });

  it('QA20K-00943 revises a paragraph while preserving first-line indent and exact line spacing', async () => {
    const { doc, model } = await importDoc('<w:p><w:pPr><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="exact"/><w:ind w:left="720" w:right="180" w:firstLine="360"/></w:pPr><w:r><w:t>Payment term</w:t></w:r></w:p>');
    const para = doc.firstChild!;
    expect(para.attrs).toMatchObject({ indLeft: 720, indRight: 180, indFirst: 360, spaceBefore: 120, spaceAfter: 240, line: 360, lineRule: 'exact' });
    const reopened = await saveAndOpen(replaceText(doc, 'Payment term', 'Payment period'), model);
    expect(reopened.doc.firstChild!.attrs).toMatchObject({ indLeft: 720, indRight: 180, indFirst: 360, spaceBefore: 120, spaceAfter: 240, line: 360, lineRule: 'exact' });
    const xml = await reopened.zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="exact"');
    expect(xml).toContain('<w:ind w:left="720" w:right="180" w:firstLine="360"');
    expect(reopened.doc.textContent).toBe('Payment period');
  });

  it('QA20K-00944 edits a clause while preserving its custom paragraph style inheritance', async () => {
    const { doc, model } = await importDoc('<w:p><w:pPr><w:pStyle w:val="ApprovalClause"/></w:pPr><w:r><w:t>Approved by finance</w:t></w:r></w:p>', async zip => {
      const styles = await zip.file('word/styles.xml')!.async('string');
      zip.file('word/styles.xml', styles.replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="ApprovalClause"><w:name w:val="Approval Clause"/><w:basedOn w:val="Heading2"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="300"/></w:pPr></w:style></w:styles>'));
    });
    expect(doc.firstChild!.attrs.styleId).toBe('ApprovalClause');
    const reopened = await saveAndOpen(replaceText(doc, 'Approved by finance', 'Approved by controller'), model);
    expect(reopened.doc.firstChild!.attrs.styleId).toBe('ApprovalClause');
    expect(reopened.model.paragraphStyles).toContainEqual({ id: 'ApprovalClause', name: 'Approval Clause' });
    const styles = await reopened.zip.file('word/styles.xml')!.async('string');
    expect(styles).toContain('w:styleId="ApprovalClause"');
    expect(styles).toContain('<w:basedOn w:val="Heading2"');
    expect(styles).toContain('<w:keepNext');
    expect(reopened.doc.textContent).toBe('Approved by controller');
  });

  it('QA20K-00945 updates a form control value but keeps the structured document tag metadata', async () => {
    const { doc, model } = await importDoc('<w:sdt><w:sdtPr><w:alias w:val="Invoice number"/><w:tag w:val="AP.InvoiceNumber"/><w:id w:val="3077"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>INV-204</w:t></w:r></w:p></w:sdtContent></w:sdt>');
    const reopened = await saveAndOpen(replaceText(doc, 'INV-204', 'INV-205'), model);
    expect(reopened.doc.textContent).toBe('INV-205');
    const xml = await reopened.zip.file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<w:sdt><w:sdtPr>[\s\S]*?<w:alias w:val="Invoice number"\/>[\s\S]*?<w:tag w:val="AP.InvoiceNumber"\/>[\s\S]*?<w:id w:val="3077"\/>[\s\S]*?<\/w:sdtPr><w:sdtContent>[\s\S]*?<w:t>INV-205<\/w:t>[\s\S]*?<\/w:sdtContent><\/w:sdt>/);
    expect(reopened.doc.firstChild!.attrs.wrap).toContain('AP.InvoiceNumber');
  });

  it('QA20K-00946 edits a sentence outside a bookmark and retains the named range around its clause', async () => {
    const { doc, model } = await importDoc('<w:p><w:r><w:t>Summary: </w:t></w:r><w:bookmarkStart w:id="9" w:name="PaymentTerms"/><w:r><w:t>net thirty</w:t></w:r><w:bookmarkEnd w:id="9"/><w:r><w:t xml:space="preserve"> days</w:t></w:r></w:p>');
    const reopened = await saveAndOpen(replaceText(doc, 'Summary:', 'Terms:'), model);
    expect(reopened.doc.textContent).toBe('Terms: net thirty days');
    const xml = await reopened.zip.file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<w:t[^>]*>Terms: <\/w:t>[\s\S]*?<w:bookmarkStart w:id="9" w:name="PaymentTerms"\/>[\s\S]*?<w:t>net thirty<\/w:t>[\s\S]*?<w:bookmarkEnd w:id="9"\/>[\s\S]*?<w:t[^>]*> days<\/w:t>/);
    expect(reopened.doc.toJSON().content?.[0].content?.filter((n: any) => n.type === 'raw_inline').map((n: any) => n.attrs.xml)).toEqual(['<w:bookmarkStart w:id="9" w:name="PaymentTerms"/>', '<w:bookmarkEnd w:id="9"/>']);
  });

  it('QA20K-00947 edits adjacent text while preserving an imported external hyperlink target', async () => {
    const { doc, model } = await importDoc('<w:p><w:r><w:t>Policy: </w:t></w:r><w:hyperlink r:id="rIdPolicy"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Travel standard</w:t></w:r></w:hyperlink></w:p>', async zip => {
      const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
      zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdPolicy" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://intranet.example.test/travel-policy" TargetMode="External"/></Relationships>'));
    });
    expect(doc.textContent).toBe('Policy: Travel standard');
    const reopened = await saveAndOpen(replaceText(doc, 'Policy:', 'Current policy:'), model);
    expect(reopened.doc.textContent).toBe('Current policy: Travel standard');
    let href: string | null = null;
    reopened.doc.descendants(node => { if (node.isText && node.text === 'Travel standard') href = node.marks.find(mark => mark.type.name === 'link')?.attrs.href ?? null; });
    expect(href).toBe('https://intranet.example.test/travel-policy');
    const rels = await reopened.zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).toContain('/hyperlink" Target="https://intranet.example.test/travel-policy"');
    expect(rels).toContain('Target="https://intranet.example.test/travel-policy"');
    expect(rels).toContain('TargetMode="External"');
  });

  it('QA20K-00948 revises a table cell while retaining the repeating header row and fixed grid', async () => {
    const table = '<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2200"/><w:gridCol w:w="3800"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:p><w:r><w:t>Owner</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Action</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Finance</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Review invoice</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const { doc, model } = await importDoc(table);
    expect(doc.firstChild!.child(0).attrs.header).toBe(true);
    const reopened = await saveAndOpen(replaceText(doc, 'Review invoice', 'Approve invoice'), model);
    const reopenedTable = reopened.doc.firstChild!;
    expect(reopenedTable.type.name).toBe('table');
    expect(reopenedTable.childCount).toBe(2);
    expect(reopenedTable.child(0).attrs.header).toBe(true);
    expect(reopenedTable.textContent).toContain('Approve invoice');
    const xml = await reopened.zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:tblLayout w:type="fixed"');
    expect(xml).toContain('<w:gridCol w:w="2200"');
    expect(xml).toContain('<w:gridCol w:w="3800"');
    expect(xml).toContain('<w:tblHeader');
  });

  it('QA20K-00949 changes a numbered procedure step and retains its non-default start override', async () => {
    const { doc, model } = await importDoc('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="8"/></w:numPr></w:pPr><w:r><w:t>Confirm amount</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="8"/></w:numPr></w:pPr><w:r><w:t>Submit invoice</w:t></w:r></w:p>', async zip => {
      zip.file('word/numbering.xml', `<?xml version="1.0"?><w:numbering ${W}><w:abstractNum w:abstractNumId="4"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="8"><w:abstractNumId w:val="4"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride></w:num></w:numbering>`);
      const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
      zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'));
      const types = await zip.file('[Content_Types].xml')!.async('string');
      zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'));
    });
    expect(doc.child(0).attrs.numId).toBe('8');
    const reopened = await saveAndOpen(replaceText(doc, 'Confirm amount', 'Verify amount'), model);
    expect(reopened.model.numbering.nums['8'].startOverrides).toEqual({ 0: 4 });
    const counter = new ListCounter(reopened.model.numbering);
    const numberedRows: Array<{ text: string; numId: string; ilvl: number; label: string }> = [];
    reopened.doc.forEach(paragraph => {
      if (paragraph.attrs.numId) {
        const numId = String(paragraph.attrs.numId);
        const ilvl = Number(paragraph.attrs.ilvl);
        numberedRows.push({ text: paragraph.textContent, numId, ilvl, label: counter.next(numId, ilvl) });
      }
    });
    expect(numberedRows).toEqual([
      { text: 'Verify amount', numId: '8', ilvl: 0, label: '4.' },
      { text: 'Submit invoice', numId: '8', ilvl: 0, label: '5.' },
    ]);
    expect(reopened.doc.textContent).toBe('Verify amountSubmit invoice');
    const numbering = await reopened.zip.file('word/numbering.xml')!.async('string');
    expect(numbering).toContain('<w:startOverride w:val="4"');
  });

  it('QA20K-00950 edits body text while preserving imported core properties for handoff', async () => {
    const { doc, model } = await importDoc('<w:p><w:r><w:t>Handoff draft</w:t></w:r></w:p>', async zip => {
      zip.file('docProps/core.xml', '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Quarterly handoff</dc:title><dc:creator>Operations</dc:creator><dc:subject>Internal review</dc:subject></cp:coreProperties>');
      const rootRels = await zip.file('_rels/.rels')!.async('string');
      zip.file('_rels/.rels', rootRels.replace('</Relationships>', '<Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'));
      const types = await zip.file('[Content_Types].xml')!.async('string');
      zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'));
    });
    const reopened = await saveAndOpen(replaceText(doc, 'Handoff draft', 'Handoff approved'), model);
    expect(reopened.doc.textContent).toBe('Handoff approved');
    const core = await reopened.zip.file('docProps/core.xml')!.async('string');
    expect(core).toContain('<dc:title>Quarterly handoff</dc:title>');
    expect(core).toContain('<dc:creator>Operations</dc:creator>');
    expect(core).toContain('<dc:subject>Internal review</dc:subject>');
    const rels = await reopened.zip.file('_rels/.rels')!.async('string');
    expect(rels).toContain('metadata/core-properties');
    expect(rels).toContain('Target="docProps/core.xml"');
  });
});
