// The compatibility report must name what the web editor can't fully show or edit, with counts,
// and stay empty for documents that have none of it.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { blankPackage } from '../../src/papyrus/docx/template';
import { scanCompat, type CompatReport } from '../../src/papyrus/docx/compat';
import { commentThreads, parseComments, readComments } from '../../src/papyrus/docx/comments';
import { formatDate } from '../../src/papyrus/docx/revisions';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const R = (id: number) => `w:id="${id}" w:author="A" w:date="2026-01-01T00:00:00Z"`;

async function pkg(body: string, extra: (zip: JSZip) => void = () => {}): Promise<JSZip> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`);
  extra(zip);
  // Round-trip through bytes like an upload.
  return JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }));
}

const byId = (r: CompatReport) => Object.fromEntries(r.items.map((i) => [i.id, i]));

async function addRel(zip: JSZip, type: string, target: string, contentType: string) {
  const relsPath = 'word/_rels/document.xml.rels';
  const rels = await zip.file(relsPath)!.async('string');
  zip.file(relsPath, rels.replace('</Relationships>', `<Relationship Id="rIdT${target.length}" Type="${type}" Target="${target}"/></Relationships>`));
  const ct = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', ct.replace('</Types>', `<Override PartName="/word/${target}" ContentType="${contentType}"/></Types>`));
}

describe('compatibility report', () => {
  it('is empty for a plain document and a new one', async () => {
    expect((await scanCompat(await pkg(`<w:p><w:r><w:t>hello</w:t></w:r></w:p>${SECT}`))).items).toEqual([]);
    expect((await scanCompat(blankPackage())).items).toEqual([]);
    expect((await scanCompat(null)).items).toEqual([]);
  });

  it('counts kept-as-is objects without double counting mc:Fallback copies', async () => {
    const textbox =
      '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
      '<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p/></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>';
    const shape = '<w:r><w:drawing><wp:anchor><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp/></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
    const chart = '<w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></wp:inline></w:drawing></w:r>';
    const smart = '<w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphic></wp:inline></w:drawing></w:r>';
    const floating = '<w:r><w:drawing><wp:anchor><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic></wp:anchor></w:drawing></w:r>';
    const math = '<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara><m:oMath/>';
    const report = await scanCompat(await pkg(`<w:p>${textbox}${textbox}${shape}${chart}${smart}${floating}${math}<w:r><w:object/></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>${SECT}`));
    const r = byId(report);
    expect(r.textbox.count).toBe(2);
    expect(r.textbox.where).toBe('內文 2 處');
    expect(r.shape.count).toBe(1);
    expect(r.chart.count).toBe(1);
    expect(r.smartart.count).toBe(1);
    expect(r.floatingImage.count).toBe(1);
    expect(r.equation.count).toBe(2);
    expect(r.ole.count).toBe(1);
    expect(r.notes.count).toBe(1);
    for (const item of report.items) {
      expect(item.effect.length).toBeGreaterThan(10);
      expect(item.effect).not.toMatch(/完整相容|完全相容/);
    }
  });

  it('reports tracked changes it can handle, and separately those it can’t', async () => {
    const body =
      `<w:p><w:pPr><w:rPr><w:ins ${R(1)}/></w:rPr></w:pPr><w:ins ${R(2)}><w:r><w:t>a</w:t></w:r></w:ins><w:del ${R(3)}><w:r><w:delText>b</w:delText></w:r></w:del>` +
      `<w:r><w:rPr><w:b/><w:rPrChange ${R(4)}><w:rPr/></w:rPrChange></w:rPr><w:t>c</w:t></w:r></w:p>` +
      `<w:tbl><w:tr><w:trPr><w:ins ${R(5)}/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>` + SECT;
    const r = byId(await scanCompat(await pkg(body)));
    expect(r.revisions.count).toBe(4);
    expect(r.revisions.where).toBe('插入 1、刪除 1、格式變更 1、段落標記 1');
    // The web has a 追蹤修訂 mode (GOV-FINDING-003): the notice must say so, not the opposite.
    expect(r.revisions.effect).toContain('也會記錄為修訂');
    expect(r.revisions.effect).not.toContain('不會記錄為修訂');
    expect(r.revisionsUnsupported.count).toBe(1);
  });

  it('counts revisions in text boxes and other kept content as unsupported, and those in a deletion once', async () => {
    const txbx =
      '<w:r><w:drawing><wp:anchor><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="t"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      `<wps:wsp><wps:txbx><w:txbxContent><w:p><w:ins ${R(1)}><w:r><w:t>in box</w:t></w:r></w:ins><w:r><w:rPr><w:b/><w:rPrChange ${R(2)}><w:rPr/></w:rPrChange></w:rPr><w:t>b</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>` +
      '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
    const math = `<m:oMath><w:ins ${R(3)}><m:r><m:t>x</m:t></m:r></w:ins></m:oMath>`;
    const del = `<w:del ${R(4)}><w:r><w:rPr><w:b/><w:rPrChange ${R(5)}><w:rPr/></w:rPrChange></w:rPr><w:delText>gone</w:delText></w:r></w:del>`;
    const r = byId(await scanCompat(await pkg(`<w:p>${txbx}${math}${del}</w:p>${SECT}`)));
    expect(r.revisions.count).toBe(1);
    expect(r.revisions.where).toBe('刪除 1');
    expect(r.revisionsUnsupported.count).toBe(3);
    expect(r.textbox.count).toBe(1);
  });

  it('lists fields that show a cached result, but not live page numbers', async () => {
    const fld = (instr: string) =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    // A PAGEREF inside the TOC's result is updated with the TOC: not counted on its own.
    const toc =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      `${fld('PAGEREF _Toc1 \\h')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const body = `<w:p>${toc}${fld('PAGE')}${fld('REF _Ref1 \\h')}<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>3</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" DATE "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>${SECT}`;
    const r = byId(await scanCompat(await pkg(body)));
    expect(r.fields.count).toBe(3);
    expect(r.fields.where).toBe('目錄（TOC） 1、交互參照（REF） 1、日期（DATE） 1');
  });

  it('names the sections with settings the editor doesn’t show (not sections themselves)', async () => {
    const s1 = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:cols w:num="2" w:space="425"/><w:lnNumType w:countBy="1"/></w:sectPr>';
    const body = `<w:p><w:pPr>${s1}</w:pPr></w:p><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:space="425"/><w:pgBorders><w:top w:val="single"/></w:pgBorders><w:pgNumType w:fmt="lowerRoman" w:chapStyle="1"/></w:sectPr>`;
    const r = byId(await scanCompat(await pkg(body)));
    expect(r.columns.where).toBe('第 1 節（共 2 節）');
    expect(r.lineNumbers.where).toBe('第 1 節（共 2 節）');
    expect(r.pageBorders.where).toBe('第 2 節（共 2 節）');
    expect(r.pageNumberFormat.count).toBe(1);
    // Different page sizes / orientations per section are supported: not listed.
    expect(Object.keys(r).sort()).toEqual(['columns', 'lineNumbers', 'pageBorders', 'pageNumberFormat']);
  });

  it('page number formats the editor shows like Word (roman, letters, Chinese) are not listed', async () => {
    const shown = '<w:p><w:pPr><w:sectPr><w:pgNumType w:fmt="lowerRoman" w:start="1"/></w:sectPr></w:pPr></w:p>';
    const body = `${shown}<w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgNumType w:fmt="taiwaneseCounting" w:start="1"/></w:sectPr>`;
    expect(byId(await scanCompat(await pkg(body))).pageNumberFormat).toBeUndefined();
    const rare = `<w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgNumType w:fmt="hebrew1"/></w:sectPr>`;
    expect(byId(await scanCompat(await pkg(rare))).pageNumberFormat.count).toBe(1);
  });

  it('content controls, locks, right-to-left text, protection, macros and fonts', async () => {
    const body =
      '<w:p><w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>' +
      `<w:r><w:rPr><w:rtl/></w:rPr><w:t>שלום</w:t></w:r></w:p>${SECT}`;
    const zip = await pkg(body, (z) => {
      z.file('word/vbaProject.bin', 'x');
      z.file('word/fonts/font1.odttf', 'x');
      z.file('word/media/image1.emf', 'x');
    });
    const settings = await zip.file('word/settings.xml')?.async('string');
    if (settings) zip.file('word/settings.xml', settings.replace(/(<w:settings[^>]*>)/, '$1<w:documentProtection w:edit="readOnly" w:enforcement="1"/>'));
    else {
      zip.file('word/settings.xml', `<w:settings ${NS}><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>`);
      await addRel(zip, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
    }
    const r = byId(await scanCompat(zip));
    expect(r.contentControls.count).toBe(1);
    expect(r.lockedControls.count).toBe(1);
    expect(r.rtl.count).toBe(1);
    expect(r.macros.count).toBe(1);
    expect(r.embeddedFonts.count).toBe(1);
    expect(r.legacyImages.count).toBe(1);
    expect(r.protection.count).toBe(1);
  });

  it('counts comments and headers/footers separately from the body', async () => {
    const zip = await pkg(`<w:p><w:r><w:t>x</w:t></w:r></w:p>${SECT}`);
    zip.file('word/comments.xml', `<w:comments ${NS}><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:comment></w:comments>`);
    await addRel(zip, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
    zip.file('word/header9.xml', `<w:hdr ${NS}><w:p><w:r><w:pict><v:shape><v:textpath string="DRAFT"/></v:shape></w:pict></w:r></w:p></w:hdr>`);
    await addRel(zip, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header9.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
    const r = byId(await scanCompat(zip));
    expect(r.comments.count).toBe(1);
    // Comments can be added and answered on the web (GOV-ISSUE-007 / GOV-197): the notice says so.
    expect(r.comments.effect).toContain('新增、回覆');
    expect(r.comments.effect).not.toContain('無法新增');
    expect(r.shape.where).toBe('頁首頁尾 1 處');
  });

  it('a watermark Word made (設計 › 浮水印) is shown and editable: not an unsupported shape', async () => {
    const zip = await pkg(`<w:p><w:r><w:t>x</w:t></w:r></w:p>${SECT}`);
    const text = '<v:shape id="PowerPlusWaterMarkObject357831064" type="#_x0000_t136" style="position:absolute;width:415pt;height:207pt;rotation:315" fillcolor="silver" stroked="f"><v:textpath style="font-family:&quot;標楷體&quot;;font-size:1pt" string="機密"/></v:shape>';
    const picture = '<v:shape id="WordPictureWatermark1" type="#_x0000_t75" style="position:absolute;width:415pt;height:207pt"><v:imagedata r:id="rId1" o:title="" gain="19661f" blacklevel="22938f"/></v:shape>';
    zip.file('word/header9.xml', `<w:hdr ${NS} xmlns:o="urn:schemas-microsoft-com:office:office"><w:p><w:r><w:pict>${text}</w:pict></w:r><w:r><w:pict>${picture}</w:pict></w:r></w:p></w:hdr>`);
    await addRel(zip, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header9.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
    expect(byId(await scanCompat(zip)).shape).toBeUndefined();
  });
});

describe('comments', () => {
  it('reads author, date, text, replies and resolved state', () => {
    const comments =
      `<w:comments ${NS}>` +
      '<w:comment w:id="1" w:author="Ann" w:date="2026-01-02T03:04:00Z" w:initials="A"><w:p w14:paraId="0000000A"><w:r><w:annotationRef/></w:r><w:r><w:t>First line</w:t></w:r></w:p><w:p w14:paraId="0000000B"><w:r><w:t>second</w:t></w:r></w:p></w:comment>' +
      '<w:comment w:id="2" w:author="Ben"><w:p w14:paraId="0000000C"><w:r><w:t>Agreed</w:t></w:r></w:p></w:comment>' +
      '</w:comments>';
    const ext = `<w15:commentsEx ${NS}><w15:commentEx w15:paraId="0000000B" w15:done="1"/><w15:commentEx w15:paraId="0000000C" w15:paraIdParent="0000000B" w15:done="0"/></w15:commentsEx>`;
    expect(parseComments(comments, ext)).toEqual([
      { id: '1', author: 'Ann', initials: 'A', date: '2026-01-02T03:04:00Z', dateUtc: null, text: 'First line\nsecond', parentId: null, done: true },
      { id: '2', author: 'Ben', initials: null, date: null, dateUtc: null, text: 'Agreed', parentId: '1', done: false },
    ]);
  });

  it('takes the real UTC time from commentsIds + commentsExtensible, and shows w:date as written', () => {
    const comments =
      `<w:comments ${NS}><w:comment w:id="1" w:author="Ann" w:date="2026-09-24T12:01:00Z"><w:p w14:paraId="0000000A"><w:r><w:t>hi</w:t></w:r></w:p></w:comment>` +
      '<w:comment w:id="2" w:author="Ben" w:date="2026-09-24T12:05:00Z"><w:p w14:paraId="0000000B"><w:r><w:t>no utc</w:t></w:r></w:p></w:comment></w:comments>';
    const ids = '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="0000000A" w16cid:durableId="7A7A7A7A"/></w16cid:commentsIds>';
    const extensible = '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"><w16cex:commentExtensible w16cex:durableId="7A7A7A7A" w16cex:dateUtc="2026-09-24T04:01:00Z"/></w16cex:commentsExtensible>';
    const [a, b] = parseComments(comments, { ids, extensible });
    expect(a.dateUtc).toBe('2026-09-24T04:01:00Z');
    const utc = new Date('2026-09-24T04:01:00Z');
    const p = (n: number) => String(n).padStart(2, '0');
    expect(formatDate(a.date, a.dateUtc)).toBe(`${utc.getFullYear()}/${p(utc.getMonth() + 1)}/${p(utc.getDate())} ${p(utc.getHours())}:${p(utc.getMinutes())}`);
    // Word wrote its local time with a "Z": the digits are shown as they are, whatever the time zone.
    expect(b.dateUtc).toBe(null);
    expect(formatDate(b.date, b.dateUtc)).toBe('2026/09/24 12:05');
  });

  it('threads: replies to replies stay in their thread; loops, missing parents and repeated ids hide nothing', () => {
    const c = (id: string, parentId: string | null) => ({ id, author: id, initials: null, date: null, dateUtc: null, text: id, parentId, done: false });
    const list = [c('1', null), c('2', '1'), c('3', '2'), c('4', '5'), c('5', '4'), c('6', 'gone'), c('1', null)];
    const threads = commentThreads(list);
    expect(threads.map((t) => [t.comment.text, t.replies.map((r) => r.text)])).toEqual([
      ['1', ['2', '3']], ['4', []], ['5', []], ['6', []], ['1', []],
    ]);
    // Every comment is shown exactly once, and keys are unique.
    expect(threads.flatMap((t) => [t.comment, ...t.replies]).length).toBe(list.length);
    expect(new Set(threads.map((t) => t.key)).size).toBe(threads.length);
  });

  it('finds the comments part through the document relationships', async () => {
    const zip = await pkg(`<w:p/>${SECT}`);
    expect(await readComments(zip)).toEqual([]);
    zip.file('word/comments.xml', `<w:comments ${NS}><w:comment w:id="5" w:author="Q"><w:p><w:r><w:t>ok</w:t></w:r></w:p></w:comment></w:comments>`);
    await addRel(zip, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
    expect((await readComments(zip)).map((c) => [c.id, c.author, c.text])).toEqual([['5', 'Q', 'ok']]);
  });
});
