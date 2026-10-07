// @vitest-environment jsdom
// Word's 浮水印 (設計 › 頁面背景 › 浮水印): a text or picture watermark is a VML shape in the header
// parts, which Word draws centred on every page behind the text. These tests pin down:
//  - the XML Word writes (so Word's own 浮水印 dialog recognises, edits and removes it), in every
//    header part the sections refer to, a header made where a section uses none;
//  - replacing and removing it (Word's shapes and ours), other header content untouched;
//  - a Word-made watermark is recognised, drawn on each page (and in the printout) and kept byte
//    for byte through unrelated edits;
//  - the compatibility notice no longer calls it an unsupported shape;
//  - the ribbon's 設計 tab, the 浮水印 menu and the 自訂浮水印 dialog;
//  - refused while 追蹤修訂 is on, and in read-only mode.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { createApp, h, nextTick, ref } from 'vue';
import { history, redo, undo } from 'prosemirror-history';
import { AllSelection, EditorState, TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { schema } from '../../src/papyrus/editor/schema';
import { scanCompat } from '../../src/papyrus/docx/compat';
import {
  WATERMARK_PRESETS, isWatermarkRun, readWatermark, textWatermark, textWatermarkSize, watermarkRunXml, type Watermark,
} from '../../src/papyrus/docx/watermark';
import { measureWatermarkText, watermarkTargets } from '../../src/papyrus/editor/watermark';
import { documentSections } from '../../src/papyrus/docx/sections';
import { readDocx } from '../../src/papyrus/docx/reader';
import { rejectAllRevisions } from '../../src/papyrus/editor/review';
import { keepSectionReferences } from '../../src/papyrus/editor/sectionReferences';
import { Slice, Fragment } from 'prosemirror-model';
import WatermarkDialog from '../../src/papyrus/vue/WatermarkDialog.vue';
import DocxEditorVue from '../../src/papyrus/vue/DocxEditor.vue';

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const PNG_URL = `data:image/png;base64,${PNG}`;

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
/** The root Word 2016 writes on a header part. */
const HDR_ROOT =
  '<w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
  'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">';

/** Word's definition of the WordArt text-path shape type (#_x0000_t136), written once per part. */
const SHAPETYPE_136 =
  '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">' +
  '<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>' +
  '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>' +
  '<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>' +
  '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>' +
  '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>' +
  '<o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>';

/** A watermark as Word 2016 writes it with 設計 › 浮水印 › 機密 1 (CONFIDENTIAL), plus header text. */
const WORD_RUN =
  `<w:r><w:rPr><w:noProof/></w:rPr><w:pict w14:anchorId="0B3A5E7C">${SHAPETYPE_136}` +
  '<v:shape id="PowerPlusWaterMarkObject357831064" o:spid="_x0000_s2049" type="#_x0000_t136" ' +
  'style="position:absolute;margin-left:0;margin-top:0;width:527.85pt;height:131.95pt;rotation:315;z-index:-251657216;' +
  'mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" ' +
  'o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/>' +
  '<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="CONFIDENTIAL"/>' +
  '<w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>';
const WORD_HEADER =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${HDR_ROOT}<w:p w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3">` +
  `<w:pPr><w:pStyle w:val="a3"/></w:pPr>${WORD_RUN}<w:r><w:t>頁首文字</w:t></w:r></w:p></w:hdr>`;

const SECT = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>';
const HEADER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class { observe() {} unobserve() {} disconnect() {} };
  Range.prototype.getClientRects ??= (() => Object.assign([], { item: () => null })) as any;
  Range.prototype.getBoundingClientRect ??= (() => ({ top: 0, bottom: 0, left: 0, right: 0, width: 0, height: 0 })) as any;
  (Element.prototype as any).scrollIntoView ??= () => {};
});
afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
  document.body.replaceChildren();
});

/** A package with this body and these header parts (word/<name> → rId<name>). */
async function pkg(body: string, headers: Record<string, string> = {}) {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}</w:body></w:document>`);
  let rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  let types = await zip.file('[Content_Types].xml')!.async('string');
  for (const [id, xml] of Object.entries(headers)) {
    zip.file(`word/${id}.xml`, xml);
    rels = rels.replace('</Relationships>', `<Relationship Id="rId_${id}" Type="${HEADER_REL}" Target="${id}.xml"/></Relationships>`);
    types = types.replace('</Types>', `<Override PartName="/word/${id}.xml" ContentType="${HEADER_TYPE}"/></Types>`);
  }
  zip.file('word/_rels/document.xml.rels', rels);
  zip.file('[Content_Types].xml', types);
  return zip;
}

async function open(zip: JSZip, options: ConstructorParameters<typeof DocxEditor>[1] = {}) {
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const host = document.createElement('div');
  document.body.append(host);
  const ed = new DocxEditor(host, options);
  await ed.open(bytes);
  return { ed, host, bytes };
}

/** One page per section (jsdom lays nothing out); `first` adds the section's second page. */
function layOut(ed: DocxEditor, pages: { section: number; inSection: number }[]) {
  const page = (ed as any).firstPage();
  (ed as any).pages = pages.map((p, i) => ({ ...page, ...p, number: i + 1, top: i * (page.height + 24) }));
  (ed as any).renderPages((ed as any).pages);
}

async function saved(ed: DocxEditor) {
  return JSZip.loadAsync(await ed.save());
}
const text = (zip: JSZip, path: string) => zip.file(path)!.async('string');

/** The header parts of the saved package that each section's references point at, by type. */
async function sectionHeaders(zip: JSZip) {
  const document = await text(zip, 'word/document.xml');
  const rels = await text(zip, 'word/_rels/document.xml.rels');
  const target = (id: string) => new RegExp(`Id="${id}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${id}"`).exec(rels);
  return [...document.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)].map((m) => {
    const out: Record<string, string> = {};
    for (const ref of m[0].matchAll(/<w:headerReference\b[^>]*>/g)) {
      const type = /w:type="(\w+)"/.exec(ref[0])![1];
      const id = /r:id="([^"]+)"/.exec(ref[0])![1];
      const t = target(id);
      out[type] = 'word/' + (t?.[1] ?? t?.[2]);
    }
    return out;
  });
}

/** Two sections: the first with 「不同的首頁」 and header1 (「甲」) as its default header; the second links to it. */
const TWO_SECTIONS =
  `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}<w:titlePg/></w:sectPr></w:pPr><w:r><w:t>第一節</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>第二節</w:t></w:r></w:p><w:sectPr>${SECT}</w:sectPr>`;
const PLAIN_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:hdr>`;

/** The watermark run our writer produces for a text watermark, with its numbers as patterns. */
const OUR_TEXT_RUN = (opts: { text: string; font: string; size?: string; rotation?: boolean; fill?: string; semi?: boolean; shapetype?: boolean }) =>
  new RegExp(
    '^' +
      escape('<w:r><w:rPr><w:noProof/></w:rPr><w:pict>') +
      (opts.shapetype === false ? '' : escape(SHAPETYPE_136)) +
      escape('<v:shape id="PowerPlusWaterMarkObject') + '\\d+' + escape('" o:spid="_x0000_s') + '\\d+' +
      escape('" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:') + '[\\d.]+' + escape('pt;height:') + '[\\d.]+' + escape('pt;') +
      (opts.rotation === false ? '' : escape('rotation:315;')) +
      escape('z-index:') + '-\\d+' + escape(';mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" ') +
      escape(`fillcolor="${opts.fill ?? 'silver'}" stroked="f">`) +
      (opts.semi === false ? '' : escape('<v:fill opacity=".5"/>')) +
      escape(`<v:textpath style="font-family:&quot;${opts.font}&quot;;font-size:${opts.size ?? '1pt'}" string="${opts.text}"/>`) +
      escape('<w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>') +
      '$',
  );
function escape(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
/** The watermark runs (w:r holding a w:pict) of a part. */
const pictRuns = (xml: string) => [...xml.matchAll(/<w:r>(?:(?!<w:r>)[\s\S])*?<w:pict[\s\S]*?<\/w:pict><\/w:r>/g)].map((m) => m[0]);

describe('浮水印: the XML Word writes', () => {
  it('a preset goes into every header part the sections use, a header made where a section needs one', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    expect(ed.watermark()).toBeNull();
    expect(ed.setWatermark(textWatermark('機密', '標楷體'))).toBe(true);
    const zip = await saved(ed);
    const [first, second] = await sectionHeaders(zip);
    // Section 1 uses 「不同的首頁」 and had no first-page header: one is made for it.
    expect(Object.keys(first).sort()).toEqual(['default', 'first']);
    expect(first.default).toBe('word/header1.xml');
    // Section 2 links to section 1's headers and gets no references of its own; no even-page header.
    expect(second).toEqual({});
    for (const part of [first.default, first.first]) {
      const xml = await text(zip, part);
      const runs = pictRuns(xml);
      expect(runs, part).toHaveLength(1);
      expect(runs[0]).toMatch(OUR_TEXT_RUN({ text: '機密', font: '標楷體' }));
      // The shape type is defined once in the part.
      expect(xml.split('id="_x0000_t136"')).toHaveLength(2);
    }
    // The header's own text stays, after the watermark in its first paragraph (as Word puts it).
    const header1 = await text(zip, 'word/header1.xml');
    expect(header1).toMatch(/<w:p>.*<w:pict>.*<\/w:pict><\/w:r><w:r><w:t>甲<\/w:t><\/w:r><\/w:p>/);
    // Distinct shape ids across the parts, and each shape its own z-index, 1024 apart (as Word).
    const both = header1 + (await text(zip, first.first));
    const ids = [...both.matchAll(/o:spid="(_x0000_s\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(2);
    const z = [...both.matchAll(/z-index:(-\d+)/g)].map((m) => Number(m[1])).sort((a, b) => b - a);
    expect(z).toEqual([-251657216, -251658240]);
    expect(ed.watermark()).toMatchObject({ kind: 'text', text: '機密', font: '標楷體', size: null, color: '#c0c0c0', semitransparent: true, layout: 'diagonal' });
  });

  it('an even-page header is made only when the document uses odd and even headers', async () => {
    const zip = await pkg(`<w:p/><w:sectPr>${SECT}</w:sectPr>`);
    zip.file('word/settings.xml', `<w:settings ${W}><w:evenAndOddHeaders/></w:settings>`);
    const { ed } = await open(zip);
    ed.setWatermark(textWatermark('草稿', '標楷體'));
    const [only] = await sectionHeaders(await saved(ed));
    expect(Object.keys(only).sort()).toEqual(['default', 'even']);
  });

  it('declares v, o, w10 and r on the header root, in a new part and in the file’s own', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:hdr>` }));
    ed.setWatermark(textWatermark('機密', '標楷體'));
    const zip = await saved(ed);
    const [first] = await sectionHeaders(zip);
    for (const part of [first.default, first.first]) {
      const root = /<w:hdr\b[^>]*>/.exec(await text(zip, part))![0];
      expect(root, part).toContain('xmlns:v="urn:schemas-microsoft-com:vml"');
      expect(root, part).toContain('xmlns:o="urn:schemas-microsoft-com:office:office"');
      expect(root, part).toContain('xmlns:w10="urn:schemas-microsoft-com:office:word"');
      expect(root, part).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    }
  });

  it('replacing takes out the old shapes (Word’s and ours); removing leaves the other header content', async () => {
    const { ed } = await open(await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: WORD_HEADER }));
    ed.setWatermark(textWatermark('草稿', '標楷體'));
    let xml = await text(await saved(ed), 'word/header1.xml');
    expect(pictRuns(xml)).toHaveLength(1);
    expect(xml).not.toContain('CONFIDENTIAL');
    expect(xml).toContain('string="草稿"');
    expect(xml).toContain('<w:t>頁首文字</w:t>');
    ed.setWatermark(textWatermark('樣本', '標楷體'));
    xml = await text(await saved(ed), 'word/header1.xml');
    expect(pictRuns(xml)).toHaveLength(1);
    expect(pictRuns(xml)[0]).toMatch(OUR_TEXT_RUN({ text: '樣本', font: '標楷體' }));

    expect(ed.setWatermark(null)).toBe(true);
    xml = await text(await saved(ed), 'word/header1.xml');
    expect(pictRuns(xml)).toHaveLength(0);
    expect(xml).not.toContain('PowerPlusWaterMarkObject');
    expect(xml).toContain('<w:t>頁首文字</w:t>');
    expect(xml).toContain('<w:pStyle w:val="a3"/>');
    expect(ed.watermark()).toBeNull();
    // Nothing left to remove: no change.
    expect(ed.setWatermark(null)).toBe(false);
  });

  it('a picture watermark: Word’s washed-out picture shape, the picture as a media part of each header', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    const picture: Watermark = { kind: 'picture', src: PNG_URL, width: 40, height: 20, scale: null, washout: true };
    expect(ed.setWatermark(picture)).toBe(true);
    const zip = await saved(ed);
    const [first] = await sectionHeaders(zip);
    for (const part of [first.default, first.first]) {
      const xml = await text(zip, part);
      const [run] = pictRuns(xml);
      expect(run).toContain('<v:shapetype id="_x0000_t75"');
      expect(run).toMatch(/<v:shape id="WordPictureWatermark\d+" o:spid="_x0000_s\d+" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:[\d.]+pt;height:[\d.]+pt;z-index:-\d+;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f">/);
      const id = /<v:imagedata r:id="([^"]+)" o:title="" gain="19661f" blacklevel="22938f"\/>/.exec(run)![1];
      const rels = await text(zip, part.replace('word/', 'word/_rels/') + '.rels');
      const rel = new RegExp(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="([^"]+)"/>`).exec(rels);
      expect(rel, part).not.toBeNull();
      const media = zip.file('word/' + rel![1])!;
      expect(media).toBeTruthy();
      expect(Buffer.from(await media.async('uint8array')).toString('base64')).toBe(PNG);
    }
    expect(await text(zip, '[Content_Types].xml')).toMatch(/Extension="png"/);
    expect(ed.watermark()).toMatchObject({ kind: 'picture', src: PNG_URL, washout: true });
    // Replaced by a text watermark: the picture's relationship goes with it.
    ed.setWatermark(textWatermark('機密', '標楷體'));
    const again = await saved(ed);
    expect(await text(again, 'word/_rels/header1.xml.rels')).not.toContain('relationships/image');
  });

  it('refused while 追蹤修訂 is on, with a notice; read-only: not at all', async () => {
    const notices: string[] = [];
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }), { onNotice: (m) => notices.push(m) });
    ed.setTrackChanges(true);
    expect(ed.setWatermark(textWatermark('機密', '標楷體'))).toBe(false);
    expect(notices.at(-1)).toMatch(/追蹤修訂.*浮水印|浮水印.*追蹤修訂/);
    expect(ed.watermark()).toBeNull();
    const ro = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }), { editable: false });
    expect(ro.ed.setWatermark(textWatermark('機密', '標楷體'))).toBe(false);
    expect(ro.ed.watermark()).toBeNull();
  });

  it('is not an undo step of the body: Ctrl+Z there leaves it (the menu removes or replaces it)', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    const view = ed.view!;
    view.dispatch(view.state.tr.insertText('字', 2));
    ed.setWatermark(textWatermark('機密', '標楷體'));
    undo(view.state, view.dispatch);
    expect(view.state.doc.textContent).toBe('第一節第二節');
    expect(ed.watermark()).toMatchObject({ text: '機密' });
    const [first] = await sectionHeaders(await saved(ed));
    expect(Object.keys(first).sort()).toEqual(['default', 'first']);
  });
});

describe('浮水印 made in Word', () => {
  const wordDoc = () => pkg(`<w:p><w:r><w:t>內文</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: WORD_HEADER });

  it('is recognised (text, font, colour, transparency, rotation)', async () => {
    const { ed } = await open(await wordDoc());
    expect(ed.watermark()).toEqual({ kind: 'text', text: 'CONFIDENTIAL', font: 'Calibri', size: null, color: '#c0c0c0', semitransparent: true, layout: 'diagonal' });
  });

  it('is drawn on a layer of each page, centred in the margins, rotated and see-through, behind the text', async () => {
    const { ed, host } = await open(await wordDoc());
    layOut(ed, [{ section: 0, inSection: 0 }, { section: 0, inSection: 1 }]);
    const pages = Array.from(host.querySelectorAll<HTMLElement>('.dx-page'));
    expect(pages).toHaveLength(2);
    for (const page of pages) {
      const layer = page.querySelector<HTMLElement>(':scope > .dx-wm-layer')!;
      expect(layer).not.toBeNull();
      expect(layer.getAttribute('aria-hidden')).toBe('true');
      // Not inside the (clipped) header band, and before it: behind the header and body text.
      expect(layer.closest('.dx-header')).toBeNull();
      expect(page.firstElementChild).toBe(layer);
      const shape = layer.querySelector<HTMLElement>('.dx-wm')!;
      expect(shape.style.transform).toBe('rotate(315deg)');
      // 527.85 pt × 131.95 pt, centred on the text area (A4, 1800 twips left and right margins).
      const w = (527.85 * 96) / 72;
      const hgt = (131.95 * 96) / 72;
      expect(parseFloat(shape.style.width)).toBeCloseTo(w, 1);
      expect(parseFloat(shape.style.height)).toBeCloseTo(hgt, 1);
      expect(parseFloat(shape.style.left) + w / 2).toBeCloseTo(120 + (11906 / 15 - 240) / 2, 1);
      expect(parseFloat(shape.style.top) + hgt / 2).toBeCloseTo(96 + (16838 / 15 - 192) / 2, 1);
      const t = shape.querySelector('text')!;
      expect(t.textContent).toBe('CONFIDENTIAL');
      expect(t.getAttribute('fill')).toBe('#c0c0c0');
      expect(t.getAttribute('fill-opacity')).toBe('0.5');
      expect(t.getAttribute('font-family')).toContain('Calibri');
    }
    // The static header shows its text but not the shape (its chip takes no room there).
    const header = pages[0].querySelector('.dx-header')!;
    expect(header.textContent).toContain('頁首文字');
    expect(header.querySelector('.dx-wm-chip')).not.toBeNull();
  });

  it('is in the printout (printHtml), as plain markup', async () => {
    const { ed } = await open(await wordDoc());
    layOut(ed, [{ section: 0, inSection: 0 }]);
    const html = ed.printHtml('t')!;
    const body = html.slice(html.indexOf('<body'));
    expect(body).toContain('dx-wm-layer');
    expect(body).toMatch(/<text[^>]*>CONFIDENTIAL<\/text>/);
    expect(body).not.toMatch(/<script/i);
  });

  it('survives an unrelated edit byte for byte (the body, and typing in the same header)', async () => {
    const { ed, bytes } = await open(await wordDoc());
    const original = await text(await JSZip.loadAsync(bytes), 'word/header1.xml');
    const view = ed.view!;
    view.dispatch(view.state.tr.insertText('新增', 1));
    let zip = await saved(ed);
    expect(await text(zip, 'word/header1.xml')).toBe(original);
    // Typed into the header: the watermark's run is written back as it was.
    ed.editHeaderFooter('header', 0);
    const hv = ed.activeView!;
    hv.dispatch(hv.state.tr.insertText('又', hv.state.doc.content.size - 1));
    ed.closeHeaderFooter();
    zip = await saved(ed);
    const xml = await text(zip, 'word/header1.xml');
    expect(xml).not.toBe(original);
    expect(pictRuns(xml)).toEqual([WORD_RUN]);
  });

  it('typing over the shape in the header keeps it (it is removed with 移除浮水印)', async () => {
    const notices: string[] = [];
    const { ed } = await open(await wordDoc(), { onNotice: (m) => notices.push(m) });
    ed.editHeaderFooter('header', 0);
    const hv = ed.activeView!;
    // Select everything in the header and type over it.
    hv.dispatch(hv.state.tr.setSelection(new AllSelection(hv.state.doc)));
    hv.dispatch(hv.state.tr.replaceSelectionWith(schema.text('新頁首')));
    expect(hv.state.doc.textContent).toContain('新頁首');
    expect(hv.state.doc.textContent).not.toContain('頁首文字');
    expect(ed.watermark()).toMatchObject({ text: 'CONFIDENTIAL' });
    expect(notices.at(-1)).toContain('移除浮水印');
    // The header editor shows it as a chip named 浮水印.
    expect(hv.dom.querySelector('.dx-wm-chip')?.textContent).toBe('浮水印');
    ed.closeHeaderFooter();
    expect(pictRuns(await text(await saved(ed), 'word/header1.xml'))).toEqual([WORD_RUN]);
  });

  it('per section: a first-page header without one shows none', async () => {
    const zip = await pkg(
      `<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/><w:headerReference w:type="first" r:id="rId_header2"/>${SECT}<w:titlePg/></w:sectPr>`,
      { header1: WORD_HEADER, header2: PLAIN_HEADER },
    );
    const { ed, host } = await open(zip);
    layOut(ed, [{ section: 0, inSection: 0 }, { section: 0, inSection: 1 }]);
    const [p1, p2] = Array.from(host.querySelectorAll('.dx-page'));
    expect(p1.querySelector('.dx-wm-layer')).toBeNull();
    expect(p2.querySelector('.dx-wm-layer')).not.toBeNull();
  });

  it('the compatibility notice no longer lists it as an unsupported shape', async () => {
    const report = await scanCompat(await wordDoc());
    expect(report.items.map((i) => i.id)).not.toContain('shape');
  });
});

describe('浮水印: review follow-ups', () => {
  const letter = (headers: Record<string, string> = {}) =>
    pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`, { header1: PLAIN_HEADER, ...headers });

  it('the text box has the text’s own proportions (measured in its font) and, turned, spans the text width', () => {
    // Calibri as the browser measures it at 100 px: CONFIDENTIAL 603.08 px wide, font box 95 + 27 px.
    const box = textWatermarkSize(textWatermark('CONFIDENTIAL', 'Calibri'), { width: 468, height: 648 }, { width: 6.0308, height: 1.22 });
    expect((box.width + box.height) / Math.SQRT2).toBeCloseTo(468, 0);
    // Word writes 527.85 pt × 131.95 pt for it on Letter.
    expect(Math.abs(box.width - 527.85)).toBeLessThan(30);
    expect(Math.abs(box.height - 131.95)).toBeLessThan(25);
    expect(box.width / box.height).toBeCloseTo(6.0308 / 1.22, 2);
  });

  it('measures the text with a canvas when there is one, and draws it with the same geometry', async () => {
    // The browser's OffscreenCanvas (jsdom has none: the text is then estimated).
    vi.stubGlobal('OffscreenCanvas', class {
      getContext() {
        return { font: '', measureText: (t: string) => ({ width: t === 'CONFIDENTIAL' ? 603.08 : 100 * t.length, fontBoundingBoxAscent: 95, fontBoundingBoxDescent: 27 }) };
      }
    });
    const { ed, host } = await open(await letter());
    ed.setWatermark(textWatermark('CONFIDENTIAL', 'Calibri'));
    const [run] = pictRuns(await text(await saved(ed), 'word/header1.xml'));
    const [, w, h] = /width:([\d.]+)pt;height:([\d.]+)pt/.exec(run)!.map(Number);
    expect((w + h) / Math.SQRT2).toBeCloseTo(468, 0);
    expect(w / h).toBeCloseTo(603.08 / 122, 1);
    layOut(ed, [{ section: 0, inSection: 0 }]);
    const shape = host.querySelector<HTMLElement>('.dx-wm')!;
    const t = shape.querySelector('text')!;
    const heightPx = (h * 96) / 72;
    // The font box fills the shape's height, the text its width.
    expect(Number(t.getAttribute('font-size'))).toBeCloseTo(heightPx / 1.22, 0);
    expect(Number(t.getAttribute('y'))).toBeCloseTo((heightPx / 1.22) * 0.95, 0);
    expect(Number(t.getAttribute('textLength'))).toBeCloseTo((w * 96) / 72, 0);
  });

  /** Word's 浮水印 gallery (機密 1): the watermark in a building-block content control of its own. */
  const GALLERY_SDT =
    '<w:sdt><w:sdtPr><w:id w:val="-1582987281"/><w:docPartObj><w:docPartGallery w:val="Watermarks"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtEndPr/>' +
    `<w:sdtContent><w:p w:rsidR="00F10A2B" w:rsidRDefault="00F10A2B"><w:pPr><w:pStyle w:val="a3"/></w:pPr>${WORD_RUN}</w:p></w:sdtContent></w:sdt>`;

  it('a gallery watermark’s content control goes when its paragraph is left empty; the header keeps a paragraph', async () => {
    const withText = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${HDR_ROOT}${GALLERY_SDT}<w:p><w:pPr><w:pStyle w:val="a3"/></w:pPr><w:r><w:t>頁首文字</w:t></w:r></w:p></w:hdr>`;
    const alone = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${HDR_ROOT}${GALLERY_SDT}</w:hdr>`;
    for (const header of [withText, alone]) {
      const { ed } = await open(await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: header }));
      expect(ed.watermark()).toMatchObject({ text: 'CONFIDENTIAL' });
      expect(ed.setWatermark(null)).toBe(true);
      const xml = await text(await saved(ed), 'word/header1.xml');
      expect(xml).not.toContain('<w:sdt');
      expect(xml).not.toContain('Watermarks');
      expect(xml).not.toContain('<w:pict');
      expect(xml).toMatch(/<w:p[ >/]/);
      if (header === withText) expect(xml).toContain('<w:t>頁首文字</w:t>');
    }
    // Replaced: the old control goes, the new watermark is in the header's first paragraph.
    const { ed } = await open(await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: withText }));
    ed.setWatermark(textWatermark('草稿'));
    const xml = await text(await saved(ed), 'word/header1.xml');
    expect(xml).not.toContain('<w:sdt');
    expect(pictRuns(xml)).toHaveLength(1);
    expect(xml).toMatch(/<w:p>.*string="草稿".*<w:t>頁首文字<\/w:t>/);
  });

  it('only the header parts the sections refer to are touched (a footer is left alone)', async () => {
    const zip = await pkg(
      `<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/><w:footerReference w:type="default" r:id="rId_footer1"/>${SECT}</w:sectPr>`,
      { header1: WORD_HEADER },
    );
    const footer = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${HDR_ROOT.replace('<w:hdr', '<w:ftr')}<w:p>${WORD_RUN}</w:p></w:ftr>`;
    zip.file('word/footer1.xml', footer);
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rId_footer1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'));
    const { ed } = await open(zip);
    ed.setWatermark(null);
    let out = await saved(ed);
    expect(await text(out, 'word/footer1.xml')).toBe(footer);
    expect(pictRuns(await text(out, 'word/header1.xml'))).toHaveLength(0);
    ed.setWatermark(textWatermark('機密'));
    out = await saved(ed);
    expect(await text(out, 'word/footer1.xml')).toBe(footer);
  });

  it('a picture watermark made in Word is recognised (自動 scale, 刷淡) and drawn washed out', async () => {
    const run =
      '<w:r><w:rPr><w:noProof/></w:rPr><w:pict w14:anchorId="7A1B2C3D"><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>' +
      '<v:shape id="WordPictureWatermark101447171" o:spid="_x0000_s2050" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:415.25pt;height:207.6pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f">' +
      '<v:imagedata r:id="rId1" o:title="seal" gain="19661f" blacklevel="22938f"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>';
    const zip = await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, {
      header1: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${HDR_ROOT}<w:p><w:pPr><w:pStyle w:val="a3"/></w:pPr>${run}</w:p></w:hdr>`,
    });
    zip.file('word/_rels/header1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>');
    zip.file('word/media/image1.png', Uint8Array.from(atob(PNG), (c) => c.charCodeAt(0)));
    const { ed, host } = await open(zip);
    // 415.25 pt is the text width (A4, 3.17 cm margins): Word's 自動.
    expect(ed.watermark()).toMatchObject({ kind: 'picture', src: PNG_URL, scale: null, washout: true });
    layOut(ed, [{ section: 0, inSection: 0 }]);
    const img = host.querySelector<HTMLImageElement>('.dx-page > .dx-wm-layer .dx-wm img')!;
    expect(img.getAttribute('src')).toBe(PNG_URL);
    expect(img.classList.contains('dx-wm-washout')).toBe(true);
    expect(parseFloat(img.parentElement!.style.width)).toBeCloseTo((415.25 * 96) / 72, 1);
    expect(host.querySelector('.dx-header')!.querySelector('img')).toBeNull();
  });

  it('a picture watermark set to 自動 reads back as 自動; o:title is the picture’s name, as Word writes it', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    ed.setWatermark({ kind: 'picture', src: PNG_URL, width: 40, height: 20, scale: null, washout: false, title: '機關"印信' });
    expect(ed.watermark()).toMatchObject({ kind: 'picture', scale: null, washout: false });
    const [run] = pictRuns(await text(await saved(ed), 'word/header1.xml'));
    expect(run).toMatch(/<v:imagedata r:id="rIdPx\d+" o:title="機關&quot;印信"\/>/);
  });

  it('pasting into the header drops a copied watermark (no second shape with the same id)', async () => {
    const { ed } = await open(await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: WORD_HEADER }));
    ed.editHeaderFooter('header', 0);
    const hv = ed.activeView!;
    let chip: any = null;
    hv.state.doc.descendants((n) => {
      if (n.type.name === 'raw_inline') chip = n;
    });
    let slice = new Slice(Fragment.from(schema.nodes.paragraph.create(null, [chip, schema.text('貼上')])), 1, 1);
    hv.someProp('transformPasted', (f: any) => {
      slice = f(slice, hv, false);
      return true;
    });
    let chips = 0;
    slice.content.descendants((n) => {
      if (n.type.name === 'raw_inline') chips++;
    });
    expect(chips).toBe(0);
    expect(slice.content.textBetween(0, slice.content.size)).toBe('貼上');
  });
});

describe('浮水印: code review follow-ups', () => {
  const withWord = () => pkg(`<w:p><w:r><w:t>內文</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: WORD_HEADER });
  /** The watermark shapes (runs, not tracked deletions) of the header being edited. */
  const shapesIn = (doc: any) => {
    const out: string[] = [];
    doc.descendants((n: any) => {
      if (n.type.name === 'raw_inline' && isWatermarkRun(n.attrs.xml)) out.push(n.attrs.xml);
    });
    return out;
  };

  it('with 追蹤修訂 on, deleting everything in the header leaves the watermark as it is, not a tracked deletion', async () => {
    const { ed } = await open(await withWord());
    ed.setTrackChanges(true);
    ed.editHeaderFooter('header', 0);
    const hv = ed.activeView!;
    hv.dispatch(hv.state.tr.setSelection(new AllSelection(hv.state.doc)));
    hv.dispatch(hv.state.tr.deleteSelection());
    expect(shapesIn(hv.state.doc)).toEqual([WORD_RUN]);
    let deletedShapes = 0;
    hv.state.doc.descendants((n) => {
      if (n.type.name === 'raw_inline' && n.attrs.label === 'del' && /PowerPlusWaterMarkObject/.test(n.attrs.xml)) deletedShapes++;
    });
    expect(deletedShapes).toBe(0);
    // Rejecting the deletion leaves one watermark too.
    rejectAllRevisions(hv.state, hv.dispatch);
    expect(shapesIn(hv.state.doc)).toEqual([WORD_RUN]);
    ed.closeHeaderFooter();
    const xml = await text(await saved(ed), 'word/header1.xml');
    expect(xml.match(/PowerPlusWaterMarkObject/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<w:del\b(?:(?!<\/w:del>)[\s\S])*PowerPlusWaterMarkObject/);
  });

  it('Ctrl+Z in the header after 浮水印 changed it brings no old watermark back', async () => {
    const { ed } = await open(await withWord());
    ed.editHeaderFooter('header', 0);
    let hv = ed.activeView!;
    // One undo step: typing over everything, the watermark put back (keepWatermarks).
    hv.dispatch(hv.state.tr.setSelection(new AllSelection(hv.state.doc)));
    hv.dispatch(hv.state.tr.replaceSelectionWith(schema.text('X')));
    ed.closeHeaderFooter();
    ed.setWatermark(textWatermark('草稿'));
    ed.editHeaderFooter('header', 0);
    hv = ed.activeView!;
    undo(hv.state, hv.dispatch);
    const shapes = shapesIn(hv.state.doc);
    expect(shapes).toHaveLength(1);
    expect(shapes[0]).toContain('string="草稿"');
    expect(hv.state.doc.textContent).toContain('頁首文字');
  });

  it('setting a watermark while the header is open: the open header shows it; Ctrl+Z there neither removes nor restores one', async () => {
    const { ed } = await open(await withWord());
    ed.editHeaderFooter('header', 0);
    let hv = ed.activeView!;
    hv.dispatch(hv.state.tr.insertText('又', hv.state.doc.content.size - 1));
    ed.setWatermark(textWatermark('草稿'));
    expect(ed.target).toBe('header');
    hv = ed.activeView!;
    expect(shapesIn(hv.state.doc).map((x) => /string="([^"]+)"/.exec(x)![1])).toEqual(['草稿']);
    undo(hv.state, hv.dispatch);
    expect(hv.state.doc.textContent).not.toContain('又');
    expect(shapesIn(hv.state.doc).map((x) => /string="([^"]+)"/.exec(x)![1])).toEqual(['草稿']);
    undo(hv.state, hv.dispatch);
    expect(shapesIn(hv.state.doc).map((x) => /string="([^"]+)"/.exec(x)![1])).toEqual(['草稿']);
  });

  it('Ctrl+Z of an earlier page setup change keeps the header 浮水印 made (it is still referenced and saved)', async () => {
    const { ed } = await open(await pkg(`<w:p><w:r><w:t>內文</w:t></w:r></w:p><w:sectPr>${SECT}</w:sectPr>`));
    const view = ed.view!;
    ed.setPageSetup({ ...ed.cursorSection().page, marginLeft: 1000, marginRight: 1000 });
    ed.setWatermark(textWatermark('機密'));
    undo(view.state, view.dispatch);
    expect(ed.cursorSection().page.marginLeft).toBe(1800);
    expect(ed.watermark()).toMatchObject({ text: '機密' });
    const zip = await saved(ed);
    const [only] = await sectionHeaders(zip);
    expect(Object.keys(only)).toEqual(['default']);
    expect(await text(zip, only.default)).toContain('string="機密"');
    expect(await text(zip, 'word/document.xml')).toContain('w:left="1800"');
  });

  it('a watermark copied from the header is not pasted into the text', async () => {
    const { ed } = await open(await withWord());
    const chip = ed.model.headerFooters[0].doc.firstChild!.firstChild!;
    expect(isWatermarkRun(chip.attrs.xml)).toBe(true);
    let slice = new Slice(Fragment.from(schema.nodes.paragraph.create(null, [chip, schema.text('貼上')])), 1, 1);
    ed.view!.someProp('transformPasted', (f: any) => {
      slice = f(slice, ed.view, false);
      return true;
    });
    expect(slice.content.textBetween(0, slice.content.size)).toBe('貼上');
    let chips = 0;
    slice.content.descendants((n) => {
      if (n.type.name === 'raw_inline') chips++;
    });
    expect(chips).toBe(0);
  });

  it('a picture watermark in several header parts is one media file, each part with its own relationship', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    ed.setWatermark({ kind: 'picture', src: PNG_URL, width: 40, height: 20, scale: null, washout: true });
    const zip = await saved(ed);
    const [first] = await sectionHeaders(zip);
    const targets = new Set<string>();
    for (const part of [first.default, first.first]) {
      const rels = await text(zip, part.replace('word/', 'word/_rels/') + '.rels');
      const t = /relationships\/image" Target="([^"]+)"/.exec(rels)![1];
      targets.add(t);
    }
    expect(targets.size).toBe(1);
    expect(Object.keys(zip.files).filter((f) => f.startsWith('word/media/') && !zip.files[f].dir)).toHaveLength(1);
  });

  it('a watermark is read once per XML, however often the header changes', () => {
    const shape = readWatermark(WORD_RUN);
    expect(shape).not.toBeNull();
    expect(readWatermark(WORD_RUN)).toBe(shape);
    expect(Object.isFrozen(shape)).toBe(true);
  });

  it('watermarkTargets: every header part the sections show, a header to make where a section shows none', async () => {
    // Section 1 shows no header; section 2 has its own.
    const zip = await pkg(
      `<w:p><w:pPr><w:sectPr>${SECT}</w:sectPr></w:pPr></w:p><w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}<w:titlePg/></w:sectPr>`,
      { header1: PLAIN_HEADER },
    );
    const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    const sections = documentSections(doc, model);
    const plan = watermarkTargets(model, sections);
    expect([...plan.parts].map(([hf, s]) => [hf.relId, s.index])).toEqual([['rId_header1', 1]]);
    expect(plan.missing.map((m) => [m.section.index, m.type])).toEqual([[0, 'default'], [1, 'first']]);
  });

  it('a later section with its own header, an earlier one without: a header is made for the earlier one', async () => {
    const zip = await pkg(
      `<w:p><w:pPr><w:sectPr>${SECT}</w:sectPr></w:pPr><w:r><w:t>一</w:t></w:r></w:p><w:p><w:r><w:t>二</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`,
      { header1: PLAIN_HEADER },
    );
    const { ed } = await open(zip);
    ed.setWatermark(textWatermark('機密'));
    const out = await saved(ed);
    const [first, second] = await sectionHeaders(out);
    expect(second).toEqual({ default: 'word/header1.xml' });
    expect(Object.keys(first)).toEqual(['default']);
    expect(first.default).not.toBe('word/header1.xml');
    for (const part of [first.default, second.default]) expect(await text(out, part)).toContain('string="機密"');
    expect(await text(out, 'word/header1.xml')).toContain('<w:t>甲</w:t>');
  });

  it('the shape box is at least 1 pt and never NaN', () => {
    const xml = watermarkRunXml({ ...textWatermark('A'), size: 0.01 }, { width: NaN, height: NaN }, { number: 1, spid: 2049, shapetype: true, zIndex: NaN });
    expect(xml).not.toContain('NaN');
    const [, w, h] = /width:([\d.]+)pt;height:([\d.]+)pt/.exec(xml)!.map(Number);
    expect(w).toBeGreaterThanOrEqual(1);
    expect(h).toBeGreaterThanOrEqual(1);
    expect(xml).toContain('z-index:-251657216');
  });

  it('escapes & < > " \' and drops control characters in the text, the font and the picture name (XML and SVG)', async () => {
    const nasty = 'A&B<C>"D\'E\u0001F';
    const { ed, host } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    ed.setWatermark({ ...textWatermark(nasty), font: 'Ca"l;i<b>ri' });
    let xml = await text(await saved(ed), 'word/header1.xml');
    // Well-formed, and the text reads back as it was (without the control character).
    const parsed = new DOMParser().parseFromString(xml, 'application/xml');
    expect(parsed.getElementsByTagName('parsererror')).toHaveLength(0);
    const textpath = Array.from(parsed.getElementsByTagNameNS('urn:schemas-microsoft-com:vml', 'textpath')).find((e) => e.hasAttribute('string'))!;
    expect(textpath.getAttribute('string')).toBe('A&B<C>"D\'EF');
    expect(textpath.getAttribute('style')).toBe('font-family:"Calibri";font-size:1pt');
    layOut(ed, [{ section: 0, inSection: 0 }]);
    const t = host.querySelector('.dx-wm text')!;
    expect(t.textContent).toBe('A&B<C>"D\'EF');
    expect(t.getAttribute('font-family')).not.toMatch(/[<;]/);
    ed.setWatermark({ kind: 'picture', src: PNG_URL, width: 40, height: 20, scale: null, washout: true, title: nasty });
    xml = await text(await saved(ed), 'word/header1.xml');
    console.log('DBG', /<v:imagedata[^>]*>/.exec(xml)?.[0]);
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    expect(doc.getElementsByTagName('parsererror')).toHaveLength(0);
    expect(doc.getElementsByTagNameNS('urn:schemas-microsoft-com:vml', 'imagedata')[0].getAttribute('o:title')).toBe('A&B<C>"D\'EF');
  });

  it('a watermark whose picture is missing can still be removed (移除浮水印 and 無浮水印)', async () => {
    const run = WORD_RUN.replace(/<v:shape id="PowerPlusWaterMarkObject357831064"[\s\S]*<\/v:shape>/,
      '<v:shape id="WordPictureWatermark5" o:spid="_x0000_s2050" type="#_x0000_t75" style="position:absolute;width:400pt;height:200pt" o:allowincell="f"><v:imagedata r:id="rIdGone" o:title=""/></v:shape>');
    const header = WORD_HEADER.replace(WORD_RUN, run);
    const { ed } = await open(await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: header }));
    expect(ed.watermark()).toBeNull();
    expect(ed.hasWatermark()).toBe(true);
    const d = await mountDialog(ed);
    expect(d.$<HTMLInputElement>('input[name="dx-wm-kind"][value="none"]').checked).toBe(true);
    d.$<HTMLButtonElement>('.dx-primary').click();
    await d.flush();
    d.done();
    expect(ed.hasWatermark()).toBe(false);
    expect(await text(await saved(ed), 'word/header1.xml')).not.toContain('WordPictureWatermark');
  });

  it('the text is drawn again with its font’s measurements once the font has loaded', async () => {
    const { ed, host } = await open(await withWord());
    layOut(ed, [{ section: 0, inSection: 0 }]);
    const before = host.querySelector('.dx-wm text')!.getAttribute('font-size');
    vi.stubGlobal('OffscreenCanvas', class {
      getContext() {
        return { font: '', measureText: () => ({ width: 603.08, fontBoundingBoxAscent: 95, fontBoundingBoxDescent: 27 }) };
      }
    });
    (ed as any).onFontsLoaded();
    const after = host.querySelector('.dx-wm text')!.getAttribute('font-size');
    expect(after).not.toBe(before);
    expect(Number(after)).toBeCloseTo(((131.95 * 96) / 72) / 1.22, 0);
  });

  it('without a canvas to measure with (jsdom), no canvas is created', async () => {
    const spy = vi.spyOn(HTMLCanvasElement.prototype, 'getContext');
    expect(measureWatermarkText('機密', '標楷體')).toBeNull();
    expect(spy).not.toHaveBeenCalled();
  });

  it('a file read that ends after the dialog was opened again does not fill the new form', async () => {
    let finish: () => void = () => {};
    vi.stubGlobal('Image', class {
      naturalWidth = 40;
      naturalHeight = 20;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      set src(_v: string) {
        finish = () => this.onload?.();
      }
    });
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    const d = await mountDialog(ed);
    await d.set('input[name="dx-wm-kind"][value="picture"]', true);
    const input = d.$<HTMLInputElement>('#dx-wm-file');
    Object.defineProperty(input, 'files', { value: [new File([Uint8Array.from(atob(PNG), (c) => c.charCodeAt(0))], 'late.png', { type: 'image/png' })], configurable: true });
    input.dispatchEvent(new Event('change'));
    for (let i = 0; i < 3; i++) await d.flush();
    // Closed and opened again before the picture has loaded.
    d.open.value = false;
    await d.flush();
    d.open.value = true;
    await d.flush();
    finish();
    for (let i = 0; i < 3; i++) await d.flush();
    expect(d.$<HTMLInputElement>('input[name="dx-wm-kind"][value="none"]').checked).toBe(true);
    expect(d.host.textContent).not.toContain('late.png');
    d.done();
  });

  it('two dialogs on the page have their own ids and radio groups', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    const a = await mountDialog(ed);
    const b = await mountDialog(ed);
    const idA = a.$<HTMLInputElement>('#dx-wm-text').id;
    const idB = b.$<HTMLInputElement>('#dx-wm-text').id;
    expect(idA).not.toBe(idB);
    expect(a.$<HTMLInputElement>('input[name="dx-wm-kind"]').name).not.toBe(b.$<HTMLInputElement>('input[name="dx-wm-kind"]').name);
    a.done();
    b.done();
  });

  it('when no header can hold it (the header is unreadable), a notice says so', async () => {
    const notices: string[] = [];
    const zip = await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: 'not xml <' });
    const { ed } = await open(zip, { onNotice: (m) => notices.push(m) });
    expect(ed.setWatermark(textWatermark('機密'))).toBe(false);
    expect(notices.at(-1)).toContain('浮水印');
    expect(notices.at(-1)).toContain('頁首');
  });

  it('removing keeps the shape type another shape in the header still uses', async () => {
    const other = '<w:r><w:pict><v:shape id="WordArt7" o:spid="_x0000_s2060" type="#_x0000_t136" style="position:absolute;width:100pt;height:40pt" fillcolor="black" stroked="f"><v:textpath style="font-family:&quot;Arial&quot;" string="LOGO"/></v:shape></w:pict></w:r>';
    const header = WORD_HEADER.replace('<w:r><w:t>頁首文字</w:t></w:r>', `<w:r><w:t>頁首文字</w:t></w:r>${other}`);
    const { ed } = await open(await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: header }));
    ed.setWatermark(null);
    const xml = await text(await saved(ed), 'word/header1.xml');
    expect(xml).not.toContain('PowerPlusWaterMarkObject');
    expect(xml.match(/<v:shapetype id="_x0000_t136"/g)).toHaveLength(1);
    expect(xml).toContain('string="LOGO"');
    // With nothing using it, the shape type goes with the watermark.
    const plain = await open(await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: WORD_HEADER }));
    plain.ed.setWatermark(null);
    expect(await text(await saved(plain.ed), 'word/header1.xml')).not.toContain('_x0000_t136');
  });
});

describe('浮水印: section references through body Ctrl+Z (re-review)', () => {
  /** Puts the body's cursor in the n-th (0-based) top-level paragraph. */
  const cursorIn = (ed: DocxEditor, n: number) => {
    const view = ed.view!;
    let at = 1;
    view.state.doc.forEach((_node, offset, index) => {
      if (index === n) at = offset + 1;
    });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
  };
  /** One page per section (jsdom lays nothing out). */
  const pagesFor = (ed: DocxEditor, count: number) => {
    const page = (ed as any).firstPage();
    (ed as any).pages = Array.from({ length: count }, (_, i) => ({ ...page, section: i, inSection: 0, number: i + 1, top: i * (page.height + 24) }));
  };
  /** Every w:headerReference / w:footerReference of the saved document points at a relationship it has. */
  async function referencesResolve(zip: JSZip) {
    const document = await text(zip, 'word/document.xml');
    const rels = await text(zip, 'word/_rels/document.xml.rels');
    for (const m of document.matchAll(/<w:(?:header|footer)Reference\b[^>]*r:id="([^"]+)"/g)) {
      expect(rels, `relationship ${m[1]}`).toContain(`Id="${m[1]}"`);
    }
  }

  it('a page setup change on a section that is not the last, undone, keeps the header 浮水印 made for it', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    const view = ed.view!;
    cursorIn(ed, 0);
    expect(ed.cursorSection().index).toBe(0);
    ed.setPageSetup({ ...ed.cursorSection().page, marginLeft: 1000, marginRight: 1000 });
    ed.setWatermark(textWatermark('機密'));
    undo(view.state, view.dispatch);
    cursorIn(ed, 0);
    expect(ed.cursorSection().page.marginLeft).toBe(1800);
    const zip = await saved(ed);
    const [first] = await sectionHeaders(zip);
    expect(Object.keys(first).sort()).toEqual(['default', 'first']);
    expect(await text(zip, first.first)).toContain('string="機密"');
    await referencesResolve(zip);
  });

  it('連結到前一節 on a middle section, then Ctrl+Z of its page setup: no reference to a header that is gone', async () => {
    const sect = (inner = '') => `<w:sectPr>${inner}${SECT}</w:sectPr>`;
    const zip = await pkg(
      `<w:p><w:pPr>${sect('<w:headerReference w:type="default" r:id="rId_header1"/>')}</w:pPr><w:r><w:t>一</w:t></w:r></w:p>` +
        `<w:p><w:pPr>${sect()}</w:pPr><w:r><w:t>二</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>三</w:t></w:r></w:p>${sect()}`,
      { header1: PLAIN_HEADER },
    );
    const { ed } = await open(zip, { confirm: () => true });
    pagesFor(ed, 3);
    const view = ed.view!;
    // The middle section gets its own header (unlinked) ...
    ed.editHeaderFooter('header', 1);
    ed.setHeaderFooterLinked(false);
    ed.closeHeaderFooter();
    // ... its margins change (an undo step, its w:sectPr now naming that header) ...
    cursorIn(ed, 1);
    expect(ed.cursorSection().index).toBe(1);
    ed.setPageSetup({ ...ed.cursorSection().page, marginLeft: 1000 });
    // ... it is linked again (its own header goes), and the margin change is undone.
    ed.editHeaderFooter('header', 1);
    ed.setHeaderFooterLinked(true);
    ed.closeHeaderFooter();
    undo(view.state, view.dispatch);
    cursorIn(ed, 1);
    expect(ed.cursorSection().page.marginLeft).toBe(1800);
    const out = await saved(ed);
    await referencesResolve(out);
    const [, middle] = await sectionHeaders(out);
    expect(middle).toEqual({});
  });

  it('a section break brought back by Ctrl+Z does not refer to a header/footer that is gone', () => {
    // The w:sectPr an undo brings back names the header it had then; that header is gone since.
    const sectPr = '<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:headerReference w:type="default" r:id="rIdGone"/><w:footerReference w:type="default" r:id="rIdKept"/></w:sectPr>';
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create({ sectPr }, schema.text('一')),
      schema.nodes.paragraph.create(null, schema.text('二')),
    ]);
    let state = EditorState.create({ schema, doc, plugins: [history(), keepSectionReferences((id) => id !== 'rIdGone')] });
    // Removing the section break (an undo step), then undoing it.
    state = state.apply(state.tr.setNodeMarkup(0, undefined, { ...state.doc.firstChild!.attrs, sectPr: null }));
    undo(state, (tr) => (state = state.apply(tr)));
    const back = state.doc.firstChild!.attrs.sectPr as string;
    expect(back).toBeTruthy();
    expect(back).not.toContain('rIdGone');
    expect(back).toContain('r:id="rIdKept"');
  });

  it('a header given content on a section that is not the last stays referenced after Ctrl+Z of an earlier page setup change', async () => {
    const zip = await pkg(`<w:p><w:pPr><w:sectPr>${SECT}</w:sectPr></w:pPr><w:r><w:t>一</w:t></w:r></w:p><w:p><w:r><w:t>二</w:t></w:r></w:p><w:sectPr>${SECT}</w:sectPr>`);
    const { ed } = await open(zip);
    pagesFor(ed, 2);
    const view = ed.view!;
    cursorIn(ed, 0);
    ed.setPageSetup({ ...ed.cursorSection().page, marginLeft: 1000 });
    // Typing into section 1's empty header makes it (claimHeaderFooter).
    ed.editHeaderFooter('header', 0);
    const hv = ed.activeView!;
    hv.dispatch(hv.state.tr.insertText('新頁首', 1));
    ed.closeHeaderFooter();
    undo(view.state, view.dispatch);
    const out = await saved(ed);
    await referencesResolve(out);
    const [first] = await sectionHeaders(out);
    expect(Object.keys(first)).toEqual(['default']);
    expect(await text(out, first.default)).toContain('新頁首');
  });

  it('rejecting a tracked deletion of a watermark from the file gives the watermark back', async () => {
    const del = `<w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z">${WORD_RUN}</w:del>`;
    const header = WORD_HEADER.replace(WORD_RUN, del);
    const { ed } = await open(await pkg(`<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rId_header1"/>${SECT}</w:sectPr>`, { header1: header }));
    expect(ed.hasWatermark()).toBe(false);
    ed.editHeaderFooter('header', 0);
    const hv = ed.activeView!;
    /** Watermark runs and tracked deletions holding one, in the header being edited. */
    const count = () => {
      let runs = 0;
      let dels = 0;
      hv.state.doc.descendants((n) => {
        if (n.type.name !== 'raw_inline') return;
        if (isWatermarkRun(n.attrs.xml)) runs++;
        else if (n.attrs.label === 'del' && /PowerPlusWaterMarkObject/.test(n.attrs.xml)) dels++;
      });
      return { runs, dels };
    };
    rejectAllRevisions(hv.state, hv.dispatch);
    expect(count()).toEqual({ runs: 1, dels: 0 });
    // Ctrl+Z gives the tracked deletion back (not a watermark as well); Ctrl+Y rejects it again.
    undo(hv.state, hv.dispatch);
    expect(count()).toEqual({ runs: 0, dels: 1 });
    redo(hv.state, hv.dispatch);
    expect(count()).toEqual({ runs: 1, dels: 0 });
    ed.closeHeaderFooter();
    expect(ed.watermark()).toMatchObject({ text: 'CONFIDENTIAL' });
  });
});

// ----- the ribbon and the dialog -----

async function mountEditor(src: Uint8Array, extra: Record<string, unknown> = {}) {
  const host = document.createElement('div');
  document.body.append(host);
  let editor: DocxEditor | null = null;
  const app = createApp({ render: () => h(DocxEditorVue, { src, onReady: (ed: DocxEditor) => (editor = ed), ...extra }) });
  app.mount(host);
  const flush = async () => { for (let i = 0; i < 7; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };
  for (let i = 0; i < 80 && !editor; i++) await flush();
  await flush();
  if (!editor) throw new Error('DocxEditorVue did not emit ready');
  return { host, editor: editor as DocxEditor, flush, done: () => app.unmount() };
}

describe('設計 › 頁面背景 › 浮水印', () => {
  it('read-only: no 設計 tab and no 浮水印 button', async () => {
    const ui = await mountEditor(await (await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER })).generateAsync({ type: 'uint8array' }), { editable: false, fileMenu: true });
    const tabs = Array.from(ui.host.querySelectorAll<HTMLButtonElement>('[role="tab"]')).map((t) => t.textContent?.trim());
    expect(tabs).toEqual(['檢視']);
    const button = Array.from(ui.host.querySelectorAll<HTMLButtonElement>('button')).find((b) => b.textContent?.includes('浮水印'));
    expect(!button || button.disabled || button.closest<HTMLElement>('[style*="display: none"]') !== null).toBe(true);
    ui.done();
  });

  it('the 設計 tab sits between 插入 and 版面配置, with 頁面背景 › 浮水印 and its menu', async () => {
    const ui = await mountEditor(await (await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER })).generateAsync({ type: 'uint8array' }));
    const tabs = Array.from(ui.host.querySelectorAll<HTMLButtonElement>('[role="tab"]')).map((t) => t.textContent?.trim());
    expect(tabs.slice(0, 4)).toEqual(['常用', '插入', '設計', '版面配置']);
    Array.from(ui.host.querySelectorAll<HTMLButtonElement>('[role="tab"]')).find((t) => t.textContent?.trim() === '設計')!.click();
    await ui.flush();
    const group = ui.host.querySelector('[role="group"][aria-label="頁面背景"]')!;
    expect(group).not.toBeNull();
    const button = group.querySelector<HTMLButtonElement>('button.dx-big[aria-haspopup="menu"]')!;
    expect(button.textContent).toContain('浮水印');
    button.click();
    await ui.flush();
    const menu = group.querySelector('[role="menu"]')!;
    const items = Array.from(menu.querySelectorAll<HTMLButtonElement>('[role="menuitem"]'));
    for (const preset of WATERMARK_PRESETS) expect(items.some((b) => b.textContent?.includes(preset)), preset).toBe(true);
    const remove = items.find((b) => b.textContent?.includes('移除浮水印'))!;
    expect(remove.disabled).toBe(true);
    expect(items.some((b) => b.textContent?.includes('自訂浮水印…'))).toBe(true);
    items.find((b) => b.getAttribute('aria-label') === '機密' || b.textContent?.trim() === '機密')!.click();
    await ui.flush();
    expect(ui.editor.watermark()).toMatchObject({ kind: 'text', text: '機密', layout: 'diagonal', semitransparent: true });
    button.click();
    await ui.flush();
    const again = Array.from(group.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')).find((b) => b.textContent?.includes('移除浮水印'))!;
    expect(again.disabled).toBe(false);
    again.click();
    await ui.flush();
    expect(ui.editor.watermark()).toBeNull();
    // 自訂浮水印… opens the dialog.
    button.click();
    await ui.flush();
    Array.from(group.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')).find((b) => b.textContent?.includes('自訂浮水印…'))!.click();
    await ui.flush();
    expect(document.querySelector('[role="dialog"]')?.textContent).toContain('浮水印');
    ui.done();
  });
});

/**
 * The dialog's controls by their own names: its ids and radio names start with the dialog's own
 * prefix (one per dialog on the page), `#dx-wm-text` finds `<prefix>-text`.
 */
const own = (sel: string) => sel.replace(/#dx-wm-(\w+)/g, '[id$="-$1"]').replace(/name="dx-wm-(\w+)"/g, 'name$="-$1"');

async function mountDialog(ed: DocxEditor) {
  const host = document.createElement('div');
  document.body.append(host);
  const open = ref(true);
  const closed = vi.fn();
  const app = createApp({ render: () => h(WatermarkDialog, { editor: ed, open: open.value, onClose: closed }) });
  app.mount(host);
  const flush = async () => { for (let i = 0; i < 5; i++) { await new Promise((r) => setTimeout(r, 0)); await nextTick(); } };
  await flush();
  const $ = <T extends Element>(sel: string) => host.querySelector<T>(own(sel))!;
  const set = async (sel: string, value: string | boolean) => {
    const el = $<HTMLInputElement>(sel);
    if (typeof value === 'boolean') {
      el.checked = value;
      el.dispatchEvent(new Event('change'));
    } else {
      el.value = value;
      el.dispatchEvent(new Event('input'));
      el.dispatchEvent(new Event('change'));
    }
    await flush();
  };
  return { host, $, set, flush, closed, open, done: () => app.unmount() };
}

describe('自訂浮水印 dialog', () => {
  it('applies a text watermark with the chosen settings and opens with them again', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    let d = await mountDialog(ed);
    // Nothing yet: 無浮水印 is chosen.
    expect(d.$<HTMLInputElement>('input[name="dx-wm-kind"][value="none"]').checked).toBe(true);
    await d.set('input[name="dx-wm-kind"][value="text"]', true);
    // The text box offers the presets.
    const options = Array.from(d.host.querySelectorAll('datalist option')).map((o) => o.getAttribute('value'));
    for (const preset of WATERMARK_PRESETS) expect(options).toContain(preset);
    await d.set('#dx-wm-text', '限閱');
    await d.set('#dx-wm-size', '72');
    await d.set('#dx-wm-color', '#ff0000');
    await d.set('#dx-wm-semi', false);
    await d.set('input[name="dx-wm-layout"][value="horizontal"]', true);
    d.$<HTMLButtonElement>('.dx-primary').click();
    await d.flush();
    expect(d.closed).toHaveBeenCalled();
    d.done();
    const zip = await saved(ed);
    const [run] = pictRuns(await text(zip, 'word/header1.xml'));
    expect(run).toMatch(OUR_TEXT_RUN({ text: '限閱', font: '標楷體', size: '72pt', rotation: false, fill: '#ff0000', semi: false }));

    d = await mountDialog(ed);
    expect(d.$<HTMLInputElement>('input[name="dx-wm-kind"][value="text"]').checked).toBe(true);
    expect(d.$<HTMLInputElement>('#dx-wm-text').value).toBe('限閱');
    expect(d.$<HTMLSelectElement>('#dx-wm-size').value).toBe('72');
    expect(d.$<HTMLInputElement>('#dx-wm-color').value).toBe('#ff0000');
    expect(d.$<HTMLInputElement>('#dx-wm-semi').checked).toBe(false);
    expect(d.$<HTMLInputElement>('input[name="dx-wm-layout"][value="horizontal"]').checked).toBe(true);
    // 無浮水印 removes it.
    await d.set('input[name="dx-wm-kind"][value="none"]', true);
    d.$<HTMLButtonElement>('.dx-primary').click();
    await d.flush();
    expect(ed.watermark()).toBeNull();
    d.done();
  });

  it('an empty text is refused with a message', async () => {
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    const d = await mountDialog(ed);
    await d.set('input[name="dx-wm-kind"][value="text"]', true);
    await d.set('#dx-wm-text', '  ');
    d.$<HTMLButtonElement>('.dx-primary').click();
    await d.flush();
    expect(d.host.querySelector('[role="alert"]')?.textContent).toContain('文字');
    expect(ed.watermark()).toBeNull();
    d.done();
  });

  it('applies a picture watermark (washed out) and opens with it again; a file that is not a picture is refused', async () => {
    vi.stubGlobal(
      'Image',
      class {
        naturalWidth = 0;
        naturalHeight = 0;
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        set src(_v: string) {
          setTimeout(() => {
            this.naturalWidth = 40;
            this.naturalHeight = 20;
            this.onload?.();
          });
        }
      },
    );
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }));
    let d = await mountDialog(ed);
    await d.set('input[name="dx-wm-kind"][value="picture"]', true);
    const input = d.$<HTMLInputElement>('#dx-wm-file');
    const choose = async (file: File) => {
      Object.defineProperty(input, 'files', { value: [file], configurable: true });
      input.dispatchEvent(new Event('change'));
      for (let i = 0; i < 5; i++) await d.flush();
    };
    await choose(new File(['%PDF-1.4'], 'a.png', { type: 'image/png' }));
    expect(d.host.querySelector('[role="alert"]')?.textContent).toMatch(/圖片/);
    await choose(new File([Uint8Array.from(atob(PNG), (c) => c.charCodeAt(0))], 'logo.png', { type: 'image/png' }));
    expect(d.host.querySelector('[role="alert"]')).toBeNull();
    expect(d.$<HTMLInputElement>('#dx-wm-washout').checked).toBe(true);
    await d.set('#dx-wm-scale', '200');
    d.$<HTMLButtonElement>('.dx-primary').click();
    await d.flush();
    d.done();
    expect(ed.watermark()).toMatchObject({ kind: 'picture', washout: true });
    const [run] = pictRuns(await text(await saved(ed), 'word/header1.xml'));
    // 200 % of 40 × 20 px: 60 × 30 pt; the file's name as the picture's title.
    expect(run).toContain('width:60pt;height:30pt');
    expect(run).toContain('o:title="logo"');

    d = await mountDialog(ed);
    expect(d.$<HTMLInputElement>('input[name="dx-wm-kind"][value="picture"]').checked).toBe(true);
    expect(d.$<HTMLInputElement>('#dx-wm-washout').checked).toBe(true);
    for (let i = 0; i < 5; i++) await d.flush();
    expect(d.$<HTMLSelectElement>('#dx-wm-scale').value).toBe('200');
    // Unchanged: 確定 changes nothing.
    const before = (ed.model.headerFooters[0].doc);
    d.$<HTMLButtonElement>('.dx-primary').click();
    await d.flush();
    expect(ed.model.headerFooters[0].doc).toBe(before);
    // 自動 is chosen and reads back as 自動 (not the percentage it comes to).
    await d.set('#dx-wm-scale', 'auto');
    d.$<HTMLButtonElement>('.dx-primary').click();
    await d.flush();
    d.done();
    d = await mountDialog(ed);
    for (let i = 0; i < 5; i++) await d.flush();
    expect(d.$<HTMLSelectElement>('#dx-wm-scale').value).toBe('auto');
    d.done();
  });

  it('while 追蹤修訂 is on, 確定 is refused with the notice and nothing changes', async () => {
    const notices: string[] = [];
    const { ed } = await open(await pkg(TWO_SECTIONS, { header1: PLAIN_HEADER }), { onNotice: (m) => notices.push(m) });
    ed.setTrackChanges(true);
    const d = await mountDialog(ed);
    await d.set('input[name="dx-wm-kind"][value="text"]', true);
    await d.set('#dx-wm-text', '密');
    d.$<HTMLButtonElement>('.dx-primary').click();
    await d.flush();
    expect(ed.watermark()).toBeNull();
    expect(notices.at(-1)).toContain('追蹤修訂');
    d.done();
  });
});
