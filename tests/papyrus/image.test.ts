// Editing a picture patches its original w:drawing (size, alt text, picture) so wrapping,
// effects, links and ids survive; a replaced picture's old media is dropped.
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { EditorState, NodeSelection, type Command } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { imageNodeView, replaceImage, setImageAttrs } from '../../src/papyrus/editor/imageView';

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const GIF = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const DRAWING =
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="1A2B3C4D">' +
  '<wp:extent cx="952500" cy="476250"/><wp:effectExtent l="19050" t="0" r="0" b="0"/>' +
  '<wp:docPr id="5" name="Picture 5" descr="old alt"><a:hlinkClick r:id="rIdLink"/></wp:docPr>' +
  '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  '<pic:nvPicPr><pic:cNvPr id="0" name="x.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg"><a:extLst>' +
  '<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext>' +
  '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/></a:ext>' +
  '</a:extLst></a:blip><a:srcRect l="1000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></pic:spPr>' +
  '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';

async function open() {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body><w:p><w:r><w:rPr><w:noProof/></w:rPr>${DRAWING}</w:r></w:p>${SECT}</w:body></w:document>`);
  zip.file('word/media/image1.png', PNG, { base64: true });
  zip.file('word/media/image2.svg', '<svg xmlns="http://www.w3.org/2000/svg"/>');
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
        '<Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.svg"/>' +
        '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>' +
        '</Relationships>',
    ),
  );
  const { doc, model } = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
  let state = EditorState.create({ schema, doc });
  let pos = -1;
  doc.descendants((n, p) => {
    if (n.type.name === 'image') pos = p;
  });
  state = state.apply(state.tr.setSelection(NodeSelection.create(state.doc, pos)));
  const run = (cmd: Command) => expect(cmd(state, (tr) => (state = state.apply(tr)))).toBe(true);
  const save = async () => JSZip.loadAsync(await writeDocx(state.doc, model));
  const xmlOf = async (z: JSZip) => z.file('word/document.xml')!.async('string');
  return { run, save, xmlOf, pos, get state() { return state; } };
}

describe('picture editing', () => {
  it('an untouched picture is written back as it was', async () => {
    const d = await open();
    expect(await d.xmlOf(await d.save())).toContain(`<w:r><w:rPr><w:noProof/></w:rPr>${DRAWING}</w:r>`);
    expect(d.state.doc.nodeAt(d.pos)!.attrs).toMatchObject({ width: 100, height: 50, alt: 'old alt' });
  });

  it('resizing patches wp:extent and the picture frame only', async () => {
    const d = await open();
    d.run(setImageAttrs({ width: 200, height: 100 }));
    const xml = await d.xmlOf(await d.save());
    const want = DRAWING.replace('<wp:extent cx="952500" cy="476250"/>', '<wp:extent cx="1905000" cy="952500"/>').replace(
      '<a:ext cx="952500" cy="476250"/>',
      '<a:ext cx="1905000" cy="952500"/>',
    );
    expect(xml).toContain(`<w:r><w:rPr><w:noProof/></w:rPr>${want}</w:r>`);
  });

  it('a size in cm is exact in EMU', async () => {
    const d = await open();
    d.run(setImageAttrs({ width: 5 * (96 / 2.54) }));
    expect(await d.xmlOf(await d.save())).toContain('<wp:extent cx="1800000" cy="476250"/>');
  });

  it('alt text is wp:docPr/@descr', async () => {
    const d = await open();
    d.run(setImageAttrs({ alt: '公司標誌 & "logo"' }));
    const xml = await d.xmlOf(await d.save());
    expect(xml).toContain(DRAWING.replace('descr="old alt"', 'descr="公司標誌 &amp; &quot;logo&quot;"'));
  });

  it('replacing the picture swaps the media but keeps the drawing (effects, link, ids)', async () => {
    const d = await open();
    d.run(replaceImage(`data:image/gif;base64,${GIF}`, 40, 10));
    expect(d.state.doc.nodeAt(d.pos)!.attrs).toMatchObject({ width: 100, height: 25 });
    const zip = await d.save();
    const xml = await d.xmlOf(zip);
    const embed = /<a:blip r:embed="([^"]+)">/.exec(xml)?.[1];
    expect(embed).toBeTruthy();
    expect(embed).not.toBe('rIdImg');
    // Old SVG version and crop gone; line, link, anchor id, docPr kept; height follows the new ratio.
    expect(xml).not.toContain('svgBlip');
    expect(xml).not.toContain('srcRect');
    expect(xml).toContain('<a14:useLocalDpi');
    expect(xml).toContain('<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>');
    expect(xml).toContain('<wp:docPr id="5" name="Picture 5" descr="old alt"><a:hlinkClick r:id="rIdLink"/></wp:docPr>');
    expect(xml).toContain('wp14:anchorId="1A2B3C4D"><wp:extent cx="952500" cy="238125"/><wp:effectExtent l="19050" t="0" r="0" b="0"/>');

    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).toContain('Id="rIdLink"');
    expect(rels).not.toContain('rIdImg');
    expect(rels).not.toContain('rIdSvg');
    const target = new RegExp(`Id="${embed}"[^>]*Target="([^"]+)"`).exec(rels)?.[1];
    expect(target).toMatch(/^media\/.*\.gif$/);
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('Extension="gif" ContentType="image/gif"');
    expect(zip.file('word/' + target)).toBeTruthy();
    expect(zip.file('word/media/image1.png')).toBeNull();
    expect(zip.file('word/media/image2.svg')).toBeNull();
  });
});

describe('resize handles', () => {
  beforeAll(() => {
    const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
    Range.prototype.getBoundingClientRect ??= zero;
    Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  });

  it('a corner drag keeps the proportions and is one undo step', async () => {
    const img = schema.nodes.image.create({ src: `data:image/png;base64,${PNG}`, width: 100, height: 50 });
    const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, [schema.text('x'), img]));
    const host = document.createElement('div');
    document.body.append(host);
    const view = new EditorView(host, {
      state: EditorState.create({ schema, doc, plugins: [history()] }),
      nodeViews: { image: imageNodeView },
    });
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, 2)));
    const wrap = view.dom.querySelector('.dx-img-wrap')!;
    expect(wrap.classList.contains('ProseMirror-selectednode')).toBe(true);
    const handle = wrap.querySelector('[data-handle="se"]')!;
    handle.dispatchEvent(new MouseEvent('pointerdown', { clientX: 0, clientY: 0, bubbles: true, cancelable: true }));
    window.dispatchEvent(new MouseEvent('pointermove', { clientX: 40, clientY: 0 }));
    window.dispatchEvent(new MouseEvent('pointerup', { clientX: 40, clientY: 0 }));
    expect(view.state.doc.nodeAt(2)!.attrs).toMatchObject({ width: 120, height: 60 });
    expect(view.state.selection).toBeInstanceOf(NodeSelection);

    // The right-edge handle stretches the width only.
    // Allow a distinct user action boundary; the history plugin groups nearby transactions.
    await new Promise((resolve) => setTimeout(resolve, 600));
    wrap.querySelector('[data-handle="e"]')!.dispatchEvent(new MouseEvent('pointerdown', { clientX: 0, clientY: 0, bubbles: true }));
    window.dispatchEvent(new MouseEvent('pointermove', { clientX: 30, clientY: 10 }));
    window.dispatchEvent(new MouseEvent('pointerup'));
    expect(view.state.doc.nodeAt(2)!.attrs).toMatchObject({ width: 150, height: 60 });

    expect(undo(view.state, view.dispatch)).toBe(true);
    expect(view.state.doc.nodeAt(2)!.attrs).toMatchObject({ width: 120, height: 60 });
    expect(undo(view.state, view.dispatch)).toBe(true);
    expect(view.state.doc.nodeAt(2)!.attrs).toMatchObject({ width: 100, height: 50 });
    view.destroy();
    host.remove();
  });
});
