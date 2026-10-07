// Text boxes and shapes (phase 8A): read from Word's own files (tests/fixtures/shapes, made by
// scripts/make-shape-fixtures.ps1), drawn at Word's position, their text edited and saved into
// both copies Word writes (DrawingML and VML), everything else kept byte for byte.
import { beforeAll, describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import type { Node as PMNode } from 'prosemirror-model';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { parseFragment, serializeXml } from '../../src/papyrus/docx/xml';
import { schema } from '../../src/papyrus/editor/schema';
import { readShape, rewriteTextBoxes, textFrames, withText, vmlLength, type ShapeContext, type ShapeModel, type SpNode } from '../../src/papyrus/docx/shapes';
import { customPaths, isKnownPreset, presetPaths, presetTextRect } from '../../src/papyrus/docx/shapeGeometry';
import { readColor, readThemeColors } from '../../src/papyrus/docx/theme';
import { WORD_ONLY, anchorPosition, placedNodes, shapeDOM, wrapBox, type PlaceFrame } from '../../src/papyrus/editor/shapeView';
import { Converter } from '../../src/papyrus/docx/convert';

const fixture = (name: string) => readFileSync(`tests/fixtures/shapes/${name}.docx`);

/** The shape nodes of a document, in order. */
function shapes(doc: PMNode): { node: PMNode; shape: ShapeModel }[] {
  const out: { node: PMNode; shape: ShapeModel }[] = [];
  doc.descendants((n) => {
    if (n.attrs?.shape) out.push({ node: n, shape: n.attrs.shape });
    return true;
  });
  return out;
}

const texts = (m: ShapeModel) => textFrames(m).map((f) => schema.nodeFromJSON(f.doc).textContent);

/** The runs holding a drawing or VML picture in a document part, as written. */
function drawingRuns(xml: string): string[] {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  return Array.from(doc.getElementsByTagNameNS(W, 'r'))
    .filter((r) => Array.from(r.children).some((c) => c.localName === 'drawing' || c.localName === 'pict' || c.localName === 'AlternateContent'))
    .filter((r) => {
      for (let p = r.parentElement; p; p = p.parentElement) if (p.localName === 'txbxContent') return false;
      return true;
    })
    .map((r) => serializeXml(r));
}

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word"';

function ctx(): ShapeContext {
  const conv = new Converter(new Map(), new Map());
  return {
    theme: readThemeColors(null),
    image: () => undefined,
    blocks: (el) => conv.blocks(el),
    doc: (b) => schema.nodes.doc.create(null, b.length ? b : [schema.nodes.paragraph.create()]),
  };
}

/** The shape model of a run child given as XML. */
function modelOf(xml: string): ShapeModel | null {
  const root = new DOMParser().parseFromString(`<root ${NS}>${xml}</root>`, 'application/xml').documentElement;
  return readShape(root.firstElementChild!, ctx());
}

const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class {
    observe() {}
    unobserve() {}
    disconnect() {}
  };
  const empty = () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= empty;
  Range.prototype.getBoundingClientRect ??= zero;
});

describe('reading Word’s flowcharts, text boxes and shapes', () => {
  it('a flowchart on a drawing canvas: every symbol, its text, and the connectors with arrowheads', async () => {
    const { doc } = await readDocx(fixture('flowchart'));
    const all = shapes(doc);
    const canvas = all.find((s) => s.shape.kind === 'canvas')!.shape;
    expect(canvas.src).toBe('dml');
    expect(canvas.drawable).toBe(true);
    expect(canvas.anchor).toMatchObject({ h: { rel: 'column', offset: 0 }, v: { rel: 'paragraph', offset: 0 }, wrap: 'topAndBottom', behind: false });
    const kids = (canvas.root as { kids: SpNode[] }).kids;
    expect(kids.filter((k) => !k.cxn).map((k) => k.geom)).toEqual([
      'flowChartTerminator', 'flowChartProcess', 'flowChartDecision', 'flowChartPredefinedProcess',
      'flowChartInputOutput', 'flowChartDocument', 'flowChartTerminator',
    ]);
    const connectors = kids.filter((k) => k.cxn);
    expect(connectors).toHaveLength(7);
    expect(connectors.map((k) => k.geom).sort()).toEqual(['bentConnector3', ...Array(6).fill('straightConnector1')].sort());
    expect(connectors.every((k) => k.line?.tail?.type === 'triangle')).toBe(true);
    expect(texts(canvas)).toEqual(['開始', '收文登錄', '需要會辦？', '會辦單位', '發文', '歸檔', '結束']);
    // Colours and lines as set in Word (spPr), over the style's theme colours.
    expect(kids[0]).toMatchObject({ fill: { color: 'DEEBF7' }, line: { color: '2F5496', width: 15875 } });
    expect(kids[2].descr).toBe('判斷是否需要其他單位會辦');
    // The loose shapes after the canvas: process, decision, an arrow and an elbow connector.
    const loose = all.filter((s) => s.shape.kind !== 'canvas').map((s) => (s.shape.root as SpNode).geom);
    expect(loose.sort()).toEqual(['bentConnector3', 'flowChartDecision', 'flowChartProcess', 'line'].sort());
    // The connector's colour comes from the theme (style lnRef accent1) of this file.
    const elbow = all.find((s) => (s.shape.root as SpNode).geom === 'bentConnector3' && s.shape.kind !== 'canvas')!.shape;
    expect((elbow.root as SpNode).line?.color).toMatch(/^[0-9A-F]{6}$/);
  });

  it('text boxes: wrapped, several paragraphs with a list, in the line, in the header, with alternative text', async () => {
    const { doc, model } = await readDocx(fixture('textboxes'));
    const all = shapes(doc).map((s) => s.shape);
    const boxes = all.filter((s) => s.kind === 'textbox');
    expect(boxes.map((b) => texts(b)[0])).toEqual(['繞圖文字方塊', '會議重點第一項：預算第二項：人力', '同列方塊']);
    expect(boxes[0].anchor).toMatchObject({ wrap: 'square', h: { rel: 'column' }, v: { rel: 'paragraph' } });
    expect(boxes[1].anchor?.wrap).toBe('none');
    expect(boxes[2].anchor).toBeNull(); // wp:inline
    expect(boxes[1]).toMatchObject({ title: '會議重點', descr: '列出兩項會議重點的文字方塊' });
    const listDoc = schema.nodeFromJSON(textFrames(boxes[1])[0].doc);
    expect(listDoc.childCount).toBe(3);
    expect(listDoc.child(1).attrs.numId).toBeTruthy();
    // The header's text box is placed from the page.
    const header = model.headerFooters.find((h) => h.kind === 'header' && h.type === 'default')!;
    const hb = shapes(header.doc)[0].shape;
    expect(texts(hb)).toEqual(['頁首文字方塊']);
    expect(hb.anchor).toMatchObject({ h: { rel: 'page' }, v: { rel: 'page' } });
    // Shapes with text are shapes, not text boxes; their text is still editable.
    expect(all.filter((s) => s.kind === 'shape').flatMap(texts)).toEqual(['乙', '甲']);
  });

  it('shapes: each preset drawn; one Word can’t be drawn here keeps Word’s look and says so', async () => {
    const { doc } = await readDocx(fixture('shapes'));
    const all = shapes(doc).map((s) => s.shape);
    const geoms = all.filter((s) => s.root.t === 'sp').map((s) => (s.root as SpNode).geom);
    for (const g of ['rect', 'roundRect', 'ellipse', 'diamond', 'parallelogram', 'rightArrow', 'leftRightArrow', 'line']) expect(geoms).toContain(g);
    const star = all.find((s) => (s.root as SpNode).geom === 'star5')!;
    expect(star.drawable).toBe(false);
    expect(all.filter((s) => s !== star).every((s) => s.drawable)).toBe(true);
    const behind = all.find((s) => s.anchor?.behind)!;
    expect(behind.anchor?.wrap).toBe('none');
    expect(all.find((s) => s.anchor?.wrap === 'topAndBottom')).toBeTruthy();
    const dashed = all.find((s) => (s.root as SpNode).geom === 'line')!;
    expect((dashed.root as SpNode).line).toMatchObject({ color: 'FF0000', dash: 'dash' });
    const alt = all.find((s) => s.title === '黃色矩形')!;
    expect(alt.descr).toBe('有標題與描述的黃色矩形');
  });

  it('groups: the children are placed through the group’s coordinates; turned shapes keep their angle', () => {
    const m = modelOf(
      '<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>' +
        '<wp:extent cx="1905000" cy="952500"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="3" name="群組 3"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
        '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="3810000" cy="1905000"/></a:xfrm></wpg:grpSpPr>' +
        '<wps:wsp><wps:cNvPr id="4" name="a"/><wps:cNvSpPr/><wps:spPr><a:xfrm rot="1800000"><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr>' +
        `<wps:txbx><w:txbxContent>${para('甲')}</w:txbxContent></wps:txbx><wps:bodyPr anchor="ctr"/></wps:wsp>` +
        '<wps:wsp><wps:cNvPr id="5" name="b"/><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="1905000" y="952500"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></wps:spPr>' +
        '<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></wps:style>' +
        `<wps:txbx><w:txbxContent>${para('乙')}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing>`,
    )!;
    expect(m.kind).toBe('group');
    expect(m.anchor).toMatchObject({ h: { rel: 'margin', align: 'center' }, v: { rel: 'page', offset: 914400 }, wrap: 'square', side: 'bothSides' });
    const placed = placedNodes(m);
    // Children space 3810000 wide mapped onto 1905000 EMU (200 px): halves.
    expect(placed.map((p) => [p.x, p.y, p.w, p.h])).toEqual([[0, 0, 100, 50], [100, 50, 100, 50]]);
    expect((placed[0].node as SpNode).rot).toBe(30);
    // The theme's accent1 (default Office theme) for the style's colours; the line shaded by half.
    expect((placed[1].node as SpNode).fill?.color).toBe('4472C4');
    expect((placed[1].node as SpNode).line?.color).toBe('223962');
    expect(texts(m)).toEqual(['甲', '乙']);
  });

  it('VML only (older Word): text boxes, rectangles, ovals, lines and flowchart types', () => {
    const m = modelOf(
      '<w:pict><v:shapetype id="_x0000_t110" coordsize="21600,21600" o:spt="110" path="m10800,l,10800,10800,21600,21600,10800xe"/>' +
        '<v:shape id="d" type="#_x0000_t110" alt="判斷" style="position:absolute;margin-left:36pt;margin-top:12pt;width:144pt;height:72pt;z-index:3;mso-position-horizontal-relative:page;mso-position-vertical-relative:text" fillcolor="#ffc000" strokecolor="red" strokeweight="2pt">' +
        `<v:textbox inset="0,0,0,0"><w:txbxContent>${para('要不要')}</w:txbxContent></v:textbox><w10:wrap type="square"/></v:shape></w:pict>`,
    )!;
    expect(m).toMatchObject({ src: 'vml', kind: 'shape', w: 144 * 12700, h: 72 * 12700, drawable: true, descr: '判斷' });
    expect(m.anchor).toMatchObject({ h: { rel: 'page', offset: 36 * 12700 }, v: { rel: 'paragraph', offset: 12 * 12700 }, wrap: 'square', behind: false, z: 3 });
    const sp = m.root as SpNode;
    expect(sp).toMatchObject({ geom: 'flowChartDecision', fill: { color: 'FFC000' }, line: { color: 'FF0000', width: 2 * 12700 } });
    expect(sp.text?.inset).toEqual([0, 0, 0, 0]);
    expect(texts(m)).toEqual(['要不要']);
    const line = modelOf('<w:pict><v:line from="0,0" to="100pt,-20pt" style="position:absolute"><v:stroke endarrow="block"/></v:line></w:pict>')!;
    expect(line.root).toMatchObject({ geom: 'line', w: 100 * 12700, h: 20 * 12700, flipV: true, line: { tail: { type: 'triangle' } } });
    expect(modelOf('<w:pict><v:oval style="width:10pt;height:10pt" filled="f"/></w:pict>')).toMatchObject({ anchor: null, root: { geom: 'ellipse', fill: null } });
    // A VML picture is not a shape (it is shown as a picture already).
    expect(modelOf('<w:pict><v:shape style="width:10pt;height:10pt"><v:imagedata r:id="rId9"/></v:shape></w:pict>')).toBeNull();
    expect(vmlLength('1in')).toBe(914400);
  });

  it('theme colours with shade, tint and luminance changes', () => {
    const theme = readThemeColors(new DOMParser().parseFromString(
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:accent1><a:srgbClr val="156082"/></a:accent1></a:clrScheme></a:themeElements></a:theme>',
      'application/xml',
    ));
    expect(theme.accent1).toBe('156082');
    const el = (x: string) => new DOMParser().parseFromString(`<a:x xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${x}</a:x>`, 'application/xml').documentElement.firstElementChild;
    expect(readColor(el('<a:schemeClr val="tx1"/>'), theme)).toEqual({ hex: '000000' });
    expect(readColor(el('<a:srgbClr val="808080"><a:tint val="50000"/></a:srgbClr>'), theme)?.hex).toBe('C0C0C0');
    expect(readColor(el('<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>'), theme)).toEqual({ hex: 'FF0000', alpha: 0.5 });
    expect(readColor(el('<a:schemeClr val="bg1"><a:lumMod val="50000"/></a:schemeClr>'), theme)?.hex).toBe('808080');
  });
});

describe('geometry', () => {
  it('draws the flowchart symbols, arrows, lines and connectors', () => {
    for (const p of ['rect', 'roundRect', 'ellipse', 'diamond', 'parallelogram', 'flowChartProcess', 'flowChartDecision', 'flowChartInputOutput',
      'flowChartDocument', 'flowChartTerminator', 'flowChartPredefinedProcess', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow',
      'line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5', 'curvedConnector3']) {
      expect(isKnownPreset(p), p).toBe(true);
      expect(presetPaths(p, 100, 50)!.length, p).toBeGreaterThan(0);
    }
    expect(presetPaths('star5', 10, 10)).toBeNull();
    expect(presetPaths('flowChartDecision', 100, 50)![0].d).toBe('M50,0L100,25L50,50L0,25Z');
    expect(presetPaths('bentConnector3', 100, 50)![0]).toEqual({ d: 'M0,0H50V50H100', fill: false, stroke: true });
    expect(presetPaths('flowChartPredefinedProcess', 80, 40)!.map((p) => p.fill)).toEqual([true, false]);
    // Text goes inside the diamond, as in Word.
    expect(presetTextRect('flowChartDecision', 100, 60)).toEqual({ l: 25, t: 15, r: 75, b: 45 });
    expect(customPaths([{ w: 10, h: 10, fill: true, stroke: true, cmds: [['M', 0, 0], ['L', 10, 0], ['L', 10, 10], ['Z']] }], 20, 40)[0].d).toBe('M0,0L20,0L20,40Z');
  });
});

describe('drawing and placing', () => {
  it('draws a text box as SVG with its text; the screen reader hears its name, then its text', async () => {
    await import('../../src/papyrus/editor/shapeView');
    const { doc } = await readDocx(fixture('textboxes'));
    const node = shapes(doc).find((s) => s.shape.title === '會議重點')!.node;
    const dom = shapeDOM(node)!;
    expect(dom.classList.contains('dx-shape-float')).toBe(true);
    expect(dom.getAttribute('role')).toBe('group');
    expect(dom.getAttribute('aria-label')).toBe('文字方塊：會議重點：列出兩項會議重點的文字方塊');
    expect(dom.querySelector('svg')!.getAttribute('aria-hidden')).toBe('true');
    expect(dom.querySelector('svg path')).toBeTruthy();
    const text = dom.querySelector<HTMLElement>('.dx-shape-text')!;
    expect(text.dataset.text).toBe('0');
    expect(text.querySelectorAll('p')).toHaveLength(3);
    expect(text.textContent).toBe('會議重點第一項：預算第二項：人力');
    expect(JSON.parse(dom.dataset.anchor!)).toMatchObject({ h: { rel: 'column' }, v: { rel: 'paragraph' }, wrap: 'none' });
    // The drawing through the schema, as a header copy and the printout get it.
    const viaSchema = (schema.nodes.raw_inline.spec.toDOM!(node) as HTMLElement);
    expect(viaSchema.classList.contains('dx-shape')).toBe(true);
  });

  it('a shape without text but with alternative text is an image; one that can’t be drawn says so', async () => {
    const { doc } = await readDocx(fixture('shapes'));
    const star = shapes(doc).find((s) => (s.shape.root as SpNode).geom === 'star5')!.node;
    const dom = shapeDOM(star)!;
    expect(dom.getAttribute('role')).toBe('img');
    expect(dom.getAttribute('aria-label')).toContain(WORD_ONLY);
    expect(dom.textContent).toContain(WORD_ONLY);
    expect(dom.querySelector('.dx-shape-unknown')).toBeTruthy();
    const plain = shapes(doc).find((s) => (s.shape.root as SpNode).geom === 'rightArrow')!.node;
    expect(shapeDOM(plain)!.getAttribute('aria-hidden')).toBe('true');
    const inline = shapes((await readDocx(fixture('textboxes'))).doc).find((s) => !s.shape.anchor)!.node;
    const idom = shapeDOM(inline)!;
    expect(idom.classList.contains('dx-shape-inline')).toBe(true);
    expect(idom.style.width).toBe(`${(1524000 + 19050) / 9525}px`);
  });

  it('places an anchored shape from the page, the margins, the column or the paragraph', () => {
    const frame: PlaceFrame = {
      page: { left: 0, top: 1000, right: 800, bottom: 2100 },
      margins: [96, 120, 96, 120],
      column: { left: 120, top: 1096, right: 680, bottom: 2004 },
      paragraph: { left: 120, top: 1300, right: 680, bottom: 1324 },
      mark: { left: 300, top: 1302, right: 300, bottom: 1320 },
    };
    const emu = (px: number) => px * 9525;
    expect(anchorPosition({ rel: 'page', offset: emu(50) }, { rel: 'page', offset: emu(40) }, 100, 20, frame)).toEqual({ x: 50, y: 1040 });
    expect(anchorPosition({ rel: 'column', offset: emu(10) }, { rel: 'paragraph', offset: emu(-5) }, 100, 20, frame)).toEqual({ x: 130, y: 1295 });
    expect(anchorPosition({ rel: 'margin', align: 'center' }, { rel: 'margin', align: 'bottom' }, 100, 20, frame)).toEqual({ x: 350, y: 1984 });
    expect(anchorPosition({ rel: 'margin', align: 'right' }, { rel: 'line', offset: 0 }, 100, 20, frame)).toEqual({ x: 580, y: 1302 });
    expect(anchorPosition({ rel: 'character', offset: emu(4) }, { rel: 'topMargin', align: 'center' }, 100, 20, frame)).toEqual({ x: 304, y: 1038 });
    expect(anchorPosition({ rel: 'page', pct: 50000 }, { rel: 'page', offset: 0 }, 100, 20, frame).x).toBe(400);
    // Without pages (a test, a detached copy): from the column and the paragraph.
    expect(anchorPosition({ rel: 'page', offset: 0 }, { rel: 'page', offset: 0 }, 1, 1, { ...frame, page: null })).toEqual({ x: 120, y: 1300 });
  });

  it('keeps the text clear of a wrapped shape: on its side, below its top, or above and below it', () => {
    const frame: PlaceFrame = {
      page: null,
      margins: [0, 0, 0, 0],
      column: { left: 0, top: 0, right: 600, bottom: 1000 },
      paragraph: { left: 0, top: 100, right: 600, bottom: 124 },
      mark: { left: 0, top: 100, right: 0, bottom: 116 },
    };
    const a = { h: { rel: 'column' }, v: { rel: 'paragraph' }, w: 200, hgt: 90, wrap: 'square' as const, dist: [0, 0, 12, 12], behind: false };
    // On the right of the column: the text keeps the left.
    expect(wrapBox(a, 380, 100, frame)).toEqual({ side: 'right', css: 'display:block;float:right;width:224px;height:90px;margin:0 8px 0 0;shape-outside:inset(0px 0 0 0)' });
    // Lower in the paragraph: text may use the part above it.
    expect(wrapBox(a, 20, 130, frame)!.css).toBe('display:block;float:left;width:224px;height:120px;margin:0 0 0 8px;shape-outside:inset(30px 0 0 0)');
    expect(wrapBox({ ...a, wrap: 'topAndBottom' }, 20, 100, frame)!.css).toContain('width:100%');
    expect(wrapBox(a, 20, -200, frame)).toBeNull();
  });
});

describe('saving', () => {
  it('an untouched document keeps every drawing byte for byte', async () => {
    for (const name of ['flowchart', 'textboxes', 'shapes']) {
      const bytes = fixture(name);
      const { doc, model } = await readDocx(bytes);
      for (const hf of model.headerFooters) hf.dirty = true;
      const out = await JSZip.loadAsync(await writeDocx(doc, model));
      const src = await JSZip.loadAsync(bytes);
      for (const part of ['word/document.xml', 'word/header1.xml', 'word/header2.xml']) {
        const before = await src.file(part)?.async('string');
        if (!before) continue;
        expect(drawingRuns((await out.file(part)!.async('string'))!), `${name} ${part}`).toEqual(drawingRuns(before));
      }
    }
  });

  it('an edited text box is written into both copies (Word 2010+ and older Word); nothing else changes', async () => {
    const { doc, model } = await readDocx(fixture('flowchart'));
    const found = shapes(doc).find((s) => s.shape.kind === 'canvas')!;
    const frame = textFrames(found.shape)[2]; // 需要會辦？
    const text = schema.nodeFromJSON(frame.doc);
    const para0 = text.firstChild!;
    const children: PMNode[] = [];
    para0.forEach((c) => children.push(c));
    const typed = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(para0.attrs, [...children, schema.text('（含跨機關）', para0.lastChild!.marks)])]);
    let pos = -1;
    doc.descendants((n, p) => {
      if (n === found.node) pos = p;
      return pos < 0;
    });
    const shape = withText(found.shape, frame.i, typed.toJSON());
    const { EditorState } = await import('prosemirror-state');
    const state = EditorState.create({ schema, doc });
    const next = state.apply(state.tr.setNodeMarkup(pos, undefined, { ...found.node.attrs, shape }, found.node.marks)).doc;
    const out = await JSZip.loadAsync(await writeDocx(next, model));
    const xml = await out.file('word/document.xml')!.async('string');
    const before = drawingRuns(await (await JSZip.loadAsync(fixture('flowchart'))).file('word/document.xml')!.async('string'));
    const after = drawingRuns(xml);
    const changed = after.filter((r, i) => r !== before[i]);
    expect(changed).toHaveLength(1);
    // In the mc:Choice (wps) and in the mc:Fallback (VML) copy.
    expect(changed[0].match(/需要會辦？（含跨機關）/g)).toHaveLength(2);
    expect(changed[0]).not.toMatch(/需要會辦？<\/w:t>/);
    // Only the two w:txbxContent differ from the run as read.
    const strip = (x: string) => x.replace(/<w:txbxContent>.*?<\/w:txbxContent>/g, (m) => (m.includes('需要會辦') ? '[box]' : m));
    const orig = before[after.indexOf(changed[0])];
    expect(strip(changed[0])).toBe(strip(orig));
  });

  it('rewriteTextBoxes finds the VML twin by its content, and leaves a model without edits alone', () => {
    const run = `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:txbx><w:txbxContent>${para('一')}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>` +
      `<mc:Fallback><w:pict><v:rect style="width:1pt;height:1pt"><v:textbox><w:txbxContent>${para('一')}</w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
    const xml = serializeXml(parseFragment(run));
    const m = modelOf(new DOMParser().parseFromString(`<root ${NS}>${run}</root>`, 'application/xml').documentElement.firstElementChild!.innerHTML)!;
    expect(rewriteTextBoxes(xml, m, () => 'x')).toBe(xml);
    const out = rewriteTextBoxes(xml, withText(m, 0, {}), () => para('二'));
    expect(out.match(/<w:t>二<\/w:t>/g)).toHaveLength(2);
    expect(out).not.toContain('<w:t>一</w:t>');
  });
});
