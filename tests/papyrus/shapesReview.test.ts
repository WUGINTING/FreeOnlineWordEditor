// Review fixes of the text box / shape display (phase 8A): watermarks, print, the VML twin of
// identical text boxes, copies getting keys and ids of their own, links in drawn text boxes,
// finding a shape's page, and the text box editor's robustness.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '../../src/papyrus/editor/schema';
import { Converter } from '../../src/papyrus/docx/convert';
import { readThemeColors } from '../../src/papyrus/docx/theme';
import { readDocx } from '../../src/papyrus/docx/reader';
import { readShape, rewriteTextBoxes, withText, type ShapeContext, type ShapeModel, type SpNode } from '../../src/papyrus/docx/shapes';
import { pageOf, shapeDOM, WORD_ONLY } from '../../src/papyrus/editor/shapeView';
import { ShapeTextSession, findShape, shapeKeys } from '../../src/papyrus/editor/shapeEdit';
import { parseFragment, serializeXml } from '../../src/papyrus/docx/xml';
import { DocxEditor } from '../../src/papyrus/editor/core';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word"';

function ctx(): ShapeContext {
  return {
    theme: readThemeColors(null),
    image: () => undefined,
    blocks: (el) => new Converter(new Map(), new Map()).blocks(el),
    doc: (b) => schema.nodes.doc.create(null, b.length ? b : [schema.nodes.paragraph.create()]),
  };
}
const rootOf = (xml: string) => new DOMParser().parseFromString(`<root ${NS}>${xml}</root>`, 'application/xml').documentElement;
const modelOf = (xml: string) => readShape(rootOf(xml).firstElementChild!, ctx());
const para = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;

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
afterEach(() => vi.restoreAllMocks());

/** The first shape of a document whose model passes `test`. */
function first(doc: PMNode, test: (m: ShapeModel) => boolean): { pos: number; node: PMNode } {
  let found: { pos: number; node: PMNode } | null = null;
  doc.descendants((n, pos) => {
    if (!found && n.attrs?.shape && test(n.attrs.shape)) found = { pos, node: n };
    return !found;
  });
  return found!;
}

describe('watermarks and the printout', () => {
  it('draws a VML WordArt watermark as faint text, not as a shape Word must edit', () => {
    const m = modelOf(
      '<w:pict><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" path="m@7,l@8,m@5,21600l@6,21600e"/>' +
        '<v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:412.4pt;height:137.45pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f">' +
        '<v:fill opacity=".5"/><v:textpath style="font-family:&quot;標楷體&quot;;font-size:1pt" string="草稿"/></v:shape></w:pict>',
    )!;
    expect(m.drawable).toBe(true);
    expect((m.root as SpNode).art).toEqual({ text: '草稿', font: '標楷體' });
    expect(m.anchor).toMatchObject({ behind: true, h: { rel: 'margin', align: 'center' }, v: { rel: 'margin', align: 'center' } });
    const dom = shapeDOM(schema.nodes.raw_inline.create({ xml: '', label: '物件', shape: m }))!;
    const text = dom.querySelector('svg text')!;
    expect(text.textContent).toBe('草稿');
    expect(text.getAttribute('fill')).toBe('#C0C0C0');
    expect(text.getAttribute('fill-opacity')).toBe('0.5');
    expect(text.parentElement!.getAttribute('transform')).toContain('rotate(315');
    expect(dom.textContent).not.toContain(WORD_ONLY);
  });

  it('prints no editor notes, and header copies show none', () => {
    const css = readFileSync('src/papyrus/editor/shapes.css', 'utf8');
    const print = css.slice(css.indexOf('@media print'));
    expect(print).toMatch(/\.dx-shape-note\s*\{\s*display:\s*none/);
    expect(print).toMatch(/\.dx-shape-unknown\s*\{\s*border:\s*0;\s*background:\s*none/);
    expect(css).toMatch(/\.dx-header \.dx-shape-note,\s*\.dx-footer \.dx-shape-note\s*\{\s*display:\s*none/);
  });
});

describe('the VML twin of two text boxes with the same text', () => {
  it('pairs them by place first', () => {
    const box = (x: number) =>
      `<wps:wsp><wps:cNvPr id="${x + 10}" name="b${x}"/><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="${x * 1000}" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${para('同')}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`;
    const vbox = (x: number) => `<v:rect style="position:absolute;left:${x * 10};top:0;width:10;height:10"><v:textbox><w:txbxContent>${para('同')}</w:txbxContent></v:textbox></v:rect>`;
    const run =
      `<w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing><wp:inline><wp:extent cx="2000" cy="1000"/><wp:docPr id="1" name="g"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1000"/></a:xfrm></wpg:grpSpPr>${box(0)}${box(1)}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>` +
      `<mc:Fallback><w:pict><v:group style="width:20pt;height:10pt" coordsize="20,10">${vbox(0)}${vbox(1)}</v:group></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
    const m = readShape(rootOf(run).firstElementChild!.firstElementChild!, ctx())!;
    const xml = serializeXml(parseFragment(run));
    const out = rewriteTextBoxes(xml, withText(m, 1, {}), () => para('改'));
    const vml = out.slice(out.indexOf('<mc:Fallback>'));
    // The second VML text box, not the first, has the new text.
    expect(vml.indexOf('<w:t>同</w:t>')).toBeLessThan(vml.indexOf('<w:t>改</w:t>'));
    expect(vml.match(/<w:t>改<\/w:t>/g)).toHaveLength(1);
  });
});

describe('copies', () => {
  it('get a key and drawing ids of their own, in the same step', async () => {
    const { doc } = await readDocx(readFileSync('tests/fixtures/shapes/flowchart.docx'));
    const state = EditorState.create({ schema, doc, plugins: [shapeKeys()] });
    const { pos, node } = first(doc, (m) => m.kind === 'canvas');
    // A copy of the canvas right after it.
    const next = state.apply(state.tr.insert(pos + 1, node));
    const a = next.doc.nodeAt(pos)!;
    const b = next.doc.nodeAt(pos + 1)!;
    const ma = a.attrs.shape as ShapeModel;
    const mb = b.attrs.shape as ShapeModel;
    expect(ma.key).not.toBe(mb.key);
    expect(ma.docId).not.toBe(mb.docId);
    expect(findShape(next.doc, mb.key)!.pos).toBe(pos + 1);
    // Its XML says the same ids; its connectors are attached to its own shapes.
    expect(b.attrs.xml).toContain(`<wp:docPr id="${mb.docId}"`);
    const kids = (mb.root as { kids: SpNode[] }).kids;
    const kidIds = kids.filter((k) => !k.cxn).map((k) => k.id);
    expect(kids.filter((k) => k.cxn).every((c) => kidIds.includes(c.st!.id) && kidIds.includes(c.end!.id))).toBe(true);
    for (const id of kidIds) expect(b.attrs.xml).toContain(`id="${id}"`);
    expect((ma.root as { kids: SpNode[] }).kids.some((k) => kidIds.includes(k.id))).toBe(false);
  });
});

describe('links in drawn text boxes', () => {
  it('open only after asking', () => {
    const text = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, schema.text('網站', [schema.marks.link.create({ href: 'https://example.gov.tw/' })])));
    const m = modelOf(
      `<w:drawing><wp:inline><wp:extent cx="1000" cy="1000"/><wp:docPr id="1" name="t"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent>${para('x')}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`,
    )!;
    const dom = shapeDOM(schema.nodes.raw_inline.create({ xml: '', shape: withText(m, 0, text.toJSON()) }))!;
    document.body.append(dom);
    const a = dom.querySelector('a')!;
    const confirm = vi.spyOn(window, 'confirm').mockReturnValue(false);
    const open = vi.spyOn(window, 'open').mockReturnValue(null);
    const e = new MouseEvent('click', { bubbles: true, cancelable: true });
    a.dispatchEvent(e);
    expect(e.defaultPrevented).toBe(true);
    expect(confirm).toHaveBeenCalledWith(expect.stringContaining('https://example.gov.tw/'));
    expect(open).not.toHaveBeenCalled();
    confirm.mockReturnValue(true);
    a.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
    expect(open).toHaveBeenCalledWith('https://example.gov.tw/', '_blank', 'noopener,noreferrer');
    dom.remove();
  });
});

describe('finding a shape’s page', () => {
  it('takes the last page starting above it (measured once per placing)', () => {
    const canvas = document.createElement('div');
    canvas.className = 'dx-canvas';
    const pages = document.createElement('div');
    pages.className = 'dx-pages';
    const host = document.createElement('div');
    canvas.append(pages, host);
    let measured = 0;
    const els = [0, 1100, 2200].map((t) => {
      const p = document.createElement('div');
      p.className = 'dx-page';
      p.getBoundingClientRect = () => {
        measured++;
        return { top: t, bottom: t + 1080, left: 0, right: 800, width: 800, height: 1080, x: 0, y: t, toJSON() {} } as DOMRect;
      };
      pages.append(p);
      return p;
    });
    const shape = document.createElement('span');
    host.append(shape);
    document.body.append(canvas);
    const cache = new Map();
    expect(pageOf(shape, 50, cache)).toBe(els[0]);
    expect(pageOf(shape, 1150, cache)).toBe(els[1]);
    expect(pageOf(shape, 1090, cache)).toBe(els[0]); // in the gap below page 1
    expect(pageOf(shape, 9999, cache)).toBe(els[2]);
    expect(measured).toBe(3);
    canvas.remove();
  });
});

describe('the text box editor', () => {
  it('closes when its shape goes away, and leaves nothing behind when it can’t start', async () => {
    const host = document.createElement('div');
    document.body.append(host);
    const ed = new DocxEditor(host);
    await ed.open(readFileSync('tests/fixtures/shapes/textboxes.docx'));
    const key = (first(ed.view!.state.doc, (m) => m.kind === 'textbox').node.attrs.shape as ShapeModel).key;
    expect(ed.editShapeText(key, 0)).toBe(true);
    const found = findShape(ed.view!.state.doc, key)!;
    (ed as any).shapeSession.host.view.dispatch(ed.view!.state.tr.delete(found.pos, found.pos + 1));
    (ed as any).shapeSession?.place();
    expect(ed.target).toBe('body');
    expect(host.querySelector('.dx-shape-edit-box')).toBeNull();

    const other = (first(ed.view!.state.doc, (m) => m.kind === 'textbox').node.attrs.shape as ShapeModel).key;
    const layer = document.createElement('div');
    const boom = () => {
      throw new Error('boom');
    };
    expect(() => new ShapeTextSession({ view: ed.view!, layer, plugins: boom, nodeViews: {}, attributes: () => ({}), dispatch: () => {}, onChange: () => {}, onClose: () => {} }, other, 0)).toThrow('boom');
    expect(layer.children).toHaveLength(0);
    expect(ed.closeShapeSession()).toBe(false);
    ed.destroy();
    host.remove();
  });
});
