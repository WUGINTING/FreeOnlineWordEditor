// Changing shapes and writing them back (phase 8B): new shapes as Word writes them, the changes
// of a shape (position, size, fill, line, wrapping, stacking, the shapes on a canvas and their
// connectors) patched into both of its copies, groups made and taken apart.
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { schema } from '../../src/papyrus/editor/schema';
import { connectionSites, connectorEnds, routeConnector } from '../../src/papyrus/docx/shapeGeometry';
import {
  SHAPE_KINDS, changedKid, connectorEndsOf, filled, groupXml, moved, newShapeXml, outlined, readRunShape, resized, siteAt, stacked, ungroupXml, wrapped,
} from '../../src/papyrus/docx/shapeOps';
import { geomKey, hasChangedShape, patchShapeXml, svgToVml } from '../../src/papyrus/docx/shapeWrite';
import type { ShapeModel, SpNode } from '../../src/papyrus/docx/shapes';

const kind = (id: string) => SHAPE_KINDS.find((k) => k.id === id)!;
const cm = 360000;

/** The first shape of a document whose model passes `test`. */
function first(doc: PMNode, test: (m: ShapeModel) => boolean): { pos: number; node: PMNode; shape: ShapeModel } {
  let found: { pos: number; node: PMNode; shape: ShapeModel } | null = null;
  doc.descendants((n, pos) => {
    if (!found && n.attrs?.shape && test(n.attrs.shape)) found = { pos, node: n, shape: n.attrs.shape };
    return !found;
  });
  return found!;
}

/** The document with one shape's model replaced, saved and read again. */
async function saveWith(file: string, test: (m: ShapeModel) => boolean, change: (m: ShapeModel) => ShapeModel) {
  const { doc, model } = await readDocx(readFileSync(`tests/fixtures/shapes/${file}.docx`));
  const f = first(doc, test);
  const state = EditorState.create({ schema, doc });
  const next = state.apply(state.tr.setNodeMarkup(f.pos, undefined, { ...f.node.attrs, shape: change(f.shape) })).doc;
  const bytes = await writeDocx(next, model);
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
  const reread = await readDocx(bytes);
  return { xml, doc: reread.doc, before: f.shape, name: f.shape.name! };
}

const byName = (doc: PMNode, name: string) => first(doc, (m) => m.name === name).shape;

describe('new shapes', () => {
  it('are written as Word 2010+ does (wps in mc:Choice, VML in mc:Fallback) and read back the same', () => {
    for (const k of SHAPE_KINDS) {
      const xml = newShapeXml({ kind: k, x: 1 * cm, y: 2 * cm, w: 3 * cm, h: k.line ? 1 * cm : 1.5 * cm, z: 251659300, docId: 123456789, name: `${k.label} 1` });
      expect(xml).toMatch(/^<w:r><w:rPr><w:noProof\/><\/w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor /);
      expect(xml).toContain('<mc:Fallback><w:pict>');
      const m = readRunShape(xml)!;
      expect(m, k.id).toMatchObject({ kind: k.textbox ? 'textbox' : 'shape', docId: 123456789, w: 3 * cm, drawable: true });
      expect(m.anchor).toMatchObject({ h: { rel: 'column', offset: 1 * cm }, v: { rel: 'paragraph', offset: 2 * cm }, wrap: 'none', behind: false, z: 251659300 });
      const root = m.root as SpNode;
      expect(root.geom).toBe(k.prst);
      if (k.line) {
        expect(root.cxn).toBe(true);
        expect(root.fill).toBeNull();
        expect(!!root.line?.tail).toBe(!!k.arrow);
      } else {
        // Room to type in, centred in a shape.
        expect(root.text?.i).toBe(0);
        expect(root.text?.anchor).toBe(k.textbox ? 't' : 'ctr');
      }
    }
  });

  it('VML paths from the drawn outlines (lines, curves, arcs)', () => {
    expect(svgToVml('M0,0L10,0L10,10Z', true)).toBe('m0,0l10,0l10,10xe');
    expect(svgToVml('M0,0H10V5', false)).toBe('nfm0,0l10,0l10,5e');
    expect(svgToVml('M0,5A5,5 0 1 0 10,5', true)).toMatch(/^m0,5c.*c.*e$/);
  });
});

describe('a changed shape', () => {
  it('moved: wp:posOffset and the VML margins; everything else as read', async () => {
    const { xml, doc, before, name } = await saveWith('textboxes', (m) => m.kind === 'textbox' && m.anchor?.wrap === 'none', (m) => moved(m, 2 * cm, 3 * cm));
    const after = byName(doc, name);
    expect(after.anchor).toMatchObject({ h: { rel: before.anchor!.h.rel, offset: 2 * cm }, v: { rel: before.anchor!.v.rel, offset: 3 * cm } });
    expect(after.w).toBe(before.w);
    const run = xml.slice(xml.lastIndexOf('<w:r>', xml.indexOf(`name="${name}"`)));
    expect(run).toMatch(/margin-left:56\.69pt/);
    expect(run).toMatch(/margin-top:85\.04pt/);
  });

  it('resized, filled, outlined: extent, a:xfrm and VML size and colours', async () => {
    const { xml, doc, name } = await saveWith('shapes', (m) => (m.root as SpNode).geom === 'diamond', (m) => outlined(filled(resized(m, 4 * cm, 2 * cm), 'FFC000'), 'C00000', 25400));
    const after = byName(doc, name);
    expect(after).toMatchObject({ w: 4 * cm, h: 2 * cm, root: { w: 4 * cm, h: 2 * cm, fill: { color: 'FFC000' }, line: { color: 'C00000', width: 25400 } } });
    const at = xml.indexOf(`name="${name}"`);
    const run = xml.slice(xml.lastIndexOf('<w:r>', at), xml.indexOf('</mc:AlternateContent>', at));
    expect(run).toContain(`<wp:extent cx="${4 * cm}" cy="${2 * cm}"/>`);
    expect(run).toContain(`<a:ext cx="${4 * cm}" cy="${2 * cm}"/>`);
    expect(run).toContain('fillcolor="#FFC000"');
    expect(run).toContain('strokecolor="#C00000"');
    expect(run).toMatch(/width:113\.39pt;height:56\.69pt/);
  });

  it('wrapped otherwise: square, top and bottom, behind, in the line of text and out again', async () => {
    let r = await saveWith('shapes', (m) => (m.root as SpNode).geom === 'ellipse' && m.kind === 'shape', (m) => wrapped(m, 'square'));
    expect(byName(r.doc, r.name).anchor).toMatchObject({ wrap: 'square', behind: false });
    expect(r.xml).toContain('<wp:wrapSquare wrapText="bothSides"/>');
    r = await saveWith('shapes', (m) => (m.root as SpNode).geom === 'ellipse' && m.kind === 'shape', (m) => wrapped(m, 'behind'));
    expect(byName(r.doc, r.name).anchor).toMatchObject({ wrap: 'none', behind: true });
    r = await saveWith('shapes', (m) => (m.root as SpNode).geom === 'ellipse' && m.kind === 'shape', (m) => wrapped(m, 'inline'));
    const inline = byName(r.doc, r.name);
    expect(inline.anchor).toBeNull();
    const at = r.xml.indexOf(`name="${r.name}"`);
    const run = r.xml.slice(r.xml.lastIndexOf('<w:r>', at), r.xml.indexOf('</mc:AlternateContent>', at));
    expect(run).toMatch(/<wp:inline [^>]*><wp:extent/);
    expect(run).not.toMatch(/position:absolute/);
    // And out of the line again, placed where it was.
    const back = wrapped(inline, 'topAndBottom', { h: 1 * cm, v: 0 }, 251660000);
    expect(back.anchor).toMatchObject({ wrap: 'topAndBottom', h: { offset: 1 * cm }, z: 251660000 });
  });

  it('restacked: relativeHeight and the VML z-index', async () => {
    const { xml, doc, name } = await saveWith('shapes', (m) => (m.root as SpNode).geom === 'rect' && m.kind === 'shape' && !m.anchor?.behind, (m) => stacked(m, 251700000));
    expect(byName(doc, name).anchor!.z).toBe(251700000);
    expect(xml).toContain('relativeHeight="251700000"');
    expect(xml).toContain('z-index:251700000');
  });

  it('back as it was: the XML as read', async () => {
    const { doc } = await readDocx(readFileSync('tests/fixtures/shapes/shapes.docx'));
    const f = first(doc, (m) => m.kind === 'shape');
    const there = moved(f.shape, 5 * cm, 5 * cm);
    expect(hasChangedShape(there)).toBe(true);
    const back = moved(there, f.shape.anchor!.h.offset!, f.shape.anchor!.v.offset!);
    expect(hasChangedShape(back)).toBe(false);
    expect(geomKey(back)).toBe(back.base);
  });
});

describe('the shapes on a canvas and their connectors', () => {
  it('a shape moved: the connectors attached to it follow, site to site', async () => {
    const { xml, doc, name } = await saveWith('flowchart', (m) => m.kind === 'canvas', (m) =>
      changedKid(m, 1, (k) => ({ ...k, x: k.x + 0.5 * cm, y: k.y + 0.3 * cm })),
    );
    const canvas = byName(doc, name);
    const kids = (canvas.root as { kids: SpNode[] }).kids;
    const shape = kids[1]; // 收文登錄
    for (const c of kids.filter((k) => k.cxn && (k.st?.id === shape.id || k.end?.id === shape.id))) {
      const ends = connectorEndsOf(c);
      const at = (conn: { id: number; idx: number }) => siteAt(kids.find((k) => k.id === conn.id)!, conn.idx)!;
      const a = at(c.st!);
      const b = at(c.end!);
      expect(Math.abs(ends.a.x - a.x) + Math.abs(ends.a.y - a.y)).toBeLessThan(2);
      expect(Math.abs(ends.b.x - b.x) + Math.abs(ends.b.y - b.y)).toBeLessThan(2);
    }
    // The VML copy of the canvas is written again from the model, texts kept.
    const run = xml.slice(xml.indexOf('<mc:Fallback>', xml.indexOf(`name="${name}"`)));
    expect(run.slice(0, run.indexOf('</mc:Fallback>'))).toContain('收文登錄');
  });

  it('routes connectors as Word does: a bent one leaving a site downwards is turned', () => {
    const box = routeConnector('bentConnector3', { x: 100, y: 100 }, { x: 40, y: 200 }, 90);
    expect(box.rot === 90 || box.rot === 270).toBe(true);
    const ends = connectorEnds(box);
    expect(ends.a.x).toBeCloseTo(100);
    expect(ends.a.y).toBeCloseTo(100);
    expect(ends.b.x).toBeCloseTo(40);
    expect(ends.b.y).toBeCloseTo(200);
    const straight = routeConnector('straightConnector1', { x: 10, y: 50 }, { x: 5, y: 10 });
    expect(straight).toEqual({ x: 5, y: 10, w: 5, h: 40, flipH: true, flipV: true });
    expect(connectionSites('flowChartDecision', 100, 50).map((s) => [s.x, s.y])).toEqual([[50, 0], [0, 25], [50, 50], [100, 25]]);
    expect(connectionSites('ellipse', 100, 100)).toHaveLength(8);
  });
});

describe('groups', () => {
  it('two shapes grouped (their text and styles kept) and taken apart again, in place', () => {
    const a = newShapeXml({ kind: kind('process'), x: 1 * cm, y: 1 * cm, w: 3 * cm, h: 1.5 * cm, z: 5, docId: 111, name: '程序 1' });
    const b = newShapeXml({ kind: kind('decision'), x: 6 * cm, y: 2 * cm, w: 3.5 * cm, h: 2 * cm, z: 6, docId: 222, name: '決策 2' });
    const ma = readRunShape(a)!;
    const mb = readRunShape(b)!;
    const xml = groupXml([{ model: ma, x: 1 * cm, y: 1 * cm, xml: a }, { model: mb, x: 6 * cm, y: 2 * cm, xml: b }], 333, [444, 555], () => '<w:p/>');
    expect(xml).toContain('<mc:Choice Requires="wpg">');
    const g = readRunShape(xml)!;
    expect(g).toMatchObject({ kind: 'group', docId: 333, w: 8.5 * cm, h: 3 * cm, anchor: { h: { offset: 1 * cm }, v: { offset: 1 * cm }, z: 6 } });
    const kids = (g.root as { kids: SpNode[] }).kids;
    expect(kids.map((k) => [k.geom, k.x, k.y, k.w, k.h, k.id])).toEqual([
      ['flowChartProcess', 0, 0, 3 * cm, 1.5 * cm, 444],
      ['flowChartDecision', 5 * cm, 1 * cm, 3.5 * cm, 2 * cm, 555],
    ]);
    expect(xml).toContain('<v:group');
    const parts = ungroupXml(xml, g, [666, 777])!;
    expect(parts.map((p) => readRunShape(p.xml)!).map((m) => [(m.root as SpNode).geom, m.anchor!.h.offset, m.anchor!.v.offset, m.w, m.h, m.docId])).toEqual([
      ['flowChartProcess', 1 * cm, 1 * cm, 3 * cm, 1.5 * cm, 666],
      ['flowChartDecision', 6 * cm, 2 * cm, 3.5 * cm, 2 * cm, 777],
    ]);
  });

  it('a structural change is not patched (it makes its XML anew)', () => {
    const a = newShapeXml({ kind: kind('process'), x: 0, y: 0, w: cm, h: cm, z: 1, docId: 1, name: 'a' });
    const b = newShapeXml({ kind: kind('process'), x: 0, y: 0, w: cm, h: cm, z: 1, docId: 2, name: 'b' });
    const g = readRunShape(groupXml([{ model: readRunShape(a)!, x: 0, y: 0, xml: a }, { model: readRunShape(b)!, x: cm, y: 0, xml: b }], 3, [4, 5], () => '<w:p/>'))!;
    const fewer = { ...g, root: { ...(g.root as { kids: SpNode[] }), kids: [(g.root as { kids: SpNode[] }).kids[0]] } } as ShapeModel;
    expect(() => patchShapeXml(a, g, fewer)).toThrow();
  });
});
