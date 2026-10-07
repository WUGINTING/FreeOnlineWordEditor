// Working with shapes in the editor (phase 8B): select, move, nudge, resize, recolour, wrap,
// restack, align, distribute, group, insert, connect, copy and paste, delete; refused while
// 追蹤修訂 is on except inserting and deleting, which are recorded as Word does.
// jsdom has no layout: a shape's box is laid out here from its model (its offsets from its
// references, which start at 0, 0), so moves and sizes can be followed in px.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { NodeSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { measureFrame, placeShapes, pastedShape } from '../../src/papyrus/editor/shapeView';
import { TRACKING_REFUSED, resizeBox } from '../../src/papyrus/editor/shapeInteract';
import { listClipboardSerializer } from '../../src/papyrus/editor/pasteLists';
import { SHAPE_KINDS } from '../../src/papyrus/docx/shapeOps';
import { schema } from '../../src/papyrus/editor/schema';
import type { ShapeModel, SpNode } from '../../src/papyrus/docx/shapes';

const PX = 9525;
const CM = 360000;

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
  // A shape's box where its model puts it: its offsets from references at 0, 0.
  const base = Element.prototype.getBoundingClientRect;
  Element.prototype.getBoundingClientRect = function (this: Element) {
    if (this.classList?.contains('dx-shape-box')) {
      // Where placeShapes put it (its references start at 0, 0 of the page), as big as drawn.
      const st = (this as HTMLElement).style;
      const left = parseFloat(st.left) || 0;
      const top = parseFloat(st.top) || 0;
      const w = parseFloat(st.width) || 0;
      const h = parseFloat(st.height) || 0;
      return { x: left, y: top, left, top, right: left + w, bottom: top + h, width: w, height: h, toJSON() {} } as DOMRect;
    }
    return base.call(this);
  };
});

let editor: DocxEditor | null = null;
let host: HTMLElement | null = null;
const notices: string[] = [];
afterEach(() => {
  editor?.destroy();
  host?.remove();
  editor = null;
  notices.length = 0;
});

async function open(name: string): Promise<DocxEditor> {
  host = document.createElement('div');
  document.body.append(host);
  editor = new DocxEditor(host, { onNotice: (m) => notices.push(m) });
  await editor.open(readFileSync(`tests/fixtures/shapes/${name}.docx`));
  placeShapes(editor.view!.dom);
  return editor;
}

function shapes(ed: DocxEditor): { pos: number; node: PMNode; shape: ShapeModel }[] {
  const out: { pos: number; node: PMNode; shape: ShapeModel }[] = [];
  ed.view!.state.doc.descendants((n, pos) => {
    if (n.attrs?.shape) out.push({ pos, node: n, shape: n.attrs.shape });
    return true;
  });
  return out;
}
const byName = (ed: DocxEditor, name: string) => shapes(ed).find((s) => s.shape.name === name)!;

/** Select shapes by name (the first is the NodeSelection). */
function select(ed: DocxEditor, ...names: string[]) {
  const view = ed.view!;
  const keys = names.map((n) => byName(ed, n).shape.key);
  view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, byName(ed, names[0]).pos)).setMeta('dx-shapes-select', { keys, kid: null }));
  placeShapes(view.dom);
}

async function saved(ed: DocxEditor): Promise<string> {
  return (await JSZip.loadAsync(await ed.save())).file('word/document.xml')!.async('string');
}

const live = () => host!.querySelector('.dx-live')!;
const wait = () => new Promise((r) => setTimeout(r, 50));

describe('selecting', () => {
  it('選取圖形 / Tab go from shape to shape; the screen reader hears which and how to go on', async () => {
    const ed = await open('shapes');
    expect(ed.selectShape(1)).toBe(true);
    placeShapes(ed.view!.dom);
    const first = ed.shapes.info()!;
    await wait();
    expect(live().textContent).toContain('已選取');
    expect(live().textContent).toContain('Esc 回到文字');
    ed.view!.someProp('handleKeyDown', (f) => f(ed.view!, new KeyboardEvent('keydown', { key: 'Tab' })));
    placeShapes(ed.view!.dom);
    expect(ed.shapes.info()!.name).not.toBe(first.name);
    ed.view!.someProp('handleKeyDown', (f) => f(ed.view!, new KeyboardEvent('keydown', { key: 'Escape' })));
    expect(ed.shapes.info()).toBeNull();
    expect(ed.view!.state.selection).not.toBeInstanceOf(NodeSelection);
  });
});

describe('moving and resizing', () => {
  it('moves by dragging (one undo step, back to the XML as read), nudges with the arrow keys', async () => {
    const ed = await open('shapes');
    const before = byName(ed, '矩形 1');
    select(ed, '矩形 1');
    expect(ed.shapes.moveBy(ed.view!, 20, 10)).toBe(true);
    const moved = byName(ed, '矩形 1').shape;
    expect(moved.anchor!.h.offset).toBe(before.shape.anchor!.h.offset! + 20 * PX);
    expect(moved.anchor!.v.offset).toBe(before.shape.anchor!.v.offset! + 10 * PX);
    expect(ed.view!.state.selection).toBeInstanceOf(NodeSelection);
    const xml = await saved(ed);
    expect(xml).toContain(`<wp:posOffset>${before.shape.anchor!.h.offset! + 20 * PX}</wp:posOffset>`);
    // Arrow: 0.2 cm; Ctrl+arrow: 0.02 cm.
    placeShapes(ed.view!.dom);
    ed.view!.someProp('handleKeyDown', (f) => f(ed.view!, new KeyboardEvent('keydown', { key: 'ArrowRight' })));
    expect(byName(ed, '矩形 1').shape.anchor!.h.offset).toBeCloseTo(moved.anchor!.h.offset! + 0.2 * CM, -2);
    placeShapes(ed.view!.dom);
    ed.view!.someProp('handleKeyDown', (f) => f(ed.view!, new KeyboardEvent('keydown', { key: 'ArrowDown', ctrlKey: true })));
    expect(byName(ed, '矩形 1').shape.anchor!.v.offset).toBeCloseTo(moved.anchor!.v.offset! + 0.02 * CM, -2);
    await wait();
    expect(live().textContent).toMatch(/已移動圖形：水平 [\d.]+ 公分，垂直 [\d.]+ 公分。/);
    ed.undo();
    ed.undo();
    ed.undo();
    expect(byName(ed, '矩形 1').shape).toBe(before.shape);
    expect(ed.isModified()).toBe(false);
  });

  it('resizes from a handle (Shift keeps the proportions) and from the ribbon in cm', async () => {
    const ed = await open('shapes');
    select(ed, '菱形 4');
    const box = { left: 0, top: 0, right: 100, bottom: 50 };
    expect(resizeBox(box, 'se', 20, 30, false)).toEqual({ left: 0, top: 0, right: 120, bottom: 80 });
    expect(resizeBox(box, 'se', 20, 0, true)).toEqual({ left: 0, top: 0, right: 120, bottom: 60 });
    expect(resizeBox(box, 'nw', 10, 10, false)).toEqual({ left: 10, top: 10, right: 100, bottom: 50 });
    const p = byName(ed, '菱形 4').shape;
    const r = ed.view!.dom.querySelector(`.dx-shape[data-shape-key="${p.key}"] > .dx-shape-box`)!.getBoundingClientRect();
    const left = r.left;
    const top = r.top;
    expect(ed.shapes.resizeTo(ed.view!, { left: left - 10, top, right: left + 150, bottom: top + 80 })).toBe(true);
    const m = byName(ed, '菱形 4').shape;
    expect(m.w).toBe(160 * PX);
    expect(m.h).toBe(80 * PX);
    expect(m.anchor!.h.offset).toBe(p.anchor!.h.offset! - 10 * PX);
    select(ed, '菱形 4');
    expect(ed.shapes.setGeometry({ w: 4, h: 2, x: 1, y: 0.5 })).toBe(true);
    expect(byName(ed, '菱形 4').shape).toMatchObject({ w: 4 * CM, h: 2 * CM, anchor: { h: { offset: 1 * CM }, v: { offset: 0.5 * CM } } });
    placeShapes(ed.view!.dom);
    const info = ed.shapes.info()!;
    expect(info).toMatchObject({ w: 4, h: 2, x: 1, y: 0.5, xFrom: '欄', yFrom: '段落' });
  });
});

describe('fill, outline, wrapping, stacking', () => {
  it('changes colours and 文繞圖, and restacks', async () => {
    const ed = await open('shapes');
    select(ed, '橢圓 3');
    ed.shapes.setFill('#ffc000');
    ed.shapes.setOutline('#c00000', 2.25);
    let root = byName(ed, '橢圓 3').shape.root as SpNode;
    expect(root.fill).toEqual({ color: 'FFC000' });
    expect(root.line).toMatchObject({ color: 'C00000', width: 2.25 * 12700 });
    ed.shapes.setFill(null);
    root = byName(ed, '橢圓 3').shape.root as SpNode;
    expect(root.fill).toBeNull();
    for (const [choice, expected] of [['square', { wrap: 'square', behind: false }], ['topAndBottom', { wrap: 'topAndBottom' }], ['behind', { wrap: 'none', behind: true }], ['front', { wrap: 'none', behind: false }]] as const) {
      select(ed, '橢圓 3');
      ed.shapes.setWrap(choice);
      expect(byName(ed, '橢圓 3').shape.anchor, choice).toMatchObject(expected);
    }
    select(ed, '橢圓 3');
    ed.shapes.setWrap('inline');
    expect(byName(ed, '橢圓 3').shape.anchor).toBeNull();
    select(ed, '橢圓 3');
    ed.shapes.setWrap('front');
    expect(byName(ed, '橢圓 3').shape.anchor).toMatchObject({ wrap: 'none', h: { rel: 'column' }, v: { rel: 'paragraph' } });

    select(ed, '矩形 1');
    const z = (n: string) => byName(ed, n).shape.anchor!.z;
    const below = z('矩形 1');
    ed.shapes.restack(1);
    expect(z('矩形 1')).toBeGreaterThan(below);
    const xml = await saved(ed);
    expect(xml).toContain(`relativeHeight="${z('矩形 1')}"`);
  });
});

describe('align, distribute, group', () => {
  it('aligns and distributes the selected shapes', async () => {
    const ed = await open('shapes');
    select(ed, '矩形 1', '平行四邊形 5', '矩形 10');
    ed.shapes.align('left');
    const lefts = ['矩形 1', '平行四邊形 5', '矩形 10'].map((n) => byName(ed, n).shape.anchor!.h.offset);
    expect(new Set(lefts).size).toBe(1);
    select(ed, '矩形 1', '平行四邊形 5', '矩形 10');
    ed.shapes.align('top');
    expect(new Set(['矩形 1', '平行四邊形 5', '矩形 10'].map((n) => byName(ed, n).shape.anchor!.v.offset)).size).toBe(1);
    // Spread out: left edges 0, 110 pt, 220 pt apart, then distributed evenly.
    select(ed, '矩形 1');
    ed.shapes.setGeometry({ x: 0 });
    select(ed, '平行四邊形 5');
    ed.shapes.setGeometry({ x: 2 });
    select(ed, '矩形 10');
    ed.shapes.setGeometry({ x: 10 });
    select(ed, '矩形 1', '平行四邊形 5', '矩形 10');
    ed.shapes.distribute('h');
    const [a, b, c] = ['矩形 1', '平行四邊形 5', '矩形 10'].map((n) => byName(ed, n).shape);
    const gap1 = b.anchor!.h.offset! - (a.anchor!.h.offset! + a.w);
    const gap2 = c.anchor!.h.offset! - (b.anchor!.h.offset! + b.w);
    expect(Math.abs(gap1 - gap2)).toBeLessThan(2 * PX);
  });

  it('groups floating shapes and takes the group apart, in place', async () => {
    const ed = await open('shapes');
    const count = shapes(ed).length;
    select(ed, '矩形 1', '橢圓 3');
    expect(ed.shapes.info()!.canGroup).toBe(true);
    expect(ed.shapes.group()).toBe(true);
    expect(shapes(ed)).toHaveLength(count - 1);
    const g = shapes(ed).find((s) => s.shape.kind === 'group')!.shape;
    expect((g.root as { kids: SpNode[] }).kids.map((k) => k.geom)).toEqual(['rect', 'ellipse']);
    expect(ed.shapes.info()!.canUngroup).toBe(true);
    const xml = await saved(ed);
    expect(xml).toContain('<mc:Choice Requires="wpg">');
    expect(ed.shapes.ungroup()).toBe(true);
    expect(shapes(ed)).toHaveLength(count);
    expect(shapes(ed).filter((s) => ['rect', 'ellipse'].includes((s.shape.root as SpNode).geom) && s.shape.kind === 'shape').length).toBeGreaterThanOrEqual(2);
    // One undo step each.
    ed.undo();
    ed.undo();
    expect(ed.isModified()).toBe(false);
  });
});

describe('inserting and connecting', () => {
  it('inserts a flowchart shape anchored at the cursor, a text box opened for typing, and a connector attached to both', async () => {
    const ed = await open('textboxes');
    const view = ed.view!;
    view.dispatch(view.state.tr.setSelection(NodeSelection.near(view.state.doc.resolve(1))));
    const kind = (id: string) => SHAPE_KINDS.find((k) => k.id === id)!;
    // Page px from the column's left and the anchor paragraph's top.
    const para = view.dom.querySelector<HTMLElement>('.dx-p')!;
    const f = measureFrame(para).frame;
    const X = f.column.left;
    const Y = f.paragraph.top;
    const P = (x: number, y: number) => ({ x: X + x, y: Y + y });
    expect(ed.shapes.insert(view, kind('process'), P(100, 200), P(200, 250))).toBe(true);
    const a = ed.shapes.info()!;
    expect(a).toMatchObject({ w: +(100 * PX / CM).toFixed(2), wrap: 'front' });
    const proc = shapes(ed).find((s) => (s.shape.root as SpNode).geom === 'flowChartProcess')!.shape;
    expect(proc.anchor).toMatchObject({ h: { rel: 'column', offset: 100 * PX }, v: { rel: 'paragraph', offset: 200 * PX } });
    expect(proc.docId).toBeGreaterThan(0);
    placeShapes(view.dom);
    ed.shapes.insert(view, kind('decision'), P(300, 200), P(400, 260));
    placeShapes(view.dom);
    const dec = shapes(ed).find((s) => (s.shape.root as SpNode).geom === 'flowChartDecision')!.shape;
    // A connector from the process's right site to the decision's left site.
    ed.shapes.insert(view, kind('arrowLine'), P(201, 225), P(299, 230));
    const line = shapes(ed).find((s) => (s.shape.root as SpNode).geom === 'straightConnector1')!.shape;
    expect((line.root as SpNode).st).toEqual({ id: proc.docId, idx: 3 });
    expect((line.root as SpNode).end).toEqual({ id: dec.docId, idx: 1 });
    // Moving the process takes the connector along.
    placeShapes(view.dom);
    select(ed, proc.name!);
    ed.shapes.moveBy(view, 0, 40);
    const after = shapes(ed).find((s) => (s.shape.root as SpNode).geom === 'straightConnector1')!.shape;
    expect(after.anchor!.v.offset).toBe(230 * PX);
    expect(after.h).toBe(35 * PX);
    expect((after.root as SpNode).flipV).toBe(true);
    // A text box opens for typing.
    ed.shapes.insert(view, kind('textbox'), P(50, 400), null);
    expect(ed.target).toBe('textbox');
    ed.run((s, d) => (d?.(s.tr.insertText('新方塊')), true));
    ed.closeShapeText(false);
    const xml = await saved(ed);
    expect(xml.match(/新方塊/g)).toHaveLength(2);
    // Unique drawing ids.
    const ids = [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe('追蹤修訂', () => {
  it('refuses moving and changing shapes; inserting and deleting are recorded (w:ins / w:del around the run)', async () => {
    const ed = await open('shapes');
    ed.setTrackChanges(true);
    select(ed, '矩形 1');
    expect(ed.shapes.moveBy(ed.view!, 10, 0)).toBe(false);
    expect(ed.shapes.setFill('#ff0000')).toBe(false);
    expect(notices).toContain(`${TRACKING_REFUSED}。`);
    const view = ed.view!;
    ed.shapes.insert(view, SHAPE_KINDS[0], { x: 10, y: 10 }, null);
    let xml = await saved(ed);
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="[^"]*" w:date="[^"]*"><w:r><w:rPr><w:noProof\/><\/w:rPr><mc:AlternateContent>/);
    select(ed, '矩形 1');
    ed.shapes.deleteSelected();
    xml = await saved(ed);
    const del = xml.indexOf('<w:del ');
    expect(del).toBeGreaterThan(0);
    expect(xml.slice(del, xml.indexOf('</w:del>', del))).toContain('name="矩形 1"');
  });
});

describe('Ctrl+drag', () => {
  it('leaves the shape and puts a copy of its own where it is dropped', async () => {
    const ed = await open('shapes');
    const before = byName(ed, '矩形 1').shape;
    select(ed, '矩形 1');
    expect(ed.shapes.duplicate(ed.view!, 40, 0)).toBe(true);
    const both = shapes(ed).filter((s) => s.shape.name === '矩形 1');
    expect(both).toHaveLength(2);
    expect(both.map((s) => s.shape.anchor!.h.offset).sort((a, b) => a! - b!)).toEqual([before.anchor!.h.offset, before.anchor!.h.offset! + 40 * PX]);
    expect(new Set(both.map((s) => s.shape.docId)).size).toBe(2);
    expect(new Set(both.map((s) => s.shape.key)).size).toBe(2);
    const xml = await saved(ed);
    const ids = [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe('copy and paste', () => {
  it('copies a shape as its run and pastes it back as a shape of its own; refuses anything pointing into a package', async () => {
    const ed = await open('shapes');
    const src = byName(ed, '矩形 1');
    const out = listClipboardSerializer(() => ed.model.numbering).serializeNode(src.node) as HTMLElement;
    const dom = (out.matches('[data-dx-shape-run]') ? out : out.querySelector('[data-dx-shape-run]'))!;
    expect(dom.getAttribute('data-dx-shape-run')).toBe(src.node.attrs.xml);
    const attrs = pastedShape(src.node.attrs.xml)!;
    expect((attrs.shape as ShapeModel).root).toMatchObject({ geom: 'rect' });
    // Pasted into the document: new key and drawing id.
    const view = ed.view!;
    view.dispatch(view.state.tr.insert(1, schema.nodes.raw_inline.create(attrs)));
    const copies = shapes(ed).filter((s) => s.shape.name === '矩形 1');
    expect(copies).toHaveLength(2);
    expect(copies[0].shape.key).not.toBe(copies[1].shape.key);
    expect(copies[0].shape.docId).not.toBe(copies[1].shape.docId);
    expect(pastedShape('<w:r><w:drawing><wp:inline><a:blip r:embed="rId5"/></wp:inline></w:drawing></w:r>')).toBeNull();
    expect(pastedShape('<w:p/>')).toBeNull();
    expect(pastedShape('not xml')).toBeNull();
  });
});

describe('anchoring as Word does', () => {
  /** Where posAtCoords "finds" the pointer: in the paragraph with this text. */
  function pointAt(ed: DocxEditor, text: string) {
    let pos = -1;
    ed.view!.state.doc.descendants((n, p) => {
      if (pos < 0 && n.type.name === 'paragraph' && n.textContent.startsWith(text)) pos = p + 1;
      return pos < 0;
    });
    ed.view!.posAtCoords = () => ({ pos, inside: -1 });
    return pos;
  }
  const paragraphOf = (ed: DocxEditor, name: string) => {
    const s = byName(ed, name);
    return ed.view!.state.doc.resolve(s.pos).parent.textContent;
  };

  it('a new shape is anchored to the paragraph where it is put, at its start', async () => {
    const ed = await open('shapes');
    const view = ed.view!;
    const start = pointAt(ed, '上及下圖案後的文字');
    ed.shapes.insert(view, SHAPE_KINDS[0], { x: 100, y: 300 }, { x: 200, y: 340 });
    const made = shapes(ed).find((s) => (s.shape.root as SpNode).geom === 'flowChartProcess')!;
    expect(made.pos).toBe(start);
    expect(view.state.doc.resolve(made.pos).parent.textContent).toContain('上及下圖案後的文字');
    expect(made.shape.anchor).toMatchObject({ v: { rel: 'paragraph' }, h: { rel: 'column' } });
  });

  it('a shape dragged elsewhere takes the nearest paragraph as its anchor (one undo step); a locked anchor stays', async () => {
    const ed = await open('shapes');
    const before = paragraphOf(ed, '矩形 1');
    select(ed, '矩形 1');
    pointAt(ed, '結尾');
    ed.shapes.moveBy(ed.view!, 0, 200);
    expect(paragraphOf(ed, '矩形 1')).toContain('結尾');
    expect(byName(ed, '矩形 1').shape.anchor!.v.rel).toBe('paragraph');
    expect(shapes(ed).filter((s) => s.shape.name === '矩形 1')).toHaveLength(1);
    const xml = await saved(ed);
    const at = xml.indexOf('name="矩形 1"');
    expect(xml.lastIndexOf('<w:p ', at)).toBeGreaterThan(xml.lastIndexOf('上及下圖案後的文字', at));
    ed.undo();
    expect(paragraphOf(ed, '矩形 1')).toBe(before);
    expect(ed.isModified()).toBe(false);

    // Locked (鎖定錨點): moved, not re-anchored.
    const s = byName(ed, '矩形 1');
    ed.view!.dispatch(ed.view!.state.tr.setNodeMarkup(s.pos, undefined, { ...s.node.attrs, shape: { ...s.shape, anchor: { ...s.shape.anchor!, locked: true } } }));
    select(ed, '矩形 1');
    pointAt(ed, '結尾');
    ed.shapes.moveBy(ed.view!, 0, 200);
    expect(paragraphOf(ed, '矩形 1')).toBe(before);
  });
});
