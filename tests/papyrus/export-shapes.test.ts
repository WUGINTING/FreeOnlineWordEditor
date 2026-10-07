// Writes the shape fixtures (tests/fixtures/shapes) changed in the editor, for checking in
// Microsoft Word (scripts/check-shapes-in-word.ps1):
//  - 8A: text box text edited; shapes, positions, sizes, colours, wrapping and connectors stay;
//  - 8B: shapes moved, resized, recoloured, wrapped, restacked, grouped, deleted, a canvas's shape
//    moved with its connectors, new shapes, connectors and a text box inserted, a shape pasted:
//    Word must show them as the editor has them, in the DrawingML copy and in the VML copy
//    (the -vml.docx files, as Word 2007 reads them).
// Opt-in: DOCX_EXPORT=<out> npx vitest run tests/papyrus/export-shapes.test.ts
//         pwsh scripts/check-shapes-in-word.ps1 -Folder <out>
import { beforeAll, describe, it } from 'vitest';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import JSZip from 'jszip';
import { closeHistory } from 'prosemirror-history';
import { NodeSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { DocxEditor } from '../../src/papyrus/editor/core';
import { textFrames, type ShapeModel, type SpNode } from '../../src/papyrus/docx/shapes';
import { schema } from '../../src/papyrus/editor/schema';
import { measureFrame, placeShapes, pastedShape } from '../../src/papyrus/editor/shapeView';
import { SHAPE_KINDS } from '../../src/papyrus/docx/shapeOps';

const out = process.env.DOCX_EXPORT;

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
  // No layout in jsdom: a shape's box is where placeShapes put it (see shapeTools.test.ts).
  const base = Element.prototype.getBoundingClientRect;
  Element.prototype.getBoundingClientRect = function (this: Element) {
    if (this.classList?.contains('dx-shape-box')) {
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

/** The text box (in `view`'s document) whose text is `text`. */
function boxWith(view: EditorView, text: string): { key: string; index: number } | null {
  let hit: { key: string; index: number } | null = null;
  view.state.doc.descendants((n) => {
    const shape = n.attrs?.shape as ShapeModel | undefined;
    for (const f of shape ? textFrames(shape) : []) {
      if (!hit && schema.nodeFromJSON(f.doc).textContent === text) hit = { key: shape!.key, index: f.i };
    }
    return !hit;
  });
  return hit;
}

/** Edits per file: text box text → appended text; `header` ones in the default header. */
const EDITS: Record<string, { text: string; add: string; header?: boolean }[]> = {
  flowchart: [{ text: '收文登錄', add: '（電子）' }, { text: '資料齊全？', add: '是' }],
  textboxes: [{ text: '繞圖文字方塊', add: '（已改）' }, { text: '同列方塊', add: '改' }, { text: '頁首文字方塊', add: '（新）', header: true }],
  shapes: [{ text: '畫布甲', add: '改' }, { text: '上及下', add: '（改）' }],
};

async function bytesOf(editor: DocxEditor): Promise<Uint8Array> {
  // (jsdom's Blob can't be read directly: the package is written out again as it is.)
  return (await JSZip.loadAsync(await editor.save())).generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe.skipIf(!out)('shape fixtures with edited text boxes, for Word', () => {
  it('writes them', async () => {
    mkdirSync(out!, { recursive: true });
    const index: { original: string; copy: string; texts: Record<string, string> }[] = [];
    for (const [name, edits] of Object.entries(EDITS)) {
      const host = document.createElement('div');
      document.body.append(host);
      const editor = new DocxEditor(host);
      const original = resolve(`tests/fixtures/shapes/${name}.docx`);
      await editor.open(readFileSync(original));
      const texts: Record<string, string> = {};
      for (const e of edits) {
        if (e.header) editor.editHeaderFooter('header');
        const box = boxWith(editor.activeView!, e.text);
        if (!box || !editor.editShapeText(box.key, box.index)) throw new Error(`no text box ${e.text} in ${name}`);
        const v = editor.activeView!;
        v.dispatch(closeHistory(v.state.tr.insertText(e.add, v.state.doc.content.size - 1)));
        editor.closeShapeText(false);
        if (e.header) editor.closeHeaderFooter(false);
        texts[e.text] = e.text + e.add;
      }
      const copy = join(out!, `${name}-edited.docx`);
      writeFileSync(copy, await bytesOf(editor));
      index.push({ original, copy: resolve(copy), texts });
      editor.destroy();
      host.remove();
    }
    writeFileSync(join(out!, 'shapes-index.json'), JSON.stringify(index, null, 2));
  }, 120_000);
});

// ----- phase 8B -----

const PT = (emu: number) => Math.round((emu / 12700) * 10) / 10;
/** Word's WdWrapType for the editor's wrapping. */
const WRAP = { square: 0, topAndBottom: 4, front: 3, behind: 5, inline: 7 } as const;

interface Expect {
  /** Shapes Word must show, by name: position (pt from its references), size (pt), wrapping, text, group members. */
  shapes: { name: string; left?: number; top?: number; width?: number; height?: number; wrap?: number; text?: string; group?: string[]; anchor?: string }[];
  /** Shapes that must be gone. */
  absent: string[];
  /** Pairs [a, b]: a stacked above b. */
  above: [string, string][];
  /** Connectors on canvases attached at both ends. */
  attached: number;
  /** How many shapes the body has. */
  count: number;
}

/** Selecting and noting what Word must show. */
class Ops {
  expect: Expect = { shapes: [], absent: [], above: [], attached: 0, count: 0 };
  constructor(private ed: DocxEditor) {}
  all(): { pos: number; node: PMNode; shape: ShapeModel }[] {
    const list: { pos: number; node: PMNode; shape: ShapeModel }[] = [];
    this.ed.view!.state.doc.descendants((n, pos) => {
      if (n.attrs?.shape) list.push({ pos, node: n, shape: n.attrs.shape });
      return true;
    });
    return list;
  }
  shape(name: string) {
    const s = this.all().find((x) => x.shape.name === name);
    if (!s) throw new Error(`no shape ${name}`);
    return s;
  }
  select(...names: string[]) {
    const view = this.ed.view!;
    placeShapes(view.dom);
    const keys = names.map((n) => this.shape(n).shape.key);
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, this.shape(names[0]).pos)).setMeta('dx-shapes-select', { keys, kid: null }));
    placeShapes(view.dom);
  }
  /** Move a shape on a canvas (its index among the canvas's shapes) by (dx, dy) px. */
  kid(canvas: string, index: number, dx: number, dy: number) {
    const view = this.ed.view!;
    placeShapes(view.dom);
    const c = this.shape(canvas);
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, c.pos)).setMeta('dx-shapes-select', { keys: [c.shape.key], kid: index }));
    placeShapes(view.dom);
    this.ed.shapes.moveBy(view, dx, dy);
  }
  /** Where the pointer is "found" (jsdom has no layout): the paragraph starting with `text`. */
  pointAt(text: string) {
    let pos = -1;
    this.ed.view!.state.doc.descendants((n, p) => {
      if (pos < 0 && n.type.name === 'paragraph' && n.textContent.startsWith(text)) pos = p + 1;
      return pos < 0;
    });
    this.ed.view!.posAtCoords = () => ({ pos, inside: -1 });
  }
  /** What Word must say of a shape as the model has it now (with the paragraph it is anchored to). */
  expectShape(name: string, text?: string, anchor?: string) {
    const m = this.shape(name).shape;
    const e: Expect['shapes'][number] = { name, width: PT(m.w), height: PT(m.h) };
    if (m.anchor) {
      e.left = PT(m.anchor.h.offset ?? 0);
      e.top = PT(m.anchor.v.offset ?? 0);
      e.wrap = m.anchor.behind ? WRAP.behind : m.anchor.wrap === 'none' ? WRAP.front : m.anchor.wrap === 'topAndBottom' ? WRAP.topAndBottom : WRAP.square;
    } else e.wrap = WRAP.inline;
    if (text != null) e.text = text;
    if (anchor != null) e.anchor = anchor;
    this.expect.shapes.push(e);
  }
}

describe.skipIf(!out)('shape fixtures changed in the editor (8B), for Word', () => {
  it('writes them with what Word must show', async () => {
    mkdirSync(out!, { recursive: true });
    const cases: { original: string; copy: string; vml: string; expect: Expect }[] = [];
    const run = async (name: string, ops: (ed: DocxEditor, x: Ops) => void) => {
      const host = document.createElement('div');
      document.body.append(host);
      const editor = new DocxEditor(host);
      const original = resolve(`tests/fixtures/shapes/${name}.docx`);
      await editor.open(readFileSync(original));
      placeShapes(editor.view!.dom);
      const x = new Ops(editor);
      ops(editor, x);
      x.expect.count = x.all().length;
      const bytes = await bytesOf(editor);
      const copy = join(out!, `${name}-8b.docx`);
      writeFileSync(copy, bytes);
      // The same file read by a Word that knows only VML (as Word 2007 does): its mc:Fallback.
      const zip = await JSZip.loadAsync(bytes);
      const doc = await zip.file('word/document.xml')!.async('string');
      zip.file('word/document.xml', doc.replace(/Requires="(wps|wpg|wpc)"/g, 'Requires="pxnone"').replace('<w:document ', '<w:document xmlns:pxnone="urn:px:none" '));
      const vml = join(out!, `${name}-8b-vml.docx`);
      writeFileSync(vml, await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
      cases.push({ original, copy: resolve(copy), vml: resolve(vml), expect: x.expect });
      editor.destroy();
      host.remove();
    };

    await run('shapes', (ed, x) => {
      x.select('矩形 1');
      ed.shapes.moveBy(ed.view!, 30, 20);
      x.select('菱形 4');
      ed.shapes.setGeometry({ w: 4, h: 2 });
      x.select('橢圓 3');
      ed.shapes.setFill('#ffc000');
      ed.shapes.setOutline('#c00000', 2.25);
      x.select('平行四邊形 5');
      ed.shapes.setWrap('square');
      x.select('箭號: 向右 6');
      ed.shapes.setWrap('behind');
      x.select('矩形 10');
      ed.shapes.restack(1);
      x.select('矩形: 圓角 2', '箭號: 左-右雙向 7');
      ed.shapes.group();
      x.select('直線接點 9');
      ed.shapes.deleteSelected();
      x.kid('畫布 13', 0, 0, 20);
      // Dragged far down: anchored to the paragraph it is dropped by (as Word does).
      x.select('矩形 10');
      x.pointAt('結尾');
      ed.shapes.moveBy(ed.view!, 0, 300);
      x.expectShape('矩形 10', undefined, '結尾');
      // A new one put by a paragraph: anchored to it.
      x.pointAt('上及下圖案後的文字');
      ed.shapes.insert(ed.view!, SHAPE_KINDS.find((k) => k.id === 'terminator')!, { x: 150, y: 50 }, { x: 250, y: 80 });
      x.expectShape(x.all().find((m) => (m.shape.root as SpNode).geom === 'flowChartTerminator')!.shape.name!, undefined, '上及下圖案後的文字');
      delete (ed.view as { posAtCoords?: unknown }).posAtCoords;
      for (const n of ['矩形 1', '菱形 4', '橢圓 3', '平行四邊形 5', '箭號: 向右 6']) x.expectShape(n);
      x.expect.shapes.push({ name: '*group*', group: ['矩形: 圓角 2', '箭號: 左-右雙向 7'] });
      x.expect.absent.push('直線接點 9');
      x.expect.above.push(['矩形 10', '星形: 五角 8']);
      x.expect.attached = 1;
    });

    await run('textboxes', (ed, x) => {
      const view = ed.view!;
      view.dispatch(view.state.tr.setSelection(NodeSelection.near(view.state.doc.resolve(1))));
      const f = measureFrame(view.dom.querySelector<HTMLElement>('.dx-p')!).frame;
      const P = (px: number, py: number) => ({ x: f.column.left + px, y: f.paragraph.top + py });
      const kind = (id: string) => SHAPE_KINDS.find((k) => k.id === id)!;
      ed.shapes.insert(view, kind('process'), P(20, 520), P(140, 570));
      placeShapes(view.dom);
      ed.shapes.insert(view, kind('decision'), P(240, 510), P(380, 580));
      placeShapes(view.dom);
      ed.shapes.insert(view, kind('arrowLine'), P(141, 545), P(239, 545));
      placeShapes(view.dom);
      ed.shapes.insert(view, kind('elbow'), P(310, 581), P(80, 640));
      placeShapes(view.dom);
      ed.shapes.insert(view, kind('textbox'), P(20, 660), P(220, 720));
      ed.run((s, d) => (d?.(s.tr.insertText('新增的文字方塊')), true));
      ed.closeShapeText(false);
      placeShapes(view.dom);
      // A copy of 甲 pasted in, then moved below it.
      const src = x.all().find((s) => textFrames(s.shape).some((t) => schema.nodeFromJSON(t.doc).textContent === '甲'))!;
      view.dispatch(view.state.tr.insert(1, schema.nodes.raw_inline.create(pastedShape(src.node.attrs.xml)!)));
      placeShapes(view.dom);
      const copy = x.all().filter((s) => s.shape.name === src.shape.name).find((s) => s.shape.key !== src.shape.key)!;
      view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, copy.pos)));
      placeShapes(view.dom);
      ed.shapes.moveBy(view, 0, 120);
      for (const s of x.all()) {
        const g = (s.shape.root as SpNode).geom;
        if (['flowChartProcess', 'flowChartDecision', 'straightConnector1', 'bentConnector3'].includes(g)) x.expectShape(s.shape.name!);
      }
      const box = x.all().find((s) => textFrames(s.shape).some((t) => schema.nodeFromJSON(t.doc).textContent === '新增的文字方塊'))!;
      x.expectShape(box.shape.name!, '新增的文字方塊');
    });
    writeFileSync(join(out!, 'shapes-8b.json'), JSON.stringify(cases, null, 2));
  }, 120_000);
});
