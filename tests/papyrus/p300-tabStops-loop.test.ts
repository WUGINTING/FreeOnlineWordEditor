// Review of persona-300 目錄: a right tab stop beyond the paragraph's right edge (a right indent,
// a table cell, a column) must not keep the page resizing: the width set on the tab is measured
// away (each tab at its own width), and the text after it ends at the paragraph's right edge.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { schema } from '../../src/papyrus/editor/schema';
import { measureTabs, tabStops, tabStopsKey } from '../../src/papyrus/editor/tabStops';

/** The fake page: a paragraph 500 px wide; the text before the tab ends at 100 px, the page number is 20 px wide. */
const PARA_RIGHT = 500;
const TAB_LEFT = 100;
const NUMBER = 20;

const rect = (left: number, top: number, width: number, height = 16) =>
  ({ left, top, width, height, right: left + width, bottom: top + height, x: left, y: top, toJSON() {} }) as DOMRect;

let observers: (() => void)[] = [];
let restore: (() => void)[] = [];

beforeEach(() => {
  observers = [];
  const RO = class {
    constructor(private cb: () => void) {}
    observe() {
      observers.push(this.cb);
    }
    disconnect() {}
    unobserve() {}
  };
  const oldRO = (globalThis as any).ResizeObserver;
  (globalThis as any).ResizeObserver = RO;
  const tabWidth = () => {
    const tab = document.querySelector<HTMLElement>('.dx-tab');
    return tab ? parseFloat(tab.style.width) || 48 : 48;
  };
  const el = vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    if (this.classList.contains('dx-tab')) return rect(TAB_LEFT, 0, tabWidth());
    if (this.tagName === 'P') return rect(0, 0, PARA_RIGHT, tabWidth() + TAB_LEFT + NUMBER > PARA_RIGHT ? 32 : 16);
    return rect(0, 0, 600, 800);
  });
  // The page number: on the tab's line, or on the next one once it does not fit.
  const oldRects = Range.prototype.getClientRects;
  Range.prototype.getClientRects = function () {
    const end = TAB_LEFT + tabWidth();
    const rects = end + NUMBER > PARA_RIGHT ? [rect(end, 0, 1), rect(0, 20, NUMBER)] : [rect(end, 0, NUMBER)];
    return Object.assign(rects, { item: (i: number) => rects[i] }) as unknown as DOMRectList;
  };
  restore = [
    () => ((globalThis as any).ResizeObserver = oldRO),
    () => el.mockRestore(),
    () => (Range.prototype.getClientRects = oldRects),
  ];
  vi.useFakeTimers();
});

afterEach(() => {
  vi.useRealTimers();
  restore.forEach((f) => f());
  document.body.innerHTML = '';
});

/** A table of contents line whose right tab (9000 twips = 600 px) lies beyond the paragraph's right edge (500 px). */
function view(onTabWidths: () => void) {
  const pPr = '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>';
  const doc = schema.nodes.doc.create(null, [
    schema.nodes.paragraph.create({ pPr }, [schema.text('第一章 總則'), schema.nodes.tab.create(), schema.text('12')]),
  ]);
  const host = document.createElement('div');
  document.body.append(host);
  const v: EditorView = new EditorView(host, {
    state: EditorState.create({ doc, plugins: [tabStops()] }),
    dispatchTransaction(tr) {
      v.updateState(v.state.apply(tr));
      if (tr.getMeta(tabStopsKey)) {
        onTabWidths();
        // The new width changes the page's height: the ResizeObserver says so.
        setTimeout(() => observers.forEach((cb) => cb()), 0);
      }
    },
  });
  return v;
}

describe('right tab stops beyond the right edge', () => {
  it('end the text at the paragraph’s right edge', () => {
    const v = view(() => {});
    const widths = measureTabs(v);
    // 500 (right edge) - 20 (page number) - 100 (where the tab starts)
    expect([...widths.values()]).toEqual([{ width: 380, leader: 'dot' }]);
    v.destroy();
  });

  it('are measured the same whatever width the tab has now (no feedback)', () => {
    const v = view(() => {});
    const tab = v.dom.querySelector<HTMLElement>('.dx-tab')!;
    tab.style.width = '480px'; // a width that makes the page number wrap
    expect([...measureTabs(v).values()]).toEqual([{ width: 380, leader: 'dot' }]);
    expect(tab.style.width).toBe('480px');
    v.destroy();
  });

  it('count from the text column the paragraph is in (a section’s column: its padding)', () => {
    const v = view(() => {});
    const para = v.dom.querySelector<HTMLElement>('p')!;
    // Text column 2 starts 150 px into the paragraph's box; the tab stop is 4500 twips (300 px) into it.
    v.dispatch(v.state.tr.setNodeMarkup(0, undefined, { ...v.state.doc.child(0).attrs, pPr: '<w:pPr><w:tabs><w:tab w:val="right" w:pos="4500"/></w:tabs></w:pPr>' }));
    v.dom.querySelector<HTMLElement>('p')!.style.paddingLeft = '150px';
    expect(para).toBeTruthy();
    expect([...measureTabs(v).values()]).toEqual([{ width: 330, leader: null }]);
    v.destroy();
  });

  it('leave a tab alone when the text after it is on the next line anyway', () => {
    const v = view(() => {});
    const old = Range.prototype.getClientRects;
    Range.prototype.getClientRects = () => Object.assign([rect(0, 20, NUMBER)], { item: () => null }) as unknown as DOMRectList;
    try {
      expect(measureTabs(v).size).toBe(0);
    } finally {
      Range.prototype.getClientRects = old;
      v.destroy();
    }
  });

  it('set the width once and stop (no endless resize loop)', () => {
    let sets = 0;
    const v = view(() => sets++);
    for (let i = 0; i < 40; i++) vi.advanceTimersByTime(100);
    expect(sets).toBe(1);
    expect(v.dom.querySelector<HTMLElement>('.dx-tab')!.style.width).toBe('380px');
    v.destroy();
  });
});
