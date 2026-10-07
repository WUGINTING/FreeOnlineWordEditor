// Table / picture / find panels: a box the user didn't change changes nothing, and the
// controls say what they are to assistive technology.
import { beforeAll, describe, expect, it } from 'vitest';
import { createApp, h, nextTick, ref } from 'vue';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import { blankPackage } from '../../src/papyrus/docx/template';
import { loadSfc } from './sfc';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
const VUE = 'src/papyrus/vue/';

beforeAll(() => {
  (globalThis as any).ResizeObserver ??= class {
    observe() {}
    unobserve() {}
    disconnect() {}
  };
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect ??= zero;
});

/** An editor on `body`, with the panel `file` mounted next to it (re-rendered on every update). */
async function setup(body: string, file: string, extra: Record<string, unknown> = {}) {
  const { DocxEditor } = await import('../../src/papyrus/editor/core');
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  const host = document.createElement('div');
  const panelHost = document.createElement('div');
  document.body.append(host, panelHost);
  const snapshot = ref<unknown>(null);
  const editor = new DocxEditor(host, { onUpdate: (s) => (snapshot.value = s) });
  await editor.open(await zip.generateAsync({ type: 'uint8array' }));
  const view = editor.view!;
  const Panel = await loadSfc(VUE + file);
  const app = createApp({ render: () => h(Panel, { editor, snapshot: snapshot.value, ...extra }) });
  app.mount(panelHost);
  const refresh = async () => {
    snapshot.value = editor.snapshot();
    await nextTick();
  };
  const done = () => {
    app.unmount();
    editor.destroy();
    host.remove();
    panelHost.remove();
  };
  return { editor, view, panel: panelHost, refresh, done };
}

const input = (el: HTMLInputElement | HTMLSelectElement, value: string, event = 'change') => {
  el.value = value;
  el.dispatchEvent(new Event('input'));
  el.dispatchEvent(new Event(event));
};
const enter = (el: HTMLElement) => el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));

describe('table panel', () => {
  const TABLE = (trPr: string) =>
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${trPr}<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`;

  it('Enter in an untouched height box keeps the height; the rule menu changes only the rule', async () => {
    const d = await setup(TABLE('<w:trPr><w:trHeight w:val="288" w:hRule="auto"/></w:trPr>'), 'TablePanel.vue');
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, 4)));
    await d.refresh();
    const row = () => d.view.state.doc.child(0).child(0).attrs;
    const box = d.panel.querySelector('input[aria-label="列高"]') as HTMLInputElement;
    expect(box.value).toBe('0.51');
    enter(box);
    box.dispatchEvent(new Event('change'));
    expect(row()).toMatchObject({ height: 288, heightRule: 'auto' });
    const rule = d.panel.querySelector('select[aria-label="列高規則"]') as HTMLSelectElement;
    input(rule, 'exact');
    await d.refresh();
    expect(row()).toMatchObject({ height: 288, heightRule: 'exact' });
    // A typed height is still applied.
    input(box, '1');
    expect(row().height).toBe(567);
    d.done();
  });

  it('the vertical alignment buttons say which one is on', async () => {
    const d = await setup(TABLE(''), 'TablePanel.vue');
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, 4)));
    await d.refresh();
    const pressed = () => Array.from(d.panel.querySelectorAll('[aria-label="垂直對齊"] button')).map((b) => b.getAttribute('aria-pressed'));
    expect(pressed()).toEqual(['true', 'false', 'false']);
    (d.panel.querySelectorAll('[aria-label="垂直對齊"] button')[2] as HTMLButtonElement).click();
    await d.refresh();
    expect(pressed()).toEqual(['false', 'false', 'true']);
    d.done();
  });
});

describe('picture panel', () => {
  it('Enter in an untouched size box keeps the picture size', async () => {
    const { schema } = await import('../../src/papyrus/editor/schema');
    const d = await setup('<w:p/>', 'ImagePanel.vue');
    // 3.3 cm is 124.72 px: shown rounded as "3.3", which must not be applied back.
    const img = schema.nodes.image.create({ src: 'data:image/gif;base64,R0lGODlhAQABAAAAACw=', width: 124.72, height: 50.3 });
    d.view.dispatch(d.view.state.tr.insert(1, img));
    d.view.dispatch(d.view.state.tr.setSelection(NodeSelection.create(d.view.state.doc, 1)));
    await d.refresh();
    const [w, hBox] = Array.from(d.panel.querySelectorAll('.dx-size input')) as HTMLInputElement[];
    expect(w.value).toBe('3.3');
    enter(w);
    w.dispatchEvent(new Event('change'));
    enter(hBox);
    expect(d.view.state.doc.nodeAt(1)!.attrs).toMatchObject({ width: 124.72, height: 50.3 });
    input(w, '5');
    expect(d.view.state.doc.nodeAt(1)!.attrs.width).toBeCloseTo(5 * (96 / 2.54), 6);
    d.done();
  });
});

describe('find panel', () => {
  it('the replace toggle has an accessible name', async () => {
    const d = await setup('<w:p><w:r><w:t>abc</w:t></w:r></w:p>', 'FindReplace.vue', { replace: false });
    await d.refresh();
    const toggle = d.panel.querySelector('.dx-toggle') as HTMLButtonElement;
    expect(toggle.getAttribute('aria-label')).toBe('顯示取代');
    d.done();
  });

  it('typing searches once the typing pauses; Enter searches at once', async () => {
    const { searchState } = await import('../../src/papyrus/editor/search');
    const d = await setup('<w:p><w:r><w:t>abc abd</w:t></w:r></w:p>', 'FindReplace.vue', { replace: false });
    await d.refresh();
    const box = d.panel.querySelector('input[aria-label="尋找"]') as HTMLInputElement;
    const typed = async (text: string) => {
      box.value = text;
      box.dispatchEvent(new Event('input'));
      await nextTick();
    };
    await typed('a');
    await typed('ab');
    expect(searchState(d.view.state).query).toBe('');
    await new Promise((r) => setTimeout(r, 250));
    expect(searchState(d.view.state).query).toBe('ab');
    expect(searchState(d.view.state).matches).toHaveLength(2);
    await typed('abd');
    enter(box);
    expect(searchState(d.view.state).query).toBe('abd');
    expect(searchState(d.view.state).current).toBe(0);
    d.done();
  });
});
