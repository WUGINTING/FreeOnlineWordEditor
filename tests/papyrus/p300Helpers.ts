// Shared set-up for the persona-300 editor tests: an editor on a small document, with a panel
// component mounted next to it (re-rendered on every update).
import { beforeAll } from 'vitest';
import { createApp, h, nextTick, ref } from 'vue';
import { blankPackage } from '../../src/papyrus/docx/template';
import { loadSfc } from './sfc';

export const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
export const SECT =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';

export function domStubs() {
  beforeAll(() => {
    (globalThis as any).ResizeObserver ??= class {
      observe() {}
      unobserve() {}
      disconnect() {}
    };
    const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
    Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
    Range.prototype.getBoundingClientRect ??= zero;
    (Element.prototype as any).scrollIntoView ??= () => {};
  });
}

/** A .docx whose body is `body` (w:p … XML). */
export async function docxOf(body: string, extra: (zip: ReturnType<typeof blankPackage>) => void = () => {}): Promise<Uint8Array> {
  const zip = blankPackage();
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${SECT}</w:body></w:document>`);
  extra(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

/** An editor on `body`, and optionally the panel `file` (src/papyrus/vue) mounted next to it. */
export async function setup(body: string, file?: string, extra: Record<string, unknown> = {}) {
  const { DocxEditor } = await import('../../src/papyrus/editor/core');
  const host = document.createElement('div');
  const panelHost = document.createElement('div');
  document.body.append(host, panelHost);
  const snapshot = ref<unknown>(null);
  const notices: string[] = [];
  const editor = new DocxEditor(host, { onUpdate: (s) => (snapshot.value = s), onNotice: (m) => notices.push(m) });
  await editor.open(await docxOf(body));
  const view = editor.view!;
  let app: ReturnType<typeof createApp> | null = null;
  if (file) {
    const Panel = await loadSfc('src/papyrus/vue/' + file);
    app = createApp({ render: () => h(Panel, { editor, snapshot: snapshot.value, ...extra }) });
    app.mount(panelHost);
  }
  const refresh = async () => {
    snapshot.value = editor.snapshot();
    await nextTick();
  };
  const done = () => {
    app?.unmount();
    editor.destroy();
    host.remove();
    panelHost.remove();
  };
  return { editor, view, panel: panelHost, host, notices, refresh, done };
}

/** A keydown on `el`; `ime` makes it a key of an input method still composing. */
export function key(el: Element, init: KeyboardEventInit & { ime?: boolean }): KeyboardEvent {
  const { ime, ...rest } = init;
  const e = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...rest, isComposing: !!ime });
  el.dispatchEvent(e);
  return e;
}

/** Types into a text box (input event, as the browser sends it). */
export async function typeInto(box: HTMLInputElement | HTMLTextAreaElement, text: string) {
  box.value = text;
  box.dispatchEvent(new Event('input'));
  await nextTick();
}
