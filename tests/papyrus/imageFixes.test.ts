// Replacing a picture only puts formats Word shows into the document (others are converted to
// PNG or refused), the writer names media by what the bytes are, and the resize handles clean
// up after themselves, capture the pointer and measure an unsized picture in page px.
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import JSZip from 'jszip';
import { EditorState, NodeSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { readDocx } from '../../src/papyrus/docx/reader';
import { writeDocx } from '../../src/papyrus/docx/writer';
import { blankPackage } from '../../src/papyrus/docx/template';
import { schema } from '../../src/papyrus/editor/schema';
import { UNSUPPORTED_IMAGE, imageNodeView, readImageFile, replaceImage } from '../../src/papyrus/editor/imageView';

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const GIF = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
const JPEG = '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAEBAQ==';
const WEBP = btoa('RIFF\x1a\x00\x00\x00WEBPVP8L\x0d\x00\x00\x00\x2f\x00\x00\x00\x10');
const HEIC = btoa('\x00\x00\x00\x18ftypheic\x00\x00\x00\x00mif1heic');
const bytes = (b64: string) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));

beforeAll(() => {
  const zero = () => ({ x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, toJSON() {} }) as DOMRect;
  Range.prototype.getBoundingClientRect ??= zero;
  Range.prototype.getClientRects ??= () => Object.assign([], { item: () => null }) as unknown as DOMRectList;
});
afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** A browser Image that loads everything but the formats in `broken`. */
function fakeImages(broken: string[] = []) {
  vi.stubGlobal(
    'Image',
    class {
      naturalWidth = 0;
      naturalHeight = 0;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      set src(v: string) {
        setTimeout(() => {
          if (broken.some((t) => v.startsWith(`data:${t}`))) this.onerror?.();
          else {
            this.naturalWidth = 40;
            this.naturalHeight = 20;
            this.onload?.();
          }
        });
      }
    },
  );
  const drawn: unknown[] = [];
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => ({ drawImage: (i: unknown) => drawn.push(i) }) as any);
  vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockImplementation(() => `data:image/png;base64,${PNG}`);
  return drawn;
}

describe('picking a replacement picture', () => {
  it('PNG, JPEG, GIF and BMP are used as they are, typed by their bytes', async () => {
    const drawn = fakeImages();
    const png = await readImageFile(new Blob([bytes(PNG)]));
    expect(png).toEqual({ src: `data:image/png;base64,${PNG}`, width: 40, height: 20 });
    // A JPEG saved as ".png".
    const jpeg = await readImageFile(new Blob([bytes(JPEG)], { type: 'image/png' }));
    expect(jpeg.src).toBe(`data:image/jpeg;base64,${JPEG}`);
    expect(drawn).toHaveLength(0);
  });

  it('other formats are converted to PNG', async () => {
    const drawn = fakeImages();
    const webp = await readImageFile(new Blob([bytes(WEBP)], { type: 'image/webp' }));
    expect(webp).toEqual({ src: `data:image/png;base64,${PNG}`, width: 40, height: 20 });
    expect(drawn).toHaveLength(1);
  });

  it('a format the browser cannot show either is refused with a message', async () => {
    fakeImages(['image/heic']);
    await expect(readImageFile(new Blob([bytes(HEIC)], { type: 'image/heic' }))).rejects.toThrow(UNSUPPORTED_IMAGE);
    expect(UNSUPPORTED_IMAGE).toMatch(/PNG/);
  });

  it('the replace command only takes formats Word shows', () => {
    const img = schema.nodes.image.create({ src: `data:image/png;base64,${PNG}`, width: 10, height: 10 });
    const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, img));
    const state = EditorState.create({ schema, doc, selection: NodeSelection.create(doc, 1) });
    expect(replaceImage(`data:image/webp;base64,${WEBP}`, 1, 1)(state)).toBe(false);
    expect(replaceImage(`data:image/gif;base64,${GIF}`, 1, 1)(state)).toBe(true);
  });
});

describe('media names follow the bytes', () => {
  it('never stores non-PNG bytes as .png', async () => {
    const { model } = await readDocx(await blankPackage().generateAsync({ type: 'uint8array' }));
    const pics = [
      `data:image/png;base64,${GIF}`, // mislabelled GIF
      `data:image/avif;base64,${btoa('\x00\x00\x00\x1cftypavif')}`, // not a PNG either
      `data:application/octet-stream;base64,${PNG}`, // a PNG without a type
    ].map((src) => schema.nodes.image.create({ src, width: 10, height: 10 }));
    const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, pics));
    const zip = await JSZip.loadAsync(await writeDocx(doc, model));
    const media = Object.keys(zip.files).filter((f) => f.startsWith('word/media/') && !zip.files[f].dir);
    expect(media.map((f) => f.split('.').pop()).sort()).toEqual(['avif', 'gif', 'png']);
    const pngs = media.filter((f) => f.endsWith('.png'));
    for (const f of pngs) expect(Array.from((await zip.file(f)!.async('uint8array')).slice(0, 4))).toEqual([0x89, 0x50, 0x4e, 0x47]);
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('Extension="avif" ContentType="image/avif"');
  });
});

describe('resize handles', () => {
  function mount(attrs: Record<string, unknown>) {
    const img = schema.nodes.image.create({ src: `data:image/png;base64,${PNG}`, ...attrs });
    const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, [schema.text('x'), img]));
    const host = document.createElement('div');
    document.body.append(host);
    const view = new EditorView(host, { state: EditorState.create({ schema, doc }), nodeViews: { image: imageNodeView } });
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, 2)));
    const wrap = view.dom.querySelector('.dx-img-wrap') as HTMLElement;
    const handle = (h: string) => wrap.querySelector(`[data-handle="${h}"]`) as HTMLElement;
    const down = (h: string, x = 0, pointerId = 1) => {
      const e = new MouseEvent('pointerdown', { clientX: x, clientY: 0, bubbles: true, cancelable: true });
      Object.defineProperty(e, 'pointerId', { value: pointerId });
      handle(h).dispatchEvent(e);
    };
    const done = () => {
      if (!view.isDestroyed) view.destroy();
      host.remove();
    };
    return { view, wrap, handle, down, done };
  }

  it('destroying the view during a drag removes its window listeners', () => {
    const d = mount({ width: 100, height: 50 });
    const added: [string, unknown][] = [];
    const removed: [string, unknown][] = [];
    const add = window.addEventListener.bind(window);
    const rem = window.removeEventListener.bind(window);
    vi.spyOn(window, 'addEventListener').mockImplementation((t: string, f: any, o?: any) => (added.push([t, f]), add(t, f, o)));
    vi.spyOn(window, 'removeEventListener').mockImplementation((t: string, f: any, o?: any) => (removed.push([t, f]), rem(t, f, o)));
    d.down('se');
    expect(added.map(([t]) => t)).toEqual(['pointermove', 'pointerup', 'pointercancel']);
    d.view.destroy();
    for (const [t, f] of added) expect(removed.some(([t2, f2]) => t2 === t && f2 === f), t).toBe(true);
    expect(() => window.dispatchEvent(new MouseEvent('pointerup'))).not.toThrow();
    d.done();
  });

  it('captures the pointer', () => {
    const d = mount({ width: 100, height: 50 });
    const capture = vi.fn();
    (d.handle('e') as any).setPointerCapture = capture;
    d.down('e', 0, 7);
    expect(capture).toHaveBeenCalledWith(7);
    window.dispatchEvent(new MouseEvent('pointerup'));
    d.done();
  });

  it('a picture without a size starts from its page size, not its zoomed screen size', () => {
    const d = mount({ width: null, height: null });
    const img = d.wrap.querySelector('img')!;
    // Shown at 200% zoom: 100 x 50 page px are 200 x 100 on screen.
    Object.defineProperty(img, 'offsetWidth', { value: 100 });
    Object.defineProperty(img, 'offsetHeight', { value: 50 });
    img.getBoundingClientRect = () => ({ x: 0, y: 0, top: 0, left: 0, right: 200, bottom: 100, width: 200, height: 100, toJSON() {} }) as DOMRect;
    d.down('e', 0);
    window.dispatchEvent(new MouseEvent('pointermove', { clientX: 40, clientY: 0 }));
    window.dispatchEvent(new MouseEvent('pointerup'));
    expect(d.view.state.doc.nodeAt(2)!.attrs).toMatchObject({ width: 120, height: 50 });
    d.done();
  });
});
