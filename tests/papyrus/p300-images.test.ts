// persona-300 A-9: pictures. Pasted pictures that can't come along are counted and the user is
// told how to add them (the clipboard's picture files are used when there is one per picture);
// big photos are made smaller as they are inserted; 插入 › 圖片 takes several files; 向右旋轉
// 90°; the document's size for the host's warning.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick } from 'vue';
import { NodeSelection, TextSelection } from 'prosemirror-state';
import JSZip from 'jszip';
import { schema } from '../../src/papyrus/editor/schema';
import {
  MAX_IMAGE_SIDE, dataUrlBytes, jpegOrientation, planImage, readImageFile, rotateImage, rotateImageRight,
} from '../../src/papyrus/editor/imageView';
import { droppedPicturesNotice } from '../../src/papyrus/editor/core';
import { domStubs, setup } from './p300Helpers';

domStubs();
afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const JPEG_OUT = '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAEBAQ==';
const bytes = (b64: string) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
const wait = (ms = 20) => new Promise((r) => setTimeout(r, ms));
/** Waits (up to 3 s) until `ready` says so: reading picture files takes a few turns. */
async function until(ready: () => boolean) {
  for (let i = 0; i < 300 && !ready(); i++) await wait(10);
}

/** A JPEG whose EXIF says orientation `o` (big-endian TIFF), padded to `size` bytes. */
function exifJpeg(o: number, size = 64): Uint8Array {
  const tiff = [0x4d, 0x4d, 0x00, 0x2a, 0, 0, 0, 8, 0, 1, 0x01, 0x12, 0, 3, 0, 0, 0, 1, 0, o, 0, 0, 0, 0, 0, 0, 0, 0];
  const app1 = [0xff, 0xe1, 0, 2 + 6 + tiff.length, 0x45, 0x78, 0x69, 0x66, 0, 0, ...tiff];
  const out = new Uint8Array(Math.max(size, 4 + app1.length + 2));
  out.set([0xff, 0xd8, ...app1, 0xff, 0xda]);
  return out;
}

/** Browser pictures of `w` × `h`; a canvas that records what is drawn and writes `out`. */
function fakeBrowser(w: number, h: number, out: (type: string) => string) {
  vi.stubGlobal(
    'Image',
    class {
      naturalWidth = 0;
      naturalHeight = 0;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      set src(_v: string) {
        setTimeout(() => {
          this.naturalWidth = w;
          this.naturalHeight = h;
          this.onload?.();
        });
      }
    },
  );
  const calls: { canvas: [number, number]; type?: string; ops: string[] }[] = [];
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (this: HTMLCanvasElement) {
    const call = { canvas: [this.width, this.height] as [number, number], ops: [] as string[] };
    calls.push(call);
    return {
      drawImage: () => call.ops.push('draw'),
      translate: () => call.ops.push('translate'),
      rotate: () => call.ops.push('rotate'),
      fillRect: () => call.ops.push('fill'),
      getImageData: (_x: number, _y: number, cw: number, ch: number) => ({ data: new Uint8ClampedArray(cw * ch * 4).fill(255) }),
    } as any;
  });
  vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockImplementation(function (type?: string) {
    calls[calls.length - 1].type = type;
    return out(type ?? 'image/png');
  });
  return calls;
}

describe('pictures made smaller as they are inserted', () => {
  it('plans: a big photo becomes 2000 px on its long side, a small graphic stays as it is, GIF is never redrawn', () => {
    expect(MAX_IMAGE_SIDE).toBe(2000);
    expect(planImage({ type: 'image/jpeg', bytes: 5e6, width: 4000, height: 3000 })).toEqual({ encode: true, format: 'jpeg', width: 2000, height: 1500 });
    expect(planImage({ type: 'image/png', bytes: 40e3, width: 800, height: 600 })).toMatchObject({ encode: false });
    expect(planImage({ type: 'image/jpeg', bytes: 300e3, width: 1200, height: 900, orientation: 6 })).toMatchObject({ encode: true, width: 1200, height: 900 });
    expect(planImage({ type: 'image/gif', bytes: 9e6, width: 3000, height: 3000 })).toMatchObject({ encode: false });
    // Never bigger: a small picture keeps its size.
    expect(planImage({ type: 'image/bmp', bytes: 30e3, width: 100, height: 50 })).toMatchObject({ encode: true, width: 100, height: 50 });
  });

  it('reads the EXIF orientation of a JPEG', () => {
    expect(jpegOrientation(exifJpeg(6))).toBe(6);
    expect(jpegOrientation(exifJpeg(1))).toBe(1);
    expect(jpegOrientation(bytes(PNG))).toBe(1);
  });

  it('a 4000 × 3000 photo is saved as a 2000 × 1500 JPEG', async () => {
    const calls = fakeBrowser(4000, 3000, () => `data:image/jpeg;base64,${JPEG_OUT}`);
    const photo = new Blob([exifJpeg(1, 3_000_000)], { type: 'image/jpeg' });
    const out = await readImageFile(photo);
    expect(out).toEqual({ src: `data:image/jpeg;base64,${JPEG_OUT}`, width: 2000, height: 1500 });
    expect(calls[0].canvas).toEqual([2000, 1500]);
    expect(calls[0].type).toBe('image/jpeg');
  });

  it('a turned photo is drawn turned even when that is not smaller; a small PNG is used as it is', async () => {
    const big = `data:image/jpeg;base64,${btoa('x'.repeat(3000))}`;
    fakeBrowser(300, 400, () => big);
    const turned = await readImageFile(new Blob([exifJpeg(6, 2000)], { type: 'image/jpeg' }));
    expect(turned.src).toBe(big);
    const calls = fakeBrowser(40, 20, () => `data:image/png;base64,${PNG}`);
    const small = await readImageFile(new Blob([bytes(PNG)], { type: 'image/png' }));
    expect(small).toEqual({ src: `data:image/png;base64,${PNG}`, width: 40, height: 20 });
    expect(calls).toHaveLength(0);
  });

  it('a redrawn picture that would be bigger is not used', async () => {
    fakeBrowser(2500, 100, () => `data:image/png;base64,${btoa('y'.repeat(4_000_000))}`);
    const file = new Blob([bytes(PNG), new Uint8Array(1_500_000)], { type: 'image/png' });
    const out = await readImageFile(file);
    expect(out.src.startsWith('data:image/png;base64,iVBOR')).toBe(true);
    expect(out.width).toBe(2500);
  });
});

/** A paste event with `html` and picture `files` on its clipboard. */
function pasteEvent(html: string, files: File[] = []) {
  const e = new Event('paste', { bubbles: true, cancelable: true }) as ClipboardEvent;
  const data = {
    types: ['text/html', 'text/plain', ...(files.length ? ['Files'] : [])],
    files,
    getData: (t: string) => (t === 'text/html' ? html : t === 'text/plain' ? 'x' : ''),
  };
  Object.defineProperty(e, 'clipboardData', { value: data });
  return e;
}

describe('pasting pictures', () => {
  const WORD_HTML = (n: number) =>
    `<html><body><!--StartFragment--><p>說明如下</p>` +
    Array.from({ length: n }, (_, i) => `<p><img width=120 height=60 src="file:///C:/Users/a/AppData/Local/Temp/msohtmlclip1/01/clip_image00${i + 1}.png" alt="照片${i + 1}"></p>`).join('') +
    `<!--EndFragment--></body></html>`;

  it('pictures that cannot come along are counted and the user is told how to add them', async () => {
    const d = await setup('<w:p/>');
    d.view.dom.dispatchEvent(pasteEvent(WORD_HTML(2)));
    await wait();
    expect(d.notices).toContain('有 2 張圖片無法一起貼上，請用『插入 › 圖片』加入。');
    expect(droppedPicturesNotice(1)).toBe('有 1 張圖片無法一起貼上，請用『插入 › 圖片』加入。');
    expect(d.view.state.doc.textContent).toContain('說明如下');
    let images = 0;
    d.view.state.doc.descendants((n) => void (n.type === schema.nodes.image && images++));
    expect(images).toBe(0);
    d.done();
  });

  it('the clipboard picture files are used in their places, at the copied size', async () => {
    fakeBrowser(400, 200, () => `data:image/png;base64,${PNG}`);
    const d = await setup('<w:p/>');
    const file = new File([bytes(PNG)], 'image.png', { type: 'image/png' });
    d.view.dom.dispatchEvent(pasteEvent(WORD_HTML(1), [file]));
    await until(() => d.view.state.doc.textContent.includes('說明如下'));
    const imgs: any[] = [];
    d.view.state.doc.descendants((n) => void (n.type === schema.nodes.image && imgs.push(n.attrs)));
    expect(imgs).toHaveLength(1);
    expect(imgs[0]).toMatchObject({ src: `data:image/png;base64,${PNG}`, width: 120, height: 60, alt: '照片1' });
    expect(d.notices.filter((n) => n.includes('無法一起貼上'))).toEqual([]);
    expect(d.view.state.doc.textContent).toContain('說明如下');
    // One undo step takes the whole paste back.
    d.editor.undo();
    expect(d.view.state.doc.textContent).toBe('');
    d.done();
  });
});

describe('插入 › 圖片 and the picture tools', () => {
  it('several pictures can be picked and are inserted in order', async () => {
    fakeBrowser(40, 20, () => `data:image/png;base64,${PNG}`);
    const d = await setup('<w:p><w:r><w:t>AB</w:t></w:r></w:p>', 'DocxToolbar.vue', { styles: [] });
    d.view.dispatch(d.view.state.tr.setSelection(TextSelection.create(d.view.state.doc, 2)));
    const picker = d.panel.querySelector('input[type="file"][accept="image/*"]') as HTMLInputElement;
    expect(picker.multiple).toBe(true);
    const GIF = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
    const files = [new File([bytes(PNG)], 'a.png', { type: 'image/png' }), new File([bytes(GIF)], 'b.gif', { type: 'image/gif' })];
    Object.defineProperty(picker, 'files', { configurable: true, value: files });
    picker.dispatchEvent(new Event('change'));
    await until(() => d.view.state.doc.firstChild!.childCount === 4);
    const para = d.view.state.doc.firstChild!;
    expect(para.childCount).toBe(4);
    expect(para.child(1).attrs.src).toBe(`data:image/png;base64,${PNG}`);
    expect(para.child(2).attrs.src).toBe(`data:image/gif;base64,${GIF}`);
    expect(para.child(3).text).toBe('B');
    d.done();
  });

  it('向右旋轉 90°: the picture is drawn turned, width and height swap, the alt text stays; saved as a changed picture', async () => {
    const calls = fakeBrowser(40, 20, () => `data:image/png;base64,${PNG}`);
    expect(await rotateImageRight(`data:image/png;base64,${PNG}`)).toBe(`data:image/png;base64,${PNG}`);
    expect(calls[0].canvas).toEqual([20, 40]);
    expect(calls[0].ops).toEqual(['translate', 'rotate', 'draw']);

    const d = await setup('<w:p/>', 'ImagePanel.vue');
    const GIF = 'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
    d.view.dispatch(d.view.state.tr.insert(1, schema.nodes.image.create({ src: `data:image/gif;base64,${GIF}`, width: 100, height: 50, alt: '平面圖' })));
    d.view.dispatch(d.view.state.tr.setSelection(NodeSelection.create(d.view.state.doc, 1)));
    await d.refresh();
    const button = Array.from(d.panel.querySelectorAll('button')).find((b) => b.textContent === '向右旋轉 90°')!;
    button.click();
    await until(() => d.view.state.doc.nodeAt(1)!.attrs.width === 50);
    expect(d.view.state.doc.nodeAt(1)!.attrs).toMatchObject({ src: `data:image/png;base64,${PNG}`, width: 50, height: 100, alt: '平面圖' });
    // The command refuses formats Word may not show.
    expect(rotateImage('data:image/webp;base64,AAAA')(d.view.state)).toBe(false);
    const zip = await JSZip.loadAsync(await d.editor.save());
    const media = Object.keys(zip.files).filter((f) => f.startsWith('word/media/'));
    expect(media.some((f) => f.endsWith('.png'))).toBe(true);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect(xml).toContain('descr="平面圖"');
    d.done();
  });
});

describe('the document size, for a warning before it is too big', () => {
  it('grows with the pictures inserted', async () => {
    const d = await setup('<w:p/>');
    const before = d.editor.estimatedSize();
    expect(before).toBeGreaterThan(1000);
    const payload = btoa(String.fromCharCode(...bytes(PNG)) + 'z'.repeat(30000));
    const src = `data:image/png;base64,${payload}`;
    expect(dataUrlBytes(src)).toBe(bytes(PNG).length + 30000);
    d.view.dispatch(d.view.state.tr.insert(1, schema.nodes.image.create({ src, width: 10, height: 10 })));
    expect(d.editor.estimatedSize() - before).toBe(dataUrlBytes(src));
    d.done();
  });

  it('DocxEditor.vue reports it (size event) after opening', async () => {
    const DocxEditorVue = (await import('../../src/papyrus/vue/DocxEditor.vue')).default;
    const sizes: { bytes: number; large: boolean }[] = [];
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp({ render: () => h(DocxEditorVue, { onSize: (s: { bytes: number; large: boolean }) => sizes.push(s) }) });
    app.mount(host);
    for (let i = 0; i < 40 && !sizes.length; i++) {
      await wait(10);
      await nextTick();
    }
    expect(sizes[0].bytes).toBeGreaterThan(1000);
    expect(sizes[0].large).toBe(false);
    app.unmount();
    host.remove();
  });
});
