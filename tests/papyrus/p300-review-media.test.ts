// persona-300 code review, pictures and pasting: drawings stay PNG (JPEG for photos only), a
// picture decoded from the file (a huge one at its new size), the clipboard's picture file used
// only when its proportions are the copied picture's, a picture Word turned or cropped, and
// Excel's style text read in one pass.
import { afterEach, describe, expect, it, vi } from 'vitest';
import { NodeSelection } from 'prosemirror-state';
import { schema } from '../../src/papyrus/editor/schema';
import { headerImageSize, readImageFile, rotateImageRight } from '../../src/papyrus/editor/imageView';
import { PASTE_PART_CLOSED, sameProportions } from '../../src/papyrus/editor/core';
import { readDrawing } from '../../src/papyrus/docx/props';
import { cssRules, inlineExcelStyles, MAX_STYLE_TEXT } from '../../src/papyrus/editor/pasteExcel';
import { domStubs, setup } from './p300Helpers';

domStubs();
afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const wait = (ms = 20) => new Promise((r) => setTimeout(r, ms));
async function until(ready: () => boolean) {
  for (let i = 0; i < 300 && !ready(); i++) await wait(10);
}
/** A data URL of `n` bytes of content. */
const urlOf = (type: string, n: number) => `data:${type};base64,${btoa('x'.repeat(n))}`;

/** A PNG header of `w` × `h` px (IHDR), padded to `size` bytes. */
function pngOf(w: number, h: number, size = 64): Uint8Array {
  const b = new Uint8Array(Math.max(size, 33));
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  new DataView(b.buffer).setUint32(16, w);
  new DataView(b.buffer).setUint32(20, h);
  return b;
}
/** A BMP header of `w` × `h` px, padded to `size` bytes. */
function bmpOf(w: number, h: number, size = 64): Uint8Array {
  const b = new Uint8Array(Math.max(size, 54));
  b.set([0x42, 0x4d]);
  new DataView(b.buffer).setInt32(18, w, true);
  new DataView(b.buffer).setInt32(22, -h, true); // top-down
  return b;
}

/**
 * A browser whose pictures are `w` × `h`, whose canvas pixels are `pixels(i)` (RGBA of pixel
 * i) and which writes `out(type)`; records what is asked of it.
 */
function fakeBrowser(w: number, h: number, pixels: (i: number) => [number, number, number, number], out: (type: string) => string) {
  const log = { types: [] as string[], draws: [] as number[][], dataUrls: 0 };
  vi.stubGlobal(
    'Image',
    class {
      naturalWidth = 0;
      naturalHeight = 0;
      onload: (() => void) | null = null;
      set src(_v: string) {
        setTimeout(() => {
          this.naturalWidth = w;
          this.naturalHeight = h;
          this.onload?.();
        });
      }
    },
  );
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function () {
    return {
      drawImage: (...a: any[]) => log.draws.push(a.slice(1)),
      translate: () => {},
      rotate: () => {},
      fillRect: () => {},
      getImageData: (_x: number, _y: number, cw: number, ch: number) => {
        const data = new Uint8ClampedArray(cw * ch * 4);
        for (let i = 0; i < cw * ch; i++) data.set(pixels(i), i * 4);
        return { data };
      },
    } as any;
  });
  vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockImplementation((type?: string) => {
    log.types.push(type ?? 'image/png');
    return out(type ?? 'image/png');
  });
  const readAsDataURL = FileReader.prototype.readAsDataURL;
  vi.spyOn(FileReader.prototype, 'readAsDataURL').mockImplementation(function (this: FileReader, blob: Blob) {
    log.dataUrls++;
    return readAsDataURL.call(this, blob);
  });
  return log;
}
const white = (): [number, number, number, number] => [255, 255, 255, 255];
/** A different colour at every pixel: a photo. */
const photo = (i: number): [number, number, number, number] => [i & 255, (i >> 8) & 255, (i * 7) & 255, 255];

describe('pictures saved as PNG unless they are photos', () => {
  it('a BMP screenshot (few colours) becomes PNG, not JPEG', async () => {
    const log = fakeBrowser(300, 200, (i) => (i % 5 ? white() : [0, 0, 0, 255]), (t) => urlOf(t, t.includes('png') ? 5000 : 1000));
    const out = await readImageFile(new Blob([bmpOf(300, 200, 180_000)], { type: 'image/bmp' }));
    expect(out.src.startsWith('data:image/png')).toBe(true);
    expect(log.types).toEqual(['image/png']);
  });

  it('a BMP photo (many colours) becomes JPEG when its PNG is much bigger, else PNG', async () => {
    fakeBrowser(300, 200, photo, (t) => urlOf(t, t.includes('png') ? 90_000 : 20_000));
    expect((await readImageFile(new Blob([bmpOf(300, 200, 180_000)]))).src.startsWith('data:image/jpeg')).toBe(true);
    vi.restoreAllMocks();
    fakeBrowser(300, 200, photo, (t) => urlOf(t, t.includes('png') ? 30_000 : 20_000));
    expect((await readImageFile(new Blob([bmpOf(300, 200, 180_000)]))).src.startsWith('data:image/png')).toBe(true);
  });

  it('WebP and SVG drawings become PNG; see-through pictures always PNG', async () => {
    const webp = new Uint8Array(64);
    webp.set([0x52, 0x49, 0x46, 0x46, 0, 0, 0, 0, 0x57, 0x45, 0x42, 0x50]);
    fakeBrowser(100, 100, white, (t) => urlOf(t, t.includes('png') ? 4000 : 1000));
    expect((await readImageFile(new Blob([webp]))).src.startsWith('data:image/png')).toBe(true);
    const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>');
    expect((await readImageFile(new Blob([svg]))).src.startsWith('data:image/png')).toBe(true);
    vi.restoreAllMocks();
    const log = fakeBrowser(100, 100, (i) => [i & 255, i >> 8, 3, i === 5 ? 0 : 255], (t) => urlOf(t, t.includes('png') ? 90_000 : 1000));
    expect((await readImageFile(new Blob([webp]))).src.startsWith('data:image/png')).toBe(true);
    expect(log.types).toEqual(['image/png']);
  });

  it('HEIC (a photo format) becomes JPEG', async () => {
    const heic = new Uint8Array(64);
    heic.set([0, 0, 0, 24, 0x66, 0x74, 0x79, 0x70, 0x68, 0x65, 0x69, 0x63]); // ....ftypheic
    fakeBrowser(400, 300, white, (t) => urlOf(t, 3000));
    expect((await readImageFile(new Blob([heic], { type: 'image/heic' }))).src.startsWith('data:image/jpeg')).toBe(true);
  });
});

describe('pictures decoded from the file', () => {
  it('reads the size in the header of PNG, GIF, BMP, JPEG and WebP files', () => {
    expect(headerImageSize(pngOf(4000, 3000), 'png')).toEqual({ width: 4000, height: 3000 });
    expect(headerImageSize(bmpOf(640, 480), 'bmp')).toEqual({ width: 640, height: 480 });
    const gif = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x20, 0x03, 0x58, 0x02]);
    expect(headerImageSize(gif, 'gif')).toEqual({ width: 800, height: 600 });
    // SOI, APP0 (16 bytes), SOF0: height 1500, width 2000.
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 16, ...new Array(14).fill(0), 0xff, 0xc0, 0, 17, 8, 0x05, 0xdc, 0x07, 0xd0, 3, 0, 0, 0, 0]);
    expect(headerImageSize(jpeg, 'jpeg')).toEqual({ width: 2000, height: 1500 });
    const vp8x = new Uint8Array(30);
    vp8x.set([0x52, 0x49, 0x46, 0x46, 0, 0, 0, 0, 0x57, 0x45, 0x42, 0x50, 0x56, 0x50, 0x38, 0x58]);
    vp8x.set([0x1f, 0x03, 0, 0x57, 0x02, 0], 24); // 800 − 1, 600 − 1
    expect(headerImageSize(vp8x, 'webp')).toEqual({ width: 800, height: 600 });
    expect(headerImageSize(new Uint8Array(8), 'png')).toBeNull();
  });

  it('a huge picture is decoded at its new size, and never read as text when it is drawn again', async () => {
    const log = fakeBrowser(1, 1, white, (t) => urlOf(t, 1000));
    const bitmap = vi.fn(async (_file: Blob, options: any) => ({ width: options.resizeWidth ?? 8000, height: options.resizeHeight ?? 6000, close: () => {} }));
    vi.stubGlobal('createImageBitmap', bitmap);
    const out = await readImageFile(new Blob([pngOf(8000, 6000, 3_000_000)], { type: 'image/png' }));
    expect(bitmap).toHaveBeenCalledTimes(1);
    expect(bitmap.mock.calls[0][1]).toMatchObject({ resizeWidth: 2000, resizeHeight: 1500, imageOrientation: 'from-image' });
    expect(out).toMatchObject({ width: 2000, height: 1500 });
    expect(log.draws[0]).toEqual([0, 0, 2000, 1500]);
    expect(log.dataUrls).toBe(0);
  });

  it('a small picture used as it is is read once, as the document keeps it', async () => {
    const log = fakeBrowser(40, 20, white, (t) => urlOf(t, 1000));
    const out = await readImageFile(new Blob([pngOf(40, 20, 200)], { type: 'image/png' }));
    expect(out.src.startsWith('data:image/png;base64,iVBOR')).toBe(true);
    expect(log.types).toEqual([]);
  });
});

/** A paste event with `html` and picture `files` on its clipboard. */
function pasteEvent(html: string, files: File[] = []) {
  const e = new Event('paste', { bubbles: true, cancelable: true }) as ClipboardEvent;
  const data = { types: ['text/html', 'text/plain', 'Files'], files, getData: (t: string) => (t === 'text/html' ? html : t === 'text/plain' ? 'x' : '') };
  Object.defineProperty(e, 'clipboardData', { value: data });
  return e;
}
const WORD_ONE = '<html><body><!--StartFragment--><p>說明如下</p><p><img width=120 height=60 src="file:///C:/Temp/msohtmlclip1/01/clip_image001.png" alt="照片"></p><!--EndFragment--></body></html>';
const images = (doc: any) => {
  const out: any[] = [];
  doc.descendants((n: any) => void (n.type === schema.nodes.image && out.push(n.attrs)));
  return out;
};

describe('the clipboard picture file of a paste', () => {
  it('proportions within 10 %, or no size given, count as the same picture', () => {
    expect(sameProportions(120, 60, 400, 200)).toBe(true);
    expect(sameProportions(120, 60, 400, 215)).toBe(true);
    expect(sameProportions(120, 60, 400, 300)).toBe(false);
    expect(sameProportions(null, null, 400, 300)).toBe(true);
  });

  it('Word text copied with one picture: a picture of the whole selection (other proportions) is not pasted as the picture', async () => {
    fakeBrowser(800, 900, white, () => urlOf('image/png', 100));
    const d = await setup('<w:p/>');
    const file = new File([pngOf(800, 900)], 'image.png', { type: 'image/png' });
    d.view.dom.dispatchEvent(pasteEvent(WORD_ONE, [file]));
    await until(() => d.view.state.doc.textContent.includes('說明如下'));
    expect(images(d.view.state.doc)).toHaveLength(0);
    expect(d.view.state.doc.textContent).toContain('照片'); // its alt text
    expect(d.notices).toContain('有 1 張圖片無法一起貼上，請用『插入 › 圖片』加入。');
    d.done();
  });

  it('pasted in the body, then the header opened while the pictures are read: the body gets the paste', async () => {
    fakeBrowser(400, 200, white, () => urlOf('image/png', 100));
    const d = await setup('<w:p/>');
    d.view.dom.dispatchEvent(pasteEvent(WORD_ONE, [new File([pngOf(400, 200)], 'a.png', { type: 'image/png' })]));
    d.editor.editHeaderFooter('header');
    await until(() => d.view.state.doc.textContent.includes('說明如下'));
    expect(images(d.view.state.doc)).toHaveLength(1);
    expect(d.editor.activeView!.state.doc.textContent).toBe('');
    d.done();
  });

  it('pasted in a header that is closed while the pictures are read: nothing is lost silently', async () => {
    fakeBrowser(400, 200, white, () => urlOf('image/png', 100));
    const d = await setup('<w:p/>');
    d.editor.editHeaderFooter('header');
    const header = d.editor.activeView!;
    header.dom.dispatchEvent(pasteEvent(WORD_ONE, [new File([pngOf(400, 200)], 'a.png', { type: 'image/png' })]));
    d.editor.closeHeaderFooter();
    await until(() => d.notices.includes(PASTE_PART_CLOSED));
    expect(d.notices).toContain(PASTE_PART_CLOSED);
    expect(d.view.state.doc.textContent).toBe('');
    d.done();
  });
});

const DRAWING = (xfrm: string, srcRect = '') =>
  '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  `<pic:blipFill><a:blip r:embed="rId9"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm${xfrm}><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

describe('turning a picture Word turned, flipped or cropped', () => {
  it('reads Word’s turn, flip and crop', () => {
    expect(readDrawing(DRAWING('')).turned).toBe(false);
    expect(readDrawing(DRAWING(' rot="5400000"')).turned).toBe(true);
    expect(readDrawing(DRAWING(' flipH="1"')).turned).toBe(true);
    expect(readDrawing(DRAWING(' rot="21600000"')).turned).toBe(false);
    expect(readDrawing(DRAWING('', '<a:srcRect l="10000" r="20000" b="50000"/>')).crop).toEqual({ left: 0.1, top: 0, right: 0.2, bottom: 0.5 });
    expect(readDrawing(DRAWING('', '<a:srcRect l="-5000"/>')).crop).toBeNull();
  });

  it('a cropped picture is turned as Word shows it: only the part Word shows', async () => {
    const log = fakeBrowser(1000, 400, white, () => 'data:image/png;base64,AAAA');
    expect(await rotateImageRight('data:image/png;base64,AAAA', { left: 0.1, top: 0, right: 0.2, bottom: 0.5 })).toBe('data:image/png;base64,AAAA');
    // Source rectangle x 100, y 0, 700 × 200, drawn at 0, 0.
    expect(log.draws[0]).toEqual([100, 0, 700, 200, 0, 0, 700, 200]);
  });

  it('向右旋轉 90° on a picture Word turned: not turned again, the user is told', async () => {
    fakeBrowser(100, 50, white, () => 'data:image/png;base64,AAAA');
    const d = await setup('<w:p/>', 'ImagePanel.vue');
    const src = 'data:image/png;base64,iVBORw0KGgo=';
    d.view.dispatch(d.view.state.tr.insert(1, schema.nodes.image.create({ src, origSrc: src, width: 100, height: 50, xml: DRAWING(' rot="5400000"') })));
    d.view.dispatch(d.view.state.tr.setSelection(NodeSelection.create(d.view.state.doc, 1)));
    await d.refresh();
    const button = Array.from(d.panel.querySelectorAll('button')).find((b) => b.textContent === '向右旋轉 90°')!;
    expect(button.getAttribute('aria-disabled')).toBe('true');
    button.click();
    await wait(30);
    await d.refresh();
    expect(d.panel.textContent).toContain('這張圖片已在 Word 中旋轉或翻轉，請用 Word 旋轉。');
    expect(d.view.state.doc.nodeAt(1)!.attrs).toMatchObject({ src, width: 100, height: 50 });
    d.done();
  });
});

describe('Excel style text', () => {
  it('rules are read in one pass, at-rule blocks included', () => {
    expect(cssRules('.a{color:red} @media print { td.xl65 { font-weight:700 } } .b{x:y}')).toEqual([
      ['.a', 'color:red'],
      [' td.xl65 ', ' font-weight:700 '],
      [' .b', 'x:y'],
    ]);
    expect(cssRules('no braces at all')).toEqual([]);
  });

  it('a long style text without braces is quick, and only the first 200 KB is read', () => {
    const start = performance.now();
    const html = `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><style>${'a'.repeat(3_000_000)}</style><table><tr><td class="xl65">1</td></tr></table></html>`;
    inlineExcelStyles(html);
    expect(performance.now() - start).toBeLessThan(1500);
    expect(MAX_STYLE_TEXT).toBe(200 * 1024);
  });
});
