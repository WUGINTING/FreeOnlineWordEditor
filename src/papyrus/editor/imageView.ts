// A picture in the page: shows resize handles while selected. Corner handles keep the
// aspect ratio; the right and bottom handles stretch one side. The new size is applied
// in one transaction when the drag ends (one undo step).

import { NodeSelection, type Command, type EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView, NodeView } from 'prosemirror-view';
import { sniffImageExt } from '../docx/writer';
import './features.css';

type Handle = 'nw' | 'ne' | 'sw' | 'se' | 'e' | 's';
const HANDLES: Handle[] = ['nw', 'ne', 'sw', 'se', 'e', 's'];
const MIN = 8; // px

export class ImageView implements NodeView {
  dom: HTMLElement;
  private img: HTMLImageElement;
  private dragging = false;
  /** Stops a drag in progress (removes its window listeners) without applying it. */
  private cancelDrag: (() => void) | null = null;

  constructor(
    private node: PMNode,
    private view: EditorView,
    private getPos: () => number | undefined,
  ) {
    this.dom = document.createElement('span');
    this.dom.className = 'dx-img-wrap';
    this.img = document.createElement('img');
    this.img.className = 'dx-img';
    this.img.draggable = false;
    this.dom.append(this.img);
    for (const h of HANDLES) {
      const el = document.createElement('span');
      el.className = `dx-img-handle dx-img-${h}`;
      el.dataset.handle = h;
      el.draggable = false;
      el.addEventListener('pointerdown', (e) => this.startDrag(e as PointerEvent, h));
      this.dom.append(el);
    }
    this.render();
  }

  private render(): void {
    const { src, alt, width, height } = this.node.attrs;
    if (this.img.getAttribute('src') !== src) this.img.src = src;
    this.img.alt = alt ?? '';
    this.setSize(width, height);
  }

  private setSize(width: number | null, height: number | null): void {
    if (width) this.img.setAttribute('width', String(Math.round(width)));
    else this.img.removeAttribute('width');
    if (height) this.img.setAttribute('height', String(Math.round(height)));
    else this.img.removeAttribute('height');
  }

  update(node: PMNode): boolean {
    if (node.type !== this.node.type) return false;
    this.node = node;
    if (!this.dragging) this.render();
    return true;
  }

  selectNode(): void {
    this.dom.classList.add('ProseMirror-selectednode');
  }

  deselectNode(): void {
    this.dom.classList.remove('ProseMirror-selectednode');
  }

  stopEvent(e: Event): boolean {
    return !!(e.target as HTMLElement | null)?.dataset?.handle;
  }

  ignoreMutation(): boolean {
    return true;
  }

  destroy(): void {
    this.cancelDrag?.();
  }

  private startDrag(e: PointerEvent, handle: Handle): void {
    if (!this.view.editable) return;
    e.preventDefault();
    e.stopPropagation();
    this.cancelDrag?.();
    const target = e.currentTarget as Element | null;
    // Keep receiving the pointer's events even when it leaves the handle or the window.
    try {
      if (e.pointerId != null) target?.setPointerCapture?.(e.pointerId);
    } catch {
      // not a pointer the element can capture
    }
    const rect = this.img.getBoundingClientRect();
    // The page view may be zoomed (CSS scale): pointer moves and client rects are in screen px.
    const zoom = rect.width && this.img.offsetWidth ? rect.width / this.img.offsetWidth : 1;
    // Start size in page (unzoomed) px.
    const w0 = this.node.attrs.width || this.img.offsetWidth || rect.width / zoom || this.img.naturalWidth || 100;
    const h0 = this.node.attrs.height || this.img.offsetHeight || rect.height / zoom || this.img.naturalHeight || 100;
    const x0 = e.clientX;
    const y0 = e.clientY;
    const sx = handle.includes('w') ? -1 : 1;
    const sy = handle.startsWith('n') ? -1 : 1;
    let w = w0;
    let h = h0;
    this.dragging = true;
    const win = this.dom.ownerDocument.defaultView ?? window;

    const move = (ev: Event) => {
      const p = ev as PointerEvent;
      const dx = ((p.clientX - x0) * sx) / zoom;
      const dy = ((p.clientY - y0) * sy) / zoom;
      if (handle === 'e') w = Math.max(MIN, w0 + dx);
      else if (handle === 's') h = Math.max(MIN, h0 + dy);
      else {
        // Corners: follow the pointer, keeping the proportions.
        const scale = Math.max(MIN / w0, MIN / h0, ((w0 + dx) / w0 + (h0 + dy) / h0) / 2);
        w = w0 * scale;
        h = h0 * scale;
      }
      this.setSize(w, h);
    };
    const stop = () => {
      win.removeEventListener('pointermove', move);
      win.removeEventListener('pointerup', up);
      win.removeEventListener('pointercancel', up);
      target?.removeEventListener('lostpointercapture', up);
      this.dragging = false;
      this.cancelDrag = null;
    };
    const up = () => {
      stop();
      if (this.view.isDestroyed) return;
      const pos = this.getPos();
      const width = Math.round(w);
      const height = Math.round(h);
      if (pos == null || (width === Math.round(w0) && height === Math.round(h0))) {
        this.render();
        return;
      }
      const tr = this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, width, height }, this.node.marks);
      this.view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
    };
    win.addEventListener('pointermove', move);
    win.addEventListener('pointerup', up);
    win.addEventListener('pointercancel', up);
    target?.addEventListener('lostpointercapture', up);
    this.cancelDrag = stop;
  }
}

export const imageNodeView = (node: PMNode, view: EditorView, getPos: () => number | undefined): NodeView =>
  new ImageView(node, view, getPos);

// ----- commands for the selected picture -----

function selected(state: EditorState): { pos: number; node: PMNode } | null {
  const sel = state.selection;
  return sel instanceof NodeSelection && sel.node.type.name === 'image' ? { pos: sel.from, node: sel.node } : null;
}

/** Change attributes of the selected picture (size, alt text ...); it stays selected. */
export const setImageAttrs = (attrs: Record<string, unknown>): Command => (state, dispatch) => {
  const img = selected(state);
  if (!img) return false;
  const next = { ...img.node.attrs, ...attrs };
  if (Object.keys(attrs).every((k) => img.node.attrs[k] === next[k])) return true;
  if (dispatch) {
    const tr = state.tr.setNodeMarkup(img.pos, undefined, next, img.node.marks);
    dispatch(tr.setSelection(NodeSelection.create(tr.doc, img.pos)));
  }
  return true;
};

/** Show another picture in the same place: the width stays, the height follows the new picture's proportions. */
export const replaceImage = (src: string, naturalWidth: number, naturalHeight: number): Command => (state, dispatch) => {
  const img = selected(state);
  // Only formats Word shows everywhere (readImageFile converts the others).
  if (!img || !/^data:image\/(png|jpeg|gif|bmp)[;,]/.test(src)) return false;
  const width = img.node.attrs.width || naturalWidth;
  const height = naturalWidth ? Math.round((width * naturalHeight) / naturalWidth) : img.node.attrs.height;
  return setImageAttrs({ src, width, height })(state, dispatch);
};

/** Picture formats every Word version shows; anything else is converted to PNG first. */
const WORD_IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp'];

/** Shown when a picture can be neither used nor converted. */
export const UNSUPPORTED_IMAGE = '無法使用這個圖片格式，請改用 PNG、JPEG、GIF 或 BMP 圖片';

/** Shown when a file that is not a picture is chosen, pasted or dropped as one: documents take no attachments. */
export const NOT_AN_IMAGE = '這裡只能插入圖片（PNG、JPEG、GIF、BMP、WebP、SVG）。PDF 等附件請用公文／流程系統附加。';

/** Signatures of common non-picture files (PDF, ZIP / Office, old Office), whatever type they claim. */
const NOT_PICTURE_SIGNATURES = [[0x25, 0x50, 0x44, 0x46], [0x50, 0x4b, 0x03, 0x04], [0xd0, 0xcf, 0x11, 0xe0]];

/**
 * Whether a file is a picture: its bytes say so, or it claims a picture type and its bytes are
 * not a known document (a PDF renamed .png is not a picture).
 */
export async function isImageFile(file: Blob): Promise<boolean> {
  const head = new Uint8Array(await readAs<ArrayBuffer>(file.slice(0, 1024), 'readAsArrayBuffer'));
  if (sniffImageExt(head)) return true;
  if (NOT_PICTURE_SIGNATURES.some((sig) => sig.every((x, i) => head[i] === x))) return false;
  return (file.type || '').toLowerCase().startsWith('image/');
}

const readAs = <T>(file: Blob, how: 'readAsDataURL' | 'readAsArrayBuffer') =>
  new Promise<T>((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as T);
    r.onerror = () => reject(r.error);
    r[how](file);
  });

/** Load a picture; null when the browser cannot show it. */
function loadImage(src: string): Promise<HTMLImageElement | null> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => resolve(null);
    img.src = src;
  });
}

// ----- pictures as they are inserted (persona-300 A-9) -----
// A phone photo is 4000 × 3000 px and 3–8 MB: a few of them made a document too big to save
// (the server takes 50 MB). As Word's 「壓縮圖片」 does for print (220 ppi), a picture is made
// at most MAX_IMAGE_SIDE px on its long side (an A4 page's text width at ~220 ppi), photos
// are saved as JPEG (quality 0.85) and pictures with transparency or few colours stay PNG. A
// photo turned by its camera (EXIF orientation) is turned for real, and its EXIF data (GPS
// position included) is left out. Nothing is made bigger: a picture that would not get
// smaller is used as it is.

/** Longest side, px, of an inserted picture (≈ 21 cm − margins at 220 ppi). */
export const MAX_IMAGE_SIDE = 2000;
/** JPEG quality of re-encoded photos. */
const JPEG_QUALITY = 0.85;
/** Distinct colours (in a sample) from which a picture counts as a photo, not a drawing. */
const MANY_COLOURS = 8192;
/** A picture of many colours becomes JPEG only when its PNG is this many times bigger. */
const PNG_MUCH_BIGGER = 2;
/** A JPEG or PNG this small (bytes) and not too large in px is used as it is. */
const SMALL_FILE = 1024 * 1024;

/** The EXIF orientation of a JPEG (1–8; 1 when it has none). */
export function jpegOrientation(b: Uint8Array): number {
  if (b[0] !== 0xff || b[1] !== 0xd8) return 1;
  let i = 2;
  while (i + 4 < b.length && b[i] === 0xff) {
    const marker = b[i + 1];
    const len = (b[i + 2] << 8) | b[i + 3];
    if (marker === 0xda || len < 2) break; // image data: no more headers
    // APP1 "Exif" + two zero bytes, then a TIFF header and IFD0.
    if (marker === 0xe1 && b[i + 4] === 0x45 && b[i + 5] === 0x78 && b[i + 6] === 0x69 && b[i + 7] === 0x66) {
      const t = i + 10;
      const le = b[t] === 0x49;
      const u16 = (o: number) => (le ? b[t + o] | (b[t + o + 1] << 8) : (b[t + o] << 8) | b[t + o + 1]);
      const u32 = (o: number) => (le ? u16(o) + u16(o + 2) * 65536 : u16(o) * 65536 + u16(o + 2));
      const ifd = u32(4);
      const count = u16(ifd);
      for (let k = 0; k < count; k++) {
        const e = ifd + 2 + k * 12;
        if (t + e + 10 > b.length) break;
        if (u16(e) === 0x0112) {
          const v = u16(e + 8);
          return v >= 1 && v <= 8 ? v : 1;
        }
      }
      return 1;
    }
    i += 2 + len;
  }
  return 1;
}

export interface ImagePlan {
  /** Draw the picture again (smaller, turned, or in another format). */
  encode: boolean;
  /** The format to save it in: 'auto' = PNG, or JPEG for a photo (see encodeImage). */
  format: 'jpeg' | 'png' | 'auto';
  width: number;
  height: number;
}

/**
 * What to do with an inserted picture: `type` its format by its bytes ('other' for formats
 * Word may not show), `bytes` its file size, `width` × `height` its size as shown (turned).
 */
export function planImage(p: { type: string; bytes: number; width: number; height: number; orientation?: number }): ImagePlan {
  const scale = Math.min(1, MAX_IMAGE_SIDE / Math.max(1, p.width, p.height));
  const width = Math.max(1, Math.round(p.width * scale));
  const height = Math.max(1, Math.round(p.height * scale));
  const smaller = scale < 1;
  const turned = (p.orientation ?? 1) > 1;
  switch (p.type) {
    case 'image/gif': // may be animated: as it is
      return { encode: false, format: 'png', width: p.width, height: p.height };
    case 'image/jpeg':
      return { encode: smaller || turned || p.bytes > SMALL_FILE, format: 'jpeg', width, height };
    case 'image/png':
      return { encode: smaller || p.bytes > SMALL_FILE, format: 'auto', width, height };
    default: // BMP (uncompressed), and WebP, HEIC, SVG … Word may not show
      return { encode: true, format: 'auto', width, height };
  }
}

/** A decoded picture, turned as its camera meant (EXIF orientation). */
interface Decoded {
  source: CanvasImageSource;
  width: number;
  height: number;
  close(): void;
  /** The file as a data URL, when it had to be read as one to be decoded (no object URLs). */
  dataUrl?: string;
}

/**
 * Decodes a picture file straight from its bytes (no copy of it as text). `size`, when given,
 * is the size to decode it at (a huge picture is then never decoded at full size, where the
 * browser can do that); the canvas draws it at the planned size either way.
 */
async function decodeImage(file: Blob, mime: string, size?: { width: number; height: number }): Promise<Decoded | null> {
  // createImageBitmap applies the EXIF orientation where the browser supports asking for it.
  if (typeof createImageBitmap === 'function') {
    try {
      const options: Record<string, unknown> = { imageOrientation: 'from-image' };
      if (size) Object.assign(options, { resizeWidth: size.width, resizeHeight: size.height, resizeQuality: 'high' });
      const bmp = await createImageBitmap(file, options as ImageBitmapOptions);
      if (bmp.width && bmp.height) return { source: bmp, width: bmp.width, height: bmp.height, close: () => bmp.close?.() };
    } catch {
      // not decodable this way (SVG …): an <img> below
    }
  }
  // An <img> of the file (an SVG is shown only with its type).
  const typed = mime && file.type !== mime ? new Blob([file], { type: mime }) : file;
  const objectUrl = typeof URL.createObjectURL === 'function' ? URL.createObjectURL(typed) : null;
  const dataUrl = objectUrl ? undefined : await readAs<string>(typed, 'readAsDataURL');
  const img = await loadImage(objectUrl ?? dataUrl!);
  const close = () => {
    if (objectUrl) URL.revokeObjectURL(objectUrl);
  };
  if (!img || !img.naturalWidth || !img.naturalHeight) {
    close();
    return null;
  }
  // Browsers show (and draw) an <img> turned by its EXIF orientation.
  return { source: img, width: img.naturalWidth, height: img.naturalHeight, close, dataUrl };
}

/** The size in px a PNG, JPEG, GIF, BMP or WebP file's header gives (as stored, before any EXIF turn); null when unknown. */
export function headerImageSize(b: Uint8Array, ext: string | null): { width: number; height: number } | null {
  const u16be = (i: number) => (b[i] << 8) | b[i + 1];
  const u16le = (i: number) => b[i] | (b[i + 1] << 8);
  const u24le = (i: number) => b[i] | (b[i + 1] << 8) | (b[i + 2] << 16);
  const u32be = (i: number) => u16be(i) * 65536 + u16be(i + 2);
  const i32le = (i: number) => b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24);
  let width = 0;
  let height = 0;
  if (ext === 'png' && b.length >= 24) [width, height] = [u32be(16), u32be(20)];
  else if (ext === 'gif' && b.length >= 10) [width, height] = [u16le(6), u16le(8)];
  else if (ext === 'bmp' && b.length >= 26) [width, height] = [Math.abs(i32le(18)), Math.abs(i32le(22))];
  else if (ext === 'webp' && b.length >= 30) {
    const chunk = String.fromCharCode(b[12], b[13], b[14], b[15]);
    if (chunk === 'VP8 ') [width, height] = [u16le(26) & 0x3fff, u16le(28) & 0x3fff];
    else if (chunk === 'VP8L') {
      const bits = (b[21] | (b[22] << 8) | (b[23] << 16) | (b[24] << 24)) >>> 0;
      [width, height] = [(bits & 0x3fff) + 1, ((bits >>> 14) & 0x3fff) + 1];
    } else if (chunk === 'VP8X') [width, height] = [1 + u24le(24), 1 + u24le(27)];
  } else if (ext === 'jpeg') {
    let i = 2;
    while (i + 9 < b.length && b[i] === 0xff) {
      const marker = b[i + 1];
      // A start of frame (C0–CF but DHT C4, JPG C8, DAC CC): height, then width.
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        [height, width] = [u16be(i + 5), u16be(i + 7)];
        break;
      }
      if (marker === 0xda) break;
      i += 2 + u16be(i + 2);
    }
  }
  return width > 0 && height > 0 ? { width, height } : null;
}

/** Whether a file is HEIC / HEIF / AVIF (an ISO "ftyp" box of those brands): photo formats. */
function isHeifFile(b: Uint8Array): boolean {
  const text = (from: number, to: number) => String.fromCharCode(...b.subarray(from, to));
  return b.length >= 12 && text(4, 8) === 'ftyp' && /^(heic|heix|hevc|hevx|heim|heis|mif1|msf1|avif|avis)$/.test(text(8, 12));
}

/**
 * What a drawn picture's pixels say (sampled): whether any are see-through, and whether it has
 * many colours (a photo) or few (a screenshot, scan, diagram or text: those stay PNG, whose
 * sharp edges JPEG would blur). Unknown counts as see-through.
 */
function pixelStats(ctx: CanvasRenderingContext2D, width: number, height: number): { transparent: boolean; manyColours: boolean } {
  try {
    const data = ctx.getImageData(0, 0, width, height).data;
    const pixels = Math.floor(data.length / 4);
    const step = Math.max(1, Math.floor(pixels / 65536));
    const enough = Math.min(MANY_COLOURS, Math.max(256, Math.ceil(pixels / step / 8)));
    const colours = new Set<number>();
    for (let p = 0; p < pixels; p += step) {
      const i = p * 4;
      if (data[i + 3] < 255) return { transparent: true, manyColours: false };
      if (colours.size < enough) colours.add((data[i] << 16) | (data[i + 1] << 8) | data[i + 2]);
    }
    return { transparent: false, manyColours: colours.size >= enough };
  } catch {
    return { transparent: true, manyColours: false };
  }
}

/**
 * The picture drawn at `plan`'s size, as a data URL; null when the browser can't. A JPEG stays
 * JPEG; anything else becomes PNG, and JPEG only when it is a photo: in a photo format (HEIC …),
 * or of many colours with a PNG much bigger than the JPEG. See-through pictures are always PNG.
 */
function encodeImage(d: Decoded, plan: ImagePlan, photoFormat: boolean): string | null {
  try {
    const canvas = document.createElement('canvas');
    canvas.width = plan.width;
    canvas.height = plan.height;
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;
    if ('imageSmoothingQuality' in ctx) ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(d.source, 0, 0, plan.width, plan.height);
    // A browser that can't write JPEG gives PNG: still a picture Word shows.
    const valid = (url: string) => (/^data:image\/(png|jpeg)[;,]/.test(url) ? url : null);
    const jpeg = () => {
      if (typeof ctx.fillRect === 'function') {
        // A JPEG has no transparency: what was see-through becomes white, not black.
        ctx.globalCompositeOperation = 'destination-over';
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, plan.width, plan.height);
      }
      return valid(canvas.toDataURL('image/jpeg', JPEG_QUALITY));
    };
    if (plan.format === 'jpeg') return jpeg();
    const stats = pixelStats(ctx, plan.width, plan.height);
    if (stats.transparent) return valid(canvas.toDataURL('image/png'));
    if (photoFormat && plan.format === 'auto') return jpeg();
    const png = valid(canvas.toDataURL('image/png'));
    if (plan.format === 'png' || !stats.manyColours) return png;
    const photo = jpeg();
    return photo && (!png || dataUrlBytes(png) > PNG_MUCH_BIGGER * dataUrlBytes(photo)) ? photo : png;
  } catch {
    return null;
  }
}

/** Bytes of a data URL's content. */
export function dataUrlBytes(src: string): number {
  const comma = src.indexOf(',');
  if (comma < 0) return 0;
  const body = src.length - comma - 1;
  if (!/;base64$/i.test(src.slice(0, comma))) return body;
  const pad = src.endsWith('==') ? 2 : src.endsWith('=') ? 1 : 0;
  return Math.max(0, Math.floor((body * 3) / 4) - pad);
}

/**
 * Read a picture file as a data URL with its size in px, ready for the document: PNG, JPEG,
 * GIF and BMP as they are unless making them smaller helps (see planImage); other formats
 * (WebP, AVIF, HEIC, SVG ...) converted to PNG or JPEG. Throws UNSUPPORTED_IMAGE when that is
 * not possible. A picture drawn again is decoded from the file itself (a huge upright one at
 * the size it will have, where the browser can); the file is read as text only when it is
 * used as it is.
 */
export async function readImageFile(file: Blob): Promise<{ src: string; width: number; height: number }> {
  const head = new Uint8Array(await readAs<ArrayBuffer>(file.slice(0, 65536), 'readAsArrayBuffer'));
  // The bytes decide, not the file name or its claimed type.
  const sniffed = sniffImageExt(head);
  const type = sniffed ? `image/${sniffed}` : file.type || '';
  const wordType = WORD_IMAGE_TYPES.includes(type);
  const planType = wordType ? type : 'other';
  const asItIs = async (read?: string) => {
    const src = read ?? (await readAs<string>(file, 'readAsDataURL'));
    return wordType ? src.replace(/^data:[^;,]*/, 'data:' + type) : src;
  };
  const orientation = type === 'image/jpeg' ? jpegOrientation(head) : 1;
  const stored = headerImageSize(head, sniffed);
  const target = stored && orientation === 1 ? planImage({ type: planType, bytes: file.size, ...stored }) : null;
  const resize = stored && target?.encode && target.width < stored.width ? target : undefined;
  const decoded = await decodeImage(file, sniffed === 'svg' ? 'image/svg+xml' : type, resize);
  if (!decoded) {
    if (wordType) return { src: await asItIs(), width: 200, height: 150 };
    throw new Error(UNSUPPORTED_IMAGE);
  }
  try {
    // Decoded smaller: the header's size is the picture's own.
    const natural = resize ? stored! : { width: decoded.width, height: decoded.height };
    const plan = planImage({ type: planType, bytes: file.size, ...natural, orientation });
    if (plan.encode) {
      const out = encodeImage(decoded, plan, type === 'image/jpeg' || isHeifFile(head));
      // Kept when smaller, or when needed: a format Word may not show, a photo to turn.
      if (out && (!wordType || orientation > 1 || dataUrlBytes(out) < file.size)) return { src: out, width: plan.width, height: plan.height };
    }
    if (!wordType) throw new Error(UNSUPPORTED_IMAGE);
    return { src: await asItIs(decoded.dataUrl), ...natural };
  } finally {
    decoded.close();
  }
}

/** What Word cuts off each side of a picture (a:srcRect), as fractions of it (see readDrawing). */
export interface PictureCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * A picture turned 90° clockwise (圖片 › 向右旋轉 90°), drawn again in its own format (JPEG stays
 * JPEG, anything else becomes PNG). A picture Word shows cropped is cropped first: the new
 * picture is what Word showed, turned (its crop is not kept, see patchDrawing). Null when the
 * browser can't draw it (EMF, TIFF … from a Word file).
 */
export async function rotateImageRight(src: string, crop: PictureCrop | null = null): Promise<string | null> {
  const img = await loadImage(src);
  const iw = img?.naturalWidth ?? 0;
  const ih = img?.naturalHeight ?? 0;
  if (!img || !iw || !ih) return null;
  const sx = Math.round(iw * (crop?.left ?? 0));
  const sy = Math.round(ih * (crop?.top ?? 0));
  const w = Math.max(1, Math.round(iw * (1 - (crop?.right ?? 0))) - sx);
  const h = Math.max(1, Math.round(ih * (1 - (crop?.bottom ?? 0))) - sy);
  try {
    const canvas = document.createElement('canvas');
    canvas.width = h;
    canvas.height = w;
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;
    ctx.translate(h, 0);
    ctx.rotate(Math.PI / 2);
    ctx.drawImage(img, sx, sy, w, h, 0, 0, w, h);
    const jpeg = /^data:image\/jpeg[;,]/i.test(src);
    const out = canvas.toDataURL(jpeg ? 'image/jpeg' : 'image/png', jpeg ? 0.92 : undefined);
    return /^data:image\/(png|jpeg)[;,]/.test(out) ? out : null;
  } catch {
    return null;
  }
}

/** Show the selected picture turned 90° clockwise (`src`, see rotateImageRight): width and height swap, the alt text stays. */
export const rotateImage = (src: string): Command => (state, dispatch) => {
  const img = selected(state);
  if (!img || !/^data:image\/(png|jpeg)[;,]/.test(src)) return false;
  const { width, height } = img.node.attrs;
  return setImageAttrs({ src, width: height, height: width })(state, dispatch);
};
