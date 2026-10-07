// Word's 浮水印 (設計 › 頁面背景 › 浮水印): a VML shape in a header part that Word draws centred on
// every page showing that header, behind the text. A text watermark is WordArt (the text-path
// shape type #_x0000_t136) named PowerPlusWaterMarkObject<n>; a picture watermark is a picture
// frame (#_x0000_t75) named WordPictureWatermark<n>. Word's own 浮水印 dialog finds, changes and
// removes watermarks by those names, so they are written exactly as Word writes them.

import { escapeAttr, parseFragment } from './xml';
import { cssFontName } from './fontNames';
import { Lru, deepFreeze, memo } from './memo';

const TEXT_NAME = 'PowerPlusWaterMarkObject';
const PICTURE_NAME = 'WordPictureWatermark';
/** Quick test on kept XML: a watermark shape's id. */
const NAMED = /\bid="(?:PowerPlusWaterMarkObject|WordPictureWatermark)/;

/**
 * The watermarks the 浮水印 menu offers: Word's zh-TW gallery (機密, 請勿複製, 草稿, 樣本, 緊急) and
 * the markings Taiwan 公文 carry (密, 限閱, 內部傳閱).
 */
export const WATERMARK_PRESETS = ['機密', '請勿複製', '草稿', '樣本', '緊急', '密', '限閱', '內部傳閱'] as const;
/** The font a new text watermark gets: 公文 are set in 標楷體. */
export const WATERMARK_FONT = '標楷體';
/** Word's default watermark colour, 銀色 (VML "silver"). */
export const WATERMARK_COLOR = '#c0c0c0';
/** The sizes (pt) Word's 浮水印 dialog lists besides 自動. */
export const WATERMARK_SIZES = [36, 40, 44, 48, 54, 60, 66, 72, 80, 90, 96, 105, 120, 144];
/** The scales (%) Word's 浮水印 dialog lists for a picture besides 自動. */
export const WATERMARK_SCALES = [500, 200, 150, 100, 50];
/**
 * The relationship id a new picture watermark is written with until it is saved: the writer
 * embeds the picture in each header part it is in and puts that part's own id in its place.
 */
export const WATERMARK_IMAGE_REF = 'rIdPxWatermarkImage';

/** A text watermark, as Word's 浮水印 dialog sets it. */
export interface TextWatermark {
  kind: 'text';
  text: string;
  /** Font name as Word writes it (「標楷體」). */
  font: string;
  /** Font size in pt; null is 自動 (as large as fits the page). */
  size: number | null;
  /** "#rrggbb". */
  color: string;
  /** 半透明: drawn at half opacity. */
  semitransparent: boolean;
  /** 斜向 (rising at 45°) or 水平. */
  layout: 'diagonal' | 'horizontal';
}

/** A picture watermark, as Word's 浮水印 dialog sets it. */
export interface PictureWatermark {
  kind: 'picture';
  /** The picture, a data: URL. */
  src: string;
  /** Its size in px at 100 %. */
  width: number;
  height: number;
  /** 縮放 in %; null is 自動 (as large as fits the text area). */
  scale: number | null;
  /** 刷淡 (washout): drawn light, as Word does. */
  washout: boolean;
  /** The picture's name (the chosen file's, without its extension), Word's o:title. */
  title?: string;
}

export type Watermark = TextWatermark | PictureWatermark;

/** A picture's name as Word's o:title has it: the file's name without its extension. */
export const pictureTitle = (fileName: string): string => oneLine(fileName.replace(/\.[A-Za-z0-9]{1,5}$/, ''));

/** A preset from the 浮水印 menu: diagonal, semi-transparent silver, sized to the page. */
export function textWatermark(text: string, font: string = WATERMARK_FONT): TextWatermark {
  return { kind: 'text', text, font, size: null, color: WATERMARK_COLOR, semitransparent: true, layout: 'diagonal' };
}

/** A watermark shape read from a header part: what the page shows. Lengths in pt. */
export interface WatermarkShape {
  kind: 'text' | 'picture';
  /** The text (empty for a picture). */
  text: string;
  font: string;
  /** The text's size in pt; null for 自動 (Word writes 1pt and fits the text to the shape). */
  size: number | null;
  /** "#rrggbb". */
  color: string;
  /** 0–1: the text's fill opacity. */
  opacity: number;
  /** Clockwise, in degrees (Word's diagonal watermark is 315). */
  rotation: number;
  width: number;
  height: number;
  /** Centred on the page instead of on the margins. */
  pageRelative: { horizontal: boolean; vertical: boolean };
  /** A picture drawn washed out (刷淡). */
  washout: boolean;
  /** A picture's name (o:title). */
  title: string;
}

/** Whether kept XML mentions a watermark shape Word recognises as one (anywhere in it). */
export function isWatermarkXml(xml: string | null | undefined): boolean {
  return !!xml && NAMED.test(xml);
}

/**
 * Whether kept XML is a watermark itself: a run (w:r) holding a watermark shape. A tracked
 * deletion that holds one (w:del) is not: it is a revision, shown and handled as one.
 */
export function isWatermarkRun(xml: string | null | undefined): boolean {
  return !!xml && RUN.test(xml) && NAMED.test(xml);
}
const RUN = /^<(?:[\w.-]+:)?r[\s>/]/;

const shapes = new Lru<WatermarkShape | null>(64);

/** The first watermark shape in kept XML, read for display; null when there is none. Cached by the XML. */
export function readWatermark(xml: string): WatermarkShape | null {
  if (!isWatermarkXml(xml)) return null;
  return memo(shapes, xml, () => deepFreeze(parseWatermark(xml)));
}

function parseWatermark(xml: string): WatermarkShape | null {
  let root: Element;
  try {
    root = parseFragment(xml);
  } catch {
    return null;
  }
  const shape = Array.from(root.getElementsByTagName('*')).find(
    (e) => e.localName === 'shape' && /^(?:PowerPlusWaterMarkObject|WordPictureWatermark)/.test(e.getAttribute('id') ?? ''),
  );
  if (!shape) return null;
  const style = cssPairs(shape.getAttribute('style'));
  const kid = (local: string) => Array.from(shape.children).find((c) => c.localName === local) ?? null;
  const textpath = kid('textpath');
  const imagedata = kid('imagedata');
  const picture = (shape.getAttribute('id') ?? '').startsWith(PICTURE_NAME);
  if (picture ? !imagedata : !textpath) return null;
  const path = cssPairs(textpath?.getAttribute('style') ?? null);
  const size = length(path['font-size']);
  const rotation = ((Number(style.rotation) || 0) % 360 + 360) % 360;
  const fill = kid('fill');
  const gain = fraction(imagedata?.getAttribute('gain') ?? null);
  const black = fraction(imagedata?.getAttribute('blacklevel') ?? null);
  return {
    kind: picture ? 'picture' : 'text',
    text: textpath?.getAttribute('string') ?? '',
    font: (path['font-family'] ?? '').replace(/^["']|["']$/g, '').trim(),
    // Word writes font-size:1pt for 自動: the text is stretched to the shape.
    size: size != null && size > 1 ? Math.round(size * 100) / 100 : null,
    color: vmlColor(shape.getAttribute('fillcolor')) ?? '#ffffff',
    opacity: Math.max(0, Math.min(1, fill?.getAttribute('on') === 'f' ? 0 : fraction(fill?.getAttribute('opacity') ?? null) ?? 1)),
    rotation,
    width: length(style.width) ?? 0,
    height: length(style.height) ?? 0,
    pageRelative: { horizontal: style['mso-position-horizontal-relative'] === 'page', vertical: style['mso-position-vertical-relative'] === 'page' },
    washout: (gain != null && gain < 1) || (black != null && black > 0),
    title: imagedata?.getAttribute('o:title') ?? '',
  };
}

/**
 * The watermark's settings as the 浮水印 dialog shows them (`src`: the picture shown for a picture
 * shape; `area`: the text area of a page showing it).
 */
export function watermarkSettings(shape: WatermarkShape, src: string | null, area?: WatermarkArea): Watermark | null {
  if (shape.kind === 'picture') {
    if (!src) return null;
    // 自動 fits the picture to the text area: as wide or as tall as it (Word's and ours).
    const auto = !!area && (Math.abs(shape.width - area.width) <= AUTO_TOLERANCE || Math.abs(shape.height - area.height) <= AUTO_TOLERANCE);
    // Otherwise shown at the size it has: 100 % of that size (the dialog works out the real scale).
    return {
      kind: 'picture', src, width: ptToPx(shape.width), height: ptToPx(shape.height), scale: auto ? null : 100, washout: shape.washout,
      ...(shape.title ? { title: shape.title } : {}),
    };
  }
  return {
    kind: 'text',
    text: shape.text,
    font: shape.font,
    size: shape.size,
    color: shape.color,
    semitransparent: shape.opacity < 1,
    layout: shape.rotation === 0 ? 'horizontal' : 'diagonal',
  };
}

/** Where a watermark goes: the text area of the page (between the margins), in pt. */
export interface WatermarkArea {
  width: number;
  height: number;
}

/** Numbers that make a new shape unique in the document. */
export interface WatermarkIds {
  /** The n of PowerPlusWaterMarkObject<n> / WordPictureWatermark<n>. */
  number: number;
  /** The n of o:spid="_x0000_s<n>". */
  spid: number;
  /** Whether the part still needs the shape type's definition (Word writes it once per part). */
  shapetype: boolean;
  /** Its z-index: Word gives each shape its own, 1024 below the last (see WATERMARK_Z_INDEX). */
  zIndex: number;
}

/** The z-index of the first watermark shape Word writes; each next one is 1024 lower. */
export const WATERMARK_Z_INDEX = -251657216;

/** A text's width and height (its font's ascent + descent) in em, measured in its font. */
export interface TextMetrics {
  width: number;
  height: number;
  /** How much of the height is above the baseline (0–1); unknown when not measured. */
  ascent?: number;
}

/**
 * The run Word writes for a watermark, placed in a paragraph of each header part:
 * `<w:r><w:rPr><w:noProof/></w:rPr><w:pict>…</w:pict></w:r>`. A picture's v:imagedata refers to
 * WATERMARK_IMAGE_REF until it is saved.
 */
export function watermarkRunXml(w: Watermark, area: WatermarkArea, ids: WatermarkIds, metrics?: TextMetrics | null): string {
  const raw = w.kind === 'text' ? textWatermarkSize(w, area, metrics) : pictureWatermarkSize(w, area);
  // At least 1 pt each way; a size that can't be worked out (NaN) is the text area's.
  const side = (n: number, fallback: number) => Math.max(1, Number.isFinite(n) ? n : Number.isFinite(fallback) ? fallback : 1);
  const box = { width: side(raw.width, area.width), height: side(raw.height, area.height / 4) };
  const style = [
    'position:absolute',
    'margin-left:0',
    'margin-top:0',
    `width:${pt(box.width)}pt`,
    `height:${pt(box.height)}pt`,
    ...(w.kind === 'text' && w.layout === 'diagonal' ? ['rotation:315'] : []),
    `z-index:${Number.isFinite(ids.zIndex) ? Math.round(ids.zIndex) : WATERMARK_Z_INDEX}`,
    'mso-position-horizontal:center',
    'mso-position-horizontal-relative:margin',
    'mso-position-vertical:center',
    'mso-position-vertical-relative:margin',
  ].join(';');
  let shape: string;
  if (w.kind === 'text') {
    const size = w.size && Number.isFinite(w.size) && w.size > 1 ? `${pt(w.size)}pt` : '1pt';
    shape =
      (ids.shapetype ? TEXT_SHAPETYPE : '') +
      `<v:shape id="${TEXT_NAME}${ids.number}" o:spid="_x0000_s${ids.spid}" type="#_x0000_t136" style="${style}" o:allowincell="f" ` +
      `fillcolor="${vmlFill(w.color)}" stroked="f">` +
      (w.semitransparent ? '<v:fill opacity=".5"/>' : '') +
      `<v:textpath style="font-family:&quot;${escapeAttr(cssFontName(w.font) || WATERMARK_FONT)}&quot;;font-size:${size}" string="${escapeAttr(oneLine(w.text))}"/>`;
  } else {
    shape =
      (ids.shapetype ? PICTURE_SHAPETYPE : '') +
      `<v:shape id="${PICTURE_NAME}${ids.number}" o:spid="_x0000_s${ids.spid}" type="#_x0000_t75" style="${style}" o:allowincell="f">` +
      `<v:imagedata r:id="${WATERMARK_IMAGE_REF}" o:title="${escapeAttr(oneLine(w.title ?? ''))}"${w.washout ? ' gain="19661f" blacklevel="22938f"' : ''}/>`;
  }
  return `<w:r><w:rPr><w:noProof/></w:rPr><w:pict>${shape}<w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>`;
}

/** The shape type a watermark of this kind uses (its id, to write the definition once per part). */
export const watermarkShapetype = (w: Watermark): string => (w.kind === 'text' ? '_x0000_t136' : '_x0000_t75');

/**
 * The text's box. Word stretches the text to its shape (fitshape), so the box has the text's own
 * proportions: its width and its font's height (ascent + descent), as `metrics` measures them in
 * the font; without them (no canvas, font not loaded) they are estimated (full-width characters
 * one em, others about 0.6 em; the height 1/0.9 em). 自動 makes the box as large as the text
 * area allows: a diagonal one, turned, spans the text area's width, as Word's does. A size in pt
 * gives the box of the text at that size.
 */
export function textWatermarkSize(w: TextWatermark, area: WatermarkArea, metrics?: TextMetrics | null): { width: number; height: number } {
  let ems = 0;
  for (const ch of oneLine(w.text) || ' ') ems += FULL_WIDTH.test(ch) ? 1 : ch === ' ' ? 0.3 : 0.6;
  const measured = metrics && metrics.width > 0 && metrics.height > 0 ? metrics : null;
  const width = measured ? measured.width : Math.max(ems, 0.6);
  const height = measured ? measured.height : BOX_HEIGHT;
  if (w.size) return { width: width * w.size, height: height * w.size };
  const scale = w.layout === 'diagonal'
    ? (Math.min(area.width, area.height) * Math.SQRT2) / (width + height)
    : Math.min(area.width / width, area.height / height);
  return { width: width * scale, height: height * scale };
}

/** The picture's box: its size (1 px = ¾ pt) times the scale, or fitted to the text area (自動). */
export function pictureWatermarkSize(w: PictureWatermark, area: WatermarkArea): { width: number; height: number } {
  const width = Math.max(1, w.width) * 0.75;
  const height = Math.max(1, w.height) * 0.75;
  const scale = w.scale ? w.scale / 100 : Math.min(area.width / width, area.height / height);
  return { width: width * scale, height: height * scale };
}

/**
 * How tall a watermark's text box is for its font size when the font can't be measured: the
 * text is then drawn at 0.9 of the box (see editor/watermark.ts), so a full-width character is
 * as tall as it is wide.
 */
const BOX_HEIGHT = 1 / 0.9;
/** How near (pt) a picture's size must be to the text area's to count as 自動. */
const AUTO_TOLERANCE = 1;
/** East Asian ideographs, kana, Hangul and full-width forms: one em wide. */
const FULL_WIDTH = /[\u2E80-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFF60\uFFE0-\uFFE6]/;

const pt = (n: number) => String(Math.round(n * 100) / 100);
const ptToPx = (n: number) => Math.round(((n * 96) / 72) * 100) / 100;
/** A watermark is one line of text. */
const oneLine = (s: string) => s.replace(/[\r\n\t]+/g, ' ').trim();

/** "a:b;c:d" → { a: 'b', c: 'd' } (VML style attributes). */
function cssPairs(style: string | null): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of (style ?? '').split(';')) {
    const i = part.indexOf(':');
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

/** A VML length ("415.4pt", "5.5in", "300px", "12cm") in pt; null when unreadable. */
function length(v: string | undefined): number | null {
  const m = /^(-?[\d.]+)\s*(pt|px|in|cm|mm|pc)?$/.exec(v ?? '');
  if (!m) return null;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return null;
  const per: Record<string, number> = { pt: 1, px: 0.75, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, pc: 12 };
  return n * (per[m[2] ?? 'px'] ?? 1);
}

/** A VML fraction (".5", "0.5", "32768f" = 32768 / 65536); null when absent or unreadable. */
function fraction(v: string | null): number | null {
  if (v == null || v === '') return null;
  const n = v.endsWith('f') ? Number(v.slice(0, -1)) / 65536 : Number(v);
  return Number.isFinite(n) ? n : null;
}

/** The colour names VML uses (Word writes "silver" for its default watermark colour). */
const VML_COLORS: Record<string, string> = {
  black: '#000000', silver: '#c0c0c0', gray: '#808080', grey: '#808080', white: '#ffffff', maroon: '#800000',
  red: '#ff0000', purple: '#800080', fuchsia: '#ff00ff', green: '#008000', lime: '#00ff00', olive: '#808000',
  yellow: '#ffff00', navy: '#000080', blue: '#0000ff', teal: '#008080', aqua: '#00ffff',
};

/** A VML colour ("silver", "#c00000", "#c00000 [3204]") as "#rrggbb"; null when unreadable. */
export function vmlColor(v: string | null): string | null {
  const s = (v ?? '').trim().toLowerCase().split(/\s+/)[0] ?? '';
  if (VML_COLORS[s]) return VML_COLORS[s];
  const hex = /^#([0-9a-f]{6}|[0-9a-f]{3})$/.exec(s)?.[1];
  if (!hex) return null;
  return '#' + (hex.length === 3 ? hex.replace(/./g, (c) => c + c) : hex);
}

/** The fillcolor Word writes: "silver" for its default colour, "#rrggbb" otherwise. */
function vmlFill(color: string): string {
  const c = vmlColor(color) ?? WATERMARK_COLOR;
  return c === WATERMARK_COLOR ? 'silver' : c;
}

// Word's definitions of the two shape types, as it writes them in a header part.

const TEXT_SHAPETYPE =
  '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">' +
  '<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>' +
  '<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>' +
  '<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>' +
  '<v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>' +
  '<v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>' +
  '<o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>';

const PICTURE_SHAPETYPE =
  '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">' +
  '<v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/>' +
  '<v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/>' +
  '<v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/>' +
  '<v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>' +
  '<o:lock v:ext="edit" aspectratio="t"/></v:shapetype>';
