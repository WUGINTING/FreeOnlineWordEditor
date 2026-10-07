// Theme colours (word/theme/theme1.xml) and DrawingML colour elements (a:srgbClr, a:schemeClr ...
// with a:shade, a:tint, a:lumMod ...), for drawing shapes as Word does.

import { NS } from './xml';

/** Theme colour name → RRGGBB. */
export type ThemeColors = Record<string, string>;

/** Word's default theme (Office 2013–2022), used when the file has none. */
export const DEFAULT_THEME: ThemeColors = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
  accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
  hlink: '0563C1', folHlink: '954F72',
};

/** The colours of a theme part (a:theme/a:themeElements/a:clrScheme); the default theme where it has none. */
export function readThemeColors(theme: Document | null): ThemeColors {
  const out: ThemeColors = { ...DEFAULT_THEME };
  const scheme = theme?.getElementsByTagNameNS(NS.a, 'clrScheme')[0];
  if (!scheme) return out;
  for (let c = scheme.firstElementChild; c; c = c.nextElementSibling) {
    const v = c.firstElementChild;
    if (!v) continue;
    const hex = v.localName === 'sysClr' ? v.getAttribute('lastClr') : v.getAttribute('val');
    if (hex && /^[0-9A-Fa-f]{6}$/.test(hex)) out[c.localName] = hex.toUpperCase();
  }
  return out;
}

/** Colour names a shape uses for its theme's text and background colours. */
const ALIAS: Record<string, string> = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };

const PRESET: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080',
  grey: '808080', lightGray: 'D3D3D3', darkGray: 'A9A9A9', orange: 'FFA500', purple: '800080', navy: '000080',
  silver: 'C0C0C0', maroon: '800000', olive: '808000', lime: '00FF00', aqua: '00FFFF', teal: '008080', fuchsia: 'FF00FF',
};

export interface Color {
  hex: string;
  /** 0..1, absent when opaque. */
  alpha?: number;
}

const clamp = (n: number) => Math.max(0, Math.min(255, Math.round(n)));

function rgb(hex: string): [number, number, number] {
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

function hex([r, g, b]: [number, number, number]): string {
  return [r, g, b].map((v) => clamp(v).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function toHsl([r, g, b]: [number, number, number]): [number, number, number] {
  const [R, G, B] = [r / 255, g / 255, b / 255];
  const max = Math.max(R, G, B);
  const min = Math.min(R, G, B);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === R ? (G - B) / d + (G < B ? 6 : 0) : max === G ? (B - R) / d + 2 : (R - G) / d + 4;
  return [h / 6, s, l];
}

function fromHsl([h, s, l]: [number, number, number]): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const ch = (t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [ch(h + 1 / 3) * 255, ch(h) * 255, ch(h - 1 / 3) * 255];
}

/**
 * The colour of a DrawingML colour element (the element itself, e.g. a:srgbClr), with its
 * modifiers applied. `placeholder` stands for a:schemeClr val="phClr" (a style's own colour).
 */
export function readColor(el: Element | null, theme: ThemeColors, placeholder?: Color | null): Color | null {
  if (!el || el.namespaceURI !== NS.a) return null;
  let base: string | null = null;
  let alpha: number | undefined;
  const val = el.getAttribute('val') ?? '';
  switch (el.localName) {
    case 'srgbClr':
      base = val;
      break;
    case 'schemeClr':
      if (val === 'phClr') {
        base = placeholder?.hex ?? null;
        alpha = placeholder?.alpha;
      } else base = theme[ALIAS[val] ?? val] ?? DEFAULT_THEME[ALIAS[val] ?? val] ?? null;
      break;
    case 'prstClr':
      base = PRESET[val] ?? null;
      break;
    case 'sysClr':
      base = el.getAttribute('lastClr') ?? (val === 'window' ? 'FFFFFF' : '000000');
      break;
    case 'scrgbClr': {
      const pc = (a: string) => (Number(el.getAttribute(a) ?? 0) / 100000) * 255;
      base = hex([pc('r'), pc('g'), pc('b')]);
      break;
    }
    case 'hslClr': {
      const h = Number(el.getAttribute('hue') ?? 0) / 21600000;
      const s = Number(el.getAttribute('sat') ?? 0) / 100000;
      const l = Number(el.getAttribute('lum') ?? 0) / 100000;
      base = hex(fromHsl([h, s, l]));
      break;
    }
    default:
      return null;
  }
  if (!base || !/^[0-9A-Fa-f]{6}$/.test(base)) return null;
  let c = rgb(base);
  for (let m = el.firstElementChild; m; m = m.nextElementSibling) {
    const v = Number(m.getAttribute('val') ?? 0) / 100000;
    switch (m.localName) {
      case 'shade':
        c = c.map((x) => x * v) as [number, number, number];
        break;
      case 'tint':
        c = c.map((x) => x + (255 - x) * (1 - v)) as [number, number, number];
        break;
      case 'lumMod':
      case 'lumOff': {
        const hsl = toHsl(c);
        hsl[2] = m.localName === 'lumMod' ? hsl[2] * v : hsl[2] + v;
        hsl[2] = Math.max(0, Math.min(1, hsl[2]));
        c = fromHsl(hsl);
        break;
      }
      case 'satMod': {
        const hsl = toHsl(c);
        hsl[1] = Math.max(0, Math.min(1, hsl[1] * v));
        c = fromHsl(hsl);
        break;
      }
      case 'alpha':
        alpha = v;
        break;
    }
  }
  const out: Color = { hex: hex(c) };
  if (alpha != null && alpha < 1) out.alpha = Math.max(0, alpha);
  return out;
}

/** The colour of the first colour element inside `parent` (a:solidFill, a:fillRef ...). */
export function colorIn(parent: Element | null, theme: ThemeColors, placeholder?: Color | null): Color | null {
  for (let c = parent?.firstElementChild ?? null; c; c = c.nextElementSibling) {
    const color = readColor(c, theme, placeholder);
    if (color) return color;
  }
  return null;
}

/** A VML colour (fillcolor="#ff0000", "red", "window [65]" ...) as RRGGBB; null when unreadable. */
export function vmlColor(v: string | null | undefined): string | null {
  if (!v) return null;
  const s = v.trim().split(/\s/)[0];
  const m = /^#?([0-9a-f]{6}|[0-9a-f]{3})$/i.exec(s);
  if (m) return (m[1].length === 3 ? m[1].replace(/./g, (x) => x + x) : m[1]).toUpperCase();
  return PRESET[s] ?? PRESET[s.toLowerCase()] ?? (s === 'window' ? 'FFFFFF' : s === 'windowText' ? '000000' : null);
}
