import { attr, child, children, numAttr, onOff, parseFragment } from './xml';
import { fontStack } from './fonts';
import { twipsToPx, halfPointsToPt } from '../units';
import { cssIdent, lineGridVars, lineHeightCss } from '../editor/schema';
import type { ParagraphStyleInfo } from './model';
import { emptyInheritance, jcToAlign, type StyleInheritance, type StyleProps } from './inheritance';
import { EAST_ASIAN_LINE, LATIN_LINE, wordLineRatio } from './fontMetrics';
import { cellMarginCss, DEFAULT_CELL_MARGINS } from './cellMargins';
import { cssFontName, eastAsiaAlias } from './fontNames';

export { cssFontName };

/**
 * Word's single line height for the fonts a style gives (see fontMetrics.ts), as CSS
 * variables: --dx-lh-l for Latin text, --dx-lh-e for East Asian text.
 */
function lineVars(props: StyleProps, css: Css): Css {
  const latin = wordLineRatio(props.fontFamily);
  const eastAsian = wordLineRatio(props.fontEastAsia);
  if (latin != null) css['--dx-lh-l'] = String(latin);
  if (eastAsian != null) css['--dx-lh-e'] = String(eastAsian);
  return css;
}

type Css = Record<string, string>;

interface RawStyle {
  id: string;
  name: string;
  type: string;
  basedOn: string | null;
  isDefault: boolean;
  own: Css;
  props: StyleProps;
  /** Table styles: cell margins (w:tblPr/w:tblCellMar) as CSS variables. */
  cellMar: Css;
}

/** The style-inheritable properties set directly in a w:pPr / w:rPr pair. */
function ownProps(pPr: Element | null, rPr: Element | null): StyleProps {
  const props: StyleProps = {};
  const align = jcToAlign(attr(child(pPr, 'jc'), 'val'));
  if (align) props.align = align;
  const b = onOff(child(rPr, 'b'));
  if (b != null) props.bold = b;
  const i = onOff(child(rPr, 'i'));
  if (i != null) props.italic = i;
  const sz = numAttr(child(rPr, 'sz'), 'val');
  if (sz != null) props.fontSize = halfPointsToPt(sz);
  const color = attr(child(rPr, 'color'), 'val');
  const hexColor = cssHex(color);
  if (hexColor) props.color = hexColor.toLowerCase();
  const outline = numAttr(child(pPr, 'outlineLvl'), 'val');
  if (outline != null) props.outlineLevel = outline;
  const fonts = child(rPr, 'rFonts');
  const latin = attr(fonts, 'ascii') ?? THEME_FONTS[attr(fonts, 'asciiTheme') ?? ''];
  const eastAsia = attr(fonts, 'eastAsia') ?? THEME_FONTS[attr(fonts, 'eastAsiaTheme') ?? ''];
  if (latin) props.fontFamily = latin;
  if (eastAsia) props.fontEastAsia = eastAsia;
  return props;
}

const THEME_FONTS: Record<string, string> = {
  minorHAnsi: 'Calibri',
  minorAscii: 'Calibri',
  majorHAnsi: 'Calibri Light',
  majorAscii: 'Calibri Light',
  minorEastAsia: 'PMingLiU',
  majorEastAsia: 'PMingLiU',
};

// Values from the file go into a <style> element: only let through what a font name or a
// colour can be, so a crafted document cannot add CSS of its own.

/** "RRGGBB" → "#RRGGBB"; anything else (auto, garbage) → null. */
export function cssHex(value: string | null | undefined): string | null {
  return value && /^[0-9A-Fa-f]{6}$/.test(value) ? '#' + value : null;
}

/**
 * Word picks a font per character: w:ascii for Latin text, w:eastAsia for CJK text, each
 * inherited on its own. CSS keeps the two as variables (--dx-font-l, --dx-font-e) and every
 * paragraph (and run) builds its font-family list from them: the Latin font first, so Latin
 * text never takes the CJK font's (wider) Latin letters, and CJK text falls through to its font.
 */
export const FONT_FAMILY = 'var(--dx-font-l), var(--dx-font-e), sans-serif';

/** CSS for w:rPr (character formatting). */
export function runPropsCss(rPr: Element | null, css: Css = {}): Css {
  if (!rPr) return css;
  const fonts = child(rPr, 'rFonts');
  if (fonts) {
    const latin = cssFontName(attr(fonts, 'ascii') ?? THEME_FONTS[attr(fonts, 'asciiTheme') ?? '']);
    const eastAsia = cssFontName(attr(fonts, 'eastAsia') ?? THEME_FONTS[attr(fonts, 'eastAsiaTheme') ?? '']);
    // Each with its other names and similar fonts, for computers without it (persona-300 B-8).
    if (latin) css['--dx-font-l'] = fontStack(latin);
    if (eastAsia) {
      css['--dx-font-e'] = fontStack(eastAsia, true);
      css['--dx-font-h'] = `"${eastAsiaAlias(eastAsia)}"`;
    }
  }
  const sz = numAttr(child(rPr, 'sz'), 'val');
  if (sz != null) css['font-size'] = `${halfPointsToPt(sz)}pt`;
  const b = onOff(child(rPr, 'b'));
  if (b != null) css['font-weight'] = b ? 'bold' : 'normal';
  const i = onOff(child(rPr, 'i'));
  if (i != null) css['font-style'] = i ? 'italic' : 'normal';
  const color = cssHex(attr(child(rPr, 'color'), 'val'));
  if (color) css.color = color;
  const caps = onOff(child(rPr, 'caps'));
  if (caps) css['text-transform'] = 'uppercase';
  const u = attr(child(rPr, 'u'), 'val');
  if (u && u !== 'none') css['text-decoration'] = 'underline';
  return css;
}

/** CSS for w:pPr (paragraph formatting that a style can set). */
export function paraPropsCss(pPr: Element | null, css: Css = {}): Css {
  if (!pPr) return css;
  const align = jcToAlign(attr(child(pPr, 'jc'), 'val'));
  if (align && ['left', 'center', 'right', 'justify'].includes(align)) css['text-align'] = align;
  const sp = child(pPr, 'spacing');
  if (sp) {
    const before = numAttr(sp, 'before');
    const after = numAttr(sp, 'after');
    if (before != null) css['margin-top'] = `${twipsToPx(before)}px`;
    if (after != null) css['margin-bottom'] = `${twipsToPx(after)}px`;
    const lh = lineHeightCss(numAttr(sp, 'line'), attr(sp, 'lineRule'));
    if (lh) css['line-height'] = lh;
    Object.assign(css, lineGridVars(numAttr(sp, 'line'), attr(sp, 'lineRule')));
  }
  const ind = child(pPr, 'ind');
  if (ind) {
    const left = numAttr(ind, 'left') ?? numAttr(ind, 'start');
    if (left != null) css['margin-left'] = `${twipsToPx(left)}px`;
    const right = numAttr(ind, 'right') ?? numAttr(ind, 'end');
    if (right != null) css['margin-right'] = `${twipsToPx(right)}px`;
    const first = numAttr(ind, 'firstLine');
    const hanging = numAttr(ind, 'hanging');
    if (first != null) css['text-indent'] = `${twipsToPx(first)}px`;
    if (hanging != null) css['text-indent'] = `${-twipsToPx(hanging)}px`;
  }
  return css;
}

function cssText(css: Css): string {
  return Object.entries(css)
    .map(([k, v]) => `${k}:${v}`)
    .join(';');
}

export interface ParsedStyles {
  css: string;
  paragraphStyles: ParagraphStyleInfo[];
  inheritance: StyleInheritance;
}

export function parseStyles(doc: Document | null): ParsedStyles {
  const base: Css = {
    '--dx-font-l': fontStack('Calibri'), '--dx-font-e': fontStack('PMingLiU', true), '--dx-font-h': `"${eastAsiaAlias('PMingLiU')}"`,
    'font-family': FONT_FAMILY, 'font-size': '11pt',
  };
  // Line height: Word's single spacing for the paragraph's fonts (--dx-lh), times the spacing.
  // The font list is rebuilt on each paragraph from the fonts its styles give (see FONT_FAMILY).
  const baseP: Css = { 'margin-top': '0px', 'margin-bottom': '0px', 'line-height': 'var(--dx-lh)', 'font-family': FONT_FAMILY };
  const raw = new Map<string, RawStyle>();
  const inheritance = emptyInheritance();

  if (doc) {
    const root = doc.documentElement;
    const defaults = child(root, 'docDefaults');
    runPropsCss(child(child(defaults, 'rPrDefault'), 'rPr'), base);
    paraPropsCss(child(child(defaults, 'pPrDefault'), 'pPr'), baseP);
    inheritance.defaults = ownProps(child(child(defaults, 'pPrDefault'), 'pPr'), child(child(defaults, 'rPrDefault'), 'rPr'));
    for (const s of children(root, 'style')) {
      const id = attr(s, 'styleId');
      if (!id) continue;
      const own: Css = {};
      paraPropsCss(child(s, 'pPr'), own);
      runPropsCss(child(s, 'rPr'), own);
      raw.set(id, {
        id,
        name: attr(child(s, 'name'), 'val') ?? id,
        type: attr(s, 'type') ?? 'paragraph',
        basedOn: attr(child(s, 'basedOn'), 'val'),
        isDefault: attr(s, 'default') === '1' || attr(s, 'default') === 'true',
        own,
        props: ownProps(child(s, 'pPr'), child(s, 'rPr')),
        cellMar: cellMarginCss(child(s, 'tblPr'), 'tblCellMar'),
      });
    }
  }

  const resolved = (id: string, seen = new Set<string>()): Css => {
    const s = raw.get(id);
    if (!s || seen.has(id)) return {};
    seen.add(id);
    return { ...(s.basedOn ? resolved(s.basedOn, seen) : {}), ...s.own };
  };

  const resolvedProps = (id: string, seen = new Set<string>()): StyleProps => {
    const s = raw.get(id);
    if (!s || seen.has(id)) return {};
    seen.add(id);
    return { ...(s.basedOn ? resolvedProps(s.basedOn, seen) : {}), ...s.props };
  };
  for (const s of raw.values()) {
    if (s.type === 'paragraph') {
      inheritance.paragraph[s.id] = resolvedProps(s.id);
      if (s.isDefault) inheritance.defaultParagraph = s.id;
    } else if (s.type === 'character') {
      inheritance.character[s.id] = resolvedProps(s.id);
    }
  }

  const rules: string[] = [];
  base['--dx-lh-l'] = String(wordLineRatio(inheritance.defaults.fontFamily) ?? LATIN_LINE);
  base['--dx-lh-e'] = String(wordLineRatio(inheritance.defaults.fontEastAsia) ?? EAST_ASIAN_LINE);
  rules.push(`.dx-doc{${cssText(base)}}`);
  // A line with East Asian text is as tall as the taller of the Latin and East Asian fonts;
  // --dx-lh-run is set on paragraphs whose runs name other fonts.
  rules.push('.dx-doc .dx-p{--dx-lh:max(var(--dx-lh-l),var(--dx-lh-run,0))}');
  rules.push('.dx-doc .dx-p.dx-cjk{--dx-lh:max(var(--dx-lh-l),var(--dx-lh-e),var(--dx-lh-run,0))}');
  rules.push(`.dx-doc .dx-p{${cssText(baseP)}}`);
  // Table styles: their paragraph and run properties apply to the text in the table, over the
  // document defaults but under paragraph styles (Word's order). :where() keeps the same
  // specificity as the paragraph-style rules that follow, so those win. Tables are tagged
  // .dx-ts-<id> (or .dx-ts--none, which gets the default table style) by editor/tableStyles.ts.
  const resolvedCellMar = (id: string, seen = new Set<string>()): Css => {
    const s = raw.get(id);
    if (!s || seen.has(id)) return {};
    seen.add(id);
    return { ...(s.basedOn ? resolvedCellMar(s.basedOn, seen) : {}), ...s.cellMar };
  };
  // Cell margins: Word's defaults on every table (so a nested table does not take its outer
  // table's), then the table style's; a table's own w:tblCellMar is set on it inline
  // (editor/tableStyles.ts) and a cell's w:tcMar on the cell (editor/schema.ts).
  rules.push(`.dx-doc :where([class*="dx-ts-"]){${cssText(DEFAULT_CELL_MARGINS)}}`);
  for (const s of raw.values()) {
    if (s.type !== 'table') continue;
    const mar = cssText(resolvedCellMar(s.id));
    if (!mar) continue;
    const selectors = [`.dx-doc :where(.dx-ts-${cssIdent(s.id)})`];
    if (s.isDefault) selectors.push('.dx-doc :where(.dx-ts--none)');
    rules.push(`${selectors.join(',')}{${mar}}`);
  }
  for (const s of raw.values()) {
    if (s.type !== 'table') continue;
    const css = cssText(lineVars(resolvedProps(s.id), resolved(s.id)));
    if (!css) continue;
    const selectors = [`.dx-doc :where(.dx-ts-${cssIdent(s.id)}) .dx-p`];
    if (s.isDefault) selectors.push('.dx-doc :where(.dx-ts--none) .dx-p');
    rules.push(`${selectors.join(',')}{${css}}`);
  }
  const paragraphStyles: ParagraphStyleInfo[] = [];
  // The default paragraph style must come before the named ones (same specificity, later wins).
  for (const s of raw.values()) {
    if (s.type === 'paragraph' && s.isDefault) {
      rules.push(`.dx-doc .dx-p{${cssText(lineVars(inheritance.paragraph[s.id] ?? {}, resolved(s.id)))}}`);
    }
  }
  for (const s of raw.values()) {
    if (s.type !== 'paragraph') continue;
    const css = lineVars(inheritance.paragraph[s.id] ?? {}, resolved(s.id));
    rules.push(`.dx-doc .dx-ps-${cssIdent(s.id)}{${cssText(css)}}`);
    paragraphStyles.push({ id: s.id, name: s.name });
  }
  for (const s of raw.values()) {
    if (s.type !== 'character') continue;
    const css = resolved(s.id);
    if (css['--dx-font-l'] || css['--dx-font-e']) css['font-family'] = FONT_FAMILY;
    rules.push(`.dx-doc .dx-cs-${cssIdent(s.id)}{${cssText(css)}}`);
  }
  return { css: rules.join('\n'), paragraphStyles, inheritance };
}

/**
 * CSS for the run properties that the editor does not model as marks, taken from a
 * run's original w:rPr (caps, character spacing, hidden text, raised text ...).
 */
export function runExtraCss(rPrXml: string): string {
  let rPr: Element;
  try {
    rPr = parseFragment(rPrXml);
  } catch {
    return '';
  }
  const css: string[] = [];
  if (onOff(child(rPr, 'caps'))) css.push('text-transform:uppercase');
  if (onOff(child(rPr, 'smallCaps'))) css.push('font-variant:small-caps');
  if (onOff(child(rPr, 'vanish'))) css.push('display:none');
  const spacing = numAttr(child(rPr, 'spacing'), 'val');
  if (spacing) css.push(`letter-spacing:${twipsToPx(spacing)}px`);
  const position = numAttr(child(rPr, 'position'), 'val');
  if (position) css.push(`position:relative;top:${-halfPointsToPt(position)}pt`);
  const fill = cssHex(attr(child(rPr, 'shd'), 'fill'));
  if (fill && !child(rPr, 'highlight')) css.push(`background-color:${fill}`);
  if (onOff(child(rPr, 'dstrike'))) css.push('text-decoration-style:double');
  const u = attr(child(rPr, 'u'), 'val');
  if (u === 'double') css.push('text-decoration-style:double');
  else if (u === 'dotted' || u === 'dottedHeavy') css.push('text-decoration-style:dotted');
  else if (u === 'dash' || u === 'dashedHeavy' || u === 'dashLong') css.push('text-decoration-style:dashed');
  else if (u === 'wave' || u === 'wavyHeavy' || u === 'wavyDouble') css.push('text-decoration-style:wavy');
  if (child(rPr, 'bdr')) css.push('border:1px solid currentColor');
  return css.join(';');
}
