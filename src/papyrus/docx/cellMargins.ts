import { attr, child, numAttr, parseFragment } from './xml';
import { twipsToPx } from '../units';

type Css = Record<string, string>;

const MARGIN_SIDES: [string, string, string][] = [
  ['top', 'top', 't'], ['left', 'left', 'l'], ['start', 'left', 'l'],
  ['bottom', 'bottom', 'b'], ['right', 'right', 'r'], ['end', 'right', 'r'],
];

/**
 * Cell margins (w:tblCellMar in a w:tblPr, or w:tcMar in a w:tcPr): the space between a cell's
 * edge and its text. As CSS variables (--dx-cm-t/-l/-b/-r) that table cells read (editor.css),
 * or with `padding` as the cell's own padding. Only twips (dxa) values are used.
 */
export function cellMarginCss(pr: Element | null, name: 'tblCellMar' | 'tcMar', as: 'vars' | 'padding' = 'vars'): Css {
  const css: Css = {};
  const mar = child(pr, name);
  if (!mar) return css;
  for (const [tag, side, short] of MARGIN_SIDES) {
    const el = child(mar, tag);
    const w = numAttr(el, 'w');
    const type = attr(el, 'type') ?? 'dxa';
    if (w == null || (type !== 'dxa' && type !== 'nil')) continue;
    const px = `${type === 'nil' ? 0 : twipsToPx(Math.max(0, w))}px`;
    css[as === 'vars' ? `--dx-cm-${short}` : `padding-${side}`] = px;
  }
  return css;
}

/** Word's cell margins when nothing sets them: 108 twips left and right, none above and below. */
export const DEFAULT_CELL_MARGINS: Css = { '--dx-cm-t': '0px', '--dx-cm-r': '7.2px', '--dx-cm-b': '0px', '--dx-cm-l': '7.2px' };

const parsed = new Map<string, Css>();

/** A cell's own margins (w:tcMar in its raw w:tcPr) as padding CSS; cached per tcPr. */
export function tcMarginCss(tcPr: string | null): string {
  if (!tcPr || !tcPr.includes('tcMar')) return '';
  let css = parsed.get(tcPr);
  if (!css) {
    try {
      css = cellMarginCss(parseFragment(tcPr), 'tcMar', 'padding');
    } catch {
      css = {};
    }
    if (parsed.size > 500) parsed.clear();
    parsed.set(tcPr, css);
  }
  return Object.entries(css).map(([k, v]) => `${k}:${v};`).join('');
}

/** A table's own margins (w:tblCellMar in its raw w:tblPr) as the CSS variables cells read. */
export function tblMarginVars(tblPr: string | null): string {
  if (!tblPr || !tblPr.includes('tblCellMar')) return '';
  try {
    const css = cellMarginCss(parseFragment(tblPr), 'tblCellMar');
    return Object.entries(css).map(([k, v]) => `${k}:${v}`).join(';');
  } catch {
    return '';
  }
}
