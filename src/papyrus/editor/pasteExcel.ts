// A range copied from Excel (persona-300 B-9). Excel's clipboard HTML keeps the cells' look in a
// <style> block (.xl65 { background:#FFFF00; text-align:center; font-weight:700; border:.5pt
// solid windowtext }) and the column widths in <col width>, neither of which the paste reads:
// the cells came in white, left-aligned, without borders and with equal columns. Before the
// HTML is parsed, those rules are written on the cells themselves, a cell's alignment on a
// paragraph around its text, its borders in data-dx-borders (read by the table cell's
// `borders` attribute, see schema.ts) and the columns' widths in data-colwidth (read by
// prosemirror-tables). Other HTML is left as it is.

/** Whether pasted HTML comes from Excel. */
export function isExcelHtml(html: string): boolean {
  return /content=["']?Excel\.Sheet|urn:schemas-microsoft-com:office:excel/i.test(html);
}

/** The properties of a cell's class rules that are taken over (not fonts and sizes: the document's stay). */
const TAKEN = /^(background|background-color|text-align|vertical-align|font-weight|font-style|color|text-decoration|border(-(top|right|bottom|left))?)$/;

type Declarations = [string, string][];

function declarations(text: string): Declarations {
  const out: Declarations = [];
  for (const part of text.split(';')) {
    const i = part.indexOf(':');
    if (i < 0) continue;
    const name = part.slice(0, i).trim().toLowerCase();
    const value = part.slice(i + 1).trim();
    if (name && value) out.push([name, value]);
  }
  return out;
}

/** Style text read from pasted HTML, at most (Excel's is a few KB; more is not a range's look). */
export const MAX_STYLE_TEXT = 200 * 1024;

/** CSS without its comments (an unclosed one runs to the end), in one pass. */
function withoutComments(css: string): string {
  let out = '';
  let at = 0;
  for (;;) {
    const open = css.indexOf('/*', at);
    if (open < 0) return out + css.slice(at);
    out += css.slice(at, open);
    const close = css.indexOf('*/', open + 2);
    if (close < 0) return out;
    at = close + 2;
  }
}

/**
 * The rules of a style sheet as [selector, declarations], in one pass over it: the rules
 * inside an @media … { } block too (its own selector is dropped), nothing for text with no
 * braces. (A regular expression for this took seconds on a long brace-free style text.)
 */
export function cssRules(css: string): [string, string][] {
  const out: [string, string][] = [];
  let start = 0; // where the next selector starts
  let open = -1; // the "{" of the rule being read
  for (let i = 0; i < css.length; i++) {
    const c = css.charCodeAt(i);
    if (c === 0x7b) {
      // "{" again before a "}": the one before opened an at-rule's block.
      if (open >= 0) start = open + 1;
      open = i;
    } else if (c === 0x7d) {
      if (open >= 0) out.push([css.slice(start, open), css.slice(open + 1, i)]);
      open = -1;
      start = i + 1;
    }
  }
  return out;
}

/** Class name → the declarations of its rules (".xl65", "td.xl65"), in the order written. */
function classRules(root: DocumentFragment): Map<string, Declarations> {
  const rules = new Map<string, Declarations>();
  let budget = MAX_STYLE_TEXT;
  for (const style of Array.from(root.querySelectorAll('style'))) {
    if (budget <= 0) break;
    const text = (style.textContent ?? '').slice(0, budget);
    budget -= text.length;
    const css = withoutComments(text.replace(/<!--|-->/g, ''));
    for (const [selector, body] of cssRules(css)) {
      const decls = declarations(body).filter(([n]) => TAKEN.test(n));
      if (!decls.length) continue;
      for (const sel of selector.split(',')) {
        const cls = /^\s*(?:td|th)?\.([\w-]+)\s*$/i.exec(sel)?.[1];
        if (cls) rules.set(cls, [...(rules.get(cls) ?? []), ...decls]);
      }
    }
  }
  return rules;
}

/** A border side from CSS ("0.5pt solid windowtext", "none") as the cell's borders attribute has it. */
function borderSide(value: string): { val: string; sz: number | null; color: string | null } | null {
  const v = value.toLowerCase().trim();
  if (!v || v === 'none' || v === '0' || /\bhidden\b/.test(v)) return { val: 'nil', sz: null, color: null };
  const style = /\b(double|dotted|dashed|solid)\b/.exec(v)?.[1] ?? 'solid';
  const width = /(\d*\.?\d+)(pt|px)/.exec(v);
  const pt = width ? Number(width[1]) * (width[2] === 'px' ? 0.75 : 1) : 0.5;
  const hex = /#([0-9a-f]{6}|[0-9a-f]{3})\b/.exec(v)?.[1];
  // windowtext, black or nothing: Word's automatic colour.
  const color = hex ? (hex.length === 3 ? hex.replace(/./g, '$&$&') : hex).toUpperCase() : 'auto';
  return { val: style === 'solid' ? 'single' : style, sz: Math.min(96, Math.max(2, Math.round(pt * 8))), color };
}

/** The cell's four sides from its border declarations (later ones win); null when it sets none. */
function cellBorders(decls: Declarations): Record<string, ReturnType<typeof borderSide>> | null {
  const sides: Record<string, ReturnType<typeof borderSide>> = {};
  let any = false;
  for (const [name, value] of decls) {
    if (name === 'border') {
      for (const s of ['top', 'right', 'bottom', 'left']) sides[s] = borderSide(value);
      any = true;
    } else if (name.startsWith('border-')) {
      sides[name.slice(7)] = borderSide(value);
      any = true;
    }
  }
  return any ? sides : null;
}

/** A column width in px from <col width=… style='width:…pt'>. */
function colWidth(col: Element): number | null {
  const w = Number(col.getAttribute('width'));
  if (w > 0) return Math.round(w);
  const pt = /width:\s*(\d*\.?\d+)pt/i.exec(col.getAttribute('style') ?? '');
  return pt ? Math.round((Number(pt[1]) * 96) / 72) : null;
}

/** Excel's clipboard HTML with its cell styles and column widths written on the cells (see above). */
export function inlineExcelStyles(html: string): string {
  if (!isExcelHtml(html) || !/<table\b/i.test(html)) return html;
  const t = document.createElement('template');
  t.innerHTML = html; // inert: nothing loads, nothing runs
  const rules = classRules(t.content);
  for (const table of Array.from(t.content.querySelectorAll('table'))) {
    const widths: (number | null)[] = [];
    for (const col of Array.from(table.querySelectorAll('col'))) {
      const span = Math.min(64, Math.max(1, Number(col.getAttribute('span')) || 1));
      for (let i = 0; i < span; i++) widths.push(colWidth(col));
    }
    // Which columns are taken by cells spanning down from rows above.
    const taken: number[][] = [];
    Array.from(table.querySelectorAll('tr')).forEach((tr, r) => {
      let c = 0;
      for (const cell of Array.from(tr.children).filter((e) => /^(td|th)$/i.test(e.tagName)) as HTMLElement[]) {
        while (taken[r]?.includes(c)) c++;
        const colspan = Math.max(1, Number(cell.getAttribute('colspan')) || 1);
        const rowspan = Math.max(1, Number(cell.getAttribute('rowspan')) || 1);
        for (let k = 1; k < rowspan; k++) (taken[r + k] ??= []).push(...Array.from({ length: colspan }, (_, i) => c + i));
        const w = widths.slice(c, c + colspan);
        if (w.length === colspan && w.every((x) => x)) cell.setAttribute('data-colwidth', w.join(','));
        c += colspan;
        // The class rules first, then the cell's own style (it wins, as in CSS).
        const decls = (cell.getAttribute('class') ?? '').split(/\s+/).flatMap((cls) => rules.get(cls) ?? []);
        const own = declarations(cell.getAttribute('style') ?? '');
        const all = [...decls, ...own.filter(([n]) => TAKEN.test(n))];
        if (!decls.length) continue;
        const style = decls.filter(([n]) => !n.startsWith('border')).map(([n, v]) => `${n}:${v}`).join(';');
        if (style) cell.setAttribute('style', `${style};${cell.getAttribute('style') ?? ''}`);
        const borders = cellBorders(all);
        if (borders) cell.setAttribute('data-dx-borders', JSON.stringify(borders));
        // Excel's cells hold text, not paragraphs: the alignment goes on a paragraph around it.
        const align = [...all].reverse().find(([n]) => n === 'text-align')?.[1];
        if (align && !cell.querySelector('p, div, table')) {
          const p = document.createElement('p');
          p.style.textAlign = align;
          p.append(...Array.from(cell.childNodes));
          cell.append(p);
        }
      }
    });
  }
  return t.innerHTML;
}

/** A cell's pasted borders (data-dx-borders, see inlineExcelStyles), checked: null when there are none. */
export function pastedBorders(value: string | null): Record<string, { val: string; sz: number | null; color: string | null }> | null {
  if (!value) return null;
  try {
    const parsed = JSON.parse(value) as Record<string, unknown>;
    const out: Record<string, { val: string; sz: number | null; color: string | null }> = {};
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const s = parsed?.[side] as { val?: unknown; sz?: unknown; color?: unknown } | undefined;
      if (!s || typeof s.val !== 'string' || !/^(nil|single|double|dotted|dashed)$/.test(s.val)) continue;
      out[side] = {
        val: s.val,
        sz: typeof s.sz === 'number' && Number.isFinite(s.sz) ? s.sz : null,
        color: typeof s.color === 'string' && /^([0-9A-F]{6}|auto)$/.test(s.color) ? s.color : null,
      };
    }
    return Object.keys(out).length ? out : null;
  } catch {
    return null;
  }
}
