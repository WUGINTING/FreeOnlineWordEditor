import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { TableMap } from 'prosemirror-tables';
import { isTyping } from './steps';
import { columnBreakIn, type ColumnGeometry, type ColumnRule } from './columnLayout';
import { formattingMarksKey, type FormattingMarksMeta } from './formattingMarks';

/**
 * Paginated view.
 *
 * The document is a single continuous editable column. Pages are drawn behind it, and
 * this plugin inserts invisible spacers so content that would cross the bottom of a page's
 * text area moves to the top of the next page:
 * - between top-level blocks, and between the lines of a paragraph (widow/orphan control);
 * - between table rows (a table continues on the next page; header rows marked to repeat
 *   are shown again at the top of each continuation);
 * - inside a row that may break across pages (Word's default): each cell's lines that do not
 *   fit continue at the top of the next page.
 * Each section has its own page size and margins; a new section starts a new page unless
 * it is continuous on the same paper. Positions are measured in the browser after every render.
 */

/** A section's page, in px. */
export interface PageGeometry {
  width: number;
  height: number;
  /** Top of the text area from the top edge (a tall header pushes it down). */
  textTop: number;
  /** Bottom of the text area from the bottom edge. */
  textBottom: number;
  marginLeft: number;
  marginRight: number;
}

export interface LayoutSection {
  geometry: PageGeometry;
  /** Whether the section starts on a new page, and which page parity it wants. */
  start: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
  pageNumberStart: number | null;
  /** w:pgNumType/@w:fmt: how PAGE fields show the section's page numbers (null: 1, 2, 3). */
  pageNumberFormat?: string | null;
  firstBlock: number;
  lastBlock: number;
  /** The section's text columns (w:cols), when it has more than one (editor/columnLayout.ts). */
  columns?: ColumnGeometry | null;
  /**
   * The columns of a section laid out in one column instead, because one of its blocks is taller
   * than a whole column (see computeBreaks' `singleColumn`); `columns` is null then.
   */
  columnsOff?: ColumnGeometry | null;
}

export interface Layout {
  sections: LayoutSection[];
  /** Gap between pages on screen, px. */
  gap: number;
  /**
   * The space before a paragraph that starts a page after a page break: kept (Word 2010 and
   * older layout), dropped (Word 2013+ layout, compatibility mode 15), or dropped only after a
   * hard page break (w:suppressSpBfAfterPgBrk). Section starts always keep it.
   */
  topSpacing?: 'keep' | 'suppress' | 'afterPageBreak';
}

/** One page as laid out on screen. */
export interface PageBox extends PageGeometry {
  /** Index of the section the page belongs to. */
  section: number;
  /**
   * Other sections with text on this page: continuous sections that start on it after `section`
   * (absent when there are none). Used to mark pages of sections Word shows in columns.
   */
  alsoSections?: number[];
  /** Index of the page inside its section (0 = the section's first page). */
  inSection: number;
  /** Printed page number (restarts where a section says so). */
  number: number;
  /** How PAGE fields show `number` on this page (its section's w:pgNumType/@w:fmt; null: 1, 2, 3). */
  numberFormat?: string | null;
  /** Top of the page on the canvas, px. */
  top: number;
  /** Left of the page on the canvas (narrower pages are centered), px. */
  left: number;
  /**
   * Table rows that break at the end of this page: from where the row's part on this page
   * ends to where it continues on the next page (canvas px). The cells' borders run through
   * that band; the editor covers it and draws the row's bottom and top edges there, like Word.
   */
  cuts?: {
    top: number;
    bottom: number;
    left: number;
    right: number;
    border: string;
    /** Header rows repeated at the top of the next page (positions of the rows, their table). */
    header?: { table: number; rows: number[] };
  }[];
  /** Lines between text columns (w:sep) on this page, page px (see columnLayout.ts). */
  rules?: ColumnRule[];
}

/**
 * A block of a section in columns that is not in its first column (or not where the text flow
 * puts it): what moves it into its column (padding deltas, px) and up to the column's top.
 */
export interface ColumnPlacement {
  pos: number;
  size: number;
  dl: number;
  dr: number;
  top: number;
}

interface Break {
  pos: number;
  height: number;
  /** A break between table rows: the table's column count and the header rows to repeat. */
  row?: { cols: number; header: number[] };
  /** A break between two lines of a paragraph (the paragraph continues on the next page). */
  line?: boolean;
  /**
   * A break inside a table cell: the row continues on the next page, and this cell's content
   * from here on moves there (before a paragraph, or between two lines of one).
   */
  cell?: 'block' | 'line';
  /**
   * A spacer that pulls what follows up (a negative margin, px), after text moved up into a
   * column: the text continues below the columns, or at the top of the next page.
   */
  collapse?: number;
}

export interface PaginationState {
  decorations: DecorationSet;
  /** Bumped to request a fresh measurement. */
  nonce: number;
  breaks: Break[];
  pages: PageBox[];
  /** Blocks moved into their text column (see ColumnPlacement). */
  placements?: ColumnPlacement[];
  /** The last change to the document was typing (measured after a pause, see TYPING_DELAY). */
  typing?: boolean;
}

/** Wait before measuring after an edit, a page setup change, a loaded picture ... (ms). */
const MEASURE_DELAY = 30;
/**
 * Wait after a keystroke: typing goes on without a measurement until a short pause (or at most
 * TYPING_MAX_WAIT into a long burst). Only when a measurement runs changes, not its result.
 */
const TYPING_DELAY = 200;
const TYPING_MAX_WAIT = 1000;
/**
 * Measurements that change the result, per document version: a layout that keeps alternating
 * between two results (e.g. a header whose height depends on the page count) stops there.
 */
const MAX_CHANGES_PER_DOC = 50;
/** Never more pages than this, whatever the page size says (a measurement stops there). */
export const MAX_PAGES = 10000;

/**
 * What decides whether the pages must be drawn again. Positions are compared to half a pixel:
 * measuring the same layout twice gives cut positions a thousandth of a pixel apart.
 */
export function pagesKey(pages: PageBox[]): string {
  return JSON.stringify(pages, (_key, value) => (typeof value === 'number' ? Math.round(value * 2) / 2 : value));
}

export const paginationKey = new PluginKey<PaginationState>('dx-pagination');

function spacer(height: number, collapse = 0): HTMLElement {
  const el = document.createElement('div');
  el.className = collapse ? 'dx-spacer dx-collapse' : 'dx-spacer';
  el.style.height = `${height}px`;
  if (collapse) el.style.marginTop = `${collapse}px`;
  el.contentEditable = 'false';
  return el;
}

/** Space between two table rows that fall on different pages. */
function rowSpacer(height: number, cols: number): HTMLElement {
  const tr = document.createElement('tr');
  tr.className = 'dx-row-spacer';
  tr.contentEditable = 'false';
  const td = document.createElement('td');
  td.colSpan = Math.max(1, cols);
  td.style.height = `${height}px`;
  tr.append(td);
  return tr;
}

/** A copy of a header row, shown again at the top of a table's continuation. */
function repeatedHeader(view: EditorView, pos: number): HTMLElement {
  const source = view.nodeDOM(pos) as HTMLElement | null;
  const tr = (source?.cloneNode(true) as HTMLElement | undefined) ?? document.createElement('tr');
  tr.classList.add('dx-row-spacer', 'dx-repeat-header');
  tr.contentEditable = 'false';
  tr.removeAttribute('data-pm-slice');
  return tr;
}

/** A short hash, so a widget is redrawn when what it copies changes. */
function hash(text: string): string {
  let h = 5381;
  for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

/**
 * Space inside a paragraph that moves its remaining lines to the next page. A full-width
 * inline block: the line before it still wraps (and is justified) as usual.
 */
function lineSpacer(height: number): HTMLElement {
  const el = document.createElement('span');
  el.className = 'dx-line-spacer';
  el.style.height = `${height}px`;
  el.contentEditable = 'false';
  return el;
}

/** Spacers inside table cells are measured with their row (see measureRow), not on their own. */
function inCell(el: HTMLElement): HTMLElement {
  el.classList.add('dx-in-cell');
  return el;
}

function decorationsFor(state: EditorState, breaks: Break[], placements: ColumnPlacement[] = [], marks = !!formattingMarksKey.getState(state)?.on): DecorationSet {
  const decos: Decoration[] = [];
  for (const p of placements) {
    const node = state.doc.nodeAt(p.pos);
    if (!node || node.nodeSize !== p.size) continue;
    const style = `--dx-cdl:${p.dl}px;--dx-cdr:${p.dr}px` + (p.top ? `;top:${p.top}px` : '');
    decos.push(Decoration.node(p.pos, p.pos + p.size, p.top ? { class: 'dx-col-shift', style } : { style }));
  }
  // A row breaking across pages shows its cells' content from the top on each page.
  const splitRows = new Set<number>();
  for (const b of breaks) {
    if (!b.cell) continue;
    const $pos = state.doc.resolve(b.pos);
    for (let d = $pos.depth; d > 0; d--) {
      if ($pos.node(d).type.name !== 'table_row') continue;
      const at = $pos.before(d);
      if (!splitRows.has(at)) {
        splitRows.add(at);
        decos.push(Decoration.node(at, at + $pos.node(d).nodeSize, { class: 'dx-row-split' }));
      }
      break;
    }
  }
  for (const b of breaks) {
    if (b.cell === 'line') {
      decos.push(Decoration.widget(b.pos, () => inCell(lineSpacer(b.height)), { side: -1, key: `cl${b.height}`, ignoreSelection: true, marks: [] }));
      continue;
    }
    if (b.cell === 'block') {
      decos.push(Decoration.widget(b.pos, () => inCell(spacer(b.height)), { side: -1, key: `cb${b.height}`, ignoreSelection: true }));
      continue;
    }
    if (b.line) {
      decos.push(Decoration.widget(b.pos, () => lineSpacer(b.height), { side: -1, key: `ls${b.height}`, ignoreSelection: true, marks: [] }));
      continue;
    }
    if (!b.row) {
      const c = b.collapse ?? 0;
      decos.push(Decoration.widget(b.pos, () => spacer(b.height, c), { side: -1, key: `sp${b.height}|${c}`, ignoreSelection: true }));
      continue;
    }
    const { cols, header } = b.row;
    decos.push(
      Decoration.widget(b.pos, () => rowSpacer(b.height, cols), { side: -2, key: `rs${b.height}|${cols}`, ignoreSelection: true }),
    );
    header.forEach((hp, i) =>
      decos.push(
        Decoration.widget(b.pos, (view) => repeatedHeader(view, hp), {
          side: -1,
          // Copied from the row's DOM: drawn again when the row, or the formatting marks shown in
          // it (顯示／隱藏編輯標記), change.
          key: `rh${hp}|${i}|${marks ? 'm' : ''}${hash(JSON.stringify(state.doc.nodeAt(hp)?.toJSON() ?? null))}`,
          ignoreSelection: true,
        }),
      ),
    );
  }
  return DecorationSet.create(state.doc, decos);
}

/** Each section's left/right padding in the body column (narrower pages are centered). */
function columnPaddings(layout: Layout): { left: number; right: number }[] {
  const maxWidth = Math.max(...layout.sections.map((s) => s.geometry.width));
  return layout.sections.map(({ geometry: g }) => {
    const side = (maxWidth - g.width) / 2;
    return { left: side + g.marginLeft, right: side + g.marginRight };
  });
}

/**
 * The body column's own left/right padding: the smallest of the sections', so every other
 * section only adds to it (see sectionColumns).
 */
export function columnBase(layout: Layout): { left: number; right: number } {
  const pads = columnPaddings(layout);
  return { left: Math.min(...pads.map((p) => p.left)), right: Math.min(...pads.map((p) => p.right)) };
}

function columnDecorations(doc: PMNode, layout: Layout): DecorationSet {
  const pads = columnPaddings(layout);
  const base = columnBase(layout);
  const columns = layout.sections.some((s) => s.columns);
  if (!columns && pads.every((p) => p.left === base.left && p.right === base.right)) return DecorationSet.empty;
  const decos: Decoration[] = [];
  let s = 0;
  doc.forEach((node, offset, index) => {
    while (s < layout.sections.length - 1 && index > layout.sections[s].lastBlock) s++;
    const extraLeft = pads[s].left - base.left;
    const extraRight = pads[s].right - base.right;
    const cols = layout.sections[s].columns;
    if (cols) {
      // A section in columns: its blocks are as wide as its first column (pagination moves them
      // into the others, see ColumnPlacement); columnLayout.css adds the two.
      const g = layout.sections[s].geometry;
      const textWidth = g.width - g.marginLeft - g.marginRight;
      const left = extraLeft + cols.lefts[0];
      const right = extraRight + textWidth - cols.lefts[0] - cols.widths[0];
      decos.push(Decoration.node(offset, offset + node.nodeSize, { class: 'dx-colsec', style: `--dx-cpl:${left}px;--dx-cpr:${right}px` }));
      return;
    }
    if (extraLeft === 0 && extraRight === 0) return;
    decos.push(
      Decoration.node(offset, offset + node.nodeSize, { style: `padding-left:${extraLeft}px;padding-right:${extraRight}px` }),
    );
  });
  return DecorationSet.create(doc, decos);
}

/** What decides where blocks sit horizontally. */
const columnKey = (layout: Layout) =>
  layout.sections
    .map((s) => `${s.lastBlock}:${s.geometry.width}:${s.geometry.marginLeft}:${s.geometry.marginRight}:${s.columns ? JSON.stringify(s.columns) : ''}`)
    .join('|');

/**
 * Sections with other page widths or side margins than the body column: their blocks get
 * the extra padding. Kept apart from the page spacers and mapped through edits, so typing
 * in a long document does not rebuild it.
 */
export function sectionColumns(getLayout: (doc: PMNode) => Layout): Plugin<{ set: DecorationSet; key: string }> {
  const key = new PluginKey<{ set: DecorationSet; key: string }>('dx-section-columns');
  const build = (doc: PMNode) => {
    const layout = getLayout(doc);
    return { set: columnDecorations(doc, layout), key: columnKey(layout) };
  };
  return new Plugin({
    key,
    state: {
      init: (_, state) => build(state.doc),
      apply(tr, value, old, state) {
        if (!tr.docChanged && !tr.getMeta(REMEASURE)) return value;
        const next = columnKey(getLayout(state.doc));
        if (next === value.key && state.doc.childCount === old.doc.childCount) {
          return { set: value.set.map(tr.mapping, tr.doc), key: next };
        }
        return build(state.doc);
      },
    },
    props: { decorations: (state) => key.getState(state)!.set },
  });
}

export function pagination(
  getLayout: (doc: PMNode) => Layout,
  onPages: (pages: PageBox[]) => void,
  /** The sections in columns to lay out in one column now (see FlowResult.singleColumn), when that changed. */
  onSingleColumn?: (sections: number[]) => void,
): Plugin<PaginationState> {
  return new Plugin<PaginationState>({
    key: paginationKey,
    state: {
      init: () => ({ decorations: DecorationSet.empty, breaks: [], pages: [], nonce: 0 }),
      apply(tr, value, _old, newState) {
        const meta = tr.getMeta(paginationKey) as { breaks: Break[]; pages: PageBox[]; placements?: ColumnPlacement[] } | undefined;
        if (meta) return { ...value, placements: undefined, ...meta, decorations: decorationsFor(newState, meta.breaks, meta.placements) };
        if (tr.getMeta(REMEASURE)) return { ...value, nonce: value.nonce + 1, typing: false };
        // Formatting marks shown or hidden: repeated header rows are copies of rows that now show
        // them or not (nothing moves, so no new measurement).
        const marks = (tr.getMeta(formattingMarksKey) as FormattingMarksMeta | undefined)?.on;
        if (marks !== undefined && value.breaks.some((b) => b.row?.header.length)) {
          return { ...value, decorations: decorationsFor(newState, value.breaks, value.placements ?? [], marks) };
        }
        if (!tr.docChanged) return value;
        // Keep spacers roughly in place until the next measurement.
        const breaks = value.breaks
          .map((b) => ({ ...b, pos: tr.mapping.map(b.pos, -1), row: b.row && { ...b.row, header: b.row.header.map((p) => tr.mapping.map(p)) } }))
          .filter((b, i, all) => i === 0 || b.pos !== all[i - 1].pos);
        const placements = value.placements?.map((p) => ({ ...p, pos: tr.mapping.map(p.pos, 1) }));
        return { ...value, breaks, placements, decorations: value.decorations.map(tr.mapping, tr.doc), typing: isTyping(tr) };
      },
    },
    props: {
      decorations: (state) => paginationKey.getState(state)!.decorations,
    },
    view(view) {
      // Debounced: a burst of keystrokes causes a single measurement. A timer (not
      // requestAnimationFrame) keeps working while the page is in a background tab.
      let timer: ReturnType<typeof setTimeout> | undefined;
      /** The pending measurement waits for a pause in typing (see TYPING_DELAY). */
      let pendingTyping = false;
      let typingSince = 0;
      let lastPages = '';
      // New spacers or pages can themselves move the content (page chrome, section widths,
      // late layout), so a changed result is measured again until it stays the same.
      let settle = 0;
      let measuredDoc: PMNode | null = null;
      let changesForDoc = 0;
      const run = () => {
        timer = undefined;
        pendingTyping = false;
        typingSince = 0;
        measure(view);
      };
      /** Measure soon (30 ms after the last request), or after a pause in typing. */
      const schedule = (typing = false) => {
        // A measurement already on its way soon covers the keystroke too.
        if (typing && timer !== undefined && !pendingTyping) return;
        clearTimeout(timer);
        if (!typing) {
          pendingTyping = false;
          typingSince = 0;
          timer = setTimeout(run, MEASURE_DELAY);
          return;
        }
        const now = Date.now();
        typingSince ||= now;
        pendingTyping = true;
        timer = setTimeout(run, Math.max(0, Math.min(TYPING_DELAY, typingSince + TYPING_MAX_WAIT - now)));
      };
      const soon = () => schedule(false);
      const measure = (v: EditorView) => {
        if (v.isDestroyed) return;
        const doc = v.state.doc;
        if (doc !== measuredDoc) {
          measuredDoc = doc;
          changesForDoc = 0;
        }
        // The result keeps alternating: wait for the next edit.
        if (changesForDoc >= MAX_CHANGES_PER_DOC) return;
        let changed = false;
        try {
          const layout = getLayout(doc);
          const result = computeBreaks(v, layout);
          const current = paginationKey.getState(v.state)!;
          // Sections in columns to lay out in one column (or in columns again): the host changes the
          // layout and has it measured again (this result is for the layout as it was).
          const single = (result.singleColumn ?? []).join();
          if (onSingleColumn && single !== layout.sections.flatMap((s, i) => (s.columnsOff ? [i] : [])).join()) {
            changesForDoc++;
            onSingleColumn(result.singleColumn ?? []);
            return;
          }
          if (
            !sameBreaks(result.breaks, current.breaks) || result.pages.length !== current.pages.length ||
            JSON.stringify(result.placements ?? []) !== JSON.stringify(current.placements ?? [])
          ) {
            v.dispatch(v.state.tr.setMeta(paginationKey, result).setMeta('addToHistory', false));
            changed = true;
          }
          // Drawn with the exact values, compared to half a pixel (see pagesKey).
          const key = pagesKey(result.pages);
          if (key !== lastPages) {
            lastPages = key;
            onPages(result.pages);
            changed = true;
          }
        } catch (err) {
          // A failed measurement leaves the pages as they are; the next change measures again.
          console.error('papyrus: page layout failed', err);
          return;
        }
        if (changed) changesForDoc++;
        if (changed && settle < 5) {
          settle++;
          soon();
        } else if (!changed) settle = 0;
      };
      // Images change height when they finish loading.
      const onLoad = (e: Event) => {
        if ((e.target as HTMLElement).tagName === 'IMG') soon();
      };
      view.dom.addEventListener('load', onLoad, true);
      window.addEventListener('resize', soon);
      // Catch-all for anything else that changes the height (late fonts, images, CSS). Text
      // growing while typing keeps waiting for the pause.
      const resize = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => schedule(pendingTyping)) : null;
      resize?.observe(view.dom);
      document.fonts?.ready.then(soon);
      soon();
      return {
        update: (_v, prev) => {
          const state = paginationKey.getState(view.state)!;
          const changedNonce = paginationKey.getState(prev)?.nonce !== state.nonce;
          if (prev.doc !== view.state.doc || changedNonce) {
            settle = 0;
            // Typing is measured after a pause; anything else (Enter, a page break, a table,
            // a paste, undo, page setup ...) right away.
            schedule(!changedNonce && state.typing === true);
          }
        },
        destroy: () => {
          clearTimeout(timer);
          resize?.disconnect();
          view.dom.removeEventListener('load', onLoad, true);
          window.removeEventListener('resize', soon);
        },
      };
    },
  });
}

/** Force a re-measure, e.g. after page size or styles changed. */
const REMEASURE = 'dx-remeasure';
export function repaginate(view: EditorView): void {
  view.dispatch(view.state.tr.setMeta(REMEASURE, true).setMeta('addToHistory', false));
}

function sameBreaks(a: Break[], b: Break[]): boolean {
  if (a.length !== b.length) return false;
  return a.every(
    (x, i) =>
      x.pos === b[i].pos && Math.abs(x.height - b[i].height) <= 1 && !!x.line === !!b[i].line && x.cell === b[i].cell &&
      Math.abs((x.collapse ?? 0) - (b[i].collapse ?? 0)) <= 1 &&
      !!x.row === !!b[i].row && (x.row?.header.join() ?? '') === (b[i].row?.header.join() ?? ''),
  );
}

/** Header rows repeat on every page of a table (w:tblHeader, the leading rows only). */
function isHeaderRow(row: PMNode): boolean {
  if (row.attrs.tblHeader === true || row.attrs.header === true || row.attrs.repeatHeader === true) return true;
  const trPr = row.attrs.trPr as string | null;
  return !!trPr && /<w:tblHeader(?![^>]*w:val="(0|false|off)")/.test(trPr);
}

/** Layout geometry of the first page of a section, and of the pages that follow. */
function makePage(layout: Layout, section: number, prev: PageBox | null, maxWidth: number): PageBox {
  const s = layout.sections[section];
  const g = s.geometry;
  const sameSection = prev?.section === section;
  return {
    ...g,
    section,
    inSection: sameSection ? prev!.inSection + 1 : 0,
    number: !sameSection && s.pageNumberStart != null ? s.pageNumberStart : (prev?.number ?? 0) + 1,
    numberFormat: s.pageNumberFormat ?? null,
    top: prev ? prev.top + prev.height + layout.gap : 0,
    left: (maxWidth - g.width) / 2,
  };
}

/**
 * Whether a text node is in a widget inside `dom` (a list marker, a spacer: not the paragraph's
 * text). Only inside it: while viewing only, the whole document is contenteditable="false", and
 * its text is still text (else no paragraph was ever split across pages there).
 */
function inWidget(t: Text, dom: HTMLElement): boolean {
  const w = t.parentElement?.closest('[contenteditable="false"]');
  return !!w && w !== dom && dom.contains(w);
}

/** The line boxes of a paragraph (client coordinates), from its text and pictures. */
export function lineBoxes(dom: HTMLElement): { top: number; bottom: number }[] {
  const rects: DOMRect[] = [];
  const range = document.createRange();
  const walker = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeType === 1) {
      const el = n as HTMLElement;
      // Pictures and text boxes in the line, not what a shape holds (editor/shapeView.ts); not
      // ProseMirror's empty separator image after a non-editable widget ending a paragraph (a
      // formatting mark ¶, a tab): after a line break it would count as a line of text.
      if (((el.tagName === 'IMG' && !el.classList.contains('ProseMirror-separator')) || el.classList.contains('dx-shape-inline')) && !el.parentElement?.closest('.dx-shape')) {
        const r = el.getBoundingClientRect();
        if (r.width > 0 || r.height > 0) rects.push(r);
      }
      continue;
    }
    // Widgets (list markers, spacers) are not the paragraph's text.
    if (!(n as Text).data.length || inWidget(n as Text, dom)) continue;
    range.selectNodeContents(n);
    for (const r of Array.from(range.getClientRects())) if (r.height > 0) rects.push(r);
  }
  rects.sort((a, b) => a.top - b.top);
  const lines: { top: number; bottom: number }[] = [];
  for (const r of rects) {
    const last = lines[lines.length - 1];
    // Pieces of one line overlap vertically (different font sizes on one line do too).
    if (last && r.top < last.bottom - Math.min(r.height, last.bottom - last.top) / 2) last.bottom = Math.max(last.bottom, r.bottom);
    else lines.push({ top: r.top, bottom: r.bottom });
  }
  // Text is shorter than its line (the line spacing adds room above and below it): widen each
  // line to its whole line box, which is what must fit on the page and what a spacer follows.
  // Consecutive lines meet halfway between their text, or at a spacer between them.
  const box = dom.getBoundingClientRect();
  const cs = getComputedStyle(dom);
  const contentTop = box.top + (parseFloat(cs.borderTopWidth) || 0) + (parseFloat(cs.paddingTop) || 0);
  const contentBottom = box.bottom - (parseFloat(cs.borderBottomWidth) || 0) - (parseFloat(cs.paddingBottom) || 0);
  const gaps = Array.from(dom.querySelectorAll('.dx-line-spacer')).map((el) => el.getBoundingClientRect());
  return lines.map((l, k) => {
    const prev = lines[k - 1];
    const next = lines[k + 1];
    const gapAbove = prev && gaps.find((g) => g.top >= prev.bottom - 1 && g.bottom <= l.top + 1);
    const gapBelow = next && gaps.find((g) => g.top >= l.bottom - 1 && g.bottom <= next.top + 1);
    const top = !prev ? Math.min(l.top, contentTop) : gapAbove ? gapAbove.bottom : (prev.bottom + l.top) / 2;
    const bottom = !next ? Math.max(l.bottom, contentBottom) : gapBelow ? gapBelow.top : (l.bottom + next.top) / 2;
    return { top, bottom };
  });
}

/** Document position of the first character on the line that starts at `clientTop`. */
function lineStart(view: EditorView, dom: HTMLElement, clientTop: number): number | null {
  const range = document.createRange();
  const walker = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT);
  const topOf = (t: Text, i: number) => {
    range.setStart(t, i);
    range.setEnd(t, i + 1);
    const rs = range.getClientRects();
    return rs.length ? rs[rs.length - 1].top : range.getBoundingClientRect().top;
  };
  for (let t = walker.nextNode() as Text | null; t; t = walker.nextNode() as Text | null) {
    if (!t.data.length || inWidget(t, dom)) continue;
    if (topOf(t, t.data.length - 1) < clientTop - 1) continue; // the whole text node is above
    let lo = 0;
    let hi = t.data.length - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (topOf(t, mid) >= clientTop - 1) hi = mid;
      else lo = mid + 1;
    }
    try {
      return view.posAtDOM(t, lo);
    } catch {
      return null;
    }
  }
  return null;
}

/** Whether a paragraph may be split across pages, and how many lines must stay together at each end. */
function splitRule(node: PMNode): { can: boolean; min: number } {
  const pPr = (node.attrs.pPr as string | null) ?? '';
  if (/<w:keepLines(?![^>]*w:val="(0|false|off)")/.test(pPr)) return { can: false, min: 0 };
  // Widow/orphan control (on unless the paragraph turns it off): two lines on each page.
  const widow = !/<w:widowControl[^>]*w:val="(0|false|off)"/.test(pPr);
  return { can: true, min: widow ? 2 : 1 };
}

/** One text line of a table cell, in page px from the top of its row, without the cell's spacers. */
export interface RowLine {
  top: number;
  bottom: number;
  /** Line number inside its paragraph, and the paragraph's line count and split rule. */
  index: number;
  count: number;
  rule: { can: boolean; min: number };
}

/** A cell of a row that may break across pages: its lines and margins (see measureRow). */
export interface RowSplitCell {
  lines: RowLine[];
  padTop: number;
  padBottom: number;
  /** Bottom of the cell's content, from the top of the row, without its spacers. */
  bottom: number;
}

interface CellLine extends RowLine {
  /** Where the line is on screen now (client y), to find its first character. */
  clientTop: number;
  para: HTMLElement;
  /** Position just before the paragraph (a spacer there moves the whole paragraph). */
  paraPos: number;
}

interface CellBox extends RowSplitCell {
  lines: CellLine[];
}

/**
 * A table row's cells as they would be without the spacers pagination put inside them:
 * the text lines of each cell and the row's natural height.
 */
function measureRow(view: EditorView, tr: HTMLTableRowElement, scale: number): { cells: CellBox[]; height: number; full: number } {
  const trTop = tr.getBoundingClientRect().top;
  const cells: CellBox[] = [];
  /** The bottom of each cell's content as it is now, with its spacers. */
  const fullBottoms: number[] = [];
  for (const td of Array.from(tr.cells)) {
    const own = Array.from(td.querySelectorAll('.dx-in-cell'))
      .filter((el) => el.closest('td, th') === td)
      .map((el) => {
        const r = el.getBoundingClientRect();
        return { top: r.top, height: r.height / scale };
      });
    const below = (clientY: number) =>
      (clientY - trTop) / scale - own.reduce((sum, sp) => (sp.top < clientY - 0.5 ? sum + sp.height : sum), 0);
    const cs = getComputedStyle(td);
    // Measured as if the cell were aligned to the top: a vertically centred (or bottom) cell's
    // text moves with the row's height, and a row breaking across pages shows it at the top.
    const firstChild = (Array.from(td.children) as HTMLElement[]).find(
      (c) => !c.classList.contains('dx-in-cell') && !c.classList.contains('column-resize-handle'),
    );
    const shift = firstChild
      ? below(firstChild.getBoundingClientRect().top) - (parseFloat(cs.paddingTop) || 0) - (parseFloat(getComputedStyle(firstChild).marginTop) || 0)
      : 0;
    const local = (clientY: number) => below(clientY) - shift;
    const lines: CellLine[] = [];
    let bottom = parseFloat(cs.paddingTop) || 0;
    let full = bottom;
    for (const child of Array.from(td.children) as HTMLElement[]) {
      if (child.classList.contains('column-resize-handle')) continue;
      const r = child.getBoundingClientRect();
      const margin = parseFloat(getComputedStyle(child).marginBottom) || 0;
      full = Math.max(full, (r.bottom - trTop) / scale - shift + margin);
      if (child.classList.contains('dx-in-cell')) continue;
      bottom = Math.max(bottom, local(r.bottom) + margin);
    }
    fullBottoms.push(full + (parseFloat(cs.paddingBottom) || 0));
    for (const para of Array.from(td.querySelectorAll('.dx-p, table')) as HTMLElement[]) {
      if (para.parentElement?.closest('.dx-shape')) continue; // a text box's text
      if (para.tagName === 'TABLE') {
        // A table in the cell moves to the next page as a whole when it does not fit (its rows
        // are not broken across pages): one line that cannot be split (GOV-ISSUE-017).
        const unit = nestedTable(view, td, para as HTMLTableElement);
        if (unit) lines.push({ top: local(unit.rect.top), bottom: local(unit.rect.bottom), clientTop: unit.rect.top, para: unit.dom, paraPos: unit.pos, index: 0, count: 1, rule: { can: false, min: 0 } });
        continue;
      }
      if (para.closest('td, th') !== td) continue; // a nested table's paragraph
      let node: PMNode;
      let paraPos: number;
      try {
        const $pos = view.state.doc.resolve(view.posAtDOM(para, 0));
        node = $pos.parent;
        paraPos = $pos.before();
      } catch {
        continue;
      }
      if (node.type.name !== 'paragraph') continue;
      let boxes = lineBoxes(para);
      if (!boxes.length) {
        const r = para.getBoundingClientRect(); // an empty paragraph: one line
        boxes = [{ top: r.top, bottom: r.bottom }];
      }
      const rule = splitRule(node);
      boxes.forEach((b, index) =>
        lines.push({ top: local(b.top), bottom: local(b.bottom), clientTop: b.top, para, paraPos, index, count: boxes.length, rule }),
      );
    }
    cells.push({ lines, padTop: parseFloat(cs.paddingTop) || 0, padBottom: parseFloat(cs.paddingBottom) || 0, bottom });
  }
  const minHeight = parseFloat(tr.style.height) || 0;
  const height = Math.max(minHeight, ...cells.map((c) => c.bottom + c.padBottom));
  return { cells, height, full: Math.max(minHeight, ...fullBottoms) };
}

/**
 * A table directly inside the cell `td` (not inside a table nested in it): its position in the
 * document, its outermost element and where that is on screen. Null for any other table.
 */
function nestedTable(view: EditorView, td: HTMLElement, table: HTMLTableElement): { pos: number; dom: HTMLElement; rect: DOMRect } | null {
  if (table.parentElement?.closest('td, th') !== td) return null;
  const para = Array.from(table.querySelectorAll<HTMLElement>('.dx-p')).find((p) => p.closest('table') === table);
  if (!para) return null;
  try {
    const $pos = view.state.doc.resolve(view.posAtDOM(para, 0));
    for (let d = $pos.depth; d > 0; d--) {
      if ($pos.node(d).type.name !== 'table') continue;
      const pos = $pos.before(d);
      const dom = view.nodeDOM(pos) as HTMLElement | null;
      if (!dom || !dom.contains(table)) return null;
      return { pos, dom, rect: dom.getBoundingClientRect() };
    }
  } catch {
    // Not in the document (being redrawn): measured next time.
  }
  return null;
}

/** Whether a row may break across pages (Word's default; not w:cantSplit, not an exact height). */
function rowCanSplit(row: PMNode, m: { cells: CellBox[] }): boolean {
  return !row.attrs.cantSplit && row.attrs.heightRule !== 'exact' && m.cells.some((c) => c.lines.length > 0);
}

/** Bottom of the row's first line (the tallest first line of its cells), from the top of the row. */
function firstLinesBottom(m: { cells: CellBox[] }): number {
  return Math.max(0, ...m.cells.filter((c) => c.lines.length).map((c) => c.lines[0].bottom));
}

/**
 * The first line of a cell that goes to the next page, given the first one that does not fit:
 * the paragraph's widow/orphan control and "keep lines together" can move it back.
 */
function cellCut(c: RowSplitCell, start: number, k: number): number {
  const l = c.lines[k];
  const first = k - l.index; // the paragraph's first line
  if (l.index === 0) return k;
  if (!l.rule.can) return Math.max(first, start);
  let cut = k;
  const rest = l.count - l.index;
  if (rest < l.rule.min) cut = Math.max(first, k - (l.rule.min - rest)); // widow: enough lines go along
  // Orphan: too few lines left behind move too (unless the paragraph already started on an earlier page).
  if (first >= start && cut - first > 0 && cut - first < l.rule.min) cut = first;
  return Math.max(cut, start);
}

/** Where computeBreaks is while it breaks a row across pages (see splitRow). */
export interface RowSplitHost {
  /** The current page's text area, canvas px. */
  area(): { top: number; bottom: number };
  /** The row's part on the current page ends at `partBottom` (canvas px): go on to the next page. */
  breakPage(partBottom: number): void;
  /**
   * Put a spacer about `height` px tall before line `k` of cell `ci`. Returns the height it was
   * given (the one already there when that is within a pixel, see computeBreaks), or null when
   * there is no place for it.
   */
  place(ci: number, k: number, height: number): number | null;
}

/**
 * Breaks a table row across pages: in every cell, the lines that do not fit on the current page
 * move to the next one (a spacer before the first of them), all continuing at the top of that
 * page, below the repeated header rows (`repeatHeight`). `rowY` is where the row starts on the
 * canvas. Each page's part may itself not fit on the next page, so this goes on, page after page,
 * until the rest of every cell fits. Returns what the spacers add to each cell.
 */
export function splitRow(cells: RowSplitCell[], rowY: number, repeatHeight: number, host: RowSplitHost): number[] {
  const start = cells.map(() => 0); // each cell's first line on the current page
  const extra = cells.map(() => 0); // what the cell's spacers add so far
  for (let segment = 0; segment < 500; segment++) {
    const area = host.area();
    const atTop = segment > 0 || rowY <= area.top + 0.5;
    const limit = area.bottom + 0.5;
    const cuts = cells.map((c, ci) => {
      let k = start[ci];
      while (k < c.lines.length && rowY + c.lines[k].bottom + extra[ci] <= limit) k++;
      if (k >= c.lines.length) return null; // the rest of this cell fits
      const cut = cellCut(c, start[ci], k);
      // Moving the cell's first line on this page again gains nothing (a line taller
      // than a page): what is left of that cell spills over.
      return cut > start[ci] || (cut === 0 && !atTop) ? cut : null;
    });
    if (cuts.every((k) => k == null)) break;
    // Where the row's part on this page ends: its tallest cell there, and the cell margin.
    const partTop = segment === 0 ? rowY : area.top;
    let partBottom = partTop;
    cells.forEach((c, ci) => {
      const k = cuts[ci];
      const end = k == null
        ? rowY + c.bottom + extra[ci]
        : k > start[ci] ? rowY + c.lines[k - 1].bottom + extra[ci] : partTop;
      partBottom = Math.max(partBottom, Math.min(end, area.bottom) + c.padBottom);
    });
    host.breakPage(partBottom);
    const next = host.area();
    cells.forEach((c, ci) => {
      const k = cuts[ci];
      if (k == null) return;
      const h = host.place(ci, k, Math.max(0, Math.round(next.top + repeatHeight + c.padTop - (rowY + c.lines[k].top + extra[ci]))));
      if (h == null) return;
      extra[ci] += h;
      start[ci] = k;
    });
  }
  return extra;
}

/**
 * How much taller the spacers inside a row breaking across pages make it: everything below the
 * row moves by that much. Measured on the cells' content with and without the spacers, not on the
 * row's own box: that also holds the row's share of the table's borders, which the row has
 * without spacers too. Counting it, every split row moved the rest of the document half a pixel
 * higher than it is, and after a few split rows the last line above a row break reached under the
 * cover drawn over the break (GOV-ISSUE-017).
 */
export function rowGrowth(view: EditorView, tr: HTMLTableRowElement, scale: number): number {
  const m = measureRow(view, tr, scale);
  return m.full - m.height;
}

/** What computeBreaks works out. */
export interface FlowResult {
  breaks: Break[];
  pages: PageBox[];
  placements?: ColumnPlacement[];
  /**
   * Sections in columns that are laid out in one column (by index): a block is taller than a
   * whole column there (it could only be cut off), or, for a section already laid out so
   * (`columnsOff`), could still be. Absent when there are none.
   */
  singleColumn?: number[];
}

export function computeBreaks(view: EditorView, layout: Layout): FlowResult {
  const root = view.dom as HTMLElement;
  // Blocks moved into their columns are measured where the text flow puts them (columnLayout.css).
  const host = root.parentElement;
  const moved = !!host && layout.sections.some((s) => s.columns) && !!root.querySelector(':scope > .dx-col-shift, :scope > .dx-collapse');
  if (moved) host!.classList.add('dx-col-measure');
  try {
    return computeFlow(view, layout);
  } finally {
    if (moved) host!.classList.remove('dx-col-measure');
  }
}

function computeFlow(view: EditorView, layout: Layout): FlowResult {
  const root = view.dom as HTMLElement;
  const rootRect = root.getBoundingClientRect();
  // The canvas may be zoomed (CSS transform): measurements are converted back to page px.
  const scale = root.offsetWidth > 0 && rootRect.width > 0 ? rootRect.width / root.offsetWidth : 1;
  const rootTop = rootRect.top + (parseFloat(getComputedStyle(root).paddingTop) || 0) * scale;

  // Current spacers, so we can measure where things would sit without them.
  const spacers = Array.from(root.querySelectorAll(':scope > .dx-spacer, .dx-row-spacer, .dx-line-spacer:not(.dx-in-cell)')).map((el) => {
    const r = el.getBoundingClientRect();
    return { top: r.top, height: r.height / scale };
  });
  // A row split across pages is taller by what the spacers in its cells add: everything below
  // it moves by that much (one spacer at the row's bottom edge stands for it).
  const splitRows = new Set<HTMLTableRowElement>();
  for (const el of Array.from(root.querySelectorAll('.dx-in-cell'))) {
    const tr = el.closest('tr');
    if (tr) splitRows.add(tr as HTMLTableRowElement);
  }
  for (const tr of splitRows) {
    const grown = rowGrowth(view, tr, scale);
    if (grown > 0) spacers.push({ top: tr.getBoundingClientRect().bottom - 1, height: grown });
  }
  spacers.sort((a, b) => a.top - b.top);
  // Spacers come in document order, so their tops increase: prefix sums + binary search.
  const prefix = [0];
  for (const sp of spacers) prefix.push(prefix[prefix.length - 1] + sp.height);
  const spacerAbove = (y: number) => {
    let lo = 0;
    let hi = spacers.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (spacers[mid].top < y - 0.5) lo = mid + 1;
      else hi = mid;
    }
    return prefix[lo];
  };
  const sections = layout.sections;
  const maxWidth = Math.max(...sections.map((s) => s.geometry.width));
  // Natural (spacer-free) canvas y of a client y: the first page's text area starts at textTop.
  const top0 = sections[0].geometry.textTop;
  const natural = (clientY: number) => (clientY - rootTop) / scale - spacerAbove(clientY) + top0;

  const sectionOf: number[] = [];
  sections.forEach((s, i) => {
    for (let b = s.firstBlock; b <= s.lastBlock; b++) sectionOf[b] = i;
  });

  const breaks: Break[] = [];
  /**
   * The spacers in table cells as they are now. A new measurement that differs by at most a pixel
   * does not replace them (sameBreaks): rows breaking across pages are laid out with the heights
   * they keep, so their lines and the covers over the breaks stay where they really are.
   */
  const shown = new Map<string, number>();
  for (const b of paginationKey.getState(view.state)?.breaks ?? []) if (b.cell) shown.set(`${b.pos}|${b.cell}`, b.height);
  const pages: PageBox[] = [makePage(layout, sectionOf[0] ?? 0, null, maxWidth)];
  let page = pages[0];
  const nextPage = (section: number) => {
    // A page size that never moves on (or absurdly many pages): stop instead of hanging.
    if (pages.length >= MAX_PAGES) throw PAGE_LIMIT;
    // A section that must start on an odd (even) page gets a blank page in between.
    const s = sections[section];
    if (section !== page.section && (s.start === 'oddPage' || s.start === 'evenPage')) {
      const probe = makePage(layout, section, page, maxWidth);
      if ((probe.number % 2 === 1) !== (s.start === 'oddPage')) {
        page = makePage(layout, page.section, page, maxWidth);
        pages.push(page);
      }
    }
    page = makePage(layout, section, page, maxWidth);
    pages.push(page);
    return page;
  };
  const contentTop = (p: PageBox) => p.top + p.textTop;
  const contentBottom = (p: PageBox) => p.top + p.height - p.textBottom;

  let added = 0; // how far spacers inserted so far move the current content down
  let forceNext = false;
  let prevEnd: number | null = null; // where the previous block ends on the canvas, incl. its bottom margin
  let prevSection = sectionOf[0] ?? 0;

  // ----- text columns (see columnLayout.ts) -----
  const placements: ColumnPlacement[] = [];
  /** Sections to lay out in one column (see FlowResult.singleColumn). */
  const singleColumn = new Set<number>();
  /** The height of a whole column (a page's text area) of a section. */
  const columnRoom = (section: number) => {
    const g = sections[section].geometry;
    return g.height - g.textTop - g.textBottom;
  };
  /**
   * The columns being filled on the current page: which one, its top (canvas), how far the blocks
   * in it are moved up from where the text flow puts them, and where each column ends (with the
   * bottom margin of its last block).
   */
  interface Columns {
    section: number;
    page: PageBox;
    geo: ColumnGeometry;
    index: number;
    top: number;
    shift: number;
    bottoms: number[];
    /** A column break at the end of the last block: the next one starts the next column. */
    breakNext: boolean;
  }
  let col: Columns | null = null;
  const openColumns = (section: number, geo: ColumnGeometry, top: number): Columns =>
    (col = { section, page, geo, index: 0, top, shift: 0, bottoms: [top], breakNext: false });
  /** The columns on this page are done: their lines (w:sep); returns where the lowest one ends. */
  const closeColumns = (): number | null => {
    const c: Columns | null = col;
    if (!c) return null;
    col = null;
    const end = Math.max(c.top, ...c.bottoms);
    const { lefts, widths, separator } = c.geo;
    if (separator) {
      for (let i = 0; i < lefts.length - 1; i++) {
        const x = c.page.marginLeft + (lefts[i] + widths[i] + lefts[i + 1]) / 2;
        (c.page.rules ??= []).push({ x: Math.round(x * 2) / 2, top: Math.round(c.top - c.page.top), bottom: Math.round(end - c.page.top) });
      }
    }
    // Text moved up into a column: what follows continues after the lowest column.
    return c.index > 0 || c.shift ? end : null;
  };
  /** A spacer before `pos` moving the flow by `h` px, up or down (up: a negative margin). */
  const pushSpacer = (pos: number, h: number) => {
    const r = Math.round(h);
    breaks.push(r >= 1 ? { pos, height: r } : { pos, height: 1, collapse: r - 1 });
  };

  try {
    view.state.doc.forEach((node, offset, index) => layoutBlock(node, offset, index));
    closeColumns();
  } catch (err) {
    if (err !== PAGE_LIMIT) throw err;
  }
  // Sections after the last block (e.g. an empty last section) have no pages.
  const out: FlowResult = placements.length ? { breaks, pages, placements } : { breaks, pages };
  if (singleColumn.size) out.singleColumn = [...singleColumn].sort((a, b) => a - b);
  return out;

  /** A block of a section in columns (see columnLayout.ts): into the column it fits in. */
  function layoutColumnBlock(
    node: PMNode,
    offset: number,
    section: number,
    geo: ColumnGeometry,
    m: { top: number; bottom: number; marginTop: number; marginBottom: number },
    fromColumns: number | null,
  ): void {
    const { top, bottom, marginTop, marginBottom } = m;
    const s = sections[section];
    const newSection = prevEnd != null && section !== prevSection;
    const startsPage = newSection && (s.start !== 'continuous' || s.geometry.width !== sections[prevSection].geometry.width ||
      s.geometry.height !== sections[prevSection].geometry.height);
    const hard = forceNext || (node.type.name === 'paragraph' && !!node.attrs.pageBreakBefore);
    const keptAtTop = (isHard: boolean) =>
      !newSection && (layout.topSpacing === 'suppress' || (layout.topSpacing === 'afterPageBreak' && isHard)) ? 0 : marginTop;
    const toNextPage = (isHard: boolean) => {
      closeColumns();
      nextPage(section);
      const kept = keptAtTop(isHard);
      pushSpacer(offset, contentTop(page) - (marginTop - kept) - prevEnd!);
      added = contentTop(page) + kept - top;
      openColumns(section, geo, contentTop(page));
    };
    if (prevEnd == null) {
      openColumns(section, geo, top + added - marginTop);
    } else if (hard || startsPage) {
      toNextPage(hard);
    } else if (!col || col.section !== section || col.page !== page) {
      // The section starts on this page after other text: below it, or below the columns just left.
      if (fromColumns != null) {
        pushSpacer(offset, fromColumns - prevEnd);
        added = fromColumns + marginTop - top;
      }
      openColumns(section, geo, top + added - marginTop);
    }
    const c = col!;
    const breakHere = columnBreakIn(node);
    const forced = c.breakNext || breakHere === 'before';
    const vTop = top + added + c.shift;
    const vBottom = bottom + added + c.shift;
    const atTop = vTop - marginTop <= c.top + 0.5;
    // A column break before anything is in the first column still leaves that column empty (Word).
    if (forced || (vBottom > contentBottom(page) + 0.5 && !atTop)) {
      if (c.index < geo.lefts.length - 1) {
        c.index++;
        c.bottoms[c.index] = c.top;
        c.shift = c.top + (layout.topSpacing === 'suppress' ? 0 : marginTop) - (top + added);
      } else {
        toNextPage(false);
      }
    }
    // Still too tall where it is now, at the top of a column: on a page whose columns start below
    // other text, the next page's whole column may hold it. A block taller than a whole column
    // could only be cut off (Word breaks it between its lines there; a block moves as a whole
    // here): the section is laid out in one column instead (FlowResult.singleColumn).
    if (bottom + added + col!.shift > contentBottom(page) + 0.5) {
      const kept = layout.topSpacing === 'suppress' ? 0 : marginTop;
      const fitsWhole = bottom - top + kept <= columnRoom(section) + 0.5;
      if (fitsWhole && col!.top > contentTop(page) + 0.5) toNextPage(false);
      else if (!fitsWhole) singleColumn.add(section);
    }
    const cc = col!;
    cc.breakNext = breakHere === 'after';
    const i = cc.index;
    const dl = Math.round((geo.lefts[i] - geo.lefts[0]) * 100) / 100;
    const dr = Math.round((geo.lefts[0] + geo.widths[0] - geo.lefts[i] - geo.widths[i]) * 100) / 100;
    const shift = Math.round(cc.shift);
    if (dl || dr || shift) placements.push({ pos: offset, size: node.nodeSize, dl, dr, top: shift });
    cc.bottoms[i] = Math.max(cc.bottoms[i] ?? cc.top, bottom + added + cc.shift + marginBottom);
    if (section !== page.section && !page.alsoSections?.includes(section)) (page.alsoSections ??= []).push(section);
  }

  function layoutBlock(node: PMNode, offset: number, index: number): void {
    const dom = view.nodeDOM(offset) as HTMLElement | null;
    if (!dom || dom.nodeType !== 1) return;
    const rect = dom.getBoundingClientRect();
    const style = getComputedStyle(dom);
    const marginTop = parseFloat(style.marginTop) || 0;
    const marginBottom = parseFloat(style.marginBottom) || 0;
    const top = natural(rect.top);
    const bottom = natural(rect.bottom);
    const section = sectionOf[index] ?? prevSection;
    const geo = sections[section].columns ?? null;
    // Leaving a section in columns: what follows goes below its columns (or to the next page).
    const fromColumns = col && (col.section !== section || !geo) ? closeColumns() : null;
    const off = sections[section].columnsOff;
    if (off && !singleColumn.has(section)) {
      // Laid out in one column for a block taller than a column: kept so while a block could
      // still be (its height here times how much narrower a column is; text and pictures grow
      // no more than that), so the section does not keep switching.
      const g = sections[section].geometry;
      const narrow = (g.width - g.marginLeft - g.marginRight) / Math.max(1, Math.min(...off.widths));
      if ((bottom - top) * narrow + marginTop > columnRoom(section) + 0.5) singleColumn.add(section);
    }
    if (geo) {
      layoutColumnBlock(node, offset, section, geo, { top, bottom, marginTop, marginBottom }, fromColumns);
      forceNext = node.type.name === 'page_break';
      prevEnd = bottom + added + marginBottom;
      prevSection = section;
      return;
    }
    const isTable = node.type.name === 'table';
    const table = isTable ? (dom.tagName === 'TABLE' ? dom : dom.querySelector('table')) : null;
    const rows = table
      ? Array.from((table as HTMLTableElement).rows).filter((r) => !r.classList.contains('dx-row-spacer'))
      : [];

    if (prevEnd != null) {
      const newSection = section !== prevSection;
      const startsPage = newSection && (sections[section].start !== 'continuous' ||
        sections[section].geometry.width !== sections[prevSection].geometry.width ||
        sections[section].geometry.height !== sections[prevSection].geometry.height);
      const wantsBreak = forceNext || startsPage || (node.type.name === 'paragraph' && node.attrs.pageBreakBefore);
      // A table moves as a whole only when not even its first row fits (or, when that row may
      // break across pages, its first line); otherwise it is split.
      let firstBottom = rows.length ? natural(rows[0].getBoundingClientRect().bottom) : bottom;
      if (rows.length && firstBottom + added > contentBottom(page) + 0.5) {
        const m = measureRow(view, rows[0], scale);
        const rowTop = natural(rows[0].getBoundingClientRect().top);
        firstBottom = rowCanSplit(node.child(0), m) ? rowTop + firstLinesBottom(m) : rowTop + m.height;
      }
      let end = isTable ? firstBottom : bottom;
      // A paragraph is split across pages when enough of its lines fit (Word's widow/orphan
      // control); it moves as a whole only when they don't.
      if (!isTable && !wantsBreak && node.type.name === 'paragraph' && bottom + added > contentBottom(page) + 0.5) {
        const rule = splitRule(node);
        const lines = rule.can ? lineBoxes(dom).map((l) => natural(l.bottom)) : [];
        if (rule.can && lines.length >= rule.min * 2) end = lines[rule.min - 1];
      }
      // After text moved up into columns, this block follows the lowest column.
      if (fromColumns != null) added = fromColumns + marginTop - top;
      const overflows = end + added > contentBottom(page) + 0.5 && top + added > contentTop(page) + 0.5;
      if (wantsBreak || overflows) {
        nextPage(section);
        const hard = forceNext || (node.type.name === 'paragraph' && node.attrs.pageBreakBefore);
        const drop = !newSection && (layout.topSpacing === 'suppress' || (layout.topSpacing === 'afterPageBreak' && hard));
        // Word drops the space before a paragraph at the top of the page (topSpacing): the
        // spacer ends that much higher, so the paragraph's text starts at the top margin.
        const kept = drop ? 0 : marginTop;
        const h = contentTop(page) - (marginTop - kept) - prevEnd;
        if (fromColumns != null) pushSpacer(offset, h);
        else breaks.push({ pos: offset, height: Math.max(0, Math.round(h)) });
        // With a spacer in between, the two blocks' vertical margins no longer collapse,
        // so this block (and all after it) moves by the margin too.
        added = contentTop(page) + kept - top;
      } else if (fromColumns != null) {
        pushSpacer(offset, fromColumns - prevEnd);
      }
    }
    // A continuous section that starts on this page: the page shows it too.
    if (section !== page.section && !page.alsoSections?.includes(section)) (page.alsoSections ??= []).push(section);

    if (table && rows.length) {
      const positions: number[] = [];
      node.forEach((_row, rowOffset) => positions.push(offset + 1 + rowOffset));
      const header: number[] = [];
      for (let i = 0; i < node.childCount && isHeaderRow(node.child(i)); i++) header.push(positions[i]);
      const headerHeight = header.length
        ? natural(rows[header.length - 1].getBoundingClientRect().bottom) - natural(rows[0].getBoundingClientRect().top)
        : 0;
      const cols = TableMap.get(node).width;
      for (let i = 0; i < rows.length && i < positions.length; i++) {
        const r = rows[i].getBoundingClientRect();
        const rowTop = natural(r.top);
        if (natural(r.bottom) + added <= contentBottom(page) + 0.5) continue;
        const rowNode = node.child(i);
        const m = measureRow(view, rows[i], scale);
        const splittable = rowCanSplit(rowNode, m);
        // A row that may not break across pages, or whose first line does not fit, moves to the
        // next page as a whole (Word breaks a row between its lines when it may).
        const firstFits = splittable && rowTop + firstLinesBottom(m) + added <= contentBottom(page) + 0.5;
        if (i > 0 && !firstFits && rowTop + added > contentTop(page) + 0.5) {
          const prevBottom = rowTop + added; // rows touch: the previous row ends where this one starts
          nextPage(section);
          // Header rows repeat, unless the break is inside the header itself.
          const repeat = i >= header.length ? header : [];
          const repeatHeight = repeat.length ? headerHeight : 0;
          breaks.push({ pos: positions[i], height: Math.max(0, Math.round(contentTop(page) - prevBottom)), row: { cols, header: repeat } });
          added = contentTop(page) + repeatHeight - rowTop;
        }
        if (!splittable || rowTop + m.height + added <= contentBottom(page) + 0.5) continue;
        // The row breaks across pages: in every cell, the lines that do not fit move to the
        // next page (a spacer before them), all continuing from the top of that page.
        // Header rows repeat above the row's continuation too (not when the row is one of them).
        const repeat = i >= header.length ? header : [];
        const repeatHeight = repeat.length ? headerHeight : 0;
        const tr = rows[i].getBoundingClientRect();
        const td = rows[i].cells[0] ? getComputedStyle(rows[i].cells[0]) : null;
        const border = td && td.borderLeftStyle !== 'none' ? `${td.borderLeftWidth} ${td.borderLeftStyle} ${td.borderLeftColor}` : 'none';
        const extra = splitRow(m.cells, rowTop + added, repeatHeight, {
          area: () => ({ top: contentTop(page), bottom: contentBottom(page) }),
          breakPage: (partBottom) => {
            const before = page;
            nextPage(section);
            (before.cuts ??= []).push({
              top: Math.min(partBottom, contentBottom(before)),
              bottom: contentTop(page),
              left: (tr.left - rootRect.left) / scale,
              right: (tr.right - rootRect.left) / scale,
              border,
              header: repeat.length ? { table: offset, rows: repeat } : undefined,
            });
          },
          place: (ci, k, h) => {
            const l = m.cells[ci].lines[k];
            const pos = l.index === 0 ? l.paraPos : lineStart(view, l.para, l.clientTop);
            if (pos == null) return null;
            const cell = l.index === 0 ? 'block' : 'line';
            // A spacer that is already there and within a pixel stays as it is (see sameBreaks), so
            // what follows is laid out with the height it really has.
            const had = shown.get(`${pos}|${cell}`);
            const height = had != null && Math.abs(had - h) <= 1 ? had : h;
            breaks.push({ pos, height, cell });
            return height;
          },
        });
        // The rows below move by what the spacers added to this row's height.
        const grownBottom = Math.max(rowTop + m.height, ...m.cells.map((c, ci) => rowTop + c.bottom + extra[ci] + c.padBottom));
        added += grownBottom - (rowTop + m.height);
      }
    } else if (node.type.name === 'paragraph' && bottom + added > contentBottom(page) + 0.5 && splitRule(node).can) {
      // The paragraph continues on the next page(s): a spacer before the first line that
      // does not fit, keeping `min` lines together at each end.
      const { min } = splitRule(node);
      const clientLines = lineBoxes(dom);
      const lines = clientLines.map((l) => ({ top: natural(l.top), bottom: natural(l.bottom) }));
      let first = 0; // first line on the current page
      while (lines.length) {
        let k = first;
        while (k < lines.length && lines[k].bottom + added <= contentBottom(page) + 0.5) k++;
        if (k >= lines.length) break;
        k = Math.min(k, lines.length - min);
        if (k - first < min || k <= first) break; // cannot split legally any more: it spills
        const pos = lineStart(view, dom, clientLines[k].top);
        if (pos == null) break;
        const prevBottom = lines[k - 1].bottom + added;
        nextPage(section);
        breaks.push({ pos, height: Math.max(0, Math.round(contentTop(page) - prevBottom)), line: true });
        added = contentTop(page) - lines[k].top;
        first = k;
      }
      // Whatever is left over a page (a very long line, a big picture) spills over.
      while (bottom + added > contentBottom(page) + 0.5 && bottom + added > page.top + page.height + layout.gap && page.height > 0) nextPage(section);
    } else {
      // A block taller than a page spills over; count the pages it reaches.
      while (bottom + added > contentBottom(page) + 0.5 && page.height > 0) {
        if (bottom + added <= page.top + page.height + layout.gap) break;
        nextPage(section);
      }
    }
    forceNext = node.type.name === 'page_break';
    prevEnd = bottom + added + marginBottom;
    prevSection = section;
  }
}

/** Thrown by computeBreaks' nextPage at MAX_PAGES. */
const PAGE_LIMIT = new Error('papyrus: page limit');
