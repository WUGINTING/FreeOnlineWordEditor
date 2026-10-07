import type { PageBox } from './pagination';
import { stripFormattingMarks } from './formattingMarks';

/**
 * Print the laid-out pages (the browser's print dialog also saves a PDF). Each page becomes
 * one sheet of its own paper size (named CSS pages, so landscape sections print landscape):
 * the page with its header/footer, plus copies of the body blocks that reach into it,
 * placed where they are on screen and clipped to the page.
 */
export function printPages(canvas: HTMLElement, pages: PageBox[], title: string, scope = ''): void {
  const frame = document.createElement('iframe');
  frame.setAttribute('aria-hidden', 'true');
  frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
  document.body.append(frame);
  const doc = frame.contentDocument!;
  doc.open();
  doc.write('<!doctype html><html><head><meta charset="utf-8"></head><body></body></html>');
  doc.close();
  if (!fillPrintDocument(doc, canvas, pages, title, scope, false, screenZoom(canvas))) {
    frame.remove();
    return;
  }

  const win = frame.contentWindow!;
  const cleanup = () => setTimeout(() => frame.remove(), 1000);
  win.addEventListener('afterprint', cleanup, { once: true });
  // Wait for the copied pictures and fonts before opening the print dialog.
  const images = Array.from(doc.images).map((img) =>
    img.complete ? Promise.resolve() : new Promise<void>((r) => ((img.onload = () => r()), (img.onerror = () => r()))),
  );
  Promise.all([...images, doc.fonts?.ready ?? Promise.resolve()]).then(() => {
    win.focus();
    win.print();
    // Browsers without afterprint: clean up anyway.
    setTimeout(() => frame.isConnected && frame.remove(), 60000);
  });
}

/**
 * The same printout as one self-contained HTML page: every style rule copied in as text and
 * pictures as data URLs, so a headless browser elsewhere (the server's PDF export) can turn it
 * into the same PDF without loading anything from this site. Null when nothing is laid out.
 */
export function printHtml(canvas: HTMLElement, pages: PageBox[], title: string, scope = ''): string | null {
  const doc = document.implementation.createHTMLDocument(title);
  const meta = doc.createElement('meta');
  meta.setAttribute('charset', 'utf-8');
  doc.head.prepend(meta);
  if (!fillPrintDocument(doc, canvas, pages, title, scope, true, screenZoom(canvas))) return null;
  return '<!doctype html>' + doc.documentElement.outerHTML;
}

/** The text of every style rule this page uses (linked stylesheets included when readable). */
function styleText(): string {
  const out: string[] = [];
  for (const sheet of Array.from(document.styleSheets)) {
    try {
      out.push(Array.from(sheet.cssRules).map((r) => r.cssText).join('\n'));
    } catch {
      // A stylesheet from another origin can't be read; it holds nothing the document needs.
    }
  }
  return out.join('\n');
}

/**
 * How many device pixels the screen has per CSS px (Windows display scaling, browser zoom).
 * Borders thinner than a device pixel (Word's ½ pt table borders) and line boxes are snapped to
 * whole device pixels, so a table's rows are a fraction of a pixel shorter or taller on a 175 %
 * screen than on paper (printing, here or in the server's headless browser, lays out at 1 whatever
 * the screen). Over a long table that adds up to a line: the rows after a page break start lower
 * than the pages were laid out, and the last row runs into the paragraph after the table
 * (GOV-ISSUE-009). The printout is therefore laid out at the screen's scale (CSS zoom) and scaled
 * back to paper size, which gives the same row heights and line breaks as the pages on screen.
 */
function screenZoom(canvas: HTMLElement): number {
  const ratio = canvas.ownerDocument.defaultView?.devicePixelRatio ?? 1;
  return Number.isFinite(ratio) && ratio > 0 && Math.abs(ratio - 1) > 1e-3 ? Math.round(ratio * 1e4) / 1e4 : 1;
}

/**
 * Fills `doc` with the print layout; false when there is nothing to print. `layoutZoom` lays each
 * sheet out at that scale and shrinks it back (see screenZoom).
 */
function fillPrintDocument(
  doc: Document,
  canvas: HTMLElement,
  pages: PageBox[],
  title: string,
  scope: string,
  inlineStyles: boolean,
  layoutZoom = 1,
): boolean {
  const body = canvas.querySelector<HTMLElement>('.dx-host > .dx-doc');
  const pageEls = Array.from(canvas.querySelectorAll<HTMLElement>('.dx-pages > .dx-page'));
  if (!body || !pages.length) return false;
  const root = canvas.closest('.dx-root');
  const canvasRect = canvas.getBoundingClientRect();
  const zoom = canvas.offsetWidth > 0 ? canvasRect.width / canvas.offsetWidth : 1;

  // Where every top-level block is on the (unzoomed) canvas.
  const blocks = Array.from(body.children)
    .filter((c): c is HTMLElement => c instanceof HTMLElement && !c.classList.contains('dx-spacer'))
    .map((el) => {
      const r = el.getBoundingClientRect();
      const style = getComputedStyle(el);
      return {
        el,
        top: (r.top - canvasRect.top) / zoom,
        left: (r.left - canvasRect.left) / zoom,
        width: r.width / zoom,
        height: r.height / zoom,
        marginTop: parseFloat(style.marginTop) || 0,
      };
    });

  const cutEls = Array.from(canvas.querySelectorAll<HTMLElement>('.dx-cuts > *')).map((el) => ({
    el,
    top: parseFloat(el.style.top) || 0,
    left: parseFloat(el.style.left) || 0,
    height: el.offsetHeight,
  }));

  doc.title = title;

  // The app's styles, the document's own styles (styles.xml) and the print layout.
  if (inlineStyles) {
    const all = doc.createElement('style');
    all.textContent = styleText();
    doc.head.append(all);
  } else {
    for (const s of Array.from(document.querySelectorAll('style, link[rel="stylesheet"]'))) doc.head.append(s.cloneNode(true));
  }
  const own = root?.querySelector(':scope > style');
  if (own) doc.head.append(own.cloneNode(true));
  const sizes = new Map<string, { w: number; h: number }>();
  for (const p of pages) sizes.set(`dxp${Math.round(p.width)}x${Math.round(p.height)}`, { w: p.width, h: p.height });
  const css = doc.createElement('style');
  css.textContent =
    '@page{margin:0}html,body{margin:0;padding:0;background:#fff}' +
    '.dx-print-page{position:relative;overflow:hidden;break-after:page;page-break-after:always;background:#fff}' +
    '.dx-print-page:last-child{break-after:auto;page-break-after:auto}' +
    '.dx-print-page .dx-page{position:absolute;left:0;top:0;box-shadow:none}' +
    '.dx-print-zoom{position:absolute;left:0;top:0;transform-origin:0 0}' +
    // !important: placed where its border box was measured, so no margin of its own (a style's
    // "space before", e.g. .dx-ps-Heading1, would otherwise push it onto the next block).
    '.dx-print-page .dx-print-block{position:absolute;box-sizing:border-box;margin:0!important}' +
    // Headers and footers are dimmed on screen only.
    '.dx-print-root .dx-header,.dx-print-root .dx-footer{opacity:1}' +
    // Paper has no gap between pages: a row-break cover is page-coloured throughout.
    '.dx-print-root .dx-cut{background:var(--dx-page)!important}' +
    '.dx-root.dx-print-root{padding:0;background:#fff;min-height:0}' +
    [...sizes].map(([name, s]) => `@page ${name}{size:${s.w}px ${s.h}px;margin:0}.${name}{page:${name}}`).join('');
  doc.head.append(css);

  const printRoot = doc.createElement('div');
  // The editor's scope class, so the document's own styles (styles.xml) apply to the copy.
  printRoot.className = `dx-root dx-print-root ${scope}`.trim();
  const contentHeight = root instanceof HTMLElement ? root.style.getPropertyValue('--dx-content-h') : '';
  if (contentHeight) printRoot.style.setProperty('--dx-content-h', contentHeight);
  pages.forEach((p, i) => {
    const sheet = doc.createElement('div');
    sheet.className = `dx-print-page dxp${Math.round(p.width)}x${Math.round(p.height)}`;
    sheet.style.cssText = `width:${p.width}px;height:${p.height}px`;
    // What the page holds, laid out at the screen's scale when that is not 1 (see screenZoom).
    let area = sheet;
    if (layoutZoom !== 1) {
      // Laid out taller than the paper before it is scaled back: the sheet must not be broken
      // into pages where that layout crosses the paper's edge (contain makes it one piece).
      sheet.style.contain = 'strict';
      area = doc.createElement('div');
      area.className = 'dx-print-zoom';
      area.style.cssText = `width:${p.width}px;height:${p.height}px;zoom:${layoutZoom};transform:scale(${1 / layoutZoom})`;
      sheet.append(area);
    }
    // Page background with its header and footer.
    const pageEl = pageEls[i]?.cloneNode(true) as HTMLElement | undefined;
    if (pageEl) {
      // The screen's 「本節在 Word 中為 2 欄」 badge is not part of the page.
      for (const badge of Array.from(pageEl.querySelectorAll('.dx-cols-badge'))) badge.remove();
      // Nor the formatting marks of its header/footer (Word doesn't print them).
      stripFormattingMarks(pageEl);
      pageEl.style.left = '0';
      pageEl.style.top = '0';
      area.append(pageEl);
    }
    // Body blocks reaching into this page, clipped by the sheet (a table split across
    // pages is shown on each page from the row the page starts with).
    const content = doc.createElement('div');
    content.className = 'dx-doc';
    content.style.cssText = 'position:absolute;inset:0';
    for (const b of blocks) {
      if (b.top + b.height <= p.top || b.top >= p.top + p.height) continue;
      const copy = b.el.cloneNode(true) as HTMLElement;
      copy.classList.add('dx-print-block');
      copy.style.top = `${b.top - p.top}px`;
      copy.style.left = `${b.left - p.left}px`;
      copy.style.width = `${b.width}px`;
      content.append(copy);
    }
    for (const sel of Array.from(content.querySelectorAll('.ProseMirror-selectednode, .selectedCell'))) {
      sel.classList.remove('ProseMirror-selectednode', 'selectedCell');
    }
    // 顯示／隱藏編輯標記 is for the screen: Word doesn't print the marks.
    stripFormattingMarks(content);
    area.append(content);
    // Where a table row breaks across pages: the cover over its cell borders and the header
    // rows repeated above its continuation (see DocxEditor.renderCuts).
    const cuts = doc.createElement('div');
    cuts.style.cssText = 'position:absolute;inset:0';
    for (const c of cutEls) {
      if (c.top + c.height <= p.top || c.top >= p.top + p.height) continue;
      const copy = c.el.cloneNode(true) as HTMLElement;
      copy.style.top = `${c.top - p.top}px`;
      copy.style.left = `${c.left - p.left}px`;
      // Repeated header rows are copies of rows that may show formatting marks.
      stripFormattingMarks(copy);
      cuts.append(copy);
    }
    area.append(cuts);
    printRoot.append(sheet);
  });
  // Text for screen readers only (「刪除：」「（刪除結束）」 around a deletion …) is not on paper,
  // nor in a PDF's text.
  for (const sr of Array.from(printRoot.querySelectorAll('.dx-sr'))) sr.remove();
  doc.body.append(printRoot);
  return true;
}
