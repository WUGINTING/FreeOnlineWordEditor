// Following a hyperlink in the page (persona-300): a link to another site opens only after the
// user saw where it goes and agreed (a document can hide any address behind friendly text); a
// link to a place in the document (a table of contents entry, w:anchor) goes there instead.
//
// When a click follows a link: always while viewing only (the text is not editable), and with
// Ctrl (⌘ on a Mac) while editing, as in Word, so a plain click still places the cursor. A
// middle click (open in a new tab) is asked about the same way, never opened by the browser.
//
// Only an absolute address opens (http:, https:, mailto:, tel:): a relative target (Word's link to
// a file beside the document, "report.docx") would resolve against this site, not the file's
// folder, so the page says it cannot open it. The question shows the site's host as the browser
// will contact it (punycode, so a look-alike name in another script shows as xn--…) and the
// whole address.
import { isSafeHref, safeHref } from '../docx/links';

/**
 * What the user is asked before a link opens: for a web address its host (punycode) and the whole
 * address; for mailto: / tel: the address.
 */
export function externalLinkQuestion(url: string): string {
  let u: URL | null = null;
  try {
    u = new URL(url);
  } catch {
    u = null;
  }
  if (u && (u.protocol === 'http:' || u.protocol === 'https:')) return `即將開啟外部網站「${u.host}」：\n${u.href}\n確定要開啟嗎？`;
  if (u?.protocol === 'mailto:') return `即將開啟郵件程式，寄信給：${decodeAnchor(u.pathname)}，確定要開啟嗎？`;
  if (u?.protocol === 'tel:') return `即將撥打電話：${decodeAnchor(u.pathname)}，確定要開啟嗎？`;
  return `即將開啟：${url}，確定要開啟嗎？`;
}

/** Said for a link to a file beside the document (a relative target), which the page cannot open. */
export const relativeLinkNotice = (target: string): string => `這個連結指向和原檔案放在一起的檔案（${target}），網頁上無法開啟；請下載文件後用 Word 開啟。`;

/** The schemes a link opens with (see isSafeHref for what may be shown as a link at all). */
const OPENS = /^(?:https?|mailto|tel):$/i;

/** Said when a link to a place in the document points at nothing (the bookmark was deleted). */
export const MISSING_ANCHOR = '找不到這個連結指向的位置（可能已被刪除）。';

export interface FollowHost {
  /** Asks the user; true to go ahead. */
  confirm: (message: string) => boolean;
  /** Opens an address in a new tab (window.open by default). */
  open?: (url: string) => void;
  /** Goes to a bookmark in the document; false when there is none of that name. */
  goToBookmark: (name: string) => boolean;
  onNotice?: (message: string) => void;
}

/**
 * Follows a link target as the page shows it: "#name" goes to that bookmark; an absolute address
 * (http:, https:, mailto:, tel:) opens in a new tab once confirmed; a relative target says it
 * cannot be opened; anything else does nothing. Returns whether something was done.
 */
export function followLink(href: string, host: FollowHost): boolean {
  if (!isSafeHref(href)) return false;
  const target = safeHref(href);
  if (target.startsWith('#')) {
    const name = decodeAnchor(target.slice(1));
    if (host.goToBookmark(name)) return true;
    host.onNotice?.(MISSING_ANCHOR);
    return false;
  }
  if (!/^[a-z][a-z0-9+.-]*:/i.test(target)) {
    host.onNotice?.(relativeLinkNotice(target));
    return false;
  }
  let url: URL;
  try {
    url = new URL(target);
  } catch {
    return false;
  }
  if (!OPENS.test(url.protocol) || (/^https?:$/i.test(url.protocol) && !url.host)) return false;
  if (!host.confirm(externalLinkQuestion(url.href))) return false;
  (host.open ?? ((u: string) => void window.open(u, '_blank', 'noopener,noreferrer')))(url.href);
  return true;
}

function decodeAnchor(s: string): string {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/**
 * A click on a link in the document: its target as the page shows it, and whether the click
 * follows it (see above). The browser itself never follows it (the caller prevents that), so a
 * link in a header shown on the page does not open a site unasked either. Null when the click is
 * not on a link.
 */
export function linkClick(event: MouseEvent, editable: boolean): { href: string; follow: boolean } | null {
  // The main button (click), or the middle one (auxclick: the browser would open a new tab).
  if (event.button !== 0 && event.button !== 1) return null;
  const target = event.target as Element | null;
  const a = (target instanceof Element ? target : (target as Node | null)?.parentElement)?.closest?.('a[href]');
  if (!a) return null;
  // While editing, a plain click places the cursor (or starts a double-click into the header).
  return { href: a.getAttribute('href') ?? '', follow: event.button === 1 || !editable || event.ctrlKey || event.metaKey };
}
