// Which hyperlink targets are allowed: one allowlist for the page (a clickable <a href>), the
// link dialog and paste.
//
// Allowed: http:, https:, mailto:, tel:, an in-document anchor (#bookmark) and a relative
// target (no scheme, e.g. "report.docx" as Word writes for a file beside the document).
// Everything else (javascript:, vbscript:, data:, file:, a UNC path "\\server\share" ...) is
// refused: it would run script in the page or make Word open a local program or network share.
//
// Saving is looser (isWritableHref): a link the DOCX already had to a file or a network share
// ("file:", "\\server\share", "notes:" ...) is written back as it was, since Word asks before
// opening such a link and dropping it would lose the author's link. Only targets that run script
// (javascript:, vbscript:, data: ...) are never written. The page never makes either clickable.

/** Schemes a link may use. */
const SAFE_SCHEME = /^(?:https?|mailto|tel)$/i;

/** Schemes that run script or carry a document: never written to the DOCX. */
const SCRIPT_SCHEME = /^(?:javascript|vbscript|livescript|mocha|data|blob)$/i;

/** C0 control characters (after normalizing): never in a link target. */
// eslint-disable-next-line no-control-regex
const CONTROL = /[\u0000-\u001f]/;

/**
 * The target as a browser (and Word) reads it: ASCII tab / line breaks anywhere and C0 controls
 * or spaces at either end are ignored, so "java\tscript:" is still javascript:.
 */
function normalized(href: string): string {
  // eslint-disable-next-line no-control-regex
  return (href || '').replace(/[\t\n\r]/g, '').replace(/^[\u0000-\u0020]+|[\u0000-\u0020]+$/g, '');
}

/** The scheme of a target ("https"), '' for a relative target or an anchor. */
function schemeOf(h: string): string {
  return /^([a-z][a-z0-9+.-]*):/i.exec(h)?.[1] ?? '';
}

/** Whether a link target is allowed (see above); an empty target is not a link. */
export function isSafeHref(href: string | null | undefined): boolean {
  const h = normalized(href ?? '');
  if (!h) return false;
  if (h.startsWith('#')) return true;
  if (/^[\\/]{2}|^\\/.test(h)) return false; // a UNC path / protocol-relative host
  if (CONTROL.test(h)) return false;
  const scheme = schemeOf(h);
  return scheme ? SAFE_SCHEME.test(scheme) : !h.includes(':') || /^[^:]*[/?#]/.test(h);
}

/**
 * Whether a link target may be written to the DOCX (see above): anything but a script-running
 * scheme or a target with control characters.
 */
export function isWritableHref(href: string | null | undefined): boolean {
  const h = normalized(href ?? '');
  if (!h || CONTROL.test(h)) return false;
  return !SCRIPT_SCHEME.test(schemeOf(h));
}

/** Whether a typed target names its scheme (or is an anchor), as the link dialog asks for. */
export function hasExplicitSafeScheme(href: string): boolean {
  const h = normalized(href);
  return isSafeHref(h) && (h.startsWith('#') || !!schemeOf(h));
}

/** Only let allowed targets become clickable in the page; anything else becomes "#". */
export function safeHref(href: string): string {
  return isSafeHref(href) ? normalized(href) : '#';
}
