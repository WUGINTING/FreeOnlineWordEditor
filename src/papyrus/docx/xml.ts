// Small helpers for walking OOXML with the standard DOM API.

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

export const REL_TYPE = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
} as const;

export const CONTENT_TYPE = {
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
} as const;

/**
 * Parse XML text. A byte order mark before it (U+FEFF once decoded; some tools, e.g. OfficeCLI,
 * write one before <?xml in every part) is not part of the document and is skipped: Word reads
 * such parts, and Edge / Chrome's parser rejects them ("Unexpected characters outside the root
 * element: ï»¿", GOV-FINDING-009). The part itself keeps its bytes.
 */
export function parseXml(text: string): Document {
  const doc = new DOMParser().parseFromString(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('Invalid XML: ' + err.textContent);
  return doc;
}

/**
 * UTF-16 XML (which Word reads) from its first bytes, as XML 1.0 appendix F tells them apart:
 * a byte order mark, or "<?" in UTF-16 without one. Null for anything else (UTF-8).
 */
export function utf16Encoding(bytes: Uint8Array): 'utf-16le' | 'utf-16be' | null {
  const [a, b, c, d] = bytes;
  if (a === 0xff && b === 0xfe) return 'utf-16le';
  if (a === 0xfe && b === 0xff) return 'utf-16be';
  if (a === 0x3c && b === 0 && c === 0x3f && d === 0) return 'utf-16le';
  if (a === 0 && b === 0x3c && c === 0 && d === 0x3f) return 'utf-16be';
  return null;
}

/**
 * The text of an XML part from its bytes: UTF-16 decoded as such (its declaration then says
 * UTF-8, the encoding the text is stored in from then on), else UTF-8 (a byte order mark is kept
 * as U+FEFF, which parseXml skips). `converted`: whether it was UTF-16.
 */
export function xmlText(bytes: Uint8Array): { text: string; converted: boolean } {
  const utf16 = utf16Encoding(bytes);
  if (!utf16) return { text: new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes), converted: false };
  const text = new TextDecoder(utf16).decode(bytes); // without its byte order mark
  return { text: text.replace(/^(\s*<\?xml\b[^>]*?\bencoding\s*=\s*)(["'])[^"']*\2/, '$1$2UTF-8$2'), converted: true };
}

/**
 * Serialize an element exactly as written in the source: same prefixes, same
 * attributes, no namespace declarations added or renamed. (XMLSerializer may
 * rename prefixes to ns1..., which breaks Word's mc:Ignorable / mc:Choice
 * Requires="wps" attributes that name prefixes inside attribute values.)
 * The result must be placed in a part whose root declares the same prefixes —
 * see WORD_NAMESPACES and the writer's root handling.
 */
export function serializeXml(node: Node): string {
  if (node.nodeType === 3 /* text */) return escapeXml(node.nodeValue ?? '');
  if (node.nodeType === 4 /* CDATA */) return `<![CDATA[${node.nodeValue ?? ''}]]>`;
  if (node.nodeType !== 1) return '';
  const el = node as Element;
  let out = '<' + el.tagName;
  for (const a of Array.from(el.attributes)) out += ` ${a.name}="${escapeAttr(a.value)}"`;
  if (!el.firstChild) return out + '/>';
  out += '>';
  for (let c: ChildNode | null = el.firstChild; c; c = c.nextSibling) out += serializeXml(c);
  return out + `</${el.tagName}>`;
}

/** Namespaces Word documents commonly use; declared on every part root we write. */
export const WORD_NAMESPACES: Record<string, string> = {
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  o: 'urn:schemas-microsoft-com:office:office',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  w10: 'urn:schemas-microsoft-com:office:word',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wne: 'http://schemas.microsoft.com/office/word/2006/wordml',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
};

/**
 * How many fragments were parsed. Parsing one builds a whole XML document, the main cost of a
 * save; tests use this to check that a save's cost does not grow with repeated properties.
 */
export const parseStats = { fragments: 0 };

/** Prefixed element / attribute names in XML text (the prefix is group 1). */
const PREFIXED_NAME = /[<\s/]([A-Za-z_][\w.-]*):[A-Za-z_]/g;

/**
 * Namespace declarations for the prefixes some fragments use. Only those are declared (a
 * document with all of WORD_NAMESPACES on its root takes about twice as long to build). Newer
 * Word versions add prefixes (e.g. w16du:dateUtc on every revision) declared only on the part's
 * root: those are bound to a placeholder so the fragment parses. serializeXml writes the names as-is.
 */
function fragmentDeclarations(xmls: string[]): string {
  const used = new Set<string>();
  for (const xml of xmls) for (const m of xml.matchAll(PREFIXED_NAME)) used.add(m[1]);
  let decl = '';
  for (const p of used) {
    if (p === 'xml' || p === 'xmlns') continue;
    decl += ` xmlns:${p}="${Object.prototype.hasOwnProperty.call(WORD_NAMESPACES, p) ? WORD_NAMESPACES[p] : `urn:px-undeclared:${p}`}"`;
  }
  return decl;
}

/** Parse a fragment written by serializeXml (prefixes but no declarations). */
export function parseFragment(xml: string): Element {
  parseStats.fragments++;
  return parseXml(`<px-root${fragmentDeclarations([xml])}>${xml}</px-root>`).documentElement.firstElementChild!;
}

/** Fragments parsed together (see parseFragments): building a document costs more than parsing a small one. */
const FRAGMENTS_PER_DOCUMENT = 64;

/**
 * Many fragments, each as parseFragment would give it (null where parseFragment throws), parsed
 * a few dozen per document: with thousands of them (Reject All restoring every deletion)
 * building a document for each is most of the work.
 */
export function parseFragments(xmls: string[]): (Element | null)[] {
  const out: (Element | null)[] = [];
  const single = (xml: string) => {
    try {
      return parseFragment(xml);
    } catch {
      return null;
    }
  };
  for (let i = 0; i < xmls.length; i += FRAGMENTS_PER_DOCUMENT) {
    const chunk = xmls.slice(i, i + FRAGMENTS_PER_DOCUMENT);
    let els: Element[] | null = null;
    // Each fragment must be exactly one element, so the root's children line up with them.
    if (chunk.every((x) => /^\s*</.test(x) && /\/?>\s*$/.test(x))) {
      try {
        const root = parseXml(`<px-root${fragmentDeclarations(chunk)}>${chunk.join('')}</px-root>`).documentElement;
        const kids: Element[] = [];
        let text = false;
        for (let c = root.firstChild; c; c = c.nextSibling) {
          if (c.nodeType === 1) kids.push(c as Element);
          else text = true;
        }
        if (!text && kids.length === chunk.length) {
          els = kids;
          parseStats.fragments += chunk.length;
        }
      } catch {
        els = null;
      }
    }
    // Something in the chunk is not one well-formed element: each on its own, as before.
    for (let k = 0; k < chunk.length; k++) out.push(els ? els[k] : single(chunk[k]));
  }
  return out;
}

/**
 * Attributes for the root element of a part we write: every namespace the
 * original root declared (plus the common Word ones) and its mc:Ignorable.
 */
export function rootAttributes(originalXml: string | null): string {
  const ns: Record<string, string> = { ...WORD_NAMESPACES };
  // A new part gets the usual list; an existing one keeps exactly what it had.
  let ignorable = originalXml ? '' : 'w14 w15 wp14';
  if (originalXml) {
    const start = /<(?!\?|!)[^>]+>/.exec(originalXml)?.[0] ?? '';
    for (const m of start.matchAll(/xmlns:([\w.-]+)="([^"]*)"/g)) ns[m[1]] = m[2];
    const ig = /mc:Ignorable="([^"]*)"/.exec(start);
    if (ig) ignorable = ig[1];
  }
  const declared = new Set(Object.keys(ns));
  ignorable = ignorable.split(/\s+/).filter((p) => declared.has(p)).join(' ');
  return (
    Object.entries(ns).map(([p, u]) => `xmlns:${p}="${escapeXml(u)}"`).join(' ') +
    (ignorable ? ` mc:Ignorable="${ignorable}"` : '')
  );
}

/** Direct element children, optionally filtered by w: local name. */
export function children(el: Element, local?: string, ns: string = NS.w): Element[] {
  const out: Element[] = [];
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (!local || (c.localName === local && c.namespaceURI === ns)) out.push(c);
  }
  return out;
}

export function child(el: Element | null | undefined, local: string, ns: string = NS.w): Element | null {
  if (!el) return null;
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (c.localName === local && c.namespaceURI === ns) return c;
  }
  return null;
}

/** First descendant with the given local name / namespace. */
export function find(el: Element | null | undefined, local: string, ns: string = NS.w): Element | null {
  if (!el) return null;
  return el.getElementsByTagNameNS(ns, local)[0] ?? null;
}

export function attr(el: Element | null | undefined, local: string, ns: string = NS.w): string | null {
  if (!el) return null;
  return el.getAttributeNS(ns, local) ?? el.getAttribute('w:' + local);
}

export function numAttr(el: Element | null | undefined, local: string, ns: string = NS.w): number | null {
  const v = attr(el, local, ns);
  if (v == null || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

/**
 * On/off properties (w:b, w:i ...): present without w:val means on;
 * w:val="0"/"false"/"off" means off. Returns null when the element is absent.
 */
export function onOff(el: Element | null): boolean | null {
  if (!el) return null;
  const v = attr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

/**
 * Characters XML 1.0 does not allow (C0 controls other than tab / LF / CR, U+FFFE, U+FFFF and
 * unpaired surrogates). A file can't contain them (the parser rejects them), but typed or pasted
 * text can; written out, they make the saved part unreadable to Word and to this editor.
 */
// eslint-disable-next-line no-control-regex
const XML_FORBIDDEN = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
const XML_FORBIDDEN_ALL = new RegExp(XML_FORBIDDEN.source, 'g');

/** The text without the characters XML forbids (the same string when there are none). */
export function xmlSafe(s: string): string {
  return XML_FORBIDDEN.test(s) ? s.replace(XML_FORBIDDEN_ALL, '') : s;
}

/** Escape text content; characters XML forbids are dropped (see xmlSafe). */
export function escapeXml(s: string): string {
  return xmlSafe(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Escape an attribute value; line breaks and tabs must be character references to survive re-parsing. */
export function escapeAttr(s: string): string {
  return escapeXml(s).replace(/\n/g, '&#10;').replace(/\r/g, '&#13;').replace(/\t/g, '&#9;');
}

/** Resolve a relationship target relative to the part that owns the .rels file. */
export function resolvePartPath(basePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const baseDir = basePart.includes('/') ? basePart.slice(0, basePart.lastIndexOf('/')) : '';
  const parts = (baseDir ? baseDir.split('/') : []).concat(target.split('/'));
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

export function relsPathFor(part: string): string {
  const i = part.lastIndexOf('/');
  return (i >= 0 ? part.slice(0, i + 1) : '') + '_rels/' + part.slice(i + 1) + '.rels';
}
