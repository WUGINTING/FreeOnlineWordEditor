import JSZip from 'jszip';
import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '../editor/schema';
import {
  NS, REL_TYPE, attr, child, children, onOff, parseXml, relsPathFor, resolvePartPath, serializeXml, utf16Encoding, xmlText,
} from './xml';
import { Converter } from './convert';
import type { DocxModel, HeaderFooterPart, HeaderFooterType } from './model';
import { parsePage } from './sections';
import { parseNumbering } from './numbering';
import { parseStyles } from './styles';
import { trackRevisionsOn } from './settings';
import { DEFAULT_THEME, readThemeColors, type ThemeColors } from './theme';

const REL_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';

export interface Relationship {
  id: string;
  type: string;
  target: string; // resolved package path, or the URL for external targets
  external: boolean;
}

export type Rels = Map<string, Relationship>;

/** Compatibility settings that change layout (see DocxModel.compat). */
function readCompat(settings: Document | null): NonNullable<DocxModel['compat']> {
  const compat = child(settings?.documentElement ?? null, 'compat');
  let mode = 12; // no compatibilityMode: a Word 2007 document
  for (const s of compat ? children(compat, 'compatSetting') : []) {
    if (attr(s, 'name') === 'compatibilityMode') mode = Number(attr(s, 'val')) || mode;
  }
  return { mode, suppressSpaceAfterPageBreak: onOff(child(compat, 'suppressSpBfAfterPgBrk')) === true };
}

/** The XML parts of a package being opened, read once (see xmlParts); dropped once it is open. */
const opening = new WeakMap<JSZip, Map<string, Uint8Array>>();

/**
 * Every XML part of the package as bytes, read once for the whole opening. A part in UTF-16
 * (which Word reads; everything else here reads parts as UTF-8) is stored in the package again
 * as UTF-8, declaration included, so the writer and the rest read it too; it is saved so. A
 * UTF-8 byte order mark stays (parseXml skips it), so a part that is not rewritten keeps its bytes.
 */
async function xmlParts(zip: JSZip): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir || !/\.(xml|rels)$/i.test(name)) continue;
    let bytes = await file.async('uint8array');
    if (utf16Encoding(bytes)) {
      const { text } = xmlText(bytes);
      zip.file(name, text);
      bytes = new TextEncoder().encode(text);
    }
    out.set(name, bytes);
  }
  return out;
}

/** A part's text (null when the package has no such part). */
async function partText(zip: JSZip, path: string): Promise<string | null> {
  const known = opening.get(zip)?.get(path);
  if (known) return xmlText(known).text;
  const file = zip.file(path);
  if (!file) return null;
  const bytes = await file.async('uint8array');
  const { text, converted } = xmlText(bytes);
  if (converted) zip.file(path, text); // see xmlParts
  return text;
}

/**
 * The relationships of a part. A .rels file that is not well-formed XML counts as empty (the
 * rest of the document still opens); its path is added to `broken` when given.
 */
export async function readRels(zip: JSZip, part: string, broken?: string[]): Promise<Rels> {
  const out: Rels = new Map();
  const path = relsPathFor(part);
  const text = await partText(zip, path);
  if (text == null) return out;
  let doc: Document;
  try {
    doc = parseXml(text);
  } catch {
    if (broken && !broken.includes(path)) broken.push(path);
    return out;
  }
  for (const r of Array.from(doc.getElementsByTagNameNS(NS.rel, 'Relationship'))) {
    const external = r.getAttribute('TargetMode') === 'External';
    const target = r.getAttribute('Target') || '';
    out.set(r.getAttribute('Id') || '', {
      id: r.getAttribute('Id') || '',
      type: r.getAttribute('Type') || '',
      target: external ? target : resolvePartPath(part, target),
      external,
    });
  }
  return out;
}

export async function findMainPart(zip: JSZip, broken?: string[]): Promise<string> {
  const rels = await readRels(zip, '', broken);
  for (const r of rels.values()) if (r.type === REL_TYPE.officeDocument) return r.target;
  return 'word/document.xml';
}

const MIME: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
  svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff',
  emf: 'image/emf', wmf: 'image/wmf',
};

async function loadImages(zip: JSZip, rels: Rels): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for (const r of rels.values()) {
    if (r.type !== REL_TYPE.image || r.external) continue;
    const f = zip.file(r.target);
    if (!f) continue;
    const ext = r.target.split('.').pop()!.toLowerCase();
    out.set(r.id, `data:${MIME[ext] ?? 'application/octet-stream'};base64,${await f.async('base64')}`);
  }
  return out;
}

async function readPart(zip: JSZip, path: string): Promise<Document | null> {
  const text = await partText(zip, path);
  return text == null ? null : parseXml(text);
}

/**
 * An optional part (styles, numbering, settings, a header ...): one that is not well-formed XML
 * is treated as missing, so the rest of the document still opens, and its path is added to
 * `broken`. The writer leaves such parts exactly as they are.
 */
async function readOptionalPart(zip: JSZip, path: string, broken: string[]): Promise<Document | null> {
  try {
    return await readPart(zip, path);
  } catch {
    if (!broken.includes(path)) broken.push(path);
    return null;
  }
}

/** Parse a header/footer part (w:hdr / w:ftr) into an editor document; null when it is missing or not XML (added to `broken`). */
export async function readHeaderFooterPart(zip: JSZip, part: string, broken: string[] = [], theme: ThemeColors = DEFAULT_THEME): Promise<PMNode | null> {
  const xml = await readOptionalPart(zip, part, broken);
  if (!xml) return null;
  const rels = await readRels(zip, part, broken);
  const conv = new Converter(rels, await loadImages(zip, rels), theme);
  const blocks = conv.blocks(xml.documentElement);
  return schema.nodes.doc.create(null, blocks.length ? blocks : [schema.nodes.paragraph.create()]);
}

/** Every header/footer referenced by any section (the last section's first), once per reference. */
async function readHeaderFooters(zip: JSZip, sectPrs: Element[], mainRels: Rels, broken: string[], theme: ThemeColors): Promise<HeaderFooterPart[]> {
  const out: HeaderFooterPart[] = [];
  for (const sectPr of sectPrs) {
    for (const kind of ['header', 'footer'] as const) {
      for (const ref of children(sectPr, kind + 'Reference')) {
        const type = (attr(ref, 'type') ?? 'default') as HeaderFooterType;
        if (type !== 'default' && type !== 'first' && type !== 'even') continue;
        const relId = attr(ref, 'id', NS.r);
        const rel = mainRels.get(relId ?? '');
        if (!rel || out.some((h) => h.relId === relId)) continue;
        const doc = await readHeaderFooterPart(zip, rel.target, broken, theme);
        if (doc) out.push({ kind, type, relId, part: rel.target, doc, dirty: false });
      }
    }
  }
  return out;
}

export interface ReadResult {
  doc: PMNode;
  model: DocxModel;
}

export async function readDocx(data: ArrayBuffer | Uint8Array | Blob): Promise<ReadResult> {
  const zip = await JSZip.loadAsync(data);
  opening.set(zip, await xmlParts(zip));
  try {
    return await readPackage(zip);
  } finally {
    opening.delete(zip);
  }
}

async function readPackage(zip: JSZip): Promise<ReadResult> {
  // Optional parts that are not well-formed XML (see DocxModel.brokenParts).
  const broken: string[] = [];
  const main = await findMainPart(zip, broken);
  // The main part itself must be readable.
  const docXml = await readPart(zip, main);
  if (!docXml) throw new Error('Not a Word document: ' + main + ' is missing');
  const rels = await readRels(zip, main, broken);

  let stylesDoc: Document | null = null;
  let numberingDoc: Document | null = null;
  let settingsDoc: Document | null = null;
  let themeDoc: Document | null = null;
  for (const r of rels.values()) {
    if (r.external) continue;
    // Only for the colours of shapes: an unreadable theme gives Word's default colours (not a notice).
    if (r.type === REL_THEME) themeDoc = await readPart(zip, r.target).catch(() => null);
    if (r.type === REL_TYPE.styles) stylesDoc = await readOptionalPart(zip, r.target, broken);
    if (r.type === REL_TYPE.numbering) numberingDoc = await readOptionalPart(zip, r.target, broken);
    if (r.type === REL_TYPE.settings) settingsDoc = await readOptionalPart(zip, r.target, broken);
  }

  const body = child(docXml.documentElement, 'body');
  if (!body) throw new Error('Word document has no body');
  const sectPr = child(body, 'sectPr');
  // Section breaks: a w:sectPr in the properties of a section's last paragraph.
  // (Also inside block wrappers such as w:sdt; not the old values kept in a w:sectPrChange.)
  const breaks = Array.from(body.getElementsByTagNameNS(NS.w, 'sectPr')).filter(
    (s) => s !== sectPr && s.parentElement?.localName !== 'sectPrChange',
  );

  // Shapes are drawn in the theme's colours.
  const theme = readThemeColors(themeDoc);
  const conv = new Converter(rels, await loadImages(zip, rels), theme);
  const blocks = conv.blocks(body);
  const doc = schema.nodes.doc.create(null, blocks.length ? blocks : [schema.nodes.paragraph.create()]);

  const styles = parseStyles(stylesDoc);
  return {
    doc,
    model: {
      zip,
      page: parsePage(sectPr),
      sectPr: sectPr ? serializeXml(sectPr) : null,
      titlePage: onOff(child(sectPr, 'titlePg')) === true,
      evenAndOdd: onOff(child(settingsDoc?.documentElement, 'evenAndOddHeaders')) === true,
      trackRevisions: trackRevisionsOn(settingsDoc),
      compat: readCompat(settingsDoc),
      numbering: parseNumbering(numberingDoc),
      paragraphStyles: styles.paragraphStyles,
      styles: styles.inheritance,
      css: styles.css,
      headerFooters: await readHeaderFooters(zip, sectPr ? [sectPr, ...breaks] : breaks, rels, broken, theme),
      relIds: [...rels.keys()],
      theme,
      ...(broken.length ? { brokenParts: broken } : {}),
    },
  };
}
