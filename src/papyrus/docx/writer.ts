import JSZip from 'jszip';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { TableMap } from 'prosemirror-tables';
import { pxToEmu } from '../units';
import {
  CONTENT_TYPE, NS, REL_TYPE, child, escapeXml, parseFragment, parseXml, relsPathFor, rootAttributes, serializeXml, xmlSafe,
} from './xml';
import { A4, type DocxModel, type HeaderFooterKind, type HeaderFooterType } from './model';
import {
  cellInfo, gridSpanOf, parseKept, patchDrawing, readCellModel, readDrawing, sameBorder, writePPr, writeRPr, writeTcPr, writeTrPr,
  BORDER_SIDES, type CellModel, type DrawingPatch, type ParaModel, type RowModel, type RunModel,
} from './props';
import { Lru, memo } from './memo';
import { LayerStack, parseLayers, type Layer } from './wrappers';
import { findMainPart, readRels, type Relationship } from './reader';
import { numberingToWrite, serializeNumbering, usedNumIds } from './numbering';
import { blankPackage } from './template';
import { referencedIds } from './sections';
import { writeComments } from './comments';
import { writeTrackRevisions } from './settings';
import { isWritableHref } from './links';
import { readTitleRecord, systemTitle, writeTitleRecord } from './titleMarker';
import { addTocStyles } from './tocStyles';
import { hasEditedText, rewriteTextBoxes, textFrames, type ShapeModel } from './shapes';
import { hasChangedShape, patchShapeXml } from './shapeWrite';
import { WATERMARK_IMAGE_REF } from './watermark';
import { columnWidths, defaultTblPr, tableCells } from './tableLayout';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Start tag of a tracked insertion or move destination wrapper. */
const REVISION_WRAPPER = /^<(?:[\w.-]+:)?(ins|moveTo)[\s>/]/;

const IMAGE_EXT: Record<string, string> = {
  'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp',
  'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/tiff': 'tiff', 'image/emf': 'emf', 'image/wmf': 'wmf',
};

interface NewRel {
  id: string;
  type: string;
  target: string; // relative to the owning part for internal targets
  external: boolean;
}

/** The format of picture bytes from their signature, as a file extension; null when unknown. */
export function sniffImageExt(b: Uint8Array): string | null {
  const at = (i: number, ...v: number[]) => v.every((x, k) => b[i + k] === x);
  if (at(0, 0x89, 0x50, 0x4e, 0x47)) return 'png';
  if (at(0, 0xff, 0xd8, 0xff)) return 'jpeg';
  if (at(0, 0x47, 0x49, 0x46, 0x38)) return 'gif';
  if (at(0, 0x42, 0x4d)) return 'bmp';
  if (at(0, 0x52, 0x49, 0x46, 0x46) && at(8, 0x57, 0x45, 0x42, 0x50)) return 'webp';
  if (at(0, 0x49, 0x49, 0x2a, 0x00) || at(0, 0x4d, 0x4d, 0x00, 0x2a)) return 'tiff';
  if (at(0, 0x01, 0x00, 0x00, 0x00) && at(40, 0x20, 0x45, 0x4d, 0x46)) return 'emf';
  if (at(0, 0xd7, 0xcd, 0xc6, 0x9a) || at(0, 0x01, 0x00, 0x09, 0x00) || at(0, 0x02, 0x00, 0x09, 0x00)) return 'wmf';
  const head = new TextDecoder().decode(b.subarray(0, 1024));
  if (/^\s*(<\?xml[^>]*>\s*)?(<!--[\s\S]*?-->\s*|<!DOCTYPE[^>]*>\s*)*<svg[\s>]/i.test(head)) return 'svg';
  return null;
}

/**
 * The extension a picture is stored under: what its bytes are, else what its data URL says
 * (never "png" for bytes that are not PNG); null when it is not a picture.
 */
function imageExt(mime: string, bytes: Uint8Array): string | null {
  const sniffed = sniffImageExt(bytes);
  if (sniffed) return sniffed;
  const type = mime.toLowerCase();
  const known = IMAGE_EXT[type] ?? { 'image/jpg': 'jpeg', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf' }[type];
  if (known && known !== 'png') return known;
  const sub = /^image\/([a-z0-9]+)$/.exec(type)?.[1];
  return sub && sub !== 'png' && sub !== 'xml' && sub !== 'rels' ? sub : null;
}

/** Something a save could not write as the editor shows it (the save itself went through). */
export interface WriteWarning {
  /** image: a picture that could not be embedded; part: a package part that was left as it was. */
  kind: 'image' | 'part';
  message: string;
  /** The package part concerned, for kind "part". */
  part?: string;
}

export interface WriteOptions {
  /** Receives what could not be written (see WriteWarning); a save never fails for these. */
  warnings?: WriteWarning[];
  /**
   * The document's name: written as the file's title (docProps/core.xml dc:title) when it has
   * none, an empty one or Word's generic "Word Document" (GOV-FINDING-012), or one the editor
   * wrote itself for an earlier name (see titleMarker.ts). A title the author set is kept;
   * without this option core.xml is not touched at all.
   */
  title?: string;
}

const REL_CORE = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CT_CORE = 'application/vnd.openxmlformats-package.core-properties+xml';
const NS_DC = 'http://purl.org/dc/elements/1.1/';
/** Titles that say nothing: none, blank, or what blank Word templates carry. */
const GENERIC_TITLE = /^\s*(word document)?\s*$/i;

/** New core properties: just the title (nothing else about the file is known). */
function coreXml(title: string): string {
  return (
    XML_HEAD +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    `xmlns:dc="${NS_DC}" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" ` +
    `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(title)}</dc:title></cp:coreProperties>`
  );
}

/**
 * Give the package the title `title` unless it has a real one (see WriteOptions.title). The
 * core properties part is edited as text, so everything else in it stays byte for byte; a
 * package without one gets a new docProps/core.xml with its relationship and content type.
 */
async function writeCoreTitle(zip: JSZip, title: string, overrides: Record<string, string>): Promise<void> {
  const rootRels = await readRels(zip, '', []);
  const rel = [...rootRels.values()].find((r) => r.type === REL_CORE && !r.external);
  if (rel) {
    const file = zip.file(rel.target);
    if (!file) return;
    const xml = await file.async('string');
    let doc: Document;
    try {
      doc = parseXml(xml);
    } catch {
      return; // not readable: left as it is
    }
    const el = doc.getElementsByTagNameNS(NS_DC, 'title')[0];
    // A real title stays, unless it is the one the editor wrote itself (renamed since; titleMarker.ts).
    const record = await readTitleRecord(zip);
    if (el && !GENERIC_TITLE.test(el.textContent ?? '') && !systemTitle(el.textContent, record)) return;
    const value = escapeXml(title);
    let out: string;
    if (el) {
      const tag = el.tagName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      // A self-closing tag keeps its attributes but not its "/".
      out = xml.replace(
        new RegExp(`<${tag}((?:\\s[^>]*?)?)\\s*(?:/>|>[\\s\\S]*?</${tag}\\s*>)`),
        (_m, attrs: string) => `<${el.tagName}${attrs}>${value}</${el.tagName}>`,
      );
    } else {
      const prefix = doc.documentElement.lookupPrefix(NS_DC);
      const open = /<(?:[\w.-]+:)?coreProperties\b[^>]*?(\/?)>/.exec(xml);
      if (!open || open[1]) return; // an empty root element: nothing sensible to add to
      const tag = prefix ? `${prefix}:title` : 'dc:title';
      const add = `<${tag}${prefix ? '' : ` xmlns:dc="${NS_DC}"`}>${value}</${tag}>`;
      out = xml.slice(0, open.index + open[0].length) + add + xml.slice(open.index + open[0].length);
    }
    if (out !== xml) zip.file(rel.target, out);
    await writeTitleRecord(zip, title, record, overrides);
    return;
  }
  // No core properties yet: add docProps/core.xml (the root .rels must be readable to point at it).
  const relsFile = zip.file('_rels/.rels');
  if (!relsFile) return;
  const rels = await relsFile.async('string');
  try {
    parseXml(rels);
  } catch {
    return;
  }
  let part = 'docProps/core.xml';
  for (let n = 1; zip.file(part); n++) part = `docProps/core${n}.xml`;
  let n = 1;
  while (rootRels.has(`rIdPxCore${n}`)) n++;
  const close = rels.lastIndexOf('</Relationships>');
  if (close < 0) return;
  zip.file(part, coreXml(title));
  zip.file('_rels/.rels', rels.slice(0, close) + `<Relationship Id="rIdPxCore${n}" Type="${REL_CORE}" Target="${part}"/>` + rels.slice(close));
  overrides['/' + part] = CT_CORE;
  await writeTitleRecord(zip, title, await readTitleRecord(zip), overrides);
}

/** Largest wp:docPr id per package entry. An entry object is replaced when its part is written, so this never goes stale. */
const drawingIdsOf = new WeakMap<object, number>();

/** The largest wp:docPr id in any XML part of the package (0 when there is none). */
async function maxDrawingId(zip: JSZip): Promise<number> {
  let max = 0;
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir || !/\.xml$/i.test(name)) continue;
    let partMax = drawingIdsOf.get(file);
    if (partMax === undefined) {
      partMax = 0;
      const xml = await file.async('string');
      for (const m of xml.matchAll(/<(?:\w+:)?docPr\b[^>]*?\sid="(\d+)"/g)) partMax = Math.max(partMax, Number(m[1]));
      drawingIdsOf.set(file, partMax);
    }
    max = Math.max(max, partMax);
  }
  return max;
}

/** Media files and drawing ids shared by every part written in one save. */
class MediaPool {
  files = new Map<string, Uint8Array>(); // package path -> bytes
  drawingId = 1;
  constructor(private zip: JSZip, readonly warnings: WriteWarning[]) {}

  /** Shared pictures added in this save, by folder and source (see `add`). */
  private shared = new Map<string, string>();

  /**
   * A new media file for a picture. With `share` (its source), a picture already added with
   * the same source is reused: a watermark's picture is in every header part, and each part gets
   * its own relationship to one file, as Word does. Other pictures always get a file of their
   * own, so a part's pictures never depend on another part's.
   */
  add(dir: string, ext: string, bytes: Uint8Array, share?: string): string {
    const key = share ? `${dir}
${share}` : null;
    const known = key ? this.shared.get(key) : undefined;
    if (known) return known;
    let n = 1;
    let path: string;
    do path = `${dir}media/px-image${n++}.${ext}`;
    while (this.zip.file(path) || this.files.has(path));
    this.files.set(path, bytes);
    if (key) this.shared.set(key, path);
    return path;
  }
}

/** Collects the relationships created while serializing one part. */
class Package {
  rels: NewRel[] = [];
  private imageBySrc = new Map<string, string>();
  private linkByHref = new Map<string, string>();
  private counter = 0;

  constructor(private usedIds: Set<string>, private part: string, readonly pool: MediaPool) {}

  private newId(): string {
    let id: string;
    do id = `rIdPx${++this.counter}`;
    while (this.usedIds.has(id));
    this.usedIds.add(id);
    return id;
  }

  /**
   * Relationship id of an embedded picture; null (with a warning) when `src` can't be embedded:
   * a remote URL, a data: URL that doesn't decode, or bytes that are not a picture. `shared`:
   * the media file may be one another part uses too (see MediaPool.add).
   */
  image(src: string, shared = false): string | null {
    const cached = this.imageBySrc.get(src);
    if (cached) return cached;
    const decoded = decodeImage(src);
    if ('error' in decoded) {
      this.pool.warnings.push({ kind: 'image', message: decoded.error });
      return null;
    }
    const { ext, bytes } = decoded;
    const path = this.pool.add(dirOf(this.part), ext, bytes, shared ? src : undefined);
    const id = this.newId();
    this.rels.push({ id, type: REL_TYPE.image, target: relativeTarget(this.part, path), external: false });
    this.imageBySrc.set(src, id);
    return id;
  }

  link(href: string): string {
    if (!isWritableHref(href)) throw new Error('refused hyperlink target');
    const cached = this.linkByHref.get(href);
    if (cached) return cached;
    const id = this.newId();
    this.rels.push({ id, type: REL_TYPE.hyperlink, target: href, external: true });
    this.linkByHref.set(href, id);
    return id;
  }
}

function dirOf(part: string): string {
  return part.includes('/') ? part.slice(0, part.lastIndexOf('/') + 1) : '';
}

function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

/** The bytes of URL-encoded data (RFC 2397 without ";base64"): %XX is a byte, anything else its UTF-8. Never throws. */
function percentToBytes(data: string): Uint8Array {
  const out: number[] = [];
  const utf8 = new TextEncoder();
  for (let i = 0; i < data.length; ) {
    if (data[i] === '%' && /^[0-9A-Fa-f]{2}$/.test(data.slice(i + 1, i + 3))) {
      out.push(parseInt(data.slice(i + 1, i + 3), 16));
      i += 3;
      continue;
    }
    const cp = data.codePointAt(i)!;
    const ch = String.fromCodePoint(cp);
    for (const b of utf8.encode(ch)) out.push(b);
    i += ch.length;
  }
  return Uint8Array.from(out);
}

type DecodedImage = { ext: string; bytes: Uint8Array } | { error: string };

/** Recent decodeImage results: the same picture is looked at more than once per save. */
const decodedImages = new Lru<DecodedImage>(16);

/** A picture's data: URL decoded for embedding, or why it can't be embedded. */
function decodeImage(src: string): DecodedImage {
  return memo(decodedImages, src, () => {
    const comma = src.indexOf(',');
    if (!/^data:/i.test(src) || comma < 0) {
      return { error: `Picture not embedded: not a data: URL (${src.slice(0, 80)})` };
    }
    const params = src.slice(5, comma).split(';');
    const mime = params[0].trim() || 'text/plain';
    const base64 = params.slice(1).some((p) => p.trim().toLowerCase() === 'base64');
    let bytes: Uint8Array;
    try {
      bytes = base64 ? base64ToBytes(src.slice(comma + 1)) : percentToBytes(src.slice(comma + 1));
    } catch {
      return { error: `Picture not embedded: its data could not be decoded (${mime})` };
    }
    const ext = imageExt(mime, bytes);
    return ext ? { ext, bytes } : { error: `Picture not embedded: the data is not a picture (${mime})` };
  });
}

// ---------------------------------------------------------------------------
// Body serialization

class BodyWriter {
  /** w14:paraId values already written; copies made by splitting/pasting must not repeat them. */
  private paraIds = new Set<string>();

  constructor(private pkg: Package, private model: DocxModel) {}

  /** Blocks, re-opening shared w:sdt / w:customXml wrappers and re-joining split paragraphs. */
  blocks(nodes: readonly PMNode[]): string {
    this.root ??= nodes;
    const stack = new LayerStack();
    let out = '';
    for (let i = 0; i < nodes.length; ) {
      const node = nodes[i];
      out += stack.moveTo(parseLayers(node.attrs.wrap));
      if (node.type.name === 'paragraph' || node.type.name === 'page_break') {
        const size = paragraphGroupSize(nodes, i);
        out += this.paragraphGroup(nodes.slice(i, i + size));
        i += size;
      } else {
        out += this.block(node);
        i++;
      }
    }
    return out + stack.closeAll();
  }

  private block(node: PMNode): string {
    switch (node.type.name) {
      case 'raw_block':
        return node.attrs.xml;
      case 'table':
        return this.table(node);
      default:
        return '';
    }
  }

  /** One w:p: a paragraph, or paragraph pieces and page breaks that came from one w:p. */
  private paragraphGroup(group: PMNode[]): string {
    const paras = group.filter((n) => n.type.name === 'paragraph');
    const attrs: Record<string, any> = paras.length
      ? paras[0].attrs
      : { ...(group[0].attrs.para ? JSON.parse(group[0].attrs.para) : {}) };
    // The section break is on the w:p's mark, which ends its last piece.
    const sectPr = paras.length ? paras[paras.length - 1].attrs.sectPr : attrs.sectPr;
    const pPr = writePPr(attrs.pPr ?? null, paraModel(attrs), sectPr ?? null);
    // One wrapper stack for the whole w:p, so a w:ins around a page break stays one w:ins.
    const stack = new LayerStack((l) => this.openLayer(l));
    let body = '';
    for (const n of group) {
      if (n.type.name === 'paragraph') {
        n.forEach((child) => {
          body += stack.moveTo(this.inlineLayers(child));
          body += this.inline(child);
        });
      } else {
        body += stack.moveTo(this.wrapperLayers(n.attrs.layers, n.attrs.href));
        body += `<w:r${n.attrs.runAttrs ?? ''}>${n.attrs.rPr ?? ''}<w:br w:type="page"/></w:r>`;
      }
    }
    body += stack.closeAll();
    return `<w:p${this.pAttrs(attrs.pAttrs)}>${pPr}${body}</w:p>`;
  }

  /** Nodes of the part being written (the first blocks() call), for choosing new revision ids. */
  private root: readonly PMNode[] | null = null;
  private openedRevisions = new Set<string>();
  private nextRevisionId: number | null = null;

  /**
   * Start tag of an inline wrapper. A tracked insertion split in two (text typed inside it is
   * not tracked) is written as two w:ins; the second gets a new w:id, as Word does.
   */
  private openLayer(l: Layer): string {
    if (!REVISION_WRAPPER.test(l.open)) return l.open;
    if (!this.openedRevisions.has(l.id)) {
      this.openedRevisions.add(l.id);
      return l.open;
    }
    return l.open.replace(/^(<[^>]*?\s(?:[\w.-]+:)?id=")[^"]*"/, `$1${this.newRevisionId()}"`);
  }

  private newRevisionId(): number {
    if (this.nextRevisionId == null) {
      let max = 0;
      const scan = (v: unknown) => {
        if (typeof v === 'string') {
          for (const m of v.matchAll(/:id=\\?"(\d+)\\?"/g)) max = Math.max(max, Number(m[1]));
        } else if (Array.isArray(v)) v.forEach(scan);
      };
      const visit = (n: PMNode) => {
        Object.values(n.attrs).forEach(scan);
        for (const m of n.marks) Object.values(m.attrs).forEach(scan);
        n.forEach(visit);
      };
      (this.root ?? []).forEach(visit);
      this.nextRevisionId = max + 1;
    }
    return this.nextRevisionId++;
  }

  private pAttrs(raw: string | null): string {
    if (!raw) return '';
    const id = /w14:paraId="([^"]+)"/.exec(raw)?.[1];
    if (!id) return raw;
    if (this.paraIds.has(id)) return raw.replace(/\s+w14:(paraId|textId)="[^"]*"/g, '');
    this.paraIds.add(id);
    return raw;
  }

  /** Wrapper layers around an inline node, with hyperlinks rebuilt from the link mark. */
  private inlineLayers(node: PMNode): Layer[] {
    const wrap = node.marks.find((m) => m.type.name === 'inlineWrap');
    const link = node.marks.find((m) => m.type.name === 'link');
    return this.wrapperLayers(wrap?.attrs.layers, link ? (link.attrs.href as string) : null);
  }

  /**
   * Wrapper layers (JSON), with the hyperlink rebuilt from its current target. A target that runs
   * script (javascript: ..., see isWritableHref in links.ts) is not written: the text stays, as
   * plain runs, without the w:hyperlink - also for such a link in a loaded file.
   */
  private wrapperLayers(json: string | null | undefined, href: string | null): Layer[] {
    const layers = parseLayers(json);
    const at = layers.findIndex((l) => l.kind === 'link');
    if (href == null || !isWritableHref(href)) {
      if (at >= 0) layers.splice(at, 1); // the user removed the link, or it is unsafe
      return layers;
    }
    const target = href.startsWith('#')
      ? `w:anchor="${escapeXml(href.slice(1))}"`
      : `r:id="${this.pkg.link(href)}"`;
    if (at >= 0) {
      layers[at] = { ...layers[at], open: `<w:hyperlink ${target}${layers[at].open}>` };
    } else {
      // A w:hyperlink can't be inside a w:ins / w:moveTo (text typed with tracked changes on).
      const ins = layers.findIndex((l) => REVISION_WRAPPER.test(l.open));
      const link = { id: 'link:' + href, open: `<w:hyperlink ${target}>`, close: '</w:hyperlink>' };
      if (ins >= 0) layers.splice(ins, 0, link);
      else layers.push(link);
    }
    return layers;
  }

  /** w:rPr for an inline node: the original run properties patched with the editor's marks. */
  private rPr(marks: readonly Mark[]): string {
    const run = marks.find((m) => m.type.name === 'run');
    return writeRPr(run?.attrs.rPr ?? null, runModel(marks));
  }

  private runOpen(marks: readonly Mark[]): string {
    const run = marks.find((m) => m.type.name === 'run');
    return `<w:r${run?.attrs.attrs ?? ''}>`;
  }

  private inline(node: PMNode): string {
    switch (node.type.name) {
      case 'text':
        return `${this.runOpen(node.marks)}${this.rPr(node.marks)}${runContent(node.text!, 'w:t')}</w:r>`;
      case 'hard_break': {
        const { type, clear } = node.attrs;
        const br = type === 'cr'
          ? '<w:cr/>'
          : `<w:br${type ? ` w:type="${escapeXml(type)}"` : ''}${clear ? ` w:clear="${escapeXml(clear)}"` : ''}/>`;
        return `${this.runOpen(node.marks)}${this.rPr(node.marks)}${br}</w:r>`;
      }
      case 'tab':
        return `${this.runOpen(node.marks)}${this.rPr(node.marks)}<w:tab/></w:r>`;
      case 'field':
        if (node.attrs.xml && !this.fieldEdited(node)) return node.attrs.xml;
        return (
          `<w:fldSimple w:instr=" ${escapeXml(node.attrs.instr)} ">` +
          // Keep the cached result exactly as it was (Word recomputes it when opening).
          `${this.runOpen(node.marks)}${this.rPr(node.marks)}${node.attrs.text ? textRun(node.attrs.text) : ''}</w:r></w:fldSimple>`
        );
      case 'image':
        return this.image(node);
      case 'raw_inline':
        return node.attrs.shape ? this.shapeRun(node) : node.attrs.xml.includes(WATERMARK_IMAGE_REF) ? this.watermarkPicture(node) : node.attrs.xml;
      default:
        return '';
    }
  }

  /**
   * A text box or shape: its run as read, with the text of the text boxes edited in the editor
   * written into it (docx/shapes.ts rewriteTextBoxes). An untouched one is its XML as read.
   */
  shapeRun(node: PMNode): string {
    const schema = node.type.schema;
    const shape = node.attrs.shape as ShapeModel;
    return rewriteTextBoxes(changedShapeXml(node), shape, (f) => this.blocks(childBlocks(schema.nodeFromJSON(f.doc))));
  }

  /**
   * A picture watermark made in the editor: each header part gets its own relationship to one
   * media file shared by all of them. One whose picture can't be embedded is
   * left out rather than written pointing at nothing.
   */
  private watermarkPicture(node: PMNode): string {
    const rid = node.attrs.src ? this.pkg.image(node.attrs.src, true) : null;
    return rid ? node.attrs.xml.split(`r:id="${WATERMARK_IMAGE_REF}"`).join(`r:id="${rid}"`) : '';
  }

  /** Whether a field read from the file was changed (its instruction or formatting). */
  private fieldEdited(node: PMNode): boolean {
    const kept = keptField(node.attrs.xml);
    if (kept.instr !== node.attrs.instr) return true;
    const run = node.marks.find((m) => m.type.name === 'run');
    return writeRPr(run?.attrs.rPr ?? null, runModel(node.marks)) !== kept.rPr;
  }

  private image(node: PMNode): string {
    const a = node.attrs;
    if (a.xml) {
      // A picture from the file: patch its drawing so wrapping, effects and ids survive.
      const patch = imagePatch(node);
      if (patch.embed !== undefined) {
        const rid = this.pkg.image(a.src);
        if (rid) patch.embed = rid;
        else delete patch.embed; // not embeddable: keep the old picture
      } else if (a.src !== a.origSrc && /^data:/i.test(a.src)) {
        // Replaced by data that can't be embedded (see pictureReplaced): the old picture stays.
        this.pkg.image(a.src); // records the warning
      }
      const xml = Object.keys(patch).length ? patchDrawing(a.xml, patch) : a.xml;
      return `${this.runOpen(node.marks)}${this.rPr(node.marks)}${xml}</w:r>`;
    }
    const rid = this.pkg.image(node.attrs.src);
    if (!rid) return '';
    const w = pxToEmu(node.attrs.width || 200);
    const h = pxToEmu(node.attrs.height || 150);
    const id = this.pkg.pool.drawingId++;
    const name = `Picture ${id}`;
    return (
      `${this.runOpen(node.marks)}${this.rPr(node.marks)}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${w}" cy="${h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="${name}" descr="${escapeXml(node.attrs.alt || '')}"/>` +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    );
  }

  private table(node: PMNode): string {
    const map = TableMap.get(node);
    const layout = tableCells(node, map);
    const { rects, cells } = layout;
    const page = this.model.page;
    const contentWidth = page.width - page.marginLeft - page.marginRight;
    // Column widths (twips). A column keeps its original width unless the user resized it.
    const { widths: grid, fromGrid: fromOriginal } = columnWidths(node, map, contentWidth, layout);
    const gridUnchanged = fromOriginal.every(Boolean);

    const tblPr = (node.attrs.tblPr as string | null) || defaultTblPr(node.attrs.styleId);
    const gridXml = gridUnchanged && node.attrs.tblGridXml
      ? node.attrs.tblGridXml
      : `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;

    let rows = '';
    for (let r = 0; r < map.height; r++) {
      const row = node.child(r);
      let out = '';
      for (let c = 0; c < map.width; ) {
        const pos = map.map[r * map.width + c];
        const cell = cells.get(pos)!;
        const rect = rects.get(pos)!;
        if (rect.left !== c) {
          c++;
          continue; // defensive: should not happen in a normalized table
        }
        c = rect.right;
        const starts = rect.top === r;
        const span = rect.right - rect.left;
        // Cells we added to square up ragged rows are not written unless the user used them.
        if (starts && cell.attrs.filler && isEmptyCell(cell)) continue;
        // What the kept w:tcPr says (parsed once per distinct w:tcPr, not once per cell).
        const kept = cellInfo(cell.attrs.tcPr ?? null);
        // The merged cell still covers the columns it had in the file (not widened / merged sideways).
        const spanUnchanged = kept.span === span;
        if (!starts) {
          const piece = (cell.attrs.mergedTc as string[] | null)?.[r - rect.top - 1];
          const rowspanUnchanged = (cell.attrs.mergedTc as string[] | null)?.length === rect.bottom - rect.top - 1;
          if (piece && rowspanUnchanged && spanUnchanged && tcGridSpan(piece) === span) {
            // Shading, borders and alignment set on the merged cell apply to all its pieces.
            const was = kept.model;
            out += patchMergedCell(piece, was, cellLooks(cell, was.width), span);
            continue;
          }
        }
        const unchangedWidth = spanUnchanged && fromOriginal.slice(rect.left, rect.right).every(Boolean);
        const have = kept.model;
        // Keep the original width (and its unit: dxa, pct, auto) unless the column was resized.
        const width = unchangedWidth && cell.attrs.tcPr ? have.width : grid.slice(rect.left, rect.right).reduce((s, w) => s + w, 0);
        const tcPr = writeTcPr(
          starts ? cell.attrs.tcPr : null,
          cellLooks(cell, width),
          { gridSpan: span, vMerge: rect.bottom - rect.top > 1 ? (starts ? 'restart' : 'continue') : null },
        );
        let content = starts ? this.blocks(childBlocks(cell)) : '<w:p/>';
        if (starts && cell.lastChild?.type.name !== 'paragraph') content += '<w:p/>';
        out += `<w:tc>${tcPr}${content}</w:tc>`;
      }
      const trPr = writeTrPr(row.attrs.trPr ?? null, rowModel(row.attrs));
      rows += `<w:tr${row.attrs.trAttrs ?? ''}>${row.attrs.tblPrEx ?? ''}${trPr}${out}</w:tr>`;
    }
    return `<w:tbl>${tblPr}${gridXml}${rows}</w:tbl>`;
  }
}

/** The id shared by the paragraph pieces and page breaks that came from one w:p. */
export function groupId(n: PMNode): string | null {
  return n.type.name === 'paragraph' ? n.attrs.brGroup : n.type.name === 'page_break' ? n.attrs.group : null;
}

/** Same paragraph properties; the section break is only on the last piece. */
function sameParagraphAttrs(a: PMNode, b: PMNode): boolean {
  return JSON.stringify({ ...a.attrs, sectPr: null }) === JSON.stringify({ ...b.attrs, sectPr: null });
}

/**
 * How many nodes from `nodes[i]` on are written as one w:p: paragraph pieces and page breaks
 * that came from it, while they keep its properties. A piece with a section break ends it,
 * as the section break is on the w:p's mark.
 */
export function paragraphGroupSize(nodes: readonly PMNode[], i: number): number {
  const node = nodes[i];
  const id = groupId(node);
  let first = node.type.name === 'paragraph' ? node : null;
  if (!id || first?.attrs.sectPr) return 1;
  let n = 1;
  while (i + n < nodes.length) {
    const next = nodes[i + n];
    if (groupId(next) !== id || next.attrs.wrap !== node.attrs.wrap) break;
    if (next.type.name === 'paragraph') {
      // A piece whose paragraph formatting was changed starts its own w:p.
      if (first && !sameParagraphAttrs(first, next)) break;
      first ??= next;
      n++;
      if (next.attrs.sectPr) break;
    } else n++;
  }
  return n;
}

/** The modeled cell properties held in a table cell's attrs. */
function cellLooks(cell: PMNode, width: number | null): CellModel {
  return { width, background: cell.attrs.background ?? null, vAlign: cell.attrs.vAlign ?? null, borders: cell.attrs.borders ?? null };
}

/** The modeled row properties held in a table row's attrs. */
function rowModel(a: Record<string, any>): RowModel {
  return { height: a.height ?? null, heightRule: a.heightRule ?? null, header: !!a.header, cantSplit: !!a.cantSplit };
}

/** Changes of a picture read from the file: size, alt text, or a replaced picture (embed set by the caller). */
function imagePatch(node: PMNode): DrawingPatch {
  const a = node.attrs;
  const p: DrawingPatch = {};
  if (a.width && a.width !== a.origWidth) p.cx = pxToEmu(a.width);
  if (a.height && a.height !== a.origHeight) p.cy = pxToEmu(a.height);
  if ((a.alt ?? '') !== readDrawing(a.xml).descr) p.descr = a.alt ?? '';
  if (pictureReplaced(node)) p.embed = null;
  return p;
}

/** A picture from the file that the user replaced with one we can embed (data that decodes to a picture). */
const pictureReplaced = (node: PMNode): boolean =>
  node.attrs.src !== node.attrs.origSrc && /^data:/i.test(node.attrs.src) && !('error' in decodeImage(node.attrs.src));

/**
 * A kept w:tc of a vertically merged cell (below its first row), with the looks the user
 * changed on the merged cell (shading, vertical alignment, border sides) applied to it too.
 */
function patchMergedCell(tc: string, before: CellModel, after: CellModel, span: number): string {
  const sides = BORDER_SIDES.filter((s) => !sameBorder(before.borders?.[s], after.borders?.[s]));
  const bg = (before.background ?? null) !== (after.background ?? null);
  const va = (before.vAlign ?? null) !== (after.vAlign ?? null);
  if (!sides.length && !bg && !va) return tc;
  const el = parseKept(tc);
  const tcPrEl = child(el, 'tcPr');
  const raw = tcPrEl ? serializeXml(tcPrEl) : null;
  const have = readCellModel(tcPrEl);
  const borders = { ...(have.borders ?? {}) };
  for (const s of sides) {
    if (after.borders?.[s]) borders[s] = after.borders[s];
    else delete borders[s];
  }
  const want: CellModel = {
    width: have.width,
    background: bg ? after.background : have.background,
    vAlign: va ? after.vAlign : have.vAlign,
    borders: Object.keys(borders).length ? borders : null,
  };
  const next = parseFragment(writeTcPr(raw, want, { gridSpan: span, vMerge: 'continue' }));
  if (tcPrEl) el.replaceChild(el.ownerDocument.importNode(next, true), tcPrEl);
  else el.insertBefore(el.ownerDocument.importNode(next, true), el.firstChild);
  return serializeXml(el);
}

const tcSpans = new Lru<number>(2000);

/** w:gridSpan of a kept w:tc as used for layout (see gridSpanOf; 1 when absent). */
function tcGridSpan(tc: string): number {
  return memo(tcSpans, tc, () => gridSpanOf(child(parseKept(tc), 'tcPr')));
}

function isEmptyCell(cell: PMNode): boolean {
  return cell.childCount === 1 && cell.firstChild!.type.name === 'paragraph' && cell.firstChild!.childCount === 0;
}

function childBlocks(node: PMNode): PMNode[] {
  const out: PMNode[] = [];
  node.forEach((n) => out.push(n));
  return out;
}

/** A w:t (or w:delText) for text; characters XML forbids are dropped ('' when nothing is left). */
function textRun(text: string, tag = 'w:t'): string {
  const t = xmlSafe(text);
  if (!t) return '';
  // Word only writes xml:space="preserve" when the edges have whitespace.
  const preserve = /^\s|\s$/.test(t) ? ' xml:space="preserve"' : '';
  return `<${tag}${preserve}>${escapeXml(t)}</${tag}>`;
}

/**
 * The content of a run for typed text: tabs and line feeds become w:tab / w:br, and, as Word
 * does with pasted text, a vertical tab a line break and a form feed a page break.
 */
function runContent(text: string, tag: 'w:t' | 'w:delText'): string {
  return text
    .split(/([\t\n\u000B\u000C])/)
    .map((p) =>
      p === '\t' ? '<w:tab/>' : p === '\n' || p === '\u000B' ? '<w:br/>' : p === '\u000C' ? '<w:br w:type="page"/>' : p ? textRun(p, tag) : '',
    )
    .join('');
}

/** What a kept w:fldSimple says: its instruction and the w:rPr of its (first) result run. */
const keptFields = new Lru<{ instr: string; rPr: string }>(500);
function keptField(xml: string): { instr: string; rPr: string } {
  return memo(keptFields, xml, () => {
    const el = parseFragment(xml);
    const r = el.getElementsByTagNameNS(NS.w, 'r')[0];
    const rPr = r ? Array.from(r.children).find((c) => c.localName === 'rPr') ?? null : null;
    return Object.freeze({ instr: (el.getAttributeNS(NS.w, 'instr') ?? '').trim(), rPr: rPr ? serializeXml(rPr) : '' });
  });
}

/**
 * What Word writes inside a w:del for an inline node (tracked deletion, editor/trackChanges.ts):
 * its run as it would be saved, with w:delText for w:t and w:delInstrText for w:instrText.
 * Null for what can't be written that way (a live field, a picture not from the file ...).
 */
export function deletedRunXml(node: PMNode): string | null {
  const run = node.marks.find((m) => m.type.name === 'run');
  const open = `<w:r${run?.attrs.attrs ?? ''}>`;
  const rPr = () => writeRPr(run?.attrs.rPr ?? null, runModel(node.marks));
  switch (node.type.name) {
    case 'text':
      return `${open}${rPr()}${runContent(node.text!, 'w:delText')}</w:r>`;
    case 'tab':
      return `${open}${rPr()}<w:tab/></w:r>`;
    case 'hard_break': {
      const { type, clear } = node.attrs;
      const br = type === 'cr'
        ? '<w:cr/>'
        : `<w:br${type ? ` w:type="${escapeXml(type)}"` : ''}${clear ? ` w:clear="${escapeXml(clear)}"` : ''}/>`;
      return `${open}${rPr()}${br}</w:r>`;
    }
    case 'image': {
      if (!node.attrs.xml || pictureReplaced(node)) return null;
      const patch = imagePatch(node);
      delete patch.embed;
      const xml = Object.keys(patch).length ? patchDrawing(node.attrs.xml, patch) : node.attrs.xml;
      return `${open}${rPr()}${xml}</w:r>`;
    }
    case 'raw_inline': {
      const xml = shapeXmlOutsideSave(node);
      if (!/^<(?:[\w.-]+:)?r[\s>/]/.test(xml)) return null;
      // Only the run's own text (not text inside a text box it holds).
      const r = parseFragment(xml);
      const rename: Record<string, string> = { t: 'delText', instrText: 'delInstrText' };
      for (const c of Array.from(r.children)) {
        if (c.namespaceURI !== NS.w || !rename[c.localName]) continue;
        const repl = r.ownerDocument.createElementNS(NS.w, (c.prefix ? c.prefix + ':' : '') + rename[c.localName]);
        for (const a of Array.from(c.attributes)) repl.setAttributeNode(a.cloneNode() as Attr);
        while (c.firstChild) repl.appendChild(c.firstChild);
        r.replaceChild(repl, c);
      }
      return serializeXml(r);
    }
    default:
      return null;
  }
}

/** A shape's run as it would be saved now (moves, sizes, colours, edited text), outside a save. */
export function currentShapeXml(node: PMNode): string {
  return shapeXmlOutsideSave(node);
}

/** A shape's run with its moves, sizes, colours ... written in (docx/shapeWrite.ts); as read when unchanged. */
function changedShapeXml(node: PMNode): string {
  const shape = node.attrs.shape as ShapeModel | null;
  return hasChangedShape(shape) ? patchShapeXml(node.attrs.xml, JSON.parse(shape!.base!) as ShapeModel, shape!) : node.attrs.xml;
}

/**
 * A raw_inline's XML outside a save (a tracked deletion of it): a shape's edited text boxes are
 * written in when that needs no new relationship (a link, a new picture); otherwise, and for
 * anything else, the XML as read.
 */
function shapeXmlOutsideSave(node: PMNode): string {
  const xml = changedShapeXml(node);
  const shape = node.attrs.shape as ShapeModel | null;
  if (!hasEditedText(shape)) return xml;
  const schema = node.type.schema;
  const docs = textFrames(shape!).filter((f) => f.edited).map((f) => schema.nodeFromJSON(f.doc));
  let needsRel = false;
  for (const d of docs) {
    d.descendants((n) => {
      if (n.marks.some((m) => m.type.name === 'link') || (n.type.name === 'image' && (!n.attrs.xml || pictureReplaced(n)))) needsRel = true;
      return !needsRel;
    });
  }
  if (needsRel) return xml;
  const pool = new MediaPool(new JSZip(), []);
  const writer = new BodyWriter(new Package(new Set(), 'word/document.xml', pool), { page: A4 } as DocxModel);
  return writer.shapeRun(node);
}

/** The modeled paragraph properties held in a paragraph node's attrs. */
export function paraModel(a: Record<string, any>): ParaModel {
  return {
    styleId: a.styleId ?? null,
    align: a.align ?? null,
    indLeft: a.indLeft ?? null,
    indRight: a.indRight ?? null,
    indFirst: a.indFirst ?? null,
    indLeftChars: a.indLeftChars ?? null,
    indRightChars: a.indRightChars ?? null,
    indFirstChars: a.indFirstChars ?? null,
    spaceBefore: a.spaceBefore ?? null,
    spaceAfter: a.spaceAfter ?? null,
    line: a.line ?? null,
    lineRule: a.lineRule ?? null,
    numId: a.numId ?? null,
    ilvl: a.ilvl ?? 0,
    pageBreakBefore: !!a.pageBreakBefore,
  };
}

/** The modeled run properties expressed by an inline node's marks. */
export function runModel(marks: readonly Mark[]): RunModel {
  const has = (n: string) => marks.find((m) => m.type.name === n);
  const font = has('font');
  return {
    charStyle: has('charStyle')?.attrs.id ?? null,
    bold: has('bold') ? has('bold')!.attrs.on !== false : null,
    italic: has('italic') ? has('italic')!.attrs.on !== false : null,
    underline: !!has('underline'),
    strike: !!has('strike'),
    superscript: !!has('superscript'),
    subscript: !!has('subscript'),
    color: has('color')?.attrs.color ?? null,
    highlight: has('highlight')?.attrs.color ?? null,
    fontFamily: font?.attrs.family ?? null,
    fontEastAsia: font?.attrs.eastAsia ?? null,
    fontSize: has('fontSize')?.attrs.pt ?? null,
  };
}

function sectPrXml(model: DocxModel): string {
  if (model.sectPr) return model.sectPr;
  const p = model.page;
  return (
    `<w:sectPr xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:pgSz w:w="${p.width}" w:h="${p.height}"/>` +
    `<w:pgMar w:top="${p.marginTop}" w:right="${p.marginRight}" w:bottom="${p.marginBottom}" w:left="${p.marginLeft}" ` +
    `w:header="${p.header}" w:footer="${p.footer}" w:gutter="0"/></w:sectPr>`
  );
}

// ---------------------------------------------------------------------------
// Package assembly

function relsXml(rels: { id: string; type: string; target: string; external: boolean }[]): string {
  return (
    XML_HEAD +
    `<Relationships xmlns="${NS.rel}">` +
    rels
      .map(
        (r) =>
          `<Relationship Id="${escapeXml(r.id)}" Type="${escapeXml(r.type)}" Target="${escapeXml(r.target)}"` +
          (r.external ? ' TargetMode="External"/>' : '/>'),
      )
      .join('') +
    '</Relationships>'
  );
}

/** Express a package path relative to the directory of `fromPart`. */
function relativeTarget(fromPart: string, path: string): string {
  const from = fromPart.split('/').slice(0, -1);
  const to = path.split('/');
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...Array(from.length - i).fill('..'), ...to.slice(i)].join('/');
}

/** Content types of the parts a package may have, by the type of the relationship that points at them. */
const PART_CONTENT_TYPES: Record<string, string> = {
  [REL_TYPE.officeDocument]: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  [REL_TYPE.styles]: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  [REL_TYPE.settings]: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  [REL_TYPE.numbering]: CONTENT_TYPE.numbering,
  [REL_TYPE.header]: CONTENT_TYPE.header,
  [REL_TYPE.footer]: CONTENT_TYPE.footer,
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes': 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes': 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments': 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable': 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings': 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme': 'application/vnd.openxmlformats-officedocument.theme+xml',
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties': 'application/vnd.openxmlformats-package.core-properties+xml',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties': 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
};

/**
 * The [Content_Types].xml to start from when the package has none (or one that is not XML):
 * the new-document template's, without overrides for parts this package doesn't have, plus
 * the known parts the package's relationships point at and its pictures' extensions.
 */
async function baseContentTypes(zip: JSZip, main: string, broken: string[]): Promise<Document> {
  const doc = parseXml(await blankPackage().file('[Content_Types].xml')!.async('string'));
  const root = doc.documentElement;
  for (const o of Array.from(root.getElementsByTagNameNS(NS.ct, 'Override'))) {
    if (!zip.file((o.getAttribute('PartName') ?? '').replace(/^\//, ''))) root.removeChild(o);
  }
  const listed = new Set(Array.from(root.getElementsByTagNameNS(NS.ct, 'Override')).map((o) => o.getAttribute('PartName')));
  const add = (part: string, type: string) => {
    if (listed.has('/' + part) || !zip.file(part)) return;
    listed.add('/' + part);
    const el = doc.createElementNS(NS.ct, 'Override');
    el.setAttribute('PartName', '/' + part);
    el.setAttribute('ContentType', type);
    root.appendChild(el);
  };
  for (const part of ['', main]) {
    for (const r of (await readRels(zip, part, broken)).values()) {
      if (!r.external && PART_CONTENT_TYPES[r.type]) add(r.target, PART_CONTENT_TYPES[r.type]);
    }
  }
  return doc;
}

async function ensureContentTypes(
  zip: JSZip,
  extensions: Set<string>,
  overrides: Record<string, string>,
  main: string,
  warnings: WriteWarning[],
): Promise<void> {
  const f = zip.file('[Content_Types].xml');
  let doc: Document | null = null;
  try {
    doc = f ? parseXml(await f.async('string')) : null;
  } catch {
    warnings.push({ kind: 'part', part: '[Content_Types].xml', message: 'The package’s content types were not readable XML and were written again' });
  }
  if (!doc) {
    doc = await baseContentTypes(zip, main, []);
    // Every picture in the package needs its extension listed.
    for (const name of Object.keys(zip.files)) {
      const ext = /\/media\/[^/]+\.([A-Za-z0-9]+)$/.exec(name)?.[1]?.toLowerCase();
      if (ext && Object.values(IMAGE_EXT).concat('jpg').includes(ext)) extensions.add(ext);
    }
  }
  const root = doc.documentElement;
  const defaults = new Set(
    Array.from(root.getElementsByTagNameNS(NS.ct, 'Default')).map((d) => (d.getAttribute('Extension') || '').toLowerCase()),
  );
  const mimeFor = (ext: string) =>
    // Only pictures are added here (see imageExt).
    Object.entries(IMAGE_EXT).find(([, e]) => e === ext)?.[0] ?? (ext === 'jpg' ? 'image/jpeg' : `image/${ext}`);
  for (const ext of extensions) {
    if (defaults.has(ext)) continue;
    const el = doc.createElementNS(NS.ct, 'Default');
    el.setAttribute('Extension', ext);
    el.setAttribute('ContentType', mimeFor(ext));
    root.insertBefore(el, root.firstChild);
  }
  const existing = new Set(
    Array.from(root.getElementsByTagNameNS(NS.ct, 'Override')).map((o) => o.getAttribute('PartName')),
  );
  for (const [part, type] of Object.entries(overrides)) {
    if (existing.has(part)) continue;
    const el = doc.createElementNS(NS.ct, 'Override');
    el.setAttribute('PartName', part);
    el.setAttribute('ContentType', type);
    root.appendChild(el);
  }
  zip.file('[Content_Types].xml', XML_HEAD + new XMLSerializer().serializeToString(root));
}

/** Relationship ids referenced (`="rIdN"`) from XML we kept verbatim. */
function rawReferences(doc: PMNode | null, extra: string[], ids: Iterable<string>): Set<string> {
  // Matched by value, since serializers may rename the r: prefix.
  const raw: string[] = [...extra];
  doc?.descendants((n) => {
    let xml: string | null = n.attrs?.xml ?? n.attrs?.sectPr ?? n.attrs?.tblPr ?? null;
    // A replaced picture no longer points at its old media (other references in it stay).
    if (xml && n.type.name === 'image' && pictureReplaced(n)) xml = patchDrawing(xml, { embed: null });
    if (xml) raw.push(xml);
    // Table XML kept as-is: merged cell pieces, cell / row properties, row exceptions.
    if (n.type.name === 'table_cell' || n.type.name === 'table_header') {
      if (n.attrs.tcPr) raw.push(n.attrs.tcPr);
      for (const tc of (n.attrs.mergedTc as string[] | null) ?? []) raw.push(tc);
    } else if (n.type.name === 'table_row') {
      if (n.attrs.trPr) raw.push(n.attrs.trPr);
      if (n.attrs.tblPrEx) raw.push(n.attrs.tblPrEx);
    }
    return true;
  });
  const all = raw.join('');
  return new Set([...ids].filter((id) => all.includes(`="${id}"`)));
}

/**
 * Rewrite one XML part (document, header or footer) from editor content.
 * Images and hyperlinks are regenerated; every other relationship is kept.
 * Returns the media the old version of the part pointed at.
 */
async function rewritePart(
  zip: JSZip,
  part: string,
  doc: PMNode,
  wrap: (inner: string, rootAttrs: string) => string,
  model: DocxModel,
  pool: MediaPool,
  opts: { extraRaw?: string[]; extraRels?: Relationship[]; tail?: string } = {},
): Promise<string[]> {
  const oldRels = await readRels(zip, part);
  const rawRefs = rawReferences(doc, opts.extraRaw ?? [], oldRels.keys());
  const kept: Relationship[] = [];
  const droppedMedia: string[] = [];
  for (const r of oldRels.values()) {
    if (rawRefs.has(r.id)) kept.push(r);
    else if (r.type === REL_TYPE.hyperlink) continue;
    else if (r.type === REL_TYPE.image && !r.external) droppedMedia.push(r.target);
    else kept.push(r);
  }
  for (const r of opts.extraRels ?? []) if (!kept.some((k) => k.id === r.id)) kept.push(r);

  const pkg = new Package(new Set([...oldRels.keys(), ...kept.map((r) => r.id)]), part, pool);
  const writer = new BodyWriter(pkg, model);
  const blocks: PMNode[] = [];
  doc.forEach((n) => blocks.push(n));
  const original = zip.file(part) ? await zip.file(part)!.async('string') : '';
  zip.file(part, wrap(writer.blocks(blocks) + (opts.tail ?? ''), rootAttributes(original || null)));

  const rels = [
    ...kept.map((r) => ({ ...r, target: r.external ? r.target : relativeTarget(part, r.target) })),
    ...pkg.rels,
  ];
  const relsPath = relsPathFor(part);
  if (model.brokenParts?.includes(relsPath)) {
    // A .rels file that was not readable XML is kept as it was, unless the part needs new
    // relationships (a picture or link added): then it has to be written again.
    if (!rels.length) return droppedMedia;
    pool.warnings.push({ kind: 'part', part: relsPath, message: 'An unreadable relationships file was replaced (it now lists only what the saved part uses)' });
  }
  zip.file(relsPath, relsXml(rels));
  return droppedMedia;
}

/** Whether a header/footer has anything worth creating a part for. */
export function hasContent(doc: PMNode): boolean {
  let found = false;
  doc.descendants((n) => {
    if ((n.isText && n.text!.trim()) || (n.isLeaf && !n.isText && n.type.name !== 'hard_break')) found = true;
    return !found;
  });
  return found;
}

const SECT_TAIL = ['textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];

/** Point the section's header/footer references at the given parts and set "different first page". */
export function patchSectPr(
  xml: string,
  refs: { kind: HeaderFooterKind; type: HeaderFooterType; relId: string }[],
  titlePage: boolean,
): string {
  const sect = parseFragment(xml);
  const doc = sect.ownerDocument;
  const wChildren = () => Array.from(sect.children).filter((c) => c.namespaceURI === NS.w);
  for (const ref of refs) {
    const name = ref.kind + 'Reference';
    const existing = wChildren().find(
      (e) => e.localName === name && (e.getAttributeNS(NS.w, 'type') || 'default') === ref.type,
    );
    const el = existing ?? doc.createElementNS(NS.w, 'w:' + name);
    for (const a of Array.from(el.attributes)) if (a.namespaceURI === NS.r) el.removeAttributeNode(a);
    el.setAttributeNS(NS.w, 'w:type', ref.type);
    el.setAttributeNS(NS.r, 'r:id', ref.relId);
    if (!existing) {
      // headerReference* then footerReference* come first in a sectPr.
      const allowed = ref.kind === 'header' ? ['headerReference'] : ['headerReference', 'footerReference'];
      const last = wChildren().filter((c) => allowed.includes(c.localName)).pop();
      sect.insertBefore(el, last ? last.nextSibling : sect.firstChild);
    }
  }
  const titlePg = wChildren().find((c) => c.localName === 'titlePg');
  if (titlePage && !titlePg) {
    const before = wChildren().find((c) => SECT_TAIL.includes(c.localName));
    sect.insertBefore(doc.createElementNS(NS.w, 'w:titlePg'), before ?? null);
  } else if (!titlePage && titlePg) {
    sect.removeChild(titlePg);
  }
  return serializeXml(sect);
}

/**
 * A copy of the package to write into. Going through JSZip's own generate + load (public API)
 * means the parts nobody changes keep their compressed bytes: JSZip passes a loaded entry
 * through unchanged when the compression matches, instead of inflating and deflating it again.
 */
async function copyPackage(src: JSZip): Promise<JSZip> {
  return JSZip.loadAsync(await src.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}

/**
 * Save the document into a .docx package. Parts the editor did not change are copied as they
 * are. Pictures that can't be embedded and parts left alone because they were unreadable are
 * reported in `options.warnings`; they never make the save fail.
 */
export async function writeDocx(doc: PMNode, model: DocxModel, options: WriteOptions = {}): Promise<Uint8Array> {
  const warnings = options.warnings ?? [];
  const broken = model.brokenParts ?? [];
  // Start from a copy of the original package so every part we don't edit survives.
  const src = model.zip ?? blankPackage();
  const zip = await copyPackage(src);

  const pool = new MediaPool(zip, warnings);
  // New pictures get drawing ids after every one already used in the package (cached per
  // part of the opened package, which is never changed).
  pool.drawingId = (await maxDrawingId(src)) + 1;
  const main = await findMainPart(zip, []);
  const mainDir = dirOf(main);
  const mainRels = await readRels(zip, main, []);
  // Ids already given to headers/footers of earlier sections count as used too.
  const usedIds = new Set([...mainRels.keys(), ...model.headerFooters.flatMap((h) => (h.relId ? [h.relId] : []))]);
  const newId = (prefix: string) => {
    let n = 1;
    while (usedIds.has(prefix + n)) n++;
    usedIds.add(prefix + n);
    return prefix + n;
  };
  const extraRels: Relationship[] = [];
  const overrides: Record<string, string> = {};
  const droppedMedia: string[] = [];

  // Headers and footers changed in the editor.
  const refs: { kind: HeaderFooterKind; type: HeaderFooterType; relId: string }[] = [];
  // Relationship ids the sections refer to: a part made in the editor whose reference was
  // undone (e.g. with the section break it belonged to) is not written.
  const referenced = referencedIds(doc, model.sectPr);
  const written = new Set<string>();
  for (const hf of model.headerFooters) {
    if (hf.pending) continue; // made in the editor but never given content
    if (hf.part && !hf.dirty) continue;
    // A part two references share is edited in place, as Word does (the editor keeps both
    // references' content in step), and written once.
    if (hf.part && written.has(hf.part)) continue;
    // Opened but left empty: don't create a part. (One a section already refers to by its
    // relationship id is always written.)
    if (!hf.part && !hf.relId && !hasContent(hf.doc)) continue;
    if (!hf.part && hf.relId && !referenced.has(hf.relId)) continue;
    let part = hf.part;
    let relId = hf.relId;
    if (!part || !relId) {
      let n = 1;
      do part = `${mainDir}${hf.kind}${n++}.xml`;
      while (zip.file(part));
      zip.file(part, ''); // reserve the name
      if (!relId) {
        // Made for the last section outside the editor's own flow: point its w:sectPr at it.
        relId = newId('rIdPxHf');
        refs.push({ kind: hf.kind, type: hf.type, relId });
      } else {
        usedIds.add(relId); // already written into its section's w:sectPr
      }
      extraRels.push({ id: relId, type: REL_TYPE[hf.kind], target: part, external: false });
      overrides['/' + part] = CONTENT_TYPE[hf.kind];
    }
    written.add(part!);
    const tag = hf.kind === 'header' ? 'w:hdr' : 'w:ftr';
    droppedMedia.push(
      ...(await rewritePart(zip, part!, hf.doc, (inner, attrs) => `${XML_HEAD}<${tag} ${attrs}>${inner}</${tag}>`, model, pool)),
    );
  }

  // Numbering: rewrite when the document uses lists.
  let numberingPart = [...mainRels.values()].find((r) => r.type === REL_TYPE.numbering)?.target;
  // Lists made in the editor that nothing uses any more (a paste that was undone ...) are left
  // out; every definition from the opened file is written, used or not.
  const numbering = numberingToWrite(model.numbering, usedNumIds([doc, ...model.headerFooters.map((hf) => hf.doc)]));
  if (numberingPart && broken.includes(numberingPart)) {
    // Unreadable when opened: kept exactly as it was. Lists made in the editor can't go into it.
    if (Object.keys(numbering.nums).length > 0) {
      warnings.push({ kind: 'part', part: numberingPart, message: 'The list definitions were not readable and were kept as they were; new lists were not saved' });
    }
  } else if (Object.keys(numbering.nums).length > 0) {
    if (!numberingPart) {
      numberingPart = mainDir + 'numbering.xml';
      extraRels.push({ id: newId('rIdPxNumbering'), type: REL_TYPE.numbering, target: numberingPart, external: false });
      overrides['/' + numberingPart] = CONTENT_TYPE.numbering;
    }
    const original = zip.file(numberingPart) ? await zip.file(numberingPart)!.async('string') : null;
    zip.file(numberingPart, serializeNumbering(numbering, rootAttributes(original)));
  }

  // Table of contents styles a table of contents made in the editor uses and the file lacks (tocStyles.ts).
  await addTocStyles(zip, mainRels, doc, model, broken);

  // Comments added, replied to, resolved, edited or deleted in the editor (docx/comments.ts);
  // untouched comment parts are left exactly as they are.
  await writeComments(doc.attrs.comments, { zip, main, mainRels, newRelId: newId, extraRels, overrides });
  // Tracked changes turned on or off in the editor (w:trackRevisions); settings.xml is untouched otherwise.
  const settingsPart = [...mainRels.values()].find((r) => r.type === REL_TYPE.settings && !r.external)?.target;
  if (settingsPart && broken.includes(settingsPart)) {
    if (model.trackRevisions) {
      warnings.push({ kind: 'part', part: settingsPart, message: 'The document settings were not readable and were kept as they were; track changes was not saved as on' });
    }
  } else {
    await writeTrackRevisions(model.trackRevisions, { zip, main, mainRels, newRelId: newId, extraRels, overrides });
  }

  const sectPr = patchSectPr(sectPrXml(model), refs, model.titlePage);
  droppedMedia.push(
    ...(await rewritePart(
      zip,
      main,
      doc,
      (inner, attrs) => `${XML_HEAD}<w:document ${attrs}><w:body>${inner}</w:body></w:document>`,
      model,
      pool,
      { extraRaw: [sectPr], extraRels, tail: sectPr },
    )),
  );

  // Write new media, then delete old media nothing points at any more.
  for (const [path, bytes] of pool.files) zip.file(path, bytes);
  const stillUsed = new Set<string>();
  const unreadableRels: string[] = [];
  for (const name of Object.keys(zip.files)) {
    const m = /^(.*)_rels\/([^/]+)\.rels$/.exec(name);
    if (!m) continue;
    for (const r of (await readRels(zip, m[1] + m[2], unreadableRels)).values()) stillUsed.add(r.target);
  }
  // A .rels file that can't be read might point at any of them: then none is deleted.
  if (!unreadableRels.length) for (const media of droppedMedia) if (!stillUsed.has(media)) zip.remove(media);

  // The file's title, only when asked for and the file has no real one (GOV-FINDING-012).
  const title = options.title?.trim();
  if (title) await writeCoreTitle(zip, title, overrides);

  const exts = new Set([...pool.files.keys()].map((p) => p.split('.').pop()!.toLowerCase()));
  await ensureContentTypes(zip, exts, overrides, main, warnings);

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}
