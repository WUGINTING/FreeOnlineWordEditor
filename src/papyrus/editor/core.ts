import { EditorState, NodeSelection, TextSelection, type Command, Plugin, type Transaction } from 'prosemirror-state';
import { EditorView, type NodeView } from 'prosemirror-view';
import { DOMSerializer, Fragment, Slice, type Node as PMNode } from 'prosemirror-model';
import { history, undo, redo, undoDepth, redoDepth, closeHistory } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { baseKeymap, chainCommands } from 'prosemirror-commands';
import { dropCursor } from 'prosemirror-dropcursor';
import { gapCursor } from 'prosemirror-gapcursor';
import { columnResizing, tableEditing, goToNextCell, fixTables, isInTable } from 'prosemirror-tables';
import {
  cssIdent, schema, fieldKind, pasteSlotIndex, pasteSlotSrc, registerImageSources, releaseImageSources, replaceUnsafeImages, setRunExtraCss, unsafeImageCount,
} from './schema';
import { adoptPastedLists, listClipboardSerializer } from './pasteLists';
import { inlineExcelStyles } from './pasteExcel';
import { runLanguages } from './runLang';
import { runExtraCss } from '../docx/styles';
import {
  clearFormatting, effectiveAlign, growFont, indent, inList, insertColumnBreak, insertHardBreak, insertImage, insertPageBreak, insertTab, tabNewRow,
  selectionHas, selectionValue, setAlign, setLink, splitParagraph, toggle, toggleFormat,
} from './commands';
import { hasExplicitSafeScheme } from '../docx/links';
import { listMarkers, madeListsSync } from './listMarkers';
import { addCounts, countDocument, countText, type WordCount } from './wordCount';
import { bookmarkPositions, documentOutline, pageReferences, updatePageReferences, type Heading } from './outline';
import { followLink, linkClick } from './linkFollow';
import { addHeadingBookmarks, blocksFromXml, newTocXml, planEntries, rebuildTocs, tocFields, type TocContext } from './toc';
import { tabStops } from './tabStops';
import { docGrid } from './docGrid';
import { columnGeometry, drawColumnRules } from './columnLayout';
import { hasMacroParts, hasMacroTemplate } from '../docx/macros';
import { documentFields, fieldPlaceholders, lockedFields, nextUnfilled, selectField, type DocField } from './fields';
import { tableStyleClasses } from './tableStyles';
import { splitTable as splitTableCommand } from './tableCommands';
import { autoFitTable, type AutoFitMode } from './tableAutoFit';
import { columnResize } from './columnResize';
import { AREA_LABEL, searchPlugin, type SearchPart } from './search';
import { attentionPlugin } from './assistant';
import { completionPlugin, suggestionPlugin } from './assistantInline';
import {
  afterParagraphSplit, commentTarget, commentsOf, insertComment, insertReply, removeComment, review, setCommentDone, setCommentText,
  type NewComment,
} from './review';
import {
  DEFAULT_AUTHOR, authorInitials, commentDate, readCommentStyles, readComments, type CommentAuthor, type CommentStyles, type DocComment,
} from '../docx/comments';
import { splitParagraphPPr } from '../docx/revisions';
import { NO_TRACK, RevisionIds, trackChanges, trackTransaction } from './trackChanges';
import { FormatPainter, type FormatPainterMode } from './formatPainter';
import {
  SHOW_MARKS_CLASS, formattingMarks, isShowMarksKey, markStaticContent, readShowMarks, recheckFormattingMarks, showFormattingMarks, writeShowMarks,
} from './formattingMarks';
import { eastAsiaFontCss, eastAsiaFontsInCss, onEastAsiaFonts, useEastAsiaFont } from '../docx/eastAsiaFonts';
import { NOT_AN_IMAGE, UNSUPPORTED_IMAGE, dataUrlBytes, imageNodeView, isImageFile, readImageFile } from './imageView';
import { columnBase, pagesKey, pagination, repaginate, sectionColumns, type Layout, type PageBox, type PageGeometry } from './pagination';
import { sectionMarks } from './sectionMarks';
import { WORD_ONLY, frameDoc, placeShapes, shapeLayout } from './shapeView';
import { ShapeTextSession, editableTextBoxes, findShape, shapeKeys, textBoxAt } from './shapeEdit';
import { ShapeTools } from './shapeInteract';
import { textFrames, withText } from '../docx/shapes';
import { readDocx } from '../docx/reader';
import { hasContent, writeDocx, type WriteWarning } from '../docx/writer';
import { blankPackage } from '../docx/template';
import {
  type DocxModel, type HeaderFooterKind, type HeaderFooterPart, type HeaderFooterType, type PageSetup,
} from '../docx/model';
import {
  documentSections, lastSectionXml, readSectPr, sectionBreakXml, sectionHeaderFooter, sectionOfBlock, startingAtBreak,
  variantFor, withColumns, withPageNumbering, withSectionStart, withPageSetup, withReference, withTitlePage, withoutReference, type PageNumbering, type Section,
  type SectionColumns, type SectionStart,
} from '../docx/sections';
import { fieldNumberFormat, formatPageNumber } from '../docx/pageNumbers';
import { isListKind } from '../docx/numbering';
import { missingFonts, type MissingFont } from '../docx/fonts';
import { printHtml, printPages } from './print';
import {
  docWatermark, keepWatermarks, keptXml, measureWatermarkText, replaceWatermarkTr, sectionHeaders, usedShapeIds, watermarkLayer,
  watermarkNodes, watermarkTargets, withoutWatermarks,
} from './watermark';
import { keepSectionReferences } from './sectionReferences';
import {
  WATERMARK_Z_INDEX, watermarkRunXml, watermarkSettings, watermarkShapetype, type Watermark, type WatermarkArea,
} from '../docx/watermark';
import { twipsToPx } from '../units';
import './editor.css';

export const PAGE_GAP = 24;

/** How the 目錄 styles docx/tocStyles.ts adds look, for the page before the file has them. */
/** A section, as the one-column layout of its columns remembers it: its place and its w:sectPr. */
const singleColumnKey = (s: Section) => `${s.index}|${s.xml ?? ''}`;

const TOC_FALLBACK_CSS: [string, string][] = [
  ['TOC2', 'margin-left:2em'],
  ['TOC3', 'margin-left:4em'],
  ['TOCHeading', 'text-align:center;font-weight:bold;margin-top:12pt;margin-bottom:6pt'],
];

/** A document this big (bytes, see DocxEditor.estimatedSize) is worth a warning: the server takes 50 MB. */
export const LARGE_DOCUMENT_BYTES = 30 * 1024 * 1024;

let editorCount = 0;

setRunExtraCss(runExtraCss);

/** Where the cursor is: the body, a header or footer, or a text box (of either). */
export type EditTarget = 'body' | HeaderFooterKind | 'textbox';

/** What a toolbar needs to know about the current selection. */
export interface EditorSnapshot {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  superscript: boolean;
  subscript: boolean;
  link: string | null;
  color: string | null;
  highlight: string | null;
  fontFamily: string | null;
  fontSize: number | null;
  /** The selection has more than one font size / colour / font: no single value is shown. */
  mixed: { fontSize: boolean; color: boolean; fontFamily: boolean };
  styleId: string | null;
  align: string | null;
  list: 'bullet' | 'decimal' | 'gongwen' | null;
  inTable: boolean;
  canUndo: boolean;
  canRedo: boolean;
  pageCount: number;
  /** Page number (1-based position among the pages) where the cursor is. */
  currentPage: number;
  /** Zoom factor of the page view (1 = 100 %). */
  zoom: number;
  /** Sections in the document, and the one the cursor is in (0-based). */
  sectionCount: number;
  section: number;
  /** The cursor's paragraph ends a section (a section break can be removed there). */
  atSectionBreak: boolean;
  /** Which part of the document the cursor is in. */
  target: EditTarget;
  titlePage: boolean;
  /** 追蹤修訂 is on: edits are recorded as tracked changes. */
  trackChanges: boolean;
  /** 複製格式 is painting with the mouse: once, or until stopped (see startFormatPainter). */
  formatPainter: FormatPainterMode | null;
  /** 顯示／隱藏編輯標記 is on: the non-printing formatting marks are shown (see setShowMarks). */
  showMarks: boolean;
}

/** Where page setup changes apply. */
export type PageSetupScope = 'section' | 'all';

/** New text columns for a section (see DocxEditor.setColumns): only the given values change. */
export type ColumnsChange = Partial<Omit<SectionColumns, 'equalWidth'>>;

export interface DocxEditorOptions {
  /** Called after every state change with toolbar-relevant info. */
  onUpdate?: (snapshot: EditorSnapshot) => void;
  /** Called when the document content changes. */
  onChange?: () => void;
  editable?: boolean;
  /** Read-only, but comments may still be added, replied to and resolved (the text stays as it is). */
  commenting?: boolean;
  /** The browser's spell check underlines (off by default; see setSpellcheck). */
  spellcheck?: boolean;
  /** A short message for the user, e.g. why an edit was refused (a locked field). */
  onNotice?: (message: string) => void;
  /** Asks the user to confirm something that cannot be undone (default: the browser's confirm). */
  confirm?: (message: string) => boolean;
  /** Who writes new comments (the signed-in user); 「使用者」 when not given. See setAuthor. */
  author?: CommentAuthor | null;
  /**
   * The document's name (e.g. the online document's title). On save it becomes the file's
   * title (docProps/core.xml dc:title) when the file has none or only Word's generic
   * "Word Document"; a title the author set is never changed. See setDocumentTitle, save.
   */
  documentTitle?: string | null;
}

const VARIANT_LABEL: Record<HeaderFooterType, string> = {
  default: '',
  first: '（第一頁）',
  even: '（偶數頁）',
};

/** The document as it was when opened or last saved; see DocxEditor.isModified. */
export interface SavePoint {
  body: PMNode;
  headerFooters: Map<HeaderFooterPart, PMNode>;
  titlePage: boolean;
  /** The last section's page setup (kept on the model, not in the body). */
  sectPr: string | null;
  page: string;
  trackRevisions: boolean;
}

/** A static header/footer drawn on a page (see renderPages). */
interface HeaderFooterBox {
  el: HTMLElement;
  /** The header/footer document it shows. */
  doc: PMNode;
  /** Its position on the page (style). */
  style: string;
  /** The field values it shows (PAGE, NUMPAGES, SECTIONPAGES). */
  fields: string;
  /** It shows the formatting marks (see DocxEditor.setShowMarks). */
  marks: boolean;
}

/** A page drawn behind the body (see renderPages). */
interface DrawnPage {
  el: HTMLElement;
  geometry: string;
  boxes: Partial<Record<HeaderFooterKind, HeaderFooterBox>>;
  /**
   * The watermark drawn on the page (浮水印), with what it was drawn from; `measured`: its text
   * was drawn by its font's measurements (not yet when the font was still loading).
   */
  watermark?: { key: string; src: string | null; el: HTMLElement; measured: boolean };
}

/** Fields a static header/footer shows live values of (see DocxEditor.fieldValue). */
const LIVE_FIELDS = ['PAGE', 'NUMPAGES', 'SECTIONPAGES'];

/**
 * Header/footer content, serialized once per header/footer document and copied onto each page;
 * with the formatting marks (顯示／隱藏編輯標記) it is a second copy that has them.
 */
const hfTemplates = new WeakMap<PMNode, { content: DocumentFragment; fields: boolean }>();
const hfMarkedTemplates = new WeakMap<PMNode, { content: DocumentFragment; fields: boolean }>();

function hfTemplate(doc: PMNode, marks = false): { content: DocumentFragment; fields: boolean } {
  const cache = marks ? hfMarkedTemplates : hfTemplates;
  let t = cache.get(doc);
  if (!t) {
    const content = DOMSerializer.fromSchema(schema).serializeFragment(doc.content) as DocumentFragment;
    if (marks) markStaticContent(content);
    t = { content, fields: !!content.querySelector('[data-field]') };
    cache.set(doc, t);
  }
  return t;
}

/** The pictures of an opened document (body, headers and footers). */
function imageSources(doc: PMNode, model: DocxModel): string[] {
  const out: string[] = [];
  for (const d of [doc, ...model.headerFooters.map((h) => h.doc)]) {
    d.descendants((node) => {
      if (node.type === schema.nodes.image && node.attrs.src) out.push(node.attrs.src as string);
      return true;
    });
  }
  return out;
}

interface HeaderFooterSession {
  hf: HeaderFooterPart;
  view: EditorView;
  pageIndex: number;
  box: HTMLElement;
}

/** Page geometry in px of a section, before headers/footers push the text area. */
function geometryOf(p: PageSetup) {
  return {
    width: twipsToPx(p.width),
    height: twipsToPx(p.height),
    textTop: twipsToPx(p.marginTop),
    textBottom: twipsToPx(p.marginBottom),
    marginLeft: twipsToPx(p.marginLeft),
    marginRight: twipsToPx(p.marginRight),
  };
}

/** Shows a field's live value (page number ...) instead of its cached text. */
class FieldView implements NodeView {
  dom: HTMLElement;
  constructor(
    private node: PMNode,
    private value: (kind: string, format: string | null) => string | null,
    private registry: Set<FieldView>,
  ) {
    this.dom = document.createElement('span');
    this.dom.className = 'dx-field';
    registry.add(this);
    this.render();
  }
  render() {
    const kind = fieldKind(this.node.attrs.instr);
    this.dom.dataset.field = kind;
    this.dom.textContent = this.value(kind, fieldNumberFormat(this.node.attrs.instr)) ?? (this.node.attrs.text || ' ');
  }
  update(node: PMNode) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }
  ignoreMutation() {
    return true;
  }
  destroy() {
    this.registry.delete(this);
  }
}

export class DocxEditor {
  /** The body editor. */
  view: EditorView | null = null;
  model!: DocxModel;
  pageCount = 1;

  private root: HTMLElement;
  /** Escape was the last key in the body: the next Tab / Shift+Tab moves the focus out. */
  private escapePressed = false;
  private canvas: HTMLElement;
  private pagesLayer: HTMLElement;
  /** Over the body: covers where a table row breaks across pages (PageBox.cuts). */
  private cutsLayer: HTMLElement;
  private host: HTMLElement;
  private hfLayer: HTMLElement;
  /** For screen readers: the headers and footers' text (the pages showing them are hidden from them). */
  private hfSummary: HTMLElement;
  private styleEl: HTMLStyleElement;
  /** The document's own CSS (styles.xml), scoped to this editor. */
  private docCss = '';
  /** Stops listening for East Asian fonts coming into use (docx/eastAsiaFonts.ts). */
  private stopFontCss = onEastAsiaFonts(() => this.syncFontCss());
  private hfSession: HeaderFooterSession | null = null;
  /** A text box being edited (editor/shapeEdit.ts), in the body or in the header/footer being edited. */
  private shapeSession: ShapeTextSession | null = null;
  /** Selecting, moving, resizing and inserting shapes (editor/shapeInteract.ts); the 圖形格式 tab uses it. */
  readonly shapes: ShapeTools;
  /** Said to screen readers: what a shape command did (moved to ..., selected ...). */
  private live: HTMLElement;
  /** Editing state (with undo history) of each header/footer, kept while the document is open. */
  private hfStates = new WeakMap<HeaderFooterPart, EditorState>();
  private hfRenderTimer: ReturnType<typeof setTimeout> | undefined;
  /** Live field views in the header/footer editor, refreshed when the page count changes. */
  private fieldViews = new Set<FieldView>();
  /** What "unchanged" means: the content when the document was opened or last saved. */
  private baseline: SavePoint | null = null;
  /**
   * Per section, how far a tall header (footer) pushes the text area down (up), px from the
   * page edge, like Word does. Measured after the headers are drawn.
   */
  private pushed = new Map<number, { top: number; bottom: number }>();
  /** The pages as last laid out. */
  pages: PageBox[] = [];
  /** Holds the zoomed canvas; sized to the canvas at the current zoom. */
  private stage: HTMLElement;
  private zoomValue = 1;
  private fitWidth = false;
  /** The body column's inline style (padding), see bodyPadding. */
  private bodyStyle = '';
  /** The last section's settings as opened, for undoing back to a file without a body w:sectPr. */
  private openedLast: { page: PageSetup; titlePage: boolean } = { page: { width: 0, height: 0, marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0, header: 0, footer: 0 }, titlePage: false };
  /** Class that scopes this editor's document styles. */
  private readonly scope = `dx-e${++editorCount}`;
  /** Set by destroy(): a document still opening is dropped when it arrives. */
  private destroyed = false;
  /** Bumped by every open(): only the latest one shows its document. */
  private loadGen = 0;
  /**
   * The drawn pages, reused from one layout to the next (see renderPages): the page element,
   * its geometry and its static header/footer boxes.
   */
  private pageEls: DrawnPage[] = [];
  /**
   * Height of a static header/footer per (section, kind, variant), measured once until that
   * header/footer, its position or the fonts change.
   */
  private hfHeights = new Map<string, { doc: PMNode; style: string; contentHeight: number; height: number }>();
  /** The page the cursor is on, as last measured (see emitUpdate). */
  private lastCursorPage = 1;
  /**
   * 複製格式 (Format Painter), shared by the body and the header/footer being edited. While it
   * paints, the pages show a paint cursor (the dx-format-painting class).
   */
  private painter = new FormatPainter(() => {
    this.root.classList.toggle('dx-format-painting', !!this.painter.mode);
    this.emitUpdate();
  });
  private pageCheck: { frame?: number; timer?: ReturnType<typeof setTimeout> } | null = null;
  /**
   * 顯示／隱藏編輯標記: the viewer's choice, not the document's (never saved in the file), kept
   * per browser like Word keeps it per installation (see editor/formattingMarks.ts).
   */
  private marksOn = readShowMarks();

  constructor(mount: HTMLElement, private options: DocxEditorOptions = {}) {
    this.root = el('div', `dx-root ${this.scope}`);
    this.styleEl = document.createElement('style');
    this.stage = el('div', 'dx-stage');
    this.canvas = el('div', 'dx-canvas');
    this.pagesLayer = el('div', 'dx-pages');
    this.host = el('div', 'dx-host');
    this.cutsLayer = el('div', 'dx-cuts');
    this.hfLayer = el('div', 'dx-hf-layer');
    this.hfSummary = el('div', 'dx-sr dx-hf-summary');
    this.canvas.append(this.hfSummary, this.pagesLayer, this.host, this.cutsLayer, this.hfLayer);
    // The pages drawn behind the text (their header / footer copies) and the table-row covers
    // are pictures of the document, not more of it: screen readers skip them (persona-300).
    // What the headers and footers say is read once, before the text (hfSummary).
    this.pagesLayer.setAttribute('aria-hidden', 'true');
    this.cutsLayer.setAttribute('aria-hidden', 'true');
    this.stage.append(this.canvas);
    this.root.append(this.styleEl, this.stage);
    mount.replaceChildren(this.root);
    this.canvas.addEventListener('dblclick', (e) => this.onDoubleClick(e));
    this.canvas.addEventListener('click', (e) => this.onLinkClick(e));
    // A middle click on a link: asked about too, never opened by the browser in a new tab.
    this.canvas.addEventListener('auxclick', (e) => this.onLinkClick(e));
    this.live = el('div', 'dx-live');
    this.live.setAttribute('role', 'status');
    this.live.setAttribute('aria-live', 'polite');
    this.root.append(this.live);
    this.shapes = new ShapeTools({
      editable: () => this.options.editable !== false && !!this.view,
      tracking: () => this.trackOn,
      notice: (m) => this.options.onNotice?.(m),
      announce: (m) => this.announce(m),
      editText: (k, i) => this.editShapeText(k, i),
      theme: () => this.model?.theme,
      documents: () => [...(this.view ? [this.view.state.doc] : []), ...(this.model?.headerFooters ?? []).map((h) => h.doc)],
    });
    window.addEventListener('resize', this.onResize);
    document.fonts?.addEventListener?.('loadingdone', this.onFontsLoaded);
    if (!DocxEditor.live.size) document.addEventListener('keydown', DocxEditor.onShowMarksKey);
    DocxEditor.live.add(this);
  }

  /**
   * The editors on the page, in the order they were made. 顯示／隱藏編輯標記 is one browser-wide
   * choice, so they all show it the same (see setShowMarks): the app shows a read-only version
   * over the document (VersionPreview) or two versions side by side (VersionSideBySide).
   */
  private static live = new Set<DocxEditor>();

  /**
   * Word's Ctrl+Shift+8 (顯示／隱藏編輯標記), in the text of an editor (the body or a
   * header/footer) or with nothing focused. One listener for all the editors on the page. Not a
   * ProseMirror key binding: those are not run in read-only mode, where showing the marks is just
   * as useful. A read-only editor takes no focus, so a key from the page body is for the editor
   * the user is looking at: the last one made that is shown (a preview opens over the document).
   */
  private static onShowMarksKey = (e: KeyboardEvent): void => {
    if (e.defaultPrevented || e.repeat || !isShowMarksKey(e)) return;
    const target = e.target as Node | null;
    const editors = [...DocxEditor.live];
    const owner = target ? editors.find((ed) => ed.root.contains(target)) : undefined;
    if (!owner && target !== document.body && target !== document.documentElement) return;
    const shown = (ed: DocxEditor) => ed.root.isConnected && (ed.root as { checkVisibility?: () => boolean }).checkVisibility?.() !== false;
    const seen = owner ?? [...editors].reverse().find(shown) ?? editors[editors.length - 1];
    if (!seen) return;
    e.preventDefault();
    seen.toggleShowMarks();
  };

  /** Late fonts change how tall headers and footers are: measure them again. */
  private onFontsLoaded = () => {
    this.hfHeights.clear();
    // A watermark's text is drawn by its font's measurements once the font is there.
    if (this.view && this.pageEls.some((p) => p.watermark && !p.watermark.measured)) this.renderPages(this.pages);
  };

  private onResize = () => {
    if (!this.fitWidth) return;
    this.applyZoom();
    this.emitUpdate();
  };

  // ----- document lifecycle -----

  async open(data: ArrayBuffer | Uint8Array | Blob): Promise<void> {
    const gen = ++this.loadGen;
    // 複製格式 belongs to the document being left.
    this.painter.reset();
    // Destroyed meanwhile, or a newer document asked for: this one is dropped.
    const stale = () => this.destroyed || gen !== this.loadGen;
    const bytes = data instanceof Blob ? data.size : data.byteLength;
    let { doc, model } = await readDocx(data);
    if (stale()) return;
    // The comments travel on the document, so editing them is undoable and counts as a change.
    const [comments, styles, macroTemplate] = await Promise.all([
      readComments(model.zip).catch((): DocComment[] => []),
      readCommentStyles(model.zip),
      hasMacroTemplate(model.zip),
    ]);
    if (stale()) return;
    this.macroTemplate = macroTemplate;
    doc = schema.nodes.doc.create({ ...doc.attrs, comments }, doc.content);
    this.setDocument(doc, model);
    // What the file has besides its pictures (text, styles, fonts …): see estimatedSize.
    this.otherBytes = Math.max(0, bytes - this.pictureBytes());
    this.commentStyles = styles;
    // Optional parts that were not readable XML: skipped on screen, kept as they are on save.
    const broken = model.brokenParts ?? [];
    if (broken.length) this.options.onNotice?.(brokenPartsNotice(broken));
  }

  async newDocument(): Promise<void> {
    await this.open(await blankPackage().generateAsync({ type: 'uint8array' }));
  }

  /** What the last save could not write as shown (see WriteWarning); the save itself went through. */
  saveWarnings: WriteWarning[] = [];

  /**
   * The document as a .docx. `options.title` (else the documentTitle option) becomes the
   * file's title when it has none or only "Word Document" (GOV-FINDING-012); without either,
   * docProps/core.xml is written exactly as it was.
   */
  async save(options: { title?: string | null } = {}): Promise<Blob> {
    if (!this.view) throw new Error('No document loaded');
    const warnings: WriteWarning[] = [];
    const title = (options.title ?? this.options.documentTitle ?? '').trim() || undefined;
    const bytes = await writeDocx(this.view.state.doc, this.model, { warnings, title });
    this.saveWarnings = warnings;
    if (warnings.length) this.options.onNotice?.(saveWarningsNotice(warnings));
    return new Blob([bytes as BlobPart], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }

  /** The current content, to be passed to markSaved() once saving it has succeeded. */
  savePoint(): SavePoint | null {
    if (!this.view) return null;
    return {
      body: this.view.state.doc,
      headerFooters: new Map(this.model.headerFooters.map((hf) => [hf, hf.doc])),
      titlePage: this.model.titlePage,
      sectPr: this.model.sectPr,
      page: JSON.stringify(this.model.page),
      trackRevisions: this.trackOn,
    };
  }

  /** The document's name, used as the file's title on save (see DocxEditorOptions.documentTitle). */
  setDocumentTitle(title: string | null): void {
    this.options.documentTitle = title;
  }

  /** The content in `point` is now what the server has. */
  markSaved(point: SavePoint | null): void {
    if (point) this.baseline = point;
  }

  /**
   * Whether the document differs from when it was opened or last saved: undoing every edit
   * makes it unchanged again. Body, headers/footers and "different first page" count.
   */
  isModified(): boolean {
    const base = this.baseline;
    if (!this.view || !base) return false;
    if (!this.view.state.doc.eq(base.body) || this.model.titlePage !== base.titlePage) return true;
    if (this.model.sectPr !== base.sectPr || JSON.stringify(this.model.page) !== base.page) return true;
    if (this.trackOn !== base.trackRevisions) return true;
    for (const hf of this.model.headerFooters) {
      if (hf.pending) continue; // opened for an earlier section, still empty
      const was = base.headerFooters.get(hf);
      // A header/footer created in this session counts once it has content.
      if (was ? !hf.doc.eq(was) : hf.doc.textContent.length > 0 || hf.doc.childCount > 1) return true;
    }
    return false;
  }

  destroy(): void {
    this.destroyed = true;
    this.stopFontCss();
    this.painter.stop();
    window.removeEventListener('resize', this.onResize);
    document.fonts?.removeEventListener?.('loadingdone', this.onFontsLoaded);
    DocxEditor.live.delete(this);
    if (!DocxEditor.live.size) document.removeEventListener('keydown', DocxEditor.onShowMarksKey);
    this.cancelPageCheck();
    this.closeShapeText(false);
    this.shapes.cancelPlacing();
    this.closeHeaderFooter();
    this.view?.destroy();
    this.view = null;
    releaseImageSources(this);
    this.root.remove();
  }

  // ----- comments (docx/comments.ts, editor/review.ts) -----

  /** Word's comment styles in the open document ("annotation reference" marks new comments). */
  private commentStyles: CommentStyles = { text: null, reference: null };

  /** The document's comments (replies and resolved state included), as they are now. */
  comments(): DocComment[] {
    return this.view ? commentsOf(this.view.state) : [];
  }

  /** Who writes new comments: the name the host page gave, else 「使用者」. */
  get author(): { name: string; initials: string } {
    const name = this.options.author?.name?.trim() || DEFAULT_AUTHOR;
    return { name, initials: this.options.author?.initials?.trim() || authorInitials(name) };
  }

  /** The signed-in user changed. */
  setAuthor(author: CommentAuthor | null): void {
    this.options.author = author;
  }

  private get canComment(): boolean {
    return (this.options.editable !== false || this.options.commenting === true) && !!this.view;
  }

  /** Comments only, in read-only mode (see DocxEditorOptions.commenting). */
  setCommenting(on: boolean): void {
    this.options.commenting = on;
    this.emitUpdate();
  }

  /** Only a comment's author may edit or delete it (and only while editing is allowed). */
  canChangeComment(id: string): boolean {
    const c = this.canComment ? this.comments().find((x) => x.id === id) : undefined;
    return !!c && c.author === this.author.name;
  }

  /**
   * Where a new comment would go: the body's selection, or the word at the cursor. Null in
   * read-only mode, while a header/footer is being edited (Word has no comments there), or
   * where there is no text.
   */
  commentTarget(): { from: number; to: number } | null {
    if (!this.canComment || this.hfSession || this.shapeSession) return null;
    return commentTarget(this.view!.state);
  }

  private newComment(text: string): NewComment {
    const { name, initials } = this.author;
    const { date, dateUtc } = commentDate();
    return { text, author: name, initials, date, dateUtc, referenceStyle: this.commentStyles.reference };
  }

  /** Add a comment on `target` (default: commentTarget()); returns its id, or null when refused. */
  addComment(text: string, target: { from: number; to: number } | null = this.commentTarget()): string | null {
    if (!this.canComment || this.hfSession || this.shapeSession || !target) return null;
    return insertComment(this.view!.state, this.view!.dispatch, target, this.newComment(text));
  }

  /** Reply to a comment's thread; returns the reply's id. */
  replyComment(id: string, text: string): string | null {
    if (!this.canComment) return null;
    return insertReply(this.view!.state, this.view!.dispatch, id, this.newComment(text));
  }

  /** Resolve (done) or reopen a comment's thread; anyone may, as in Word. */
  resolveComment(id: string, done: boolean): boolean {
    return this.canComment && setCommentDone(this.view!.state, this.view!.dispatch, id, done);
  }

  /** Change the text of one of your own comments. */
  editComment(id: string, text: string): boolean {
    return this.canChangeComment(id) && setCommentText(this.view!.state, this.view!.dispatch, id, text);
  }

  /** Delete one of your own comments (with its replies, when it starts a thread). */
  deleteComment(id: string): boolean {
    return this.canChangeComment(id) && removeComment(this.view!.state, this.view!.dispatch, id);
  }

  // ----- tracked changes (editor/trackChanges.ts) -----

  /** Revision ids given out in this document (body and headers / footers share them). */
  private revisionIds = this.newRevisionIds();

  private newRevisionIds(): RevisionIds {
    return new RevisionIds(() => [this.view?.state.doc, ...(this.model?.headerFooters ?? []).map((h) => h.doc)]);
  }

  private get trackOn(): boolean {
    return this.model?.trackRevisions === true;
  }

  /** Whether this Tab / Shift+Tab comes right after Escape (and so leaves the document). */
  private leaveOnTab(): boolean {
    const leave = this.escapePressed;
    this.escapePressed = false;
    return leave;
  }

  /** Whether 追蹤修訂 is on (from the file's settings when it was opened, or as set since). */
  trackChanges(): boolean {
    return this.trackOn;
  }

  /**
   * Turn 追蹤修訂 on or off, as Word's Review > Track Changes (Ctrl+Shift+E): while on, what is
   * typed, pasted, deleted or formatted is recorded with the author and time. Saved as
   * w:trackRevisions in settings.xml, as Word does. Not in read-only mode.
   */
  setTrackChanges(on: boolean): boolean {
    if (!this.view || this.options.editable === false) return false;
    if (this.trackOn === on) return true;
    this.model.trackRevisions = on;
    // The views read the setting again (the marks new text gets at the cursor).
    for (const v of [this.view, this.hfSession?.view, this.shapeSession?.view]) v?.dispatch(v.state.tr.setSelection(v.state.selection).setMeta('addToHistory', false));
    this.options.onChange?.();
    this.emitUpdate();
    return true;
  }

  /** Section and page setup changes can't be recorded as tracked changes: refused while tracking. */
  private refusedWhileTracking(): boolean {
    if (!this.trackOn) return false;
    this.options.onNotice?.('追蹤修訂開啟時，無法變更分節符號或版面設定，請先關閉「追蹤修訂」。');
    return true;
  }

  /**
   * The editor's <style>: the document's styles, then the East Asian font aliases of the fonts in
   * use and the rule that puts them first in runs with w:hint="eastAsia" (docx/eastAsiaFonts.ts).
   * print.ts copies it into the printout, so the PDF draws those symbols the same way.
   */
  private syncFontCss(): void {
    const css = `${this.docCss}\n${eastAsiaFontCss()}`;
    if (this.styleEl.textContent !== css) this.styleEl.textContent = css;
  }

  // ----- editing API -----

  /** The editor that has the cursor: the body, a header/footer or a text box being edited. */
  get activeView(): EditorView | null {
    return this.shapeSession?.view ?? this.hfSession?.view ?? this.view;
  }

  get target(): EditTarget {
    return this.shapeSession ? 'textbox' : this.hfSession?.hf.kind ?? 'body';
  }

  /**
   * Run a command on the part with the cursor (the body, or the header/footer being edited).
   * The text gets the keyboard afterwards, unless `focus` is false: a dialog that stays open
   * (the 符號 dialog) keeps it.
   */
  run(command: Command, options: { focus?: boolean } = {}): boolean {
    const view = this.activeView;
    if (!view) return false;
    // Another command ends 複製格式, as in Word.
    this.painter.stop();
    // Page and column breaks only make sense in the body.
    if ((this.hfSession || this.shapeSession) && (command === insertPageBreak || command === insertColumnBreak)) return false;
    const ok = command(view.state, view.dispatch, view);
    if (options.focus !== false) view.focus();
    return ok;
  }

  /**
   * Word's 複製格式 button: copy the formatting of the selection (or of the text at the cursor)
   * and paint it onto the next text selected with the mouse; a click paints the word under it.
   * `sticky` (a double-click on the button) keeps painting until Escape or the button again.
   * Works in the body and in a header/footer being edited. Not in read-only mode.
   */
  startFormatPainter(sticky = false): boolean {
    const view = this.activeView;
    if (!view || this.options.editable === false) return false;
    this.painter.start(view, sticky);
    view.focus();
    return true;
  }

  /** Stop 複製格式 painting (as Escape does). */
  stopFormatPainter(): void {
    this.painter.stop();
  }

  /** The 複製格式 button's click: stop painting if it is on, otherwise start (see startFormatPainter). */
  toggleFormatPainter(sticky = false): boolean {
    if (!this.painter.mode) return this.startFormatPainter(sticky);
    this.painter.stop();
    this.activeView?.focus();
    return true;
  }

  /** Ctrl+Shift+C: copy the formatting of the selection for pasteFormat, without painting. A native copy or cut forgets it. */
  copyFormat(): boolean {
    const view = this.activeView;
    return !!view && this.painter.copy(view);
  }

  /** Ctrl+Shift+V: paste what copyFormat copied onto the selection; false when there is none (not the button's). */
  pasteFormat(): boolean {
    const view = this.activeView;
    if (!view || this.options.editable === false) return false;
    return this.painter.paste(view);
  }

  /** 顯示／隱藏編輯標記 is on (see setShowMarks). */
  get showMarks(): boolean {
    return this.marksOn;
  }

  /**
   * Word's 顯示／隱藏編輯標記 (Show/Hide ¶): show or hide the non-printing formatting marks
   * (¶ · ° □ → ↵) in the body, the header/footer being edited and the headers/footers drawn on
   * the pages. Display only: the document, its undo history and the saved file are unchanged,
   * nothing moves on the pages, and the printout never shows them. Works in read-only mode too.
   * The choice is the browser's: every editor on the page follows (a version shown over or next
   * to the document), and it is remembered for the next documents.
   */
  setShowMarks(on: boolean): void {
    let changed = false;
    for (const editor of new Set([this, ...DocxEditor.live])) changed = editor.applyShowMarks(on) || changed;
    if (changed) writeShowMarks(on);
  }

  /** Show or hide the marks in this editor only; whether that changed anything. */
  private applyShowMarks(on: boolean): boolean {
    if (on === this.marksOn) return false;
    this.marksOn = on;
    for (const view of [this.view, this.hfSession?.view]) if (view) showFormattingMarks(view, on);
    // The static header/footer copies on the pages are drawn again with or without them.
    if (this.view) this.renderPages(this.pages);
    this.emitUpdate();
    return true;
  }

  /** The 顯示／隱藏編輯標記 button and Ctrl+Shift+8: show the marks if hidden, else hide them. */
  toggleShowMarks(): boolean {
    this.setShowMarks(!this.marksOn);
    return true;
  }

  can(command: Command): boolean {
    const view = this.activeView;
    return !!view && command(view.state, undefined, view);
  }

  undo(): boolean {
    return this.run(this.historyCommand('undo'));
  }

  redo(): boolean {
    return this.run(this.historyCommand('redo'));
  }

  /**
   * Word's 分割表格 (表格版面配置 › 合併, Ctrl+Shift+Enter in a table) at the cursor, in the body
   * or a header/footer being edited. One undo step; refused while 追蹤修訂 is on and not in
   * read-only mode. See splitTable in tableCommands.ts.
   */
  splitTable(): boolean {
    if (this.options.editable === false) return false;
    return this.run(splitTableCommand);
  }

  /**
   * Word's 自動調整 (表格版面配置 › 儲存格大小) for the table with the cursor, in the body or a
   * header/footer being edited: fit the columns to their content (measured on the page) or to
   * the text area, or fix their current widths. One undo step; refused while 追蹤修訂 is on and
   * not in read-only mode. See autoFitTable in tableAutoFit.ts.
   */
  autoFitTable(mode: AutoFitMode): boolean {
    if (this.options.editable === false) return false;
    const p = this.cursorSection().page;
    return this.run(autoFitTable(mode, { textWidth: p.width - p.marginLeft - p.marginRight }));
  }

  /** Width of the text area in px where the cursor is (page minus left/right margins). */
  get contentWidth(): number {
    const p = this.cursorSection().page;
    return twipsToPx(p.width - p.marginLeft - p.marginRight);
  }

  /**
   * Insert a picture from a file; returns whether it was inserted. Formats Word may not show
   * (WebP, SVG ...) are converted to PNG. A file that is not a picture (a PDF ...) or a picture
   * the browser cannot read is refused with a notice (onNotice), never inserted: documents take
   * no attachments (GOV-FINDING-007).
   */
  async insertImageFile(file: File): Promise<boolean> {
    const image = await this.readPicture(file);
    return !!image && this.run(insertImage(image.src, image.width, image.height, this.contentWidth));
  }

  /** A picture file read for the document (see readImageFile); null, with a notice, when it can't be used. */
  private async readPicture(file: File): Promise<{ src: string; width: number; height: number } | null> {
    if (!(await isImageFile(file))) {
      this.options.onNotice?.(NOT_AN_IMAGE);
      return null;
    }
    try {
      return await readImageFile(file);
    } catch {
      this.options.onNotice?.(UNSUPPORTED_IMAGE);
      return null;
    }
  }

  // ----- fonts this computer doesn't have (persona-300 B-8) -----

  /**
   * The fonts the document's text shows, each once, in no particular order: the document
   * defaults', those of the styles its paragraphs, runs and tables use (body, headers and
   * footers; each style with what it takes from w:basedOn) and the runs' own. A style that
   * styles.xml defines but no text uses is not counted (Word's templates define dozens).
   */
  documentFonts(): string[] {
    const out = new Set<string>();
    if (!this.view || !this.model) return [];
    const used = new Set<string>();
    for (const d of [this.view.state.doc, ...this.model.headerFooters.map((h) => h.doc)]) {
      d.descendants((node) => {
        const style = node.attrs.styleId as string | null | undefined;
        if (node.type === schema.nodes.paragraph && style) used.add('dx-ps-' + cssIdent(style));
        if (node.type.name === 'table') used.add(style ? 'dx-ts-' + cssIdent(style) : 'dx-ts--none');
        for (const mark of node.marks) {
          if (mark.type === schema.marks.charStyle && mark.attrs.id) used.add('dx-cs-' + cssIdent(String(mark.attrs.id)));
          if (mark.type !== schema.marks.font) continue;
          if (mark.attrs.family) out.add(mark.attrs.family);
          if (mark.attrs.eastAsia) out.add(mark.attrs.eastAsia);
        }
        return true;
      });
    }
    // The style rules (docx/styles.ts, one per line, already resolved along w:basedOn): the
    // document's and the default paragraph style's (no style class) and those of used styles.
    for (const rule of this.model.css.split('\n')) {
      const brace = rule.indexOf('{');
      const classes = rule.slice(0, brace).match(/dx-(?:ps|cs|ts)-[\w-]+/g);
      if (brace < 0 || (classes && !classes.some((c) => used.has(c)))) continue;
      for (const m of rule.slice(brace).matchAll(/--dx-font-[le]:\s*"([^"]+)"/g)) out.add(m[1]);
    }
    return [...out];
  }

  /**
   * The document's fonts this computer doesn't have, with the similar font shown instead (if
   * any): the page and its PDF then look different from Word, which a host can say
   * (「這台電腦沒有『標楷體』，畫面與 PDF 的字型和 Word 會不同」). Call it once the page's fonts have
   * loaded (document.fonts.ready); DocxEditor.vue's `fonts` event does.
   */
  missingFonts(): MissingFont[] {
    return missingFonts(this.documentFonts());
  }

  // ----- the document's size (persona-300 A-9) -----

  /** Bytes of the opened file other than its pictures. */
  private otherBytes = 0;

  /** Bytes of the pictures the body and the headers / footers show (each picture once). */
  private pictureBytes(): number {
    if (!this.view) return 0;
    const seen = new Set<string>();
    for (const d of [this.view.state.doc, ...(this.model?.headerFooters ?? []).map((h) => h.doc)]) {
      d.descendants((node) => {
        if (node.type === schema.nodes.image && typeof node.attrs.src === 'string') seen.add(node.attrs.src);
        return true;
      });
    }
    let n = 0;
    for (const src of seen) n += dataUrlBytes(src);
    return n;
  }

  /**
   * About how big the .docx will be when saved, in bytes: the opened file without its pictures,
   * plus the pictures shown now. Pictures are most of a big document, so this is close enough
   * to warn before the server's limit (50 MB) is near. It walks the document: a host calls it
   * after changes (DocxEditor.vue's `size` event does), not on every key.
   */
  estimatedSize(): number {
    return this.otherBytes + this.pictureBytes();
  }

  // ----- navigation pane and page references (editor/outline.ts) -----

  /** The body's headings (paragraphs with an outline level), in document order. */
  outline(): Heading[] {
    return this.view ? documentOutline(this.view.state.doc, this.model?.styles) : [];
  }

  /** Scrolls a heading to the top of the view and puts the cursor at its start. */
  goToHeading(pos: number): void {
    const view = this.view;
    if (!view || pos < 0 || pos >= view.state.doc.content.size) return;
    view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(pos + 1))));
    const dom = view.nodeDOM(pos);
    if (dom instanceof HTMLElement) dom.scrollIntoView?.({ block: 'start' });
    view.focus();
  }

  /**
   * Puts the cursor where a bookmark starts and scrolls there (a link to a place in the
   * document, e.g. a table of contents entry). False when the body has no bookmark of that name
   * (Word's bookmark names ignore case).
   */
  goToBookmark(name: string): boolean {
    const view = this.view;
    if (!view || !name) return false;
    const marks = bookmarkPositions(view.state.doc);
    let pos = marks.get(name);
    if (pos == null) for (const [n, p] of marks) if (n.toLowerCase() === name.toLowerCase()) pos = p;
    if (pos == null) return false;
    const $pos = view.state.doc.resolve(pos);
    view.dispatch(view.state.tr.setSelection(TextSelection.near($pos)));
    const dom = $pos.depth > 0 ? view.nodeDOM($pos.before()) : null;
    if (dom instanceof HTMLElement) dom.scrollIntoView?.({ block: 'start' });
    view.focus();
    return true;
  }

  /** A click on a link: followed when viewing only, or with Ctrl while editing (editor/linkFollow.ts). */
  private onLinkClick(e: MouseEvent): void {
    const link = linkClick(e, this.options.editable !== false);
    if (!link) return;
    // The browser never opens a link by itself (not even one in a header drawn on a page).
    e.preventDefault();
    if (link.follow) this.followLink(link.href);
  }

  /**
   * Follows a link target as the page shows it (editor/linkFollow.ts): a place in the document is
   * gone to; another site opens in a new tab once the user agreed (the confirm option). The link's
   * right-click menu uses this too.
   */
  followLink(href: string): boolean {
    return followLink(href, {
      confirm: this.options.confirm ?? ((m: string) => window.confirm(m)),
      goToBookmark: (name) => this.goToBookmark(name),
      onNotice: this.options.onNotice,
    });
  }

  /** The document is attached to a macro-enabled template (.dotm), see hasMacros. */
  private macroTemplate = false;

  /**
   * The file holds macros (Word's own programs: word/vbaProject.bin, as a .docm has), ActiveX
   * controls, or is attached to a macro-enabled template (.dotm) Word runs where it has it. The
   * page never runs them and keeps them as they are; the host tells the user (see MACRO_NOTICE).
   */
  hasMacros(): boolean {
    return this.macroTemplate || hasMacroParts(this.model?.zip ?? null);
  }

  /** The page a body position is on; null before layout. */
  private pageAt(pos: number): PageBox | null {
    const view = this.view;
    if (!view || !this.pages.length) return null;
    try {
      const y = view.coordsAtPos(pos).top;
      const canvasY = (y - this.canvas.getBoundingClientRect().top) / this.zoomValue;
      return this.pages.find((p) => canvasY < p.top + p.height + PAGE_GAP / 2) ?? this.pages[this.pages.length - 1];
    } catch {
      return null;
    }
  }

  /** The printed page number (restarting per section as set) at a body position; null before layout. */
  pageNumberAt(pos: number): number | null {
    return this.pageAt(pos)?.number ?? null;
  }

  /**
   * The page number at a body position as Word prints it there (a PAGEREF's result): in that
   * section's page number format (i, ii … / 一, 二 …), or the one the field's \* switch asks
   * for. Null before layout.
   */
  pageLabelAt(pos: number, instr = ''): string | null {
    const n = this.pageNumberAt(pos);
    if (n == null) return null;
    return formatPageNumber(n, fieldNumberFormat(instr) ?? this.pageAt(pos)?.numberFormat);
  }

  /** How many PAGEREF results the body shows (a table of contents' page numbers ...). */
  pageReferenceCount(): number {
    return this.view ? this.countOnce('refs', this.view.state.doc, (d) => pageReferences(d).length) : 0;
  }

  /** Counts of a document (page references, tables of contents), kept for as long as it is shown. */
  private docCounts = new WeakMap<PMNode, Record<string, number>>();

  /** `count(doc)`, worked out once per document (the ribbon asks on every update). */
  private countOnce(key: string, doc: PMNode, count: (doc: PMNode) => number): number {
    let counts = this.docCounts.get(doc);
    if (!counts) this.docCounts.set(doc, (counts = {}));
    return (counts[key] ??= count(doc));
  }

  /**
   * Updates the page numbers shown by PAGEREF fields (the table of contents, 「見第 N 頁」)
   * from this page layout, like Word's 更新功能變數; one undo step. Returns how many changed.
   */
  updatePageReferences(): number {
    const view = this.view;
    if (!view || this.options.editable === false) return 0;
    const tr = view.state.tr;
    const n = updatePageReferences(tr, (pos, instr) => this.pageLabelAt(pos, instr));
    if (n) view.dispatch(tr);
    return n;
  }

  // ----- table of contents (editor/toc.ts) -----

  /** How many tables of contents (TOC fields) the body has. */
  tableOfContentsCount(): number {
    return this.view ? this.countOnce('tocs', this.view.state.doc, (d) => tocFields(d).length) : 0;
  }

  /** The 目錄 styles a table of contents made here used: added on save when the file lacks them (docx/tocStyles.ts). */
  private noteTocStyles(ids: Iterable<string>): void {
    const made = (this.model.tocStylesUsed ??= new Set());
    for (const id of ids) made.add(id);
  }

  private tocContext(): TocContext {
    const p = this.cursorSection().page;
    return {
      styles: this.model.styles,
      paragraphStyles: this.model.paragraphStyles,
      numbering: this.model.numbering,
      pageLabel: (pos) => this.pageLabelAt(pos),
      textWidth: p.width - p.marginLeft - p.marginRight,
    };
  }

  /** A table of contents can't be made or rebuilt as tracked changes: refused while tracking. */
  private tocRefusedWhileTracking(): boolean {
    if (!this.trackOn) return false;
    this.options.onNotice?.('追蹤修訂開啟時，無法插入或重建目錄，請先關閉「追蹤修訂」。');
    return true;
  }

  /**
   * 參考資料 › 目錄: inserts a table of contents of the headings 1-3 at the cursor, as Word's
   * does (a TOC field; see editor/toc.ts), with page numbers from this page layout; one undo step.
   * A document that already has one gets it rebuilt instead. Only in the body's own paragraphs.
   */
  insertTableOfContents(): boolean {
    const view = this.view;
    if (!view || this.options.editable === false || this.hfSession) return false;
    if (this.tocRefusedWhileTracking()) return false;
    if (tocFields(view.state.doc).length) {
      this.updateTableOfContents();
      this.options.onNotice?.('這份文件已有目錄，已依目前的標題更新目錄。');
      return true;
    }
    const { state } = view;
    const $from = state.selection.$from;
    const block = $from.depth >= 1 ? state.doc.child($from.index(0)) : null;
    if (!block || block.type !== schema.nodes.paragraph) {
      this.options.onNotice?.('目錄只能插入在正文的段落中（不能在表格內）。');
      return false;
    }
    const ctx = this.tocContext();
    this.addTocCss();
    const entries = planEntries(state.doc, ctx, { from: 1, to: 3 });
    const blocks = blocksFromXml(newTocXml(entries, ctx));
    this.noteTocStyles(blocks.map((b) => b.attrs.styleId as string | null).filter((id): id is string => !!id));
    const tr = state.tr;
    const blockPos = $from.before(1);
    addHeadingBookmarks(tr, entries);
    const at = tr.mapping.map(blockPos);
    let end: number;
    if (block.content.size === 0 && !block.attrs.sectPr) {
      // An empty paragraph becomes the table of contents.
      tr.replaceWith(at, at + block.nodeSize, blocks);
      end = at + blocks.reduce((s, b) => s + b.nodeSize, 0);
    } else {
      const where = $from.parentOffset === 0 ? at : tr.mapping.map(blockPos + block.nodeSize);
      tr.insert(where, blocks);
      end = where + blocks.reduce((s, b) => s + b.nodeSize, 0);
    }
    tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(end, tr.doc.content.size)), -1)).scrollIntoView();
    view.dispatch(tr);
    this.refreshTocPagesLater();
    view.focus();
    return true;
  }

  /**
   * 參考資料 › 更新目錄: rebuilds each table of contents from the headings now (added, removed or
   * renamed headings, other levels), as Word's 「更新整個目錄」 does, then the page numbers of every
   * PAGEREF (「見第 N 頁」 too) from this page layout; one undo step. A table of contents that
   * already lists the headings as they are only gets its page numbers updated.
   */
  updateTableOfContents(): { rebuilt: boolean; pages: number } {
    const view = this.view;
    if (!view || this.options.editable === false) return { rebuilt: false, pages: 0 };
    const tr = view.state.tr;
    let rebuilt = false;
    if (tocFields(tr.doc).length && !this.tocRefusedWhileTracking()) {
      this.addTocCss();
      const r = rebuildTocs(tr, this.tocContext());
      rebuilt = r.rebuilt;
      this.noteTocStyles(r.styleIds);
    }
    // Page numbers of the positions as laid out now (before the rebuild).
    const back = tr.mapping.invert();
    const pages = updatePageReferences(tr, (pos, instr) => this.pageLabelAt(back.map(pos), instr));
    if (tr.docChanged) view.dispatch(tr);
    // The rebuilt table of contents may move the headings: their page numbers once laid out again.
    if (rebuilt) this.refreshTocPagesLater();
    return { rebuilt, pages };
  }

  /**
   * A file without Word's 目錄 styles gets them on save (docx/tocStyles.ts); until then the page
   * shows the entries as those styles will (a file's own styles are left to its styles.xml).
   */
  private addTocCss(): void {
    const have = new Set(this.model.paragraphStyles.map((s) => s.id));
    const rules = TOC_FALLBACK_CSS.filter(([id]) => !have.has(id)).map(([id, css]) => `.${this.scope} .dx-doc .dx-ps-${id}{${css}}`);
    const text = rules.join('\n');
    if (text && !this.styleEl.textContent?.includes(text)) this.styleEl.textContent += '\n' + text;
  }

  private tocPagesTimer: ReturnType<typeof setTimeout> | undefined;

  /**
   * Once the pages are laid out again after a table of contents was made or rebuilt, its page
   * numbers follow (not an undo step of its own: undo takes back the table of contents whole).
   * Given up when the document changes in the meantime.
   */
  private refreshTocPagesLater(): void {
    const view = this.view;
    if (!view) return;
    const doc = view.state.doc;
    let last = '';
    let same = 0;
    let tries = 0;
    clearTimeout(this.tocPagesTimer);
    const tick = () => {
      if (this.destroyed || this.view !== view || view.state.doc !== doc) return;
      const key = pagesKey(this.pages);
      same = key === last ? same + 1 : 0;
      last = key;
      if (same < 2 && ++tries < 40) {
        this.tocPagesTimer = setTimeout(tick, 150);
        return;
      }
      const tr = view.state.tr;
      if (updatePageReferences(tr, (pos, instr) => this.pageLabelAt(pos, instr))) view.dispatch(tr.setMeta('addToHistory', false));
    };
    this.tocPagesTimer = setTimeout(tick, 150);
  }

  // ----- word count -----

  /** Words and characters of the body and its text boxes (Word counts text boxes; each East Asian character is a word). */
  wordCount(): WordCount {
    if (!this.view) return { words: 0, chars: 0, charsWithSpaces: 0 };
    const doc = this.view.state.doc;
    return addCounts([countDocument(doc), ...editableTextBoxes(doc, true).map((t) => countDocument(frameDoc(t.frame)))]);
  }

  /** The same for the selection (null when nothing is selected). */
  selectionWordCount(): WordCount | null {
    const state = this.view?.state;
    if (!state || state.selection.empty) return null;
    const { from, to } = state.selection;
    return countText(state.doc.textBetween(from, to, '\n', '￼'));
  }

  // ----- spell check -----

  /** Set by setSpellcheck; until then the `spellcheck` option decides. */
  private spellcheckChoice: boolean | null = null;

  private get spellcheck(): boolean {
    return this.spellcheckChoice ?? this.options.spellcheck === true;
  }

  /**
   * Turns the browser's own spell check on or off (red underlines; mostly Latin text: browsers
   * don't check Chinese). Off by default, like the page's other editing surfaces.
   */
  setSpellcheck(on: boolean): void {
    this.spellcheckChoice = on;
    // A no-op update makes the views read their attributes again.
    for (const v of [this.view, this.hfSession?.view, this.shapeSession?.view]) v?.dispatch(v.state.tr.setMeta('addToHistory', false));
  }

  get spellcheckOn(): boolean {
    return this.spellcheck;
  }

  // ----- fields to fill in (content controls in a template; editor/fields.ts) -----

  /** The body's fields, in document order. */
  fields(): DocField[] {
    return this.view ? documentFields(this.view.state.doc) : [];
  }

  /** Required fields still empty or showing their placeholder. */
  unfilledRequired(): DocField[] {
    return this.fields().filter((f) => f.required && !f.filled);
  }

  /** Selects a field's content and scrolls to it. */
  goToField(id: string): boolean {
    return !!this.view && selectField(this.view, id);
  }

  /** Selects the next field that isn't filled in (after the cursor, wrapping around). */
  nextUnfilledField(requiredOnly = false): DocField | null {
    return this.view ? nextUnfilled(this.view, requiredOnly) : null;
  }

  /** Re-measure pages (call after changing page size, fonts, container width ...). */
  repaginate(): void {
    if (this.view) repaginate(this.view);
  }

  // ----- sections & page setup -----

  /** The document's sections (page setup, headers/footers) as they are now. */
  sections(doc: PMNode | undefined = this.view?.state.doc): Section[] {
    return documentSections(doc ?? emptyDoc(), this.model);
  }

  /** The section the cursor is in (or whose header/footer is being edited). */
  cursorSection(): Section {
    const sections = this.sections();
    if (this.hfSession) return sections[this.pages[this.hfSession.pageIndex]?.section ?? 0] ?? sections[sections.length - 1];
    const block = this.view ? this.view.state.selection.$from.index(0) : 0;
    return sectionOfBlock(sections, block);
  }

  /** Change paper size, orientation and margins of the cursor's section, or of every section. */
  setPageSetup(page: PageSetup, scope: PageSetupScope = 'section'): void {
    this.setSectionSetup({ page }, scope);
  }

  /**
   * Change the page numbering (Word's 頁碼格式) of the cursor's section, or of every section:
   * `format` (w:fmt; null = 1, 2, 3) and `start` (restart at a number; null = continue from the
   * previous section).
   */
  setPageNumbering(numbering: Partial<PageNumbering>, scope: PageSetupScope = 'section'): void {
    this.setSectionSetup({ numbering }, scope);
  }

  /**
   * The 版面設定 dialog's changes as one undoable step, for the cursor's section or every
   * section: page size and margins, page numbering, "different first page", text columns and
   * where the section starts (新頁 / 接續本頁). Only what is given changes, and a section already
   * set that way keeps its w:sectPr as written. Refused (with a notice) while 追蹤修訂 is on,
   * like every section change.
   */
  setSectionSetup(
    changes: {
      page?: PageSetup;
      numbering?: Partial<PageNumbering>;
      titlePage?: boolean;
      columns?: ColumnsChange;
      start?: SectionStart;
    },
    scope: PageSetupScope = 'section',
  ): void {
    const view = this.view;
    if (!view || this.refusedWhileTracking()) return;
    const { page, numbering, titlePage, columns, start } = changes;
    const edit = (s: Section, xml: string): string => {
      let out = page ? withPageSetup(xml, page) : xml;
      if (start) out = withSectionStart(out, start);
      if (numbering) out = withPageNumbering(out, numbering);
      if (columns) out = withColumns(out, columns);
      if (titlePage !== undefined && titlePage !== s.titlePage) out = withTitlePage(out, titlePage);
      return out;
    };
    const targets = scope === 'all' ? this.sections() : [this.cursorSection()];
    const tr = view.state.tr;
    for (const s of targets) {
      if (s.pos == null) {
        // The last section's w:sectPr is a document attribute, so this can be undone too.
        const before = this.lastSectPr();
        const xml = edit(s, before);
        if (xml !== before) tr.setDocAttribute('sectPr', xml);
      } else {
        const node = tr.doc.nodeAt(s.pos);
        if (!node?.attrs.sectPr) continue;
        const xml = edit(s, node.attrs.sectPr);
        if (xml !== node.attrs.sectPr) tr.setNodeMarkup(s.pos, undefined, { ...node.attrs, sectPr: xml });
      }
    }
    if (!tr.docChanged) return;
    if (page || titlePage !== undefined) this.pushed.clear();
    view.dispatch(tr);
    repaginate(view);
    // Another first-page header/footer may show: draw the pages again.
    if (titlePage !== undefined) this.renderPages(this.pages);
  }

  /**
   * Set the text columns (Word's 版面配置 > 欄) of the cursor's section, or of every section:
   * `count` equal columns `space` twips apart, with or without a line between them (w:cols).
   * The page shows and prints them as Word does (editor/columnLayout.ts). One undo step; refused
   * while 追蹤修訂 is on.
   */
  setColumns(columns: ColumnsChange, scope: PageSetupScope = 'section'): void {
    this.setSectionSetup({ columns }, scope);
  }

  /**
   * Whether some section has more than one text column (w:cols). The page, its printout and the
   * PDF made from printHtml() show the columns (editor/columnLayout.ts).
   */
  hasMultiColumnSections(): boolean {
    return this.sections().some((s) => s.columns.count > 1);
  }

  /** Sections in columns laid out in one column here (by singleColumnKey), see columnsShownAsOne. */
  private singleColumn = new Set<string>();

  /**
   * Whether a section Word lays out in columns is shown (and printed, and made into a PDF) in one
   * column here: one of its blocks (a long paragraph that may not break, a tall table or
   * picture) is taller than a whole column, which Word breaks between its lines and the page
   * cannot (see editor/pagination.ts). The host says so before making a PDF.
   */
  columnsShownAsOne(): boolean {
    if (!this.singleColumn.size) return false;
    return this.sections().some((s) => s.columns.count > 1 && this.singleColumn.has(singleColumnKey(s)));
  }

  /** The pagination found these sections (by index) are to be laid out in one column: measured again so. */
  private setSingleColumn(indices: number[]): void {
    const view = this.view;
    if (!view) return;
    const sections = this.sections();
    const next = new Set(indices.map((i) => sections[i]).filter((s): s is Section => !!s).map(singleColumnKey));
    if (next.size === this.singleColumn.size && [...next].every((k) => this.singleColumn.has(k))) return;
    this.singleColumn = next;
    repaginate(view);
  }

  /**
   * Insert a column break at the cursor (Word's 分欄符號, Ctrl+Shift+Enter; w:br w:type="column"):
   * in Word the text after it starts the next column. The editor shows it as a marked line and
   * doesn't start a new page for it. Only in body paragraphs (not in tables, headers or footers).
   */
  insertColumnBreak(): boolean {
    return !!this.view && !this.hfSession && this.run(insertColumnBreak);
  }

  /** A page number as Word shows it in a page number format (w:fmt; null = 1, 2, 3). */
  pageNumberText(n: number, format: string | null): string {
    return formatPageNumber(n, format);
  }

  /**
   * The page number a section's first page has when it continues from the previous section
   * (from the current layout; 1 for the first section).
   */
  continuedPageNumber(index: number): number {
    if (index <= 0) return 1;
    const before = this.pages.filter((p) => p.section < index);
    return before.length ? before[before.length - 1].number + 1 : 1;
  }

  /**
   * Start a new section at the cursor (Word's "section break, next page"): the text before
   * the cursor ends a section with this section's page setup and headers/footers.
   */
  insertSectionBreak(): boolean {
    const view = this.view;
    if (!view || this.hfSession || this.refusedWhileTracking()) return false;
    const { $from } = view.state.selection;
    if ($from.depth !== 1 || $from.parent.type !== schema.nodes.paragraph) return false;
    const section = this.cursorSection();
    const xml = sectionBreakXml(section);
    const tr = view.state.tr.deleteSelection();
    const pos = tr.selection.from;
    const $pos = tr.doc.resolve(pos);
    const para = $pos.parent;
    const before = $pos.before();
    // The paragraph's own section break (if it ends the section) stays with its second half.
    // The section now starting at the new break begins on a new page ("next page" break) and
    // continues the page numbering; the new section before it keeps the original start.
    const secondSectPr = para.attrs.sectPr ? startingAtBreak(para.attrs.sectPr) : null;
    // The original paragraph mark (its revision, w:pPrChange) stays with the second half, as
    // after Enter; the first half ends with the new mark that holds the new section break.
    const pPr = splitParagraphPPr(para.attrs.pPr).last;
    tr.split(pos, 1, [{ type: schema.nodes.paragraph, attrs: { ...para.attrs, pPr, sectPr: secondSectPr, pAttrs: null, brGroup: null } }]);
    afterParagraphSplit(tr, before);
    tr.setNodeMarkup(before, undefined, { ...tr.doc.nodeAt(before)!.attrs, sectPr: xml });
    if (section.pos != null && section.pos !== before) {
      const end = tr.mapping.map(section.pos);
      const node = tr.doc.nodeAt(end);
      if (node?.attrs.sectPr) tr.setNodeMarkup(end, undefined, { ...node.attrs, sectPr: startingAtBreak(node.attrs.sectPr) });
    } else if (section.pos == null) {
      tr.setDocAttribute('sectPr', startingAtBreak(this.lastSectPr()));
    }
    view.dispatch(tr.scrollIntoView());
    view.focus();
    return true;
  }

  /** Remove the section break ending the cursor's paragraph: its section joins the next one. */
  removeSectionBreak(): boolean {
    const view = this.view;
    if (!view || this.hfSession || this.refusedWhileTracking()) return false;
    const { $from } = view.state.selection;
    if ($from.depth < 1) return false;
    const pos = $from.before(1);
    const node = view.state.doc.nodeAt(pos);
    if (!node?.attrs.sectPr) return false;
    view.dispatch(view.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, sectPr: null }));
    view.focus();
    return true;
  }

  // ----- zoom & navigation -----

  get zoom(): number {
    return this.zoomValue;
  }

  /** Zoom the page view (1 = 100 %), or 'fit' to fit the widest page into the window. */
  setZoom(zoom: number | 'fit'): void {
    this.fitWidth = zoom === 'fit';
    if (zoom !== 'fit') this.zoomValue = Math.min(4, Math.max(0.25, zoom));
    this.applyZoom();
    // Other text is on screen now: where formatting marks show spaces follows.
    for (const view of [this.view, this.hfSession?.view]) if (view) recheckFormattingMarks(view);
    this.emitUpdate();
  }

  /** Scroll page `index` (0-based) into view. */
  scrollToPage(index: number): void {
    const page = this.pagesLayer.children[Math.max(0, Math.min(index, this.pagesLayer.children.length - 1))] as HTMLElement | undefined;
    page?.scrollIntoView({ block: 'start', behavior: 'smooth' });
  }

  /**
   * Print the pages (the browser's print dialog can also save a PDF). `title` names the printout
   * (the PDF's suggested file name): the host passes the document's name; else the page title.
   */
  print(title?: string | null): void {
    if (!this.view) return;
    // The page being edited shows a live editor instead of its header: draw it statically.
    this.closeShapeText(false);
    this.closeHeaderFooter(false);
    this.shapes.clearOverlay();
    printPages(this.canvas, this.pages, title?.trim() || document.title || 'document', this.scope);
  }

  /**
   * The printout as one self-contained HTML page (styles and pictures inside), for a PDF made
   * elsewhere, e.g. the server's headless browser: the same pages as print(). Null before layout.
   */
  printHtml(title = document.title || 'document'): string | null {
    if (!this.view) return null;
    this.closeShapeText(false);
    this.closeHeaderFooter(false);
    this.shapes.clearOverlay();
    return printHtml(this.canvas, this.pages, title, this.scope);
  }

  private applyZoom(): void {
    const width = parseFloat(this.canvas.style.width) || 0;
    const height = parseFloat(this.canvas.style.minHeight) || 0;
    if (this.fitWidth && width > 0) {
      const available = this.root.clientWidth - 32;
      if (available > 0) this.zoomValue = Math.min(4, Math.max(0.25, available / width));
    }
    const z = this.zoomValue;
    this.canvas.style.transform = z === 1 ? '' : `scale(${z})`;
    this.stage.style.width = `${width * z}px`;
    this.stage.style.height = `${height * z}px`;
  }

  /**
   * The page the cursor is on, for an update sent while editing: the last one measured, checked
   * again at the next animation frame (finding it forces a layout, which is cheaper once the
   * browser lays out the edit anyway). A changed page sends another update.
   */
  private deferredCursorPage(): number {
    if (this.hfSession || !this.view || this.pages.length < 2) return (this.lastCursorPage = this.cursorPage());
    if (!this.pageCheck && !this.destroyed) {
      const check = () => {
        this.cancelPageCheck();
        if (this.destroyed || !this.view) return;
        const page = this.cursorPage();
        if (page === this.lastCursorPage) return;
        this.lastCursorPage = page;
        const snap = this.snapshot();
        if (snap) this.options.onUpdate?.(snap);
      };
      this.pageCheck = {};
      // A frame, or soon anyway where frames don't run (a hidden tab).
      if (typeof requestAnimationFrame === 'function') this.pageCheck.frame = requestAnimationFrame(check);
      this.pageCheck.timer = setTimeout(check, 100);
    }
    return this.lastCursorPage;
  }

  private cancelPageCheck(): void {
    const c = this.pageCheck;
    if (!c) return;
    this.pageCheck = null;
    if (c.frame != null) cancelAnimationFrame(c.frame);
    clearTimeout(c.timer);
  }

  /** 1-based index of the page the cursor is on. */
  private cursorPage(): number {
    if (this.hfSession) return this.hfSession.pageIndex + 1;
    const view = this.view;
    if (!view || this.pages.length < 2) return 1;
    try {
      const y = view.coordsAtPos(view.state.selection.head).top;
      const canvasY = (y - this.canvas.getBoundingClientRect().top) / this.zoomValue;
      const i = this.pages.findIndex((p) => canvasY < p.top + p.height + PAGE_GAP / 2);
      return (i < 0 ? this.pages.length - 1 : i) + 1;
    } catch {
      return 1;
    }
  }

  // ----- headers & footers -----

  /** Start editing the header or footer shown on a page (like double-clicking it in Word). */
  editHeaderFooter(kind: HeaderFooterKind, pageIndex = 0): void {
    if (!this.view || this.options.editable === false) return;
    this.closeShapeText(false);
    this.closeHeaderFooter(false);
    const pages = this.pages.length ? this.pages : [this.firstPage()];
    pageIndex = Math.max(0, Math.min(pageIndex, pages.length - 1));
    const page = pages[pageIndex];
    const sections = this.sections();
    const section = sections[page.section] ?? sections[sections.length - 1];
    const type = variantFor(section, page.inSection, page.number, this.model.evenAndOdd);
    let hf = sectionHeaderFooter(this.model, sections, section.index, kind, type);
    if (!hf) {
      // A new part gets its section's reference once it has content (see claimHeaderFooter),
      // so a later section break copies the reference and opening an empty one changes nothing.
      hf = { kind, type, relId: this.newRelId(), part: null, doc: emptyDoc(), dirty: false, pending: true };
      this.model.headerFooters.push(hf);
    }

    const box = el('div', `dx-hf-box dx-hf-${kind}`);
    // Pending parts are dropped when their editor closes: a pending one here was made just now.
    const created = hf.pending ? hf : null;
    let view: EditorView;
    try {
      box.style.cssText = `top:${page.top}px;left:${page.left}px;height:${page.height}px;width:${page.width}px`;
      box.dataset.margins = pageMargins(page);
      const area = el('div', 'dx-hf-area');
      area.style.cssText = this.hfPosition(kind, section.page);
      box.append(area, this.hfBar(kind, type, page, section, sections.length));
      this.hfLayer.append(box);

      const current = hf;
      const sectionPos = section.pos;
      // Reopening a header/footer continues where it was left, undo history included.
      const kept = this.hfStates.get(hf);
      view = new EditorView(area, {
        state: kept && kept.doc === hf.doc ? kept : EditorState.create({ schema, doc: hf.doc, plugins: this.plugins(false) }),
        attributes: () => ({ class: 'dx-doc dx-hf-doc', spellcheck: String(this.spellcheck), translate: 'no', 'aria-label': kind === 'header' ? '頁首' : '頁尾', lang: 'zh-TW' }),
        nodeViews: {
          field: (node) => new FieldView(node, (k, f) => this.fieldValue(k, this.hfSession?.pageIndex ?? 0, f), this.fieldViews),
          image: imageNodeView,
        },
        dispatchTransaction: (sent) =>
          this.safeDispatch(view, sent, (tr) => {
            if (!tr.docChanged) return;
            current.doc = view.state.doc;
            current.dirty = true;
            // Other references to the same part show the same content (Word edits the part in place).
            for (const other of this.model.headerFooters) {
              if (other !== current && current.part && other.part === current.part) {
                other.doc = current.doc;
                other.dirty = true;
              }
            }
            if (current.pending) this.claimHeaderFooter(current, sectionPos);
            this.options.onChange?.();
            this.scheduleHeaderFooterRender();
          }),
        handleDOMEvents: { paste: (_v, event) => this.notePaste(event) },
        handlePaste: (v, event, slice) => this.onPaste(v, event, slice),
        handleDrop: (_v, event) => this.onDrop(event as DragEvent),
        transformPastedHTML: (html) => this.safePastedHtml(html),
        // A copied watermark is not pasted again: it would be a second shape with the same ids.
        transformPasted: (slice) => withoutWatermarks(this.adoptPastedLists(slice)),
        clipboardSerializer: this.clipboardSerializer,
      });
    } catch (err) {
      // Nothing of the failed editor stays on the page; a part made just for it goes too.
      box.remove();
      if (created) this.model.headerFooters = this.model.headerFooters.filter((h) => h !== created);
      throw err;
    }
    this.hfSession = { hf, view, pageIndex, box };
    // A kept editing state has the marks as they were when it closed.
    showFormattingMarks(view, this.marksOn);
    this.root.classList.add('dx-hf-mode');
    this.renderPages(this.pages);
    view.focus();
    this.emitUpdate();
  }

  /** Leave header/footer editing and return to the body. */
  closeHeaderFooter(focusBody = true): void {
    const s = this.hfSession;
    if (!s) return;
    this.closeShapeText(false);
    this.hfSession = null;
    clearTimeout(this.hfRenderTimer);
    this.hfStates.set(s.hf, s.view.state);
    s.view.destroy();
    s.box.remove();
    // Opened for an earlier section and left empty: nothing to keep.
    if (s.hf.pending) this.model.headerFooters = this.model.headerFooters.filter((h) => h !== s.hf);
    this.root.classList.remove('dx-hf-mode');
    this.renderPages(this.pages);
    if (focusBody) this.view?.focus();
    this.emitUpdate();
  }

  // ----- find & replace across the body, headers and footers (editor/search.ts) -----

  /** A header/footer part's key for the find panel: parts two sections share are one. */
  private hfKey(hf: HeaderFooterPart): string {
    return `${hf.kind}:${hf.part ?? hf.relId ?? hf.type}`;
  }

  /** The header/footer parts shown on the pages (each once), with the page that shows each first. */
  private shownHeaderFooters(): { hf: HeaderFooterPart; pageIndex: number; label: string }[] {
    if (!this.view || !this.model) return [];
    const pages = this.pages.length ? this.pages : [this.firstPage()];
    const sections = this.sections();
    const out = new Map<string, { hf: HeaderFooterPart; pageIndex: number; label: string }>();
    for (const kind of ['header', 'footer'] as const) {
      pages.forEach((page, pageIndex) => {
        const section = sections[page.section] ?? sections[sections.length - 1];
        const type = variantFor(section, page.inSection, page.number, this.model.evenAndOdd);
        const hf = sectionHeaderFooter(this.model, sections, section.index, kind, type);
        if (!hf || hf.pending || out.has(this.hfKey(hf))) return;
        const label = AREA_LABEL[kind] + VARIANT_LABEL[type] + (sections.length > 1 ? `－第 ${section.index + 1} 節` : '');
        out.set(this.hfKey(hf), { hf, pageIndex, label });
      });
    }
    return [...out.values()];
  }

  /**
   * What the find panel searches: the body, then each header part and each footer part shown
   * on a page (once, however many pages or sections show it), each followed by its text boxes
   * (those that can be drawn, so edited, here). Footnotes and comments are not searched.
   * Read-only: the body only (headers/footers and text boxes can't be opened).
   */
  searchParts(): SearchPart[] {
    if (!this.view) return [];
    const parts: SearchPart[] = [{ id: 'body', area: 'body', label: AREA_LABEL.body, doc: this.view.state.doc }];
    if (this.options.editable === false) return parts;
    const hfParts: SearchPart[] = [];
    for (const { hf, label } of this.shownHeaderFooters()) {
      const open = this.hfSession && this.hfKey(this.hfSession.hf) === this.hfKey(hf) ? this.hfSession.view.state.doc : hf.doc;
      hfParts.push({ id: this.hfKey(hf), area: hf.kind, label, doc: open });
    }
    // Text boxes after the part holding them: the body's after the body, a header's after it.
    const boxes = (host: SearchPart): SearchPart[] =>
      editableTextBoxes(host.doc).map(({ key, frame }) => {
        const id = textBoxPartId(host.id, key, frame.i);
        const live = this.shapeSession && this.activeSearchPart() === id ? this.shapeSession.view.state.doc : frameDoc(frame);
        return { id, area: 'textbox' as const, label: host.id === 'body' ? AREA_LABEL.textbox : `${AREA_LABEL.textbox}（${host.label}）`, doc: live };
      });
    parts.push(...boxes(parts[0]));
    for (const p of hfParts) parts.push(p, ...boxes(p));
    return parts;
  }

  /** Whether the document may be edited (false in read-only mode). */
  get editable(): boolean {
    return this.options.editable !== false;
  }

  /** The searchParts id of where the cursor is. */
  activeSearchPart(): string {
    const host = this.hfSession ? this.hfKey(this.hfSession.hf) : 'body';
    return this.shapeSession ? textBoxPartId(host, this.shapeSession.key, this.shapeSession.index) : host;
  }

  /**
   * Make a search part the one being edited, as Word does when Find moves into a header: the
   * body, or the header/footer (opened on the first page that shows it). Returns its editor;
   * null when it can't be opened. The focus stays where it was.
   */
  openSearchPart(id: string): EditorView | null {
    if (!this.view) return null;
    if (id === this.activeSearchPart()) return this.activeView;
    const focused = document.activeElement as HTMLElement | null;
    const box = parseTextBoxPartId(id);
    const host = box?.host ?? id;
    this.closeShapeText(false);
    if (host === 'body') this.closeHeaderFooter(false);
    else if (!this.hfSession || this.hfKey(this.hfSession.hf) !== host) {
      const shown = this.shownHeaderFooters().find((s) => this.hfKey(s.hf) === host);
      if (!shown) return null;
      this.editHeaderFooter(shown.hf.kind, shown.pageIndex);
    }
    if (box) this.editShapeText(box.key, box.index);
    if (this.activeSearchPart() !== id) return null;
    if (focused && focused !== document.activeElement && focused.isConnected && !this.root.contains(focused)) focused.focus();
    return this.activeView;
  }

  /**
   * Change a search part with one transaction (`build` gets its current state), as one undo
   * step of that part: in the body or an open header/footer through its editor; in a closed
   * header/footer on its kept editing state, so opening it and pressing Ctrl+Z undoes it.
   * Tracked as a change while 追蹤修訂 is on. Returns whether it changed.
   */
  changeSearchPart(id: string, build: (state: EditorState) => Transaction | null): boolean {
    if (!this.view || this.options.editable === false) return false;
    const live = id === 'body' ? this.view : id === this.activeSearchPart() ? this.activeView : null;
    if (live) {
      const before = live.state.doc;
      const tr = build(live.state);
      if (!tr || !tr.docChanged) return false;
      live.dispatch(tr);
      return live.state.doc !== before;
    }
    const box = parseTextBoxPartId(id);
    if (box) return this.changeTextBoxes(box.host, [{ id, ...box, build }], false).length > 0;
    const hf = this.shownHeaderFooters().find((s) => this.hfKey(s.hf) === id)?.hf;
    if (!hf) return false;
    const kept = this.hfStates.get(hf);
    const state = kept && kept.doc === hf.doc ? kept : EditorState.create({ schema, doc: hf.doc, plugins: this.plugins(false) });
    let next: EditorState;
    try {
      const built = build(state);
      const tr = built && built.docChanged ? trackTransaction(state, built) : null;
      if (!tr) return false;
      next = state.apply(tr);
    } catch (err) {
      console.error('papyrus: edit failed', err);
      this.options.onNotice?.('這項編輯無法完成，文件保持在上一個狀態。');
      return false;
    }
    this.setKeptPart(hf, next);
    return true;
  }

  /** A closed header/footer part's editing state is now `next` (every reference to the part shows it). */
  private setKeptPart(hf: HeaderFooterPart, next: EditorState): void {
    this.hfStates.set(hf, next);
    for (const other of this.model.headerFooters) {
      if (other === hf || (hf.part && other.part === hf.part)) {
        other.doc = next.doc;
        other.dirty = true;
      }
    }
    this.options.onChange?.();
    this.renderPages(this.pages);
    this.emitUpdate();
  }

  // ----- one undo step over several parts (全部取代 in the body, headers and footers) -----
  // Each part (body, each header/footer) has its own undo history, as the editor has one editor
  // per part. After a change made in several parts at once (changeSearchParts), the next Ctrl+Z
  // (or 復原) undoes it in every part, and the Ctrl+Y after that redoes it everywhere; as long
  // as nothing else was edited meanwhile. A part edited since keeps its own history.

  /** The parts the last grouped change (or its undo) left, each with the content it had then. */
  private group: { kind: 'undo' | 'redo'; parts: { id: string; doc: PMNode }[]; from: { view: EditorView; doc: PMNode } } | null = null;

  /** A search part's current content (see searchParts). */
  private partDoc(id: string): PMNode | null {
    return this.searchParts().find((p) => p.id === id)?.doc ?? null;
  }

  /**
   * Undo / redo as a command: the grouped one when it applies, else the part's own. Asked
   * whether it can run (no dispatch), it only says so and changes nothing.
   */
  private historyCommand(kind: 'undo' | 'redo'): Command {
    const own = kind === 'undo' ? undo : redo;
    return (state, dispatch, view) => (!!view && (dispatch ? this.groupStep(kind, view) : this.groupApplies(kind, view))) || own(state, dispatch, view);
  }

  /**
   * Whether a part still has the content the grouped change left (the same content: an edit
   * undone since gives an equal document, not the same object).
   */
  private static sameDoc(a: PMNode | null | undefined, b: PMNode): boolean {
    return !!a && (a === b || a.eq(b));
  }

  /** Whether an undo / redo from `view` is the grouped one. */
  private groupApplies(kind: 'undo' | 'redo', view: EditorView): boolean {
    const g = this.group;
    if (!g || g.kind !== kind) return false;
    // From where it was made (still unchanged), or from a part it changed (still as it left it).
    const id = view === this.view ? 'body' : this.activeSearchPart();
    const own = g.parts.find((p) => p.id === id);
    if (own ? !DocxEditor.sameDoc(view.state.doc, own.doc) : g.from.view !== view || !DocxEditor.sameDoc(view.state.doc, g.from.doc)) return false;
    return g.parts.some((p) => DocxEditor.sameDoc(this.partDoc(p.id), p.doc));
  }

  /** The grouped undo / redo, when it applies (see above); returns whether it ran. */
  private groupStep(kind: 'undo' | 'redo', view: EditorView): boolean {
    const g = this.group;
    if (!g || !this.groupApplies(kind, view)) return false;
    this.group = null;
    const step = kind === 'undo' ? undo : redo;
    const done: { id: string; doc: PMNode }[] = [];
    for (const p of g.parts) {
      if (!DocxEditor.sameDoc(this.partDoc(p.id), p.doc)) continue; // edited since: its own history
      const live = p.id === 'body' ? this.view : p.id === this.activeSearchPart() ? this.activeView : null;
      if (live) {
        if (step(live.state, live.dispatch, live)) done.push({ id: p.id, doc: live.state.doc });
        continue;
      }
      const hf = this.shownHeaderFooters().find((s) => this.hfKey(s.hf) === p.id)?.hf;
      const kept = hf && this.hfStates.get(hf);
      if (!hf || !kept || !DocxEditor.sameDoc(hf.doc, kept.doc)) continue;
      let next: EditorState | null = null;
      step(kept, (tr) => (next = kept.apply(tr)));
      if (next) {
        this.setKeptPart(hf, next);
        done.push({ id: p.id, doc: (next as EditorState).doc });
      }
    }
    if (done.length) this.group = { kind: kind === 'undo' ? 'redo' : 'undo', parts: done, from: { view, doc: view.state.doc } };
    return done.length > 0;
  }

  /**
   * Change several search parts (see changeSearchPart) as one step for undo: the next Ctrl+Z or
   * 復原 — in whichever of these parts, or where the cursor was — undoes all of them, as Word's
   * 全部取代 does. Returns the ids of the parts that changed.
   */
  changeSearchParts(changes: { id: string; build: (state: EditorState) => Transaction | null }[]): string[] {
    const changed: string[] = [];
    // A closed text box is changed in the part holding it, in the same undo step as that part.
    const boxes = new Map<string, { id: string; key: string; index: number; build: (state: EditorState) => Transaction | null }[]>();
    const active = this.activeSearchPart();
    for (const c of changes) {
      const box = c.id === active ? null : parseTextBoxPartId(c.id);
      if (box) boxes.set(box.host, [...(boxes.get(box.host) ?? []), { id: c.id, ...box, build: c.build }]);
      else if (this.changeSearchPart(c.id, c.build)) changed.push(c.id);
    }
    const hosts: string[] = [];
    for (const [host, list] of boxes) {
      const done = this.changeTextBoxes(host, list, changed.includes(host));
      if (done.length && !changed.includes(host)) hosts.push(host);
      changed.push(...done);
    }
    const view = this.activeView;
    if (changed.length && view) {
      const parts = [...changed, ...hosts].map((id) => ({ id, doc: this.partDoc(id)! })).filter((p) => p.doc);
      this.group = { kind: 'undo', parts, from: { view, doc: view.state.doc } };
      this.emitUpdate();
    }
    return changed;
  }

  // ----- text boxes (editor/shapeEdit.ts) -----

  /**
   * Edit the text of text box `index` of the shape with model key `key` (docx/shapes.ts), in
   * the body or in the header/footer being edited, like double-clicking it in Word. Returns
   * whether it opened; a shape that can't be drawn here can't be edited (a notice says so).
   * `at`: where the cursor goes (client coordinates, e.g. of the double-click).
   */
  editShapeText(key: string, index: number, at?: { x: number; y: number }): boolean {
    if (!this.view || this.options.editable === false) return false;
    this.closeShapeText(false);
    const hf = this.hfSession;
    const view = hf && findShape(hf.view.state.doc, key) ? hf.view : this.view;
    const found = findShape(view.state.doc, key);
    if (!found || !textFrames(found.shape).some((f) => f.i === index)) return false;
    if (!found.shape.drawable) {
      this.options.onNotice?.(`${WORD_ONLY}。`);
      return false;
    }
    const session = new ShapeTextSession(
      {
        view,
        layer: this.hfLayer,
        plugins: () => this.plugins(false, true),
        nodeViews: { image: imageNodeView },
        attributes: () => ({ spellcheck: String(this.spellcheck), translate: 'no', lang: 'zh-TW' }),
        dispatch: (v, tr, after) => this.safeDispatch(v, tr, after),
        onChange: () => this.options.onChange?.(),
        onClose: (s) => {
          if (this.shapeSession !== s) return;
          this.shapeSession = null;
          this.emitUpdate();
        },
      },
      key,
      index,
    );
    this.shapeSession = session;
    session.focus(at);
    this.emitUpdate();
    return true;
  }

  /** Tell screen readers (the live region). */
  announce(message: string): void {
    this.live.textContent = '';
    // A new text in a fresh task, so the same message twice is said twice.
    setTimeout(() => (this.live.textContent = message), 30);
  }

  /** 插入 › 圖案: place a new shape of that kind (SHAPE_KINDS id) with a click or a drag on the page. */
  startInsertShape(kind: string): boolean {
    this.closeShapeText(false);
    const view = this.activeView;
    return !!view && this.options.editable !== false && this.shapes.startPlacing(kind, view);
  }

  /** 選取圖形: select the next (previous) shape from the cursor; Tab goes on to the next. */
  selectShape(dir: 1 | -1 = 1): boolean {
    this.closeShapeText(false);
    const view = this.activeView;
    return !!view && this.shapes.selectNext(view, dir);
  }

  /** Stop editing a text box (its text is already in the document). */
  closeShapeText(focus = true): void {
    this.shapeSession?.close(focus);
  }

  /**
   * For commands that change the document around the cursor (a table of contents, going to a
   * bookmark ...): ends the text box editing first, so they work on the body. Returns whether one was open.
   */
  closeShapeSession(): boolean {
    const open = !!this.shapeSession;
    this.closeShapeText(false);
    return open;
  }

  /**
   * Change closed text boxes of one part (the body, or a header/footer) with `build` on each
   * one's text, as one step of that part (joined to the step just made there when `join`).
   * Returns the search part ids of the text boxes that changed.
   */
  private changeTextBoxes(
    host: string,
    list: { id: string; key: string; index: number; build: (state: EditorState) => Transaction | null }[],
    join: boolean,
  ): string[] {
    const hfLive = this.hfSession && this.hfKey(this.hfSession.hf) === host ? this.hfSession.view : null;
    const live = host === 'body' ? this.view : hfLive;
    const hf = live ? null : this.shownHeaderFooters().find((s) => this.hfKey(s.hf) === host)?.hf;
    if (!live && !hf) return [];
    const kept = hf ? this.hfStates.get(hf) : null;
    const state = live ? live.state : kept && kept.doc === hf!.doc ? kept : EditorState.create({ schema, doc: hf!.doc, plugins: this.plugins(false) });
    const tr = state.tr;
    const done: string[] = [];
    for (const c of list) {
      const found = findShape(tr.doc, c.key);
      const frame = found && textFrames(found.shape).find((f) => f.i === c.index);
      if (!found || !frame) continue;
      // The text box's text as its own editor would change it (tracked while 追蹤修訂 is on).
      const text = EditorState.create({ schema, doc: frameDoc(frame), plugins: this.plugins(false) });
      let next: PMNode;
      try {
        const built = c.build(text);
        const tracked = built && built.docChanged ? trackTransaction(text, built) : null;
        if (!tracked) continue;
        next = text.apply(tracked).doc;
      } catch (err) {
        console.error('papyrus: edit failed', err);
        continue;
      }
      tr.setNodeMarkup(found.pos, undefined, { ...found.node.attrs, shape: withText(found.shape, c.index, next.toJSON() as Record<string, unknown>) }, found.node.marks);
      done.push(c.id);
    }
    if (!done.length) return [];
    tr.setMeta(NO_TRACK, true);
    // prosemirror-history puts an "appended" transaction in the step before it.
    if (join) tr.setMeta('appendedTransaction', state.tr);
    else closeHistory(tr);
    if (live) live.dispatch(tr);
    else this.setKeptPart(hf!, state.apply(tr));
    return done;
  }

  /** A header/footer made for an earlier section got content: its section now refers to it. */
  private claimHeaderFooter(hf: HeaderFooterPart, sectionPos: number | null): void {
    const view = this.view;
    if (!view || !hf.relId) return;
    const tr = view.state.tr.setMeta('addToHistory', false);
    if (sectionPos == null) {
      tr.setDocAttribute('sectPr', withReference(this.lastSectPr(), hf.kind, hf.type, hf.relId));
    } else {
      const node = view.state.doc.nodeAt(sectionPos);
      if (!node?.attrs.sectPr) return;
      tr.setNodeMarkup(sectionPos, undefined, { ...node.attrs, sectPr: withReference(node.attrs.sectPr, hf.kind, hf.type, hf.relId) });
    }
    hf.pending = false;
    view.dispatch(tr);
  }

  /** The last section's w:sectPr, with its page setup written in when the file had none. */
  private lastSectPr(): string {
    const xml = this.view?.state.doc.attrs.sectPr ?? this.model.sectPr;
    return xml ?? lastSectionXml(null, this.model.page, this.model.titlePage);
  }

  /** Keeps the model's copy of the last section (used by the writer and layout) in step with the document. */
  private syncLastSection(): void {
    const xml = (this.view?.state.doc.attrs.sectPr as string | null | undefined) ?? null;
    if (xml === this.model.sectPr) return;
    if (xml == null) {
      // Undone back to a file that had no body-level w:sectPr.
      Object.assign(this.model, { sectPr: null, page: { ...this.openedLast.page }, titlePage: this.openedLast.titlePage });
      return;
    }
    const { page, titlePage } = readSectPr(xml);
    this.model.sectPr = xml;
    this.model.page = page;
    this.model.titlePage = titlePage;
  }

  /** A relationship id not used by the package or by any header/footer. */
  private newRelId(): string {
    const used = new Set([...(this.model.relIds ?? []), ...this.model.headerFooters.map((h) => h.relId)]);
    let n = 1;
    while (used.has(`rIdPxHf${n}`)) n++;
    return `rIdPxHf${n}`;
  }

  /** Word's "Different First Page", for the section being edited (or the cursor's). */
  setTitlePage(on: boolean): void {
    const section = this.cursorSection();
    if (section.titlePage === on || this.refusedWhileTracking()) return;
    if (section.pos == null && this.view) {
      this.view.dispatch(this.view.state.tr.setDocAttribute('sectPr', withTitlePage(this.lastSectPr(), on)));
    } else if (this.view && section.pos != null) {
      const node = this.view.state.doc.nodeAt(section.pos)!;
      this.view.dispatch(
        this.view.state.tr.setNodeMarkup(section.pos, undefined, { ...node.attrs, sectPr: withTitlePage(node.attrs.sectPr, on) }),
      );
    }
    const s = this.hfSession;
    if (s) this.editHeaderFooter(s.hf.kind, s.pageIndex);
    else this.renderPages(this.pages);
  }

  /**
   * Whether the header/footer being edited is linked to the previous section's (Word's 「連結到前一節」):
   * null when that does not apply (not editing one, or the first section).
   */
  headerFooterLinked(): boolean | null {
    const at = this.hfSessionSection();
    if (!at || at.section.index === 0) return null;
    return !this.ownHeaderFooter(at.section, at.kind, at.type, at.last);
  }

  /**
   * GOV-088, Word's 「連結到前一節」 for the header/footer being edited. Off: the section gets its own,
   * starting as a copy of what it showed. On: its own is dropped (after asking, when it has content)
   * and it shows the previous section's again.
   */
  setHeaderFooterLinked(link: boolean): void {
    const s = this.hfSession;
    const at = this.hfSessionSection();
    const view = this.view;
    if (!s || !at || !view || at.section.index === 0 || this.headerFooterLinked() === link) return;
    if (this.refusedWhileTracking()) return this.editHeaderFooter(at.kind, s.pageIndex);
    const { section, kind, type } = at;
    const sectPr = (xml: string) => {
      const tr = view.state.tr.setMeta('addToHistory', false);
      if (section.pos == null) tr.setDocAttribute('sectPr', xml);
      else tr.setNodeMarkup(section.pos, undefined, { ...view.state.doc.nodeAt(section.pos)!.attrs, sectPr: xml });
      view.dispatch(tr);
    };
    const current = section.pos == null ? this.lastSectPr() : (view.state.doc.nodeAt(section.pos)!.attrs.sectPr as string);
    if (link) {
      const own = this.ownHeaderFooter(section, kind, type, at.last)!;
      const name = kind === 'header' ? '頁首' : '頁尾';
      const ask = this.options.confirm ?? ((m: string) => window.confirm(m));
      if (hasContent(own.doc) && !ask(`要刪除這一節自己的${name}，改為和前一節相同嗎？`)) {
        return this.editHeaderFooter(kind, s.pageIndex); // the check box shows the state again
      }
      sectPr(withoutReference(current, kind, type));
      // One made in the editor goes; one from the file stays in the package, no longer used.
      if (!own.part) this.model.headerFooters = this.model.headerFooters.filter((h) => h !== own);
    } else {
      const hf: HeaderFooterPart = { kind, type, relId: this.newRelId(), part: null, doc: s.hf.doc, dirty: true, pending: false };
      this.model.headerFooters.push(hf);
      sectPr(withReference(current, kind, type, hf.relId!));
    }
    this.options.onChange?.();
    this.editHeaderFooter(kind, s.pageIndex);
  }

  /** The section, variant and header/footer kind of the header/footer being edited. */
  private hfSessionSection(): { section: Section; kind: HeaderFooterKind; type: HeaderFooterType; last: boolean } | null {
    const s = this.hfSession;
    if (!s) return null;
    const sections = this.sections();
    const page = this.pages[s.pageIndex] ?? this.firstPage();
    const section = sections[page.section] ?? sections[sections.length - 1];
    const type = variantFor(section, page.inSection, page.number, this.model.evenAndOdd);
    return { section, kind: s.hf.kind, type, last: section.index === sections.length - 1 };
  }

  /** The section's own header/footer of this variant (not one it shows from an earlier section). */
  private ownHeaderFooter(section: Section, kind: HeaderFooterKind, type: HeaderFooterType, last: boolean): HeaderFooterPart | undefined {
    const relId = section.refs[kind][type];
    if (relId) return this.model.headerFooters.find((h) => h.relId === relId && h.kind === kind);
    return last ? this.model.headerFooters.find((h) => !h.relId && !h.pending && h.kind === kind && h.type === type) : undefined;
  }

  // ----- 浮水印 (設計 › 頁面背景 › 浮水印; docx/watermark.ts, editor/watermark.ts) -----

  /**
   * The document's watermark, as Word's 浮水印 dialog shows it: the one in the header of the
   * cursor's section, else in any header. Null when there is none.
   */
  watermark(): Watermark | null {
    if (!this.view || !this.model) return null;
    const sections = this.sections();
    const cursor = this.cursorSection();
    const own = sectionHeaderFooter(this.model, sections, cursor.index, 'header', 'default');
    const shown = sectionHeaders(this.model, sections);
    for (const hf of [own, ...shown.keys()]) {
      if (!hf || hf.pending) continue;
      const wm = docWatermark(hf.doc);
      // A picture's 自動 is judged against the page of a section showing the part.
      const settings = wm && watermarkSettings(wm.shape, wm.src, watermarkArea(shown.get(hf) ?? cursor));
      if (settings) return settings;
    }
    return null;
  }

  /**
   * Whether a header the sections show holds a watermark shape, even one that can't be shown
   * (a picture missing from the package): 移除浮水印 can then take it out.
   */
  hasWatermark(): boolean {
    if (!this.view || !this.model) return false;
    return [...sectionHeaders(this.model, this.sections()).keys()].some((hf) => watermarkNodes(hf.doc).length > 0);
  }

  /**
   * Word's 浮水印: put the watermark (null: 移除浮水印) into every header part the sections refer
   * to (default, first page and even pages), replacing the watermarks there, Word's and ours;
   * nothing else in the headers changes. A section that shows no header of a kind it uses (its
   * default one, the first-page one with 不同的首頁, the even-page one with 奇偶頁不同) gets one
   * made for it, holding just the watermark. Returns whether anything changed.
   *
   * Only the header parts the sections refer to (or show, linked to an earlier section) are
   * touched; footers and header parts no section uses are left as they are.
   *
   * NOT AN UNDO STEP, by decision: the headers keep their own undo histories apart from the
   * body's, and one Ctrl+Z step across the body and several header parts would need a step type
   * of its own. So Ctrl+Z (in the body or in a header) does not take a watermark out or bring an
   * old one back; the 浮水印 menu removes (移除浮水印) or replaces it. Refused while 追蹤修訂 is on
   * (a watermark can't be recorded as a tracked change here), and in read-only mode.
   */
  setWatermark(watermark: Watermark | null): boolean {
    const view = this.view;
    if (!view || this.options.editable === false) return false;
    if (this.trackOn) {
      this.options.onNotice?.('追蹤修訂開啟時無法設定或移除浮水印（這項變更無法記錄為修訂），請先關閉「追蹤修訂」。');
      return false;
    }
    // The header being edited is closed while its parts change and opened again afterwards.
    const session = this.hfSession ? { kind: this.hfSession.hf.kind, pageIndex: this.hfSession.pageIndex } : null;
    this.closeHeaderFooter(false);

    const sections = this.sections();
    // Where it goes (editor/watermark.ts): the header parts the sections show, and a header made
    // for each section that uses a type of header and shows none.
    const targets = new Map<HeaderFooterPart, Section>();
    const made: HeaderFooterPart[] = [];
    const newRefs = new Map<Section, { type: HeaderFooterType; relId: string }[]>();
    if (watermark) {
      const plan = watermarkTargets(this.model, sections);
      for (const [hf, section] of plan.parts) targets.set(hf, section);
      for (const { section, type } of plan.missing) {
        const hf: HeaderFooterPart = { kind: 'header', type, relId: this.newRelId(), part: null, doc: emptyDoc(), dirty: true, pending: false };
        this.model.headerFooters.push(hf);
        made.push(hf);
        targets.set(hf, section);
        newRefs.set(section, [...(newRefs.get(section) ?? []), { type, relId: hf.relId! }]);
      }
      if (!targets.size) {
        this.options.onNotice?.('找不到可以放浮水印的頁首（這份文件的頁首無法讀取），浮水印沒有設定。');
        if (session) this.editHeaderFooter(session.kind, session.pageIndex);
        return false;
      }
    }
    // The header parts the sections show (their own, or an earlier section's): the only ones
    // cleared or given the watermark. One part referenced twice is changed once (the other
    // reference shows the same content).
    const shown = new Set<HeaderFooterPart>([...targets.keys(), ...sectionHeaders(this.model, sections).keys()]);
    const parts = new Map<string | HeaderFooterPart, HeaderFooterPart>();
    for (const hf of shown) {
      const key = hf.part ?? hf;
      if (!parts.has(key) || targets.has(hf)) parts.set(key, hf);
    }

    const ids = usedShapeIds([view.state.doc, ...this.model.headerFooters.map((h) => h.doc)]);
    let spid = Math.max(2048, ...ids.spids);
    const newNumber = () => {
      let n: number;
      do n = 100000000 + Math.floor(Math.random() * 900000000);
      while (ids.numbers.has(String(n)));
      ids.numbers.add(String(n));
      return n;
    };
    // Word gives each shape its own z-index, 1024 below the one before.
    let zIndex = WATERMARK_Z_INDEX + 1024;
    // The text measured once in its font: the shape has the text's proportions (Word stretches
    // the text to its shape).
    const metrics = watermark?.kind === 'text' ? measureWatermarkText(watermark.text, watermark.font) : null;
    let changed = made.length > 0;
    for (const hf of parts.values()) {
      const section = watermark ? targets.get(hf) : undefined;
      changed = this.changeHeaderFooterDoc(hf, (state) => {
        let node: PMNode | null = null;
        if (watermark && section) {
          const area = watermarkArea(section);
          // The shape type is defined once in a part: unless something else there defines it.
          const old = new Set(watermarkNodes(state.doc).map((w) => w.node));
          const defined = keptXml(state.doc, (n) => old.has(n)).includes(`id="${watermarkShapetype(watermark)}"`);
          zIndex -= 1024;
          const xml = watermarkRunXml(watermark, area, { number: newNumber(), spid: ++spid, shapetype: !defined, zIndex }, metrics);
          const picture = watermark.kind === 'picture' ? watermark : null;
          node = schema.nodes.raw_inline.create({
            xml,
            label: '浮水印',
            src: picture?.src ?? null,
            width: picture?.width ?? null,
            height: picture?.height ?? null,
          });
        }
        return replaceWatermarkTr(state, node);
      }) || changed;
    }

    if (newRefs.size) {
      // The sections refer to the headers made for them (not an undo step either, see above).
      const tr = view.state.tr.setMeta('addToHistory', false);
      for (const [section, refs] of newRefs) {
        if (section.pos == null) {
          tr.setDocAttribute('sectPr', refs.reduce((xml, r) => withReference(xml, 'header', r.type, r.relId), this.lastSectPr()));
        } else {
          const node = tr.doc.nodeAt(section.pos)!;
          const sectPr = refs.reduce((xml, r) => withReference(xml, 'header', r.type, r.relId), node.attrs.sectPr as string);
          tr.setNodeMarkup(section.pos, undefined, { ...node.attrs, sectPr });
        }
      }
      view.dispatch(tr);
    }
    if (changed) this.options.onChange?.();
    if (session) this.editHeaderFooter(session.kind, session.pageIndex);
    else this.renderPages(this.pages);
    this.emitUpdate();
    return changed;
  }

  /**
   * Change a header/footer's content with one transaction on its kept editing state (so its
   * undo history and selection carry on when it is opened), and every reference to the same part
   * with it. Returns whether it changed.
   */
  private changeHeaderFooterDoc(hf: HeaderFooterPart, build: (state: EditorState) => Transaction | null): boolean {
    const kept = this.hfStates.get(hf);
    const state = kept && kept.doc === hf.doc ? kept : EditorState.create({ schema, doc: hf.doc, plugins: this.plugins(false) });
    const tr = build(state);
    if (!tr || !tr.docChanged) return false;
    const next = state.apply(tr);
    this.hfStates.set(hf, next);
    for (const other of this.model.headerFooters) {
      if (other === hf || (hf.part && other.part === hf.part)) {
        other.doc = next.doc;
        other.dirty = true;
      }
    }
    return true;
  }

  /** Insert a PAGE / NUMPAGES field at the cursor. */
  insertField(instr: 'PAGE' | 'NUMPAGES'): boolean {
    return this.run((state, dispatch) => {
      if (dispatch) {
        const node = schema.nodes.field.create({ instr, text: '1' }, null, state.storedMarks ?? state.selection.$from.marks());
        dispatch(state.tr.replaceSelectionWith(node, true).scrollIntoView());
      }
      return true;
    });
  }

  private hfPosition(kind: HeaderFooterKind, p: PageSetup): string {
    const sides = `left:${twipsToPx(p.marginLeft)}px;right:${twipsToPx(p.marginRight)}px`;
    return kind === 'header' ? `top:${twipsToPx(p.header)}px;${sides}` : `bottom:${twipsToPx(p.footer)}px;${sides}`;
  }

  /** The small label + controls shown while editing a header/footer. */
  private hfBar(kind: HeaderFooterKind, type: HeaderFooterType, page: PageBox, section: Section, sectionCount: number): HTMLElement {
    const bar = el('div', 'dx-hf-bar');
    const edge = kind === 'header' ? page.textTop : page.height - page.textBottom;
    bar.style.cssText = kind === 'header' ? `top:${edge}px` : `top:${edge}px;transform:translateY(-100%)`;
    const label = el('span', 'dx-hf-label');
    label.textContent =
      (kind === 'header' ? '頁首' : '頁尾') + VARIANT_LABEL[type] + (sectionCount > 1 ? `－第 ${section.index + 1} 節` : '');

    const first = document.createElement('label');
    first.className = 'dx-hf-check';
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = section.titlePage;
    box.addEventListener('change', () => this.setTitlePage(box.checked));
    first.append(box, document.createTextNode('首頁不同'));

    // 連結到前一節 (not for the first section, which has none before it). The bar is made before the
    // session starts: worked out from the section itself.
    const linked = section.index === 0 ? null : !this.ownHeaderFooter(section, kind, type, section.index === sectionCount - 1);
    const link = document.createElement('label');
    link.className = 'dx-hf-check dx-hf-link';
    const linkBox = document.createElement('input');
    linkBox.type = 'checkbox';
    linkBox.checked = !!linked;
    linkBox.addEventListener('change', () => this.setHeaderFooterLinked(linkBox.checked));
    link.title = '勾選：和前一節用同一個' + (kind === 'header' ? '頁首' : '頁尾') + '；取消勾選：這一節用自己的';
    link.append(linkBox, document.createTextNode('連結到前一節'));

    const button = (text: string, title: string, fn: () => void) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = text;
      b.title = title;
      b.addEventListener('mousedown', (e) => e.preventDefault()); // keep the cursor in the header
      b.addEventListener('click', fn);
      return b;
    };
    bar.append(
      label,
      first,
      ...(linked == null ? [] : [link]),
      button('插入頁碼', '在游標處插入目前頁碼', () => this.insertField('PAGE')),
      button('插入總頁數', '在游標處插入總頁數', () => this.insertField('NUMPAGES')),
      button('關閉頁首及頁尾', '回到內文 (Esc)', () => this.closeHeaderFooter()),
    );
    return bar;
  }

  private onDoubleClick(e: MouseEvent): void {
    if (!this.view || this.options.editable === false) return;
    // Inside the text box being edited: a word is selected, as usual.
    if (this.shapeSession?.box.contains(e.target as Node)) return;
    // A text box or a shape with text: edit its text (in the body, or in the header being edited).
    const box = (e.target as Element).closest?.('.dx-pages') ? null : textBoxAt(e.target as Element);
    if (box) {
      if ('wordOnly' in box) this.options.onNotice?.(`${WORD_ONLY}。`);
      else this.editShapeText(box.key, box.index, { x: e.clientX, y: e.clientY });
      return;
    }
    const s = this.hfSession;
    if (s && s.box.querySelector('.dx-hf-area')!.contains(e.target as Node)) return;
    if ((e.target as HTMLElement).closest?.('.dx-hf-bar')) return;
    const rect = this.canvas.getBoundingClientRect();
    const x = (e.clientX - rect.left) / this.zoomValue;
    const y = (e.clientY - rect.top) / this.zoomValue;
    const pageIndex = this.pages.findIndex((p) => y >= p.top && y <= p.top + p.height && x >= p.left && x <= p.left + p.width);
    if (pageIndex < 0) return;
    const page = this.pages[pageIndex];
    const yInPage = y - page.top;
    if (yInPage < page.textTop) this.editHeaderFooter('header', pageIndex);
    else if (yInPage > page.height - page.textBottom) this.editHeaderFooter('footer', pageIndex);
    else if (s) this.closeHeaderFooter();
  }

  private scheduleHeaderFooterRender(): void {
    clearTimeout(this.hfRenderTimer);
    this.hfRenderTimer = setTimeout(() => this.renderPages(this.pages), 150);
  }

  /**
   * A live field's value on a page, as Word shows it: PAGE in the section's page number format
   * (w:pgNumType/@w:fmt); NUMPAGES (all pages) and SECTIONPAGES as 1, 2, 3. A \* switch
   * (`format`: PAGE \* roman ...) overrides either.
   */
  private fieldValue(kind: string, pageIndex: number, format: string | null = null): string | null {
    const page = this.pages[pageIndex];
    if (kind === 'PAGE') return formatPageNumber(page?.number ?? pageIndex + 1, format ?? page?.numberFormat);
    if (kind === 'NUMPAGES') return formatPageNumber(Math.max(1, this.pages.length), format);
    if (kind === 'SECTIONPAGES') {
      return formatPageNumber(page ? this.pages.filter((p) => p.section === page.section).length : this.pageCount, format);
    }
    return null;
  }

  // ----- internals -----

  /** Page geometry per section for the pagination plugin; headers taller than the margin push the text area. */
  private layout(doc: PMNode): Layout {
    const compat = this.model?.compat;
    return {
      gap: PAGE_GAP,
      topSpacing: !compat ? 'keep' : compat.mode >= 15 ? 'suppress' : compat.suppressSpaceAfterPageBreak ? 'afterPageBreak' : 'keep',
      sections: this.sections(doc).map((s) => {
        const geometry = this.geometry(s);
        const columns = columnGeometry(s.xml, geometry.width - geometry.marginLeft - geometry.marginRight);
        // A section with a block taller than a column: laid out in one column (see columnsShownAsOne).
        const off = !!columns && this.singleColumn.has(singleColumnKey(s));
        return {
          geometry,
          start: s.start,
          pageNumberStart: s.pageNumberStart,
          pageNumberFormat: s.pageNumberFormat,
          firstBlock: s.firstBlock,
          lastBlock: s.lastBlock,
          // Text columns, shown as Word lays them out (editor/columnLayout.ts).
          columns: off ? null : columns,
          columnsOff: off ? columns : null,
        };
      }),
    };
  }

  private geometry(s: Section): PageGeometry {
    const g = geometryOf(s.page);
    const push = this.pushed.get(s.index);
    return { ...g, textTop: Math.max(g.textTop, push?.top ?? 0), textBottom: Math.max(g.textBottom, push?.bottom ?? 0) };
  }

  /** The first page before anything was measured. */
  private firstPage(): PageBox {
    const s = this.sections()[0];
    return { ...this.geometry(s), section: 0, inSection: 0, number: s.pageNumberStart ?? 1, numberFormat: s.pageNumberFormat, top: 0, left: 0 };
  }

  private setDocument(doc: PMNode, model: DocxModel): void {
    if (this.destroyed) return;
    this.closeShapeText(false);
    this.closeHeaderFooter(false);
    // The new document is built first (its plugins read the model as they start) and shown only
    // once it is ready: one that fails leaves the open document, and its model, as they were.
    const before = { model: this.model, pushed: this.pushed, openedLast: this.openedLast, revisionIds: this.revisionIds, bodyStyle: this.bodyStyle };
    const host = el('div', 'dx-host');
    this.host.after(host);
    let view: EditorView;
    try {
      this.model = model;
      this.pushed = new Map();
      this.singleColumn = new Set();
      this.openedLast = { page: { ...model.page }, titlePage: model.titlePage };
      this.revisionIds = this.newRevisionIds();
      // The last section's w:sectPr travels with the document (undoable page setup).
      if (doc.attrs.sectPr !== model.sectPr) doc = schema.nodes.doc.create({ ...doc.attrs, sectPr: model.sectPr }, doc.content);
      let state = EditorState.create({ schema, doc, plugins: this.plugins(true) });
      const fix = fixTables(state);
      if (fix) state = state.apply(fix.setMeta('addToHistory', false));
      this.bodyStyle = this.bodyPadding(state.doc);
      view = new EditorView(host, {
        state,
        editable: () => this.options.editable !== false,
        // A function, so a new padding (other page setup, taller header) is applied on the next update.
        // translate="no" (here and in the header/footer editor): a browser's page translation rewrites
        // the text in place, and the editor would take that for typing and save the translation.
        // An accessible name and the page's language for screen readers (persona-300).
        attributes: () => ({ class: 'dx-doc', spellcheck: String(this.spellcheck), translate: 'no', style: this.bodyStyle, 'aria-label': '文件內容', lang: 'zh-TW' }),
        nodeViews: { image: imageNodeView },
        dispatchTransaction: (sent) =>
          this.safeDispatch(view, sent, (tr) => {
            if (!tr.docChanged) return;
            this.syncLastSection();
            this.options.onChange?.();
            // Section breaks or page setup changed: the column may need other margins.
            this.syncBodyStyle();
          }),
        handleDOMEvents: { paste: (_v, event) => this.notePaste(event) },
        handlePaste: (v, event, slice) => this.onPaste(v, event, slice),
        handleDrop: (_view, event) => this.onDrop(event as DragEvent),
        transformPastedHTML: (html) => this.safePastedHtml(html),
        // A watermark copied from a header is not pasted into the text (it would be a second
        // shape with the same ids, and it belongs in the header).
        transformPasted: (slice) => withoutWatermarks(this.adoptPastedLists(slice)),
        clipboardSerializer: this.clipboardSerializer,
      });
    } catch (err) {
      host.remove();
      this.model = before.model;
      this.pushed = before.pushed;
      this.openedLast = before.openedLast;
      this.revisionIds = before.revisionIds;
      this.bodyStyle = before.bodyStyle;
      throw err;
    }

    this.view?.destroy();
    this.host.remove();
    this.host = host;
    this.view = view;
    // The document's styles apply to this editor only (a version preview may be open next to it).
    this.docCss = model.css
      .split('\n')
      .map((rule) => {
        const open = rule.indexOf('{');
        if (open < 0) return rule;
        const selectors = rule.slice(0, open).split(',').map((sel) => (sel.trim().startsWith('.dx-doc') ? `.${this.scope} ${sel.trim()}` : sel));
        return selectors.join(',') + rule.slice(open);
      })
      .join('\n');
    for (const font of eastAsiaFontsInCss(model.css)) useEastAsiaFont(font);
    this.syncFontCss();
    this.pages = [];
    this.pageEls = [];
    this.pagesLayer.replaceChildren();
    this.hfHeights.clear();
    this.lastCursorPage = 1;
    registerImageSources(this, imageSources(view.state.doc, model));
    this.baseline = this.savePoint();
    this.renderPages([]);
    this.emitUpdate();
  }

  /**
   * Applies a transaction to a view (as a tracked change while 追蹤修訂 is on). A plugin or
   * callback that fails leaves the view on its last good state, redrawn, and tells the user.
   */
  private safeDispatch(view: EditorView, sent: Transaction, after: (tr: Transaction) => void): void {
    try {
      // With 追蹤修訂 on, the edit as a tracked change (null: it can't be recorded, refused).
      const tr = trackTransaction(view.state, sent);
      if (!tr) return this.emitUpdate();
      view.updateState(view.state.apply(tr));
      after(tr);
      this.emitUpdate();
    } catch (err) {
      console.error('papyrus: edit failed', err);
      try {
        view.updateState(view.state);
      } catch {
        // the view can't even be redrawn: leave it
      }
      this.options.onNotice?.('這項編輯無法完成，文件保持在上一個狀態。');
    }
  }

  // ----- pasting pictures (persona-300 A-9) -----
  // Pasted HTML may only show pictures it carries (data: URLs): a Word clipboard's
  // file:///…clip_image or a web page's https:// picture is never fetched. When the clipboard
  // also has the picture files, one for each such picture (a picture copied from Word or from a
  // browser), those are used in their places; otherwise the pictures are left out and the user
  // is told how many, and how to add them.

  /** The picture files on the clipboard of the paste being handled (see notePaste). */
  private pasteFiles: File[] | null = null;
  /** The paste being handled uses these files for its pictures (see safePastedHtml). */
  private pasteSlots: { token: string; files: File[] } | null = null;
  /** Pictures the paste being handled leaves out. */
  private pasteDropped = 0;

  /** A paste starts: its picture files, for safePastedHtml (called before it). */
  private notePaste(event: Event): boolean {
    this.pasteFiles = imageFiles((event as ClipboardEvent).clipboardData);
    this.pasteSlots = null;
    this.pasteDropped = 0;
    return false;
  }

  /**
   * Pasted or dropped HTML: pictures that aren't pictures (remote or unreadable) become their
   * alt text, or placeholders for the clipboard's picture files (see onPaste); Excel's cell
   * styles and column widths are written on the cells (pasteExcel.ts).
   */
  private safePastedHtml(html: string): string {
    // Excel keeps its cells' look in a style block and its column widths in <col>: on the cells.
    html = inlineExcelStyles(html);
    const files = this.pasteFiles ?? [];
    this.pasteFiles = null;
    const unsafe = unsafeImageCount(html);
    if (!unsafe) return html;
    if (files.length === unsafe) {
      const token = Math.random().toString(36).slice(2, 10);
      this.pasteSlots = { token, files };
      return replaceUnsafeImages(html, (img, i) => {
        const size = ['width', 'height']
          .map((a) => (/^\d+(\.\d+)?$/.test(img.getAttribute(a) ?? '') ? ` ${a}="${img.getAttribute(a)}"` : ''))
          .join('');
        const alt = (img.getAttribute('alt') ?? '').replace(/[&<>"]/g, (c) => `&#${c.charCodeAt(0)};`);
        return `<img src="${pasteSlotSrc(token, i)}" alt="${alt}"${size}>`;
      });
    }
    this.pasteDropped += unsafe;
    return replaceUnsafeImages(html);
  }

  /** Ctrl+V, the ribbon's 貼上 and the browser's paste. */
  private onPaste(view: EditorView, event: ClipboardEvent, slice: Slice): boolean {
    const slots = this.pasteSlots;
    const dropped = this.pasteDropped;
    this.pasteSlots = null;
    this.pasteDropped = 0;
    this.pasteFiles = null;
    if (slots) {
      void this.pasteWithPictures(view, slice, slots);
      return true;
    }
    if (dropped) this.options.onNotice?.(droppedPicturesNotice(dropped));
    const data = event.clipboardData;
    // With HTML, the HTML is pasted, as the ribbon's 貼上 does: Excel puts a picture of the copied
    // range next to its table, which must not win over the cells (persona-300 B-9).
    if (data && Array.from(data.types ?? []).includes('text/html') && data.getData('text/html').trim()) return false;
    return this.handleFiles(data);
  }

  /** Something dropped on the page: picture files are inserted (a picture dragged from a web page too). */
  private onDrop(event: DragEvent): boolean {
    const dropped = this.pasteDropped;
    this.pasteDropped = 0;
    this.pasteSlots = null;
    const handled = this.handleFiles(event.dataTransfer);
    if (dropped && !imageFiles(event.dataTransfer).length) this.options.onNotice?.(droppedPicturesNotice(dropped));
    return handled;
  }

  /**
   * Pastes `slice` once the clipboard's picture files are read, each in its placeholder's place,
   * into the part the paste was made in (the body, or the header / footer being edited then),
   * even when the cursor moved to another part meanwhile. A header / footer closed meanwhile
   * gets nothing, and the user is told.
   */
  private async pasteWithPictures(view: EditorView, slice: Slice, slots: { token: string; files: File[] }): Promise<void> {
    const pictures = await Promise.all(slots.files.map((f) => this.readPicture(f).catch(() => null)));
    if (!this.view) return; // the editor is gone
    if (view.isDestroyed) {
      this.options.onNotice?.(PASTE_PART_CLOSED);
      return;
    }
    let missing = 0;
    const maxWidth = this.contentWidth;
    const map = (fragment: Fragment): Fragment => {
      const out: PMNode[] = [];
      fragment.forEach((node) => {
        const k = node.type === schema.nodes.image ? pasteSlotIndex(node.attrs.src, slots.token) : -1;
        if (k < 0) {
          out.push(node.isLeaf ? node : node.copy(map(node.content)));
          return;
        }
        let pic = pictures[k];
        // Not the copied picture: Word puts a picture of the whole selection (text and all) on the
        // clipboard when text is copied with one picture. Its proportions differ from the picture's.
        if (pic && !sameProportions(node.attrs.width, node.attrs.height, pic.width, pic.height)) pic = null;
        if (!pic) {
          missing++;
          if (node.attrs.alt) out.push(schema.text(node.attrs.alt, node.marks));
          return;
        }
        // The size the copied picture had, else its own (no wider than the text).
        let width: number = node.attrs.width || pic.width;
        let height: number = node.attrs.height || (node.attrs.width ? Math.round((node.attrs.width * pic.height) / pic.width) : pic.height);
        if (width > maxWidth && maxWidth > 0) {
          height = Math.round((height * maxWidth) / width);
          width = Math.round(maxWidth);
        }
        out.push(schema.nodes.image.create({ ...node.attrs, src: pic.src, width, height }, null, node.marks));
      });
      return Fragment.from(out);
    };
    const content = map(slice.content);
    const tr = view.state.tr.replaceSelection(new Slice(content, slice.openStart, slice.openEnd));
    view.dispatch(tr.scrollIntoView().setMeta('paste', true).setMeta('uiEvent', 'paste'));
    if (missing) this.options.onNotice?.(droppedPicturesNotice(missing));
  }

  /**
   * Pastes HTML as Ctrl+V does, with the picture files that came with it (the ribbon's 貼上 reads
   * the clipboard itself).
   */
  pasteHtml(html: string, files: File[] = []): void {
    const view = this.activeView;
    if (!view) return;
    view.focus();
    this.pasteFiles = files.filter((f) => f.type.startsWith('image/'));
    this.pasteSlots = null;
    this.pasteDropped = 0;
    view.pasteHTML(html);
  }

  /** Copied list paragraphs say which document and list they come from (see pasteLists.ts). */
  private clipboardSerializer = listClipboardSerializer(() => this.model.numbering);

  /** Turns pasted list placeholders into lists of this document (see pasteLists.ts). */
  adoptPastedLists(slice: Slice): Slice {
    return adoptPastedLists(slice, this.model.numbering);
  }

  /**
   * The body column's padding: the first page's top, the last page's bottom and — when all
   * sections share them — the side margins (otherwise each section's blocks get their own,
   * see pagination.ts).
   */
  private bodyPadding(doc: PMNode): string {
    const layout = this.layout(doc);
    const geos = layout.sections.map((s) => s.geometry);
    const side = columnBase(layout);
    return `padding:${geos[0].textTop}px ${side.right}px ${geos[geos.length - 1].textBottom}px ${side.left}px`;
  }

  /** Apply a changed body padding; returns whether it changed. */
  private syncBodyStyle(): boolean {
    const view = this.view;
    if (!view) return false;
    const style = this.bodyPadding(view.state.doc);
    if (style === this.bodyStyle) return false;
    this.bodyStyle = style;
    view.dom.setAttribute('style', style);
    repaginate(view);
    return true;
  }

  /**
   * Pictures pasted or dropped as files are inserted. Other files alone (a PDF ...) are refused
   * with a notice; pasted or dropped text that comes with them is handled as usual.
   */
  private handleFiles(data: DataTransfer | undefined | null): boolean {
    const files = Array.from(data?.files ?? []);
    const images = files.filter((f) => f.type.startsWith('image/'));
    if (!images.length) {
      const types = Array.from(data?.types ?? []);
      if (!files.length || types.includes('text/plain') || types.includes('text/html')) return false;
      this.options.onNotice?.(NOT_AN_IMAGE);
      return true;
    }
    (async () => {
      for (const f of images) await this.insertImageFile(f);
    })();
    return true;
  }

  private plugins(body: boolean, textBox = false): Plugin[] {
    const listIn: Command = (state, dispatch, view) => inList(state) && indent(1)(state, dispatch, view);
    const listOut: Command = (state, dispatch, view) => inList(state) && indent(-1)(state, dispatch, view);
    const keys: Record<string, Command> = {
      // Through the editor: one Ctrl+Z also undoes a change made in several parts (全部取代).
      'Mod-z': this.historyCommand('undo'),
      'Mod-y': this.historyCommand('redo'),
      'Mod-Shift-z': this.historyCommand('redo'),
      // Styles are looked up when the key is pressed: the model changes when a file is opened.
      'Mod-b': (s, d, v) => toggleFormat('bold', this.model.styles)(s, d, v),
      'Mod-i': (s, d, v) => toggleFormat('italic', this.model.styles)(s, d, v),
      'Mod-u': toggle('underline'),
      'Mod-Shift-x': toggle('strike'),
      'Mod-e': (s, d, v) => setAlign('center', this.model.styles)(s, d, v),
      'Mod-l': (s, d, v) => setAlign('left', this.model.styles)(s, d, v),
      'Mod-r': (s, d, v) => setAlign('right', this.model.styles)(s, d, v),
      'Mod-j': (s, d, v) => setAlign('justify', this.model.styles)(s, d, v),
      'Mod-Space': clearFormatting,
      // Word's font size keys: Ctrl+] / Ctrl+[ one point, Ctrl+Shift+> / < one step of the size list.
      'Mod-]': (s, d, v) => growFont(1, this.model.styles)(s, d, v),
      'Mod-[': (s, d, v) => growFont(-1, this.model.styles)(s, d, v),
      'Mod->': (s, d, v) => growFont('up', this.model.styles)(s, d, v),
      'Mod-<': (s, d, v) => growFont('down', this.model.styles)(s, d, v),
      // The same keys by their place on the keyboard (prosemirror-keymap looks them up by key
      // code when the character is not found): layouts where Shift+. / Shift+, is not > / <.
      'Mod-Shift-.': (s, d, v) => growFont('up', this.model.styles)(s, d, v),
      'Mod-Shift-,': (s, d, v) => growFont('down', this.model.styles)(s, d, v),
      // Word's 複製格式 shortcuts: copy / paste formatting. Ctrl+Shift+V pastes only what
      // Ctrl+Shift+C copied (until a Ctrl+C / Ctrl+X); otherwise it is left to the browser, whose
      // Ctrl+Shift+V pastes as plain text.
      'Mod-Shift-c': (_s, _d, v) => !!v && this.painter.copy(v),
      'Mod-Shift-C': (_s, _d, v) => !!v && this.painter.copy(v),
      'Mod-Shift-v': (_s, _d, v) => !!v && this.options.editable !== false && this.painter.paste(v),
      'Mod-Shift-V': (_s, _d, v) => !!v && this.options.editable !== false && this.painter.paste(v),
      // Word's shortcut for Review > Track Changes.
      'Mod-Shift-e': () => this.setTrackChanges(!this.trackOn),
      'Mod-Shift-E': () => this.setTrackChanges(!this.trackOn),
      Enter: splitParagraph,
      'Shift-Enter': insertHardBreak,
      // Tab types a tab, as in Word; right after Escape it leaves the document instead, so a
      // keyboard user has a way out forwards too (GOV-ISSUE-013; the same convention as
      // CodeMirror and other web editors). Returning false lets the browser move the focus.
      Tab: (st, d, v) => (this.leaveOnTab() ? false : chainCommands(goToNextCell(1), tabNewRow, listIn, insertTab)(st, d, v)),
      'Shift-Tab': (st, d, v) => (this.leaveOnTab() ? false : chainCommands(goToNextCell(-1), listOut)(st, d, v)),
    };
    // Word's Ctrl+Shift+Enter: in a table 分割表格 (Split Table), elsewhere a column break (body
    // only). Not in read-only mode.
    const splitTable: Command = (s, d, v) => this.options.editable !== false && splitTableCommand(s, d, v);
    keys['Mod-Shift-Enter'] = splitTable;
    if (body) {
      keys['Mod-Enter'] = insertPageBreak;
      keys['Mod-Shift-Enter'] = chainCommands(splitTable, insertColumnBreak);
      keys.Escape = () => {
        this.escapePressed = true;
        this.options.onNotice?.('已離開文字輸入：按 Tab 移到下一個控制項，按其他鍵繼續編輯。');
        return true;
      };
    } else {
      keys.Escape = () => {
        this.closeHeaderFooter();
        return true;
      };
    }
    // A header's watermark stays when the header's text is typed over (see keepWatermarks).
    const watermarks = body ? [] : [keepWatermarks((m) => this.options.onNotice?.(m))];
    const plugins: Plugin[] = [
      history(),
      // Before listMarkers: undo / redo of a paste takes its list definitions out / back first.
      madeListsSync(
        () => this.model.numbering,
        () => [...(this.view ? [this.view.state.doc] : []), ...this.model.headerFooters.map((hf) => hf.doc)],
      ),
      trackChanges({
        enabled: () => this.trackOn && this.options.editable !== false,
        author: () => this.author.name,
        onNotice: (message) => this.options.onNotice?.(message),
        ids: this.revisionIds,
      }),
      // An assistant's grey completion after the cursor and its suggested changes drawn in the text
      // (editor/assistantInline.ts). Before the key bindings: Tab and Escape are the completion's while one shows.
      ...(body ? [completionPlugin(), suggestionPlugin()] : []),
      // Any key other than Escape / Tab / a modifier ends "Escape, then Tab leaves".
      new Plugin({
        props: {
          handleKeyDown: (_view, event) => {
            if (!['Escape', 'Tab', 'Shift', 'Control', 'Alt', 'Meta'].includes(event.key)) this.escapePressed = false;
            return false;
          },
        },
      }),
      // Before the keymaps: Escape ends 複製格式 before it leaves the text or the header/footer.
      this.painter.plugin(),
      // A selected shape takes the arrow keys, Tab, Esc and Delete (before the keys above).
      ...(textBox ? [] : [this.shapes.plugin()]),
      // Enter on a selected text box (or a shape with text) edits its text.
      keymap({
        Enter: (state) => {
          const sel = state.selection;
          const shape = sel instanceof NodeSelection ? sel.node.attrs.shape : null;
          const first = shape ? textFrames(shape)[0] : null;
          return !!first && this.editShapeText(shape.key, first.i);
        },
      }),
      keymap(keys),
      keymap(baseKeymap),
      dropCursor(),
      gapCursor(),
      // Before columnResizing: it takes the press on a column border (drags as Word does).
      columnResize({ cellMinWidth: 24, zoom: () => this.zoomValue }),
      columnResizing({ cellMinWidth: 24 }),
      tableEditing(),
      listMarkers(() => this.model.numbering),
      tableStyleClasses(),
      fieldPlaceholders(),
      lockedFields((f) => this.options.onNotice?.(`「${f.title}」已鎖定${f.locked ? '，內容不能修改' : '，欄位不能刪除'}。`)),
      searchPlugin(),
      review(),
      runLanguages(),
      tabStops(),
      shapeLayout(() => this.model?.numbering ?? null),
      shapeKeys(),
      formattingMarks(() => this.marksOn),
      ...watermarks,
    ];
    if (body) {
      plugins.push(
        pagination(
          (doc) => this.layout(doc),
          (pages) => this.renderPages(pages),
          (sections) => this.setSingleColumn(sections),
        ),
        sectionColumns((doc) => this.layout(doc)),
        docGrid((doc) => this.sections(doc)),
        sectionMarks(),
        // Where the body was edited lately: what the assistant reads first (editor/assistant.ts).
        attentionPlugin(),
        // Ctrl+Z / Ctrl+Y of a page setup change don't take away (or bring back) the references to
        // headers/footers made or unlinked since, which are not undo steps.
        keepSectionReferences((relId) =>
          !!this.model.relIds?.includes(relId) || this.model.headerFooters.some((h) => h.relId === relId)),
      );
    }
    return plugins;
  }

  /**
   * Draws the pages behind the body: each page, its header and footer (with live page numbers)
   * and the covers where table rows break. Page elements and header/footer boxes are kept from
   * the last drawing and changed only where the layout or the content changed.
   */
  private renderPages(pages: PageBox[]): void {
    if (!pages.length) pages = [this.firstPage()];
    this.pages = pages;
    this.pageCount = pages.length;
    const sections = this.sections();
    const layer = this.pagesLayer;
    const s = this.hfSession;
    while (this.pageEls.length > pages.length) this.pageEls.pop()!.el.remove();
    // No picture may be taller than the smallest text area (set before headers are measured).
    const content = Math.max(40, Math.min(...pages.map((p) => p.height - p.textTop - p.textBottom)));
    this.root.style.setProperty('--dx-content-h', `${content}px`);
    const measured: { kind: HeaderFooterKind; section: number; height: () => number }[] = [];
    pages.forEach((pg, i) => {
      let page = this.pageEls[i];
      if (!page) {
        page = { el: el('div', 'dx-page'), geometry: '', boxes: {} };
        this.pageEls.push(page);
        layer.append(page.el);
      }
      const geometry = `top:${pg.top}px;left:${pg.left}px;width:${pg.width}px;height:${pg.height}px`;
      if (page.geometry !== geometry) page.el.style.cssText = page.geometry = geometry;
      // Shapes placed from the page's margins read them here (editor/shapeView.ts).
      page.el.dataset.margins = pageMargins(pg);
      const section = sections[pg.section] ?? sections[sections.length - 1];
      const type = variantFor(section, pg.inSection, pg.number, this.model.evenAndOdd);
      for (const kind of ['header', 'footer'] as const) {
        const hf = sectionHeaderFooter(this.model, sections, section.index, kind, type);
        let box = page.boxes[kind];
        // The page being edited shows the live editor instead of a static copy.
        if (!hf || (s && s.pageIndex === i && s.hf === hf)) {
          box?.el.remove();
          delete page.boxes[kind];
          continue;
        }
        const marks = this.marksOn;
        const template = hfTemplate(hf.doc, marks);
        if (!box || box.doc !== hf.doc || box.marks !== marks) {
          const boxEl = el('div', `dx-${kind} dx-doc${marks ? ` ${SHOW_MARKS_CLASS}` : ''}`);
          boxEl.append(template.content.cloneNode(true));
          if (box) box.el.replaceWith(boxEl);
          else if (kind === 'header') page.el.insertBefore(boxEl, page.boxes.footer?.el ?? null);
          else page.el.append(boxEl);
          box = page.boxes[kind] = { el: boxEl, doc: hf.doc, style: '', fields: '', marks };
          if (template.fields) box.fields = '\0'; // not filled in yet
        }
        const style = this.hfPosition(kind, section.page);
        if (box.style !== style) box.el.style.cssText = box.style = style;
        if (template.fields) {
          const values = `${pg.numberFormat ?? ''}|${pg.number}|` + LIVE_FIELDS.map((k) => this.fieldValue(k, i)).join('|');
          if (box.fields !== values) {
            this.fillFields(box.el, i);
            box.fields = values;
          }
        }
        const b = box;
        measured.push({ kind, section: pg.section, height: () => this.hfHeight(pg.section, kind, type, b, content) });
      }
      this.renderWatermark(page, pg, section, sectionHeaderFooter(this.model, sections, section.index, 'header', type));
    });
    this.renderCuts(pages);
    this.renderColumnRules(pages);
    this.renderHeaderFooterSummary();
    if (s) {
      const area = s.box.querySelector('.dx-hf-area') as HTMLElement;
      measured.push({ kind: s.hf.kind, section: pages[s.pageIndex]?.section ?? 0, height: () => area.offsetHeight });
    }
    const width = Math.max(...pages.map((p) => p.left + p.width));
    const last = pages[pages.length - 1];
    this.canvas.style.width = `${width}px`;
    this.canvas.style.minHeight = `${last.top + last.height}px`;
    this.applyZoom();
    for (const f of this.fieldViews) f.render();
    this.updateTextArea(measured, sections);
    // Shapes placed from the page (body, header and footer copies), and a text box being edited.
    if (placeShapes(this.canvas, this.model.numbering) && this.view) repaginate(this.view);
    this.shapeSession?.place();
    this.shapes.draw();
    this.emitUpdate();
  }

  /**
   * The text of every header and footer the pages show, once each, for screen readers (the
   * pages are hidden from them): 「頁首：臺北市政府 函。頁尾：第 1 頁」, with 「（首頁）」
   * 「－第 2 節」 … where the document has those.
   */
  private renderHeaderFooterSummary(): void {
    const parts = this.shownHeaderFooters()
      .map(({ hf, label }) => ({ label, text: hf.doc.textBetween(0, hf.doc.content.size, ' ', ' ').replace(/\s+/g, ' ').trim() }))
      .filter((p) => p.text)
      .map((p) => `${p.label}：${p.text}`);
    const text = parts.length ? parts.join('。') + '。' : '';
    if (this.hfSummary.textContent !== text) this.hfSummary.textContent = text;
  }

  /**
   * The page's 浮水印: the watermark of the header the page shows (Word draws it on each page of
   * that header, a first-page header without one shows none), on a layer behind the header, the
   * footer and the text. Drawn again only when it or the page changed.
   */
  private renderWatermark(page: DrawnPage, pg: PageBox, section: Section, hf: HeaderFooterPart | undefined): void {
    const wm = hf && !hf.pending ? docWatermark(hf.doc) : null;
    const p = section.page;
    const box = {
      width: pg.width,
      height: pg.height,
      top: twipsToPx(p.marginTop),
      right: twipsToPx(p.marginRight),
      bottom: twipsToPx(p.marginBottom),
      left: twipsToPx(p.marginLeft),
    };
    // A text is drawn by its font's measurements: drawn again once the font has loaded.
    const metrics = wm?.shape.kind === 'text' ? measureWatermarkText(wm.shape.text, wm.shape.font) : null;
    const key = wm ? JSON.stringify([wm.shape, box, metrics]) : '';
    if ((page.watermark?.key ?? '') === key && page.watermark?.src === wm?.src) return;
    page.watermark?.el.remove();
    delete page.watermark;
    const el = wm ? watermarkLayer(wm, box, metrics) : null;
    if (!wm || !el) return;
    page.el.prepend(el);
    page.watermark = { key, src: wm.src, el, measured: wm.shape.kind !== 'text' || !!metrics };
  }

  /** Live values of a static header/footer's fields (a field result split over runs shows once). */
  private fillFields(box: HTMLElement, pageIndex: number): void {
    let last: HTMLElement | null = null;
    for (const f of Array.from(box.querySelectorAll<HTMLElement>('[data-field]'))) {
      const value = this.fieldValue(f.dataset.field ?? '', pageIndex, f.dataset.format ?? null);
      if (value == null) continue;
      const continues = last && last.nextElementSibling === f && last.dataset.field === f.dataset.field;
      f.textContent = continues ? '' : value;
      last = f;
    }
  }

  /**
   * How tall a static header/footer is: measured on the first page showing it, then kept per
   * (section, kind, variant) until the header/footer, its position, the pictures' size limit or
   * the fonts change.
   */
  private hfHeight(section: number, kind: HeaderFooterKind, type: HeaderFooterType, box: HeaderFooterBox, contentHeight: number): number {
    const key = `${section}|${kind}|${type}`;
    const hit = this.hfHeights.get(key);
    if (hit && hit.doc === box.doc && hit.style === box.style && hit.contentHeight === contentHeight) return hit.height;
    const height = box.el.offsetHeight;
    // Fonts still loading: measured again next time.
    if (document.fonts?.status !== 'loading') this.hfHeights.set(key, { doc: box.doc, style: box.style, contentHeight, height });
    return height;
  }

  /** The lines between text columns (w:sep) on each page, as laid out (editor/columnLayout.ts). */
  private renderColumnRules(pages: PageBox[]): void {
    pages.forEach((pg, i) => {
      const page = this.pageEls[i]?.el;
      if (page) drawColumnRules(page, pg.rules);
    });
  }

  /**
   * Where a table row breaks across pages, its cells' borders would run through the footer, the
   * gap and the next header: that band is covered (page colour, desk colour in the gap), with the
   * row's bottom edge on one page and its top edge on the next. Headers and footers stay on top.
   */
  private renderCuts(pages: PageBox[]): void {
    const cuts: HTMLElement[] = [];
    pages.forEach((pg, i) => {
      const next = pages[i + 1];
      for (const c of pg.cuts ?? []) {
        const band = el('div', 'dx-cut');
        const height = Math.max(0, c.bottom - c.top);
        const gapStart = Math.max(0, pg.top + pg.height - c.top);
        const gapEnd = next ? Math.min(height, next.top - c.top) : height;
        band.style.cssText =
          `top:${c.top}px;left:${c.left - 1}px;width:${c.right - c.left + 2}px;height:${height}px;` +
          `border-top:${c.border};border-bottom:${c.border};` +
          `background:linear-gradient(var(--dx-page) ${gapStart}px,var(--dx-desk) ${gapStart}px ${gapEnd}px,var(--dx-page) ${gapEnd}px)`;
        cuts.push(band);
        const header = c.header && this.headerCopy(c.header.table, c.header.rows);
        if (header) {
          header.style.cssText = `top:${c.bottom}px;left:${c.left}px;width:${c.right - c.left}px`;
          cuts.push(header);
        }
      }
    });
    this.cutsLayer.replaceChildren(...cuts);
  }

  /** A copy of a table's header rows (with its columns and style), shown above a row's continuation. */
  private headerCopy(tablePos: number, rows: number[]): HTMLElement | null {
    const view = this.view;
    const table = (view?.nodeDOM(tablePos) as HTMLElement | null)?.querySelector('table');
    if (!view || !table) return null;
    const copy = table.cloneNode(false) as HTMLElement;
    const cols = table.querySelector('colgroup');
    if (cols) copy.append(cols.cloneNode(true));
    const body = document.createElement('tbody');
    for (const pos of rows) {
      const tr = view.nodeDOM(pos) as HTMLElement | null;
      if (tr) body.append(tr.cloneNode(true));
    }
    copy.append(body);
    copy.style.width = '100%';
    // Inside .dx-doc and, below it, the table's style class, as in the body: the document's table
    // style rules (".dx-doc :where(.dx-ts-X) .dx-p") only apply there. With both classes on one
    // element, the header rows took the default paragraph spacing, came out taller than the rows
    // the page was laid out with and covered the first lines of the row's continuation
    // (GOV-ISSUE-017). The table's own cell margins (custom properties on its wrapper) come
    // along; the wrapper's section padding does not, the copy is placed at the row's left edge.
    // The rows' formatting marks show here too when they do in the body (顯示／隱藏編輯標記).
    const wrap = el('div', `dx-cut-header dx-doc${this.marksOn ? ` ${SHOW_MARKS_CLASS}` : ''}`);
    const source = table.parentElement;
    const tableWrap = el('div', source?.className ?? '');
    if (source) {
      for (const prop of Array.from(source.style)) {
        if (prop.startsWith('--')) tableWrap.style.setProperty(prop, source.style.getPropertyValue(prop));
      }
    }
    tableWrap.append(copy);
    wrap.append(tableWrap);
    return wrap;
  }

  /** Word moves the body down (or up) when a header (or footer) is taller than its margin; per section. */
  private updateTextArea(boxes: { kind: HeaderFooterKind; section: number; height: () => number }[], sections: Section[]): void {
    const next = new Map<number, { top: number; bottom: number }>();
    for (const b of boxes) {
      const sec = sections[b.section];
      if (!sec) continue;
      const push = next.get(b.section) ?? { top: 0, bottom: 0 };
      const height = b.height();
      if (b.kind === 'header') push.top = Math.max(push.top, twipsToPx(sec.page.header) + height);
      else push.bottom = Math.max(push.bottom, twipsToPx(sec.page.footer) + height);
      next.set(b.section, push);
    }
    const same =
      next.size === this.pushed.size &&
      [...next].every(([k, v]) => {
        const o = this.pushed.get(k);
        return !!o && Math.abs(o.top - v.top) < 0.5 && Math.abs(o.bottom - v.bottom) < 0.5;
      });
    if (same) return;
    this.pushed = next;
    if (this.view && !this.syncBodyStyle()) repaginate(this.view);
  }

  snapshot(): EditorSnapshot | null {
    return this.makeSnapshot(false);
  }

  private makeSnapshot(deferPage: boolean): EditorSnapshot | null {
    const view = this.activeView;
    if (!view) return null;
    const state = view.state;
    const { from, to, empty, $from } = state.selection;
    const here = state.storedMarks ?? $from.marks();
    const has = (name: string) => {
      const type = schema.marks[name];
      return empty ? !!type.isInSet(here) : state.doc.rangeHasMark(from, to, type);
    };
    const attr = (name: string, key: string) => schema.marks[name].isInSet(here)?.attrs[key] ?? null;
    const para = $from.parent.type === schema.nodes.paragraph ? $from.parent : null;
    // What is in effect, styles included (a style's 18 pt blue shows as 18 pt blue).
    const size = selectionValue(state, 'fontSize', this.model.styles);
    const color = selectionValue(state, 'color', this.model.styles);
    const font = selectionValue(state, 'font', this.model.styles);
    const sections = this.sections();
    const section = this.cursorSection();
    let list: EditorSnapshot['list'] = null;
    if (para?.attrs.numId) {
      const n = this.model.numbering;
      const abs = n.abstracts[n.nums[para.attrs.numId]?.abstractId ?? ''];
      list = isListKind(abs, 'bullet') ? 'bullet' : isListKind(abs, 'gongwen') ? 'gongwen' : 'decimal';
    }
    return {
      // What is in effect, styles included (a heading style's bold shows as bold).
      bold: selectionHas(state, 'bold', this.model.styles),
      italic: selectionHas(state, 'italic', this.model.styles),
      underline: has('underline'),
      strike: has('strike'),
      superscript: has('superscript'),
      subscript: has('subscript'),
      link: attr('link', 'href'),
      color: color.value as string | null,
      highlight: attr('highlight', 'color'),
      fontFamily: font.value as string | null,
      fontSize: size.value as number | null,
      mixed: { fontSize: size.mixed, color: color.mixed, fontFamily: font.mixed },
      styleId: para?.attrs.styleId ?? null,
      align: para ? effectiveAlign(para.attrs, this.model.styles) : null,
      list,
      inTable: isInTable(state),
      canUndo: undoDepth(state) > 0 || this.groupApplies('undo', view),
      canRedo: redoDepth(state) > 0 || this.groupApplies('redo', view),
      pageCount: this.pageCount,
      currentPage: deferPage ? this.deferredCursorPage() : (this.lastCursorPage = this.cursorPage()),
      zoom: this.zoomValue,
      sectionCount: sections.length,
      section: section.index,
      atSectionBreak: !this.hfSession && $from.depth >= 1 && !!state.doc.nodeAt($from.before(1))?.attrs.sectPr,
      target: this.target,
      titlePage: section.titlePage,
      trackChanges: this.trackOn,
      formatPainter: this.painter.mode,
      showMarks: this.marksOn,
    };
  }

  private emitUpdate(): void {
    const snap = this.makeSnapshot(true);
    if (snap) this.options.onUpdate?.(snap);
  }
}

/** Friendly names of package parts, for notices. */
function partName(part: string): string {
  const file = part.split('/').pop() ?? part;
  if (/^header\d*\.xml$/.test(file)) return '頁首';
  if (/^footer\d*\.xml$/.test(file)) return '頁尾';
  if (file === 'styles.xml') return '樣式';
  if (file === 'numbering.xml') return '編號';
  if (file === 'settings.xml') return '文件設定';
  if (file.endsWith('.rels')) return '關聯';
  return file;
}

/** The picture files of a paste or drop. */
function imageFiles(data: DataTransfer | null | undefined): File[] {
  return Array.from(data?.files ?? []).filter((f) => f.type.startsWith('image/'));
}

/**
 * Word's 插入連結 (Ctrl+K, the ribbon's 連結): asks for the address, the selection's link shown,
 * and sets it on the selection (empty: the link is removed). Only http(s)://, mailto: and #
 * links (docx/links.ts). Returns whether the link changed. Works with the ribbon hidden too.
 */
export function promptLink(editor: DocxEditor | null | undefined): boolean {
  if (!editor?.activeView) return false;
  const current = editor.snapshot()?.link ?? '';
  const href = window.prompt('連結網址（留空 = 移除連結）', current || 'https://');
  if (href === null) return false;
  const clean = href.trim();
  if (clean && !hasExplicitSafeScheme(clean)) {
    window.alert('只支援 http://、https://、mailto: 或 # 開頭的連結');
    return false;
  }
  return editor.run(setLink(clean || null));
}

/** Shown when a paste with pictures could not go in: the header / footer it was made in was closed meanwhile. */
export const PASTE_PART_CLOSED = '頁首／頁尾已關閉，剛才貼上的內容沒有加入，請再貼上一次。';

/**
 * Whether a picture file (`w` × `h` px) can be the pasted picture shown at `width` × `height`:
 * the same proportions, within 10 % (yes when the HTML gives no size).
 */
export function sameProportions(width: number | null | undefined, height: number | null | undefined, w: number, h: number): boolean {
  if (!width || !height || !w || !h) return true;
  return Math.abs((width / height) / (w / h) - 1) <= 0.1;
}

/** Told when pasted pictures could not come along (persona-300 A-9). */
export function droppedPicturesNotice(n: number): string {
  return `有 ${n} 張圖片無法一起貼上，請用『插入 › 圖片』加入。`;
}

export function brokenPartsNotice(parts: string[]): string {
  const names = [...new Set(parts.map(partName))].join('、');
  return `這份文件有部分內容格式錯誤（${names}），已略過顯示，存檔時會原樣保留。`;
}

export function saveWarningsNotice(warnings: WriteWarning[]): string {
  const images = warnings.filter((w) => w.kind === 'image').length;
  const parts = [...new Set(warnings.filter((w) => w.kind === 'part').map((w) => partName(w.part ?? '')))];
  const out: string[] = [];
  if (images) out.push(`${images} 張圖片無法寫入，已保留原圖或略過`);
  if (parts.length) out.push(`部分內容（${parts.join('、')}）無法寫入新內容，已原樣保留`);
  return `已存檔，但${out.join('；')}。`;
}

/** A page's margins for placing shapes: "top,right,bottom,left" px (see editor/shapeView.ts). */
function pageMargins(p: PageGeometry): string {
  return [p.textTop, p.marginRight, p.textBottom, p.marginLeft].map((v) => Math.round(v * 100) / 100).join(',');
}

/** The find panel's id of a text box: the part holding it, its shape and which text box. */
function textBoxPartId(host: string, key: string, index: number): string {
  return `textbox|${host}|${key}|${index}`;
}

function parseTextBoxPartId(id: string): { host: string; key: string; index: number } | null {
  const m = /^textbox\|(.+)\|([^|]+)\|(\d+)$/.exec(id);
  return m ? { host: m[1], key: m[2], index: Number(m[3]) } : null;
}

/** A section's text area (between the margins) in pt: where its watermark is centred and fitted. */
function watermarkArea(section: Section): WatermarkArea {
  const p = section.page;
  return { width: (p.width - p.marginLeft - p.marginRight) / 20, height: (p.height - p.marginTop - p.marginBottom) / 20 };
}

function emptyDoc(): PMNode {
  return schema.nodes.doc.create(null, schema.nodes.paragraph.create());
}

function el(tag: string, className: string): HTMLElement {
  const e = document.createElement(tag);
  e.className = className;
  return e;
}


