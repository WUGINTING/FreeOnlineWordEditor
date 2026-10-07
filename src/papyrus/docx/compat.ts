// What in an opened .docx the web editor cannot fully show or edit.
//
// The editor keeps everything it doesn't understand and writes it back unchanged, but the
// user should know which content is only kept, not shown or editable, and what that means.
// The scan reads the package itself (not the editor's model) so it reports what the file
// really contains.

import type JSZip from 'jszip';
import { NS, REL_TYPE, attr, parseXml } from './xml';
import { findMainPart, readRels } from './reader';
import { COMMENTS_REL } from './comments';
import { isShownPageFormat } from './pageNumbers';

export interface CompatItem {
  id: string;
  /** What it is, e.g. 文字方塊. */
  title: string;
  count: number;
  /** How many / where, e.g. 內文 2 處、頁首頁尾 1 處. */
  where: string;
  /** What happens to it in the web editor. */
  effect: string;
}

export interface CompatReport {
  items: CompatItem[];
}

/** How to get the untouched file back (the page's version history keeps the upload). */
export const KEEP_ORIGINAL = '上傳的第 1 版就是原檔，可在「版本紀錄」下載；在網頁上存檔也不會刪除上列內容。';

const MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const VML = 'urn:schemas-microsoft-com:vml';
const OFFICE = 'urn:schemas-microsoft-com:office:office';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
const WPC = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas';
const GRAPHIC = {
  picture: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  chart: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  diagram: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
};
const REL_FOOTNOTES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes';
const REL_ENDNOTES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes';

/** Fields the editor shows live; every other field shows its cached result. */
const LIVE_FIELDS = new Set(['PAGE', 'NUMPAGES', 'SECTIONPAGES']);
/** Fields whose cached result is plain, editable text that never needs updating. */
const STATIC_FIELDS = new Set(['HYPERLINK', '']);
const FIELD_NAMES: Record<string, string> = {
  TOC: '目錄', REF: '交互參照', PAGEREF: '頁碼參照', NOTEREF: '註腳參照', SEQ: '圖表編號', DATE: '日期',
  TIME: '時間', CREATEDATE: '建立日期', SAVEDATE: '儲存日期', PRINTDATE: '列印日期', STYLEREF: '樣式參照',
  DOCPROPERTY: '文件屬性', AUTHOR: '作者', TITLE: '標題', FILENAME: '檔名', MERGEFIELD: '合併列印欄位',
  IF: '條件', FORMTEXT: '表單欄位', FORMCHECKBOX: '表單核取方塊', FORMDROPDOWN: '表單下拉清單', INDEX: '索引',
  TA: '引文', TOA: '引文目錄', CITATION: '引文', BIBLIOGRAPHY: '參考文獻', LISTNUM: '清單編號', '=': '公式計算',
};

type Story = 'body' | 'hf' | 'notes';
const STORY_NAME: Record<Story, string> = { body: '內文', hf: '頁首頁尾', notes: '註腳' };

/** Counts per story, turned into "內文 2 處、頁首頁尾 1 處". */
class Tally {
  private counts = new Map<string, Map<Story, number>>();
  add(id: string, story: Story, n = 1): void {
    if (n <= 0) return;
    let m = this.counts.get(id);
    if (!m) this.counts.set(id, (m = new Map()));
    m.set(story, (m.get(story) ?? 0) + n);
  }
  total(id: string): number {
    let t = 0;
    for (const n of this.counts.get(id)?.values() ?? []) t += n;
    return t;
  }
  where(id: string, unit = '處'): string {
    const m = this.counts.get(id);
    if (!m) return '';
    return (['body', 'hf', 'notes'] as Story[])
      .filter((s) => m.has(s))
      .map((s) => `${STORY_NAME[s]} ${m.get(s)} ${unit}`)
      .join('、');
  }
}

/** Visit every element, skipping mc:Fallback (the same content as its mc:Choice, for old readers). */
function walk(root: Element, visit: (el: Element) => boolean | void): void {
  const stack: Element[] = [root];
  while (stack.length) {
    const el = stack.pop()!;
    if (el.namespaceURI === NS.mc && el.localName === 'Fallback') continue;
    if (visit(el) === false) continue;
    for (let c = el.lastElementChild; c; c = c.previousElementSibling) stack.push(c);
  }
}

const isW = (el: Element | null, ...locals: string[]) => !!el && el.namespaceURI === NS.w && locals.includes(el.localName);
const on = (el: Element) => {
  const v = attr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off');
};

function hasAncestor(el: Element, test: (e: Element) => boolean, stopAt?: Element): boolean {
  for (let p = el.parentElement; p && p !== stopAt; p = p.parentElement) if (test(p)) return true;
  return false;
}

interface FieldTally {
  kinds: Map<string, number>;
}

function scanStory(root: Element, story: Story, t: Tally, fields: FieldTally, revisions: Map<string, number>, unsupportedRev: Tally): void {
  // Fields inside another field (PAGEREF in a TOC ...) are updated with it: only the outermost count.
  const fieldStack: { instr: string; done: boolean; nested: boolean }[] = [];
  const addField = (instr: string) => {
    const kind = (instr.trim().split(/\s+/)[0] || '').toUpperCase();
    if (LIVE_FIELDS.has(kind) || STATIC_FIELDS.has(kind)) return;
    fields.kinds.set(kind, (fields.kinds.get(kind) ?? 0) + 1);
    t.add('fields', story);
  };
  const rev = (k: string) => revisions.set(k, (revisions.get(k) ?? 0) + 1);

  walk(root, (el) => {
    const ns = el.namespaceURI;
    const local = el.localName;
    if (ns === NS.w) {
      switch (local) {
        case 'txbxContent':
          t.add('textbox', story);
          break;
        case 'object':
          t.add('ole', story);
          unsupportedRev.add('revisions', story, revisionsInside(el));
          return false;
        case 'altChunk':
          t.add('ole', story);
          break;
        case 'control':
          t.add('macros', story);
          break;
        case 'footnoteReference':
        case 'endnoteReference':
          t.add('notes', story);
          break;
        case 'sdt':
          t.add('contentControls', story);
          break;
        case 'lock': {
          const v = attr(el, 'val');
          if (v && v !== 'unlocked' && isW(el.parentElement, 'sdtPr')) t.add('lockedControls', story);
          break;
        }
        case 'framePr':
          t.add('frames', story);
          break;
        case 'bidi':
          if (isW(el.parentElement, 'pPr') && on(el)) t.add('rtl', story);
          break;
        case 'rtl':
          if (on(el)) t.add('rtl', story);
          break;
        case 'fldSimple':
          if (!fieldStack.length) addField(attr(el, 'instr') ?? '');
          break;
        case 'fldChar': {
          const type = attr(el, 'fldCharType');
          if (type === 'begin') fieldStack.push({ instr: '', done: false, nested: fieldStack.length > 0 });
          else if (type === 'separate' || type === 'end') {
            const f = fieldStack[fieldStack.length - 1];
            if (f && !f.done) {
              f.done = true;
              if (!f.nested) addField(f.instr);
            }
            if (type === 'end') fieldStack.pop();
          }
          break;
        }
        case 'instrText': {
          const f = fieldStack[fieldStack.length - 1];
          if (f && !f.done) f.instr += el.textContent ?? '';
          break;
        }
        // Tracked changes.
        case 'ins':
        case 'del':
        case 'moveFrom':
        case 'moveTo':
        case 'rPrChange':
        case 'pPrChange': {
          // Resolved with the deletion holding it (the editor lists that one only).
          if (hasAncestor(el, (p) => isW(p, 'del', 'moveFrom'), root)) break;
          const parent = el.parentElement;
          if (story === 'notes' || isW(parent, 'trPr', 'numPr') || inKeptContent(el, root)) unsupportedRev.add('revisions', story);
          else if (local === 'rPrChange' || local === 'pPrChange') rev('format');
          else if (isW(parent, 'rPr') && isW(parent!.parentElement, 'pPr')) rev('paraMark');
          else if (isW(parent, 'rPr')) {
            // Only paragraph marks carry these in rPr.
          } else rev(local === 'ins' || local === 'moveTo' ? 'ins' : 'del');
          break;
        }
        case 'cellIns':
        case 'cellDel':
        case 'cellMerge':
        case 'tblPrChange':
        case 'tblPrExChange':
        case 'trPrChange':
        case 'tcPrChange':
        case 'tblGridChange':
        case 'sectPrChange':
        case 'numberingChange':
          unsupportedRev.add('revisions', story);
          break;
      }
      return;
    }
    if (ns === MATH && (local === 'oMathPara' || (local === 'oMath' && !isM(el.parentElement, 'oMathPara')))) {
      t.add('equation', story);
      unsupportedRev.add('revisions', story, revisionsInside(el));
      return false;
    }
    if (ns === NS.a && local === 'graphicData') {
      const uri = el.getAttribute('uri') ?? '';
      if (uri === GRAPHIC.chart || uri.includes('/chartex')) t.add('chart', story);
      else if (uri === GRAPHIC.diagram) t.add('smartart', story);
      else if (uri === GRAPHIC.picture && isAnchored(el)) t.add('floatingImage', story);
      return;
    }
    if ((ns === WPS && local === 'wsp' && !hasAncestor(el, (p) => p.namespaceURI === WPG && p.localName === 'wgp')) ||
        (ns === WPG && local === 'wgp') || (ns === WPC && local === 'wpc')) {
      // A text box is reported as such; any other drawn shape as a shape.
      if (!el.getElementsByTagNameNS(WPS, 'txbx').length) t.add('shape', story);
      return;
    }
    if (ns === VML && ['shape', 'rect', 'roundrect', 'oval', 'line', 'polyline', 'arc', 'group'].includes(local)) {
      if (hasAncestor(el, (p) => p.namespaceURI === VML && p.localName === 'group')) return;
      if (hasAncestor(el, (p) => isW(p, 'object'))) return;
      // Word's 浮水印 (PowerPlusWaterMarkObject / WordPictureWatermark) is shown on the pages and
      // can be changed or removed with 設計 › 浮水印: not an unsupported shape.
      if (local === 'shape' && /^(?:PowerPlusWaterMarkObject|WordPictureWatermark)/.test(el.getAttribute('id') ?? '')) return;
      const picture = el.getElementsByTagNameNS(VML, 'imagedata').length > 0 && local !== 'group';
      const textbox = el.getElementsByTagNameNS(VML, 'textbox').length > 0;
      if (!picture && !textbox) t.add('shape', story);
      return;
    }
    if (ns === OFFICE && local === 'OLEObject') t.add('ole', story);
  });
}

/**
 * Inside content the editor keeps as-is (text boxes, drawings, objects, equations, alternate
 * content ...): what is there can't be seen or reviewed in the editor. That is anything inside
 * a run other than its own properties, or inside an equation or mc:AlternateContent.
 */
function inKeptContent(el: Element, stopAt: Element): boolean {
  for (let prev = el, p = el.parentElement; p && p !== stopAt; prev = p, p = p.parentElement) {
    if (p.namespaceURI === MATH || (p.namespaceURI === NS.mc && p.localName === 'AlternateContent') || isW(p, 'txbxContent')) return true;
    if (isW(p, 'r') && !isW(prev, 'rPr')) return true;
  }
  return false;
}

const REVISIONS = ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange'];

/** Tracked changes inside an element the scan doesn't walk into (those inside a deletion go with it). */
function revisionsInside(el: Element): number {
  let n = 0;
  for (const local of REVISIONS) {
    for (const r of Array.from(el.getElementsByTagNameNS(NS.w, local))) {
      if (!hasAncestor(r, (p) => isW(p, 'del', 'moveFrom'), el)) n++;
    }
  }
  return n;
}

function isM(el: Element | null, local: string): boolean {
  return !!el && el.namespaceURI === MATH && el.localName === local;
}

function isAnchored(el: Element): boolean {
  return hasAncestor(el, (p) => p.namespaceURI === NS.wp && p.localName === 'anchor');
}

// ----- sections -----

interface SectionTally {
  count: number;
  columns: number[];
  lineNumbers: number[];
  pageBorders: number[];
  vAlign: number[];
  vertical: number[];
  pageNumberFormat: number[];
}

function scanSections(body: Element): SectionTally {
  const out: SectionTally = { count: 0, columns: [], lineNumbers: [], pageBorders: [], vAlign: [], vertical: [], pageNumberFormat: [] };
  const sects = Array.from(body.getElementsByTagNameNS(NS.w, 'sectPr')).filter((s) => !isW(s.parentElement, 'sectPrChange'));
  sects.forEach((s, i) => {
    const n = i + 1;
    out.count++;
    const kid = (local: string) => Array.from(s.children).find((c) => isW(c, local)) ?? null;
    const cols = kid('cols');
    if (cols && (Number(attr(cols, 'num') ?? 1) > 1 || Array.from(cols.children).filter((c) => isW(c, 'col')).length > 1)) out.columns.push(n);
    if (kid('lnNumType')) out.lineNumbers.push(n);
    if (kid('pgBorders')) out.pageBorders.push(n);
    const va = attr(kid('vAlign'), 'val');
    if (va && va !== 'top') out.vAlign.push(n);
    const td = attr(kid('textDirection'), 'val');
    if (td && !['lrTb', 'tb', 'lrTbV'].includes(td)) out.vertical.push(n);
    const fmt = attr(kid('pgNumType'), 'fmt');
    // Roman, letters and the Chinese formats are shown like Word (docx/pageNumbers.ts): only
    // chapter numbers (1-1, 1-2 …) and rarer formats are listed.
    if (!isShownPageFormat(fmt) || attr(kid('pgNumType'), 'chapStyle')) out.pageNumberFormat.push(n);
  });
  return out;
}

function sectionWhere(list: number[], total: number): string {
  if (total <= 1) return '整份文件';
  if (list.length === total) return `全部 ${total} 節`;
  return `第 ${list.join('、')} 節（共 ${total} 節）`;
}

// ----- the scan -----

const LARGE_BYTES = 3 * 1024 * 1024;
const LARGE_PARAGRAPHS = 3000;
const BROWSER_UNSUPPORTED_IMAGES = /\.(emf|wmf|tif|tiff)$/i;

async function text(zip: JSZip, path: string): Promise<string | null> {
  const f = zip.file(path);
  return f ? f.async('string') : null;
}

function parse(xml: string | null): Element | null {
  if (!xml) return null;
  try {
    return parseXml(xml).documentElement;
  } catch {
    return null;
  }
}

/** Scan an opened package. A new (blank) document gives an empty report. */
export async function scanCompat(zip: JSZip | null): Promise<CompatReport> {
  if (!zip) return { items: [] };
  const main = await findMainPart(zip);
  const mainXml = await text(zip, main);
  const root = parse(mainXml);
  if (!root) return { items: [] };
  const rels = [...(await readRels(zip, main)).values()].filter((r) => !r.external);

  const t = new Tally();
  const unsupportedRev = new Tally();
  const fields: FieldTally = { kinds: new Map() };
  const revisions = new Map<string, number>();
  const body = Array.from(root.children).find((c) => isW(c, 'body')) ?? root;
  scanStory(body, 'body', t, fields, revisions, unsupportedRev);

  // Headers/footers: tracked changes there are handled while editing them, like the body's.
  for (const r of rels.filter((r) => r.type === REL_TYPE.header || r.type === REL_TYPE.footer)) {
    const el = parse(await text(zip, r.target));
    if (el) scanStory(el, 'hf', t, fields, revisions, unsupportedRev);
  }
  // Footnotes / endnotes: not editable here at all.
  for (const r of rels.filter((r) => r.type === REL_FOOTNOTES || r.type === REL_ENDNOTES)) {
    const el = parse(await text(zip, r.target));
    if (el) scanStory(el, 'notes', new Tally(), { kinds: new Map() }, new Map(), unsupportedRev);
  }

  const items: CompatItem[] = [];
  const add = (id: string, title: string, effect: string, where?: string, count?: number) => {
    const n = count ?? t.total(id);
    if (n > 0) items.push({ id, title, count: n, where: where ?? t.where(id), effect });
  };

  const KEPT = '原樣保留，存檔不會遺失';
  add('textbox', '文字方塊', `${KEPT}；網頁上顯示在 Word 的位置，按兩下可以編輯裡面的文字（可尋找取代、計入字數），存檔只改寫該文字方塊的文字；移動、調整大小要在 Word 中進行。`);
  add('shape', '圖案、繪圖物件', `${KEPT}；常見的圖案（方塊、流程圖、箭號、線條與接點）照 Word 的位置畫出，圖案裡的文字按兩下可以編輯；無法移動、縮放或改樣式。其他圖形以預覽圖或框線顯示，標示「這個圖形只能在 Word 修改」。`);
  add('smartart', 'SmartArt 圖形', `${KEPT}；網頁上以標籤或預覽圖顯示，無法修改文字與版面。`);
  add('chart', '圖表', `${KEPT}；網頁上以標籤或預覽圖顯示，無法修改資料或樣式。`);
  add('equation', '方程式', `${KEPT}；網頁上以「公式」標籤顯示，看不到也無法編輯公式內容。`);
  add('ole', '內嵌物件（Excel、PowerPoint 等 OLE 物件）', `${KEPT}；網頁上以標籤或預覽圖顯示，無法開啟或修改。`);
  add('floatingImage', '浮動（文繞圖）圖片', `${KEPT}；網頁上以預覽圖放在文字行中，位置與文繞圖效果和 Word 不同，也無法移動或調整大小。`);
  add('notes', '註腳／章節附註', `${KEPT}；內文中只顯示「註腳」標籤，註腳內容在網頁上看不到也無法編輯。`);

  const commentsRel = rels.find((r) => r.type === COMMENTS_REL);
  const commentsEl = commentsRel ? parse(await text(zip, commentsRel.target)) : null;
  const comments = commentsEl ? commentsEl.getElementsByTagNameNS(NS.w, 'comment').length : 0;
  add('comments', '註解（留言）', `${KEPT}；可在「留言」面板閱讀、跳到留言位置，也可以新增、回覆、標示為已解決或重新開啟，並編輯或刪除自己的留言（Word 中看得到）。留言內容在網頁上只能是純文字：編輯過的留言，原本的粗體等格式會變成一般文字。`, `共 ${comments} 則`, comments);

  const revTotal = [...revisions.values()].reduce((s, n) => s + n, 0);
  if (revTotal) {
    const names: [string, string][] = [['ins', '插入'], ['del', '刪除'], ['format', '格式變更'], ['paraMark', '段落標記']];
    const where = names.filter(([k]) => revisions.get(k)).map(([k, n]) => `${n} ${revisions.get(k)}`).join('、');
    add('revisions', '追蹤修訂', '網頁上會標示出來，可逐一或全部接受／拒絕。按工具列「追蹤修訂」（Ctrl+Shift+E）開啟後，網頁上的輸入、刪除與格式變更也會記錄為修訂（作者與時間），在 Word 中可以看到；追蹤修訂開啟時，表格結構、分節與版面設定等少數變更會先擋下並說明。', where, revTotal);
  }
  const unsupported = unsupportedRev.total('revisions');
  if (unsupported) {
    add('revisionsUnsupported', '無法在網頁上處理的修訂（表格列／儲存格、分節設定、編號、註腳中的修訂）',
      `${KEPT}；網頁上看不到，也無法接受或拒絕，「全部接受」不包含這些，請在 Word 中處理。`, unsupportedRev.where('revisions'), unsupported);
  }

  if (fields.kinds.size) {
    const kinds = [...fields.kinds].map(([k, n]) => `${FIELD_NAMES[k] ? `${FIELD_NAMES[k]}（${k}）` : k} ${n}`).join('、');
    add('fields', '需要更新的欄位', '顯示上次在 Word 中計算的結果，網頁上不會自動更新（例如目錄頁碼、交互參照、日期）；在 Word 中開啟後全選再按 F9 即可更新。', kinds);
  }
  add('contentControls', '內容控制項', '其中的文字可以直接編輯，控制項本身原樣保留；但下拉選單、日期選擇、核取方塊等功能在網頁上無作用。');
  add('lockedControls', '已鎖定的內容控制項', '在 Word 中設為「無法刪除／無法編輯」的控制項，在網頁上沒有鎖定，仍可修改或刪除，請留意。');
  add('frames', '框架／首字放大', '網頁上以一般段落顯示，位置與 Word 不同；設定原樣保留。');
  add('rtl', '由右至左的文字（阿拉伯文、希伯來文等）', '網頁上的文字方向、排列與游標移動可能不正確；設定原樣保留。');

  const sections = scanSections(body);
  const sect = (id: string, title: string, list: number[], effect: string) =>
    add(id, title, effect, sectionWhere(list, sections.count), list.length);
  sect('columns', '多欄版面（分欄）', sections.columns,
    '網頁上仍以單欄編輯與顯示（分頁位置會和 Word 不同），網頁產生的 PDF 也是單欄；分欄設定與分欄符號原樣保留（在網頁「版面設定」設定的分欄也一樣），用 Word 開啟或列印時會分欄。');
  sect('lineNumbers', '行號', sections.lineNumbers, '網頁上不顯示行號；設定原樣保留。');
  sect('pageBorders', '頁面框線', sections.pageBorders, '網頁上不顯示頁面框線；設定原樣保留。');
  sect('sectionVAlign', '頁面垂直對齊', sections.vAlign, '網頁上內容一律從頁面頂端開始；設定原樣保留。');
  sect('verticalText', '直書（文字方向）', sections.vertical, '網頁上以橫書顯示；設定原樣保留。');
  sect('pageNumberFormat', '頁碼格式（章節頁碼或特殊格式）', sections.pageNumberFormat, '網頁上的頁碼不含章節編號（1-1 顯示為 1），特殊格式以 1, 2, 3 顯示；在 Word 中仍依原本格式。');

  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const macros = names.filter((n) => /(^|\/)vbaProject\.bin$/i.test(n) || /(^|\/)activeX\//i.test(n)).length;
  add('macros', '巨集或 ActiveX 控制項', '網頁上不會執行；存檔時原樣保留。', `${macros + t.total('macros')} 個`, macros + t.total('macros'));
  const fonts = names.filter((n) => /(^|\/)fonts\/[^/]+\.(odttf|ttf|otf|fntdata)$/i.test(n)).length;
  add('embeddedFonts', '內嵌字型', '網頁上改用電腦已安裝的字型顯示，字形與換行可能和 Word 不同；內嵌字型原樣保留。', `${fonts} 個字型檔`, fonts);
  const images = names.filter((n) => /(^|\/)media\//.test(n) && BROWSER_UNSUPPORTED_IMAGES.test(n)).length;
  add('legacyImages', '瀏覽器無法顯示的圖片格式（EMF／WMF／TIFF）', '網頁上可能顯示為空白或標籤；圖片原樣保留，在 Word 中正常顯示。', `${images} 張`, images);

  const settingsRel = rels.find((r) => r.type === REL_TYPE.settings);
  const settings = settingsRel ? parse(await text(zip, settingsRel.target)) : null;
  const protection = settings
    ? Array.from(settings.children).filter((c) => isW(c, 'documentProtection', 'writeProtection')).filter((c) => {
        const enforced = attr(c, 'enforcement');
        return c.localName === 'writeProtection' || enforced === '1' || enforced === 'true' || enforced === 'on';
      }).length
    : 0;
  add('protection', '文件保護（限制編輯）', '網頁上不會套用 Word 的限制編輯設定，所有內容都能修改；設定原樣保留，在 Word 中仍然有效。', '整份文件', protection ? 1 : 0);

  const paragraphs = body.getElementsByTagNameNS(NS.w, 'p').length;
  const bytes = mainXml!.length;
  if (bytes >= LARGE_BYTES || paragraphs >= LARGE_PARAGRAPHS) {
    items.push({
      id: 'large',
      title: '大型文件',
      count: 1,
      where: `${paragraphs} 個段落`,
      effect: '開啟、捲動與存檔可能較慢，請耐心等候。',
    });
  }
  return { items };
}
