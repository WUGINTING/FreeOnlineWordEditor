<script lang="ts">
/** Ids for the ribbon's tabs and panel (aria-controls), one ribbon per editor on the page. */
let ribbons = 0;
</script>

<script setup lang="ts">
import { tl } from './locale';
// The ribbon: Word's tabs (檔案、常用、插入、設計、版面配置、參考資料、校閱、檢視, and the 表格 /
// 圖片 tabs while a table or picture is selected), each command where Word has it, so someone
// who knows Word finds it without looking.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import { Selection } from 'prosemirror-state';
import { selectAll } from 'prosemirror-commands';
import {
  addColumnAfter, addColumnBefore, addRowAfter, addRowBefore, deleteColumn, deleteRow, deleteTable,
  mergeCells, splitCell,
} from 'prosemirror-tables';
import { promptLink, type DocxEditor, type EditorSnapshot } from '../editor/core';
import { schema } from '../editor/schema';
import {
  FONT_SIZES, clearFormatting, continueNumbering, indent, insertBlankPage, insertPageBreak, insertTable, restartNumbering, setAlign, setFont, setLineSpacing,
  setMark, setParagraphStyle, toggle, toggleFormat, toggleList,
} from '../editor/commands';
import { isEastAsianFont } from '../docx/fonts';
import type { ListKind } from '../docx/numbering';
import type { PageSetup, ParagraphStyleInfo } from '../docx/model';
import { clipboardCommand, pasteClipboard } from './editorActions';
import { selectedImage } from '../editor/commands';
import { UNSUPPORTED_IMAGE } from '../editor/imageView';
import TablePanel from './TablePanel.vue';
import ColorPicker, { HIGHLIGHT_PALETTE, STANDARD_COLORS } from './ColorPicker.vue';
import ImagePanel from './ImagePanel.vue';
import ShapePanel from './ShapePanel.vue';
import { SHAPE_KINDS } from '../docx/shapeOps';
import { distributeColumns } from '../editor/tableCommands';
import SymbolGallery from './SymbolGallery.vue';
import {
  acceptAllRevisions, acceptRevision, collectRevisions, goToRevision, rejectAllRevisions, rejectRevision, reviewSummary, revisionSpoken,
  toggleRevisionMarks,
} from '../editor/review';
import { ICONS, MARGIN_PRESETS, PAPERS, near } from './ribbon';
import { composing } from './keys';
import { WATERMARK_PRESETS, textWatermark } from '../docx/watermark';

const props = withDefaults(
  defineProps<{
    editor: DocxEditor | null;
    snapshot: EditorSnapshot | null;
    styles: ParagraphStyleInfo[];
    /** The page has its own 檔案 menu (save, download, versions…): 「檔案」 emits `file`. */
    fileMenu?: boolean;
    /** Viewing only: just 檔案 and 檢視. */
    readOnly?: boolean;
    /** What the editor around the ribbon shows (the 檢視 / 校閱 toggles). */
    outlineOpen?: boolean;
    commentsOpen?: boolean;
    commentCount?: number;
    /** The 修訂窗格 is open. */
    revisionsOpen?: boolean;
    spellcheck?: boolean;
    /** The zoom choice: a factor ('1', '1.5' …) or 'fit'. */
    zoom?: string;
  }>(),
  { fileMenu: false, readOnly: false, outlineOpen: false, commentsOpen: false, commentCount: 0, revisionsOpen: false, spellcheck: false, zoom: '1' },
);

const emit = defineEmits<{
  /** Open the find panel (true: with replace). */
  (e: 'find', replace: boolean): void;
  /** Open the page setup dialog. */
  (e: 'page-setup'): void;
  /** Open the paragraph dialog (indents, spacing, line spacing). */
  (e: 'paragraph'): void;
  /** Start a new comment on the selection (G09). */
  (e: 'comment'): void;
  /** 「檔案」 was pressed (only with `fileMenu`). */
  (e: 'file'): void;
  (e: 'toggle-outline'): void;
  (e: 'toggle-comments'): void;
  (e: 'toggle-revisions'): void;
  (e: 'toggle-spellcheck'): void;
  (e: 'zoom', value: string): void;
  /** 插入 › 符號 › 其他符號…: open the 符號 dialog. */
  (e: 'symbol'): void;
  /** 設計 › 浮水印 › 自訂浮水印…: open the 浮水印 dialog. */
  (e: 'watermark'): void;
  /** 參考資料 › 更新目錄: rebuilds the table of contents from the headings, then its page numbers. */
  (e: 'update-toc'): void;
  /** 參考資料 › 目錄: inserts a table of contents (headings 1-3) at the cursor. */
  (e: 'insert-toc'): void;
  /** A short message for the user. */
  (e: 'notice', message: string): void;
}>();

const uid = `dx-ribbon-${++ribbons}`;

/** A shape is selected: show 圖形格式. */
const shapeSelected = computed(() => {
  void props.snapshot;
  return !!props.editor?.shapes?.info();
});
/** 插入 › 圖案: the kinds of shape, by group (文字方塊 has a button of its own). */
const SHAPE_MENU = [
  { label: '流程圖', kinds: SHAPE_KINDS.filter((k) => k.label.startsWith('流程圖')) },
  { label: '基本圖案', kinds: SHAPE_KINDS.filter((k) => ['rect', 'roundRect', 'ellipse'].includes(k.id)) },
  { label: '箭號', kinds: SHAPE_KINDS.filter((k) => k.id.endsWith('Arrow') && !k.line) },
  { label: '線條', kinds: SHAPE_KINDS.filter((k) => k.line) },
];
function insertShape(id: string) {
  props.editor?.startInsertShape(id);
}

/** A picture is selected: show the picture tools. */
const imageSelected = computed(() => {
  void props.snapshot; // re-evaluated on every editor update
  const state = props.editor?.activeView?.state;
  return !!state && !!selectedImage(state);
});

const FONTS = [
  { label: '新細明體', value: 'PMingLiU' },
  { label: '標楷體', value: 'DFKai-SB' },
  { label: '微軟正黑體', value: 'Microsoft JhengHei' },
  { label: 'Calibri', value: 'Calibri' },
  { label: 'Arial', value: 'Arial' },
  { label: 'Times New Roman', value: 'Times New Roman' },
  { label: 'Courier New', value: 'Courier New' },
];
const SIZES = FONT_SIZES;
const SPACING = [
  { label: '單行', value: 1 },
  { label: '1.15', value: 1.15 },
  { label: '1.5 倍', value: 1.5 },
  { label: '2 倍', value: 2 },
];
const ZOOMS = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2];

// Friendly names for the built-in styles Word uses.
const STYLE_NAMES: Record<string, string> = {
  normal: '內文', title: '標題', subtitle: '副標題',
  'heading 1': '標題 1', 'heading 2': '標題 2', 'heading 3': '標題 3',
  'heading 4': '標題 4', 'heading 5': '標題 5', 'heading 6': '標題 6',
  'list paragraph': '清單段落', quote: '引文',
};
const PREFERRED = ['normal', 'title', 'subtitle', 'heading 1', 'heading 2', 'heading 3', 'heading 4'];
/** Word's 樣式 gallery shows these first. */
const GALLERY = ['normal', 'heading 1', 'heading 2', 'heading 3', 'title'];

const styleOptions = computed(() => {
  const list = props.styles.map((s) => {
    const known = STYLE_NAMES[s.name.toLowerCase()];
    return { id: s.id, key: s.name.toLowerCase(), label: known != null ? tl(known) : s.name };
  });
  const rank = (k: string) => {
    const i = PREFERRED.indexOf(k);
    return i < 0 ? PREFERRED.length : i;
  };
  return list.sort((a, b) => rank(a.key) - rank(b.key) || a.label.localeCompare(b.label));
});
const gallery = computed(() =>
  GALLERY.map((k) => styleOptions.value.find((o) => o.key === k)).filter((o): o is NonNullable<typeof o> => !!o),
);

const defaultStyleId = computed(() => props.styles.find((s) => s.name.toLowerCase() === 'normal')?.id ?? null);
const currentStyle = computed(() => s.value?.styleId ?? defaultStyleId.value);

const s = computed(() => props.snapshot);
/** Tracked changes in the part being edited (body, or the header/footer being edited). */
const rev = computed(() => {
  void props.snapshot;
  return reviewSummary(props.editor?.activeView?.state);
});
const run = (cmd: Parameters<DocxEditor['run']>[0]) => props.editor?.run(cmd);
/** Said by screen readers (aria-live): where 上一個 / 下一個 修訂 went, its author and kind (persona-300). */
const spoken = ref('');
function goRevision(dir: 1 | -1) {
  if (!run(goToRevision(dir))) return;
  const state = props.editor?.activeView?.state;
  const r = reviewSummary(state).current;
  if (!state || !r) return;
  const all = collectRevisions(state.doc);
  spoken.value = tl('第 {0} 處，共 {1} 處：{2}', all.findIndex((x) => x.key === r.key) + 1, all.length, revisionSpoken(r));
}
/** The open document's styles, so bold / alignment act on what the text really shows. */
const styles = () => props.editor?.model.styles;
const inBody = computed(() => s.value?.target === 'body');

function applyStyle(id: string) {
  run(setParagraphStyle(id === defaultStyleId.value ? null : id || null));
}
function onStyle(e: Event) {
  applyStyle((e.target as HTMLSelectElement).value);
}
function onFont(e: Event) {
  const v = (e.target as HTMLSelectElement).value;
  // As in Word: a Latin font leaves the Chinese text's font (標楷體 stays on 中文).
  run(setFont(v || null, isEastAsianFont(v)));
}
function onSize(e: Event) {
  const v = Number((e.target as HTMLSelectElement).value);
  run(setMark(schema.marks.fontSize, v ? { pt: v } : null));
}
function onSpacing(e: Event) {
  const v = Number((e.target as HTMLSelectElement).value);
  run(setLineSpacing(v || null));
}
/** 字型色彩 (null: 自動, the style's colour). */
function onColor(color: string | null) {
  run(setMark(schema.marks.color, color ? { color } : null));
}
/** 螢光標記 (null: 無色彩). */
function onHighlight(color: string | null) {
  run(setMark(schema.marks.highlight, color ? { color } : null));
}
function onLink() {
  // The same as Ctrl+K (editor/core.ts promptLink): the allowlist the page, paste and the DOCX writer use.
  promptLink(props.editor);
}
function list(kind: ListKind) {
  if (props.editor) run(toggleList(kind, props.editor.model.numbering));
}
// 多層次清單 (persona-300 B-2): the 公文 list, and restarting / continuing a list's numbers.
const numberingCan = computed(() => {
  void props.snapshot;
  const ed = props.editor;
  if (!ed?.activeView || menu.value !== 'multilevel') return { restart: false, continue: false };
  return { restart: ed.can(restartNumbering(ed.model.numbering)), continue: ed.can(continueNumbering(ed.model.numbering)) };
});
function renumber(restart: boolean) {
  const ed = props.editor;
  if (ed) run(restart ? restartNumbering(ed.model.numbering) : continueNumbering(ed.model.numbering));
}

const fileInput = ref<HTMLInputElement | null>(null);
/** 插入 › 圖片: one or more pictures, inserted one after the other in the order picked. */
async function onImage(e: Event) {
  const input = e.target as HTMLInputElement;
  const files = Array.from(input.files ?? []);
  input.value = '';
  for (const file of files) {
    try {
      await props.editor?.insertImageFile(file);
    } catch {
      window.alert(tl(UNSUPPORTED_IMAGE));
    }
  }
}

// ----- 剪貼簿 -----
// The browser lets a button copy and cut (as Ctrl+C / Ctrl+X do), and paste only with the
// user's permission; without it, the user is told to press Ctrl+V.
function clip(command: 'cut' | 'copy') {
  clipboardCommand(props.editor, command, (m) => emit('notice', m));
}
function paste() {
  return pasteClipboard(props.editor, (m) => emit('notice', m));
}
function selectEverything() {
  run(selectAll);
}

// ----- 頁首及頁尾 -----
const pageIndex = () => Math.max(0, (s.value?.currentPage ?? 1) - 1);
function editHeaderFooter(kind: 'header' | 'footer') {
  props.editor?.editHeaderFooter(kind, pageIndex());
}
/** 頁碼 › 頁面底端: the footer shows the page number (centred when the footer was empty). */
function pageNumberAtBottom() {
  const ed = props.editor;
  if (!ed) return;
  ed.editHeaderFooter('footer', pageIndex());
  const view = ed.activeView;
  if (!view || ed.target !== 'footer') return;
  const empty = view.state.doc.textContent.trim() === '';
  view.dispatch(view.state.tr.setSelection(Selection.atEnd(view.state.doc)));
  if (empty) ed.run(setAlign('center', styles()));
  ed.insertField('PAGE');
}

// ----- 設計 -----
/** 浮水印 › a preset: diagonal, semi-transparent silver text in every header (DocxEditor.setWatermark). */
function presetWatermark(text: string) {
  props.editor?.setWatermark(textWatermark(text));
}
/** 移除浮水印 is offered only when there is one (even one that can't be shown). */
const hasWatermark = () => !!props.editor?.hasWatermark();

// ----- 版面配置 -----
function cursorPage(): PageSetup | null {
  return props.editor?.view ? props.editor.cursorSection().page : null;
}
function setPage(change: (p: PageSetup) => PageSetup) {
  const p = cursorPage();
  if (!p || !props.editor) return;
  const next = change(p);
  if (next.marginLeft + next.marginRight >= next.width - 567 || next.marginTop + next.marginBottom >= next.height - 567) {
    emit('notice', tl('邊界太大，頁面上沒有可以寫字的地方了。請改用「自訂邊界…」調整。'));
    return;
  }
  props.editor.setPageSetup(next);
}
const marginsOf = (m: (typeof MARGIN_PRESETS)[number]) => (p: PageSetup) => ({ ...p, marginTop: m.tb, marginBottom: m.tb, marginLeft: m.lr, marginRight: m.lr });
const isMargins = (p: PageSetup | null, m: (typeof MARGIN_PRESETS)[number]) =>
  !!p && near(p.marginTop, m.tb) && near(p.marginBottom, m.tb) && near(p.marginLeft, m.lr) && near(p.marginRight, m.lr);
function setOrientation(landscape: boolean) {
  setPage((p) => ((p.width > p.height) === landscape ? p : { ...p, width: p.height, height: p.width }));
}
function setPaper(paper: (typeof PAPERS)[number]) {
  setPage((p) => (p.width > p.height ? { ...p, width: paper.h, height: paper.w } : { ...p, width: paper.w, height: paper.h }));
}
const isPaper = (p: PageSetup | null, paper: (typeof PAPERS)[number]) =>
  !!p && near(Math.min(p.width, p.height), paper.w) && near(Math.max(p.width, p.height), paper.h);
const cursorColumns = () => (props.editor?.view ? props.editor.cursorSection().columns.count : 1);
function setColumns(count: number) {
  props.editor?.setColumns({ count });
  if (count > 1) emit('notice', tl('已設為 {0} 欄。', count));
}

// ----- 參考資料 / 校閱 -----
const tocCount = computed(() => {
  void props.snapshot;
  const ed = props.editor;
  return tab.value === 'references' && ed ? ed.pageReferenceCount() + ed.tableOfContentsCount() : 0;
});
const countOpen = ref(false);
const counts = computed(() => {
  void props.snapshot;
  const ed = props.editor;
  return countOpen.value && ed ? { all: ed.wordCount(), selected: ed.selectionWordCount() } : null;
});
const n = (v: number) => v.toLocaleString('zh-TW');
function closeCount(e: KeyboardEvent) {
  e.preventDefault();
  e.stopPropagation();
  countOpen.value = false;
}

// ----- Tabs -----
type TabId = 'home' | 'insert' | 'design' | 'layout' | 'references' | 'review' | 'view' | 'tableDesign' | 'tableLayout' | 'picture' | 'shape';
interface Tab {
  id: TabId;
  label: string;
  /** Shown only while a table or picture is selected (Word's contextual tabs). */
  context?: string;
}
const inTable = computed(() => !!s.value?.inTable);
/** 合併儲存格 needs two or more cells selected (persona-300: it did nothing, silently). */
const canMerge = computed(() => {
  void props.snapshot;
  return inTable.value && !!props.editor?.can(mergeCells);
});
const tabs = computed<Tab[]>(() => {
  if (props.readOnly) return [{ id: 'view', label: tl('檢視') }];
  const out: Tab[] = [
    { id: 'home', label: tl('常用') },
    { id: 'insert', label: tl('插入') },
    { id: 'design', label: tl('設計') },
    { id: 'layout', label: tl('版面配置') },
    { id: 'references', label: tl('參考資料') },
    { id: 'review', label: tl('校閱') },
    { id: 'view', label: tl('檢視') },
  ];
  if (inTable.value) out.push({ id: 'tableDesign', label: tl('表格設計'), context: tl('表格工具') }, { id: 'tableLayout', label: tl('表格版面配置'), context: tl('表格工具') });
  if (imageSelected.value) out.push({ id: 'picture', label: tl('圖片格式'), context: tl('圖片工具') });
  if (shapeSelected.value) out.push({ id: 'shape', label: tl('圖形格式'), context: tl('繪圖工具') });
  return out;
});
const tab = ref<TabId>(props.readOnly ? 'view' : 'home');
// A table or picture tab goes away with its table or picture: back to 常用 (as in Word).
watch(tabs, (list) => {
  if (!list.some((t) => t.id === tab.value)) tab.value = list[0].id;
});
const tabId = (id: TabId) => `${uid}-tab-${id}`;
const panelId = `${uid}-panel`;
const activeTab = computed(() => tabs.value.find((t) => t.id === tab.value) ?? tabs.value[0]);

// 摺疊功能區 (Ctrl+F1, or double-click a tab, as in Word): only the tabs show; a tab opens its
// panel over the page until a command is used or the focus leaves.
const COLLAPSE_KEY = 'papyrus.ribbon.collapsed';
function rememberedCollapsed(): boolean {
  try {
    return localStorage.getItem(COLLAPSE_KEY) === '1';
  } catch {
    return false; // storage unavailable: the full ribbon
  }
}
const collapsed = ref(rememberedCollapsed());
const peek = ref(false);
const panelShown = computed(() => !collapsed.value || peek.value);
function toggleCollapsed() {
  collapsed.value = !collapsed.value;
  peek.value = false;
  try {
    localStorage.setItem(COLLAPSE_KEY, collapsed.value ? '1' : '0');
  } catch {
    // not remembered; still applies now
  }
}
function selectTab(id: TabId) {
  if (collapsed.value && peek.value && tab.value === id) {
    peek.value = false;
    return;
  }
  tab.value = id;
  if (collapsed.value) peek.value = true;
}
/** A command in a panel shown over the page closes it (not one that opens a menu). */
function onPanelClick(e: MouseEvent) {
  if (!collapsed.value || !peek.value) return;
  const b = (e.target as Element).closest?.('button');
  if (b && !b.hasAttribute('aria-haspopup') && !b.hasAttribute('aria-expanded')) peek.value = false;
}

const tabList = ref<HTMLElement | null>(null);
function onTabKey(e: KeyboardEvent) {
  if (composing(e)) return;
  const items = Array.from(tabList.value?.querySelectorAll<HTMLElement>('[role="tab"]') ?? []);
  const i = items.indexOf(e.target as HTMLElement);
  if (i < 0) return;
  let next = -1;
  if (e.key === 'ArrowRight') next = (i + 1) % items.length;
  else if (e.key === 'ArrowLeft') next = (i + items.length - 1) % items.length;
  else if (e.key === 'Home') next = 0;
  else if (e.key === 'End') next = items.length - 1;
  else if (e.key === 'ArrowDown') {
    e.preventDefault();
    focus();
    return;
  } else if (e.key === 'Escape') {
    e.preventDefault();
    if (peek.value) peek.value = false;
    else props.editor?.activeView?.focus();
    return;
  }
  if (next < 0) return;
  e.preventDefault();
  items[next].focus();
  tab.value = tabs.value[next].id;
}

// ----- 檔案 -----
function onFile(e: MouseEvent) {
  if (props.fileMenu) emit('file');
  else toggleMenu('file', e);
}
async function downloadDocx() {
  const ed = props.editor;
  if (!ed) return;
  const url = URL.createObjectURL(await ed.save());
  const a = document.createElement('a');
  a.href = url;
  a.download = 'document.docx';
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ----- Menus (邊界、方向、大小、欄、分隔設定、頁碼、檔案) -----
// A button with aria-haspopup opens a list of choices; Up / Down move in it, Escape closes it
// back on the button, a click elsewhere closes it.
const menu = ref<string | null>(null);
let menuTrigger: HTMLElement | null = null;
function toggleMenu(id: string, e: MouseEvent) {
  if (menu.value === id) {
    closeMenu(false);
    return;
  }
  menu.value = id;
  menuTrigger = e.currentTarget as HTMLElement;
  nextTick(() => menuTrigger?.parentElement?.querySelector<HTMLElement>('.dx-menu [role^="menuitem"]:not(:disabled)')?.focus());
}
function closeMenu(refocus = true) {
  menu.value = null;
  if (refocus) menuTrigger?.focus();
}
function pick(action: () => void) {
  closeMenu(false);
  action();
}
function onMenuKey(e: KeyboardEvent) {
  if (composing(e)) return;
  const items = Array.from((e.currentTarget as HTMLElement).querySelectorAll<HTMLElement>('[role^="menuitem"]:not(:disabled)'));
  const i = items.indexOf(document.activeElement as HTMLElement);
  const to = (k: number) => {
    e.preventDefault();
    e.stopPropagation();
    items[(k + items.length) % items.length]?.focus();
  };
  if (e.key === 'Escape') {
    e.preventDefault();
    e.stopPropagation();
    closeMenu(true);
  } else if (e.key === 'ArrowDown') to(i + 1);
  else if (e.key === 'ArrowUp') to(i - 1);
  else if (e.key === 'Home') to(0);
  else if (e.key === 'End') to(items.length - 1);
  else if (e.key === 'Tab') closeMenu(false);
}

// Table size picker. From the keyboard it opens with the focus on the first cell; the arrow keys
// pick the size (one Tab stop, not 64) and Escape closes it, back on 「插入表格」.
const tableOpen = ref(false);
const hover = ref({ r: 0, c: 0 });
const tableButton = ref<HTMLButtonElement | null>(null);
const gridCells = ref<HTMLElement | null>(null);
/** The cell that is the grid's Tab stop: the one picked, or the first. */
const gridStop = computed(() => (hover.value.r ? hover.value : { r: 1, c: 1 }));
function toggleTable(e: MouseEvent) {
  tableOpen.value = !tableOpen.value;
  tableForm.value = false;
  // detail 0: Enter / Space, not a mouse click.
  if (tableOpen.value && e.detail === 0) nextTick(() => focusCell(1, 1));
}
function focusCell(r: number, c: number) {
  gridCells.value?.querySelectorAll<HTMLElement>('.dx-cell')[(r - 1) * 8 + (c - 1)]?.focus();
}
const GRID_MOVES: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
function onGridKey(e: KeyboardEvent) {
  if (composing(e)) return;
  if (e.key === 'Escape') {
    e.preventDefault();
    e.stopPropagation();
    tableOpen.value = false;
    tableForm.value = false;
    hover.value = { r: 0, c: 0 };
    tableButton.value?.focus();
    return;
  }
  const move = GRID_MOVES[e.key];
  if (tableForm.value || !move || e.altKey || e.ctrlKey || e.metaKey) return; // Enter / Space pick the cell, Tab leaves
  e.preventDefault();
  e.stopPropagation();
  const clamp = (k: number) => Math.min(8, Math.max(1, k));
  focusCell(clamp(gridStop.value.r + move[0]), clamp(gridStop.value.c + move[1]));
}
// 插入表格… (persona-300): any size, as Word's dialog (up to 63 columns; rows up to MAX_TABLE_ROWS).
const MAX_TABLE_COLS = 63;
const MAX_TABLE_ROWS = 500;
const tableForm = ref(false);
const tableRows = ref(2);
const tableCols = ref(5);
function openTableForm() {
  tableForm.value = true;
  nextTick(() => tableButton.value?.parentElement?.querySelector<HTMLInputElement>('.dx-tform input')?.select());
}
function submitTableForm() {
  const r = Math.round(Number(tableRows.value));
  const c = Math.round(Number(tableCols.value));
  if (!(r >= 1 && r <= MAX_TABLE_ROWS && c >= 1 && c <= MAX_TABLE_COLS)) {
    emit('notice', tl('列數請輸入 1～{0}，欄數請輸入 1～{1}。', MAX_TABLE_ROWS, MAX_TABLE_COLS));
    return;
  }
  tableForm.value = false;
  pickTable(r, c);
}
function pickTable(r: number, c: number) {
  tableOpen.value = false;
  if (!props.editor) return;
  run(insertTable(r, c, props.editor.contentWidth));
  // As in Word, the new table's tools come up.
  tab.value = 'tableDesign';
}

// One Tab stop for the panel's controls (GOV-ISSUE-013; WAI-ARIA toolbar pattern): the control
// used last has tabindex 0, every other -1; Left / Right / Home / End move between the controls,
// and Escape goes back to the text. A text box keeps those keys for its own text. Controls come
// and go (tabs, table / picture tools, disabled buttons), so the tab stops follow the DOM.
// Menus and the table grid have their own keys.
const ribbon = ref<HTMLElement | null>(null);
const root = ref<HTMLElement | null>(null);
const CONTROL = 'button, select, input, textarea';
const TYPES_TEXT = new Set(['text', 'search', 'number', 'email', 'url', 'tel', 'password']);
let current: HTMLElement | null = null;
function usable(el: HTMLElement): boolean {
  if ((el as HTMLButtonElement).disabled || el.closest('.dx-popup')) return false;
  for (let node: HTMLElement | null = el; node && node !== root.value; node = node.parentElement) {
    if (node.hidden || node.style.display === 'none') return false;
  }
  return true;
}
function controls(): HTMLElement[] {
  return root.value ? Array.from(root.value.querySelectorAll<HTMLElement>(CONTROL)).filter(usable) : [];
}
function syncTabStops() {
  if (!root.value) return;
  const items = controls();
  if (!current || !items.includes(current)) current = items[0] ?? null;
  for (const el of root.value.querySelectorAll<HTMLElement>(CONTROL)) {
    if (el.closest('.dx-popup')) continue; // menus and the table grid have their own
    const index = el === current ? 0 : -1;
    if (el.tabIndex !== index) el.tabIndex = index;
  }
}
function onFocusIn(e: FocusEvent) {
  const el = e.target as HTMLElement;
  if (el !== current && el.matches?.(CONTROL) && usable(el)) {
    current = el;
    syncTabStops();
  }
}
function typesText(el: HTMLElement) {
  return el instanceof HTMLTextAreaElement || (el instanceof HTMLInputElement && TYPES_TEXT.has(el.type)) || el.isContentEditable;
}
function onKeydown(e: KeyboardEvent) {
  // An input method's Esc in a box of the table / picture tools drops its candidates only (A-7).
  if (e.defaultPrevented || composing(e)) return;
  if (e.key === 'Escape') {
    e.preventDefault();
    if (peek.value) {
      peek.value = false;
      tabList.value?.querySelector<HTMLElement>('[aria-selected="true"]')?.focus();
    } else props.editor?.activeView?.focus();
    return;
  }
  if (e.altKey || e.ctrlKey || e.metaKey || e.shiftKey) return;
  if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key) || typesText(e.target as HTMLElement)) return;
  const items = controls();
  const i = items.indexOf(e.target as HTMLElement);
  if (i < 0) return;
  const count = items.length;
  const next = e.key === 'Home' ? 0 : e.key === 'End' ? count - 1 : (i + (e.key === 'ArrowLeft' ? count - 1 : 1)) % count;
  e.preventDefault();
  items[next].focus();
}
/** Puts the keyboard on the ribbon, on the control used last (Alt+F10). */
function focus() {
  const focusCurrent = () => {
    syncTabStops();
    current?.focus();
  };
  if (!collapsed.value || peek.value) return focusCurrent();
  peek.value = true;
  nextTick(focusCurrent);
}
defineExpose({ focus, toggleCollapsed, insertLink: onLink });

/** A click outside a menu, the table grid, the word count or the shown-over panel closes it. */
function onDocumentDown(e: MouseEvent) {
  const t = e.target as Node;
  if (menu.value && !menuTrigger?.parentElement?.contains(t)) closeMenu(false);
  if (tableOpen.value && !tableButton.value?.parentElement?.contains(t)) tableOpen.value = false;
  if (countOpen.value && !(t instanceof Element && t.closest('.dx-rcount-pop'))) countOpen.value = false;
  if (peek.value && !ribbon.value?.contains(t)) peek.value = false;
}

let observer: MutationObserver | null = null;
onMounted(() => {
  syncTabStops();
  document.addEventListener('mousedown', onDocumentDown, true);
  if (!root.value) return;
  observer = new MutationObserver(syncTabStops);
  observer.observe(root.value, { subtree: true, childList: true, attributes: true, attributeFilter: ['disabled', 'hidden', 'style'] });
});
onBeforeUnmount(() => {
  observer?.disconnect();
  document.removeEventListener('mousedown', onDocumentDown, true);
});
</script>

<template>
  <!-- Word's ribbon: 檔案, the tabs (one Tab stop, arrow keys), and the tab's panel (one Tab stop,
       roving tabindex). Alt+F10 from the text comes to the panel, Escape goes back. -->
  <div ref="ribbon" class="dx-ribbon" :class="{ 'dx-collapsed': collapsed, 'dx-peek': peek }">
    <div class="dx-tabrow">
      <div class="dx-pop">
        <button
          type="button"
          class="dx-file"
          :aria-haspopup="fileMenu ? 'dialog' : 'menu'"
          :aria-expanded="fileMenu ? undefined : menu === 'file'"
          :title="tl('檔案：儲存、下載、列印、版本紀錄等')"
          @click="onFile"
        >{{ tl('檔案') }}</button>
        <div v-if="menu === 'file'" class="dx-menu dx-popup" role="menu" :aria-label="tl('檔案')" @keydown="onMenuKey">
          <button type="button" role="menuitem" tabindex="-1" :title="tl('列印／另存 PDF（在列印視窗選擇「另存為 PDF」）')" @click="pick(() => editor?.print())">
            <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.print" />{{ tl('列印') }}
          </button>
          <button type="button" role="menuitem" tabindex="-1" :title="tl('下載 Word 檔（.docx）')" @click="pick(downloadDocx)">
            <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.download" />{{ tl('下載') }}
          </button>
        </div>
      </div>
      <div ref="tabList" class="dx-tabs" role="tablist" :aria-label="tl('功能區索引標籤')" @keydown="onTabKey">
        <button
          v-for="t in tabs"
          :id="tabId(t.id)"
          :key="t.id"
          type="button"
          role="tab"
          class="dx-rtab"
          :class="{ 'dx-context': t.context }"
          :aria-selected="t.id === activeTab.id"
          :aria-controls="panelId"
          :tabindex="t.id === activeTab.id ? 0 : -1"
          :title="t.context ? tl('{0}：{1}', t.context, t.label) : undefined"
          @click="selectTab(t.id)"
          @dblclick="toggleCollapsed"
        >{{ t.label }}</button>
      </div>
      <button
        type="button"
        class="dx-collapse"
        :aria-pressed="collapsed"
        :title="collapsed ? tl('固定功能區：一直顯示工具 (Ctrl+F1)') : tl('摺疊功能區：只顯示索引標籤，按標籤才顯示工具 (Ctrl+F1)')"
        :aria-label="collapsed ? tl('固定功能區') : tl('摺疊功能區')"
        @click="toggleCollapsed"
      >
        <svg class="dx-ico" viewBox="0 0 24 24" v-html="collapsed ? ICONS.pin : ICONS.collapse" />
      </button>
    </div>

    <div :id="panelId" class="dx-tabpanel" role="tabpanel" :aria-labelledby="tabId(activeTab.id)" v-show="panelShown">
      <div
        ref="root"
        class="dx-toolbar"
        role="toolbar"
        :aria-label="tl('文件格式工具列')"
        aria-keyshortcuts="Alt+F10"
        @mousedown.self.prevent
        @keydown="onKeydown"
        @focusin="onFocusIn"
        @click="onPanelClick"
      >
        <!-- ===== 常用 ===== -->
        <div v-show="tab === 'home'" class="dx-panel">
          <div class="dx-rgroup" role="group" :aria-label="tl('復原')">
            <div class="dx-rbody dx-stack">
              <button type="button" class="dx-sm" :title="tl('復原 (Ctrl+Z)')" :disabled="!s?.canUndo" @click="editor?.undo()">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.undo" /><span>{{ tl('復原') }}</span>
              </button>
              <button type="button" class="dx-sm" :title="tl('取消復原 (Ctrl+Y)')" :disabled="!s?.canRedo" @click="editor?.redo()">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.redo" /><span>{{ tl('取消復原') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('復原') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('剪貼簿')">
            <div class="dx-rbody">
              <button type="button" class="dx-big" :title="tl('貼上 (Ctrl+V)')" @click="paste">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.paste" /><span>{{ tl('貼上') }}</span>
              </button>
              <div class="dx-stack">
                <button type="button" class="dx-sm" :title="tl('剪下 (Ctrl+X)')" @click="clip('cut')">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.cut" /><span>{{ tl('剪下') }}</span>
                </button>
                <button type="button" class="dx-sm" :title="tl('複製 (Ctrl+C)')" @click="clip('copy')">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.copy" /><span>{{ tl('複製') }}</span>
                </button>
                <!-- 複製格式: behaviour in editor/formatPainter.ts; a click paints once, a double-click keeps painting until Esc. -->
                <button
                  type="button"
                  class="dx-sm"
                  :title="tl('複製格式：複製所選文字的格式，再用滑鼠選取要套用的文字（按兩下可連續套用，按 Esc 結束） (鍵盤：Ctrl+Shift+C 複製格式，再按 Ctrl+Shift+V 貼上格式)')"
                  :class="{ on: s?.formatPainter }"
                  :aria-pressed="!!s?.formatPainter"
                  @click="editor?.toggleFormatPainter(false)"
                  @dblclick="editor?.startFormatPainter(true)"
                >
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.formatPainter" /><span>{{ tl('複製格式') }}</span>
                </button>
              </div>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('剪貼簿') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('字型')">
            <div class="dx-rbody dx-rows">
              <div class="dx-rrow">
                <select :title="tl('字型')" :aria-label="tl('字型')" class="dx-w-font" :value="s?.fontFamily ?? ''" @change="onFont">
                  <option value="">{{ s?.mixed?.fontFamily ? tl('（多種字型）') : tl('（預設字型）') }}</option>
                  <option v-if="s?.fontFamily && !FONTS.some((f) => f.value === s!.fontFamily)" :value="s.fontFamily">{{ s.fontFamily }}</option>
                  <option v-for="f in FONTS" :key="f.value" :value="f.value">{{ f.label }}</option>
                </select>
                <select :title="tl('字型大小')" :aria-label="tl('字型大小')" class="dx-w-size" :value="s?.fontSize ?? ''" @change="onSize">
                  <option value="">{{ s?.mixed?.fontSize ? tl('多種') : '—' }}</option>
                  <option v-if="s?.fontSize && !SIZES.includes(s.fontSize)" :value="s.fontSize">{{ s.fontSize }}</option>
                  <option v-for="size in SIZES" :key="size" :value="size">{{ size }}</option>
                </select>
                <button type="button" :title="tl('清除格式 (Ctrl+Space)')" :aria-label="tl('清除格式')" @click="run(clearFormatting)">⌫</button>
              </div>
              <div class="dx-rrow">
                <button type="button" :title="tl('粗體 (Ctrl+B)')" :aria-label="tl('粗體')" :class="{ on: s?.bold }" :aria-pressed="!!(s?.bold)" @click="run(toggleFormat('bold', styles()))"><b>B</b></button>
                <button type="button" :title="tl('斜體 (Ctrl+I)')" :aria-label="tl('斜體')" :class="{ on: s?.italic }" :aria-pressed="!!(s?.italic)" @click="run(toggleFormat('italic', styles()))"><i>I</i></button>
                <button type="button" :title="tl('底線 (Ctrl+U)')" :aria-label="tl('底線')" :class="{ on: s?.underline }" :aria-pressed="!!(s?.underline)" @click="run(toggle('underline'))"><u>U</u></button>
                <button type="button" :title="tl('刪除線')" :aria-label="tl('刪除線')" :class="{ on: s?.strike }" :aria-pressed="!!(s?.strike)" @click="run(toggle('strike'))"><s>S</s></button>
                <button type="button" :title="tl('下標')" :aria-label="tl('下標')" :class="{ on: s?.subscript }" :aria-pressed="!!(s?.subscript)" @click="run(toggle('subscript'))">x₂</button>
                <button type="button" :title="tl('上標')" :aria-label="tl('上標')" :class="{ on: s?.superscript }" :aria-pressed="!!(s?.superscript)" @click="run(toggle('superscript'))">x²</button>
                <!-- Word's colour menus: named standard colours, 無色彩 / 自動, and 其他色彩… (persona-300). -->
                <ColorPicker :label="tl('螢光標記')" :colors="HIGHLIGHT_PALETTE" :value="s?.highlight ?? null" :none-label="tl('無色彩')" @pick="onHighlight">
                  <span class="dx-cp-glyph dx-hl">ab</span>
                </ColorPicker>
                <ColorPicker :label="tl('文字顏色')" :colors="STANDARD_COLORS" :value="s?.color ?? null" :mixed="!!s?.mixed?.color" :none-label="tl('自動')" @pick="onColor">
                  <span class="dx-cp-glyph">A</span>
                </ColorPicker>
              </div>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('字型') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('段落')">
            <div class="dx-rbody dx-rows">
              <div class="dx-rrow">
                <button type="button" :title="tl('項目符號清單')" :aria-label="tl('項目符號清單')" :class="{ on: s?.list === 'bullet' }" :aria-pressed="!!(s?.list === 'bullet')" @click="list('bullet')">
                  <svg viewBox="0 0 16 16"><circle cx="3" cy="4" r="1.2" /><circle cx="3" cy="8" r="1.2" /><circle cx="3" cy="12" r="1.2" /><path d="M6 4h8M6 8h8M6 12h8" /></svg>
                </button>
                <button type="button" :title="tl('編號清單')" :aria-label="tl('編號清單')" :class="{ on: s?.list === 'decimal' }" :aria-pressed="!!(s?.list === 'decimal')" @click="list('decimal')">
                  <svg viewBox="0 0 16 16"><text x="0.5" y="6" font-size="5">1</text><text x="0.5" y="13" font-size="5">2</text><path d="M6 4h8M6 11h8" /></svg>
                </button>
                <div class="dx-pop">
                  <button
                    type="button"
                    :title="tl('多層次清單：公文編號（一、（一）1.（1）甲、（甲））、重新編號、接續編號')"
                    :aria-label="tl('多層次清單')"
                    aria-haspopup="menu"
                    :aria-expanded="menu === 'multilevel'"
                    :class="{ on: s?.list === 'gongwen' }"
                    @click="toggleMenu('multilevel', $event)"
                  >
                    <svg viewBox="0 0 16 16"><text x="0" y="6.5" font-size="6">{{ tl('一') }}</text><text x="3" y="14" font-size="6">1</text><path d="M7 4.5h8M9 12h6" /></svg>
                  </button>
                  <div v-if="menu === 'multilevel'" class="dx-menu dx-popup" role="menu" :aria-label="tl('多層次清單')" @keydown="onMenuKey">
                    <button type="button" role="menuitemradio" tabindex="-1" :aria-checked="s?.list === 'gongwen'" @click="pick(() => list('gongwen'))">
                      <b>{{ tl('公文編號') }}</b><small>{{ tl('一、（一）1.（1）甲、（甲）；按 Tab 到下一層') }}</small>
                    </button>
                    <button type="button" role="menuitem" tabindex="-1" :disabled="!numberingCan.restart" @click="pick(() => renumber(true))">
                      <b>{{ tl('重新從 1 開始編號') }}</b><small>{{ tl('這一項和後面的項目重新編號') }}</small>
                    </button>
                    <button type="button" role="menuitem" tabindex="-1" :disabled="!numberingCan.continue" @click="pick(() => renumber(false))">
                      <b>{{ tl('接續編號') }}</b><small>{{ tl('接著前一個同樣的清單編號') }}</small>
                    </button>
                  </div>
                </div>
                <button type="button" :title="tl('減少縮排')" :aria-label="tl('減少縮排')" @click="run(indent(-1))">⇤</button>
                <button type="button" :title="tl('增加縮排')" :aria-label="tl('增加縮排')" @click="run(indent(1))">⇥</button>
                <!-- 顯示／隱藏編輯標記: display only, behaviour in editor/formattingMarks.ts. -->
                <button
                  type="button"
                  :title="tl('顯示／隱藏編輯標記：顯示段落標記、空格、定位字元等不會列印的符號 (Ctrl+Shift+8)')"
                  :aria-label="tl('顯示／隱藏編輯標記')"
                  :class="{ on: s?.showMarks }"
                  :aria-pressed="!!s?.showMarks"
                  @click="editor?.toggleShowMarks()"
                >¶</button>
              </div>
              <div class="dx-rrow">
                <button type="button" :title="tl('靠左對齊')" :aria-label="tl('靠左對齊')" :class="{ on: !s?.align || s.align === 'left' }" :aria-pressed="!!(!s?.align || s.align === 'left')" @click="run(setAlign('left', styles()))">
                  <svg viewBox="0 0 16 16"><path d="M2 3h12M2 6.5h8M2 10h12M2 13.5h8" /></svg>
                </button>
                <button type="button" :title="tl('置中')" :aria-label="tl('置中')" :class="{ on: s?.align === 'center' }" :aria-pressed="!!(s?.align === 'center')" @click="run(setAlign('center', styles()))">
                  <svg viewBox="0 0 16 16"><path d="M2 3h12M4 6.5h8M2 10h12M4 13.5h8" /></svg>
                </button>
                <button type="button" :title="tl('靠右對齊')" :aria-label="tl('靠右對齊')" :class="{ on: s?.align === 'right' }" :aria-pressed="!!(s?.align === 'right')" @click="run(setAlign('right', styles()))">
                  <svg viewBox="0 0 16 16"><path d="M2 3h12M6 6.5h8M2 10h12M6 13.5h8" /></svg>
                </button>
                <button type="button" :title="tl('左右對齊')" :aria-label="tl('左右對齊')" :class="{ on: s?.align === 'justify' }" :aria-pressed="!!(s?.align === 'justify')" @click="run(setAlign('justify', styles()))">
                  <svg viewBox="0 0 16 16"><path d="M2 3h12M2 6.5h12M2 10h12M2 13.5h12" /></svg>
                </button>
                <select :title="tl('行距')" :aria-label="tl('行距')" class="dx-w-size" value="" @change="onSpacing">
                  <option value="" disabled>{{ tl('行距') }}</option>
                  <option v-for="o in SPACING" :key="o.value" :value="o.value">{{ tl(o.label) }}</option>
                </select>
              </div>
            </div>
            <div class="dx-rlabel">
              <span aria-hidden="true">{{ tl('段落') }}</span>
              <button type="button" class="dx-launch" :title="tl('段落設定（縮排、段前段後距離、行距）')" :aria-label="tl('段落設定')" @click="emit('paragraph')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.launcher" />
              </button>
            </div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('樣式')">
            <div class="dx-rbody dx-rows">
              <div class="dx-rrow dx-gallery">
                <button
                  v-for="o in gallery"
                  :key="o.id"
                  type="button"
                  class="dx-style"
                  :class="['dx-style-' + o.key.replace(' ', ''), { on: currentStyle === o.id }]"
                  :aria-pressed="currentStyle === o.id"
                  :title="tl('樣式：{0}', o.label)"
                  @click="applyStyle(o.id)"
                >{{ o.label }}</button>
              </div>
              <div class="dx-rrow">
                <select :title="tl('段落樣式')" :aria-label="tl('段落樣式')" class="dx-w-style" :value="currentStyle ?? ''" @change="onStyle">
                  <option v-if="!defaultStyleId" value="">{{ tl('內文') }}</option>
                  <option v-for="o in styleOptions" :key="o.id" :value="o.id">{{ o.label }}</option>
                </select>
              </div>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('樣式') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('編輯')">
            <div class="dx-rbody dx-stack">
              <button type="button" class="dx-sm" :title="tl('尋找 (Ctrl+F)')" @click="emit('find', false)">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.find" /><span>{{ tl('尋找') }}</span>
              </button>
              <button type="button" class="dx-sm" :title="tl('取代 (Ctrl+H)')" @click="emit('find', true)">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.replace" /><span>{{ tl('取代') }}</span>
              </button>
              <button type="button" class="dx-sm" :title="tl('全選 (Ctrl+A)')" @click="selectEverything">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.selectAll" /><span>{{ tl('全選') }}</span>
              </button>

            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('編輯') }}</div>
          </div>
        </div>

        <!-- ===== 插入 ===== -->
        <div v-show="tab === 'insert'" class="dx-panel">
          <div class="dx-rgroup" role="group" :aria-label="tl('頁面')">
            <div class="dx-rbody">
              <button
                type="button"
                class="dx-big"
                :title="tl('插入空白頁：在游標處加入一整頁空白，游標後面的文字移到下一頁（同 Word「插入 › 空白頁」）')"
                :disabled="!inBody"
                @click="run(insertBlankPage)"
              >
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.blankPage" /><span>{{ tl('空白頁') }}</span>
              </button>
              <button type="button" class="dx-big" :title="tl('插入分頁符號 (Ctrl+Enter)')" :disabled="!inBody" @click="run(insertPageBreak)">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.pageBreak" /><span>{{ tl('分頁符號') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('頁面') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('表格')">
            <div class="dx-rbody">
              <div class="dx-pop">
                <button ref="tableButton" type="button" class="dx-big" :title="tl('插入表格')" :aria-expanded="tableOpen" @click="toggleTable">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.table" /><span>{{ tl('表格 ▾') }}</span>
                </button>
                <div v-if="tableOpen" class="dx-grid dx-popup" @mouseleave="hover = { r: 0, c: 0 }" @keydown="onGridKey">
                  <form v-if="tableForm" class="dx-tform" role="group" :aria-label="tl('插入表格')" @submit.prevent="submitTableForm">
                    <label>{{ tl('欄數') }} <input v-model.number="tableCols" type="number" min="1" :max="MAX_TABLE_COLS" /></label>
                    <label>{{ tl('列數') }} <input v-model.number="tableRows" type="number" min="1" :max="MAX_TABLE_ROWS" /></label>
                    <div class="dx-tform-actions">
                      <button type="submit">{{ tl('確定') }}</button>
                      <button type="button" @click="tableForm = false">{{ tl('取消') }}</button>
                    </div>
                  </form>
                  <template v-else>
                  <div class="dx-grid-label">{{ hover.r ? tl('{0} × {1} 表格', hover.r, hover.c) : tl('插入表格') }}</div>
                  <div ref="gridCells" class="dx-grid-cells">
                    <template v-for="r in 8" :key="r">
                      <button
                        v-for="c in 8"
                        :key="c"
                        type="button"
                        class="dx-cell"
                        :class="{ on: r <= hover.r && c <= hover.c }"
                        :tabindex="r === gridStop.r && c === gridStop.c ? 0 : -1"
                        :aria-label="tl('{0} 列 {1} 欄', r, c)"
                        @mouseenter="hover = { r, c }"
                        @focus="hover = { r, c }"
                        @click="pickTable(r, c)"
                      />
                    </template>
                  </div>
                  <button type="button" class="dx-grid-more" :title="tl('指定欄數與列數（最多 63 欄）')" @click="openTableForm">{{ tl('插入表格…') }}</button>
                  </template>
                </div>
              </div>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('表格') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('圖例')">
            <div class="dx-rbody">
              <button type="button" class="dx-big" :title="tl('插入圖片')" @click="fileInput?.click()">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.picture" /><span>{{ tl('圖片') }}</span>
              </button>
              <input ref="fileInput" type="file" accept="image/*" multiple hidden @change="onImage" />
              <div class="dx-pop">
                <button type="button" class="dx-big" :title="tl('插入圖案：選好後在頁面上按一下或拖曳')" aria-haspopup="menu" :aria-expanded="menu === 'shapes'" @click="toggleMenu('shapes', $event)">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.shapes" /><span>{{ tl('圖案 ▾') }}</span>
                </button>
                <div v-if="menu === 'shapes'" class="dx-menu dx-popup dx-shape-menu" role="menu" :aria-label="tl('圖案')" @keydown="onMenuKey">
                  <template v-for="g in SHAPE_MENU" :key="g.label">
                    <div class="dx-menu-head" role="presentation">{{ tl(g.label) }}</div>
                    <button v-for="k in g.kinds" :key="k.id" type="button" role="menuitem" tabindex="-1" :title="tl('插入{0}', tl(k.label))" @click="pick(() => insertShape(k.id))">{{ tl(k.label).replace(/^[^:：]*[:：]\s*/, '') }}</button>
                  </template>
                </div>
              </div>
              <button type="button" class="dx-big" :title="tl('插入文字方塊：在頁面上按一下或拖曳')" @click="insertShape('textbox')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.textbox" /><span>{{ tl('文字方塊') }}</span>
              </button>
              <button type="button" class="dx-big" :title="tl('選取游標後的圖形或文字方塊（再按 Tab 到下一個，方向鍵移動，Enter 編輯文字，Esc 回到文字）')" @click="editor?.selectShape(1)">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.select" /><span>{{ tl('選取圖形') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('圖例') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('連結')">
            <div class="dx-rbody">
              <button type="button" class="dx-big" :title="tl('插入連結')" :class="{ on: !!s?.link }" @click="onLink">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.link" /><span>{{ tl('連結') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('連結') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('留言')">
            <div class="dx-rbody">
              <button type="button" class="dx-big" :title="tl('在選取的文字上新增留言 (Ctrl+Alt+M)')" :disabled="!inBody" @click="emit('comment')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.comment" /><span>{{ tl('留言') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('留言') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('頁首及頁尾')">
            <div class="dx-rbody">
              <button type="button" class="dx-big" :title="tl('編輯頁首（也可以在頁面上方空白處按兩下）')" @click="editHeaderFooter('header')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.header" /><span>{{ tl('頁首') }}</span>
              </button>
              <button type="button" class="dx-big" :title="tl('編輯頁尾（也可以在頁面下方空白處按兩下）')" @click="editHeaderFooter('footer')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.footer" /><span>{{ tl('頁尾') }}</span>
              </button>
              <div class="dx-pop">
                <button type="button" class="dx-big" :title="tl('頁碼')" aria-haspopup="menu" :aria-expanded="menu === 'pageNumber'" @click="toggleMenu('pageNumber', $event)">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.pageNumber" /><span>{{ tl('頁碼 ▾') }}</span>
                </button>
                <div v-if="menu === 'pageNumber'" class="dx-menu dx-popup" role="menu" :aria-label="tl('頁碼')" @keydown="onMenuKey">
                  <button type="button" role="menuitem" tabindex="-1" :title="tl('在頁尾插入頁碼（頁尾原本空白時置中）')" @click="pick(pageNumberAtBottom)">{{ tl('頁面底端') }}</button>
                  <button
                    type="button"
                    role="menuitem"
                    tabindex="-1"
                    :title="tl('在游標處插入目前頁碼（編輯頁首或頁尾時）')"
                    :disabled="inBody"
                    @click="pick(() => editor?.insertField('PAGE'))"
                  >{{ tl('目前位置') }}</button>
                  <button type="button" role="menuitem" tabindex="-1" :title="tl('頁碼的格式、起始頁碼與首頁不同')" @click="pick(() => emit('page-setup'))">{{ tl('頁碼格式…') }}</button>
                </div>
              </div>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('頁首及頁尾') }}</div>
          </div>

          <!-- 符號: the gallery (SymbolGallery.vue) inserts at the cursor; 其他符號 opens the 符號 dialog. -->
          <div class="dx-rgroup" role="group" :aria-label="tl('符號')">
            <div class="dx-rbody">
              <SymbolGallery v-slot="{ open, toggle }" :editor="editor" :disabled="readOnly" @more="emit('symbol')">
                <button
                  type="button"
                  class="dx-big"
                  :title="tl('符號：插入鍵盤上沒有的符號，例如 ※、①、℃')"
                  aria-haspopup="menu"
                  :aria-expanded="open"
                  :disabled="readOnly"
                  @click="toggle"
                >
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.symbol" /><span>{{ tl('符號 ▾') }}</span>
                </button>
              </SymbolGallery>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('符號') }}</div>
          </div>
        </div>

        <!-- ===== 設計 ===== -->
        <div v-show="tab === 'design'" class="dx-panel">
          <!-- 頁面背景, as in Word: 浮水印 (頁面色彩 and 頁面框線 go next to it). -->
          <div class="dx-rgroup" role="group" :aria-label="tl('頁面背景')">
            <div class="dx-rbody">
              <div class="dx-pop">
                <button
                  type="button"
                  class="dx-big"
                  :title="tl('浮水印：在每一頁的文字後面加上「機密」、「草稿」等淡色文字或圖片')"
                  aria-haspopup="menu"
                  :aria-expanded="menu === 'watermark'"
                  @click="toggleMenu('watermark', $event)"
                >
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.watermark" /><span>{{ tl('浮水印 ▾') }}</span>
                </button>
                <div v-if="menu === 'watermark'" class="dx-menu dx-popup dx-wm-menu" role="menu" :aria-label="tl('浮水印')" @keydown="onMenuKey">
                  <div class="dx-wm-gallery">
                    <button
                      v-for="p in WATERMARK_PRESETS"
                      :key="p"
                      type="button"
                      role="menuitem"
                      tabindex="-1"
                      class="dx-wm-tile"
                      :title="tl('浮水印「{0}」：斜向、半透明，每一頁都有', p)"
                      :aria-label="p"
                      @click="pick(() => presetWatermark(p))"
                    >
                      <span class="dx-wm-sheet" aria-hidden="true"><span class="dx-wm-word">{{ p }}</span></span>
                      <span class="dx-wm-name">{{ p }}</span>
                    </button>
                  </div>
                  <button type="button" role="menuitem" tabindex="-1" :title="tl('選擇文字、字型、色彩或圖片')" @click="pick(() => emit('watermark'))">{{ tl('自訂浮水印…') }}</button>
                  <button type="button" role="menuitem" tabindex="-1" :title="tl('移除每一頁的浮水印')" :disabled="!hasWatermark()" @click="pick(() => editor?.setWatermark(null))">{{ tl('移除浮水印') }}</button>
                </div>
              </div>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('頁面背景') }}</div>
          </div>
        </div>

        <!-- ===== 版面配置 ===== -->
        <div v-show="tab === 'layout'" class="dx-panel">
          <div class="dx-rgroup" role="group" :aria-label="tl('版面設定')">
            <div class="dx-rbody">
              <div class="dx-pop">
                <button type="button" class="dx-big" :title="tl('邊界')" aria-haspopup="menu" :aria-expanded="menu === 'margins'" @click="toggleMenu('margins', $event)">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.margins" /><span>{{ tl('邊界 ▾') }}</span>
                </button>
                <div v-if="menu === 'margins'" class="dx-menu dx-popup" role="menu" :aria-label="tl('邊界')" @keydown="onMenuKey">
                  <button
                    v-for="m in MARGIN_PRESETS"
                    :key="m.id"
                    type="button"
                    role="menuitemradio"
                    tabindex="-1"
                    :aria-checked="isMargins(cursorPage(), m)"
                    @click="pick(() => setPage(marginsOf(m)))"
                  ><b>{{ tl(m.label) }}</b><small>{{ tl(m.note) }}</small></button>
                  <button type="button" role="menuitem" tabindex="-1" @click="pick(() => emit('page-setup'))">{{ tl('自訂邊界…') }}</button>
                </div>
              </div>
              <div class="dx-pop">
                <button type="button" class="dx-big" :title="tl('方向')" aria-haspopup="menu" :aria-expanded="menu === 'orientation'" @click="toggleMenu('orientation', $event)">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.orientation" /><span>{{ tl('方向 ▾') }}</span>
                </button>
                <div v-if="menu === 'orientation'" class="dx-menu dx-popup" role="menu" :aria-label="tl('方向')" @keydown="onMenuKey">
                  <button type="button" role="menuitemradio" tabindex="-1" :aria-checked="!!cursorPage() && cursorPage()!.width <= cursorPage()!.height" @click="pick(() => setOrientation(false))">{{ tl('直向') }}</button>
                  <button type="button" role="menuitemradio" tabindex="-1" :aria-checked="!!cursorPage() && cursorPage()!.width > cursorPage()!.height" @click="pick(() => setOrientation(true))">{{ tl('橫向') }}</button>
                </div>
              </div>
              <div class="dx-pop">
                <button type="button" class="dx-big" :title="tl('大小（紙張）')" aria-haspopup="menu" :aria-expanded="menu === 'size'" @click="toggleMenu('size', $event)">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.size" /><span>{{ tl('大小 ▾') }}</span>
                </button>
                <div v-if="menu === 'size'" class="dx-menu dx-popup" role="menu" :aria-label="tl('大小')" @keydown="onMenuKey">
                  <button
                    v-for="p in PAPERS"
                    :key="p.id"
                    type="button"
                    role="menuitemradio"
                    tabindex="-1"
                    :aria-checked="isPaper(cursorPage(), p)"
                    @click="pick(() => setPaper(p))"
                  >{{ tl(p.label) }}</button>
                  <button type="button" role="menuitem" tabindex="-1" @click="pick(() => emit('page-setup'))">{{ tl('其他紙張大小…') }}</button>
                </div>
              </div>
              <div class="dx-pop">
                <button type="button" class="dx-big" :title="tl('欄（分欄）')" aria-haspopup="menu" :aria-expanded="menu === 'columns'" @click="toggleMenu('columns', $event)">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.columns" /><span>{{ tl('欄 ▾') }}</span>
                </button>
                <div v-if="menu === 'columns'" class="dx-menu dx-popup" role="menu" :aria-label="tl('欄')" @keydown="onMenuKey">
                  <button
                    v-for="(label, i) in [tl('一欄'), tl('二欄'), tl('三欄')]"
                    :key="label"
                    type="button"
                    role="menuitemradio"
                    tabindex="-1"
                    :aria-checked="cursorColumns() === i + 1"
                    @click="pick(() => setColumns(i + 1))"
                  >{{ label }}</button>
                  <button type="button" role="menuitem" tabindex="-1" @click="pick(() => emit('page-setup'))">{{ tl('其他欄…') }}</button>
                </div>
              </div>
              <div class="dx-pop">
                <button type="button" class="dx-big" :title="tl('分隔設定：分頁符號、分欄符號、分節符號')" aria-haspopup="menu" :aria-expanded="menu === 'breaks'" @click="toggleMenu('breaks', $event)">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.breaks" /><span>{{ tl('分隔設定 ▾') }}</span>
                </button>
                <div v-if="menu === 'breaks'" class="dx-menu dx-popup" role="menu" :aria-label="tl('分隔設定')" @keydown="onMenuKey">
                  <button type="button" role="menuitem" tabindex="-1" :title="tl('插入分頁符號 (Ctrl+Enter)')" :disabled="!inBody" @click="pick(() => run(insertPageBreak))">
                    <b>{{ tl('分頁符號') }}</b><small>{{ tl('後面的文字從下一頁開始') }}</small>
                  </button>
                  <button
                    type="button"
                    role="menuitem"
                    tabindex="-1"
                    :title="tl('插入分欄符號 (Ctrl+Shift+Enter)：後面的文字從下一欄開始')"
                    :disabled="!inBody || s?.inTable"
                    @click="pick(() => editor?.insertColumnBreak())"
                  ><b>{{ tl('分欄符號') }}</b><small>{{ tl('後面的文字從下一欄開始') }}</small></button>
                  <button
                    type="button"
                    role="menuitem"
                    tabindex="-1"
                    :title="tl('插入分節符號（下一頁）：游標前後可以有不同的紙張方向、邊界與頁首頁尾')"
                    :disabled="!inBody"
                    @click="pick(() => editor?.insertSectionBreak())"
                  ><b>{{ tl('分節符號（下一頁）') }}</b><small>{{ tl('新的一節從下一頁開始，可以有不同的版面') }}</small></button>
                  <button
                    v-if="s?.atSectionBreak"
                    type="button"
                    role="menuitem"
                    tabindex="-1"
                    :title="tl('移除這個分節符號：這一節併入下一節，改用下一節的版面與頁首頁尾')"
                    @click="pick(() => editor?.removeSectionBreak())"
                  ><b>{{ tl('移除分節符號') }}</b><small>{{ tl('這一節併入下一節') }}</small></button>
                </div>
              </div>
            </div>
            <div class="dx-rlabel">
              <span aria-hidden="true">{{ tl('版面設定') }}</span>
              <button type="button" class="dx-launch" :title="tl('版面設定（紙張大小、方向、邊界）')" :aria-label="tl('版面設定')" @click="emit('page-setup')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.launcher" />
              </button>
            </div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('段落')">
            <div class="dx-rbody">
              <button type="button" class="dx-big" :title="tl('縮排與間距：左右縮排、段前段後距離、行距')" @click="emit('paragraph')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.paragraph" /><span>{{ tl('縮排與間距') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('段落') }}</div>
          </div>
        </div>

        <!-- ===== 參考資料 ===== -->
        <div v-show="tab === 'references'" class="dx-panel">
          <div class="dx-rgroup" role="group" :aria-label="tl('目錄')">
            <div class="dx-rbody">
              <button
                type="button"
                class="dx-big"
                :title="tl('目錄：在游標處插入依標題 1～3 建立的目錄（和 Word 的目錄相同，在 Word 按 F9 也能更新）')"
                @click="emit('insert-toc')"
              >
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.tocInsert" /><span>{{ tl('目錄') }}</span>
              </button>
              <button
                type="button"
                class="dx-big"
                :title="tocCount ? tl('更新目錄：依目前的標題重建目錄，並重新計算目錄與交互參照的頁碼') : tl('更新目錄：這份文件沒有目錄或頁碼參照')"
                :disabled="!tocCount"
                @click="emit('update-toc')"
              >
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.toc" /><span>{{ tl('更新目錄') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('目錄') }}</div>
          </div>
        </div>

        <!-- ===== 校閱 ===== -->
        <div v-show="tab === 'review'" class="dx-panel">
          <div class="dx-rgroup" role="group" :aria-label="tl('校對')">
            <div class="dx-rbody">
              <button
                type="button"
                class="dx-big"
                :class="{ on: spellcheck }"
                :aria-pressed="spellcheck"
                :title="tl('拼字檢查：使用瀏覽器的拼字檢查（主要檢查英文；多數瀏覽器不檢查中文）')"
                @click="emit('toggle-spellcheck')"
              >
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.spell" /><span>{{ tl('拼字檢查') }}</span>
              </button>
              <div class="dx-pop dx-rcount-pop">
                <button type="button" class="dx-big" :title="tl('字數統計')" :aria-expanded="countOpen" @click="countOpen = !countOpen">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.wordCount" /><span>{{ tl('字數統計') }}</span>
                </button>
                <div v-if="countOpen && counts" class="dx-rcount dx-popup" role="dialog" :aria-label="tl('字數統計')" @keydown.esc="!composing($event) && closeCount($event)">
                  <table>
                    <tbody>
                      <tr><th>{{ tl('頁數') }}</th><td>{{ n(s?.pageCount ?? 1) }}</td></tr>
                      <tr><th>{{ tl('字數') }}</th><td>{{ n(counts.all.words) }}</td></tr>
                      <tr><th>{{ tl('字元（不含空白）') }}</th><td>{{ n(counts.all.chars) }}</td></tr>
                      <tr><th>{{ tl('字元（含空白）') }}</th><td>{{ n(counts.all.charsWithSpaces) }}</td></tr>
                      <tr v-if="counts.selected"><th>{{ tl('選取範圍字數') }}</th><td>{{ n(counts.selected.words) }}</td></tr>
                    </tbody>
                  </table>
                  <p>{{ tl('本文的統計，不含頁首頁尾。') }}</p>
                  <button type="button" @click="countOpen = false">{{ tl('關閉') }}</button>
                </div>
              </div>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('校對') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('留言')">
            <div class="dx-rbody">
              <button type="button" class="dx-big" :title="tl('在選取的文字上新增留言 (Ctrl+Alt+M)')" :disabled="!inBody" @click="emit('comment')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.comment" /><span>{{ tl('新增留言') }}</span>
              </button>
              <button
                type="button"
                class="dx-big"
                :class="{ on: commentsOpen }"
                :aria-pressed="commentsOpen"
                :disabled="!commentCount"
                :title="commentCount ? tl('顯示或隱藏留言（共 {0} 則）', commentCount) : tl('這份文件沒有留言')"
                @click="emit('toggle-comments')"
              >
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.comments" /><span>{{ tl('顯示留言') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('留言') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('追蹤')">
            <div class="dx-rbody">
              <button
                type="button"
                class="dx-big"
                :class="{ on: s?.trackChanges }"
                :aria-pressed="!!s?.trackChanges"
                :aria-label="tl('追蹤修訂')"
                :title="tl('追蹤修訂 (Ctrl+Shift+E)：開啟時，輸入、刪除與格式變更都會記錄為修訂（作者與時間），可在 Word 中檢視、接受或拒絕')"
                @click="editor?.setTrackChanges(!s?.trackChanges)"
              >
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.track" /><span>{{ tl('追蹤修訂') }}</span>
              </button>
              <button
                type="button"
                class="dx-big"
                :class="{ on: rev.show }"
                :aria-pressed="rev.show"
                :title="rev.show ? tl('隱藏修訂標記（顯示接受全部修訂後的樣子）') : tl('顯示修訂標記')"
                @click="run(toggleRevisionMarks)"
              >
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.markup" /><span>{{ tl('顯示標記') }}</span>
              </button>
              <button
                type="button"
                class="dx-big"
                :class="{ on: revisionsOpen }"
                :aria-pressed="revisionsOpen"
                :title="tl('修訂窗格：列出每一處修訂的作者、類型、內容與時間，可逐一接受或拒絕')"
                @click="emit('toggle-revisions')"
              >
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.comments" /><span>{{ tl('修訂窗格') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('追蹤') }}</div>
          </div>

          <div class="dx-rgroup dx-review" role="group" :aria-label="tl('變更')">
            <div class="dx-rbody">
              <button type="button" class="dx-big" :title="tl('接受游標處或選取範圍中的修訂')" :disabled="!rev.here" @click="run(acceptRevision)">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.accept" /><span>{{ tl('接受') }}</span>
              </button>
              <button type="button" class="dx-big" :title="tl('拒絕游標處或選取範圍中的修訂')" :disabled="!rev.here" @click="run(rejectRevision)">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.reject" /><span>{{ tl('拒絕') }}</span>
              </button>
              <div class="dx-stack">
                <button type="button" class="dx-sm" :title="tl('上一個修訂')" :aria-label="tl('上一個修訂')" :disabled="!rev.count" @click="goRevision(-1)">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.prev" /><span>{{ tl('上一個') }}</span>
                </button>
                <button type="button" class="dx-sm" :title="tl('下一個修訂')" :aria-label="tl('下一個修訂')" :disabled="!rev.count" @click="goRevision(1)">
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.next" /><span>{{ tl('下一個') }}</span>
                </button>
                <span v-if="rev.count" class="dx-review-label" :title="tl('這部分有 {0} 處修訂', rev.count)">{{ tl('修訂 {0}', rev.count) }}</span>
                <span class="dx-sr" aria-live="polite">{{ spoken }}</span>
              </div>
              <div class="dx-stack">
                <button type="button" class="dx-sm" :title="tl('接受這部分的所有修訂')" :disabled="!rev.count" @click="run(acceptAllRevisions)">{{ tl('全部接受') }}</button>
                <button type="button" class="dx-sm" :title="tl('拒絕這部分的所有修訂')" :disabled="!rev.count" @click="run(rejectAllRevisions)">{{ tl('全部拒絕') }}</button>
              </div>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('變更') }}</div>
          </div>
        </div>

        <!-- ===== 檢視 ===== -->
        <div v-show="tab === 'view'" class="dx-panel">
          <div class="dx-rgroup" role="group" :aria-label="tl('顯示')">
            <div class="dx-rbody">
              <button
                type="button"
                class="dx-big"
                :class="{ on: outlineOpen }"
                :aria-pressed="outlineOpen"
                :title="tl('導覽窗格：依標題瀏覽文件')"
                @click="emit('toggle-outline')"
              >
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.nav" /><span>{{ tl('導覽窗格') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('顯示') }}</div>
          </div>

          <div class="dx-rgroup" role="group" :aria-label="tl('縮放')">
            <div class="dx-rbody">
              <label class="dx-big dx-zoom" :title="tl('縮放')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.zoom" />
                <select :value="zoom" :aria-label="tl('縮放比例')" @change="emit('zoom', ($event.target as HTMLSelectElement).value)">
                  <option v-for="z in ZOOMS" :key="z" :value="String(z)">{{ Math.round(z * 100) }}%</option>
                  <option value="fit">{{ tl('頁寬') }}</option>
                </select>
              </label>
              <button type="button" class="dx-big" :title="tl('縮放到 100%')" :class="{ on: zoom === '1' }" :aria-pressed="zoom === '1'" @click="emit('zoom', '1')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.zoom100" /><span>100%</span>
              </button>
              <button type="button" class="dx-big" :title="tl('頁寬：頁面和視窗一樣寬')" :class="{ on: zoom === 'fit' }" :aria-pressed="zoom === 'fit'" @click="emit('zoom', 'fit')">
                <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.pageWidth" /><span>{{ tl('頁寬') }}</span>
              </button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('縮放') }}</div>
          </div>
        </div>

        <!-- ===== 表格設計 / 表格版面配置 (contextual) ===== -->
        <div v-show="tab === 'tableDesign'" class="dx-panel">
          <div class="dx-rgroup" role="group" :aria-label="tl('表格樣式')">
            <div class="dx-rbody">
              <TablePanel v-if="s?.inTable" :editor="editor" :snapshot="snapshot" />
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('網底、框線與儲存格') }}</div>
          </div>
        </div>
        <div v-show="tab === 'tableLayout'" class="dx-panel">
          <div v-if="s?.inTable" class="dx-rgroup" role="group" :aria-label="tl('刪除')">
            <div class="dx-rbody dx-stack">
              <button type="button" class="dx-sm" :title="tl('刪除列')" @click="run(deleteRow)">{{ tl('刪除列') }}</button>
              <button type="button" class="dx-sm" :title="tl('刪除欄')" @click="run(deleteColumn)">{{ tl('刪除欄') }}</button>
              <button type="button" class="dx-sm" :title="tl('刪除表格')" @click="run(deleteTable)">{{ tl('刪除表格') }}</button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('刪除') }}</div>
          </div>
          <div v-if="s?.inTable" class="dx-rgroup" role="group" :aria-label="tl('列與欄')">
            <div class="dx-rbody">
              <button type="button" class="dx-big" :title="tl('在上方插入列')" @click="run(addRowBefore)"><span class="dx-glyph">⬆</span><span>{{ tl('上方插入') }}</span></button>
              <button type="button" class="dx-big" :title="tl('在下方插入列')" @click="run(addRowAfter)"><span class="dx-glyph">⬇</span><span>{{ tl('下方插入') }}</span></button>
              <button type="button" class="dx-big" :title="tl('在左側插入欄')" @click="run(addColumnBefore)"><span class="dx-glyph">⬅</span><span>{{ tl('左方插入') }}</span></button>
              <button type="button" class="dx-big" :title="tl('在右側插入欄')" @click="run(addColumnAfter)"><span class="dx-glyph">➡</span><span>{{ tl('右方插入') }}</span></button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('列與欄') }}</div>
          </div>
          <div v-if="s?.inTable" class="dx-rgroup" role="group" :aria-label="tl('合併')">
            <div class="dx-rbody dx-stack">
              <button
                type="button"
                class="dx-sm"
                :title="canMerge ? tl('合併選取的儲存格') : tl('請先選取兩個以上的儲存格')"
                :aria-description="canMerge ? undefined : tl('請先選取兩個以上的儲存格')"
                :disabled="!canMerge"
                @click="run(mergeCells)"
              >{{ tl('合併儲存格') }}</button>
              <button type="button" class="dx-sm" :title="tl('分割儲存格')" @click="run(splitCell)">{{ tl('分割儲存格') }}</button>
              <button type="button" class="dx-sm" :title="tl('分割表格：從游標所在的列起分成兩個表格 (Ctrl+Shift+Enter)')" @click="editor?.splitTable()">{{ tl('分割表格') }}</button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('合併') }}</div>
          </div>
          <div v-if="s?.inTable" class="dx-rgroup" role="group" :aria-label="tl('儲存格大小')">
            <div class="dx-rbody dx-stack">
              <div class="dx-pop">
                <button
                  type="button"
                  class="dx-sm"
                  :title="tl('自動調整：依內容或視窗寬度調整欄寬，或固定欄寬')"
                  aria-haspopup="menu"
                  :aria-expanded="menu === 'autoFit'"
                  @click="toggleMenu('autoFit', $event)"
                >
                  <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.autoFit" /><span>{{ tl('自動調整 ▾') }}</span>
                </button>
                <div v-if="menu === 'autoFit'" class="dx-menu dx-popup" role="menu" :aria-label="tl('自動調整')" @keydown="onMenuKey">
                  <button type="button" role="menuitem" tabindex="-1" :title="tl('依儲存格內容調整欄寬')" @click="pick(() => editor?.autoFitTable('contents'))">{{ tl('自動調整成內容大小') }}</button>
                  <button type="button" role="menuitem" tabindex="-1" :title="tl('讓表格和版面同寬')" @click="pick(() => editor?.autoFitTable('window'))">{{ tl('自動調整成視窗大小') }}</button>
                  <button type="button" role="menuitem" tabindex="-1" :title="tl('欄寬固定為目前的寬度')" @click="pick(() => editor?.autoFitTable('fixed'))">{{ tl('固定欄寬') }}</button>
                </div>
              </div>
              <button type="button" class="dx-sm" :title="tl('平均分配欄寬（選取多欄時只分配這些欄）')" @click="run(distributeColumns)">{{ tl('平均分配欄寬') }}</button>
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('儲存格大小') }}</div>
          </div>
        </div>

        <!-- ===== 圖形格式 (contextual) ===== -->
        <div v-show="tab === 'shape'" class="dx-panel">
          <div class="dx-rgroup" role="group" :aria-label="tl('圖形')">
            <div class="dx-rbody">
              <ShapePanel v-if="shapeSelected" :editor="editor" :snapshot="snapshot" />
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('圖形') }}</div>
          </div>
        </div>

        <!-- ===== 圖片格式 (contextual) ===== -->
        <div v-show="tab === 'picture'" class="dx-panel">
          <div class="dx-rgroup" role="group" :aria-label="tl('圖片')">
            <div class="dx-rbody">
              <ImagePanel v-if="imageSelected" :editor="editor" :snapshot="snapshot" />
            </div>
            <div class="dx-rlabel" aria-hidden="true">{{ tl('大小與替代文字') }}</div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.dx-ribbon {
  position: relative;
  z-index: 25;
  background: #fff;
  border-bottom: 1px solid #d2d0ce;
  /* In rem, so the ribbon grows with the browser's font size setting (persona-300 B-10). */
  font: 0.8125rem system-ui, 'Microsoft JhengHei', sans-serif;
  color: #323130;
}
/* ----- Tab row ----- */
.dx-tabrow {
  display: flex;
  align-items: stretch;
  gap: 2px;
  padding: 0 8px;
  background: #f3f2f1;
  border-bottom: 1px solid #e1dfdd;
}
.dx-tabs {
  display: flex;
  flex-wrap: wrap;
  flex: 1;
  min-width: 0;
}
.dx-file,
.dx-rtab {
  height: 2rem;
  padding: 0 12px;
  border: 0;
  border-radius: 0;
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: pointer;
  white-space: nowrap;
}
/* 檔案 in Word's blue, as in Word. */
.dx-file {
  background: #185abd;
  color: #fff;
  min-width: 56px;
}
.dx-file:hover,
.dx-file[aria-expanded='true'] {
  background: #124a9c;
}
.dx-rtab {
  position: relative;
}
.dx-rtab:hover {
  background: #e1dfdd;
}
.dx-rtab[aria-selected='true'] {
  color: #185abd;
  font-weight: 600;
}
.dx-rtab[aria-selected='true']::after {
  content: '';
  position: absolute;
  left: 10px;
  right: 10px;
  bottom: 2px;
  height: 3px;
  border-radius: 2px;
  background: #185abd;
}
/* Word's contextual tabs (表格工具, 圖片工具) stand out in their own colour. */
.dx-rtab.dx-context {
  color: #8764b8;
  background: #f4f0fa;
}
.dx-rtab.dx-context[aria-selected='true']::after {
  background: #8764b8;
}
.dx-collapse {
  align-self: center;
  width: 28px;
  height: 28px;
  padding: 0;
  border: 0;
  border-radius: 4px;
  background: transparent;
  color: inherit;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}
.dx-collapse:hover {
  background: #e1dfdd;
}
.dx-file:focus-visible,
.dx-rtab:focus-visible,
.dx-collapse:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: -2px;
}
/* A collapsed ribbon shows a tab's panel over the page. */
.dx-peek .dx-tabpanel {
  position: absolute;
  left: 0;
  right: 0;
  top: 100%;
  background: #fff;
  border-bottom: 1px solid #d2d0ce;
  box-shadow: 0 6px 12px rgba(0, 0, 0, 0.18);
}
/* ----- Panel ----- */
.dx-toolbar {
  padding: 4px 8px 0;
}
/* Groups wrap onto more rows as a whole (GOV-FINDING-011: no sideways scrolling). */
.dx-panel {
  display: flex;
  flex-wrap: wrap;
  align-items: stretch;
  min-height: 5.5rem;
}
.dx-rgroup {
  display: flex;
  flex-direction: column;
  min-width: 0;
  padding: 0 8px;
  border-right: 1px solid #e1dfdd;
  margin-bottom: 4px;
}
.dx-rgroup:last-child {
  border-right: 0;
}
.dx-rbody {
  flex: 1;
  display: flex;
  flex-wrap: wrap;
  align-items: flex-start;
  gap: 2px;
}
.dx-rows {
  flex-direction: column;
  gap: 4px;
  padding-top: 2px;
}
.dx-rrow {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 2px;
}
.dx-stack {
  display: flex;
  flex-direction: column;
  align-items: stretch;
  gap: 1px;
}
.dx-rlabel {
  position: relative;
  padding: 2px 1.75rem 3px;
  text-align: center;
  font-size: 0.75rem;
  color: #605e5c;
  white-space: nowrap;
}
/* The small ↘ button at a group's corner opens its dialog (Word's dialog box launcher). */
button.dx-launch {
  position: absolute;
  right: -6px;
  bottom: 0;
  height: 1.5rem;
  min-width: 1.5rem;
  width: 1.5rem;
  padding: 0;
}
.dx-launch .dx-ico {
  width: 12px;
  height: 12px;
}
/* Tools in a group: the table and picture tools, too. */
.dx-group {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 2px;
  min-width: 0;
}
button,
select {
  height: 1.75rem;
  min-width: 1.75rem;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 0 6px;
}
select {
  border-color: #c8c6c4;
  padding: 0 4px;
}
button:hover:not(:disabled) {
  background: #f3f2f1;
}
button:disabled {
  opacity: 0.35;
  cursor: default;
}
button.on {
  background: #deecf9;
  color: #185abd;
}
button:focus-visible,
select:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}
svg {
  width: 16px;
  height: 16px;
  stroke: currentColor;
  stroke-width: 1.5;
  fill: currentColor;
}
svg path {
  fill: none;
}
/* The ribbon's own icons (ribbon.ts ICONS). */
svg.dx-ico {
  width: 18px;
  height: 18px;
  flex: none;
  fill: none;
  stroke-width: 1.6;
  stroke-linecap: round;
  stroke-linejoin: round;
}
.dx-ico :deep(.f) {
  fill: currentColor;
  stroke: none;
}
.dx-ico :deep(.w) {
  fill: #fff;
}
.dx-ico :deep(text) {
  font-family: system-ui, sans-serif;
  font-weight: 600;
}
/* A large button: icon over its name. */
.dx-big {
  flex-direction: column;
  justify-content: flex-start;
  gap: 2px;
  height: 4rem;
  min-width: 3.25rem;
  padding: 6px 4px 2px;
  font-size: 0.75rem;
  line-height: 1.2;
  white-space: nowrap;
}
.dx-big .dx-ico {
  width: 26px;
  height: 26px;
}
.dx-glyph {
  font-size: 20px;
  line-height: 26px;
}
/* A small button: icon beside its name; at least 24 px high, as a finger or a shaky hand needs. */
.dx-sm {
  justify-content: flex-start;
  gap: 4px;
  height: 1.5rem;
  padding: 0 6px 0 4px;
  font-size: 0.75rem;
  white-space: nowrap;
}
.dx-sm .dx-ico {
  width: 16px;
  height: 16px;
}
.dx-zoom {
  display: inline-flex;
  align-items: center;
  justify-content: flex-start;
}
.dx-zoom select {
  height: 1.625rem;
  font-size: 0.75rem;
}
.dx-w-style {
  width: 8.25rem;
}
.dx-w-font {
  width: 8rem;
}
.dx-w-size {
  width: 4rem;
}
/* 樣式 gallery: each style shown as it looks. */
.dx-gallery .dx-style {
  height: 1.875rem;
  min-width: 3.25rem;
  border-color: #e1dfdd;
  background: #fff;
  font-size: 0.75rem;
}
.dx-gallery .dx-style.on {
  border-color: #185abd;
  background: #deecf9;
}
.dx-style-heading1 {
  font-weight: 700;
  font-size: 14px !important;
  color: #2f5496;
}
.dx-style-heading2 {
  font-weight: 700;
  color: #2f5496;
}
.dx-style-heading3 {
  color: #1f3763;
}
.dx-style-title {
  font-size: 15px !important;
}
.dx-cp-glyph {
  font-weight: 600;
  line-height: 1;
}
.dx-cp-glyph.dx-hl {
  padding: 0 2px;
}
/* ----- Pop-ups: menus, the table grid, the word count ----- */
.dx-pop {
  position: relative;
  display: inline-flex;
}
.dx-menu,
.dx-grid,
.dx-rcount {
  position: absolute;
  z-index: 40;
  top: calc(100% + 2px);
  left: 0;
  background: #fff;
  border: 1px solid #c8c6c4;
  border-radius: 6px;
  box-shadow: 0 6px 16px rgba(0, 0, 0, 0.18);
}
.dx-menu {
  display: flex;
  flex-direction: column;
  min-width: 180px;
  padding: 4px;
}
.dx-menu button {
  height: auto;
  min-height: 30px;
  justify-content: flex-start;
  flex-wrap: wrap;
  gap: 2px 8px;
  padding: 4px 10px 4px 26px;
  text-align: left;
  white-space: nowrap;
  position: relative;
}
.dx-menu button .dx-ico {
  position: absolute;
  left: 4px;
}
.dx-menu button[aria-checked='true']::before {
  content: '✓';
  position: absolute;
  left: 9px;
  color: #185abd;
  font-weight: 700;
}
.dx-menu small {
  flex-basis: 100%;
  color: #605e5c;
  font-size: 11px;
}
/* 浮水印's gallery: each preset on a small page, as Word's 浮水印 menu shows them. */
.dx-wm-gallery {
  display: grid;
  grid-template-columns: repeat(4, 76px);
  gap: 4px;
  padding: 2px 2px 6px;
  border-bottom: 1px solid #e1dfdd;
  margin-bottom: 4px;
}
.dx-menu .dx-wm-tile {
  flex-direction: column;
  flex-wrap: nowrap;
  align-items: center;
  gap: 2px;
  padding: 4px;
  white-space: normal;
}
.dx-wm-sheet {
  position: relative;
  display: grid;
  place-items: center;
  width: 44px;
  height: 58px;
  overflow: hidden;
  background: #fff;
  border: 1px solid #c8c6c4;
}
.dx-wm-word {
  transform: rotate(-45deg);
  color: #a6a6a6;
  font-family: '標楷體', 'DFKai-SB', serif;
  font-size: 13px;
  white-space: nowrap;
}
.dx-wm-name {
  font-size: 12px;
}
.dx-grid {
  padding: 8px;
}
.dx-grid-label {
  margin-bottom: 6px;
  text-align: center;
}
.dx-grid-cells {
  display: grid;
  grid-template-columns: repeat(8, 16px);
  gap: 2px;
}
.dx-cell {
  width: 16px;
  height: 16px;
  min-width: 0;
  padding: 0;
  border: 1px solid #c8c6c4;
  border-radius: 2px;
}
.dx-grid-more {
  width: 100%;
  margin-top: 6px;
  border-color: #c8c6c4;
}
.dx-tform {
  display: flex;
  flex-direction: column;
  gap: 6px;
  min-width: 10rem;
}
.dx-tform label {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 8px;
}
.dx-tform input {
  width: 4.5rem;
  height: 1.75rem;
  font: inherit;
}
.dx-tform-actions {
  display: flex;
  justify-content: flex-end;
  gap: 4px;
}
.dx-tform-actions button {
  border-color: #c8c6c4;
}
.dx-cell.on {
  background: #d2e3fc;
  border-color: #1a73e8;
}
.dx-rcount {
  padding: 10px 12px;
  min-width: 220px;
}
.dx-rcount table {
  width: 100%;
  border-collapse: collapse;
}
.dx-rcount th {
  text-align: left;
  font-weight: normal;
  padding: 2px 12px 2px 0;
}
.dx-rcount td {
  text-align: right;
  font-variant-numeric: tabular-nums;
}
.dx-rcount p {
  margin: 6px 0;
  color: #605e5c;
  font-size: 12px;
}
.dx-rcount button {
  border-color: #c8c6c4;
}
.dx-sr {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
}
.dx-review-label {
  font-size: 12px;
  color: #605e5c;
  padding: 0 4px;
  white-space: nowrap;
}
.dx-shape-menu {
  max-height: 70vh;
  overflow-y: auto;
}
.dx-menu-head {
  padding: 4px 10px 2px;
  font-size: 11px;
  color: #5f6368;
}
</style>
