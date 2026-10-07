<script lang="ts">
/**
 * The browser's device pixel ratio when the editor was first loaded on this page. Zooming the
 * browser in (Ctrl + +) raises it; a sidebar, DevTools or a narrow window do not.
 */
const LOADED_PIXEL_RATIO = typeof window !== 'undefined' && window.devicePixelRatio > 0 ? window.devicePixelRatio : 1;
</script>

<script setup lang="ts">
import { tl } from './locale';
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, watch } from 'vue';
import { DocxEditor, LARGE_DOCUMENT_BYTES, promptLink, type EditorSnapshot } from '../editor/core';
import type { MissingFont } from '../docx/fonts';
import { isShowMarksKey } from '../editor/formattingMarks';
import type { ParagraphStyleInfo } from '../docx/model';
import DocxToolbar from './DocxToolbar.vue';
import FindReplace from './FindReplace.vue';
import CompatNotice from './CompatNotice.vue';
import CommentsPanel from './CommentsPanel.vue';
import RevisionsPanel from './RevisionsPanel.vue';
import FieldsPanel from './FieldsPanel.vue';
import OutlinePanel from './OutlinePanel.vue';
import PageSetupDialog from './PageSetupDialog.vue';
import ParagraphDialog from './ParagraphDialog.vue';
import SymbolDialog from './SymbolDialog.vue';
import WatermarkDialog from './WatermarkDialog.vue';
import ContextMenu from './ContextMenu.vue';
import { TextSelection } from 'prosemirror-state';
import { scanCompat, type CompatReport } from '../docx/compat';
import type { CommentAuthor, DocComment } from '../docx/comments';
import { commentAtSelection, commentedText, latestLoads, revealComment } from '../editor/review';
import { composing } from './keys';

const props = withDefaults(
  defineProps<{
    /** A .docx to open. When omitted, an empty document is created. */
    src?: Blob | ArrayBuffer | Uint8Array | null;
    editable?: boolean;
    toolbar?: boolean;
    /** The signed-in user, as the author of new comments (「使用者」 when not given). */
    author?: CommentAuthor | null;
    /**
     * The page has its own 檔案 menu (save, download, versions…): the ribbon's 「檔案」 emits
     * `file`, and the ribbon shows while viewing only, too (檔案 and 檢視).
     */
    fileMenu?: boolean;
    /** Read-only, but comments may be added, replied to and resolved (editor option `commenting`). */
    commenting?: boolean;
  }>(),
  { src: null, editable: true, toolbar: true, author: null, fileMenu: false, commenting: false },
);

const emit = defineEmits<{
  (e: 'ready', editor: DocxEditor): void;
  (e: 'change'): void;
  (e: 'error', error: unknown): void;
  /** What in the opened document the web editor can't fully show or edit (G01). */
  (e: 'compat', report: CompatReport): void;
  /** The ribbon's 「檔案」 (with `fileMenu`). */
  (e: 'file'): void;
  /**
   * About how big the saved .docx will be (DocxEditor.estimatedSize), after opening and after
   * changes; `large` past LARGE_DOCUMENT_BYTES (30 MB; the server takes 50 MB), for a warning.
   */
  (e: 'size', info: { bytes: number; large: boolean }): void;
  /** The opened document's fonts this computer doesn't have (DocxEditor.missingFonts), once its fonts loaded; [] when none. */
  (e: 'fonts', missing: MissingFont[]): void;
}>();

const mount = ref<HTMLElement | null>(null);
const editor = shallowRef<DocxEditor | null>(null);
const snapshot = shallowRef<EditorSnapshot | null>(null);
const styles = shallowRef<ParagraphStyleInfo[]>([]);
const loading = ref(false);

// Documents open one after another; once a newer src arrives, an older load is skipped or
// its results (comments, compatibility notice, errors) are dropped.
const queueLoad = latestLoads();
function load(src: typeof props.src) {
  // The previous document's notice and comments never stay on the next one.
  compat.value = null;
  commentDraft.value = null;
  commentsOpen.value = false;
  fieldsOpen.value = false;
  return queueLoad(async (current) => {
    const ed = editor.value;
    if (!ed) return;
    loading.value = true;
    try {
      if (src) await ed.open(src);
      else await ed.newDocument();
      if (!current()) return;
      styles.value = ed.model.paragraphStyles;
      lastSize = -1;
      if (editor.value) reportSize();
      // Which of its fonts this computer lacks, once the page's fonts have loaded (persona-300 B-8).
      void (document.fonts?.ready ?? Promise.resolve()).then(() => current() && emit('fonts', ed.missingFonts()));
      await loadReview(ed, current);
    } catch (err) {
      if (current()) emit('error', err);
    } finally {
      if (current()) loading.value = false;
    }
  });
}

onMounted(async () => {
  editor.value = new DocxEditor(mount.value!, {
    editable: props.editable,
    commenting: props.commenting,
    spellcheck: spellcheck.value,
    onNotice: showNotice,
    author: props.author,
    onUpdate: (snap) => (snapshot.value = snap),
    onChange: () => {
      emit('change');
      scheduleSize();
    },
  });
  const zoom = rememberedZoom();
  if (zoom) setZoomChoice(zoom);
  await load(props.src);
  emit('ready', editor.value);
  reportSize();
});

// The document's size, for the host's warning before it gets too big to save (persona-300 A-9):
// worked out once the edits pause, and sent when it changed.
let sizeTimer: ReturnType<typeof setTimeout> | undefined;
let lastSize = -1;
function reportSize() {
  clearTimeout(sizeTimer);
  const bytes = editor.value?.view ? editor.value.estimatedSize() : 0;
  if (!bytes || bytes === lastSize) return;
  lastSize = bytes;
  emit('size', { bytes, large: bytes > LARGE_DOCUMENT_BYTES });
}
function scheduleSize() {
  clearTimeout(sizeTimer);
  sizeTimer = setTimeout(reportSize, 800);
}

watch(
  () => props.src,
  (src) => load(src),
);
watch(
  () => props.commenting,
  (on) => editor.value?.setCommenting(on),
);
watch(
  () => props.author,
  (author) => editor.value?.setAuthor(author ?? null),
);

onBeforeUnmount(() => {
  clearTimeout(sizeTimer);
  editor.value?.destroy();
});

/** Returns the current document as a .docx Blob. */
async function save(): Promise<Blob> {
  if (!editor.value) throw new Error('Editor not ready');
  return editor.value.save();
}

/** Saves and triggers a browser download. */
async function download(filename = 'document.docx'): Promise<void> {
  const blob = await save();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename.endsWith('.docx') ? filename : filename + '.docx';
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Find / replace panel: toolbar button, Ctrl+F (find), Ctrl+H (replace).
const findOpen = ref(false);
const findReplace = ref(false);
const findPanel = ref<InstanceType<typeof FindReplace> | null>(null);
function openFind(replace: boolean) {
  if (!editor.value?.view) return;
  findReplace.value = replace || (findOpen.value && findReplace.value);
  if (findOpen.value) findPanel.value?.focus(replace);
  else findOpen.value = true;
}
/** Puts the keyboard in the document, where the cursor is. */
function focusDocument() {
  editor.value?.activeView?.focus();
}

// Closing the panel puts the keyboard back in the document, at the current hit or where the
// cursor was (GOV-ISSUE-013: focus was left on the page, 43 Tab presses away from the text).
function closeFind() {
  findOpen.value = false;
  nextTick(focusDocument);
}
// Alt+F10 (as in Word, TinyMCE, CKEditor) puts the keyboard on the toolbar; Escape there comes
// back to the text (DocxToolbar.vue).
const toolbarRef = ref<InstanceType<typeof DocxToolbar> | null>(null);
const showToolbar = computed(() => props.toolbar && (props.editable || props.fileMenu));
function onFindKey(e: KeyboardEvent) {
  if (composing(e)) return; // an input method's keys (注音 …) are never shortcuts
  // Shift+F10 / the menu key: the right-click menu at the cursor (GOV-184).
  if ((e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10' && !e.ctrlKey && !e.altKey && !e.metaKey)) && openMenuAtCursor()) {
    e.preventDefault();
    return;
  }
  if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && e.key === 'F10' && toolbarRef.value) {
    e.preventDefault();
    toolbarRef.value.focus();
    return;
  }
  // Ctrl+F1: 摺疊功能區, as in Word.
  if (e.ctrlKey && !e.altKey && !e.metaKey && !e.shiftKey && e.key === 'F1' && toolbarRef.value) {
    e.preventDefault();
    toolbarRef.value.toggleCollapsed();
    return;
  }
  // Ctrl+Shift+8: 顯示／隱藏編輯標記, from the ribbon as well as the text, as in Word. Handled here
  // first; the editor's own page-wide listener then sees defaultPrevented and leaves it.
  if (!e.defaultPrevented && !e.repeat && isShowMarksKey(e) && editor.value) {
    e.preventDefault();
    editor.value.toggleShowMarks();
    return;
  }
  // Ctrl+Alt+M: new comment, as in Word.
  if ((e.ctrlKey || e.metaKey) && e.altKey && !e.shiftKey && e.code === 'KeyM') {
    e.preventDefault();
    startComment();
    return;
  }
  if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;
  const key = e.key.toLowerCase();
  // Word's Ctrl+K (insert link) and Ctrl+D (font; the browser's bookmark otherwise) in the document.
  if (key === 'd') return e.preventDefault();
  if (key === 'k' && props.editable && editor.value?.activeView?.dom.contains(e.target as Node)) {
    e.preventDefault();
    // Not through the ribbon: it may be hidden (toolbar off).
    promptLink(editor.value);
    return;
  }
  if (key !== 'f' && (key !== 'h' || !props.editable)) return;
  e.preventDefault();
  openFind(key === 'h');
}

// GOV-184: Word's right-click menu over the text (ContextMenu.vue). Ctrl+right-click keeps the
// browser's own menu (its spelling suggestions).
const menuAt = ref<{ x: number; y: number } | null>(null);
/** When the menu was last opened from the keyboard: the menu key's own contextmenu event is ignored. */
let keyboardMenuAt = 0;
function onContextMenu(e: MouseEvent) {
  if (Date.now() - keyboardMenuAt < 500) {
    e.preventDefault();
    return;
  }
  const view = editor.value?.activeView;
  if (e.ctrlKey || e.metaKey || !view || !view.dom.contains(e.target as Node)) return;
  e.preventDefault();
  // Outside the selection: the cursor moves where the click was first, as in Word.
  let hit: { pos: number } | null = null;
  try {
    hit = view.posAtCoords({ left: e.clientX, top: e.clientY });
  } catch {
    // no position there (e.g. over a picture being laid out): the selection stays
  }
  const sel = view.state.selection;
  if (hit && (hit.pos < sel.from || hit.pos > sel.to)) {
    view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(hit.pos))));
  }
  menuAt.value = { x: e.clientX, y: e.clientY };
}
/** Opens the menu under the cursor, when the keyboard is in the text. */
function openMenuAtCursor(): boolean {
  const view = editor.value?.activeView;
  if (!view || !view.hasFocus()) return false;
  const c = view.coordsAtPos(view.state.selection.head);
  keyboardMenuAt = Date.now();
  menuAt.value = { x: c.left, y: c.bottom };
  return true;
}
function closeMenu(refocus: boolean) {
  menuAt.value = null;
  if (refocus) focusDocument();
}

// Compatibility notice (G01) and the comments panel (G09).
const compat = shallowRef<CompatReport | null>(null);
// The comments are on the editor's document (read when it opens; changes are undoable).
const comments = computed<DocComment[]>(() => {
  void snapshot.value; // re-evaluated on every editor update
  return editor.value?.comments() ?? [];
});
const commentsOpen = ref(false);
async function loadReview(ed: DocxEditor, current: () => boolean) {
  // A failed scan must never keep the document from opening.
  const scanned = await scanCompat(ed.model.zip).catch((): CompatReport => ({ items: [] }));
  if (!current()) return;
  // Parts that were not readable XML (docx/reader.ts brokenParts) are listed too.
  const broken = ed.model.brokenParts ?? [];
  const report: CompatReport = broken.length
    ? {
        ...scanned,
        items: [
          ...scanned.items,
          {
            id: 'broken-parts',
            title: '格式錯誤的內容',
            count: broken.length,
            where: broken.join('、'),
            effect: '這些部分無法讀取，網頁上略過不顯示；存檔時會原樣保留，不會修改。',
          },
        ],
      }
    : scanned;
  compat.value = report;
  // A template's fields to fill in come first; otherwise the comments, if any.
  fieldsOpen.value = ed.fields().length > 0;
  commentsOpen.value = !fieldsOpen.value && ed.comments().length > 0;
  emit('compat', report);
}
const activeComment = computed(() => {
  void snapshot.value; // re-evaluated on every editor update
  const state = editor.value?.view?.state;
  return state && comments.value.length ? commentAtSelection(state) : null;
});
const commentAnchors = computed(() => {
  void snapshot.value;
  const doc = editor.value?.view?.state.doc;
  const out: Record<string, string> = {};
  if (doc && commentsOpen.value) for (const c of comments.value) out[c.id] = commentedText(doc, c.id);
  return out;
});
/** Clicking commented text or a comment mark opens the comments panel. */
function onDocClick(e: MouseEvent) {
  if (comments.value.length && (e.target as HTMLElement).closest?.('[data-comment-id]')) commentsOpen.value = true;
}
function showComment(id: string) {
  const view = editor.value?.view;
  if (view) revealComment(view, id);
}

// Writing comments: 「新增留言」 (toolbar, panel, Ctrl+Alt+M) opens the panel with a box for the
// selection (or the word at the cursor); replies, resolving, editing and deleting go straight
// to the editor, one undo step each.
const commentDraft = shallowRef<{ from: number; to: number; doc: unknown; quote: string } | null>(null);
const commentAuthor = computed(() => {
  void snapshot.value;
  void props.author; // setAuthor() runs in the watcher below
  return editor.value?.author?.name ?? null;
});
function startComment() {
  const ed = editor.value;
  if ((!props.editable && !props.commenting) || !ed?.view) return;
  const target = ed.commentTarget();
  if (!target) {
    showNotice(ed.target === 'textbox' ? '文字方塊不能加入留言。' : ed.target !== 'body' ? '頁首頁尾不能加入留言。' : '這裡不能加入留言，請選取要留言的文字。');
    return;
  }
  const doc = ed.view.state.doc;
  commentDraft.value = { ...target, doc, quote: doc.textBetween(target.from, target.to, '\n', '') };
  commentsOpen.value = true;
}
function addComment(text: string) {
  const ed = editor.value;
  const draft = commentDraft.value;
  if (!ed?.view || !draft) return;
  if (ed.view.state.doc !== draft.doc) {
    // The text changed while the comment was being written: ask again where it goes.
    commentDraft.value = null;
    startComment();
    if (commentDraft.value) showNotice('文件已變更，請確認留言的位置後再送出。');
    return;
  }
  if (ed.addComment(text, { from: draft.from, to: draft.to }) == null) showNotice('無法在這裡加入留言。');
  commentDraft.value = null;
}
function replyComment(id: string, text: string) {
  if (editor.value?.replyComment(id, text) == null) showNotice('無法回覆這則留言。');
}
function editComment(id: string, text: string) {
  if (!editor.value?.editComment(id, text) && !editor.value?.canChangeComment(id)) showNotice('只能編輯自己的留言。');
}
function removeComment(id: string) {
  if (!editor.value?.deleteComment(id)) showNotice('只能刪除自己的留言。');
}
watch(commentsOpen, (open) => {
  if (!open) commentDraft.value = null;
});

// A short message over the page (e.g. an edit refused in a locked field), gone after a few seconds.
const notice = ref('');
let noticeTimer: ReturnType<typeof setTimeout> | undefined;
function showNotice(message: string) {
  notice.value = message;
  clearTimeout(noticeTimer);
  noticeTimer = setTimeout(() => (notice.value = ''), 3000);
}

// Fields to fill in (content controls of a template; editor/fields.ts). One side panel at a time.
const fieldsOpen = ref(false);
// 修訂窗格 (persona-300): the tracked changes of the part being edited.
const revisionsOpen = ref(false);
watch(revisionsOpen, (open) => {
  if (open) fieldsOpen.value = commentsOpen.value = false;
});

// Navigation pane (left): the headings, and updating a table of contents' page numbers.
const outlineOpen = ref(false);
const headings = computed(() => {
  void snapshot.value;
  return outlineOpen.value ? editor.value?.outline() ?? [] : [];
});
const activeHeading = computed(() => {
  // Re-evaluated on every editor update: a selection change alone leaves `headings` the same
  // (cached) array, so depending on it only would keep the first computed row highlighted.
  void snapshot.value;
  const at = editor.value?.view?.state.selection.from ?? -1;
  let pos: number | null = null;
  for (const h of headings.value) if (h.pos <= at) pos = h.pos;
  return pos;
});
const pageRefs = computed(() => {
  void snapshot.value;
  return outlineOpen.value ? editor.value?.pageReferenceCount() ?? 0 : 0;
});
function updatePages() {
  const n = editor.value?.updatePageReferences() ?? 0;
  showNotice(n ? `已更新 ${n} 個頁碼。` : '頁碼都已是最新的。');
}
/** 參考資料 › 更新目錄: the entries from the headings (editor/toc.ts), then the page numbers. */
function updateToc() {
  const r = editor.value?.updateTableOfContents();
  if (!r) return;
  showNotice(r.rebuilt ? '已依目前的標題更新目錄。' : r.pages ? `已更新 ${r.pages} 個頁碼。` : '目錄已是最新的。');
}
const fields = computed(() => {
  void snapshot.value; // re-evaluated on every editor update (cached per document)
  return editor.value?.fields() ?? [];
});
const activeField = computed(() => {
  void snapshot.value;
  const sel = editor.value?.view?.state.selection;
  return (sel && fields.value.find((f) => sel.from >= f.from && sel.to <= f.to)?.id) ?? null;
});
watch(fieldsOpen, (open) => {
  if (open) commentsOpen.value = revisionsOpen.value = false;
});
watch(commentsOpen, (open) => {
  if (open) fieldsOpen.value = revisionsOpen.value = false;
});

// Word count (the whole body, or the selection) and the browser's spell check.
const counts = computed(() => {
  void snapshot.value;
  const ed = editor.value;
  return ed ? { all: ed.wordCount(), selected: ed.selectionWordCount() } : null;
});
const n = (v: number) => v.toLocaleString('zh-TW');
const countTitle = computed(() => {
  const c = counts.value;
  if (!c) return '';
  const part = (label: string, w: { words: number; chars: number; charsWithSpaces: number }) =>
    `${label}：字數 ${n(w.words)}，字元（不含空白）${n(w.chars)}，字元（含空白）${n(w.charsWithSpaces)}`;
  return [c.selected && part('選取範圍', c.selected), part('整份文件（本文）', c.all)].filter(Boolean).join('\n');
});
const SPELL_KEY = 'papyrus.spellcheck';
const spellcheck = ref(false);
try {
  spellcheck.value = localStorage.getItem(SPELL_KEY) === '1';
} catch {
  // storage unavailable: off
}
function toggleSpellcheck() {
  spellcheck.value = !spellcheck.value;
  editor.value?.setSpellcheck(spellcheck.value);
  try {
    localStorage.setItem(SPELL_KEY, spellcheck.value ? '1' : '0');
  } catch {
    // not remembered; still applies now
  }
}

// Page setup dialog, zoom and page navigation.
const pageSetupOpen = ref(false);
const paragraphOpen = ref(false);
/** 插入 › 符號 › 其他符號…: the 符號 dialog (only while editing). */
const symbolOpen = ref(false);
/** 設計 › 浮水印 › 自訂浮水印…: the 浮水印 dialog (only while editing). */
const watermarkOpen = ref(false);
const ZOOMS = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2];
const zoomChoice = ref<string>('1');
/** The user chose a zoom: never changed for them after that, and remembered on this computer (persona-300 B-10). */
let zoomPicked = false;
const ZOOM_KEY = 'papyrus.zoom';
function onZoom(e: Event) {
  setZoomChoice((e.target as HTMLSelectElement).value);
}
function setZoomChoice(v: string) {
  zoomPicked = true;
  zoomChoice.value = v;
  editor.value?.setZoom(v === 'fit' ? 'fit' : Number(v));
  try {
    localStorage.setItem(ZOOM_KEY, v);
  } catch {
    // not remembered; still applies now
  }
}
/** The zoom chosen last time on this computer, if any. */
function rememberedZoom(): string | null {
  try {
    const v = localStorage.getItem(ZOOM_KEY);
    return v === 'fit' || ZOOMS.map(String).includes(v ?? '') ? v : null;
  } catch {
    return null; // storage unavailable: fitted as usual
  }
}
/**
 * How far the browser itself was zoomed in since the page loaded (Ctrl + +, as people with poor
 * eyesight do): the device pixel ratio now over the one at load; 1 when not known. (A window's
 * outer width over its inner width is not it: DevTools or a sidebar narrow the page as well.)
 */
function browserZoom(): number {
  const now = window.devicePixelRatio;
  return now > 0 ? now / LOADED_PIXEL_RATIO : 1;
}
const zoomLabel = computed(() => `${Math.round((snapshot.value?.zoom ?? 1) * 100)}%`);

// Narrow windows (GOV-FINDING-025; editing stays a desktop feature, the notice stays): until the
// user picks a zoom, a page wider than the space for it is shown smaller so it fits, like Word
// mobile's fit to screen, but never below 50 % (still readable; the rest scrolls). Back to 100 %
// once it fits again. Only the view is scaled: pages and page breaks stay the same.
const MIN_AUTO_ZOOM = 0.5;
/** The page width (canvas px at 100 %) the automatic zoom was last worked out for. */
let fittedWidth = 0;
function autoFit(resized = false) {
  const ed = editor.value;
  const area = mount.value;
  if (zoomPicked || !ed || !area) return;
  const pages = ed.pages ?? [];
  const width = pages.length ? Math.max(...pages.map((p) => p.left + p.width)) : 0;
  if (!width || (!resized && width === fittedWidth)) return; // nothing that matters changed
  const available = area.clientWidth - 32; // less the page area's side padding (as core.ts applyZoom)
  if (available <= 0) return; // not laid out (hidden)
  fittedWidth = width;
  // A window made narrow by zooming the browser in is not made up for by showing the page
  // smaller: that would undo the zoom the user wanted (persona-300 B-10).
  const fits = available >= width || browserZoom() > 1.1;
  const zoom = fits ? 1 : Math.max(MIN_AUTO_ZOOM, available / width);
  zoomChoice.value = fits ? '1' : 'fit';
  if (Math.abs(ed.zoom - zoom) > 0.001) ed.setZoom(zoom);
}
watch(snapshot, () => autoFit());
const onWindowResize = () => autoFit(true);
let areaObserver: ResizeObserver | null = null;
onMounted(() => {
  window.addEventListener('resize', onWindowResize);
  // Also when a side panel opens or closes, not only when the window changes.
  if (typeof ResizeObserver === 'function' && mount.value) {
    areaObserver = new ResizeObserver(onWindowResize);
    areaObserver.observe(mount.value);
  }
});
onBeforeUnmount(() => {
  window.removeEventListener('resize', onWindowResize);
  areaObserver?.disconnect();
});
function goToPage(e: Event) {
  const input = e.target as HTMLInputElement;
  const n = Math.round(Number(input.value));
  const count = snapshot.value?.pageCount ?? 1;
  if (Number.isFinite(n) && n >= 1 && n <= count) editor.value?.scrollToPage(n - 1);
  input.value = '';
}

defineExpose({ save, download, editor, open: load, compat, startComment });
</script>

<template>
  <div class="dx-vue" @keydown="onFindKey">
    <!-- First Tab stop: straight into the text instead of through every toolbar button (GOV-ISSUE-013). -->
    <a v-if="showToolbar" class="dx-skip" href="#" @click.prevent="focusDocument">{{ tl('跳到文件內容') }}</a>
    <DocxToolbar
      v-if="showToolbar"
      ref="toolbarRef"
      :editor="editor"
      :snapshot="snapshot"
      :styles="styles"
      :file-menu="fileMenu"
      :read-only="!editable"
      :outline-open="outlineOpen"
      :comments-open="commentsOpen"
      :comment-count="comments.length"
      :revisions-open="revisionsOpen"
      :spellcheck="spellcheck"
      :zoom="zoomChoice"
      @find="openFind"
      @page-setup="pageSetupOpen = true"
      @paragraph="paragraphOpen = true"
      @symbol="symbolOpen = editable"
      @watermark="watermarkOpen = editable"
      @comment="startComment"
      @file="emit('file')"
      @toggle-outline="outlineOpen = !outlineOpen"
      @toggle-comments="commentsOpen = !commentsOpen"
      @toggle-revisions="revisionsOpen = !revisionsOpen"
      @toggle-spellcheck="toggleSpellcheck"
      @zoom="setZoomChoice"
      @update-toc="updateToc"
      @insert-toc="editor?.insertTableOfContents()"
      @notice="showNotice"
    />
    <PageSetupDialog :editor="editor" :open="pageSetupOpen" @close="pageSetupOpen = false" />
    <ParagraphDialog :editor="editor" :open="paragraphOpen" @close="paragraphOpen = false" />
    <SymbolDialog :editor="editor" :open="symbolOpen" :editable="editable" @close="symbolOpen = false" />
    <WatermarkDialog :editor="editor" :open="watermarkOpen" @close="watermarkOpen = false" />
    <ContextMenu
      :editor="editor"
      :snapshot="snapshot"
      :editable="editable"
      :at="menuAt"
      @close="closeMenu"
      @comment="startComment"
      @paragraph="paragraphOpen = true"
      @notice="showNotice"
    />
    <div class="dx-find-anchor">
      <FindReplace
        v-if="findOpen"
        ref="findPanel"
        v-model:replace="findReplace"
        :editor="editor"
        :snapshot="snapshot"
        :can-replace="editable"
        @close="closeFind"
      />
    </div>
    <CompatNotice :report="compat" />
    <div class="dx-body-row">
      <OutlinePanel
        v-if="outlineOpen"
        :headings="headings"
        :active-pos="activeHeading"
        :page-refs="pageRefs"
        :editable="editable"
        @go="(pos) => editor?.goToHeading(pos)"
        @update-pages="updatePages"
        @close="outlineOpen = false"
      />
      <div class="dx-scroll" @click="onDocClick" @contextmenu="onContextMenu">
        <div ref="mount" />
        <div v-if="loading" class="dx-loading">{{ tl('載入中…') }}</div>
        <div v-if="notice" class="dx-notice" role="status">{{ notice }}</div>
      </div>
      <FieldsPanel
        v-if="fieldsOpen && fields.length"
        :fields="fields"
        :active-id="activeField"
        :editable="editable"
        @select="(id) => editor?.goToField(id)"
        @next="editor?.nextUnfilledField()"
        @close="fieldsOpen = false"
      />
      <RevisionsPanel v-if="revisionsOpen" :editor="editor" :snapshot="snapshot" :editable="editable" @close="revisionsOpen = false" @announce="showNotice" />
      <CommentsPanel
        v-if="commentsOpen && (comments.length || commentDraft)"
        :comments="comments"
        :anchors="commentAnchors"
        :active-id="activeComment"
        :editable="editable || commenting"
        :author="commentAuthor"
        :draft="commentDraft"
        @select="showComment"
        @close="commentsOpen = false"
        @new="startComment"
        @add="addComment"
        @cancel-draft="commentDraft = null"
        @reply="replyComment"
        @resolve="(id: string, done: boolean) => editor?.resolveComment(id, done)"
        @edit="editComment"
        @remove="removeComment"
      />
    </div>
    <div class="dx-status">
      <span aria-live="polite">{{ tl('第') }} {{ snapshot?.currentPage ?? 1 }} {{ tl('頁，共') }} {{ snapshot?.pageCount ?? 1 }} {{ tl('頁') }}</span>
      <span v-if="(snapshot?.sectionCount ?? 1) > 1" class="dx-status-item">{{ tl('第') }} {{ (snapshot?.section ?? 0) + 1 }} {{ tl('節，共') }} {{ snapshot?.sectionCount }} {{ tl('節') }}</span>
      <label class="dx-status-item">
        {{ tl('跳至') }}
        <input class="dx-page-input" type="number" min="1" :max="snapshot?.pageCount ?? 1" :placeholder="String(snapshot?.currentPage ?? 1)" :aria-label="tl('跳至頁碼')" @change="goToPage" @keydown.enter="!composing($event) && goToPage($event)" />
        {{ tl('頁') }}
      </label>
      <span v-if="counts" class="dx-status-item" :title="countTitle">
        {{ counts.selected ? `已選 ${n(counts.selected.words)}／共 ${n(counts.all.words)} 字` : `字數 ${n(counts.all.words)}` }}
      </span>
      <button v-if="editable" type="button" class="dx-status-btn" :aria-pressed="spellcheck"
              :title="tl('使用瀏覽器的拼字檢查（主要檢查英文；多數瀏覽器不檢查中文）')" @click="toggleSpellcheck">
        {{ tl('拼字檢查：') }}{{ spellcheck ? '開' : '關' }}
      </button>
      <label class="dx-status-item dx-zoom">
        {{ tl('縮放') }}
        <select :value="zoomChoice" :aria-label="tl('縮放')" @change="onZoom">
          <option v-for="z in ZOOMS" :key="z" :value="String(z)">{{ Math.round(z * 100) }}%</option>
          <option value="fit">{{ tl('符合寬度（') }}{{ zoomLabel }}{{ tl('）') }}</option>
        </select>
      </label>
      <button type="button" class="dx-status-btn" :aria-pressed="outlineOpen" :title="tl('依標題瀏覽文件（導覽窗格）')" @click="outlineOpen = !outlineOpen">
        {{ tl('導覽') }}
      </button>
      <button v-if="fields.length" type="button" class="dx-status-btn" :aria-pressed="fieldsOpen" @click="fieldsOpen = !fieldsOpen">
        {{ tl('欄位') }} {{ fields.filter((f) => f.filled).length }}/{{ fields.length }}
      </button>
      <button v-if="comments.length" type="button" class="dx-status-btn" :aria-pressed="commentsOpen" @click="commentsOpen = !commentsOpen">
        {{ tl('留言') }} {{ comments.length }}
      </button>
    </div>
  </div>
</template>

<style scoped>
/* Shown only while it has the keyboard focus. */
.dx-skip {
  position: absolute;
  left: -9999px;
  z-index: 20;
  padding: 6px 12px;
  background: #1a73e8;
  color: #fff;
  border-radius: 4px;
  font: 14px system-ui, 'Microsoft JhengHei', sans-serif;
}
.dx-skip:focus {
  left: 8px;
  top: 8px;
}
.dx-vue {
  position: relative;
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
}
/* Zero-height anchor: the find panel floats over the top of the page area. */
.dx-find-anchor {
  position: relative;
  height: 0;
  z-index: 30;
}
.dx-body-row {
  display: flex;
  flex: 1;
  min-height: 0;
}
.dx-status {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  column-gap: 4px;
}
.dx-status-item {
  margin-left: 12px;
  display: inline-flex;
  align-items: center;
  gap: 4px;
}
.dx-zoom {
  margin-left: auto;
}
.dx-page-input {
  width: 4em;
  padding: 0 2px;
  font: inherit;
}
.dx-status select {
  font: inherit;
  padding: 0 2px;
}
.dx-status-btn {
  margin-left: 12px;
  border: 0;
  background: transparent;
  font: inherit;
  color: #1967d2;
  cursor: pointer;
  padding: 0 4px;
  border-radius: 3px;
}
.dx-status-btn:hover {
  background: rgba(0, 0, 0, 0.06);
}
.dx-scroll {
  position: relative;
  flex: 1;
  min-width: 0;
  min-height: 0;
  overflow: auto;
  background: #e8eaed;
}
.dx-notice {
  position: sticky;
  bottom: 16px;
  width: fit-content;
  max-width: 80%;
  margin: 0 auto;
  padding: 6px 14px;
  border-radius: 4px;
  background: #3c4043;
  color: #fff;
  font: 13px system-ui, 'Microsoft JhengHei', sans-serif;
  box-shadow: 0 2px 6px rgba(0, 0, 0, 0.3);
  z-index: 20;
}
.dx-loading {
  position: absolute;
  inset: 0;
  display: grid;
  place-items: center;
  background: rgba(232, 234, 237, 0.7);
  font: 14px system-ui, sans-serif;
  color: #3c4043;
}
.dx-status {
  padding: 2px 12px;
  font: 12px system-ui, 'Microsoft JhengHei', sans-serif;
  color: #5f6368;
  background: #f8f9fa;
  border-top: 1px solid #dadce0;
}
</style>
