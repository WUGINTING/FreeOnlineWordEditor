<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue';
import { commentThreads, type CommentThread, type DocComment } from '../docx/comments';
import { formatDate } from '../docx/revisions';
import { composing } from './keys';

// G09: the document's comments. Clicking one scrolls to what it comments on. While editing is
// allowed: add a comment on the selection, reply to a thread, resolve / reopen it, and edit or
// delete your own comments (every change can be undone like any other edit).
const props = withDefaults(
  defineProps<{
    comments: DocComment[];
    /** Commented text by comment id. */
    anchors: Record<string, string>;
    /** Comment whose range holds the cursor. */
    activeId: string | null;
    /** Comments can be added and changed (false: read-only). */
    editable?: boolean;
    /** Name of the signed-in user: they may edit and delete comments with this author. */
    author?: string | null;
    /** A new comment being written: the text it will be on. */
    draft?: { quote: string } | null;
  }>(),
  { editable: false, author: null, draft: null },
);
const emit = defineEmits<{
  (e: 'select', id: string): void;
  (e: 'close'): void;
  /** Start a new comment on the selection (the host works out where). */
  (e: 'new'): void;
  (e: 'add', text: string): void;
  (e: 'cancel-draft'): void;
  (e: 'reply', id: string, text: string): void;
  (e: 'resolve', id: string, done: boolean): void;
  (e: 'edit', id: string, text: string): void;
  (e: 'remove', id: string): void;
}>();

// Replies to replies stay in their thread; a reply loop or a missing parent never hides a comment.
const threads = computed<CommentThread[]>(() => commentThreads(props.comments));
const mine = (c: DocComment) => props.editable && !!props.author && c.author === props.author;

function snippet(text: string): string {
  const t = text.replace(/\s+/g, ' ').trim();
  return t.length > 60 ? t.slice(0, 60) + '…' : t;
}

// The new comment's box.
const draftText = ref('');
const draftBox = ref<HTMLTextAreaElement | null>(null);
watch(
  () => props.draft,
  async (d, old) => {
    if (!d) return;
    if (!old) draftText.value = '';
    await nextTick();
    draftBox.value?.focus();
  },
  { immediate: true },
);
function addDraft() {
  const text = draftText.value.trim();
  if (!text) return;
  awaitingNew = new Set(props.comments.map((c) => c.id));
  emit('add', text);
  draftText.value = '';
  // GOV-093: the box goes away; the keyboard stays in the panel (on 新增留言), then moves to the new
  // comment when it is listed (the host lists it on its next update).
  focusIn('.dx-comments-new');
}
let awaitingNew: Set<string> | null = null;
watch(
  () => props.comments,
  (list) => {
    if (!awaitingNew) return;
    const added = list.find((c) => !awaitingNew!.has(c.id));
    if (!added) return;
    awaitingNew = null;
    focusIn('.dx-comment-main', added.id);
  },
);

// One reply box and one edit box open at a time.
const replyTo = ref<string | null>(null);
const replyText = ref('');
const editing = ref<string | null>(null);
const editText = ref('');
const panel = ref<HTMLElement | null>(null);
const focusBox = async (cls: string) => {
  await nextTick();
  (panel.value?.querySelector(`.${cls}`) as HTMLTextAreaElement | null)?.focus();
};
function startReply(t: CommentThread) {
  editing.value = null;
  replyTo.value = t.comment.id;
  replyText.value = '';
  void focusBox('dx-comment-reply-box');
}
/**
 * Once Vue has updated the panel: focuses the first `selector` element, the one whose data-id is
 * `id` when given, or the `inner` element inside that one.
 */
function focusIn(selector: string, id?: string, inner?: string) {
  void nextTick(() => {
    const all = Array.from(panel.value?.querySelectorAll<HTMLElement>(selector) ?? []);
    const el = id == null ? all[0] : all.find((e) => e.dataset.id === id);
    ((inner ? el?.querySelector(inner) : el) as HTMLElement | null | undefined)?.focus();
  });
}
// GOV-093: when a reply or edit box closes (sent or cancelled), the keyboard goes back to the
// button that opened it instead of falling to the page.
function closeReply() {
  const id = replyTo.value;
  replyTo.value = null;
  if (id != null) focusIn('.dx-comment-thread', id, '.dx-comment-reply-open');
}
function sendReply() {
  const text = replyText.value.trim();
  if (!text || replyTo.value == null) return;
  emit('reply', replyTo.value, text);
  replyText.value = '';
  closeReply();
}
function startEdit(c: DocComment) {
  replyTo.value = null;
  editing.value = c.id;
  editText.value = c.text;
  void focusBox('dx-comment-edit-box');
}
function closeEdit() {
  const id = editing.value;
  editing.value = null;
  if (id != null) focusIn('.dx-comment-own', id, '.dx-comment-edit-open');
}
function saveEdit() {
  const text = editText.value.trim();
  if (!text || editing.value == null) return;
  emit('edit', editing.value, text);
  closeEdit();
}
/**
 * Ctrl+Enter sends, Esc cancels (as in Word's comment box). Not while an input method (注音 …)
 * is composing: its Esc only drops the candidates, never the comment being written (A-7).
 */
function keys(e: KeyboardEvent, send: () => void, cancel: () => void) {
  if (composing(e)) return;
  if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
    e.preventDefault();
    send();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    e.stopPropagation();
    cancel();
  }
}
const count = computed(() => props.comments.length);

// GOV-051 / GOV-060: 只看未解決, and 上一則 / 下一則 through the threads listed (in the panel's order).
const openOnly = ref(false);
const openCount = computed(() => threads.value.filter((t) => !t.comment.done).length);
const shown = computed(() => (openOnly.value ? threads.value.filter((t) => !t.comment.done) : threads.value));
/** The listed thread holding the cursor's comment (-1: none). */
const current = computed(() =>
  shown.value.findIndex((t) => t.comment.id === props.activeId || t.replies.some((r) => r.id === props.activeId)),
);
function resolve(t: CommentThread) {
  emit('resolve', t.comment.id, !t.comment.done);
  // Resolved while only open ones are listed: its card goes away; the keyboard stays in the panel.
  if (openOnly.value && !t.comment.done) focusIn('.dx-comments-filter');
}
/** Selects the comment `step` threads from the current one (from the first or the last when none is). */
function go(step: 1 | -1) {
  const list = shown.value;
  if (!list.length) return;
  const at = current.value < 0 ? (step > 0 ? 0 : list.length - 1) : current.value + step;
  const t = list[at];
  if (!t) return;
  emit('select', t.comment.id);
  // Selecting moves the focus into the document: back to this button, so the next press goes on;
  // to the other one at the first / last comment (this one is off there).
  const last = step > 0 ? at === list.length - 1 : at === 0;
  focusIn((step > 0) !== last ? '.dx-comments-next' : '.dx-comments-prev');
}
</script>

<template>
  <aside ref="panel" class="dx-comments" aria-label="留言">
    <header>
      <span class="dx-comments-title">留言（{{ count }}）</span>
      <button v-if="editable" type="button" class="dx-comments-new" title="在選取的文字上新增留言 (Ctrl+Alt+M)" @click="emit('new')">＋ 新增留言</button>
      <button type="button" class="dx-comments-close" title="關閉留言面板" aria-label="關閉留言面板" @click="emit('close')">×</button>
    </header>
    <div v-if="threads.length" class="dx-comments-nav">
      <button type="button" class="dx-comments-prev" title="上一則留言" aria-label="上一則留言"
              :disabled="!shown.length || current === 0" @click="go(-1)">↑ 上一則</button>
      <button type="button" class="dx-comments-next" title="下一則留言" aria-label="下一則留言"
              :disabled="!shown.length || current === shown.length - 1" @click="go(1)">↓ 下一則</button>
      <button type="button" class="dx-comments-filter" :aria-pressed="openOnly"
              title="只列出還沒解決的留言" @click="openOnly = !openOnly">只看未解決（{{ openCount }}）</button>
    </div>
    <p v-if="!editable" class="dx-comments-note">唯讀：只能檢視留言。</p>
    <p v-if="openOnly && !shown.length" class="dx-comments-note">沒有未解決的留言。</p>
    <form v-if="editable && draft" class="dx-comment-draft" @submit.prevent="addDraft">
      <blockquote>{{ draft.quote ? snippet(draft.quote) : '（游標位置）' }}</blockquote>
      <textarea
        ref="draftBox"
        v-model="draftText"
        class="dx-comment-box"
        rows="3"
        aria-label="新留言"
        placeholder="輸入留言…（Ctrl+Enter 送出）"
        @keydown="keys($event, addDraft, () => emit('cancel-draft'))"
      />
      <div class="dx-comment-actions">
        <button type="submit" class="dx-comment-primary" :disabled="!draftText.trim()">留言</button>
        <button type="button" @click="emit('cancel-draft')">取消</button>
      </div>
    </form>
    <ol>
      <li
        v-for="t in shown"
        :key="t.key"
        :class="{ active: activeId === t.comment.id || t.replies.some((r) => r.id === activeId), done: t.comment.done }"
      >
        <div class="dx-comment-card">
          <template v-for="(c, i) in [t.comment, ...t.replies]" :key="i">
            <div :class="i ? 'dx-comment-reply' : 'dx-comment-root'">
              <button type="button" class="dx-comment-main" :data-id="c.id" @click="emit('select', c.id)">
                <div class="dx-comment-head">
                  <span class="dx-comment-author">{{ c.author || '未知作者' }}</span>
                  <span class="dx-comment-date">{{ formatDate(c.date, c.dateUtc) }}</span>
                  <span v-if="!i && t.comment.done" class="dx-comment-done">已解決</span>
                </div>
                <blockquote v-if="!i && anchors[c.id]">{{ snippet(anchors[c.id]) }}</blockquote>
                <div v-if="editing !== c.id" class="dx-comment-text">{{ c.text || (i ? '' : '（空白留言）') }}</div>
              </button>
              <form v-if="editing === c.id" class="dx-comment-edit" @submit.prevent="saveEdit">
                <textarea
                  v-model="editText"
                  class="dx-comment-box dx-comment-edit-box"
                  rows="3"
                  aria-label="編輯留言"
                  @keydown="keys($event, saveEdit, closeEdit)"
                />
                <div class="dx-comment-actions">
                  <button type="submit" class="dx-comment-primary" :disabled="!editText.trim()">儲存</button>
                  <button type="button" @click="closeEdit">取消</button>
                </div>
              </form>
              <div v-else-if="mine(c)" class="dx-comment-actions dx-comment-own" :data-id="c.id">
                <button type="button" class="dx-comment-edit-open" :aria-label="`編輯${c.author}的留言`" @click="startEdit(c)">編輯</button>
                <button
                  type="button"
                  :title="!i && t.replies.length ? `連同 ${t.replies.length} 則回覆一起刪除（可復原）` : '刪除（可復原）'"
                  :aria-label="`刪除${c.author}的留言`"
                  @click="emit('remove', c.id)"
                >刪除</button>
              </div>
            </div>
          </template>
          <form v-if="editable && replyTo === t.comment.id" class="dx-comment-edit" @submit.prevent="sendReply">
            <textarea
              v-model="replyText"
              class="dx-comment-box dx-comment-reply-box"
              rows="2"
              aria-label="回覆"
              placeholder="回覆…（Ctrl+Enter 送出）"
              @keydown="keys($event, sendReply, closeReply)"
            />
            <div class="dx-comment-actions">
              <button type="submit" class="dx-comment-primary" :disabled="!replyText.trim()">回覆</button>
              <button type="button" @click="closeReply">取消</button>
            </div>
          </form>
          <div v-else-if="editable" class="dx-comment-actions dx-comment-thread" :data-id="t.comment.id">
            <button type="button" class="dx-comment-reply-open" @click="startReply(t)">回覆</button>
            <button type="button" :aria-pressed="t.comment.done" @click="resolve(t)">
              {{ t.comment.done ? '重新開啟' : '標示為已解決' }}
            </button>
          </div>
        </div>
      </li>
    </ol>
  </aside>
</template>

<style scoped>
.dx-comments {
  width: 18rem;
  flex: none;
  display: flex;
  flex-direction: column;
  min-height: 0;
  background: #f8f9fa;
  border-left: 1px solid #dadce0;
  /* In rem (the browser's font size setting); buttons at least 24 px high with 14 px text (persona-300 B-10). */
  font: 0.875rem system-ui, 'Microsoft JhengHei', sans-serif;
  color: #3c4043;
}
header {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 8px 12px 4px;
}
.dx-comments-title {
  flex: 1;
  font-weight: 600;
}
.dx-comments-close {
  min-width: 1.75rem;
  min-height: 1.75rem;
  border: 0;
  background: transparent;
  font-size: 1.125rem;
  line-height: 1;
  cursor: pointer;
  color: #5f6368;
  padding: 2px 6px;
  border-radius: 4px;
}
.dx-comments-close:hover,
.dx-comments-new:hover,
.dx-comment-actions button:hover:not(:disabled) {
  background: rgba(0, 0, 0, 0.06);
}
.dx-comments-new {
  min-height: 1.75rem;
  border: 0;
  background: transparent;
  font: inherit;
  color: #1967d2;
  cursor: pointer;
  padding: 2px 6px;
  border-radius: 4px;
}
.dx-comments-nav {
  display: flex;
  gap: 4px;
  padding: 0 12px 6px;
}
.dx-comments-nav {
  flex-wrap: wrap;
}
.dx-comments-nav button {
  min-height: 1.75rem;
  border: 1px solid #dadce0;
  background: #fff;
  font: inherit;
  color: #3c4043;
  cursor: pointer;
  padding: 1px 8px;
  border-radius: 4px;
}
.dx-comments-nav button:disabled {
  color: #9aa0a6;
  cursor: default;
}
.dx-comments-nav button:hover:not(:disabled) {
  background: rgba(0, 0, 0, 0.06);
}
.dx-comments-filter {
  margin-left: auto;
}
.dx-comments-filter[aria-pressed='true'] {
  background: #e8f0fe !important;
  border-color: #1967d2 !important;
  color: #1967d2 !important;
}
.dx-comments-note {
  margin: 0 12px 8px;
  font-size: 0.8125rem;
  color: #5f6368;
}
ol {
  list-style: none;
  margin: 0;
  padding: 0 8px 12px;
  overflow: auto;
  flex: 1;
  min-height: 0;
}
li {
  margin-bottom: 8px;
}
.dx-comment-card,
.dx-comment-draft {
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  padding: 8px 10px;
}
.dx-comment-draft {
  margin: 0 8px 8px;
  border-color: #f0b400;
  box-shadow: 0 0 0 2px rgba(240, 180, 0, 0.35);
}
.dx-comment-card:hover {
  border-color: #f0b400;
}
.dx-comment-main {
  display: block;
  width: 100%;
  text-align: left;
  font: inherit;
  color: inherit;
  background: transparent;
  border: 0;
  padding: 0;
  cursor: pointer;
}
.dx-comment-main:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: 2px;
}
li.active .dx-comment-card {
  border-color: #f0b400;
  box-shadow: 0 0 0 2px rgba(240, 180, 0, 0.35);
}
li.done .dx-comment-card {
  opacity: 0.7;
}
.dx-comment-head {
  display: flex;
  flex-wrap: wrap;
  align-items: baseline;
  gap: 6px;
}
.dx-comment-author {
  font-weight: 600;
}
.dx-comment-date {
  font-size: 0.8125rem;
  color: #80868b;
}
.dx-comment-done {
  font-size: 11px;
  color: #137333;
  background: #e6f4ea;
  border-radius: 3px;
  padding: 0 4px;
}
blockquote {
  margin: 6px 0;
  padding: 2px 8px;
  border-left: 3px solid #f0b400;
  color: #5f6368;
}
.dx-comment-text {
  margin-top: 4px;
  white-space: pre-wrap;
  word-break: break-word;
}
.dx-comment-reply {
  margin-top: 8px;
  padding-top: 6px;
  border-top: 1px solid #f1f3f4;
}
.dx-comment-box {
  display: block;
  box-sizing: border-box;
  width: 100%;
  margin-top: 4px;
  padding: 4px 6px;
  font: inherit;
  color: inherit;
  border: 1px solid #dadce0;
  border-radius: 4px;
  resize: vertical;
}
.dx-comment-box:focus {
  outline: 2px solid #1a73e8;
  outline-offset: -1px;
}
.dx-comment-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
  margin-top: 4px;
}
.dx-comment-actions button {
  min-height: 1.75rem;
  border: 0;
  background: transparent;
  font: inherit;
  color: #1967d2;
  cursor: pointer;
  padding: 2px 8px;
  border-radius: 3px;
}
.dx-comment-actions button:disabled {
  color: #9aa0a6;
  cursor: default;
}
.dx-comment-actions .dx-comment-primary {
  background: #1a73e8;
  color: #fff;
}
.dx-comment-actions .dx-comment-primary:hover:not(:disabled) {
  background: #1967d2;
}
.dx-comment-actions .dx-comment-primary:disabled {
  background: #dadce0;
  color: #fff;
}
.dx-comment-own {
  justify-content: flex-end;
}
.dx-comment-thread {
  margin-top: 6px;
  padding-top: 4px;
  border-top: 1px solid #f1f3f4;
}
</style>
