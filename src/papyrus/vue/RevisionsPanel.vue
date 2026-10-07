<script setup lang="ts">
import { tl } from './locale';
// 修訂窗格 (persona-300; Word's 審閱窗格): every tracked change of the part being edited (the
// body, or the header / footer open), in order, with who made it, what it is (插入、刪除、格式變更
// …), its text and when (民國). Clicking one goes there; 接受 / 拒絕 act on that one only (one
// undo step each). The author is written out, not only a colour.
import { computed, onBeforeUnmount, shallowRef, watch } from 'vue';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import type { DocxEditor, EditorSnapshot } from '../editor/core';
import {
  collectRevisions, goToRevisionKey, resolveRevisionKey, revisionExcerpt, revisionLabel, reviewSummary, type Revision,
} from '../editor/review';
import { authorColor, rocDate } from '../docx/revisions';

const props = withDefaults(defineProps<{ editor: DocxEditor | null; snapshot: EditorSnapshot | null; editable?: boolean }>(), { editable: true });
const emit = defineEmits<{ (e: 'close'): void; (e: 'announce', message: string): void }>();

interface Row {
  key: string;
  label: string;
  author: string;
  when: string;
  excerpt: string;
  color: string;
}

const view = () => props.editor?.activeView ?? null;

// The rows follow the document, not every keystroke (persona-300 review): typing waits until
// it pauses (REBUILD_DELAY); another part, the panel's own 接受 / 拒絕 and a pointer or the
// focus coming to the panel bring them up to date at once. A row that did not change is the
// same object, so only changed rows are drawn again.
const REBUILD_DELAY = 250;
const rows = shallowRef<Row[]>([]);
let shown: { view: EditorView | null; doc: PMNode | null } = { view: null, doc: null };
let timer: ReturnType<typeof setTimeout> | undefined;
function rebuild() {
  clearTimeout(timer);
  timer = undefined;
  const v = view();
  const doc = v?.state.doc ?? null;
  if (v === shown.view && doc === shown.doc) return;
  shown = { view: v, doc };
  if (!doc) {
    rows.value = [];
    return;
  }
  const old = new Map(rows.value.map((r) => [r.key, r]));
  const next = collectRevisions(doc).map((r: Revision): Row => {
    const row = {
      key: r.key,
      label: revisionLabel(r),
      author: r.author || '未知作者',
      when: rocDate(r.date, r.dateUtc),
      excerpt: revisionExcerpt(doc, r),
      color: authorColor(r.author),
    };
    const was = old.get(r.key);
    return was && (Object.keys(row) as (keyof Row)[]).every((k) => was[k] === row[k]) ? was : row;
  });
  if (next.length !== rows.value.length || next.some((r, i) => r !== rows.value[i])) rows.value = next;
}
watch(
  () => props.snapshot,
  () => {
    const v = view();
    if (v === shown.view && v?.state.doc === shown.doc) return;
    // Another part (or the first time): at once. The same part edited: once typing pauses.
    if (v !== shown.view || !shown.doc) rebuild();
    else if (timer === undefined) timer = setTimeout(rebuild, REBUILD_DELAY);
  },
  { immediate: true },
);
onBeforeUnmount(() => clearTimeout(timer));
const current = computed(() => {
  void props.snapshot;
  return reviewSummary(view()?.state).current?.key ?? null;
});
const where = computed(() => (props.snapshot?.target === 'header' ? '頁首' : props.snapshot?.target === 'footer' ? '頁尾' : '正文'));

function go(row: Row) {
  props.editor?.run(goToRevisionKey(row.key));
}
function resolveOne(row: Row, accept: boolean) {
  if (props.editor?.run(resolveRevisionKey(row.key, accept))) {
    rebuild();
    emit('announce', `已${accept ? '接受' : '拒絕'}：${row.label}，${row.author}`);
  }
}
</script>

<template>
  <aside class="dx-revs" :aria-label="tl('修訂窗格')" @pointerdown.capture="rebuild" @focusin="rebuild">
    <header>
      <span class="dx-revs-title">{{ tl('修訂（') }}{{ where }} {{ rows.length }} {{ tl('處）') }}</span>
      <button type="button" class="dx-revs-close" :title="tl('關閉修訂窗格')" :aria-label="tl('關閉修訂窗格')" @click="emit('close')">×</button>
    </header>
    <p v-if="!rows.length" class="dx-revs-note">{{ tl('這部分沒有修訂。') }}</p>
    <ol>
      <li v-for="(r, i) in rows" :key="r.key" :class="{ active: current === r.key }">
        <button type="button" class="dx-revs-main" :aria-current="current === r.key ? 'true' : undefined" @click="go(r)">
          <span class="dx-revs-head">
            <span class="dx-revs-kind" :style="{ borderColor: r.color, color: r.color }">{{ r.label }}</span>
            <span class="dx-revs-author">{{ r.author }}</span>
            <span class="dx-revs-when">{{ r.when }}</span>
          </span>
          <span v-if="r.excerpt" class="dx-revs-text" :class="{ 'dx-revs-deleted': r.label === '刪除' || r.label === '移出' }">{{ r.excerpt }}</span>
          <span class="dx-sr">{{ tl('第') }} {{ i + 1 }} {{ tl('處，共') }} {{ rows.length }} {{ tl('處') }}</span>
        </button>
        <div v-if="editable" class="dx-revs-actions">
          <button type="button" :aria-label="`接受：${r.label}，${r.author}`" @click="resolveOne(r, true)">{{ tl('接受') }}</button>
          <button type="button" :aria-label="`拒絕：${r.label}，${r.author}`" @click="resolveOne(r, false)">{{ tl('拒絕') }}</button>
        </div>
      </li>
    </ol>
  </aside>
</template>

<style scoped>
.dx-revs {
  width: 18rem;
  flex: none;
  display: flex;
  flex-direction: column;
  min-height: 0;
  background: #f8f9fa;
  border-left: 1px solid #dadce0;
  font: 0.875rem system-ui, 'Microsoft JhengHei', sans-serif;
  color: #3c4043;
}
header {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 8px 12px 4px;
}
.dx-revs-title {
  flex: 1;
  font-weight: 600;
}
.dx-revs-close {
  min-width: 1.75rem;
  min-height: 1.75rem;
  border: 0;
  background: transparent;
  font-size: 1.125rem;
  cursor: pointer;
  color: #5f6368;
  border-radius: 4px;
}
.dx-revs-note {
  margin: 0 12px 8px;
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
  margin-bottom: 6px;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  padding: 6px 8px;
}
li.active {
  border-color: #1a73e8;
  box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.25);
}
.dx-revs-main {
  display: flex;
  flex-direction: column;
  gap: 2px;
  width: 100%;
  text-align: left;
  font: inherit;
  color: inherit;
  background: transparent;
  border: 0;
  padding: 0;
  cursor: pointer;
}
.dx-revs-main:focus-visible,
.dx-revs-actions button:focus-visible,
.dx-revs-close:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: 2px;
}
.dx-revs-head {
  display: flex;
  flex-wrap: wrap;
  align-items: baseline;
  gap: 6px;
}
.dx-revs-kind {
  padding: 0 4px;
  border: 1px solid;
  border-radius: 3px;
  font-size: 0.8125rem;
}
.dx-revs-author {
  font-weight: 600;
}
.dx-revs-when {
  font-size: 0.8125rem;
  color: #5f6368;
}
.dx-revs-text {
  overflow-wrap: anywhere;
}
.dx-revs-deleted {
  text-decoration: line-through;
}
.dx-revs-actions {
  display: flex;
  justify-content: flex-end;
  gap: 4px;
  margin-top: 4px;
}
.dx-revs-actions button {
  min-height: 1.75rem;
  padding: 2px 10px;
  border: 1px solid #dadce0;
  border-radius: 4px;
  background: #fff;
  font: inherit;
  color: #1967d2;
  cursor: pointer;
}
.dx-sr {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
}
</style>
