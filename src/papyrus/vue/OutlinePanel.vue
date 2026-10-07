<script setup lang="ts">
import type { Heading } from '../editor/outline';

// Navigation pane (Word's 導覽窗格): the document's headings; a click goes to one. Below them,
// updating the page numbers a table of contents (PAGEREF fields) shows.
defineProps<{
  headings: Heading[];
  /** Position of the heading the cursor is under. */
  activePos: number | null;
  /** How many table-of-contents / 「見第 N 頁」 page numbers the document shows. */
  pageRefs: number;
  editable: boolean;
}>();
const emit = defineEmits<{
  (e: 'go', pos: number): void;
  (e: 'update-pages'): void;
  (e: 'close'): void;
}>();
</script>

<template>
  <aside class="dx-outline" aria-label="導覽">
    <header>
      <span class="dx-outline-title">導覽</span>
      <button type="button" class="dx-outline-close" title="關閉導覽窗格" aria-label="關閉導覽窗格" @click="emit('close')">×</button>
    </header>
    <p v-if="!headings.length" class="dx-outline-empty">這份文件沒有標題。套用「標題 1」「標題 2」等樣式的段落會列在這裡。</p>
    <ol v-else>
      <li v-for="h in headings" :key="h.pos" :style="{ paddingLeft: `${h.level * 14}px` }" :class="{ active: activePos === h.pos }">
        <button type="button" :title="h.text" @click="emit('go', h.pos)">{{ h.text }}</button>
      </li>
    </ol>
    <footer v-if="editable && pageRefs">
      <button type="button" class="dx-outline-update" title="依網頁上的分頁重算目錄與「見第 N 頁」的頁碼（可復原）" @click="emit('update-pages')">
        更新目錄頁碼
      </button>
      <p>頁碼依網頁上的分頁計算；在 Word 開啟後也可按 F9 更新。</p>
    </footer>
  </aside>
</template>

<style scoped>
.dx-outline {
  width: 240px;
  flex: none;
  display: flex;
  flex-direction: column;
  min-height: 0;
  background: #f8f9fa;
  border-right: 1px solid #dadce0;
  font: 13px system-ui, 'Microsoft JhengHei', sans-serif;
  color: #3c4043;
}
header {
  display: flex;
  align-items: center;
  padding: 8px 12px 4px;
}
.dx-outline-title {
  flex: 1;
  font-weight: 600;
}
.dx-outline-close {
  border: 0;
  background: transparent;
  font-size: 16px;
  line-height: 1;
  cursor: pointer;
  color: #5f6368;
  padding: 2px 6px;
  border-radius: 4px;
}
.dx-outline-close:hover {
  background: rgba(0, 0, 0, 0.06);
}
.dx-outline-empty {
  margin: 4px 12px;
  font-size: 12px;
  color: #5f6368;
}
ol {
  list-style: none;
  margin: 0;
  padding: 0 4px 8px;
  overflow: auto;
  flex: 1;
  min-height: 0;
}
li button {
  display: block;
  width: 100%;
  text-align: left;
  font: inherit;
  color: inherit;
  background: transparent;
  border: 0;
  border-radius: 4px;
  padding: 3px 8px;
  cursor: pointer;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
li button:hover {
  background: rgba(0, 0, 0, 0.06);
}
li.active button {
  background: #e8f0fe;
  color: #1967d2;
  font-weight: 600;
}
footer {
  border-top: 1px solid #dadce0;
  padding: 8px 12px;
}
footer p {
  margin: 4px 0 0;
  font-size: 11px;
  color: #80868b;
}
.dx-outline-update {
  font: inherit;
  padding: 3px 8px;
  border: 1px solid #dadce0;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}
.dx-outline-update:hover {
  background: #f1f3f4;
}
</style>
