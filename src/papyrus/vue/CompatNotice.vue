<script setup lang="ts">
import { ref, watch } from 'vue';
import { KEEP_ORIGINAL, type CompatReport } from '../docx/compat';

// G01: tells the user, when a document opens, what in it the web editor can't fully show or edit.
const props = withDefaults(
  defineProps<{
    report: CompatReport | null;
    /** How to get the original file back; the page may word it for its own version history. */
    keepOriginal?: string;
  }>(),
  { keepOriginal: KEEP_ORIGINAL },
);

const dismissed = ref(false);
const open = ref(false);
watch(
  () => props.report,
  () => {
    dismissed.value = false;
    open.value = false;
  },
);
</script>

<template>
  <div v-if="report && report.items.length && !dismissed" class="dx-compat" role="status">
    <div class="dx-compat-bar">
      <span class="dx-compat-icon" aria-hidden="true">!</span>
      <span class="dx-compat-text">這份文件有 {{ report.items.length }} 項內容在網頁上無法完整顯示或編輯</span>
      <button type="button" class="dx-compat-link" :aria-expanded="open" @click="open = !open">{{ open ? '收合' : '查看詳情' }}</button>
      <button type="button" class="dx-compat-close" title="關閉提示" aria-label="關閉提示" @click="dismissed = true">×</button>
    </div>
    <div v-if="open" class="dx-compat-panel">
      <p class="dx-compat-lead">網頁版只支援 Word 的部分功能。下列內容都會原樣保存在檔案中，但在網頁上的顯示或編輯方式有限制：</p>
      <ul>
        <li v-for="item in report.items" :key="item.id">
          <div class="dx-compat-title">
            {{ item.title }}<span v-if="item.where" class="dx-compat-where">{{ item.where }}</span>
          </div>
          <div class="dx-compat-effect">{{ item.effect }}</div>
        </li>
      </ul>
      <p class="dx-compat-keep"><b>保留原檔：</b>{{ keepOriginal }}</p>
    </div>
  </div>
</template>

<style scoped>
.dx-compat {
  font: 13px system-ui, 'Microsoft JhengHei', sans-serif;
  color: #5f4100;
  background: #fef7e0;
  border-bottom: 1px solid #f6d57a;
}
.dx-compat-bar {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 12px;
  min-height: 28px;
}
.dx-compat-icon {
  display: inline-grid;
  place-items: center;
  width: 16px;
  height: 16px;
  border-radius: 50%;
  background: #f9ab00;
  color: #fff;
  font-weight: 700;
  font-size: 11px;
  flex: none;
}
.dx-compat-text {
  flex: 1;
  min-width: 0;
}
.dx-compat-link,
.dx-compat-close {
  border: 0;
  background: transparent;
  color: #1967d2;
  font: inherit;
  cursor: pointer;
  padding: 2px 4px;
  border-radius: 4px;
}
.dx-compat-close {
  color: #5f6368;
  font-size: 16px;
  line-height: 1;
}
.dx-compat-link:hover,
.dx-compat-close:hover {
  background: rgba(0, 0, 0, 0.06);
}
.dx-compat-link:focus-visible,
.dx-compat-close:focus-visible {
  outline: 2px solid #1a73e8;
}
.dx-compat-panel {
  max-height: 40vh;
  overflow: auto;
  padding: 4px 16px 10px 36px;
  color: #3c4043;
}
.dx-compat-lead,
.dx-compat-keep {
  margin: 4px 0 8px;
}
ul {
  margin: 0;
  padding: 0;
  list-style: none;
}
li {
  padding: 6px 0;
  border-top: 1px solid rgba(0, 0, 0, 0.06);
}
.dx-compat-title {
  font-weight: 600;
}
.dx-compat-where {
  font-weight: 400;
  color: #5f6368;
  margin-left: 8px;
}
.dx-compat-effect {
  margin-top: 2px;
}
.dx-compat-keep {
  margin-top: 8px;
  padding-top: 8px;
  border-top: 1px solid rgba(0, 0, 0, 0.1);
}
</style>
