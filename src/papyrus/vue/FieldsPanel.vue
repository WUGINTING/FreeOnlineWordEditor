<script setup lang="ts">
import { computed } from 'vue';
import type { DocField } from '../editor/fields';

// Fields to fill in (a template's content controls): which are done, which are still empty.
// Clicking one selects it in the document.
const props = defineProps<{
  fields: DocField[];
  /** Field whose content holds the cursor. */
  activeId: string | null;
  editable: boolean;
}>();
const emit = defineEmits<{
  (e: 'select', id: string): void;
  (e: 'next'): void;
  (e: 'close'): void;
}>();

const KIND: Record<DocField['kind'], string> = {
  text: '文字', richText: '文字', date: '日期', list: '選項', checkbox: '核取方塊', picture: '圖片',
};
const done = computed(() => props.fields.filter((f) => f.filled).length);
const missing = computed(() => props.fields.filter((f) => f.required && !f.filled).length);

function snippet(f: DocField): string {
  if (f.placeholder) return '（尚未填寫）';
  const t = f.text.replace(/￼/g, '').replace(/\s+/g, ' ').trim();
  if (!t) return '（空白）';
  return t.length > 40 ? t.slice(0, 40) + '…' : t;
}
</script>

<template>
  <aside class="dx-fields" aria-label="填寫欄位">
    <header>
      <span class="dx-fields-title">欄位（已填 {{ done }} / {{ fields.length }}）</span>
      <button type="button" class="dx-fields-close" title="關閉欄位面板" aria-label="關閉欄位面板" @click="emit('close')">×</button>
    </header>
    <p v-if="missing" class="dx-fields-missing" role="status">還有 {{ missing }} 個必填欄位沒有填寫</p>
    <p v-else class="dx-fields-ok" role="status">必填欄位都已填寫</p>
    <button v-if="editable && done < fields.length" type="button" class="dx-fields-next" @click="emit('next')">下一個未填欄位</button>
    <ol>
      <li v-for="f in fields" :key="f.id" :class="{ active: activeId === f.id, filled: f.filled, required: f.required && !f.filled }">
        <button type="button" class="dx-field-card" @click="emit('select', f.id)">
          <span class="dx-field-state" :aria-label="f.filled ? '已填寫' : '未填寫'">{{ f.filled ? '✓' : f.required ? '!' : '○' }}</span>
          <span class="dx-field-body">
            <span class="dx-field-name">{{ f.title }}<span v-if="f.required" class="dx-field-req">必填</span><span v-if="f.locked" class="dx-field-lock">鎖定</span></span>
            <span class="dx-field-meta">{{ KIND[f.kind] }}・{{ snippet(f) }}</span>
          </span>
        </button>
      </li>
    </ol>
  </aside>
</template>

<style scoped>
.dx-fields {
  width: 260px;
  flex: none;
  display: flex;
  flex-direction: column;
  min-height: 0;
  background: #f8f9fa;
  border-left: 1px solid #dadce0;
  font: 13px system-ui, 'Microsoft JhengHei', sans-serif;
  color: #3c4043;
}
header {
  display: flex;
  align-items: center;
  padding: 8px 12px 4px;
}
.dx-fields-title {
  flex: 1;
  font-weight: 600;
}
.dx-fields-close {
  border: 0;
  background: transparent;
  font-size: 16px;
  line-height: 1;
  cursor: pointer;
  color: #5f6368;
  padding: 2px 6px;
  border-radius: 4px;
}
.dx-fields-close:hover {
  background: rgba(0, 0, 0, 0.06);
}
.dx-fields-missing,
.dx-fields-ok {
  margin: 0 12px 6px;
  font-size: 12px;
}
.dx-fields-missing {
  color: #b3261e;
}
.dx-fields-ok {
  color: #137333;
}
.dx-fields-next {
  margin: 0 12px 8px;
  font: inherit;
  padding: 3px 8px;
  border: 1px solid #dadce0;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}
.dx-fields-next:hover {
  background: #f1f3f4;
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
}
.dx-field-card {
  display: flex;
  gap: 8px;
  width: 100%;
  text-align: left;
  font: inherit;
  color: inherit;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  padding: 6px 8px;
  cursor: pointer;
}
.dx-field-card:hover,
li.active .dx-field-card {
  border-color: #1a73e8;
}
li.active .dx-field-card {
  box-shadow: 0 0 0 2px rgba(26, 115, 232, 0.25);
}
.dx-field-card:focus-visible {
  outline: 2px solid #1a73e8;
}
.dx-field-state {
  flex: none;
  width: 1.2em;
  text-align: center;
  font-weight: 700;
  color: #80868b;
}
li.filled .dx-field-state {
  color: #137333;
}
li.required .dx-field-state {
  color: #b3261e;
}
.dx-field-body {
  display: flex;
  flex-direction: column;
  min-width: 0;
}
.dx-field-name {
  font-weight: 600;
  overflow-wrap: anywhere;
}
.dx-field-req,
.dx-field-lock {
  margin-left: 6px;
  font-size: 11px;
  font-weight: 400;
  border-radius: 3px;
  padding: 0 4px;
}
.dx-field-req {
  color: #b3261e;
  background: #fce8e6;
}
.dx-field-lock {
  color: #5f6368;
  background: #f1f3f4;
}
.dx-field-meta {
  font-size: 12px;
  color: #5f6368;
  overflow-wrap: anywhere;
}
</style>
