<script setup lang="ts">
import { nextTick, onBeforeUnmount, onMounted, ref } from 'vue';
import type { DocxEditor } from '../editor/core';
import { menuSymbols, symbolCode } from '../editor/symbols';
import { readRecentSymbols } from '../editor/recentSymbols';
import { insertSymbolAndRemember } from '../editor/symbolActions';
import { ICONS } from './ribbon';

// Word's 插入 › 符號 gallery on the ribbon: 20 symbols, the recently used ones first, in rows of
// five, and 「其他符號(M)…」 below them (emits `more`: the page opens the 符號 dialog). A click, Enter
// or Space inserts at the cursor and closes the gallery. The arrow keys move in the rows and
// columns (Down from the last row and Up from the first go to 其他符號), Home / End to the row's
// ends, M to 其他符號, Escape back to the button. The ribbon supplies the button (the default slot,
// given `open` and `toggle`), so it looks like the ribbon's other buttons; it sits in a `.dx-popup`,
// which the ribbon's arrow-key order leaves out.
const props = defineProps<{ editor: DocxEditor | null; disabled?: boolean }>();
const emit = defineEmits<{ (e: 'more'): void }>();

const COLUMNS = 5;
const open = ref(false);
const choices = ref<string[]>([]);
const root = ref<HTMLElement | null>(null);
const menu = ref<HTMLElement | null>(null);
let trigger: HTMLElement | null = null;

const cells = () => Array.from(menu.value?.querySelectorAll<HTMLElement>('.dx-sym') ?? []);
const moreItem = () => menu.value?.querySelector<HTMLElement>('.dx-symmore') ?? null;

function toggle(e: MouseEvent) {
  if (open.value) return close(false);
  if (props.disabled || !props.editor?.editable) return;
  trigger = e.currentTarget as HTMLElement;
  choices.value = menuSymbols(readRecentSymbols());
  open.value = true;
  nextTick(() => {
    fit();
    cells()[0]?.focus();
  });
}

/** The group is the 插入 tab's last: in a narrow window the gallery would stick out on the right. */
function fit() {
  const el = menu.value;
  if (!el) return;
  const width = document.documentElement.clientWidth;
  const box = el.getBoundingClientRect();
  if (box.right > width - 8) el.style.left = `${Math.round(Math.max(8 - box.left, width - 8 - box.right))}px`;
}

function close(refocus: boolean) {
  open.value = false;
  if (refocus) trigger?.focus();
}

function pick(ch: string) {
  close(false);
  if (props.editor) insertSymbolAndRemember(props.editor, ch);
}

function more() {
  close(false);
  emit('more');
}

function onKey(e: KeyboardEvent) {
  const list = cells();
  const extra = moreItem();
  const i = list.indexOf(e.target as HTMLElement);
  const go = (el: HTMLElement | null | undefined) => {
    e.preventDefault();
    e.stopPropagation();
    el?.focus();
  };
  if (e.key === 'Escape') {
    e.preventDefault();
    e.stopPropagation();
    close(true);
    return;
  }
  if (e.key === 'Tab') {
    close(false);
    return;
  }
  if ((e.key === 'm' || e.key === 'M') && !e.ctrlKey && !e.altKey && !e.metaKey) {
    go(null);
    more();
    return;
  }
  const lastRow = Math.floor((list.length - 1) / COLUMNS) * COLUMNS;
  if (i < 0) {
    // On 「其他符號…」: Up back to the gallery's last row, Down round to the first symbol.
    if (e.key === 'ArrowUp' || e.key === 'End') return go(list[lastRow]);
    if (e.key === 'ArrowDown' || e.key === 'Home') return go(list[0]);
    return;
  }
  const start = i - (i % COLUMNS);
  if (e.key === 'ArrowRight') go(list[(i + 1) % list.length]);
  else if (e.key === 'ArrowLeft') go(list[(i + list.length - 1) % list.length]);
  else if (e.key === 'ArrowDown') go(i + COLUMNS < list.length ? list[i + COLUMNS] : extra);
  else if (e.key === 'ArrowUp') go(i - COLUMNS >= 0 ? list[i - COLUMNS] : extra);
  else if (e.key === 'Home') go(e.ctrlKey ? list[0] : list[start]);
  else if (e.key === 'End') go(e.ctrlKey ? extra : list[Math.min(list.length - 1, start + COLUMNS - 1)]);
}

/** A click outside closes it. */
function onDocumentDown(e: MouseEvent) {
  if (open.value && !root.value?.contains(e.target as Node)) close(false);
}
onMounted(() => document.addEventListener('mousedown', onDocumentDown, true));
onBeforeUnmount(() => document.removeEventListener('mousedown', onDocumentDown, true));

const label = (ch: string) => `${ch}（${symbolCode(ch)}）`;
</script>

<template>
  <div ref="root" class="dx-pop">
    <slot :open="open" :toggle="toggle" />
    <div v-if="open" ref="menu" class="dx-menu dx-popup dx-symmenu" role="menu" aria-label="符號" @keydown="onKey">
      <div class="dx-symgrid">
        <button
          v-for="ch in choices"
          :key="ch"
          type="button"
          role="menuitem"
          tabindex="-1"
          class="dx-sym"
          :title="label(ch)"
          :aria-label="label(ch)"
          @click="pick(ch)"
        >{{ ch }}</button>
      </div>
      <button type="button" role="menuitem" tabindex="-1" class="dx-symmore" title="開啟「符號」對話方塊，選擇更多符號或輸入字元代碼" @click="more">
        <svg class="dx-ico" viewBox="0 0 24 24" v-html="ICONS.symbol" />其他符號(M)…
      </button>
    </div>
  </div>
</template>

<style scoped>
.dx-pop {
  position: relative;
  display: inline-flex;
}
/* As the ribbon's other menus (DocxToolbar.vue .dx-menu). */
.dx-menu {
  position: absolute;
  z-index: 40;
  top: calc(100% + 2px);
  left: 0;
  display: flex;
  flex-direction: column;
  padding: 4px;
  background: #fff;
  border: 1px solid #c8c6c4;
  border-radius: 6px;
  box-shadow: 0 6px 16px rgba(0, 0, 0, 0.18);
}
.dx-menu button {
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
}
.dx-menu button:hover {
  background: #f3f2f1;
}
.dx-menu button:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: -2px;
}
/* Word's gallery: rows of five, 「其他符號…」 below. */
.dx-symgrid {
  display: grid;
  grid-template-columns: repeat(5, 34px);
  gap: 2px;
  padding: 2px 2px 4px;
  border-bottom: 1px solid #e1dfdd;
  margin-bottom: 4px;
}
.dx-menu .dx-sym {
  justify-content: center;
  width: 34px;
  height: 34px;
  padding: 0;
  border-color: #e1dfdd;
  font-size: 18px;
  line-height: 1;
}
.dx-menu .dx-symmore {
  gap: 8px;
  min-height: 30px;
  padding: 4px 10px 4px 4px;
  white-space: nowrap;
}
.dx-ico {
  width: 16px;
  height: 16px;
  flex: none;
  fill: none;
  stroke: currentColor;
  stroke-width: 1.6;
  stroke-linecap: round;
  stroke-linejoin: round;
}
</style>
