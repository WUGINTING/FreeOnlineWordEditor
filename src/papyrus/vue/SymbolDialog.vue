<script setup lang="ts">
import { tl } from './locale';
import { computed, nextTick, ref, watch } from 'vue';
import type { DocxEditor } from '../editor/core';
import { SYMBOL_CATEGORIES, parseSymbolCode, symbolCategory, symbolCode, symbolHex } from '../editor/symbols';
import { readRecentSymbols } from '../editor/recentSymbols';
import { insertSymbolAndRemember } from '../editor/symbolActions';

// Word's 符號 dialog (插入 › 符號 › 其他符號…): a 子集 list, the subset's symbols in a grid, the
// selected symbol's 字元代碼 (which can also be typed), and the 最近使用過的符號 row. 插入, Enter or a
// double-click inserts at the cursor (the body, or the header/footer being edited) and, as in
// Word, the dialog stays open, with the keyboard, so several symbols can go in one after
// another; 關閉 or Esc closes it, back to the text. Each insertion is one undo step
// (commands.ts insertSymbol). In the grids (WAI-ARIA grid) the arrow keys, Home / End (Ctrl: the
// first / last symbol) and PageUp / PageDown move and select; Space selects, Enter inserts.
const props = withDefaults(defineProps<{ editor: DocxEditor | null; open: boolean; editable?: boolean }>(), { editable: true });
const emit = defineEmits<{ (e: 'close'): void }>();

/** Symbols per row of the grid. */
const COLUMNS = 16;
/** Rows PageUp / PageDown move (the rows the grid shows). */
const PAGE_ROWS = 8;

const categoryId = ref(SYMBOL_CATEGORIES[0].id);
const category = computed(() => SYMBOL_CATEGORIES.find((c) => c.id === categoryId.value) ?? SYMBOL_CATEGORIES[0]);
/** The grid's rows: each symbol with its index in the subset. */
const rows = computed(() => {
  const out: { ch: string; i: number }[][] = [];
  category.value.chars.forEach((ch, i) => {
    if (i % COLUMNS === 0) out.push([]);
    out[out.length - 1].push({ ch, i });
  });
  return out;
});
/** The grid's Tab stop (roving tabindex): the index of the symbol last moved to. */
const active = ref(0);
const recent = ref<string[]>([]);
const recentActive = ref(0);
/** The 字元代碼 box (hexadecimal): what 插入 inserts. Picking a symbol fills it in. */
const code = ref('');
const error = ref('');
const dialog = ref<HTMLElement | null>(null);
const grid = ref<HTMLElement | null>(null);
const recentGrid = ref<HTMLElement | null>(null);

/** The character the code box stands for, or null while it isn't one. */
const current = computed(() => {
  const r = parseSymbolCode(code.value);
  return r.ok ? r.char : null;
});
/** Read-only (or no document): nothing can be inserted, and the dialog doesn't show. */
const canInsert = computed(() => props.editable && !!props.editor);
const shown = computed(() => props.open && props.editable);

watch(
  shown,
  (open) => {
    if (!open) return;
    recent.value = readRecentSymbols();
    recentActive.value = 0;
    error.value = '';
    const chars = category.value.chars;
    active.value = Math.min(active.value, chars.length - 1);
    if (!current.value) code.value = symbolHex(chars[active.value]);
    nextTick(() => focusCell(active.value));
  },
  { immediate: true },
);
// The document became read-only while the dialog was open.
watch(
  () => props.editable,
  (editable) => {
    if (!editable && props.open) emit('close');
  },
);

function close() {
  emit('close');
  // Back to the text, where the symbols went in.
  nextTick(() => props.editor?.activeView?.focus());
}

/** Select a symbol: its code goes into the code box, and it is what 插入 inserts. */
function choose(ch: string) {
  code.value = symbolHex(ch);
  error.value = '';
}

const gridCells = () => Array.from(grid.value?.querySelectorAll<HTMLElement>('[role="gridcell"]') ?? []);
const recentCells = () => Array.from(recentGrid.value?.querySelectorAll<HTMLElement>('[role="gridcell"]') ?? []);
function focusCell(i: number) {
  gridCells()[i]?.focus();
}

function onCategory() {
  active.value = 0;
  choose(category.value.chars[0]);
}

/** A code typed in: the symbol is selected (and scrolled to) in its subset's grid when it is in one. */
function onCode() {
  error.value = '';
  const ch = current.value;
  if (!ch) return;
  const here = category.value.chars.indexOf(ch);
  if (here >= 0) active.value = here;
  else {
    const other = symbolCategory(ch);
    if (!other) return;
    categoryId.value = other.id;
    active.value = other.chars.indexOf(ch);
  }
  nextTick(() => gridCells()[active.value]?.scrollIntoView?.({ block: 'nearest' }));
}

/** Enter in the code box inserts (not while an input method is still composing). */
function onCodeEnter(e: KeyboardEvent) {
  if (e.isComposing || e.keyCode === 229) return;
  e.preventDefault();
  insert();
}

/**
 * Insert `ch` (default: the code box's character) at the cursor. The dialog stays open and keeps
 * the keyboard: on the control it was on, or, from the recently used row (which has just put that
 * symbol first), on the row's first symbol.
 */
function insert(ch?: string) {
  let value = ch ?? null;
  if (!value) {
    const r = parseSymbolCode(code.value);
    if (!r.ok) {
      error.value = r.error;
      return;
    }
    value = r.char;
  }
  const ed = props.editor;
  if (!ed || !canInsert.value) return;
  const back = document.activeElement as HTMLElement | null;
  const fromRecent = !!back && !!recentGrid.value?.contains(back);
  const changed = insertSymbolAndRemember(ed, value, { focus: false });
  if (changed) {
    recent.value = readRecentSymbols();
    recentActive.value = 0; // the symbol just used is first now
  }
  // Moving the focused cell to the front of the row takes the focus off it (the browser blurs a
  // moved element): put it back once the row is drawn again. Nothing moved when nothing went in.
  nextTick(() => {
    if (fromRecent && changed) recentCells()[0]?.focus();
    else if (back?.isConnected) {
      if (document.activeElement !== back) back.focus();
    } else focusCell(active.value);
  });
}

/** Where a grid key moves to from `i` in a list of `count` in rows of `columns`, or null. */
function moveTo(e: KeyboardEvent, i: number, count: number, columns: number): number | null {
  const last = count - 1;
  const start = i - (i % columns);
  const page = PAGE_ROWS * columns;
  switch (e.key) {
    case 'ArrowRight':
      return Math.min(last, i + 1);
    case 'ArrowLeft':
      return Math.max(0, i - 1);
    case 'ArrowDown':
      return i + columns <= last ? i + columns : i;
    case 'ArrowUp':
      return i - columns >= 0 ? i - columns : i;
    case 'PageDown': {
      // The same column a page further down, or in the lowest row that has it.
      let to = i + page;
      while (to > last && to - columns > i) to -= columns;
      return to > last ? i : to;
    }
    case 'PageUp': {
      let to = i - page;
      while (to < 0 && to + columns < i) to += columns;
      return to < 0 ? i : to;
    }
    case 'Home':
      return e.ctrlKey ? 0 : start;
    case 'End':
      return e.ctrlKey ? last : Math.min(last, start + columns - 1);
    default:
      return null;
  }
}

/** Keys in a grid of `list` (rows of `columns`) whose Tab stop is `index`; `move` focuses a cell. */
function gridKey(e: KeyboardEvent, list: readonly string[], columns: number, index: number, move: (i: number) => void) {
  if (e.altKey || e.metaKey) return;
  if (e.key === 'Enter') {
    e.preventDefault();
    if (list[index]) insert(list[index]);
    return;
  }
  if (e.key === ' ') {
    e.preventDefault();
    if (list[index]) choose(list[index]);
    return;
  }
  const next = moveTo(e, index, list.length, columns);
  if (next == null) return;
  e.preventDefault();
  move(next);
  choose(list[next]);
}

function onGridKey(e: KeyboardEvent) {
  gridKey(e, category.value.chars, COLUMNS, active.value, (i) => {
    active.value = i;
    focusCell(i);
  });
}

function onRecentKey(e: KeyboardEvent) {
  gridKey(e, recent.value, Math.max(1, recent.value.length), recentActive.value, (i) => {
    recentActive.value = i;
    recentCells()[i]?.focus();
  });
}

/** A recently used symbol got the focus (found by its character: the row may just have changed order). */
function focusRecent(ch: string) {
  recentActive.value = Math.max(0, recent.value.indexOf(ch));
}

function trapFocus(e: KeyboardEvent) {
  if (e.key !== 'Tab' || !dialog.value) return;
  const items = Array.from(dialog.value.querySelectorAll<HTMLElement>('select, input, button, [tabindex="0"]')).filter(
    (el) => !el.hasAttribute('disabled'),
  );
  if (!items.length) return;
  const first = items[0];
  const last = items[items.length - 1];
  if (e.shiftKey && document.activeElement === first) {
    e.preventDefault();
    last.focus();
  } else if (!e.shiftKey && document.activeElement === last) {
    e.preventDefault();
    first.focus();
  }
}

const label = (ch: string) => tl('{0}（{1}）', ch, symbolCode(ch));
</script>

<template>
  <div v-if="shown" class="dx-dialog-backdrop" @mousedown.self="close" @keydown.esc="close">
    <div ref="dialog" class="dx-dialog" role="dialog" aria-modal="true" aria-labelledby="dx-sy-title" @keydown="trapFocus">
      <h3 id="dx-sy-title">{{ tl('符號') }}</h3>

      <label class="dx-sy-subset">
        {{ tl('子集') }}
        <select v-model="categoryId" @change="onCategory">
          <option v-for="c in SYMBOL_CATEGORIES" :key="c.id" :value="c.id">{{ tl(c.label) }}</option>
        </select>
      </label>

      <div ref="grid" class="dx-sy-grid" role="grid" :aria-label="tl('符號')" @keydown="onGridKey">
        <div v-for="(row, r) in rows" :key="r" class="dx-sy-row" role="row">
          <div
            v-for="cell in row"
            :key="cell.ch"
            class="dx-sy-cell"
            role="gridcell"
            :tabindex="cell.i === active ? 0 : -1"
            :aria-selected="cell.ch === current"
            :aria-label="label(cell.ch)"
            :title="label(cell.ch)"
            @focus="active = cell.i"
            @click="active = cell.i; choose(cell.ch)"
            @dblclick="insert(cell.ch)"
          >{{ cell.ch }}</div>
        </div>
      </div>

      <div class="dx-sy-recent">
        <span id="dx-sy-recent-label">{{ tl('最近使用過的符號') }}</span>
        <div
          v-if="recent.length"
          ref="recentGrid"
          class="dx-sy-grid dx-sy-recent-grid"
          role="grid"
          aria-labelledby="dx-sy-recent-label"
          @keydown="onRecentKey"
        >
          <div class="dx-sy-row" role="row">
            <div
              v-for="(ch, i) in recent"
              :key="ch"
              class="dx-sy-cell"
              role="gridcell"
              :tabindex="i === recentActive ? 0 : -1"
              :aria-selected="ch === current"
              :aria-label="label(ch)"
              :title="label(ch)"
              @focus="focusRecent(ch)"
              @click="focusRecent(ch); choose(ch)"
              @dblclick="insert(ch)"
            >{{ ch }}</div>
          </div>
        </div>
        <span v-else class="dx-sy-none">{{ tl('尚未使用過符號') }}</span>
      </div>

      <!-- The code is read out with the symbol (its cell's name), so this line isn't announced again. -->
      <div class="dx-sy-info">
        <span class="dx-sy-glyph" aria-hidden="true">{{ current ?? '' }}</span>
        <span>{{ tl('字元代碼：{0}', current ? symbolCode(current) : '—') }}</span>
        <label class="dx-sy-code">
          {{ tl('字元代碼（16 進位）') }}
          <input
            v-model="code"
            type="text"
            maxlength="8"
            spellcheck="false"
            autocomplete="off"
            :placeholder="tl('例如 2605')"
            @input="onCode"
            @keydown.enter="onCodeEnter"
          />
        </label>
      </div>

      <p v-if="error" class="dx-error" role="alert">{{ error }}</p>
      <div class="dx-actions">
        <button type="button" class="dx-primary" :disabled="!canInsert" :title="tl('在游標處插入選取的符號（對話方塊不會關閉，可以連續插入）')" @click="insert()">{{ tl('插入') }}</button>
        <button type="button" @click="close">{{ tl('關閉') }}</button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.dx-dialog-backdrop {
  position: fixed;
  inset: 0;
  z-index: 1060;
  display: grid;
  place-items: center;
  background: rgba(32, 33, 36, 0.4);
  font: 14px system-ui, 'Microsoft JhengHei', sans-serif;
}
.dx-dialog {
  box-sizing: border-box;
  width: min(560px, calc(100vw - 32px));
  max-height: calc(100vh - 32px);
  overflow: auto;
  background: #fff;
  border-radius: 8px;
  padding: 16px 20px;
  box-shadow: 0 4px 24px rgba(0, 0, 0, 0.25);
  color: #202124;
}
h3 {
  margin: 0 0 10px;
  font-size: 17px;
}
.dx-sy-subset {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 8px;
}
select,
input[type='text'] {
  font: inherit;
  padding: 2px 4px;
}
/* The symbols: 16 to a row (fewer on a narrow screen: the rows scroll sideways inside). */
.dx-sy-grid {
  max-height: 248px;
  overflow: auto;
  border: 1px solid #dadce0;
  border-radius: 4px;
}
.dx-sy-row {
  display: flex;
}
.dx-sy-cell {
  flex: none;
  box-sizing: border-box;
  width: 31px;
  height: 31px;
  display: grid;
  place-items: center;
  border-right: 1px solid #eceff1;
  border-bottom: 1px solid #eceff1;
  font-size: 18px;
  line-height: 1;
  cursor: default;
  user-select: none;
  font-family: 'Microsoft JhengHei', 'PMingLiU', 'Segoe UI Symbol', sans-serif;
}
.dx-sy-cell:hover {
  background: #f1f3f4;
}
.dx-sy-cell[aria-selected='true'] {
  background: #1a73e8;
  color: #fff;
}
.dx-sy-cell:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: -3px;
}
.dx-sy-cell[aria-selected='true']:focus-visible {
  outline-color: #fff;
}
.dx-sy-recent {
  margin-top: 10px;
}
.dx-sy-recent > span {
  display: block;
  margin-bottom: 4px;
  font-size: 13px;
  color: #5f6368;
}
.dx-sy-recent-grid {
  max-height: none;
}
.dx-sy-none {
  font-size: 13px;
  color: #5f6368;
}
.dx-sy-info {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px 16px;
  margin-top: 12px;
}
.dx-sy-glyph {
  display: grid;
  place-items: center;
  width: 44px;
  height: 44px;
  border: 1px solid #dadce0;
  border-radius: 4px;
  font-size: 28px;
  font-family: 'Microsoft JhengHei', 'PMingLiU', 'Segoe UI Symbol', sans-serif;
}
.dx-sy-code {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  margin-left: auto;
}
.dx-sy-code input {
  width: 6em;
  font-family: ui-monospace, Consolas, monospace;
}
.dx-error {
  color: #b3261e;
  font-size: 13px;
}
.dx-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  margin-top: 12px;
}
.dx-actions button {
  font: inherit;
  padding: 4px 14px;
  border: 1px solid #dadce0;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}
.dx-actions .dx-primary {
  background: #1a73e8;
  border-color: #1a73e8;
  color: #fff;
}
.dx-actions button:disabled {
  opacity: 0.5;
  cursor: default;
}
</style>
