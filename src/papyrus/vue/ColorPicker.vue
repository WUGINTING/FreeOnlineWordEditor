<script lang="ts">
/** A colour with the name Word gives it in Chinese. */
export interface NamedColor {
  name: string;
  value: string;
}

/** Word's 標準色彩 (text colour, cell shading). */
export const STANDARD_COLORS: NamedColor[] = [
  { name: '深紅', value: '#c00000' },
  { name: '紅色', value: '#ff0000' },
  { name: '橙色', value: '#ffc000' },
  { name: '黃色', value: '#ffff00' },
  { name: '淺綠', value: '#92d050' },
  { name: '綠色', value: '#00b050' },
  { name: '淺藍', value: '#00b0f0' },
  { name: '藍色', value: '#0070c0' },
  { name: '深藍', value: '#002060' },
  { name: '紫色', value: '#7030a0' },
  { name: '黑色', value: '#000000' },
  { name: '深灰', value: '#595959' },
  { name: '灰色', value: '#a6a6a6' },
  { name: '淺灰', value: '#d9d9d9' },
  { name: '白色', value: '#ffffff' },
];

let menus = 0;

/** Word's 螢光標記 colours (w:highlight names, docx/props.ts HIGHLIGHT_COLORS). */
export const HIGHLIGHT_PALETTE: NamedColor[] = [
  { name: '黃色', value: '#ffff00' },
  { name: '亮綠色', value: '#00ff00' },
  { name: '青綠色', value: '#00ffff' },
  { name: '粉紅色', value: '#ff00ff' },
  { name: '藍色', value: '#0000ff' },
  { name: '紅色', value: '#ff0000' },
  { name: '深藍色', value: '#000080' },
  { name: '藍綠色', value: '#008080' },
  { name: '綠色', value: '#008000' },
  { name: '紫色', value: '#800080' },
  { name: '深紅色', value: '#800000' },
  { name: '深黃色', value: '#808000' },
  { name: '灰色 50%', value: '#808080' },
  { name: '灰色 25%', value: '#c0c0c0' },
  { name: '黑色', value: '#000000' },
];
</script>

<script setup lang="ts">
import { tl } from './locale';
// A Word-like colour menu (persona-300): the standard colours as named swatches (each with an
// aria-label and title, 「紅色」「深藍」…), 自動 / 無色彩, and 「其他色彩…」 (the browser's own
// colour picker) for any other colour. The button shows the current colour; Escape closes the
// menu back on it, a click elsewhere closes it.
import { computed, nextTick, onBeforeUnmount, onMounted, ref } from 'vue';
import { composing } from './keys';

const props = withDefaults(
  defineProps<{
    /** What it colours (「文字顏色」): the button's and the menu's name. */
    label: string;
    colors: NamedColor[];
    /** The current colour ("#rrggbb"), shown on the button. */
    value?: string | null;
    /** The first choice: 「自動」 (text colour) or 「無色彩」 (highlight, shading); none when not given. */
    noneLabel?: string | null;
    /** More than one colour in the selection. */
    mixed?: boolean;
  }>(),
  { value: null, noneLabel: null, mixed: false },
);
const emit = defineEmits<{ (e: 'pick', value: string | null): void }>();

const open = ref(false);
const root = ref<HTMLElement | null>(null);
const button = ref<HTMLButtonElement | null>(null);
const lower = (v: string | null | undefined) => (v ?? '').toLowerCase();
const menuId = `dx-cp-${++menus}`;
/** The button's name: what it colours and the colour it shows now (「文字顏色：深藍」). */
const buttonName = computed(() => {
  const named = props.colors.find((c) => c.value === lower(props.value))?.name;
  const current = props.mixed
    ? tl('多種顏色')
    : !props.value
      ? props.noneLabel ?? tl('無')
      : named != null ? tl(named) : tl('其他色彩 {0}', props.value.toUpperCase());
  return tl('{0}：{1}', props.label, current);
});

function toggle() {
  open.value = !open.value;
  if (open.value) nextTick(() => root.value?.querySelector<HTMLElement>('.dx-cp-menu button')?.focus());
}
function close(refocus = true) {
  open.value = false;
  if (refocus) button.value?.focus();
}
function pick(v: string | null) {
  close();
  emit('pick', v);
}
function onOther(e: Event) {
  pick((e.target as HTMLInputElement).value);
}
// Escape closes an open menu only; with the menu closed it goes on to the ribbon / page.
function onKey(e: KeyboardEvent) {
  if (!open.value || composing(e) || e.key !== 'Escape') return;
  e.preventDefault();
  e.stopPropagation();
  close();
}
// Arrow keys move between the swatches (five in a row).
function onGridKey(e: KeyboardEvent) {
  const moves: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -5, ArrowDown: 5 };
  const d = moves[e.key];
  if (!d) return;
  const items = Array.from(root.value?.querySelectorAll<HTMLElement>('.dx-cp-swatch') ?? []);
  const i = items.indexOf(document.activeElement as HTMLElement);
  if (i < 0) return;
  e.preventDefault();
  e.stopPropagation();
  items[Math.max(0, Math.min(items.length - 1, i + d))]?.focus();
}
function onDocumentDown(e: MouseEvent) {
  if (open.value && !root.value?.contains(e.target as Node)) close(false);
}
onMounted(() => document.addEventListener('mousedown', onDocumentDown, true));
onBeforeUnmount(() => document.removeEventListener('mousedown', onDocumentDown, true));
</script>

<template>
  <span ref="root" class="dx-cp" @keydown="onKey">
    <button
      ref="button"
      type="button"
      class="dx-cp-button"
      :title="mixed ? tl('{0}（選取範圍有多種顏色）', label) : label"
      :aria-label="buttonName"
      aria-haspopup="dialog"
      :aria-expanded="open"
      :aria-controls="open ? menuId : undefined"
      @click="toggle"
    >
      <slot />
      <span class="dx-cp-bar" :class="{ 'dx-cp-mixed': mixed }" :style="{ background: value ?? 'transparent' }" />
      <span class="dx-cp-caret" aria-hidden="true">▾</span>
    </button>
    <div v-if="open" :id="menuId" class="dx-cp-menu dx-popup" role="dialog" :aria-label="label">
      <button v-if="noneLabel" type="button" class="dx-cp-none" :aria-pressed="!value" @click="pick(null)">{{ noneLabel }}</button>
      <div class="dx-cp-grid" role="group" :aria-label="tl('標準色彩')" @keydown="onGridKey">
        <button
          v-for="c in colors"
          :key="c.value"
          type="button"
          class="dx-cp-swatch"
          :title="tl(c.name)"
          :aria-label="tl(c.name)"
          :aria-pressed="lower(value) === c.value"
          :style="{ background: c.value }"
          @click="pick(c.value)"
        />
      </div>
      <label class="dx-cp-other">
        {{ tl('其他色彩…') }}
        <input type="color" :aria-label="tl('{0}：其他色彩', label)" :value="value ?? '#000000'" @change="onOther" />
      </label>
    </div>
  </span>
</template>

<style scoped>
.dx-cp {
  position: relative;
  display: inline-flex;
}
.dx-cp-button {
  position: relative;
  display: inline-flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 1px;
  height: 1.75rem;
  min-width: 2rem;
  padding: 0 14px 0 6px;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: pointer;
}
.dx-cp-button:hover {
  background: #f3f2f1;
}
.dx-cp-button:focus-visible,
.dx-cp-menu button:focus-visible,
.dx-cp-other:focus-within {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}
.dx-cp-bar {
  width: 16px;
  height: 4px;
  border: 1px solid rgba(0, 0, 0, 0.25);
}
.dx-cp-mixed {
  background: linear-gradient(90deg, #d93025, #188038, #1a73e8) !important;
}
.dx-cp-caret {
  position: absolute;
  right: 3px;
  font-size: 0.625rem;
  color: #605e5c;
}
.dx-cp-menu {
  position: absolute;
  z-index: 45;
  top: calc(100% + 2px);
  left: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
  padding: 8px;
  background: #fff;
  border: 1px solid #c8c6c4;
  border-radius: 6px;
  box-shadow: 0 6px 16px rgba(0, 0, 0, 0.18);
}
.dx-cp-none,
.dx-cp-other {
  display: flex;
  align-items: center;
  justify-content: flex-start;
  min-height: 1.75rem;
  padding: 0 8px;
  border: 1px solid #e1dfdd;
  border-radius: 4px;
  background: #fff;
  font: inherit;
  cursor: pointer;
  white-space: nowrap;
}
.dx-cp-other {
  position: relative;
  overflow: hidden;
}
.dx-cp-other input {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  opacity: 0;
  cursor: pointer;
}
.dx-cp-grid {
  display: grid;
  grid-template-columns: repeat(5, 1.5rem);
  gap: 4px;
}
.dx-cp-swatch {
  width: 1.5rem;
  height: 1.5rem;
  min-width: 0;
  padding: 0;
  border: 1px solid rgba(0, 0, 0, 0.3);
  border-radius: 3px;
  cursor: pointer;
}
.dx-cp-swatch[aria-pressed='true'] {
  outline: 2px solid #185abd;
  outline-offset: 1px;
}
</style>
