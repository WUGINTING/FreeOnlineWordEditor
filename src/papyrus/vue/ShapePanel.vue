<script setup lang="ts">
import { tl } from './locale';
// 圖形格式 (the ribbon's tab while a shape is selected, like 圖片格式): fill and outline, 文繞圖,
// stacking, align and distribute, group, and the position and size in cm (editor/shapeInteract.ts).
import { computed, ref, watch } from 'vue';
import type { DocxEditor, EditorSnapshot } from '../editor/core';
import type { WrapChoice } from '../docx/shapeOps';
import ColorPicker, { STANDARD_COLORS } from './ColorPicker.vue';
import { onEnter } from './keys';

const props = defineProps<{ editor: DocxEditor | null; snapshot: EditorSnapshot | null }>();

const info = computed(() => {
  void props.snapshot; // re-evaluated on every editor update
  return props.editor?.shapes.info() ?? null;
});
const tools = () => props.editor?.shapes;

const WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
const WRAPS: { value: WrapChoice; label: string }[] = [
  { value: 'inline', label: '與文字排列' },
  { value: 'square', label: '矩形' },
  { value: 'topAndBottom', label: '上及下' },
  { value: 'front', label: '文字在前' },
  { value: 'behind', label: '文字在後' },
];

const x = ref('');
const y = ref('');
const w = ref('');
const h = ref('');
const lock = ref(false);
const show = (v: number | null | undefined) => (v == null ? '' : String(v));
function sync() {
  x.value = show(info.value?.x);
  y.value = show(info.value?.y);
  w.value = show(info.value?.w);
  h.value = show(info.value?.h);
}
watch(() => [info.value?.x, info.value?.y, info.value?.w, info.value?.h], sync, { immediate: true });

const num = (t: string) => {
  const n = Number(t.trim());
  return Number.isFinite(n) ? n : null;
};

function applyPosition() {
  const i = info.value;
  if (!i) return;
  const nx = num(x.value);
  const ny = num(y.value);
  if (nx == null || ny == null) return sync();
  if (nx === i.x && ny === i.y) return;
  tools()?.setGeometry({ x: nx, y: ny });
}

function applySize(which: 'w' | 'h') {
  const i = info.value;
  if (!i) return;
  let nw = num(w.value);
  let nh = num(h.value);
  if (nw == null || nh == null || nw < 0 || nh < 0) return sync();
  if (nw === i.w && nh === i.h) return;
  if (lock.value && i.w > 0 && i.h > 0) {
    if (which === 'w') nh = Math.round(((nw * i.h) / i.w) * 100) / 100;
    else nw = Math.round(((nh * i.w) / i.h) * 100) / 100;
  }
  tools()?.setGeometry({ w: nw, h: nh });
}

const alignOpen = ref(false);
function align(how: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom') {
  alignOpen.value = false;
  tools()?.align(how);
}
function distribute(axis: 'h' | 'v') {
  alignOpen.value = false;
  tools()?.distribute(axis);
}
function onWrap(e: Event) {
  tools()?.setWrap((e.target as HTMLSelectElement).value as WrapChoice);
}
function onWidth(e: Event) {
  const v = Number((e.target as HTMLSelectElement).value);
  tools()?.setOutline(info.value?.line ?? '#000000', v);
}
</script>

<template>
  <div v-if="info" class="dx-group dx-shape-panel" role="group" :aria-label="tl('圖形格式')">
    <span class="dx-shape-name" aria-live="off">{{ info.count > 1 ? `已選取 ${info.count} 個圖形` : info.name }}</span>
    <template v-if="info.drawable">
      <span class="dx-field" role="group" :aria-label="tl('圖案樣式')">
        <ColorPicker v-if="!info.line1" :label="tl('圖案填滿')" :colors="STANDARD_COLORS" :value="info.fill" :none-label="tl('無填滿')" @pick="(c) => tools()?.setFill(c)">
          <span class="dx-cp-glyph">{{ tl('填滿') }}</span>
        </ColorPicker>
        <ColorPicker :label="tl('圖案外框')" :colors="STANDARD_COLORS" :value="info.line" :none-label="tl('無外框')" @pick="(c) => tools()?.setOutline(c, info?.lineWidth ?? undefined)">
          <span class="dx-cp-glyph">{{ tl('外框') }}</span>
        </ColorPicker>
        <select :aria-label="tl('外框粗細（點）')" :title="tl('外框粗細')" :value="info.lineWidth ?? ''" @change="onWidth">
          <option v-if="info.lineWidth == null" value="" disabled>{{ tl('粗細') }}</option>
          <option v-for="v in WIDTHS" :key="v" :value="v">{{ v }} {{ tl('點') }}</option>
          <option v-if="info.lineWidth != null && !WIDTHS.includes(info.lineWidth)" :value="info.lineWidth">{{ info.lineWidth }} {{ tl('點') }}</option>
        </select>
      </span>
      <label class="dx-field">
        {{ tl('文繞圖') }}
        <select :aria-label="tl('文繞圖')" :value="info.wrap" @change="onWrap">
          <option v-for="o in WRAPS" :key="o.value" :value="o.value">{{ o.label }}</option>
        </select>
      </label>
      <span class="dx-field" role="group" :aria-label="tl('排列')">
        <button type="button" :title="tl('上移一層')" :disabled="!info.floating" @click="tools()?.restack(1)">{{ tl('上移一層') }}</button>
        <button type="button" :title="tl('下移一層')" :disabled="!info.floating" @click="tools()?.restack(-1)">{{ tl('下移一層') }}</button>
        <span class="dx-pop">
          <button type="button" :title="tl('對齊與均分')" aria-haspopup="menu" :aria-expanded="alignOpen" :disabled="!info.floating" @click="alignOpen = !alignOpen">{{ tl('對齊 ▾') }}</button>
          <span v-if="alignOpen" class="dx-menu dx-popup" role="menu" :aria-label="tl('對齊')" @keydown.esc.stop.prevent="alignOpen = false">
            <button type="button" role="menuitem" @click="align('left')">{{ tl('靠左對齊') }}</button>
            <button type="button" role="menuitem" @click="align('center')">{{ tl('水平置中') }}</button>
            <button type="button" role="menuitem" @click="align('right')">{{ tl('靠右對齊') }}</button>
            <button type="button" role="menuitem" @click="align('top')">{{ tl('靠上對齊') }}</button>
            <button type="button" role="menuitem" @click="align('middle')">{{ tl('垂直置中') }}</button>
            <button type="button" role="menuitem" @click="align('bottom')">{{ tl('靠下對齊') }}</button>
            <button type="button" role="menuitem" :disabled="info.count < 3" @click="distribute('h')">{{ tl('水平均分') }}</button>
            <button type="button" role="menuitem" :disabled="info.count < 3" @click="distribute('v')">{{ tl('垂直均分') }}</button>
          </span>
        </span>
        <button type="button" :title="tl('群組 (Ctrl+G)：先按住 Shift 點選多個圖形')" :disabled="!info.canGroup" @click="tools()?.group()">{{ tl('群組') }}</button>
        <button type="button" :title="tl('取消群組 (Ctrl+Shift+G)')" :disabled="!info.canUngroup" @click="tools()?.ungroup()">{{ tl('取消群組') }}</button>
      </span>
      <span class="dx-field" role="group" :aria-label="tl('位置')">
        <template v-if="info.floating">
          {{ tl('水平') }}
          <input v-model="x" type="text" inputmode="decimal" :aria-label="`水平位置（公分，相對於${info.xFrom}）`" :title="`相對於${info.xFrom}`" @change="applyPosition" @keydown="onEnter(applyPosition)($event)" />
          {{ tl('垂直') }}
          <input v-model="y" type="text" inputmode="decimal" :aria-label="`垂直位置（公分，相對於${info.yFrom}）`" :title="`相對於${info.yFrom}`" @change="applyPosition" @keydown="onEnter(applyPosition)($event)" />
          {{ tl('公分') }}
        </template>
      </span>
      <span class="dx-field" role="group" :aria-label="tl('大小')">
        {{ tl('寬') }}
        <input v-model="w" type="text" inputmode="decimal" :aria-label="tl('寬度（公分）')" @change="applySize('w')" @keydown="onEnter(() => applySize('w'))($event)" />
        {{ tl('高') }}
        <input v-model="h" type="text" inputmode="decimal" :aria-label="tl('高度（公分）')" @change="applySize('h')" @keydown="onEnter(() => applySize('h'))($event)" />
        {{ tl('公分') }}
        <label class="dx-check" :title="tl('調整寬度或高度時維持長寬比（拖曳控點時按住 Shift 也可以）')"><input v-model="lock" type="checkbox" />{{ tl('鎖定比例') }}</label>
      </span>
    </template>
    <span v-else class="dx-msg">{{ tl('這個圖形只能在 Word 修改') }}</span>
    <button type="button" :title="tl('選取下一個圖形 (Tab)')" @click="editor?.selectShape(1)">{{ tl('選取下一個') }}</button>
    <button type="button" :title="tl('刪除選取的圖形 (Delete)')" @click="tools()?.deleteSelected()">{{ tl('刪除') }}</button>
  </div>
</template>

<style scoped>
.dx-shape-panel {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 2px 8px;
  font-size: 12px;
}
.dx-shape-name {
  font-weight: 600;
  max-width: 260px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.dx-field {
  display: inline-flex;
  align-items: center;
  gap: 3px;
  white-space: nowrap;
}
button,
select {
  height: 28px;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: pointer;
  padding: 0 6px;
}
select {
  border-color: #dadce0;
  padding: 0 2px;
}
button:hover:not(:disabled) {
  background: #f1f3f4;
}
button:disabled {
  opacity: 0.45;
  cursor: default;
}
button:focus-visible,
select:focus-visible,
input:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}
input[type='text'] {
  width: 48px;
  height: 24px;
  box-sizing: border-box;
  border: 1px solid #dadce0;
  border-radius: 4px;
  padding: 0 4px;
  font: inherit;
}
.dx-check {
  display: inline-flex;
  align-items: center;
  gap: 3px;
}
.dx-pop {
  position: relative;
}
.dx-menu {
  position: absolute;
  z-index: 20;
  top: 100%;
  left: 0;
  display: flex;
  flex-direction: column;
  min-width: 120px;
  padding: 4px 0;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 4px;
  box-shadow: 0 2px 6px rgba(60, 64, 67, 0.3);
}
.dx-menu button {
  text-align: left;
  border-radius: 0;
}
.dx-msg {
  color: #5f6368;
}
</style>
