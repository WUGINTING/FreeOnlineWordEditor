<script setup lang="ts">
import { tl } from './locale';
// Table appearance for the selected cells / rows: shading, borders, vertical alignment,
// row height, header row, "don't split row", sorting rows. (平均分配欄寬 is in 表格版面配置 ›
// 儲存格大小, as in Word.)
import { computed, ref, watch } from 'vue';
import type { DocxEditor, EditorSnapshot } from '../editor/core';
import {
  distributeColumns, setCantSplit, setCellBackground, setCellBorders, setCellVAlign, setColumnWidth, setHeaderRow, setRowHeight,
  setRowHeightRule, sortRefusal, sortTable, tableInfo, type BorderMode, type HeightRule,
} from '../editor/tableCommands';
import { onEnter } from './keys';
import ColorPicker, { STANDARD_COLORS } from './ColorPicker.vue';

const props = defineProps<{ editor: DocxEditor | null; snapshot: EditorSnapshot | null }>();

const info = computed(() => {
  void props.snapshot; // re-evaluated on every editor update
  const state = props.editor?.activeView?.state;
  return state ? tableInfo(state) : null;
});
const run = (cmd: Parameters<DocxEditor['run']>[0]) => props.editor?.run(cmd);
const vAlignTop = computed(() => !!info.value && (!info.value.vAlign || info.value.vAlign === 'top'));

// Sorting rows by the cursor's column: refused (with the reason) for tables with merged cells.
const keepFirst = ref(true);
const sortBlocked = computed(() => {
  void props.snapshot;
  const state = props.editor?.activeView?.state;
  return state ? sortRefusal(state) : '游標不在表格中。';
});

const BORDERS: { value: BorderMode; label: string }[] = [
  { value: 'all', label: '所有框線' },
  { value: 'outer', label: '外框線' },
  { value: 'inner', label: '內框線' },
  { value: 'none', label: '無框線' },
];

function onBorders(e: Event) {
  const el = e.target as HTMLSelectElement;
  const mode = el.value as BorderMode;
  el.value = '';
  if (mode) run(setCellBorders(mode));
}
function onShading(color: string | null) {
  run(setCellBackground(color));
}

// Row height, shown in cm or pt (stored in twips).
const TWIPS: Record<string, number> = { cm: 1440 / 2.54, pt: 20 };
const unit = ref<'cm' | 'pt'>('cm');
const heightText = ref('');
const rule = ref<HeightRule>('atLeast');
const shownHeight = () => {
  const h = info.value?.height;
  return h ? String(+(h / TWIPS[unit.value]).toFixed(2)) : '';
};
function syncHeight() {
  heightText.value = shownHeight();
  const r = info.value?.heightRule;
  rule.value = r === 'exact' || r === 'auto' ? r : 'atLeast';
}
watch(() => [info.value?.height, info.value?.heightRule, unit.value], syncHeight, { immediate: true });

function applyHeight() {
  const text = heightText.value.trim();
  // The box still shows the (rounded) current height: nothing to change.
  if (text === shownHeight()) return;
  if (!text) {
    run(setRowHeight(null));
    return;
  }
  const n = Number(text);
  if (!Number.isFinite(n) || n <= 0) {
    syncHeight();
    return;
  }
  run(setRowHeight(Math.round(n * TWIPS[unit.value]), rule.value));
}
function applyRule() {
  run(setRowHeightRule(rule.value));
}

// 欄寬（公分） of the cursor's column (or the selected columns): px in the document, cm here, exact in twips.
const PX_PER_CM = 96 / 2.54;
const widthText = ref('');
const shownWidth = () => {
  const w = info.value?.colWidth;
  return w ? String(+(w / PX_PER_CM).toFixed(2)) : '';
};
watch(() => info.value?.colWidth, () => (widthText.value = shownWidth()), { immediate: true });
function applyWidth() {
  const text = widthText.value.trim();
  if (text === shownWidth()) return;
  const cm = Number(text);
  if (!Number.isFinite(cm) || cm < 0.1 || cm > 55) {
    widthText.value = shownWidth();
    return;
  }
  run(setColumnWidth(Math.round((cm * 1440) / 2.54) / 15));
}
</script>

<template>
  <div class="dx-group dx-table-panel" role="group" :aria-label="tl('表格外觀')">
    <ColorPicker :label="tl('儲存格底色')" :colors="STANDARD_COLORS" :value="info?.background ?? null" :none-label="tl('無色彩')" @pick="onShading">
      <span>{{ tl('底色') }}</span>
    </ColorPicker>
    <button type="button" :title="tl('移除儲存格底色')" :disabled="!info?.background" @click="run(setCellBackground(null))">{{ tl('無底色') }}</button>

    <select :title="tl('框線（套用到選取的儲存格）')" value="" @change="onBorders">
      <option value="" disabled>{{ tl('框線') }}</option>
      <option v-for="b in BORDERS" :key="b.value" :value="b.value">{{ b.label }}</option>
    </select>

    <span class="dx-seg" role="group" :aria-label="tl('垂直對齊')">
      <button
        type="button"
        :title="tl('垂直置上')"
        :aria-label="tl('垂直置上')"
        :class="{ on: vAlignTop }"
        :aria-pressed="vAlignTop"
        @click="run(setCellVAlign('top'))"
      >
        <svg viewBox="0 0 16 16"><path d="M2 2.5h12M5 6h6M5 9h6" /></svg>
      </button>
      <button
        type="button"
        :title="tl('垂直置中')"
        :aria-label="tl('垂直置中')"
        :class="{ on: info?.vAlign === 'middle' }"
        :aria-pressed="info?.vAlign === 'middle'"
        @click="run(setCellVAlign('middle'))"
      >
        <svg viewBox="0 0 16 16"><path d="M5 6.5h6M5 9.5h6M2 2h12M2 14h12" /></svg>
      </button>
      <button
        type="button"
        :title="tl('垂直置下')"
        :aria-label="tl('垂直置下')"
        :class="{ on: info?.vAlign === 'bottom' }"
        :aria-pressed="info?.vAlign === 'bottom'"
        @click="run(setCellVAlign('bottom'))"
      >
        <svg viewBox="0 0 16 16"><path d="M5 7h6M5 10h6M2 13.5h12" /></svg>
      </button>
    </span>

    <span class="dx-height" :title="tl('列高（空白 = 自動）')">
      {{ tl('列高') }}
      <input v-model="heightText" type="text" inputmode="decimal" :placeholder="tl('自動')" :aria-label="tl('列高')" @change="applyHeight" @keydown="onEnter(applyHeight)($event)" />
      <select v-model="unit" :aria-label="tl('列高單位')">
        <option value="cm">{{ tl('公分') }}</option>
        <option value="pt">{{ tl('點') }}</option>
      </select>
      <select v-model="rule" :aria-label="tl('列高規則')" @change="applyRule">
        <option value="atLeast">{{ tl('最小') }}</option>
        <option value="exact">{{ tl('固定') }}</option>
        <option v-if="rule === 'auto'" value="auto">{{ tl('自動') }}</option>
      </select>
    </span>

    <span class="dx-height" :title="tl('游標所在欄（或選取的欄）的寬度')">
      {{ tl('欄寬') }}
      <input v-model="widthText" type="text" inputmode="decimal" :placeholder="info?.colWidth ? '' : '不一'" :aria-label="tl('欄寬（公分）')" @change="applyWidth" @keydown="onEnter(applyWidth)($event)" />
      {{ tl('公分') }}
    </span>

    <label class="dx-check" :title="tl('在每一頁頂端重複此列（Word 的「重複標題列」）')">
      <input type="checkbox" :checked="info?.header" @change="run(setHeaderRow(($event.target as HTMLInputElement).checked))" />
      {{ tl('標題列（跨頁重複）') }}
    </label>
    <label class="dx-check" :title="tl('此列不會被分頁拆開')">
      <input type="checkbox" :checked="info?.cantSplit" @change="run(setCantSplit(($event.target as HTMLInputElement).checked))" />
      {{ tl('列不可跨頁拆開') }}
    </label>

    <span class="dx-sort" role="group" :aria-label="tl('依游標所在欄排序')">
      <button type="button" :disabled="!!sortBlocked" :title="sortBlocked ?? '依游標所在欄由小到大排序（數字依大小、日期依先後、文字依筆畫）'" @click="run(sortTable(false, keepFirst))">{{ tl('排序↑') }}</button>
      <button type="button" :disabled="!!sortBlocked" :title="sortBlocked ?? '依游標所在欄由大到小排序'" @click="run(sortTable(true, keepFirst))">{{ tl('排序↓') }}</button>
      <label class="dx-check" :title="tl('第一列是標題時不參與排序（已設為「標題列」的列一律保留在上方）')">
        <input v-model="keepFirst" type="checkbox" />
        {{ tl('保留第一列') }}
      </label>
    </span>
  </div>
</template>

<style scoped>
.dx-table-panel {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 2px 4px;
  font-size: 12px;
}
button,
select {
  height: 28px;
  min-width: 28px;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
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
  opacity: 0.35;
  cursor: default;
}
button.on {
  background: #e8f0fe;
  color: #1967d2;
}
button:focus-visible,
select:focus-visible,
input:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}
svg {
  width: 16px;
  height: 16px;
  stroke: currentColor;
  stroke-width: 1.5;
  fill: none;
}
.dx-seg {
  display: inline-flex;
}
.dx-height {
  display: inline-flex;
  align-items: center;
  gap: 2px;
}
.dx-height input {
  width: 44px;
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
  gap: 2px;
  cursor: pointer;
  white-space: nowrap;
}
</style>
