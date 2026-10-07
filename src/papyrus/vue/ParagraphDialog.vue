<script setup lang="ts">
import { nextTick, ref, watch } from 'vue';
import type { DocxEditor } from '../editor/core';
import { setParagraphFormat } from '../editor/commands';

// Word's 段落 dialog for the selected paragraphs: indents (公分 or 字元), first-line or hanging
// indent, space before / after, line spacing. It shows the paragraph's own settings; an empty
// box is 「依樣式」 (the style decides). Only what is changed is written.
const props = defineProps<{ editor: DocxEditor | null; open: boolean }>();
const emit = defineEmits<{ (e: 'close'): void }>();

const TWIPS_PER_CM = 1440 / 2.54;
const round2 = (n: number) => Math.round(n * 100) / 100;

type LineKind = '' | 'single' | '1.5' | 'double' | 'multiple' | 'atLeast' | 'exact';
interface Form {
  unit: 'cm' | 'char';
  left: string;
  right: string;
  special: 'none' | 'first' | 'hanging';
  by: string;
  before: string;
  after: string;
  line: LineKind;
  lineValue: string;
}

const form = ref<Form>(blank());
let initial: Form = blank();
const error = ref('');
const dialog = ref<HTMLElement | null>(null);
let returnFocus: HTMLElement | null = null;

function blank(): Form {
  return { unit: 'cm', left: '', right: '', special: 'none', by: '', before: '', after: '', line: '', lineValue: '' };
}

/** The form for the cursor's paragraph (its own attributes). */
function read(a: Record<string, any>): Form {
  const chars = a.indLeftChars != null || a.indRightChars != null || a.indFirstChars != null;
  const len = (twips: number | null, hundredths: number | null) =>
    chars ? (hundredths == null ? '' : String(round2(hundredths / 100))) : twips == null ? '' : String(round2(twips / TWIPS_PER_CM));
  const first = chars ? a.indFirstChars : a.indFirst;
  const f: Form = {
    unit: chars ? 'char' : 'cm',
    left: len(a.indLeft, a.indLeftChars),
    right: len(a.indRight, a.indRightChars),
    special: first == null || first === 0 ? 'none' : first > 0 ? 'first' : 'hanging',
    by: first == null || first === 0 ? '' : len(Math.abs(a.indFirst ?? 0), a.indFirstChars == null ? null : Math.abs(a.indFirstChars)),
    before: a.spaceBefore == null ? '' : String(round2(a.spaceBefore / 20)),
    after: a.spaceAfter == null ? '' : String(round2(a.spaceAfter / 20)),
    line: '',
    lineValue: '',
  };
  if (a.line != null) {
    if (a.lineRule === 'exact' || a.lineRule === 'atLeast') {
      f.line = a.lineRule;
      f.lineValue = String(round2(a.line / 20));
    } else {
      f.line = a.line === 240 ? 'single' : a.line === 360 ? '1.5' : a.line === 480 ? 'double' : 'multiple';
      f.lineValue = String(round2(a.line / 240));
    }
  }
  return f;
}

watch(
  () => props.open,
  (open) => {
    const view = props.editor?.activeView;
    if (!open || !view) return;
    const para = view.state.selection.$from.parent;
    form.value = read(para.type.name === 'paragraph' ? para.attrs : {});
    initial = { ...form.value };
    error.value = '';
    returnFocus = document.activeElement as HTMLElement | null;
    nextTick(() => dialog.value?.querySelector<HTMLElement>('select, input, button')?.focus());
  },
  { immediate: true },
);

function close() {
  emit('close');
  nextTick(() => returnFocus?.focus?.());
}

function trapFocus(e: KeyboardEvent) {
  if (e.key !== 'Tab' || !dialog.value) return;
  const items = Array.from(dialog.value.querySelectorAll<HTMLElement>('select, input, button')).filter((el) => !el.hasAttribute('disabled'));
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

/** '' → null (依樣式); a number → the number; anything else → NaN (refused). */
const num = (s: string): number | null => (s.trim() === '' ? null : Number(s));

function apply() {
  const f = form.value;
  const values = [f.left, f.right, f.by, f.before, f.after, f.lineValue].map(num);
  if (values.some((v) => v != null && (!Number.isFinite(v) || v < 0))) {
    error.value = '請輸入 0 以上的數字，或留空（依樣式）。';
    return;
  }
  const patch: Record<string, number | string | null> = {};
  const unitChanged = f.unit !== initial.unit;
  // Indents: the unit decides which attributes hold them; the other kind is cleared.
  const setLength = (key: 'Left' | 'Right', value: string, was: string) => {
    if (!unitChanged && value === was) return;
    const n = num(value);
    if (f.unit === 'char') {
      patch[`ind${key}Chars`] = n == null ? null : Math.round(n * 100);
      patch[`ind${key}`] = null;
    } else {
      patch[`ind${key}`] = n == null ? null : Math.round(n * TWIPS_PER_CM);
      patch[`ind${key}Chars`] = null;
    }
  };
  setLength('Left', f.left, initial.left);
  setLength('Right', f.right, initial.right);
  if (unitChanged || f.special !== initial.special || f.by !== initial.by) {
    const n = f.special === 'none' ? null : num(f.by) ?? 0;
    const signed = n == null ? null : f.special === 'hanging' ? -n : n;
    if (f.unit === 'char') {
      patch.indFirstChars = signed == null ? null : Math.round(signed * 100);
      patch.indFirst = null;
    } else {
      patch.indFirst = signed == null ? null : Math.round(signed * TWIPS_PER_CM);
      patch.indFirstChars = null;
    }
  }
  if (f.before !== initial.before) patch.spaceBefore = num(f.before) == null ? null : Math.round(num(f.before)! * 20);
  if (f.after !== initial.after) patch.spaceAfter = num(f.after) == null ? null : Math.round(num(f.after)! * 20);
  if (f.line !== initial.line || f.lineValue !== initial.lineValue) {
    const v = num(f.lineValue);
    const lines: Record<string, number> = { single: 240, '1.5': 360, double: 480 };
    if (!f.line) {
      patch.line = null;
      patch.lineRule = null;
    } else if (lines[f.line]) {
      patch.line = lines[f.line];
      patch.lineRule = 'auto';
    } else if (f.line === 'multiple') {
      patch.line = Math.round((v ?? 1) * 240);
      patch.lineRule = 'auto';
    } else {
      patch.line = Math.round((v ?? 12) * 20);
      patch.lineRule = f.line;
    }
  }
  if (Object.keys(patch).length) props.editor?.run(setParagraphFormat(patch));
  close();
}

function onLineKind() {
  const f = form.value;
  if (f.line === 'multiple' && !f.lineValue) f.lineValue = '1.15';
  if ((f.line === 'exact' || f.line === 'atLeast') && !f.lineValue) f.lineValue = '12';
}
</script>

<template>
  <div v-if="open" class="dx-dialog-backdrop" @mousedown.self="close" @keydown.esc="!($event.isComposing || $event.keyCode === 229) && close()">
    <div ref="dialog" class="dx-dialog" role="dialog" aria-modal="true" aria-labelledby="dx-pd-title" @keydown="trapFocus">
      <h3 id="dx-pd-title">段落</h3>
      <p class="dx-hint">空白表示依段落樣式。只會寫入你改過的項目。</p>

      <fieldset>
        <legend>縮排</legend>
        <div class="dx-row">
          單位
          <label><input v-model="form.unit" type="radio" name="dx-pd-unit" value="cm" /> 公分</label>
          <label><input v-model="form.unit" type="radio" name="dx-pd-unit" value="char" /> 字元</label>
        </div>
        <div class="dx-grid">
          <label>左 <input v-model="form.left" type="text" inputmode="decimal" placeholder="依樣式" /></label>
          <label>右 <input v-model="form.right" type="text" inputmode="decimal" placeholder="依樣式" /></label>
          <label>
            指定方式
            <select v-model="form.special">
              <option value="none">（無）</option>
              <option value="first">第一行</option>
              <option value="hanging">凸排</option>
            </select>
          </label>
          <label>位移 <input v-model="form.by" type="text" inputmode="decimal" :disabled="form.special === 'none'" placeholder="0" /></label>
        </div>
      </fieldset>

      <fieldset>
        <legend>間距</legend>
        <div class="dx-grid">
          <label>與前段距離（點）<input v-model="form.before" type="text" inputmode="decimal" placeholder="依樣式" /></label>
          <label>與後段距離（點）<input v-model="form.after" type="text" inputmode="decimal" placeholder="依樣式" /></label>
          <label>
            行距
            <select v-model="form.line" @change="onLineKind">
              <option value="">依樣式</option>
              <option value="single">單行間距</option>
              <option value="1.5">1.5 倍行高</option>
              <option value="double">2 倍行高</option>
              <option value="multiple">多行（倍數）</option>
              <option value="atLeast">最小行高（點）</option>
              <option value="exact">固定行高（點）</option>
            </select>
          </label>
          <label>
            {{ form.line === 'atLeast' || form.line === 'exact' ? '點數' : '倍數' }}
            <input v-model="form.lineValue" type="text" inputmode="decimal" :disabled="!['multiple', 'atLeast', 'exact'].includes(form.line)" />
          </label>
        </div>
      </fieldset>

      <p v-if="error" class="dx-error" role="alert">{{ error }}</p>
      <div class="dx-actions">
        <button type="button" @click="close">取消</button>
        <button type="button" class="dx-primary" @click="apply">套用</button>
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
  width: min(480px, calc(100vw - 32px));
  max-height: calc(100vh - 32px);
  overflow: auto;
  background: #fff;
  border-radius: 8px;
  padding: 16px 20px;
  box-shadow: 0 4px 24px rgba(0, 0, 0, 0.25);
  color: #202124;
}
h3 {
  margin: 0 0 4px;
  font-size: 17px;
}
.dx-hint {
  margin: 0 0 10px;
  font-size: 12px;
  color: #5f6368;
}
fieldset {
  border: 1px solid #dadce0;
  border-radius: 6px;
  margin: 0 0 12px;
  padding: 8px 12px;
}
legend {
  padding: 0 4px;
  font-size: 13px;
  color: #5f6368;
}
.dx-row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px 12px;
  margin: 4px 0;
}
.dx-grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 6px 12px;
}
.dx-grid label {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 6px;
}
input[type='text'] {
  width: 5.5em;
}
.dx-error {
  color: #b3261e;
  font-size: 13px;
}
.dx-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
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
</style>
