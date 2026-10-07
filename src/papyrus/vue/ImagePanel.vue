<script setup lang="ts">
import { tl } from './locale';
// Tools for the selected picture: size (px or cm, proportions locked by default),
// turn 90° to the right, replace, delete, alt text.
import { computed, onBeforeUnmount, ref, watch } from 'vue';
import { deleteSelection } from 'prosemirror-commands';
import { NodeSelection } from 'prosemirror-state';
import type { DocxEditor, EditorSnapshot } from '../editor/core';
import { selectedImage, setImageAlt } from '../editor/commands';
import { UNSUPPORTED_IMAGE, readImageFile, replaceImage, rotateImage, rotateImageRight, setImageAttrs } from '../editor/imageView';
import { onEnter } from './keys';
import { readDrawing } from '../docx/props';

const props = defineProps<{ editor: DocxEditor | null; snapshot: EditorSnapshot | null }>();

const image = computed(() => {
  void props.snapshot; // re-evaluated on every editor update
  const state = props.editor?.activeView?.state;
  return state ? selectedImage(state) : null;
});
const run = (cmd: Parameters<DocxEditor['run']>[0]) => props.editor?.run(cmd);

const PX_PER_CM = 96 / 2.54;
const unit = ref<'px' | 'cm'>('cm');
const lock = ref(true);
const widthText = ref('');
const heightText = ref('');
const altText = ref('');

const show = (px: number | null) => (px ? String(unit.value === 'cm' ? +(px / PX_PER_CM).toFixed(2) : Math.round(px)) : '');
function sync() {
  const a = image.value?.attrs;
  widthText.value = show(a?.width ?? null);
  heightText.value = show(a?.height ?? null);
  altText.value = a?.alt ?? '';
}
watch(() => [image.value?.attrs.width, image.value?.attrs.height, image.value?.attrs.alt, image.value?.attrs.src, unit.value], sync, { immediate: true });

/** Input in the current unit -> px (cm keeps full precision, so 5 cm is exactly 5 cm in Word). */
function toPx(text: string): number | null {
  const n = Number(text.trim());
  if (!Number.isFinite(n) || n <= 0) return null;
  return unit.value === 'cm' ? n * PX_PER_CM : Math.round(n);
}

function applySize(which: 'width' | 'height') {
  const a = image.value?.attrs;
  const text = which === 'width' ? widthText.value : heightText.value;
  // The box still shows the (rounded) current size: nothing to change.
  if (a && text.trim() === show(a[which] ?? null)) return;
  const value = toPx(text);
  if (!a || value == null) {
    sync();
    return;
  }
  const w0 = a.width || value;
  const h0 = a.height || value;
  let width = which === 'width' ? value : w0;
  let height = which === 'height' ? value : h0;
  if (lock.value && a.width && a.height) {
    if (which === 'width') height = unit.value === 'cm' ? (value * h0) / w0 : Math.round((value * h0) / w0);
    else width = unit.value === 'cm' ? (value * w0) / h0 : Math.round((value * w0) / h0);
  }
  run(setImageAttrs({ width, height }));
}

// Alt text (GOV-ISSUE-011): applied to the picture as it is typed, so a save or leaving the
// document right after typing keeps it (the box's change event only comes when it loses
// focus, and clicking a disabled button or the toolbar background never takes it). One
// typing session (until the box is left or Enter) is one undo step.
let altSession: { time: number; first: boolean; pos: number; view: unknown } | null = null;
function applyAlt(end = false) {
  const view = props.editor?.activeView;
  const sel = view?.state.selection;
  const selected = !!view && sel instanceof NodeSelection && sel.node.type.name === 'image';
  // The picture being described: the selected one, else (it was just deselected) where it was.
  const pos = selected ? sel!.from : altSession?.view === view ? altSession?.pos : undefined;
  if (view && pos != null && view.state.doc.nodeAt(pos)?.type.name === 'image') {
    if (!altSession || altSession.pos !== pos || altSession.view !== view) altSession = { time: Date.now(), first: true, pos, view };
    const before = view.state.doc;
    // Not through editor.run(): that would take the focus from the box.
    setImageAlt(pos, altText.value, altSession)(view.state, view.dispatch);
    if (view.state.doc !== before) altSession.first = false;
  }
  if (end) altSession = null;
}
function onAltInput(e: Event) {
  // Wait for the composed text of an input method.
  if ((e as InputEvent).isComposing) return;
  applyAlt();
}
// The panel goes away while the box still has the focus (the picture was deselected): keep what was typed.
onBeforeUnmount(() => {
  if (altSession) applyAlt(true);
});

const fileInput = ref<HTMLInputElement | null>(null);
const message = ref('');

/**
 * How Word shows the selected picture from the file: turned or flipped there, and cropped.
 * A picture Word turned or flipped is not turned here: drawn again turned, Word would turn it
 * once more (its own turn stays in the file).
 */
const inWord = computed(() => {
  const xml = image.value?.attrs.xml as string | null | undefined;
  if (!xml || image.value?.attrs.src !== image.value?.attrs.origSrc) return { turned: false, crop: null };
  try {
    const d = readDrawing(xml);
    return { turned: d.turned, crop: d.crop };
  } catch {
    return { turned: false, crop: null };
  }
});
const TURNED_IN_WORD = '這張圖片已在 Word 中旋轉或翻轉，請用 Word 旋轉。';

/** 向右旋轉 90° (persona-300 A-9): the picture is drawn again turned, its width and height swap. */
async function rotateRight() {
  const src = image.value?.attrs.src as string | undefined;
  message.value = '';
  if (!src) return;
  if (inWord.value.turned) {
    message.value = tl(TURNED_IN_WORD);
    return;
  }
  // Cropped in Word: what Word shows is what is turned.
  const turned = await rotateImageRight(src, inWord.value.crop);
  // Still the same picture selected (the drawing takes a moment).
  if (image.value?.attrs.src !== src) return;
  if (!turned || !run(rotateImage(turned))) message.value = tl('這張圖片無法在網頁上旋轉，請用 Word 旋轉。');
}
async function onReplace(e: Event) {
  const input = e.target as HTMLInputElement;
  const file = input.files?.[0];
  input.value = '';
  message.value = '';
  if (!file) return;
  try {
    // Formats Word may not show (WebP, HEIC, SVG ...) come back converted to PNG.
    const { src, width, height } = await readImageFile(file);
    run(replaceImage(src, width, height));
  } catch {
    message.value = tl(UNSUPPORTED_IMAGE);
  }
}
</script>

<template>
  <div class="dx-group dx-image-panel" role="group" :aria-label="tl('圖片')">
    <span class="dx-size">
      {{ tl('寬') }}
      <input v-model="widthText" type="text" inputmode="decimal" :aria-label="tl('圖片寬度')" @change="applySize('width')" @keydown="onEnter(() => applySize('width'))($event)" />
      {{ tl('高') }}
      <input v-model="heightText" type="text" inputmode="decimal" :aria-label="tl('圖片高度')" @change="applySize('height')" @keydown="onEnter(() => applySize('height'))($event)" />
      <select v-model="unit" :aria-label="tl('尺寸單位')">
        <option value="cm">{{ tl('公分') }}</option>
        <option value="px">{{ tl('像素') }}</option>
      </select>
    </span>
    <label class="dx-check" :title="tl('調整寬度或高度時維持長寬比')">
      <input v-model="lock" type="checkbox" />
      {{ tl('鎖定比例') }}
    </label>
    <button
      type="button"
      :title="inWord.turned ? tl(TURNED_IN_WORD) : tl('把圖片向右轉 90 度（順時針）')"
      :aria-disabled="inWord.turned ? 'true' : undefined"
      @click="rotateRight"
    >{{ tl('向右旋轉 90°') }}</button>
    <button type="button" :title="tl('換成另一張圖片（保留目前寬度與版面設定）')" @click="fileInput?.click()">{{ tl('替換圖片') }}</button>
    <input ref="fileInput" type="file" accept="image/*" hidden @change="onReplace" />
    <span v-if="message" class="dx-msg" role="alert">{{ message }}</span>
    <button type="button" :title="tl('刪除圖片')" @click="run(deleteSelection)">{{ tl('刪除') }}</button>
    <span class="dx-alt">
      {{ tl('替代文字') }}
      <input
        v-model="altText"
        type="text"
        :aria-label="tl('替代文字')"
        :placeholder="tl('描述這張圖片')"
        @input="onAltInput"
        @change="applyAlt(true)"
        @blur="applyAlt(true)"
        @keydown="onEnter(() => applyAlt(true))($event)"
      />
    </span>
  </div>
</template>

<style scoped>
.dx-image-panel {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 2px 6px;
  font-size: 12px;
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
button:hover {
  background: #f1f3f4;
}
button:focus-visible,
select:focus-visible,
input:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}
.dx-size,
.dx-alt,
.dx-check {
  display: inline-flex;
  align-items: center;
  gap: 3px;
  white-space: nowrap;
}
input[type='text'] {
  height: 24px;
  box-sizing: border-box;
  border: 1px solid #dadce0;
  border-radius: 4px;
  padding: 0 4px;
  font: inherit;
}
.dx-size input {
  width: 52px;
}
.dx-msg {
  color: #c5221f;
}
.dx-alt input {
  width: 160px;
}
</style>
