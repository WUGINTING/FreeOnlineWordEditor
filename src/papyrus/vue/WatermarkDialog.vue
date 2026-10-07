<script lang="ts">
/** Ids for the dialog's lists, one per dialog on the page. */
let dialogs = 0;
</script>

<script setup lang="ts">
import { tl } from './locale';
import { computed, nextTick, ref, watch } from 'vue';
import type { DocxEditor } from '../editor/core';
import {
  WATERMARK_COLOR, WATERMARK_FONT, WATERMARK_PRESETS, WATERMARK_SCALES, WATERMARK_SIZES, pictureTitle, type Watermark,
} from '../docx/watermark';
import { localNames } from '../docx/eastAsiaFonts';
import { UNSUPPORTED_IMAGE, isImageFile, readImageFile } from '../editor/imageView';

// Word's 自訂浮水印 dialog (設計 › 浮水印 › 自訂浮水印…): 無浮水印, 圖片浮水印 (a picture, its
// 縮放 and 刷淡) or 文字浮水印 (文字, 字型, 大小, 色彩, 半透明, 配置). It opens with the document's
// watermark; 確定 sets it in every header (DocxEditor.setWatermark), or changes nothing when
// nothing was changed.
const props = defineProps<{ editor: DocxEditor | null; open: boolean }>();
const emit = defineEmits<{ (e: 'close'): void }>();

const uid = `dx-wm-${++dialogs}`;
/** Pictures larger than this are refused: every header part holds a copy. */
const MAX_BYTES = 10 * 1024 * 1024;
const FONTS = ['標楷體', '新細明體', '微軟正黑體', '細明體', 'Calibri', 'Arial', 'Times New Roman'];

interface Picture {
  src: string;
  /** Size in px at 100 %. */
  width: number;
  height: number;
  /** The chosen file's name; '' for the document's picture. */
  name: string;
  /** Its name in the document (Word's o:title). */
  title: string;
}
interface Form {
  kind: 'none' | 'picture' | 'text';
  text: string;
  font: string;
  /** 'auto' or pt. */
  size: string;
  color: string;
  semi: boolean;
  layout: 'diagonal' | 'horizontal';
  picture: Picture | null;
  /** 'auto' or %. */
  scale: string;
  washout: boolean;
}

const form = ref<Form>(blank());
let initial = '';
const error = ref('');
const dialog = ref<HTMLElement | null>(null);
const fileInput = ref<HTMLInputElement | null>(null);
let returnFocus: HTMLElement | null = null;
/** Bumped on each opening: a picture measured for an earlier one is dropped. */
let opening = 0;

function blank(): Form {
  return {
    kind: 'none', text: WATERMARK_PRESETS[0], font: WATERMARK_FONT, size: 'auto', color: WATERMARK_COLOR, semi: true,
    layout: 'diagonal', picture: null, scale: 'auto', washout: true,
  };
}

/** The dialog's fonts: 標楷體 and the common ones, the document's East Asian font, the watermark's own. */
const fonts = computed(() => {
  const out = [...FONTS];
  const add = (name: string | null | undefined) => {
    if (!name) return;
    const names = localNames(name).map((n) => n.toLowerCase());
    if (!out.some((f) => names.includes(f.toLowerCase()))) out.push(name);
  };
  add(props.editor?.model?.styles?.defaults.fontEastAsia);
  add(form.value.font);
  return out;
});
/** The size list, with the watermark's own size when Word's list doesn't have it. */
const sizes = computed(() => {
  const own = Number(form.value.size);
  return Number.isFinite(own) && own > 0 && !WATERMARK_SIZES.includes(own) ? [...WATERMARK_SIZES, own].sort((a, b) => a - b) : WATERMARK_SIZES;
});
const scales = computed(() => {
  const own = Number(form.value.scale);
  return Number.isFinite(own) && own > 0 && !WATERMARK_SCALES.includes(own) ? [...WATERMARK_SCALES, own].sort((a, b) => b - a) : WATERMARK_SCALES;
});

/** The form for a watermark (the document's, when it has one). */
function read(w: Watermark | null): Form {
  const f = blank();
  if (w?.kind === 'text') {
    Object.assign(f, {
      kind: 'text', text: w.text, font: w.font || WATERMARK_FONT, size: w.size ? String(w.size) : 'auto', color: w.color,
      semi: w.semitransparent, layout: w.layout,
    });
  } else if (w?.kind === 'picture') {
    Object.assign(f, {
      kind: 'picture', picture: { src: w.src, width: w.width, height: w.height, name: '', title: w.title ?? '' }, scale: w.scale ? String(w.scale) : 'auto', washout: w.washout,
    });
  }
  return f;
}

watch(
  () => props.open,
  (open) => {
    if (!open || !props.editor) return;
    const n = ++opening;
    form.value = read(props.editor.watermark());
    initial = JSON.stringify(form.value);
    error.value = '';
    returnFocus = document.activeElement as HTMLElement | null;
    const picture = form.value.picture;
    // The document's picture is shown at the size it has: its real 縮放 is worked out from the
    // picture's own size once it has loaded.
    if (picture) {
      naturalSize(picture.src).then((size) => {
        if (!size || n !== opening || form.value.picture?.src !== picture.src) return;
        // 自動 stays 自動; a set size shows as the percentage of the picture's own size it is.
        const scale = form.value.scale === 'auto' ? 'auto' : String(Math.round((picture.width / size.width) * 100));
        const measured = { ...form.value, picture: { ...picture, width: size.width, height: size.height }, scale };
        // Unchanged still means unchanged.
        if (JSON.stringify(form.value) === initial) initial = JSON.stringify(measured);
        form.value = measured;
      });
    }
    nextTick(() => dialog.value?.querySelector<HTMLElement>('input[type="radio"]:checked')?.focus());
  },
  { immediate: true },
);

function naturalSize(src: string): Promise<{ width: number; height: number } | null> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve(img.naturalWidth && img.naturalHeight ? { width: img.naturalWidth, height: img.naturalHeight } : null);
    img.onerror = () => resolve(null);
    img.src = src;
  });
}

function close() {
  emit('close');
  nextTick(() => returnFocus?.focus?.());
}

function trapFocus(e: KeyboardEvent) {
  if (e.key !== 'Tab' || !dialog.value) return;
  const items = Array.from(dialog.value.querySelectorAll<HTMLElement>('select, input, button')).filter(
    (el) => !(el as HTMLInputElement).disabled && !el.classList.contains('dx-file-input'),
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

/** 選擇圖片…: a PNG, JPEG or GIF picture (other picture formats are turned into PNG, as 插入圖片 does). */
async function onFile(e: Event) {
  const input = e.target as HTMLInputElement;
  const file = input.files?.[0];
  try {
    input.value = '';
  } catch {
    // some browsers refuse; the same file can then not be chosen twice in a row
  }
  if (!file) return;
  // A file read after the dialog was closed and opened again is dropped.
  const n = opening;
  error.value = '';
  if (file.size > MAX_BYTES) {
    error.value = tl('圖片檔案太大（超過 10 MB），請改用較小的圖片。');
    return;
  }
  const isImage = await isImageFile(file);
  if (n !== opening) return;
  if (!isImage) {
    error.value = tl('請選擇圖片檔（PNG、JPEG 或 GIF）。');
    return;
  }
  let image: { src: string; width: number; height: number };
  try {
    image = await readImageFile(file);
  } catch {
    if (n === opening) error.value = tl(UNSUPPORTED_IMAGE);
    return;
  }
  if (n !== opening) return;
  if (!/^data:image\/(?:png|jpeg|gif)[;,]/i.test(image.src)) {
    error.value = tl('浮水印圖片請使用 PNG、JPEG 或 GIF 格式。');
    return;
  }
  form.value = { ...form.value, kind: 'picture', picture: { ...image, name: file.name, title: pictureTitle(file.name) } };
}

/** The watermark the form sets; a message when it can't be set. */
function watermarkOf(f: Form): Watermark | string | null {
  if (f.kind === 'none') return null;
  if (f.kind === 'picture') {
    if (!f.picture) return tl('請先選擇浮水印要用的圖片。');
    return {
      kind: 'picture', src: f.picture.src, width: f.picture.width, height: f.picture.height,
      scale: f.scale === 'auto' ? null : Number(f.scale), washout: f.washout, title: f.picture.title,
    };
  }
  const text = f.text.replace(/\s+/g, ' ').trim();
  if (!text) return tl('請輸入浮水印的文字。');
  if (!/^#[0-9a-f]{6}$/i.test(f.color)) return tl('請選擇色彩。');
  return {
    kind: 'text', text, font: f.font || WATERMARK_FONT, size: f.size === 'auto' ? null : Number(f.size),
    color: f.color.toLowerCase(), semitransparent: f.semi, layout: f.layout,
  };
}

function apply() {
  const ed = props.editor;
  const w = watermarkOf(form.value);
  if (typeof w === 'string') {
    error.value = w;
    return;
  }
  // 無浮水印 removes any watermark shape there is (even one whose picture is missing, which the
  // dialog can't show); otherwise nothing changed means no change to the document.
  if (ed && w === null) {
    if (ed.hasWatermark()) ed.setWatermark(null);
  } else if (ed && JSON.stringify(form.value) !== initial) {
    ed.setWatermark(w);
  }
  close();
}
</script>

<template>
  <div v-if="open" class="dx-dialog-backdrop" @mousedown.self="close" @keydown.esc="close">
    <div ref="dialog" class="dx-dialog" role="dialog" aria-modal="true" :aria-labelledby="`${uid}-title`" @keydown="trapFocus">
      <h3 :id="`${uid}-title`">{{ tl('浮水印') }}</h3>
      <p class="dx-hint">{{ tl('浮水印會出現在每一頁的文字後面，也會列印出來；它存在各節的頁首中，Word 也能修改或移除。') }}</p>

      <label class="dx-choice"><input v-model="form.kind" type="radio" :name="`${uid}-kind`" value="none" /> {{ tl('無浮水印') }}</label>

      <label class="dx-choice"><input v-model="form.kind" type="radio" :name="`${uid}-kind`" value="picture" /> {{ tl('圖片浮水印') }}</label>
      <fieldset :disabled="form.kind !== 'picture'" :aria-label="tl('圖片浮水印')">
        <div class="dx-line">
          <input :id="`${uid}-file`" ref="fileInput" class="dx-file-input" type="file" accept="image/png,image/jpeg,image/gif" tabindex="-1" @change="onFile" />
          <button type="button" @click="fileInput?.click()">{{ tl('選擇圖片…') }}</button>
          <span class="dx-file-name">{{ form.picture ? form.picture.name || tl('目前的圖片') : tl('尚未選擇圖片') }}</span>
        </div>
        <div class="dx-line">
          <label>{{ tl('縮放') }}
            <select :id="`${uid}-scale`" v-model="form.scale">
              <option value="auto">{{ tl('自動') }}</option>
              <option v-for="s in scales" :key="s" :value="String(s)">{{ s }}%</option>
            </select>
          </label>
          <label><input :id="`${uid}-washout`" v-model="form.washout" type="checkbox" /> {{ tl('刷淡') }}</label>
        </div>
      </fieldset>

      <label class="dx-choice"><input v-model="form.kind" type="radio" :name="`${uid}-kind`" value="text" /> {{ tl('文字浮水印') }}</label>
      <fieldset :disabled="form.kind !== 'text'" :aria-label="tl('文字浮水印')">
        <div class="dx-grid">
          <label :for="`${uid}-text`">{{ tl('文字') }}</label>
          <input :id="`${uid}-text`" v-model="form.text" type="text" :list="`${uid}-presets`" maxlength="255" />
          <datalist :id="`${uid}-presets`">
            <option v-for="p in WATERMARK_PRESETS" :key="p" :value="p" />
          </datalist>
          <label :for="`${uid}-font`">{{ tl('字型') }}</label>
          <select :id="`${uid}-font`" v-model="form.font">
            <option v-for="f in fonts" :key="f" :value="f">{{ f }}</option>
          </select>
          <label :for="`${uid}-size`">{{ tl('大小') }}</label>
          <select :id="`${uid}-size`" v-model="form.size">
            <option value="auto">{{ tl('自動') }}</option>
            <option v-for="s in sizes" :key="s" :value="String(s)">{{ s }}</option>
          </select>
          <label :for="`${uid}-color`">{{ tl('色彩') }}</label>
          <span class="dx-line">
            <input :id="`${uid}-color`" v-model="form.color" type="color" />
            <label><input :id="`${uid}-semi`" v-model="form.semi" type="checkbox" /> {{ tl('半透明') }}</label>
          </span>
          <span>{{ tl('配置') }}</span>
          <span class="dx-line" role="radiogroup" :aria-label="tl('配置')">
            <label><input v-model="form.layout" type="radio" :name="`${uid}-layout`" value="diagonal" /> {{ tl('斜向') }}</label>
            <label><input v-model="form.layout" type="radio" :name="`${uid}-layout`" value="horizontal" /> {{ tl('水平') }}</label>
          </span>
        </div>
      </fieldset>

      <p v-if="error" class="dx-error" role="alert">{{ error }}</p>
      <div class="dx-actions">
        <button type="button" @click="close">{{ tl('取消') }}</button>
        <button type="button" class="dx-primary" @click="apply">{{ tl('確定') }}</button>
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
  width: min(460px, calc(100vw - 32px));
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
.dx-choice {
  display: block;
  margin: 6px 0 4px;
  font-weight: 600;
}
fieldset {
  border: 1px solid #dadce0;
  border-radius: 6px;
  margin: 0 0 8px;
  padding: 8px 12px;
}
fieldset:disabled {
  color: #9aa0a6;
}
.dx-line {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 6px 12px;
  margin: 2px 0;
}
.dx-grid {
  display: grid;
  grid-template-columns: auto minmax(0, 1fr);
  align-items: center;
  gap: 6px 12px;
}
.dx-grid input[type='text'],
.dx-grid select {
  width: 100%;
  box-sizing: border-box;
}
.dx-file-input {
  display: none;
}
.dx-file-name {
  font-size: 13px;
  color: #5f6368;
  overflow-wrap: anywhere;
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
.dx-actions button,
.dx-line > button {
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
