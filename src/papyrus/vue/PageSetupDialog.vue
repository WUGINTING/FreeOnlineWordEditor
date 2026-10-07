<script setup lang="ts">
import { tl } from './locale';
import { computed, nextTick, ref, watch } from 'vue';
import type { ColumnsChange, DocxEditor, PageSetupScope } from '../editor/core';
import type { PageSetup } from '../docx/model';
import type { PageNumbering, SectionColumns, SectionStart } from '../docx/sections';

const props = defineProps<{ editor: DocxEditor | null; open: boolean }>();
const emit = defineEmits<{ (e: 'close'): void }>();

// Paper sizes in twips (1 cm = 567 twips), portrait.
const PAPERS = [
  { id: 'A4', label: 'A4（21 × 29.7 公分）', w: 11906, h: 16838 },
  { id: 'A3', label: 'A3（29.7 × 42 公分）', w: 16838, h: 23811 },
  { id: 'A5', label: 'A5（14.8 × 21 公分）', w: 8391, h: 11906 },
  { id: 'B4', label: 'B4（25.7 × 36.4 公分）', w: 14570, h: 20636 },
  { id: 'B5', label: 'B5（18.2 × 25.7 公分）', w: 10318, h: 14570 },
  { id: 'Letter', label: 'Letter（8.5 × 11 英吋）', w: 12240, h: 15840 },
  { id: 'Legal', label: 'Legal（8.5 × 14 英吋）', w: 12240, h: 20160 },
];
// Page number formats (w:pgNumType/@w:fmt), as Word's 頁碼格式 lists them.
const NUMBER_FORMATS = [
  { id: 'decimal', label: '1, 2, 3 …' },
  { id: 'numberInDash', label: '- 1 -, - 2 -, - 3 - …' },
  { id: 'lowerRoman', label: 'i, ii, iii …（小寫羅馬數字）' },
  { id: 'upperRoman', label: 'I, II, III …（大寫羅馬數字）' },
  { id: 'lowerLetter', label: 'a, b, c …' },
  { id: 'upperLetter', label: 'A, B, C …' },
  { id: 'taiwaneseCounting', label: '一, 二, 三 … 十一（繁）' },
  { id: 'chineseCounting', label: '一, 二, 三 … 十一（簡）' },
  { id: 'taiwaneseCountingThousand', label: '一, 二, 三 … 一○, 一一（逐位）' },
  { id: 'ideographLegalTraditional', label: '壹, 貳, 參 …' },
  { id: 'ideographTraditional', label: '甲, 乙, 丙 …（第 11 頁起為數字）' },
];
const MAX_START = 32767;
const TWIPS_PER_CM = 1440 / 2.54;
const cm = (twips: number) => Math.round((twips / TWIPS_PER_CM) * 100) / 100;
const twips = (value: number) => Math.round(value * TWIPS_PER_CM);

const paper = ref('A4');
const landscape = ref(false);
const width = ref(21);
const height = ref(29.7);
const margins = ref({ top: 2.54, bottom: 2.54, left: 3.17, right: 3.17, header: 1.5, footer: 1.75 });
const scope = ref<PageSetupScope>('section');
const sectionInfo = ref({ index: 0, count: 1 });
const error = ref('');
/** The page as opened: a field left as shown keeps its exact value (centimetres are rounded). */
let original: PageSetup | null = null;

// 頁碼: the cursor's section's page number format and start, and "different first page".
const numberFormat = ref('decimal');
const restart = ref(false);
const startAt = ref(1);
const titlePage = ref(false);
/** The first page's number when the section continues from the previous one. */
const continued = ref(1);
/** A format the document has that the list doesn't offer (kept unless changed). */
const otherFormat = ref<string | null>(null);
let originalNumbering: { format: string; start: number | null; titlePage: boolean } | null = null;
const dialog = ref<HTMLElement | null>(null);
let returnFocus: HTMLElement | null = null;

// 分欄 (w:cols): Word shows and prints the columns; the editor keeps one column (GOV-FINDING-014).
const COLUMN_COUNTS = [1, 2, 3];
/** Narrowest column the dialog allows (cm). */
const MIN_COLUMN_CM = 1;
const ONE_COLUMN_NOTE = '網頁上也分欄顯示與列印；段落與表格整段移到下一欄（Word 會在段落中間換欄），分欄位置可能和 Word 略有不同。';
const columnCount = ref(1);
/** Space between columns, cm. */
const columnSpace = ref(1.27);
const columnSeparator = ref(false);
let originalColumns: SectionColumns | null = null;
/** A column count from the file that the list doesn't offer (4 or more). */
const otherCount = ref<number | null>(null);
// Where the section starts (w:type): a title over two columns needs the columns' section to
// continue on the title's page (Word's 分節符號（接續本頁）).
const START_LABELS: Record<SectionStart, string> = {
  nextPage: '新頁',
  continuous: '接續本頁（不換頁）',
  evenPage: '自偶數頁起（文件原有）',
  oddPage: '自奇數頁起（文件原有）',
  nextColumn: '下一欄（文件原有）',
};
const sectionStart = ref<SectionStart>('nextPage');
let originalStart: SectionStart = 'nextPage';
const startOptions = computed(() =>
  (Object.keys(START_LABELS) as SectionStart[]).filter((k) => k === 'nextPage' || k === 'continuous' || k === originalStart),
);

// Fill the form from the cursor's section each time the dialog opens.
watch(
  () => props.open,
  (open) => {
    const ed = props.editor;
    if (!open || !ed?.view) return;
    const s = ed.cursorSection();
    const p = s.page;
    const short = Math.min(p.width, p.height);
    const long = Math.max(p.width, p.height);
    landscape.value = p.width > p.height;
    paper.value = PAPERS.find((x) => Math.abs(x.w - short) < 30 && Math.abs(x.h - long) < 30)?.id ?? 'custom';
    width.value = cm(p.width);
    height.value = cm(p.height);
    margins.value = {
      top: cm(p.marginTop), bottom: cm(p.marginBottom), left: cm(p.marginLeft), right: cm(p.marginRight),
      header: cm(p.header), footer: cm(p.footer),
    };
    sectionInfo.value = { index: s.index, count: ed.sections().length };
    const fmt = s.pageNumberFormat ?? 'decimal';
    numberFormat.value = fmt;
    otherFormat.value = NUMBER_FORMATS.some((f) => f.id === fmt) ? null : fmt;
    restart.value = s.pageNumberStart != null;
    startAt.value = s.pageNumberStart ?? 1;
    titlePage.value = s.titlePage;
    continued.value = ed.continuedPageNumber(s.index);
    originalNumbering = { format: fmt, start: s.pageNumberStart, titlePage: s.titlePage };
    const c = s.columns;
    columnCount.value = c.count;
    otherCount.value = COLUMN_COUNTS.includes(c.count) ? null : c.count;
    columnSpace.value = cm(c.space);
    columnSeparator.value = c.separator;
    originalColumns = { ...c };
    sectionStart.value = originalStart = s.start;
    scope.value = 'section';
    error.value = '';
    original = { ...p };
    // Keyboard users land in the dialog, and go back where they were when it closes.
    returnFocus = document.activeElement as HTMLElement | null;
    nextTick(() => dialog.value?.querySelector<HTMLElement>('select, input, button')?.focus());
  },
  { immediate: true },
);

function close() {
  emit('close');
  nextTick(() => returnFocus?.focus?.());
}

/** Tab / Shift+Tab stay inside the dialog. */
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

/** The value in twips: unchanged fields keep the document's own value exactly. */
function keep(value: number, ...candidates: number[]): number {
  return candidates.find((t) => cm(t) === value) ?? twips(value);
}

function onPaper() {
  const p = PAPERS.find((x) => x.id === paper.value);
  if (!p) return;
  width.value = cm(landscape.value ? p.h : p.w);
  height.value = cm(landscape.value ? p.w : p.h);
}
function onOrientation(value: boolean) {
  if (value === landscape.value) return;
  landscape.value = value;
  [width.value, height.value] = [height.value, width.value];
}

const multi = computed(() => sectionInfo.value.count > 1);

/** The columns as the dialog would set them, or null when they are as the section has them. */
function columnsChange(): ColumnsChange | null {
  const o = originalColumns;
  const space = o ? keep(columnSpace.value, o.space) : twips(columnSpace.value);
  const count = columnCount.value;
  const separator = count > 1 && columnSeparator.value;
  if (o && count === o.count && space === o.space && separator === o.separator) return null;
  return { count, space, separator };
}

/** The file sets each column's width and nothing about the columns was changed: kept as it is. */
const unequalKept = computed(() => {
  const o = originalColumns;
  return !!o && !o.equalWidth && columnCount.value === o.count && keep(columnSpace.value, o.space) === o.space;
});

/** What the section's first page shows with these settings. */
const preview = computed(() => {
  const n = restart.value ? startAt.value : continued.value;
  const ed = props.editor;
  return ed && Number.isInteger(n) && n >= 0 ? ed.pageNumberText(n, numberFormat.value) || '（空白）' : '';
});

/** Only the column values that differ from the section's. */
function changedColumns(c: ColumnsChange): ColumnsChange {
  const o = originalColumns;
  if (!o) return c;
  const out: ColumnsChange = {};
  if (c.count !== o.count) out.count = c.count;
  if (c.space !== o.space) out.space = c.space;
  if (c.separator !== o.separator) out.separator = c.separator;
  return out;
}

function apply() {
  const ed = props.editor;
  if (!ed) return;
  const m = margins.value;
  const values = [width.value, height.value, m.top, m.bottom, m.left, m.right, m.header, m.footer];
  if (values.some((v) => !Number.isFinite(v) || v < 0)) {
    error.value = '請輸入 0 以上的數字。';
    return;
  }
  if (width.value < 5 || height.value < 5 || width.value > 150 || height.value > 150) {
    error.value = '紙張大小需介於 5 到 150 公分。';
    return;
  }
  if (m.left + m.right >= width.value - 1 || m.top + m.bottom >= height.value - 1) {
    error.value = '邊界太大，頁面上沒有可以寫字的地方了。';
    return;
  }
  if (restart.value && (!Number.isInteger(startAt.value) || startAt.value < 0 || startAt.value > MAX_START)) {
    error.value = `起始頁碼請輸入 0 到 ${MAX_START} 的整數。`;
    return;
  }
  const n = columnCount.value;
  if (n > 1) {
    if (!Number.isFinite(columnSpace.value) || columnSpace.value < 0) {
      error.value = '欄間距請輸入 0 以上的數字。';
      return;
    }
    const text = width.value - m.left - m.right;
    if ((text - (n - 1) * columnSpace.value) / n < MIN_COLUMN_CM) {
      error.value = `欄間距太大：${n} 欄時每欄至少要 ${MIN_COLUMN_CM} 公分寬，請縮小間距或邊界。`;
      return;
    }
  }
  const o = original;
  const page: PageSetup = o
    ? {
        // Width/height may have been swapped by the orientation: either original side counts.
        width: keep(width.value, o.width, o.height),
        height: keep(height.value, o.height, o.width),
        marginTop: keep(m.top, o.marginTop),
        marginBottom: keep(m.bottom, o.marginBottom),
        marginLeft: keep(m.left, o.marginLeft),
        marginRight: keep(m.right, o.marginRight),
        header: keep(m.header, o.header),
        footer: keep(m.footer, o.footer),
      }
    : {
        width: twips(width.value), height: twips(height.value),
        marginTop: twips(m.top), marginBottom: twips(m.bottom), marginLeft: twips(m.left), marginRight: twips(m.right),
        header: twips(m.header), footer: twips(m.footer),
      };
  // Only what was changed is written (a section already set that way keeps its w:sectPr).
  const pageSame = !!o && (Object.keys(page) as (keyof PageSetup)[]).every((k) => page[k] === o[k]);
  const on = originalNumbering;
  const numbering: Partial<PageNumbering> = {};
  if (!on || numberFormat.value !== on.format) numbering.format = numberFormat.value;
  const start = restart.value ? startAt.value : null;
  if (!on || start !== on.start) numbering.start = start;
  const numberingChanged = Object.keys(numbering).length > 0;
  const titleChanged = !on || titlePage.value !== on.titlePage;
  const columns = columnsChange();
  const startChanged = sectionStart.value !== originalStart;
  // Nothing changed for a single section: nothing to do (no edit, no "unsaved").
  if (scope.value === 'section' && pageSame && !numberingChanged && !titleChanged && !columns && !startChanged) {
    close();
    return;
  }
  ed.setSectionSetup(
    {
      // 整份文件: every section gets this page setup, as before.
      page: scope.value === 'all' || !pageSame ? page : undefined,
      numbering: numberingChanged ? numbering : undefined,
      titlePage: titleChanged ? titlePage.value : undefined,
      // For the cursor's section only what changed; 整份文件 gives every section these columns.
      columns: !columns ? undefined : scope.value === 'all' ? columns : changedColumns(columns),
      start: startChanged ? sectionStart.value : undefined,
    },
    scope.value,
  );
  close();
}
</script>

<template>
  <div v-if="open" class="dx-dialog-backdrop" @mousedown.self="close" @keydown.esc="!($event.isComposing || $event.keyCode === 229) && close()">
    <div ref="dialog" class="dx-dialog" role="dialog" aria-modal="true" aria-labelledby="dx-ps-title" @keydown="trapFocus">
      <h3 id="dx-ps-title">{{ tl('版面設定') }}</h3>

      <fieldset>
        <legend>{{ tl('紙張') }}</legend>
        <label class="dx-row">
          {{ tl('大小') }}
          <select v-model="paper" @change="onPaper">
            <option v-for="p in PAPERS" :key="p.id" :value="p.id">{{ p.label }}</option>
            <option value="custom">{{ tl('自訂') }}</option>
          </select>
        </label>
        <div class="dx-row">
          <label>{{ tl('寬') }} <input v-model.number="width" type="number" step="0.1" min="5" @input="paper = 'custom'" /> {{ tl('公分') }}</label>
          <label>{{ tl('高') }} <input v-model.number="height" type="number" step="0.1" min="5" @input="paper = 'custom'" /> {{ tl('公分') }}</label>
        </div>
        <div class="dx-row" role="radiogroup" :aria-label="tl('方向')">
          {{ tl('方向') }}
          <label><input type="radio" name="dx-ps-orient" :checked="!landscape" @change="onOrientation(false)" /> {{ tl('直向') }}</label>
          <label><input type="radio" name="dx-ps-orient" :checked="landscape" @change="onOrientation(true)" /> {{ tl('橫向') }}</label>
        </div>
      </fieldset>

      <fieldset>
        <legend>{{ tl('邊界（公分）') }}</legend>
        <div class="dx-grid">
          <label>{{ tl('上') }} <input v-model.number="margins.top" type="number" step="0.1" min="0" /></label>
          <label>{{ tl('下') }} <input v-model.number="margins.bottom" type="number" step="0.1" min="0" /></label>
          <label>{{ tl('左') }} <input v-model.number="margins.left" type="number" step="0.1" min="0" /></label>
          <label>{{ tl('右') }} <input v-model.number="margins.right" type="number" step="0.1" min="0" /></label>
          <label>{{ tl('頁首距頁緣') }} <input v-model.number="margins.header" type="number" step="0.1" min="0" /></label>
          <label>{{ tl('頁尾距頁緣') }} <input v-model.number="margins.footer" type="number" step="0.1" min="0" /></label>
        </div>
      </fieldset>

      <fieldset>
        <legend>{{ tl('頁碼') }}</legend>
        <label class="dx-row">
          {{ tl('頁碼格式') }}
          <select v-model="numberFormat" :aria-label="tl('頁碼格式')">
            <option v-for="f in NUMBER_FORMATS" :key="f.id" :value="f.id">{{ f.label }}</option>
            <option v-if="otherFormat" :value="otherFormat">{{ tl('文件原有格式（') }}{{ otherFormat }}{{ tl('）') }}</option>
          </select>
        </label>
        <div class="dx-row" role="radiogroup" :aria-label="tl('頁碼編排')">
          <label><input v-model="restart" type="radio" name="dx-ps-restart" :value="false" /> {{ tl('接續前一節') }}</label>
          <label>
            <input v-model="restart" type="radio" name="dx-ps-restart" :value="true" /> {{ tl('起始頁碼') }}
            <input
              v-model.number="startAt"
              type="number"
              step="1"
              min="0"
              :max="MAX_START"
              :aria-label="tl('起始頁碼')"
              class="dx-start"
              @focus="restart = true"
            />
          </label>
        </div>
        <p class="dx-hint">
          {{ tl('這一節第一頁的頁碼：') }}<strong>{{ preview }}</strong>
          <template v-if="!restart && sectionInfo.index === 0">{{ tl('（第一節從 1 起算）') }}</template>
        </p>
        <label class="dx-row">
          <input v-model="titlePage" type="checkbox" /> {{ tl('首頁不同') }}
          <span class="dx-hint-inline">{{ tl('這一節第一頁用另外的頁首頁尾，空白就不顯示頁碼（內容在頁首頁尾編輯時設定）') }}</span>
        </label>
        <details class="dx-hint">
          <summary>{{ tl('封面不顯示頁碼、正文從 1 起算') }}</summary>
          <ol>
            <li>{{ tl('游標放在封面最後，按「版面配置 › 分隔設定 › 分節符號（下一頁）」插入分節符號。') }}</li>
            <li>{{ tl('游標在封面時勾選「首頁不同」（或讓封面的頁尾保持空白）。') }}</li>
            <li>{{ tl('游標移到正文，選「起始頁碼」並填 1，套用範圍選「目前這一節」。') }}</li>
          </ol>
        </details>
      </fieldset>

      <fieldset>
        <legend>{{ tl('分欄') }}</legend>
        <label v-if="sectionInfo.index > 0" class="dx-row">
          {{ tl('這一節的開始位置') }}
          <select v-model="sectionStart" :aria-label="tl('這一節的開始位置')">
            <option v-for="k in startOptions" :key="k" :value="k">{{ START_LABELS[k] }}</option>
          </select>
          <span class="dx-hint-inline">{{ tl('標題單欄、內文分欄時，把內文這一節設為「接續本頁」，兩者就在同一頁。') }}</span>
        </label>
        <div class="dx-row">
          <label>
            {{ tl('欄數') }}
            <select v-model.number="columnCount" :aria-label="tl('欄數')">
              <option v-for="n in COLUMN_COUNTS" :key="n" :value="n">{{ n === 1 ? '1（不分欄）' : n }}</option>
              <option v-if="otherCount" :value="otherCount">{{ otherCount }}{{ tl('（文件原有）') }}</option>
            </select>
          </label>
          <label>
            {{ tl('間距') }}
            <input
              v-model.number="columnSpace"
              type="number"
              step="0.05"
              min="0"
              :aria-label="tl('欄間距')"
              :disabled="columnCount === 1"
            />
            {{ tl('公分') }}
          </label>
        </div>
        <div class="dx-row">
          <label><input v-model="columnSeparator" type="checkbox" :disabled="columnCount === 1" /> {{ tl('分隔線') }}</label>
          <label :title="tl('各欄寬度相同（依紙張寬度、邊界與間距平均分配）')">
            <input type="checkbox" :checked="!unequalKept" disabled /> {{ tl('欄寬相等') }}
          </label>
        </div>
        <p v-if="unequalKept" class="dx-hint">{{ tl('這一節在 Word 中各欄寬度不同；變更欄數或間距後會改為欄寬相等。') }}</p>
        <p v-if="columnCount > 1" class="dx-hint dx-cols-note" role="note">{{ ONE_COLUMN_NOTE }}</p>
        <p class="dx-hint">{{ tl('要讓後面的文字從下一欄開始，請在該處按「版面配置 › 分隔設定 › 分欄符號」（Ctrl+Shift+Enter）。') }}</p>
      </fieldset>

      <fieldset>
        <legend>{{ tl('套用範圍') }}</legend>
        <label class="dx-row">
          <input v-model="scope" type="radio" name="dx-ps-scope" value="section" />
          {{ multi ? `目前這一節（第 ${sectionInfo.index + 1} 節，共 ${sectionInfo.count} 節）` : '整份文件' }}
        </label>
        <label v-if="multi" class="dx-row"><input v-model="scope" type="radio" name="dx-ps-scope" value="all" /> {{ tl('整份文件（所有分節）') }}</label>
      </fieldset>

      <p v-if="error" class="dx-error" role="alert">{{ error }}</p>
      <div class="dx-actions">
        <button type="button" @click="close">{{ tl('取消') }}</button>
        <button type="button" class="dx-primary" @click="apply">{{ tl('套用') }}</button>
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
  margin: 0 0 12px;
  font-size: 17px;
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
input[type='number'] {
  width: 5.5em;
  padding: 2px 4px;
}
select {
  padding: 2px 4px;
}
input.dx-start {
  width: 4.5em;
  margin-left: 4px;
}
.dx-hint {
  margin: 4px 0;
  font-size: 12.5px;
  color: #5f6368;
}
.dx-hint strong {
  color: #202124;
}
.dx-cols-note {
  padding: 4px 8px;
  border-left: 3px solid #1a73e8;
  background: #e8f0fe;
  color: #174ea6;
}
.dx-hint-inline {
  flex-basis: 100%;
  font-size: 12.5px;
  color: #5f6368;
}
.dx-hint ol {
  margin: 4px 0 0;
  padding-left: 20px;
}
.dx-hint summary {
  cursor: pointer;
}
.dx-error {
  color: #c5221f;
  margin: 0 0 8px;
}
.dx-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
.dx-actions button {
  padding: 5px 14px;
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
