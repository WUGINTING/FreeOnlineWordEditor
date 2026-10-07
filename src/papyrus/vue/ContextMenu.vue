<script setup lang="ts">
import { tl } from './locale';
import { computed, nextTick, onBeforeUnmount, ref, watch } from 'vue';
import type { Command, EditorState } from 'prosemirror-state';
import { selectAll } from 'prosemirror-commands';
import {
  addColumnAfter, addColumnBefore, addRowAfter, addRowBefore, deleteColumn, deleteRow, deleteTable, isInTable, mergeCells, splitCell,
} from 'prosemirror-tables';
import type { DocxEditor, EditorSnapshot } from '../editor/core';
import type { Mark } from 'prosemirror-model';
import { setLink } from '../editor/commands';
import { schema } from '../editor/schema';
import { isSafeHref, safeHref } from '../docx/links';
import { clipboardCommand, pasteClipboard, promptLink } from './editorActions';

// GOV-184: Word's right-click menu over the text (also Shift+F10 / the menu key): the clipboard,
// 段落…, 超連結, 新增留言 and, in a table, its rows and columns. The same commands as the ribbon.
const props = defineProps<{
  editor: DocxEditor | null;
  snapshot: EditorSnapshot | null;
  editable: boolean;
  /** Where it opens (viewport px); null: closed. */
  at: { x: number; y: number } | null;
}>();
const emit = defineEmits<{
  /** `refocus`: the keyboard goes back to the text (Esc, Tab, a command that stays in the text). */
  (e: 'close', refocus: boolean): void;
  (e: 'comment'): void;
  (e: 'paragraph'): void;
  (e: 'notice', message: string): void;
}>();

interface Item {
  key: string;
  label: string;
  keys?: string;
  disabled?: boolean;
  /** What it does; `stay`: the command leaves the keyboard where it puts it (a dialog, the comment box). */
  run: () => void;
  stay?: boolean;
}
type Entry = Item | 'sep';

const notice = (m: string) => emit('notice', m);

/**
 * The link the selection is on (read from the document when the menu opens, not from the ribbon's
 * snapshot, which follows a moment later): also at a link's first letter, where typing would not
 * extend it.
 */
function linkAt(state: EditorState): string | null {
  const { from, to, $from } = state.selection;
  let href: string | null = null;
  const check = (marks: readonly Mark[]) => {
    const m = schema.marks.link.isInSet(marks);
    if (m && href == null) href = m.attrs.href as string;
  };
  if (from === to) {
    check($from.marks());
    if ($from.nodeAfter) check($from.nodeAfter.marks);
  } else {
    state.doc.nodesBetween(from, to, (n) => {
      if (n.isText) check(n.marks);
    });
  }
  return href;
}
/** A ProseMirror command as an item, off when it cannot apply here. */
function command(key: string, label: string, cmd: Command): Item {
  const view = props.editor?.activeView;
  return { key, label, disabled: !view || !cmd(view.state), run: () => props.editor?.run(cmd) };
}

const entries = computed<Entry[]>(() => {
  if (!props.at) return [];
  const ed = props.editor;
  const view = ed?.activeView;
  void props.snapshot; // the items follow the editor while the menu is open
  const empty = !view || view.state.selection.empty;
  const link = view ? linkAt(view.state) : null;
  const inTable = !!view && isInTable(view.state);
  const edit = props.editable;
  const out: Entry[] = [];
  if (edit) out.push({ key: 'cut', label: '剪下', keys: 'Ctrl+X', disabled: empty, run: () => clipboardCommand(ed, 'cut', notice) });
  out.push({ key: 'copy', label: '複製', keys: 'Ctrl+C', disabled: empty, run: () => clipboardCommand(ed, 'copy', notice) });
  if (edit) out.push({ key: 'paste', label: '貼上', keys: 'Ctrl+V', run: () => void pasteClipboard(ed, notice) });
  out.push({ key: 'all', label: '全選', keys: 'Ctrl+A', run: () => ed?.run(selectAll) });
  if (edit) {
    out.push('sep', { key: 'paragraph', label: '段落…', run: () => emit('paragraph'), stay: true });
    out.push({ key: 'link', label: link ? '編輯超連結…' : '超連結…', run: () => promptLink(ed, link) });
    if (link) out.push({ key: 'unlink', label: '移除超連結', run: () => ed?.run(setLink(null)) });
  }
  if (link && isSafeHref(link)) {
    const href = safeHref(link);
    // Another site opens only once confirmed (the editor's own confirm, as for a click); a place in
    // the document is gone to (linkFollow.ts).
    out.push({ key: 'open-link', label: '開啟超連結', run: () => void ed?.followLink(href) });
  }
  if (edit && ed?.target === 'body') {
    out.push({ key: 'comment', label: '新增留言', keys: 'Ctrl+Alt+M', run: () => emit('comment'), stay: true });
  }
  if (edit && inTable) {
    out.push(
      'sep',
      command('row-before', '在上方插入列', addRowBefore),
      command('row-after', '在下方插入列', addRowAfter),
      command('col-before', '在左側插入欄', addColumnBefore),
      command('col-after', '在右側插入欄', addColumnAfter),
      command('merge', '合併儲存格', mergeCells),
      command('split', '分割儲存格', splitCell),
      command('del-row', '刪除列', deleteRow),
      command('del-col', '刪除欄', deleteColumn),
      command('del-table', '刪除表格', deleteTable),
    );
  }
  return out;
});
const items = computed(() => entries.value.filter((e): e is Item => e !== 'sep'));

const menu = ref<HTMLElement | null>(null);
const pos = ref({ left: 0, top: 0 });

/** The enabled items' buttons, in order. */
const buttons = () => Array.from(menu.value?.querySelectorAll<HTMLButtonElement>('[role="menuitem"]:not([disabled])') ?? []);

watch(
  () => props.at,
  async (at) => {
    if (!at) return;
    pos.value = { left: at.x, top: at.y };
    await nextTick();
    const el = menu.value;
    if (!el) return;
    // Kept inside the window: opens to the left / upwards when there is no room.
    const r = el.getBoundingClientRect();
    const w = window.innerWidth;
    const h = window.innerHeight;
    pos.value = {
      left: at.x + r.width > w ? Math.max(4, at.x - r.width) : at.x,
      top: at.y + r.height > h ? Math.max(4, h - r.height - 4) : at.y,
    };
    buttons()[0]?.focus();
  },
  { immediate: true },
);

function choose(item: Item) {
  if (item.disabled) return;
  // Closed first: the command then finds the text focused and the selection where it was.
  emit('close', !item.stay);
  item.run();
}

function onKey(e: KeyboardEvent) {
  const list = buttons();
  const at = list.indexOf(document.activeElement as HTMLButtonElement);
  const move = (i: number) => {
    e.preventDefault();
    list[(i + list.length) % list.length]?.focus();
  };
  if (e.key === 'ArrowDown') move(at + 1);
  else if (e.key === 'ArrowUp') move(at < 0 ? -1 : at - 1);
  else if (e.key === 'Home') move(0);
  else if (e.key === 'End') move(-1);
  else if (e.key === 'Escape' || e.key === 'Tab') {
    e.preventDefault();
    e.stopPropagation();
    emit('close', true);
  }
}

/** A press anywhere outside closes it (the click itself goes on to what is under it). */
function onPointerDown(e: Event) {
  if (props.at && menu.value && !menu.value.contains(e.target as Node)) emit('close', false);
}
const onScroll = (e: Event) => {
  if (props.at && !(menu.value && menu.value.contains(e.target as Node))) emit('close', false);
};
const onBlur = () => props.at && emit('close', false);
watch(
  () => !!props.at,
  (open) => {
    const add = open ? 'addEventListener' : 'removeEventListener';
    document[add]('mousedown', onPointerDown, true);
    document[add]('scroll', onScroll, true);
    window[add]('resize', onBlur);
    window[add]('blur', onBlur);
  },
  { immediate: true },
);
onBeforeUnmount(() => {
  document.removeEventListener('mousedown', onPointerDown, true);
  document.removeEventListener('scroll', onScroll, true);
  window.removeEventListener('resize', onBlur);
  window.removeEventListener('blur', onBlur);
});
</script>

<template>
  <div
    v-if="at"
    ref="menu"
    class="dx-context-menu"
    role="menu"
    :aria-label="tl('文字功能')"
    :style="{ left: `${pos.left}px`, top: `${pos.top}px` }"
    @keydown="onKey"
    @contextmenu.prevent
  >
    <template v-for="(entry, i) in entries" :key="entry === 'sep' ? `sep${i}` : entry.key">
      <div v-if="entry === 'sep'" class="dx-context-sep" role="separator" />
      <button
        v-else
        type="button"
        role="menuitem"
        :data-key="entry.key"
        :disabled="entry.disabled"
        tabindex="-1"
        @mousedown.prevent
        @click="choose(entry)"
      >
        <span>{{ entry.label }}</span><span v-if="entry.keys" class="dx-context-keys">{{ entry.keys }}</span>
      </button>
    </template>
    <div class="dx-context-hint">{{ tl('Ctrl+右鍵：瀏覽器原本的選單（拼字建議）') }}</div>
  </div>
</template>

<style scoped>
.dx-context-menu {
  position: fixed;
  z-index: 1000;
  min-width: 200px;
  padding: 4px 0;
  background: #fff;
  border: 1px solid #c8c8c8;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.18);
  font: 13px system-ui, 'Microsoft JhengHei', sans-serif;
  color: #202124;
}
button {
  display: flex;
  justify-content: space-between;
  gap: 24px;
  width: 100%;
  padding: 5px 14px;
  border: 0;
  background: transparent;
  font: inherit;
  color: inherit;
  text-align: left;
  cursor: pointer;
}
button:hover:not(:disabled),
button:focus-visible {
  background: #e8f0fe;
  outline: none;
}
button:disabled {
  color: #9aa0a6;
  cursor: default;
}
.dx-context-keys {
  color: #5f6368;
}
.dx-context-sep {
  height: 1px;
  margin: 4px 0;
  background: #e0e0e0;
}
.dx-context-hint {
  margin-top: 4px;
  padding: 4px 14px 2px;
  border-top: 1px solid #eee;
  font-size: 11px;
  color: #5f6368;
}
</style>
