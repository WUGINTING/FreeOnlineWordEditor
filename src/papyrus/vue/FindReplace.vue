<script setup lang="ts">
import { tl } from './locale';
// Find / replace panel (Ctrl+F / Ctrl+H). Searches the body (tables included) and every
// header and footer part shown on the pages, each part once (GOV-ISSUE-012); moving to a hit
// in a header/footer opens it for editing, as Word does. Text boxes, footnotes and comments
// are not searched, and the panel says so.
//
// Replace all first lists every match with its context and where it is; the user unticks the
// ones to keep (GOV-ISSUE-010). A match that may be the start of a longer term (第十二條 in
// 第十二條之一, 2025 in 20251) is flagged and unticked by default: replacing it changes a
// different citation, and the count on the button shows how many will be replaced.
// Undo: one Ctrl+Z (or 復原) right after undoes the whole replace-all, in the body and in every
// header/footer part (DocxEditor.changeSearchParts).
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue';
import type { Command, EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import type { DocxEditor, EditorSnapshot } from '../editor/core';
import {
  AREA_LABEL, clearSearch, findMatches, matchContext, nextMatchIndex, replaceCurrent, replaceMatchesTr, searchState, selectMatch,
  setSearch, type SearchArea, type SearchMatch, type SearchPart,
} from '../editor/search';
import { composing, onEnter } from './keys';

const props = withDefaults(
  defineProps<{ editor: DocxEditor | null; snapshot: EditorSnapshot | null; replace: boolean; canReplace?: boolean }>(),
  { canReplace: true },
);
const emit = defineEmits<{ (e: 'close'): void; (e: 'update:replace', value: boolean): void }>();

const query = ref('');
const replacement = ref('');
const caseSensitive = ref(false);
/** 區分全形／半形: off, as in Word, so 114 finds １１４. */
const widthSensitive = ref(false);
const message = ref('');
const root = ref<HTMLElement | null>(null);
const findInput = ref<HTMLInputElement | null>(null);
const replaceInput = ref<HTMLInputElement | null>(null);
const replaceAllButton = ref<HTMLButtonElement | null>(null);

/** Element ids of this panel (a version preview may show a second editor). */
const uid = `dx-find-${Math.random().toString(36).slice(2, 8)}`;

const view = (): EditorView | null => props.editor?.activeView ?? null;
/** Run a search command without taking the focus away from the panel. */
function exec(cmd: Command, v = view()): boolean {
  return !!v && cmd(v.state, v.dispatch, v);
}

// ----- matches in every part -----

/** Matches in the parts not being edited, found once per content and query. */
const found = new WeakMap<PMNode, { key: string; matches: SearchMatch[] }>();
function matchesIn(doc: PMNode, q: string, cs: boolean, ws: boolean): SearchMatch[] {
  const key = `${cs ? 1 : 0}${ws ? 1 : 0}${q}`;
  let f = found.get(doc);
  if (!f || f.key !== key) {
    f = { key, matches: findMatches(doc, q, cs, ws) };
    found.set(doc, f);
  }
  return f.matches;
}

interface PartMatches {
  part: SearchPart;
  matches: SearchMatch[];
}

/** Every part's matches of the current search, which part is being edited, the total and the current one's number. */
const results = computed(() => {
  void props.snapshot; // re-evaluated on every editor update
  const ed = props.editor;
  const v = view();
  const none = { parts: [] as PartMatches[], active: -1, total: 0, current: -1 };
  if (!ed || !v) return none;
  // The editor's search is the one to show (the typing may not have been searched yet).
  const s = searchState(v.state);
  if (!s.query) return none;
  const activeId = ed.activeSearchPart();
  const parts = ed.searchParts().map((part) => ({ part, matches: part.id === activeId ? s.matches : matchesIn(part.doc, s.query, s.caseSensitive, s.widthSensitive) }));
  const active = parts.findIndex((p) => p.part.id === activeId);
  let total = 0;
  let current = -1;
  parts.forEach((p, i) => {
    if (i === active && s.current >= 0) current = total + s.current;
    total += p.matches.length;
  });
  return { parts, active, total, current };
});

const countLabel = computed(() => {
  const r = results.value;
  if (!query.value) return '';
  // Still waiting for the typing to pause: not 「找不到」 yet (run-188).
  if (searching.value) return tl('搜尋中…');
  if (!r.total) return tl('找不到');
  return `${r.current >= 0 ? r.current + 1 : '–'} / ${r.total}`;
});

/** 「正文 6、頁首 1、頁尾 1」: areas without matches left out. */
function byArea(counts: Record<SearchArea, number>): string {
  return (['body', 'header', 'footer', 'textbox'] as SearchArea[])
    .filter((a) => counts[a])
    .map((a) => `${tl(AREA_LABEL[a])} ${counts[a]}`)
    .join(tl('、'));
}
function areaCounts(parts: { area: SearchArea; n: number }[]): Record<SearchArea, number> {
  const counts: Record<SearchArea, number> = { body: 0, header: 0, footer: 0, textbox: 0 };
  for (const p of parts) counts[p.area] += p.n;
  return counts;
}
/** Where the matches are, when the document has headers or footers. */
const breakdown = computed(() => {
  const r = results.value;
  if (!r.total || r.parts.length < 2) return '';
  return byArea(areaCounts(r.parts.map((p) => ({ area: p.part.area, n: p.matches.length }))));
});

/** What is searched: said in the panel so a 「找不到」 is not taken for the whole file. */
const scope = computed(() =>
  props.editor && !props.editor.editable
    ? tl('搜尋範圍：正文（含表格）。唯讀時不含頁首頁尾、文字方塊、註腳及註解。')
    : tl('搜尋範圍：正文（含表格）、頁首、頁尾、文字方塊。不含註腳及註解。'),
);

function scrollToCurrent() {
  nextTick(() => {
    const el = view()?.dom.querySelector('.dx-search-current') as HTMLElement | null;
    el?.scrollIntoView?.({ block: 'center', inline: 'nearest' });
  });
}

// ----- searching -----

/** Typing in the find box searches once the typing pauses (big documents). */
const SEARCH_DELAY = 150; // ms
let pending: ReturnType<typeof setTimeout> | null = null;
/** A search is owed to the typing (shown as 「搜尋中…」 rather than a stale count). */
const searching = ref(false);
function cancelPending() {
  if (pending) clearTimeout(pending);
  pending = null;
  searching.value = false;
}
function search() {
  cancelPending();
  message.value = '';
  exec(setSearch(query.value, caseSensitive.value, widthSensitive.value));
  scrollToCurrent();
}
/** Run the search the typing still owes, before acting on its matches. */
function flush() {
  if (pending) search();
}
watch(query, () => {
  cancelPending();
  closeReview(false);
  const v = view();
  if (v && searchState(v.state).query === query.value) return;
  pending = setTimeout(search, SEARCH_DELAY);
  searching.value = true;
});
watch([caseSensitive, widthSensitive], () => {
  closeReview(false);
  search();
});

// Header/footer editing started or ended: search where the cursor now is.
let searched: EditorView | null = null;
watch(
  () => props.snapshot?.target,
  () => {
    const v = view();
    if (v === searched) return;
    if (searched && !searched.isDestroyed) exec(clearSearch, searched);
    searched = v;
    search();
  },
);

/** Go to match `index` (-1: the last) of a part, opening it for editing when it is another part. */
function moveTo(part: SearchPart, index: number) {
  const ed = props.editor;
  if (!ed) return;
  const before = view();
  const focused = root.value?.contains(document.activeElement) ? (document.activeElement as HTMLElement) : null;
  const next = ed.openSearchPart(part.id);
  if (!next) {
    message.value = tl('無法開啟{0}', part.label);
    return;
  }
  if (next !== before) {
    if (before && !before.isDestroyed) exec(clearSearch, before);
    searched = next;
    exec(setSearch(query.value, caseSensitive.value, widthSensitive.value), next);
  }
  exec(selectMatch(index), next);
  // Opening a header/footer focuses it: typing Enter again should go on searching.
  focused?.focus();
  scrollToCurrent();
}

/** The next (1) or previous (-1) match: in this part, else in the next part that has one (wrapping around). */
function go(dir: 1 | -1) {
  flush();
  message.value = '';
  const r = results.value;
  const v = view();
  if (!r.total || !v) return;
  const i = r.active >= 0 ? nextMatchIndex(v.state, dir) : -1;
  if (i >= 0) {
    exec(selectMatch(i));
    scrollToCurrent();
    return;
  }
  const n = r.parts.length;
  let k = r.active >= 0 ? r.active : dir > 0 ? -1 : 0;
  for (let step = 0; step < n; step++) {
    k = (k + dir + n) % n;
    if (r.parts[k].matches.length) break;
  }
  moveTo(r.parts[k].part, dir > 0 ? 0 : -1);
}

function replaceOne() {
  flush();
  closeReview(false);
  if (!exec(replaceCurrent(replacement.value, false))) {
    message.value = tl('沒有可取代的項目');
    return;
  }
  go(1);
}

// ----- replace all: review every match first (GOV-ISSUE-010) -----

interface ReviewItem {
  id: number;
  partId: string;
  area: SearchArea;
  where: string;
  from: number;
  to: number;
  before: string;
  text: string;
  after: string;
  longer: boolean;
  checked: boolean;
}
const reviewItems = ref<ReviewItem[] | null>(null);
/** The content each part had when the list was made: replacing checks it is still the same. */
let reviewDocs = new Map<string, PMNode>();
const reviewList = ref<HTMLElement | null>(null);
const checkedCount = computed(() => reviewItems.value?.filter((it) => it.checked).length ?? 0);
const flaggedCount = computed(() => reviewItems.value?.filter((it) => it.longer).length ?? 0);

function openReview() {
  flush();
  message.value = '';
  const r = results.value;
  if (!r.total) return;
  reviewDocs = new Map(r.parts.map((p) => [p.part.id, p.part.doc]));
  let id = 0;
  reviewItems.value = r.parts.flatMap((p) =>
    p.matches.map((m) => {
      const ctx = matchContext(p.part.doc, m);
      return { id: id++, partId: p.part.id, area: p.part.area, where: p.part.label, from: m.from, to: m.to, ...ctx, checked: !ctx.longer };
    }),
  );
  nextTick(() => reviewList.value?.querySelector<HTMLInputElement>('input[type="checkbox"]')?.focus());
}

function closeReview(refocus = true) {
  if (!reviewItems.value) return;
  reviewItems.value = null;
  reviewDocs = new Map();
  if (refocus) nextTick(() => (replaceAllButton.value?.disabled ? replaceInput.value : replaceAllButton.value)?.focus());
}

function checkAll(on: boolean) {
  for (const it of reviewItems.value ?? []) it.checked = on;
}

function confirmReview() {
  const ed = props.editor;
  const items = reviewItems.value;
  if (!ed || !items) return;
  const parts = ed.searchParts();
  // Edited meanwhile (typing in the page, undo ...): the positions no longer hold.
  const stale = [...reviewDocs].some(([id, doc]) => parts.find((p) => p.id === id)?.doc !== doc);
  if (stale) {
    openReview();
    message.value = tl('文件已變更，已重新列出相符項目，請再確認。');
    return;
  }
  const groups = new Map<string, SearchMatch[]>();
  let skipped = 0;
  for (const it of items) {
    if (!it.checked) skipped++;
    else groups.set(it.partId, [...(groups.get(it.partId) ?? []), { from: it.from, to: it.to }]);
  }
  const changed = ed.changeSearchParts(
    [...groups].map(([id, matches]) => ({ id, build: (state: EditorState) => replaceMatchesTr(state, matches, replacement.value) })),
  );
  const done = changed.map((id) => ({ area: parts.find((p) => p.id === id)!.area, n: groups.get(id)!.length }));
  closeReview();
  const total = done.reduce((n, d) => n + d.n, 0);
  const rest = skipped ? tl('，略過 {0} 處', skipped) : '';
  message.value = total ? tl('已取代 {0} 處（{1}）{2}', total, byArea(areaCounts(done)), rest) : tl('沒有取代任何項目{0}', rest);
}

// ----- keys, closing, focus -----

function onFindKey(e: KeyboardEvent) {
  // Enter picking a character of an input method (注音 …) is not 「下一個」.
  if (e.key === 'Enter' && !composing(e)) {
    e.preventDefault();
    go(e.shiftKey ? -1 : 1);
  }
}

/** Enter in 「取代為」 replaces, but not the Enter that picks a character of an input method. */
const onReplaceKey = onEnter(() => replaceOne());
/** Esc closes the panel (the review first), but not the Esc that drops an input method's candidates. */
function onEscape(e: KeyboardEvent, action: () => void, stop = false) {
  if (composing(e)) return;
  e.preventDefault();
  if (stop) e.stopPropagation();
  action();
}

function close() {
  const v = view();
  // The search the typing still owes (even one not scheduled yet), so the match to select exists.
  if (v && query.value && searchState(v.state).query !== query.value) search();
  // As in Word, closing leaves the found text selected, so typing goes on from there (GOV-113).
  const current = v ? searchState(v.state).current : -1;
  if (v && current >= 0) exec(selectMatch(current), v);
  cancelPending();
  exec(clearSearch, v);
  if (props.editor?.view && props.editor.view !== v) exec(clearSearch, props.editor.view);
  emit('close');
  v?.focus();
}

/** Focus the find box (or the replace box), starting from the selected text like Word. */
function focus(replaceMode = props.replace) {
  const v = view();
  if (v) {
    const { from, to, empty } = v.state.selection;
    const text = empty ? '' : v.state.doc.textBetween(from, to, '\n', '\n');
    if (text && !text.includes('\n') && text.length <= 200) query.value = text;
  }
  nextTick(() => {
    const el = replaceMode && query.value ? replaceInput.value : findInput.value;
    el?.focus();
    el?.select();
  });
}

onMounted(() => {
  searched = view();
  focus();
  search();
});
onBeforeUnmount(() => {
  cancelPending();
  if (searched && !searched.isDestroyed) exec(clearSearch, searched);
});

defineExpose({ focus });
</script>

<template>
  <div ref="root" class="dx-find" role="search" :aria-label="tl('尋找與取代')" @keydown.esc="onEscape($event, close)">
    <div class="dx-find-row">
      <span v-if="!canReplace" class="dx-toggle-space" />
      <button
        v-else
        type="button"
        class="dx-toggle"
        :title="replace ? tl('隱藏取代') : tl('顯示取代 (Ctrl+H)')"
        :aria-label="replace ? tl('隱藏取代') : tl('顯示取代')"
        :aria-expanded="replace"
        @click="emit('update:replace', !replace)"
      >{{ replace ? '▾' : '▸' }}</button>
      <input
        ref="findInput"
        v-model="query"
        type="text"
        class="dx-find-input"
        :placeholder="tl('尋找')"
        :aria-label="tl('尋找')"
        :aria-describedby="`${uid}-scope`"
        @keydown="onFindKey"
      />
      <span class="dx-count" aria-live="polite">{{ countLabel }}</span>
      <button type="button" :title="tl('上一個 (Shift+Enter)')" :aria-label="tl('上一個')" :disabled="!results.total" @click="go(-1)">↑</button>
      <button type="button" :title="tl('下一個 (Enter)')" :aria-label="tl('下一個')" :disabled="!results.total" @click="go(1)">↓</button>
      <button type="button" :title="tl('關閉 (Esc)')" :aria-label="tl('關閉')" @click="close">✕</button>
    </div>
    <div v-if="replace && canReplace" class="dx-find-row">
      <span class="dx-toggle-space" />
      <input
        ref="replaceInput"
        v-model="replacement"
        type="text"
        class="dx-find-input"
        :placeholder="tl('取代為')"
        :aria-label="tl('取代為')"
        @keydown="onReplaceKey"
      />
      <button type="button" :disabled="!results.total" @click="replaceOne">{{ tl('取代') }}</button>
      <button ref="replaceAllButton" type="button" :title="tl('先列出所有相符項目，確認後再取代')" :disabled="!results.total" @click="openReview">{{ tl('全部取代') }}</button>
    </div>
    <div class="dx-find-row dx-find-options">
      <span class="dx-toggle-space" />
      <label><input v-model="caseSensitive" type="checkbox" /> {{ tl('區分大小寫') }}</label>
      <label :title="tl('不勾選時，「114」也會找到全形的「１１４」')"><input v-model="widthSensitive" type="checkbox" /> {{ tl('區分全形／半形') }}</label>
      <span v-if="breakdown" class="dx-find-where">{{ breakdown }}</span>
      <span v-if="message" class="dx-find-msg" role="status">{{ message }}</span>
    </div>
    <div :id="`${uid}-scope`" class="dx-find-row dx-find-scope">
      <span class="dx-toggle-space" />
      <span>{{ scope }}</span>
    </div>

    <div
      v-if="reviewItems"
      ref="reviewList"
      class="dx-review"
      role="dialog"
      :aria-labelledby="`${uid}-review`"
      @keydown.esc="onEscape($event, () => closeReview(), true)"
    >
      <div :id="`${uid}-review`" class="dx-review-title">{{ tl('將「{0}」取代為「{1}」：共 {2} 處，請確認要取代的項目', query, replacement, reviewItems.length) }}</div>
      <p v-if="flaggedCount" class="dx-review-note">
        {{ tl('標示「後面還有文字」的 {0} 處可能是較長的詞（如「第十二條之一」），預設不取代；確認要取代再勾選。', flaggedCount) }}
      </p>
      <div class="dx-review-bulk">
        <button type="button" @click="checkAll(true)">{{ tl('全選') }}</button>
        <button type="button" @click="checkAll(false)">{{ tl('全不選') }}</button>
      </div>
      <ul class="dx-review-list" :aria-label="tl('相符項目')">
        <li v-for="it in reviewItems" :key="it.id" :class="{ 'dx-review-longer': it.longer }">
          <label>
            <input v-model="it.checked" type="checkbox" />
            <span class="dx-review-where">{{ it.where }}</span>
            <span class="dx-review-text">{{ it.before }}<mark>{{ it.text }}</mark>{{ it.after }}</span>
            <span v-if="it.longer" class="dx-review-hint">{{ tl('後面還有文字，可能是較長的詞') }}</span>
          </label>
        </li>
      </ul>
      <div class="dx-review-actions">
        <button type="button" class="dx-primary" :disabled="!checkedCount" @click="confirmReview">{{ tl('全部取代（{0} 處）', checkedCount) }}</button>
        <button type="button" @click="closeReview()">{{ tl('取消') }}</button>
      </div>
    </div>
  </div>
</template>

<style scoped>
.dx-find {
  position: absolute;
  top: 8px;
  right: 20px;
  z-index: 30;
  display: flex;
  flex-direction: column;
  gap: 4px;
  max-width: calc(100% - 40px);
  padding: 8px;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 8px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
  font: 13px system-ui, 'Microsoft JhengHei', sans-serif;
  color: #3c4043;
}
.dx-find-row {
  display: flex;
  align-items: center;
  gap: 4px;
}
.dx-find-input {
  width: 200px;
  min-width: 0;
  height: 28px;
  box-sizing: border-box;
  border: 1px solid #dadce0;
  border-radius: 4px;
  padding: 0 6px;
  font: inherit;
}
.dx-count {
  min-width: 52px;
  text-align: center;
  font-size: 12px;
  color: #5f6368;
  white-space: nowrap;
}
button {
  height: 28px;
  min-width: 28px;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: pointer;
  padding: 0 6px;
}
button:hover:not(:disabled) {
  background: #f1f3f4;
}
button:disabled {
  opacity: 0.35;
  cursor: default;
}
button:focus-visible,
input:focus-visible {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}
.dx-toggle,
.dx-toggle-space {
  width: 20px;
  min-width: 20px;
  padding: 0;
}
.dx-find-options {
  flex-wrap: wrap;
  font-size: 12px;
}
.dx-find-options label {
  display: inline-flex;
  align-items: center;
  gap: 3px;
  cursor: pointer;
}
.dx-find-where {
  color: #5f6368;
}
.dx-find-msg {
  margin-left: auto;
  color: #188038;
}
.dx-find-scope {
  max-width: 360px;
  font-size: 11px;
  color: #5f6368;
  align-items: flex-start;
}
.dx-review {
  display: flex;
  flex-direction: column;
  gap: 6px;
  max-width: 420px;
  margin-top: 4px;
  padding-top: 8px;
  border-top: 1px solid #dadce0;
}
.dx-review-title {
  font-weight: 600;
}
.dx-review-note {
  margin: 0;
  font-size: 12px;
  color: #b06000;
}
.dx-review-bulk {
  display: flex;
  gap: 4px;
}
.dx-review-list {
  max-height: 260px;
  overflow-y: auto;
  margin: 0;
  padding: 0;
  list-style: none;
  border: 1px solid #dadce0;
  border-radius: 4px;
}
.dx-review-list li + li {
  border-top: 1px solid #f1f3f4;
}
.dx-review-list label {
  display: flex;
  flex-wrap: wrap;
  align-items: baseline;
  gap: 2px 6px;
  padding: 4px 6px;
  cursor: pointer;
}
.dx-review-list label:focus-within {
  background: #e8f0fe;
}
.dx-review-where {
  flex: none;
  padding: 0 4px;
  border-radius: 3px;
  background: #f1f3f4;
  font-size: 11px;
  color: #5f6368;
}
.dx-review-text {
  flex: 1 1 200px;
  min-width: 0;
  overflow-wrap: anywhere;
}
.dx-review-text mark {
  background: #fde293;
  color: inherit;
}
.dx-review-hint {
  flex-basis: 100%;
  padding-left: 22px;
  font-size: 11px;
  color: #b06000;
}
.dx-review-longer .dx-review-text mark {
  background: #fad2cf;
}
.dx-review-actions {
  display: flex;
  justify-content: flex-end;
  gap: 4px;
}
.dx-primary:not(:disabled) {
  background: #1a73e8;
  color: #fff;
}
.dx-primary:hover:not(:disabled) {
  background: #1765cc;
}
</style>
