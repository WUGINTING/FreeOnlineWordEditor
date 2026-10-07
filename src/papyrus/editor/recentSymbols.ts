// Word's 最近使用過的符號 (recently used symbols) for 插入 › 符號: the ribbon gallery shows them
// first and the 符號 dialog shows them in their own row. Kept in the browser (localStorage), one
// list for every document, as Word keeps one list per user. The stored value may be anything
// (another version, another page, a hand edit), so it is checked on every read.
import { MAX_RECENT_SYMBOLS, isSymbol } from './symbols';

/** The localStorage key of the list (a JSON array of characters, most recent first). */
export const RECENT_SYMBOLS_KEY = 'papyrus.recentSymbols';

/** Where the list is kept: localStorage, or a stand-in in tests. */
export type SymbolStore = Pick<Storage, 'getItem' | 'setItem'>;

/** The page's localStorage, or null where it can't be used (blocked, private mode, no window). */
function browserStore(): SymbolStore | null {
  try {
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

/** `ch` put first in `list`, without a second copy, and at most MAX_RECENT_SYMBOLS long. */
export function pushRecentSymbol(list: readonly string[], ch: string): string[] {
  return [ch, ...list.filter((x) => x !== ch)].slice(0, MAX_RECENT_SYMBOLS);
}

/** The recently used symbols, most recent first; empty when nothing (valid) is stored. */
export function readRecentSymbols(store: SymbolStore | null = browserStore()): string[] {
  let raw: string | null = null;
  try {
    raw = store?.getItem(RECENT_SYMBOLS_KEY) ?? null;
  } catch {
    return [];
  }
  if (!raw) return [];
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return [];
  }
  if (!Array.isArray(value)) return [];
  const out: string[] = [];
  for (const ch of value) {
    if (out.length >= MAX_RECENT_SYMBOLS) break;
    if (isSymbol(ch) && !out.includes(ch)) out.push(ch);
  }
  return out;
}

/**
 * Record that `ch` was inserted: it becomes the first recently used symbol. Returns the new
 * list (still right when it can't be stored: the storage is only a convenience).
 */
export function rememberSymbol(ch: string, store: SymbolStore | null = browserStore()): string[] {
  const list = readRecentSymbols(store);
  if (!isSymbol(ch)) return list;
  const next = pushRecentSymbol(list, ch);
  try {
    store?.setItem(RECENT_SYMBOLS_KEY, JSON.stringify(next));
  } catch {
    // not remembered (storage full or blocked); the symbol was still inserted
  }
  return next;
}
