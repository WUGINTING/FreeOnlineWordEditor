// The interface in other languages. The source text, Traditional Chinese, is the key: tl('檔案')
// gives 「檔案」 in zh-TW and what locales/<language>.ts says elsewhere; a text a language does not
// have is shown as it is. Numbered places take values: tl('第 {0} 頁，共 {1} 頁', page, pages).
//
// One language for the whole page (not per editor). It is read as texts are shown; set it before
// the editor is created, since some of what is already on the page keeps its language.

import en from './locales/en';
import zhCN from './locales/zh-CN';

/** The language the interface is written in. */
export const DEFAULT_LOCALE = 'zh-TW';

const catalogs = new Map<string, Record<string, string>>([
  ['en', en],
  ['zh-CN', zhCN],
]);

let current = DEFAULT_LOCALE;
let messages: Record<string, string> | undefined;
const listeners = new Set<() => void>();
/** Called on every lookup: the Vue layer makes what is on screen follow the language (vue/locale.ts). */
let onRead: () => void = () => {};

/** `text` in the interface's language, its {0}, {1} … replaced by `values`. */
export function tl(text: string, ...values: unknown[]): string {
  onRead();
  if (typeof text !== 'string') return text;
  const out = messages?.[text] ?? text;
  if (!values.length) return out;
  return out.replace(/\{(\d+)\}/g, (place, i) => (Number(i) < values.length ? String(values[Number(i)] ?? '') : place));
}

/** The languages the interface has, the one it is written in first. */
export function locales(): string[] {
  return [DEFAULT_LOCALE, ...catalogs.keys()];
}

/**
 * The interface language for a language tag (a browser's navigator.language, an <html lang>):
 * the tag itself when there is one, else the closest: Chinese in simplified script zh-CN, other
 * Chinese zh-TW, another language's regional form that language, anything unknown English.
 */
export function matchLocale(tag: string | null | undefined): string {
  const wanted = (tag ?? '').trim();
  if (!wanted) return DEFAULT_LOCALE;
  const known = locales();
  const same = known.find((l) => l.toLowerCase() === wanted.toLowerCase());
  if (same) return same;
  if (/^zh\b/i.test(wanted)) return /-(hans|cn|sg|my)\b/i.test(wanted) ? 'zh-CN' : DEFAULT_LOCALE;
  const language = wanted.split('-')[0].toLowerCase();
  return known.find((l) => l.toLowerCase() === language) ?? known.find((l) => l.toLowerCase().startsWith(`${language}-`)) ?? 'en';
}

/** The interface's language. */
export function getLocale(): string {
  return current;
}

/** Shows the interface in `locale` (see matchLocale for tags such as en-US); returns the language used. */
export function setLocale(locale: string | null | undefined): string {
  const next = matchLocale(locale);
  if (next === current) return current;
  current = next;
  messages = catalogs.get(next);
  for (const listener of [...listeners]) listener();
  return current;
}

/**
 * Adds texts to a language, or a language of your own: { '檔案': 'Datei', … }. A text given here
 * replaces the one the language had.
 */
export function addMessages(locale: string, texts: Record<string, string>): void {
  const merged = { ...catalogs.get(locale), ...texts };
  catalogs.set(locale, merged);
  if (locale === current) {
    messages = merged;
    for (const listener of [...listeners]) listener();
  }
}

/** Calls `listener` when the language or its texts change; returns how to stop. */
export function onLocaleChange(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** For the layer that draws the interface: `read` runs on every lookup (so it can follow the language). */
export function trackLocaleReads(read: () => void): void {
  onRead = read;
}
