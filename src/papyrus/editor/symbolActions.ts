// 插入 › 符號 from the ribbon gallery and the 符號 dialog: insert, then remember as recently used.
// Kept apart from recentSymbols.ts, which is only the stored list.
import type { DocxEditor } from './core';
import { insertSymbol } from './commands';
import { rememberSymbol } from './recentSymbols';

/**
 * Insert `ch` at the cursor of `editor` (the body, or the header/footer being edited), with the
 * document's styles deciding whether it takes w:hint="eastAsia", and, when it went in, make it the
 * first recently used symbol. Nothing in read-only mode; a symbol that didn't go in (refused while
 * tracking changes, say) isn't remembered. `focus: false` leaves the keyboard where it is (the 符號
 * dialog stays open). Returns whether the document changed.
 */
export function insertSymbolAndRemember(editor: DocxEditor, ch: string, options: { focus?: boolean } = {}): boolean {
  const view = editor.activeView;
  if (!view || !editor.editable) return false;
  const before = view.state.doc;
  editor.run(insertSymbol(ch, editor.model.styles), options);
  const changed = view.state.doc !== before;
  if (changed) rememberSymbol(ch);
  return changed;
}
