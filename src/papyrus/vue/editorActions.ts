// Actions the ribbon (DocxToolbar.vue) and the right-click menu (ContextMenu.vue) share, so both do
// exactly the same: the clipboard and 超連結.
import { promptLink as corePromptLink, type DocxEditor } from '../editor/core';
import { tl } from './locale';

type Notice = (message: string) => void;

/**
 * 剪下 / 複製 from a button or menu: the browser lets a click do what Ctrl+X / Ctrl+C do; when it
 * refuses, the user is told the keys.
 */
export function clipboardCommand(editor: DocxEditor | null | undefined, command: 'cut' | 'copy', notice: Notice): void {
  const view = editor?.activeView;
  if (!view) return;
  view.focus();
  let ok = false;
  try {
    ok = document.execCommand(command);
  } catch {
    ok = false;
  }
  if (!ok) notice(command === 'cut' ? tl('請按 Ctrl+X 剪下。') : tl('請按 Ctrl+C 複製。'));
}

/** 貼上 from a button or menu: only with the user's permission to read the clipboard; else they are told to press Ctrl+V. */
export async function pasteClipboard(editor: DocxEditor | null | undefined, notice: Notice): Promise<void> {
  const view = editor?.activeView;
  if (!view) return;
  view.focus();
  try {
    if (navigator.clipboard?.read) {
      const items = await navigator.clipboard.read();
      const pictures: File[] = [];
      for (const item of items) {
        const image = item.types.find((t) => t.startsWith('image/'));
        if (image) pictures.push(new File([await item.getType(image)], 'image', { type: image }));
      }
      const html = items.find((item) => item.types.includes('text/html'));
      // Through the editor's own paste handling, as Ctrl+V: HTML wins over a picture of it (Excel),
      // and the pictures come along where the HTML shows them (Word, a web page).
      if (html) return void editor?.pasteHtml(await (await html.getType('text/html')).text(), pictures);
      if (pictures.length) {
        await editor?.insertImageFile(pictures[0]);
        return;
      }
      for (const item of items) {
        if (item.types.includes('text/plain')) return void view.pasteText(await (await item.getType('text/plain')).text());
      }
      return;
    }
    if (navigator.clipboard?.readText) return void view.pasteText(await navigator.clipboard.readText());
  } catch {
    // no permission: below
  }
  notice(tl('瀏覽器不允許按鈕讀取剪貼簿，請按 Ctrl+V 貼上。'));
}

/** 超連結: the editor's own (editor/core.ts promptLink), so the ribbon, Ctrl+K and this menu ask the same way. */
export function promptLink(editor: DocxEditor | null | undefined, _current?: string | null): void {
  corePromptLink(editor);
}
