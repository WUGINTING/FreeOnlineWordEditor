// The whole editor (ribbon, panels, dialogs) in a page element, for pages that are not Vue
// applications: plain HTML, React, Angular, a server-rendered page … It runs its own small Vue
// application inside the element.

import { createApp, h, shallowRef, type App } from 'vue';
import DocxEditorVue from './vue/DocxEditor.vue';
import type { DocxEditor } from './editor/core';
import { setLocale } from './i18n';
import type { CommentAuthor } from './docx/comments';
import type { CompatReport } from './docx/compat';
import type { MissingFont } from './docx/fonts';

/** A .docx as the editor takes it. */
export type DocxSource = Blob | ArrayBuffer | Uint8Array;

export interface CreateDocxEditorOptions {
  /** A .docx to open. When omitted, an empty document is created. */
  src?: DocxSource | null;
  editable?: boolean;
  toolbar?: boolean;
  /** Who writes new comments (「使用者」 when not given). */
  author?: CommentAuthor | null;
  /** The page has its own 檔案 menu: the ribbon's 「檔案」 calls `onFile` instead of opening its own. */
  fileMenu?: boolean;
  /** Read-only, but comments may be added, replied to and resolved. */
  commenting?: boolean;
  /**
   * The interface language, for the whole page: 'zh-TW' (the default), 'zh-CN', 'en', or a tag
   * such as navigator.language (see setLocale).
   */
  locale?: string;
  /** The editor is there and the first document is open. */
  onReady?: (editor: DocxEditor) => void;
  onChange?: () => void;
  /** A document could not be opened. */
  onError?: (error: unknown) => void;
  /** What in the opened document the web editor can't fully show or edit. */
  onCompat?: (report: CompatReport) => void;
  /** The ribbon's 「檔案」 (with `fileMenu`). */
  onFile?: () => void;
  /** About how big the saved .docx will be, after opening and after changes. */
  onSize?: (info: { bytes: number; large: boolean }) => void;
  /** The opened document's fonts this computer doesn't have ([] when none). */
  onFonts?: (missing: MissingFont[]) => void;
}

export interface DocxEditorHandle {
  /** Opens another document in place of the current one (null: an empty document). */
  open(src: DocxSource | null): void;
  /** The current document as a .docx. */
  save(): Promise<Blob>;
  /** Saves and has the browser download the file. */
  download(filename?: string): Promise<void>;
  /** The editor itself (commands, selection, undo …); null until `onReady`. */
  readonly editor: DocxEditor | null;
  /** Removes the editor from the page. */
  destroy(): void;
}

type Exposed = {
  open(src: DocxSource | null): unknown;
  save(): Promise<Blob>;
  download(filename?: string): Promise<void>;
  editor: DocxEditor | null;
};

/**
 * Puts the editor in `target` (an element, or a CSS selector of one). The element decides the
 * editor's size: give it a height.
 */
export function createDocxEditor(target: Element | string, options: CreateDocxEditorOptions = {}): DocxEditorHandle {
  const el = typeof target === 'string' ? document.querySelector(target) : target;
  if (!el) throw new Error(`createDocxEditor: no element for ${String(target)}`);

  if (options.locale) setLocale(options.locale);
  const component = shallowRef<Exposed | null>(null);
  const app: App = createApp({
    render: () =>
      h(DocxEditorVue, {
        ref: component,
        src: options.src ?? null,
        editable: options.editable ?? true,
        toolbar: options.toolbar ?? true,
        author: options.author ?? null,
        fileMenu: options.fileMenu ?? false,
        commenting: options.commenting ?? false,
        onReady: (editor: DocxEditor) => options.onReady?.(editor),
        onChange: () => options.onChange?.(),
        onError: (error: unknown) => options.onError?.(error),
        onCompat: (report: CompatReport) => options.onCompat?.(report),
        onFile: () => options.onFile?.(),
        onSize: (info: { bytes: number; large: boolean }) => options.onSize?.(info),
        onFonts: (missing: MissingFont[]) => options.onFonts?.(missing),
      }),
  });
  app.mount(el);

  const mounted = (): Exposed => {
    if (!component.value) throw new Error('createDocxEditor: the editor was destroyed');
    return component.value;
  };
  return {
    open(next) {
      void mounted().open(next);
    },
    save: () => mounted().save(),
    download: (filename) => mounted().download(filename),
    get editor() {
      return component.value?.editor ?? null;
    },
    destroy() {
      app.unmount();
      component.value = null;
    },
  };
}
