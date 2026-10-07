// Stand-ins used when DocxEditor.vue is compiled in a test (editorLoad.test.ts): an editor
// whose open() finishes only when the test says so, and empty child components.
import { defineComponent } from 'vue';

/** Resolves a pending open() by document name. */
export const gates = new Map<string, () => void>();

export class DocxEditor {
  model: { zip: unknown; paragraphStyles: unknown[] } = { zip: null, paragraphStyles: [] };
  view = null;
  activeView = null;
  opened: string[] = [];

  constructor(_host: HTMLElement, _options: unknown) {}

  open(src: { name: string; zip: unknown }): Promise<void> {
    return new Promise((resolve) => {
      gates.set(src.name, () => {
        gates.delete(src.name);
        this.model = { zip: src.zip, paragraphStyles: [] };
        this.opened.push(src.name);
        resolve();
      });
    });
  }

  newDocument(): Promise<void> {
    return Promise.resolve();
  }

  /** No fields to fill in (editor/fields.ts). */
  fields(): unknown[] {
    return [];
  }

  wordCount() {
    return { words: 0, chars: 0, charsWithSpaces: 0 };
  }

  selectionWordCount(): null {
    return null;
  }

  setSpellcheck(): void {}

  /** Fonts this computer lacks (DocxEditor.vue's `fonts` event): none. */
  missingFonts(): unknown[] {
    return [];
  }

  /** The document's size (DocxEditor.vue's `size` event). */
  estimatedSize(): number {
    return 0;
  }

  /** Comments (editor/review.ts): none, and nowhere to add one. */
  comments(): unknown[] {
    return [];
  }

  author = { name: '使用者', initials: '使' };
  setAuthor(): void {}
  commentTarget(): null {
    return null;
  }
  addComment(): null {
    return null;
  }
  replyComment(): null {
    return null;
  }
  resolveComment(): boolean {
    return false;
  }
  editComment(): boolean {
    return false;
  }
  deleteComment(): boolean {
    return false;
  }
  canChangeComment(): boolean {
    return false;
  }

  outline(): unknown[] {
    return [];
  }

  pageReferenceCount(): number {
    return 0;
  }

  destroy(): void {}
}

export default defineComponent({ name: 'Stub', render: () => null });
