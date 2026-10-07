import { defineComponent } from 'vue';
import { DocxEditor as BaseDocxEditor } from './papyrus/editorLoadStubs';

export const instances: SpellEditor[] = [];

export class SpellEditor extends (BaseDocxEditor as any) {
  spellcheckOn = false;
  view: any;
  constructor(host: HTMLElement, options: any) {
    super(host, options);
    this.spellcheckOn = options?.spellcheck === true;
    this.view = { dom: document.createElement('div') };
    this.view.dom.setAttribute('spellcheck', String(this.spellcheckOn));
    instances.push(this);
  }
  setSpellcheck(on: boolean): void {
    this.spellcheckOn = on;
    this.view.dom.setAttribute('spellcheck', String(on));
  }
}

export const DocxEditor = SpellEditor;
export default defineComponent({ name: 'SpellcheckStub', render: () => null });
