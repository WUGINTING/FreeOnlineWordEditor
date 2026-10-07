import { defineComponent } from 'vue';
import { DocxEditor as BaseDocxEditor } from './papyrus/editorLoadStubs';

export const instances: ZoomEditor[] = [];
export class ZoomEditor extends (BaseDocxEditor as any) {
  zoomCalls: Array<number | 'fit'> = [];
  view = { dom: document.createElement('div') };
  constructor(host: HTMLElement, options: any) {
    super(host, options);
    instances.push(this);
  }
  setZoom(value: number | 'fit'): void { this.zoomCalls.push(value); }
}
export const DocxEditor = ZoomEditor;
export default defineComponent({ name: 'ZoomStub', render: () => null });
