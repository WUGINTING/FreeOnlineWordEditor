// Public API
export { createDocxEditor, type CreateDocxEditorOptions, type DocxEditorHandle, type DocxSource } from './mount';
export { default as DocxEditorVue } from './vue/DocxEditor.vue';
export { default as DocxToolbar } from './vue/DocxToolbar.vue';
export { DocxEditor, type EditorSnapshot, type DocxEditorOptions } from './editor/core';
export { readDocx } from './docx/reader';
export { writeDocx } from './docx/writer';
export { blankPackage } from './docx/template';
export { schema } from './editor/schema';
export * as commands from './editor/commands';
export type { DocxModel, PageSetup } from './docx/model';
export { MACRO_NOTICE } from './docx/macros';
export { externalLinkQuestion } from './editor/linkFollow';
export { compareStructure, structureChanges, type StructureChange } from './docx/structureDiff';
