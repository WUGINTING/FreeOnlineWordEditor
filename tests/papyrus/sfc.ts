// Loads a .vue single-file component in tests (the vitest config has no Vue plugin):
// compiles its <script setup> with the template inlined, with relative imports made absolute.
import { mkdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { compileScript, parse } from '@vue/compiler-sfc';
import type { Component } from 'vue';

export async function loadSfc(path: string): Promise<Component> {
  const { descriptor } = parse(readFileSync(path, 'utf8'), { filename: path });
  const { content } = compileScript(descriptor, { id: 'test', inlineTemplate: true });
  const code = content.replace(/from '(\.\.?\/[^']+)'/g, (_, p: string) => `from '${resolve(dirname(path), p).replace(/\\/g, '/')}'`);
  // Inside the project root (Vite does not serve files from elsewhere), outside src. In a git
  // worktree whose node_modules links to another checkout, the cache there resolves outside this
  // root: an ignored folder of the worktree itself is used instead.
  const modules = resolve('node_modules');
  let linked = false;
  try {
    linked = realpathSync(modules) !== modules;
  } catch {
    // no node_modules folder: the default below
  }
  const dir = linked ? resolve('build/papyrus-sfc') : resolve('node_modules/.cache/papyrus-sfc');
  mkdirSync(dir, { recursive: true });
  const file = join(dir, `${basename(path, '.vue')}-${process.pid}.ts`);
  writeFileSync(file, code);
  try {
    return (await import(/* @vite-ignore */ file.replace(/\\/g, '/'))).default;
  } finally {
    rmSync(file, { force: true });
  }
}
