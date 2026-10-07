// What the interface's languages need: every Chinese text in the editor's source (src/papyrus) is
// either an interface text, which each language file (src/papyrus/locales) must have, or listed in
// not-interface.json (document content, text matched against documents, font names …).
//
//   node tests/papyrus/i18n/scan.mjs            what each language lacks, and what it no longer needs
//   node tests/papyrus/i18n/scan.mjs --keys     the interface texts, one JSON string per line
//
// tests/papyrus/i18n.test.ts runs the same check.
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, resolve, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '../../..');
const require = createRequire(join(root, 'package.json'));
const ts = require('typescript');
const { parse } = require('@vue/compiler-sfc');

/** Chinese characters and full-width punctuation. */
export const CJK = /[㐀-鿿　-〿＀-￯]/;
const SOURCE = 'src/papyrus';
const LOCALES = 'src/papyrus/locales';

const walk = (dir) =>
  readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    return statSync(path).isDirectory() ? walk(path) : [path];
  });

/**
 * The Chinese texts of the source: `texts` (text → the places it is written, "file:line") for string
 * literals and for whatever tl() is given, `templates` for template literals with values in them
 * (as their text with {0}, {1} … places: an interface text is written tl('… {0} …', value) instead),
 * and `problems` for text left as it is in a Vue template.
 */
export function scanSource() {
  const texts = new Map();
  const templates = new Map();
  const problems = [];
  const add = (text, where, always = false) => {
    if (!always && !CJK.test(text)) return;
    if (!texts.has(text)) texts.set(text, []);
    texts.get(text).push(where);
  };

  const scanCode = (file, code, firstLine, what) => {
    const sf = ts.createSourceFile('x.ts', code, ts.ScriptTarget.Latest, true);
    const visit = (node) => {
      const where = () => `${file}:${firstLine + sf.getLineAndCharacterOfPosition(node.getStart()).line}`;
      if (ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) {
        const call = node.parent;
        const givenToTl = ts.isCallExpression(call) && ts.isIdentifier(call.expression) && call.expression.text === 'tl' && call.arguments[0] === node;
        if (!ts.isImportDeclaration(call) && !ts.isExportDeclaration(call)) add(node.text, where(), givenToTl);
      } else if (ts.isTemplateExpression(node)) {
        const text = node.head.text + node.templateSpans.map((s, i) => `{${i}}${s.literal.text}`).join('');
        if (CJK.test(text)) templates.set(text, [...(templates.get(text) ?? []), where()]);
      }
      ts.forEachChild(node, visit);
    };
    visit(sf);
  };

  for (const path of walk(join(root, SOURCE))) {
    const file = path.slice(root.length + 1).split(sep).join('/');
    if (file.startsWith(`${LOCALES}/`)) continue;
    const src = readFileSync(path, 'utf8');
    if (file.endsWith('.ts')) scanCode(file, src, 1, 'script');
    if (!file.endsWith('.vue')) continue;
    const { descriptor } = parse(src, { filename: file });
    for (const block of [descriptor.script, descriptor.scriptSetup]) if (block) scanCode(file, block.content, block.loc.start.line, 'script');
    const visit = (node) => {
      const line = node.loc?.start.line ?? 0;
      if (node.type === 2 && CJK.test(node.content)) problems.push(`${file}:${line} template: text not through tl(): ${node.content.trim().slice(0, 40)}`);
      else if (node.type === 5 && node.content.content) scanCode(file, `(${node.content.content})`, line, 'template');
      else if (node.type === 1) {
        for (const p of node.props) {
          if (p.type === 6 && p.value && CJK.test(p.value.content)) problems.push(`${file}:${p.loc.start.line} template: attribute ${p.name} not through tl(): ${p.value.content.slice(0, 40)}`);
          // v-for="x in list" and the like are not expressions by themselves: only their literals matter.
          else if (p.type === 7 && p.exp?.content && CJK.test(p.exp.content)) scanCode(file, p.name === 'for' ? `[${p.exp.content.replace(/^.*?\s(?:in|of)\s/, '')}]` : `(${p.exp.content})`, p.loc.start.line, 'template');
        }
      }
      for (const child of node.children ?? []) visit(child);
      for (const branch of node.branches ?? []) visit(branch);
    };
    if (descriptor.template?.ast) visit(descriptor.template.ast);
  }
  return { texts, templates, problems };
}

/** The texts that are not interface text: text → why. */
export function notInterface() {
  const list = JSON.parse(readFileSync(join(here, 'not-interface.json'), 'utf8'));
  return new Map(list.map((entry) => [entry.text, entry.why]));
}

/** The languages' texts: locale → { source text: translation }. */
export async function catalogs() {
  const out = {};
  for (const name of readdirSync(join(root, LOCALES)).filter((n) => n.endsWith('.ts'))) {
    // The files are plain objects in TypeScript: turned into JavaScript and loaded.
    const code = ts.transpileModule(readFileSync(join(root, LOCALES, name), 'utf8'), { compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ESNext } }).outputText;
    out[name.replace(/\.ts$/, '')] = (await import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`)).default;
  }
  return out;
}

const places = (text) => [...new Set((text.match(/\{\d+\}/g) ?? []))].sort().join('');

/**
 * The whole check: `keys` the interface texts, and per language what it lacks (`missing`), has but
 * the source no longer uses (`unused`), and translates with other {0} places than the source (`places`).
 */
export async function check() {
  const { texts, templates, problems } = scanSource();
  const exempt = notInterface();
  const keys = [...texts.keys()].filter((text) => !exempt.has(text));
  for (const [text, where] of templates) {
    if (!exempt.has(text)) problems.push(`${where[0]}: a template literal with Chinese text and values, neither through tl() nor listed: ${text.slice(0, 60)}`);
  }
  const unusedExempt = [...exempt.keys()].filter((text) => !texts.has(text) && !templates.has(text));
  const languages = {};
  for (const [locale, catalog] of Object.entries(await catalogs())) {
    languages[locale] = {
      missing: keys.filter((key) => !(key in catalog) || !String(catalog[key]).trim()),
      unused: Object.keys(catalog).filter((key) => !texts.has(key)),
      places: keys.filter((key) => key in catalog && places(key) !== places(String(catalog[key]))),
    };
  }
  return { keys, texts, templates, problems, unusedExempt, languages };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const result = await check();
  if (process.argv.includes('--keys')) {
    for (const key of result.keys) console.log(JSON.stringify(key));
  } else {
    console.log(`interface texts: ${result.keys.length}`);
    for (const problem of result.problems) console.log(`  problem: ${problem}`);
    for (const text of result.unusedExempt) console.log(`  not-interface.json lists a text the source no longer has: ${JSON.stringify(text)}`);
    for (const [locale, l] of Object.entries(result.languages)) {
      console.log(`${locale}: ${l.missing.length} missing, ${l.unused.length} no longer used, ${l.places.length} with other {n} places`);
      for (const key of l.missing.slice(0, 40)) console.log(`  missing ${JSON.stringify(key)}  (${result.texts.get(key)[0]})`);
      if (l.missing.length > 40) console.log(`  … and ${l.missing.length - 40} more`);
      for (const key of l.unused) console.log(`  no longer used ${JSON.stringify(key)}`);
      for (const key of l.places) console.log(`  other places ${JSON.stringify(key)}`);
    }
    const bad = result.problems.length + result.unusedExempt.length + Object.values(result.languages).reduce((n, l) => n + l.missing.length + l.unused.length + l.places.length, 0);
    process.exit(bad ? 1 : 0);
  }
}
