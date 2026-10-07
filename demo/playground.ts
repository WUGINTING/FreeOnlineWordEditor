// The test page: every option and call of createDocxEditor, with the events it reports.
// For trying the editor out before putting it in an application of your own.
import { createDocxEditor, type CreateDocxEditorOptions, type DocxEditorHandle } from '../src/papyrus';
import { SAMPLES, SHELL_CSS, currentLang, sampleTitle, sampleUrl, shellHeader, words } from './i18n';

const lang = currentLang();
const t = words(lang);
const app = document.getElementById('app')!;
document.title = `Papyrus DOCX ${t.playground}`;

const style = document.createElement('style');
style.textContent = `${SHELL_CSS}
  #app{display:grid;grid-template-rows:auto 1fr;background:#edf0f2}
  .bench{min-height:0;display:grid;grid-template-columns:300px 1fr;gap:12px;padding:12px}
  .panel{min-height:0;overflow:auto;background:#fff;border:1px solid #d2d9de;padding:12px 14px;font-size:13px}
  .panel h2{font-size:13px;margin:14px 0 6px;color:#17364c;text-transform:uppercase;letter-spacing:.5px}.panel h2:first-child{margin-top:0}
  .panel label{display:flex;gap:8px;align-items:center;padding:2px 0;font-family:ui-monospace,Consolas,monospace}
  .panel p{margin:4px 0;color:#64747d;font-size:12px}
  .panel select{width:100%;font:inherit;padding:4px}
  .buttons{display:flex;flex-wrap:wrap;gap:6px}
  .buttons button{font:inherit;padding:4px 10px;cursor:pointer;border:1px solid #aab6be;border-radius:4px;background:#f6f8f9}
  .buttons button:hover{background:#e8eef1}.buttons button:disabled{opacity:.5;cursor:default}
  .log-head{display:flex;justify-content:space-between;align-items:baseline}
  .log{margin:0;padding:0;list-style:none;font:12px/1.5 ui-monospace,Consolas,monospace}
  .log li{border-top:1px solid #edf0f2;padding:2px 0;word-break:break-word}.log b{color:#1f6f8f}.log .err b{color:#b3261e}
  .stage{min-height:0;background:#fff;border:1px solid #d2d9de;overflow:hidden;display:grid}
  .stage-empty{place-self:center;color:#64747d;font-size:14px}
  @media (max-width:820px){.bench{grid-template-columns:1fr;grid-template-rows:auto minmax(480px,1fr)}}
`;
document.head.append(style);

/** An element with these properties and children. */
function el<K extends keyof HTMLElementTagNameMap>(tag: K, props: Record<string, unknown> = {}, ...children: (Node | string)[]): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  Object.assign(node, props);
  node.append(...children);
  return node;
}

// ----- the editor and what it was made with -----

type Flag = 'editable' | 'toolbar' | 'commenting' | 'fileMenu';
const flags: Record<Flag, boolean> = { editable: true, toolbar: true, commenting: false, fileMenu: false };
let handle: DocxEditorHandle | null = null;
/** The document shown, kept so the editor can be made again with it. */
let current: Blob | null = null;

const stage = el('div', { className: 'stage' });
const log = el('ol', { className: 'log' });
/** The line counting a run of changes, while it is the newest line. */
let changeCount: Text | null = null;
let changes = 0;

function report(name: string, detail = '', error = false) {
  const time = new Date().toLocaleTimeString(lang === 'en' ? 'en-GB' : lang, { hour12: false });
  changeCount = null;
  log.prepend(el('li', { className: error ? 'err' : '' }, `${time} `, el('b', {}, name), detail ? ` ${detail}` : ''));
  while (log.children.length > 200) log.lastElementChild!.remove();
}

function create() {
  handle?.destroy();
  stage.replaceChildren();
  const options: CreateDocxEditorOptions = {
    ...flags,
    src: current,
    onReady: (editor) => report('onReady', `editable=${editor.editable}`),
    onChange: () => {
      // One line for a run of changes, not one per keystroke.
      if (changeCount) return void (changeCount.textContent = ` ×${++changes}`);
      report('onChange', '×1');
      changes = 1;
      changeCount = log.firstElementChild!.lastChild as Text;
    },
    onError: (error) => report('onError', error instanceof Error ? error.message : String(error), true),
    onCompat: (r) => report('onCompat', JSON.stringify(r).slice(0, 160)),
    onFile: () => report('onFile'),
    onSize: ({ bytes, large }) => {
      report('onSize', `${bytes.toLocaleString()} ${t.bytes}${large ? ' (large)' : ''}`);
    },
    onFonts: (missing) => report('onFonts', missing.length ? missing.map((f) => f.label).join(', ') : '[]'),
  };
  handle = createDocxEditor(stage, options);
  report('createDocxEditor', JSON.stringify(flags));
  sync();
}

function open(doc: Blob | null, what: string) {
  current = doc;
  if (!handle) return create();
  handle.open(doc);
  report('open', what);
}

// ----- the controls -----

const boxes = (Object.keys(flags) as Flag[]).map((flag) => {
  const box = el('input', { type: 'checkbox', checked: flags[flag] });
  box.onchange = () => {
    flags[flag] = box.checked;
    create();
  };
  return el('label', {}, box, flag);
});

const samples = el('select');
samples.add(new Option('—', ''));
for (const sample of SAMPLES) samples.add(new Option(sampleTitle(sample, lang), sample.file));
samples.onchange = async () => {
  if (!samples.value) return;
  try {
    const response = await fetch(sampleUrl(samples.value));
    if (!response.ok) throw new Error(String(response.status));
    open(await response.blob(), samples.selectedOptions[0].text);
  } catch (error) {
    report(t.loadFailed, String(error), true);
  }
};

const file = el('input', { type: 'file', accept: '.docx', hidden: true });
file.onchange = () => {
  const own = file.files?.[0];
  if (own) open(own, own.name);
  file.value = '';
  samples.value = '';
};

const button = (text: string, run: () => unknown) => {
  const b = el('button', { type: 'button' }, text);
  b.onclick = async () => {
    try {
      await run();
    } catch (error) {
      report(text, error instanceof Error ? error.message : String(error), true);
    }
  };
  return b;
};

const needsEditor: HTMLButtonElement[] = [];
const withEditor = (text: string, run: (h: DocxEditorHandle) => unknown) => {
  const b = button(text, () => run(handle!));
  needsEditor.push(b);
  return b;
};

const actions = el(
  'div',
  { className: 'buttons' },
  button(t.openFile, () => file.click()),
  button(t.blank, () => {
    samples.value = '';
    open(null, t.blank);
  }),
  withEditor(t.save, async (h) => {
    const blob = await h.save();
    report('save', `${t.saved}: ${blob.size.toLocaleString()} ${t.bytes}`);
  }),
  withEditor(t.download, (h) => h.download('papyrus-docx.docx')),
  withEditor(t.roundTrip, async (h) => {
    const blob = await h.save();
    open(blob, `${t.reopened}: ${blob.size.toLocaleString()} ${t.bytes}`);
  }),
  withEditor(t.destroy, (h) => {
    h.destroy();
    handle = null;
    stage.replaceChildren(el('div', { className: 'stage-empty' }, t.destroyed));
    report('destroy');
    sync();
  }),
  button(t.create, create),
);

function sync() {
  for (const b of needsEditor) b.disabled = !handle;
}

const panel = el(
  'aside',
  { className: 'panel' },
  el('h2', {}, t.options),
  ...boxes,
  el('p', {}, t.optionsNote),
  el('h2', {}, t.sample),
  samples,
  el('h2', {}, t.actions),
  actions,
  file,
  el('div', { className: 'log-head' }, el('h2', {}, t.events), button(t.clear, () => log.replaceChildren())),
  log,
);

app.append(shellHeader(lang, 'playground', t), el('main', { className: 'bench' }, panel, stage));
create();

// public/browser-check.js waits for this: the page could start in this browser.
(window as unknown as { __appStarted?: boolean }).__appStarted = true;
