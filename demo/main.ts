// The demo page: one of the sample documents, or a .docx from this computer, in the editor.
// Nothing leaves the browser; 「檔案 › 下載」 in the ribbon saves the edited document.
import { createDocxEditor } from '../src/papyrus';
import { SAMPLES, SHELL_CSS, currentLang, sampleTitle, sampleUrl, shellHeader, words } from './i18n';

const lang = currentLang();
const t = words(lang);
const app = document.getElementById('app')!;
document.title = `Papyrus DOCX ${t.demo}`;

const style = document.createElement('style');
style.textContent = `${SHELL_CSS}
  #app{display:grid;grid-template-rows:auto 1fr auto;background:#edf0f2}
  .workspace{min-height:0;display:flex;flex-direction:column;padding:0 18px 10px}
  .doc-head{display:flex;flex-wrap:wrap;align-items:center;gap:8px 32px;min-height:56px;padding:8px 12px}
  .picker{font-size:14px;font-weight:600;display:flex;gap:12px;align-items:center}
  .picker select{min-width:250px;padding:7px 10px;background:#fff;border:1px solid #aab6be;border-radius:4px;font:inherit;color:#1d2c36}
  .picker input{font:inherit;font-weight:400}
  .doc-head .state{margin-left:auto;color:#2d6f60}
  .editor-wrap{min-height:0;flex:1;background:#fff;border:1px solid #d2d9de;box-shadow:0 2px 8px #17364c17;overflow:hidden}
  .demo-footer{min-height:30px;padding:4px 26px;display:flex;flex-wrap:wrap;gap:2px 24px;align-items:center;justify-content:space-between;color:#64747d;font-size:11px;background:#e5e9eb}
`;
document.head.append(style);

const status = document.createElement('div');
status.className = 'state';
status.setAttribute('role', 'status');

const samplePicker = document.createElement('label');
samplePicker.className = 'picker';
const select = document.createElement('select');
for (const sample of SAMPLES) select.add(new Option(sampleTitle(sample, lang), sample.file));
samplePicker.append(t.sample, select);

const filePicker = document.createElement('label');
filePicker.className = 'picker';
const file = document.createElement('input');
file.type = 'file';
file.accept = '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document';
filePicker.append(t.own, file);

const head = document.createElement('section');
head.className = 'doc-head';
head.append(samplePicker, filePicker, status);

const wrap = document.createElement('div');
wrap.className = 'editor-wrap';
const workspace = document.createElement('main');
workspace.className = 'workspace';
workspace.append(head, wrap);

const footer = document.createElement('footer');
footer.className = 'demo-footer';
for (const text of [t.fictional, lang === 'zh-TW' ? '' : t.editorLanguage, t.local]) {
  if (!text) continue;
  const span = document.createElement('span');
  span.textContent = text;
  footer.append(span);
}

app.append(shellHeader(lang, 'demo', t), workspace, footer);

const editor = createDocxEditor(wrap, {
  onError: (error) => (status.textContent = `${t.openFailed}${t.open}${error instanceof Error ? error.message : error}${t.close}`),
});

async function loadSample(name: string) {
  status.textContent = t.loading;
  try {
    const response = await fetch(sampleUrl(name));
    if (!response.ok) throw new Error(String(response.status));
    const doc = await response.blob();
    // A sample chosen meanwhile wins.
    if (select.value !== name) return;
    editor.open(doc);
    status.textContent = `${t.showingSample}${t.colon}${select.selectedOptions[0].text}`;
  } catch (error) {
    status.textContent = `${t.loadFailed}${t.open}${error instanceof Error ? error.message : error}${t.close}`;
  }
}
select.onchange = () => void loadSample(select.value);
file.onchange = () => {
  const own = file.files?.[0];
  if (own) {
    editor.open(own);
    status.textContent = `${t.showingOwn}${t.colon}${own.name}`;
  }
  file.value = '';
};
void loadSample(select.value);

// public/browser-check.js waits for this: the page could start in this browser.
(window as unknown as { __appStarted?: boolean }).__appStarted = true;
