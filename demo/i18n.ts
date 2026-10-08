// The demo pages' own words in three languages; the editor's interface follows the same choice.
import { setLocale } from '../src/papyrus';

export type Lang = 'en' | 'zh-TW' | 'zh-CN';
export const LANGS: { id: Lang; label: string }[] = [
  { id: 'en', label: 'English' },
  { id: 'zh-TW', label: '繁體中文' },
  { id: 'zh-CN', label: '简体中文' },
];
const KEY = 'papyrus-demo-lang';

const en = {
  tagline: 'Word documents in the browser',
  /** Between a label and what it names; around a detail. */
  colon: ': ',
  open: ' (',
  close: ')',
  demo: 'Demo',
  playground: 'Playground',
  language: 'Language',
  sample: 'Sample document',
  own: 'or open your own .docx',
  loading: 'Loading the document',
  loadFailed: 'The document could not be loaded',
  openFailed: 'This file could not be opened',
  showingSample: 'Sample',
  showingOwn: 'Your file',
  fictional: 'The sample documents are fictional (in Traditional Chinese)',
  local: 'Files stay in your browser; nothing is uploaded',
  // playground
  options: 'Options',
  optionsNote: 'Changing an option starts the editor again with the same document.',
  actions: 'Actions',
  openFile: 'Open a .docx',
  blank: 'New empty document',
  save: 'save()',
  download: 'download()',
  roundTrip: 'Save and open again',
  destroy: 'destroy()',
  create: 'createDocxEditor()',
  events: 'Events',
  clear: 'Clear',
  destroyed: 'The editor is not on the page. Press createDocxEditor().',
  saved: 'saved',
  reopened: 'saved and opened again',
  bytes: 'bytes',
};
type Words = typeof en;

const zhTW: Words = {
  tagline: '線上 Word 文件編輯',
  colon: '：',
  open: '（',
  close: '）',
  demo: '示範',
  playground: '測試畫面',
  language: '語言',
  sample: '示範文件',
  own: '或開啟自己的 .docx',
  loading: '載入文件中',
  loadFailed: '文件載入失敗',
  openFailed: '這個檔案打不開',
  showingSample: '示範文件',
  showingOwn: '您的檔案',
  fictional: '示範文件為虛構內容，不代表任何機關的實際公文',
  local: '檔案只在您的瀏覽器裡處理，不會上傳',
  options: '選項',
  optionsNote: '改變選項會用同一份文件重新建立編輯器。',
  actions: '操作',
  openFile: '開啟 .docx',
  blank: '新的空白文件',
  save: 'save()',
  download: 'download()',
  roundTrip: '存檔後重新開啟',
  destroy: 'destroy()',
  create: 'createDocxEditor()',
  events: '事件',
  clear: '清除',
  destroyed: '編輯器不在頁面上，請按 createDocxEditor()。',
  saved: '已存檔',
  reopened: '已存檔並重新開啟',
  bytes: '位元組',
};

const zhCN: Words = {
  tagline: '在线 Word 文档编辑',
  colon: '：',
  open: '（',
  close: '）',
  demo: '演示',
  playground: '测试页面',
  language: '语言',
  sample: '示例文档',
  own: '或打开自己的 .docx',
  loading: '正在加载文档',
  loadFailed: '文档加载失败',
  openFailed: '这个文件无法打开',
  showingSample: '示例文档',
  showingOwn: '您的文件',
  fictional: '示例文档为虚构内容（繁体中文），不代表任何机关的实际公文',
  local: '文件只在您的浏览器里处理，不会上传',
  options: '选项',
  optionsNote: '更改选项会用同一份文档重新创建编辑器。',
  actions: '操作',
  openFile: '打开 .docx',
  blank: '新建空白文档',
  save: 'save()',
  download: 'download()',
  roundTrip: '保存后重新打开',
  destroy: 'destroy()',
  create: 'createDocxEditor()',
  events: '事件',
  clear: '清除',
  destroyed: '编辑器不在页面上，请按 createDocxEditor()。',
  saved: '已保存',
  reopened: '已保存并重新打开',
  bytes: '字节',
};

const WORDS: Record<Lang, Words> = { en, 'zh-TW': zhTW, 'zh-CN': zhCN };

/** The sample documents (public/demo-docs) and their titles in English; the file name carries the Chinese one. */
export const SAMPLES: { file: string; en: string }[] = [
  { file: '01_民眾陳情案件回復函.docx', en: "Reply to a citizen's petition" },
  { file: '02_行政業務簽呈.docx', en: 'Internal memo for approval' },
  { file: '03_跨單位會議通知.docx', en: 'Notice of an inter-office meeting' },
  { file: '04_專案工作會議紀錄.docx', en: 'Minutes of a project meeting' },
  { file: '05_服務措施調整公告.docx', en: 'Public notice of a service change' },
  { file: '06_年度業務執行計畫.docx', en: 'Annual work plan' },
  { file: '07_公共服務改善進度報告.docx', en: 'Progress report on service improvement' },
  { file: '08_年度政策推動成果報告.docx', en: 'Annual policy results report' },
  { file: '09_跨機關資料協作函.docx', en: 'Letter on inter-agency data cooperation' },
  { file: '10_資訊服務採購需求說明.docx', en: 'Requirements for an IT service procurement' },
];

/** A sample's title as the page's language shows it. */
export function sampleTitle(sample: { file: string; en: string }, lang: Lang): string {
  return lang === 'en' ? sample.en : sample.file.replace(/^\d+_/, '').replace(/\.docx$/, '');
}

/** The address of a sample document, wherever the site is served from. */
export const sampleUrl = (file: string) => `${import.meta.env.BASE_URL}demo-docs/${encodeURIComponent(file)}`;

function known(value: string | null | undefined): Lang | null {
  return LANGS.some((l) => l.id === value) ? (value as Lang) : null;
}

/** The page's language: ?lang= in the address, else the one chosen before, else the browser's. */
export function currentLang(): Lang {
  const asked = known(new URLSearchParams(location.search).get('lang'));
  if (asked) return asked;
  try {
    const stored = known(localStorage.getItem(KEY));
    if (stored) return stored;
  } catch {
    // no storage (private window): the browser's language
  }
  const browser = navigator.language || '';
  if (/^zh\b/i.test(browser)) return /-(CN|SG|Hans)\b/i.test(browser) ? 'zh-CN' : 'zh-TW';
  return 'en';
}

/** Remembers the language and shows the page in it. */
export function chooseLang(lang: Lang): void {
  try {
    localStorage.setItem(KEY, lang);
  } catch {
    // not remembered: the address still says it
  }
  const url = new URL(location.href);
  url.searchParams.set('lang', lang);
  location.href = url.toString();
}

/** The words of the page's language. Also marks the page as being in it. */
export function words(lang: Lang): Words {
  setLocale(lang);
  document.documentElement.lang =lang === 'en' ? 'en' : lang === 'zh-CN' ? 'zh-Hans-CN' : 'zh-Hant-TW';
  return WORDS[lang];
}

/** The pages' shared header styles. */
export const SHELL_CSS = `
  *{box-sizing:border-box} html,body,#app{margin:0;width:100%;height:100%;font-family:"Microsoft JhengHei","PingFang TC",system-ui,sans-serif;color:#19232c}
  .demo-header{min-height:58px;background:#17364c;color:#fff;display:flex;flex-wrap:wrap;align-items:center;padding:6px 28px;gap:6px 28px;border-bottom:3px solid #38a78e}
  .brand{font-weight:700;letter-spacing:.5px;font-size:19px}.brand span{font-weight:400;margin-left:14px;font-size:15px;color:#d5e6e8}
  .demo-nav{display:flex;gap:4px}.demo-nav a{color:#d5e6e8;text-decoration:none;font-size:14px;padding:5px 12px;border-radius:4px}
  .demo-nav a:hover{background:#ffffff1f}.demo-nav a[aria-current=page]{background:#ffffff2b;color:#fff;font-weight:600}
  .demo-lang{margin-left:auto;display:flex;align-items:center;gap:8px;font-size:13px;color:#d5e6e8}
  .demo-lang select{font:inherit;padding:3px 6px;border-radius:4px;border:1px solid #6d8797;background:#fff;color:#1d2c36}
  .state{font-size:13px;color:#b5e7db}
`;

/** The header both pages share: name, the two pages, the language. */
export function shellHeader(lang: Lang, page: 'demo' | 'playground', t: Words): HTMLElement {
  const header = document.createElement('header');
  header.className = 'demo-header';
  const brand = document.createElement('div');
  brand.className = 'brand';
  brand.append('Papyrus DOCX');
  const tagline = document.createElement('span');
  tagline.textContent = t.tagline;
  brand.append(tagline);

  const nav = document.createElement('nav');
  nav.className = 'demo-nav';
  for (const [id, href, text] of [['demo', './', t.demo], ['playground', './playground.html', t.playground]] as const) {
    const a = document.createElement('a');
    a.href = `${href}?lang=${lang}`;
    a.textContent = text;
    if (id === page) a.setAttribute('aria-current', 'page');
    nav.append(a);
  }
  const source = document.createElement('a');
  source.href = 'https://github.com/WUGINTING/FreeOnlineWordEditor';
  source.textContent = 'GitHub';
  nav.append(source);

  const picker = document.createElement('label');
  picker.className = 'demo-lang';
  picker.append(t.language);
  const select = document.createElement('select');
  for (const l of LANGS) select.add(new Option(l.label, l.id, false, l.id === lang));
  select.onchange = () => chooseLang(select.value as Lang);
  picker.append(select);

  header.append(brand, nav, picker);
  return header;
}
