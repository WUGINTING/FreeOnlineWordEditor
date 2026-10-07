# Papyrus DOCX

[English](README.md) | **繁體中文** | [简体中文](README.zh-CN.md)

[![npm](https://img.shields.io/npm/v/papyrus-docx.svg)](https://www.npmjs.com/package/papyrus-docx)
[![CI](https://github.com/WUGINTING/FreeOnlineWordEditor/actions/workflows/ci.yml/badge.svg)](https://github.com/WUGINTING/FreeOnlineWordEditor/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

免費、開源，在瀏覽器裡開啟、編輯、儲存 Word（.docx）的編輯器，**Word ⇄ 線上編輯無損轉換**。
Vue 3 + ProseMirror，MIT 授權。檔案全程在瀏覽器裡處理，不需要伺服器。

- **線上示範：** <https://wuginting.github.io/FreeOnlineWordEditor/>
- **測試畫面**（所有選項與事件都能試，適合整合前先玩玩看）：
  <https://wuginting.github.io/FreeOnlineWordEditor/playground.html>

這個編輯器是用「無塵室」方式從零寫的：只參考了 SuperDoc 的**公開功能介紹**和
Office Open XML（ECMA-376）公開標準，**沒有**看過或複製 SuperDoc 的原始碼。
紀錄在 [CLEAN_ROOM.md](CLEAN_ROOM.md)。

## 馬上試試看

需要 Node 20 以上。

```bash
npm install
npm run dev
```

打開終端機顯示的網址，就會看到示範頁：可以選一份示範文件，或開啟自己電腦裡的 .docx，
編輯後用工具列的「檔案 › 下載」存回 Word 檔。`/playground.html` 是測試畫面。

## 核心原則：無損轉換

打開 → 編輯 → 存檔，**只有使用者改的地方會變**，其餘部分與原檔逐元素相同：

- 每個段落、文字、表格格子都保留 Word 原本的格式設定原稿（w:pPr / w:rPr / w:tcPr / w:trPr，
  以及 rsid、paraId 等屬性），存檔時只修補使用者改過的那一項。
- 編輯器看不懂的內容（圖表、公式、書籤、欄位代碼、內容控制項、追蹤修訂、註解標記……）
  一律原樣保留。

## 用在自己的專案

三種引入方式，挑一種就好：

| 你的專案 | 用法 |
|---|---|
| Vue 3 專案 | [安裝套件，用 Vue 元件](#1-vue-3-專案) |
| React、Angular、純 JavaScript 等有打包工具的專案 | [安裝套件，用 `createDocxEditor`](#2-其他框架或純-javascript) |
| 沒有打包工具的網頁（JSP、PHP、靜態 HTML……） | [一行 `<script>`](#3-一行-script不需要打包工具) |

### 安裝

```bash
npm install papyrus-docx
```

想用 main 分支上還沒發佈的最新程式，可以直接從 GitHub 安裝（安裝時會自動建置）：

```bash
npm install github:WUGINTING/FreeOnlineWordEditor
```

套件附 TypeScript 型別，不需要另外安裝 `@types`。

### 1. Vue 3 專案

```vue
<script setup lang="ts">
import { ref } from 'vue';
import { DocxEditorVue } from 'papyrus-docx';
import 'papyrus-docx/style.css';

const file = ref<Blob | null>(null); // 要開啟的 .docx；null 則是空白文件
const editor = ref<InstanceType<typeof DocxEditorVue> | null>(null);

async function save() {
  const blob: Blob = await editor.value!.save(); // 編輯後的 .docx
  // 上傳、另存……
}
</script>

<template>
  <div style="height: 80vh">
    <DocxEditorVue ref="editor" :src="file" @error="console.error" />
  </div>
</template>
```

編輯器會填滿外層元素，所以外層要有高度。

| 屬性 | 說明 |
|---|---|
| `src` | 要開啟的 .docx（`Blob`、`ArrayBuffer` 或 `Uint8Array`）；不給就是空白文件 |
| `editable` | 能不能編輯（預設 `true`） |
| `toolbar` | 要不要顯示工具列（預設 `true`） |
| `commenting` | 唯讀，但可以新增、回覆、解決註解 |
| `author` | 新註解的作者 |
| `fileMenu` | 頁面有自己的「檔案」選單時設為 `true`，工具列的「檔案」改為送出 `file` 事件 |

事件：`ready`（編輯器已就緒）、`change`（內容有變）、`error`（檔案打不開）、
`compat`（這份文件裡有哪些內容網頁上無法完整顯示或編輯）、`size`、`fonts`、`file`。
元件上可以呼叫 `save()`、`download(檔名)`、`open(檔案)`。

### 2. 其他框架或純 JavaScript

`createDocxEditor` 把整個編輯器（含工具列）放進你指定的元素，不需要懂 Vue：

```ts
import { createDocxEditor } from 'papyrus-docx';
import 'papyrus-docx/style.css';

const editor = createDocxEditor('#editor', {   // 元素，或 CSS 選擇器
  src: file,                                   // 要開啟的 .docx；不給就是空白文件
  onChange: () => {},
  onError: (error) => console.error(error),
});

editor.open(anotherFile);            // 換一份文件
const blob = await editor.save();    // 編輯後的 .docx
await editor.download('文件.docx');  // 讓瀏覽器下載
editor.destroy();                    // 離開頁面時收掉
```

選項與上表的屬性相同；事件改成 `onReady`、`onChange`、`onError`、`onCompat`、`onSize`、
`onFonts`、`onFile`。`editor.editor` 是編輯器本體，可以用來下指令、讀選取範圍、復原等。

React 的寫法：

```tsx
import { useEffect, useRef } from 'react';
import { createDocxEditor, type DocxEditorHandle } from 'papyrus-docx';
import 'papyrus-docx/style.css';

export function WordEditor({ file }: { file: Blob | null }) {
  const host = useRef<HTMLDivElement>(null);
  const editor = useRef<DocxEditorHandle | null>(null);
  useEffect(() => {
    editor.current = createDocxEditor(host.current!, { src: file });
    return () => editor.current?.destroy();
  }, [file]);
  return <div ref={host} style={{ height: '80vh' }} />;
}
```

### 3. 一行 script（不需要打包工具）

`dist/papyrus-docx.standalone.iife.js` 一個檔案就包含全部（Vue、樣式都在裡面），
從 CDN 載入（如下），或把檔案放到自己的網站，然後用全域的 `PapyrusDocx`：

```html
<div id="editor" style="height: 80vh"></div>
<script src="https://cdn.jsdelivr.net/npm/papyrus-docx@0.1.0/dist/papyrus-docx.standalone.iife.js"></script>
<script>
  var editor = PapyrusDocx.createDocxEditor('#editor');
  // editor.open(file)、editor.save()、editor.download('文件.docx')
</script>
```

完整範例見 [examples/script-tag.html](examples/script-tag.html)（先 `npm run build`，再用瀏覽器打開）。
上面網址裡的 `0.1.0` 是套件的版本：請寫明要用的版本，新版出來時你的頁面才不會跟著變。

偏好 ES 模組的話，用 `papyrus-docx/standalone`（`dist/papyrus-docx.standalone.js`），內容相同。

### 介面語言

介面有繁體中文（`zh-TW`，預設）、簡體中文（`zh-CN`）與英文（`en`）。請在建立編輯器之前選好語言：

```ts
import { setLocale, createDocxEditor } from 'papyrus-docx';

setLocale(navigator.language);   // 'en-US' → en、'zh-Hans-CN' → zh-CN、'zh-HK' → zh-TW，其他 → en
// 或： createDocxEditor('#editor', { locale: 'en' })   ·   <DocxEditorVue locale="en" />
```

語言是整個頁面共用一種，不是每個編輯器各自設定。`setLocale` 會回傳實際使用的語言；`locales()` 列出有哪些語言。

想改某一句，或加上自己的語言，用原文（繁體中文）當索引把文字交給它：
`addMessages('de', { '檔案': 'Datei', '尋找': 'Suchen' })`。某個語言沒有的句子會以繁體中文顯示。
完整的文字清單在 [src/papyrus/locales/en.ts](src/papyrus/locales/en.ts)；`npm run i18n` 會列出每個語言檔還缺哪些。
字型名稱、樣式名稱與文件本身的內容不會被翻譯。

### 只要讀寫檔案，不要畫面

```ts
import { readDocx, writeDocx } from 'papyrus-docx';

const { doc, model } = await readDocx(bytes);   // Uint8Array / ArrayBuffer / Blob
const saved = await writeDocx(doc, model);      // 編輯後的 .docx（Uint8Array）
```

匯出的項目都列在 [src/papyrus/index.ts](src/papyrus/index.ts)。

## 測試

```bash
npm test             # 全部測試
npm run type-check   # 型別檢查
```

拿自己的 Word 檔做「存檔前後逐元素比對」（檔案不會被修改）：

```bash
DOCX_SAMPLES=<放 .docx 的資料夾> npm test
```

`scripts/` 裡的 PowerShell 腳本會用桌面版 Microsoft Word 開啟存檔後的檔案做比對
（需要 Windows 與 Word）。開發期間以 31 份真實文件驗證過：存檔前後逐元素相同、
Microsoft Word 比對 31/31 相同。那些文件不屬於這個儲存庫，所以沒有附上。

## 專案結構

```
src/papyrus/
  docx/reader.ts     .docx 封包 → 編輯器文件（找出各零件、樣式、編號、圖片、頁首頁尾）
  docx/convert.ts    WordprocessingML → 編輯器內容；看不懂的一律原樣保留
  docx/props.ts      段落 / 文字 / 格子格式：保留原稿，只修補改過的項目
  docx/wrappers.ts   內容控制項、超連結、修訂等外框的保存與重建
  docx/writer.ts     編輯器文件 → .docx（以原檔為底，只重寫內文與改過的頁首頁尾）
  docx/styles.ts     styles.xml → CSS
  docx/numbering.ts  清單編號規則與計數
  editor/schema.ts   文件結構定義（ProseMirror schema）
  editor/pagination.ts  量測區塊高度，插入換頁墊片
  editor/core.ts     不綁框架的編輯器主體（含頁首頁尾編輯）
  vue/               Vue 元件與工具列
  mount.ts           createDocxEditor：不用 Vue 也能放進頁面
demo/                示範網站：示範頁與測試畫面（npm run dev）
examples/            一行 script 的用法範例
public/demo-docs/    示範文件（虛構內容）
tests/               測試
scripts/             Microsoft Word 驗證腳本
```

## 已知限制

- 畫面分頁以段落 / 表格為單位；只影響顯示，不影響檔案內容。
- 圖表、公式、SmartArt、註腳內容只能原樣保留，不能在網頁上編輯（文字方塊與常見圖案可以顯示、改字、移動與新增）。
- 追蹤修訂：畫面顯示「接受所有修訂」後的樣子；修訂記錄原樣保留在檔案裡。
- 需要較新的瀏覽器：Chrome / Edge 105、Firefox 121、Safari 15.5 以上。
- Node 22.12 在含中文字的資料夾路徑下會當掉（Node 本身的問題），請放在英文路徑或升級 Node。

## 授權

[MIT](LICENSE)
