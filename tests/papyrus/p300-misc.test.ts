// persona-300 13–15: the page setup help names where the buttons are now; :has() rules stand on
// their own (a browser without :has() drops only them); print() names the printout as asked.
import { describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { domStubs, setup } from './p300Helpers';

vi.mock('../../src/papyrus/editor/print', async (orig) => ({ ...(await orig<object>()), printPages: vi.fn() }));
domStubs();

describe('help text and CSS', () => {
  it('PageSetupDialog points to 版面配置 › 分隔設定, not old toolbar buttons', () => {
    const src = readFileSync('src/papyrus/vue/PageSetupDialog.vue', 'utf8');
    expect(src).not.toMatch(/工具列「分節」|工具列「分欄符號」/);
    expect(src).toContain('版面配置 › 分隔設定 › 分節符號（下一頁）');
    expect(src).toContain('版面配置 › 分隔設定 › 分欄符號');
  });

  it('every CSS rule with :has() has no other selector in it', () => {
    for (const file of ['src/papyrus/editor/editor.css', 'src/papyrus/editor/features.css']) {
      const css = readFileSync(file, 'utf8').replace(/\/\*[\s\S]*?\*\//g, '');
      for (const m of css.matchAll(/([^{}]+)\{/g)) {
        // Commas outside brackets separate selectors (not those of :is(td, th)).
        const selectors: string[] = [];
        let depth = 0;
        let cur = '';
        for (const ch of m[1]) {
          if (ch === '(') depth++;
          if (ch === ')') depth--;
          if (ch === ',' && !depth) {
            selectors.push(cur.trim());
            cur = '';
          } else cur += ch;
        }
        if (cur.trim()) selectors.push(cur.trim());
        if (selectors.some((s) => s.includes(':has('))) expect(selectors, `${file}: ${m[1].trim()}`).toHaveLength(1);
      }
    }
  });
});

describe('print title', () => {
  it('uses the title the host gives, else the page title', async () => {
    const { printPages } = await import('../../src/papyrus/editor/print');
    const d = await setup('<w:p><w:r><w:t>內容</w:t></w:r></w:p>');
    document.title = '線上文件';
    d.editor.print('115年度預算說明');
    d.editor.print();
    d.editor.print('  ');
    expect((printPages as any).mock.calls.map((c: unknown[]) => c[2])).toEqual(['115年度預算說明', '線上文件', '線上文件']);
    d.done();
  });
});
