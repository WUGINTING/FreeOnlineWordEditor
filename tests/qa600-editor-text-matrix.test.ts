import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readDocx } from '../src/papyrus/docx/reader';
import { writeDocx } from '../src/papyrus/docx/writer';
import { blankPackage } from '../src/papyrus/docx/template';
import { content } from './papyrus/helpers';

// Two hundred distinct everyday text inputs, each opened, saved, and reopened as DOCX.
// Interaction behaviors are covered by the existing focused tests listed in the inventory.
const groups: Array<[string, string[]]> = [
  ['English office prose', ['Meeting notes', 'Project update', 'Please review', 'Action required', 'For your information', 'Draft proposal', 'Quarterly report', 'Customer response', 'Approval request', 'Next steps']],
  ['Traditional Chinese office prose', ['會議紀錄', '專案進度', '請協助確認', '敬請核示', '供您參考', '提案草稿', '季度報告', '客戶回覆', '核准申請', '後續辦理']],
  ['mixed Chinese and English', ['預計 Monday 完成', '請 review 附件', 'Project 進度更新', '會議安排 via Teams', '請回覆 status', 'Budget 尚待確認', 'Draft 已完成', 'Customer 意見整理', 'Deadline 為 Friday', 'Action items 如下']],
  ['punctuation', ['逗號，句號。', '冒號：分號；', '引號「中文」與“英文”', '括號（內文）(text)', '破折號——連字號-', '問句？驚嘆！', '省略號……', '斜線／反斜線\\', '百分比 25%，金額 NT$1,200', '項目 A、B、C；備註：無']],
  ['whitespace and paragraphs', ['前後空格  保留', '包含	定位符', '含兩個  空白', '段首文字', '尾端文字 ', '單字間隔 A B C', '全形　空格', '混用 tab	空格 ', '一行短文。另一句。', '空白段落前後']],
  ['numbers and dates', ['2026/09/25', '2026-09-25 09:30', '民國 115 年 9 月 25 日', '1234567890', '負數 -42.5', '比例 3/5', '第 1.2.3 節', '版本 v2.0.1', '電話 02-1234-5678', '統編 12345678']],
  ['currency and amounts', ['NT$ 1,250', 'US$ 3,400.50', '€ 99.95', '¥ 1200', '新台幣壹萬元整', '含稅 5%', '未稅 18,000 元', '預算上限 250 萬', '退款 −300 元', '金額：零元整']],
  ['special Unicode', ['Cafe\u0301', 'café', 'Ångström', 'naïve résumé', '中文𠮷字', 'emoji 😀', '家庭 👨‍👩‍👧‍👦', '旗幟 🇹🇼', '變化符號 ❤️', '數學 ∑≤∞']],
  ['Japanese and Korean', ['お知らせ', '会議の議題', '資料をご確認ください', '請求書を送付します', '안녕하세요', '회의 일정 안내', '검토 부탁드립니다', '고객 문의 답변', '한글과 中文', '東京 서울 台北']],
  ['Arabic and Hebrew', ['مرحبا بالعالم', 'تقرير الاجتماع', 'يرجى المراجعة', 'שלום עולם', 'סיכום פגישה', 'English العربية mixed', 'עברית English mixed', 'رقم الطلب 12345', 'מספר הזמנה 54321', 'RTL punctuation (نص)']],
  ['Thai and Vietnamese', ['สวัสดีครับ', 'บันทึกการประชุม', 'กรุณาตรวจสอบ', 'Xin chào', 'Biên bản cuộc họp', 'Vui lòng xác nhận', 'ภาษาไทย English', 'Tiếng Việt 中文', 'Đặng Thị Ánh', 'ข้อมูลลูกค้า']],
  ['XML-sensitive text', ['A & B', '<標題>', 'A > B', 'A < B', 'R&D 計畫', 'AT&T 回覆', '輸入 <tag> 文字', '條件 x <= y', '文字 & 符號', '報價 < NT$500']],
  ['office identifiers', ['案件編號 DOC-1001', 'PO-2026-0098', '客戶代碼 CUST-88', '工單 #WK-502', '文件編號 HR-F-012', '專案代碼 PRJ_2026_A', '帳號 user.name@example.com', '網址 https://example.com/a?b=1&c=2', '版本 1.0-beta', '路徑 A/B/C']],
  ['short and long forms', ['好', 'OK', 'No.', '已閱。', '收到，謝謝。', '請查收附件並回覆。', '請於期限前完成檢查並更新追蹤表。', '經討論後決議維持原方案，待資料補齊後再行確認。', '本次會議確認需求範圍、交付日期及後續負責窗口。', '若有問題，請直接回覆此信並附上相關文件供團隊查核。']],
  ['repeated and similar words', ['報告 報告 報告', 'test Test TEST', '文件文件文件', 'A AA AAA AAAA', '同名同名不同段', '改版 改版完成', 'review reviewer reviewed', '計畫與計畫書', '付款與付費', '資料資料；資料。']],
  ['email and messaging', ['主旨：下週會議', 'Hi Alex, please advise.', 'Dear team, 請確認。', 'CC: sales@example.com', '回覆：Re: 報價單', '請勿直接轉寄 (Internal)', 'Thanks & regards', '已收到您的訊息，謝謝！', '附件 2 份，請查收。', '聯絡窗口：王小姐 ext. 123']],
  ['lists as ordinary text', ['1. 第一項', '2) 第二項', '(a) Subtask', '• 重點一', '－ 備註一', '一、總則', '（一）適用範圍', 'A. Option one', '※ 注意事項', '① 項目一']],
  ['line-like content', ['日期：\n地點：', '姓名：王小明\n部門：行政', '第一行\n第二行', '上方\n\n下方', 'To: Alice\nFrom: Bob', '欄位 A: 1\n欄位 B: 2', '標題\n內文', '開始\n中間\n結束', '中文\nEnglish\n日本語', '說明：\n\t縮排文字']],
  ['formal wording', ['茲通知各單位配合辦理。', '敬請查照並於期限內回覆。', '本案奉核後據以執行。', '附件如後，請參閱。', '如有疑義，請洽承辦人。', '請依規定完成簽核程序。', '會議結論如附件所示。', '以上報告，敬請核示。', '請於三個工作日內提供資料。', '本函副知相關單位知悉。']],
  ['technical office terms', ['API 連線正常', 'SLA 99.9%', 'CPU 使用率 72%', 'JSON 欄位 id/name', '資料庫備份完成', '登入 MFA 已啟用', 'HTTPS 憑證到期日', 'UTF-8 編碼測試', 'HTTP 404 / 500', 'CSV 匯入 1,024 筆']],
  ['mixed scripts and symbols', ['中文ABC123！', 'A中B文C', 'αβγ Δ', '温度 23°C', '電阻 10 Ω', '箭頭 → 下一步', '核取 ☑ 完成', '星號 ★ 重點', '分隔 | 欄位 | 內容', '商標 ™ 與版權 ©']],
];

const escapeXml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const encodeRuns = (s: string) => s.split(/([\t\n])/).map((part) =>
  part === '\t' ? '<w:r><w:tab/></w:r>' : part === '\n' ? '<w:r><w:br/></w:r>' :
    part ? `<w:r><w:t xml:space="preserve">${escapeXml(part)}</w:t></w:r>` : '',
).join('');
const cases = groups.flatMap(([group, inputs]) => inputs.map((input) => ({ group, input }))).slice(0, 200);

describe('QA-600 editor text stability variants', () => {
  it.each(cases)('$group: $input', async ({ input }) => {
    const zip = blankPackage();
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${encodeRuns(input)}</w:p><w:sectPr/></w:body></w:document>`);
    const original = await readDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(original.doc.textContent).toBe(input.replace(/[\t\n]/g, ''));
    const saved = await writeDocx(original.doc, original.model);
    const reopened = await readDocx(saved);
    expect(reopened.doc.textContent).toBe(input.replace(/[\t\n]/g, ''));
    expect(content(reopened.doc.toJSON())).toEqual(content(original.doc.toJSON()));
  });
});
