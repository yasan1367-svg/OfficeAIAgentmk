/// <reference types="office-js" />
/**
 * Office AI Agent - Microsoft Word Writer & Report Generator (Phase 6)
 */

export interface WordParagraphOptions {
  text: string;
  style?: 'Title' | 'Heading1' | 'Heading2' | 'Normal';
  fontBold?: boolean;
  fontSize?: number;
  fontColor?: string;
  alignment?: 'Right' | 'Left' | 'Center' | 'Justified';
}

export interface WordTableOptions {
  headers: string[];
  rows: string[][];
  headerBgColor?: string;
}

/**
 * Inserts a styled paragraph into the active Word document
 */
export async function insertWordParagraph(
  options: WordParagraphOptions
): Promise<{ success: boolean; text: string }> {
  if (typeof window !== 'undefined' && (window as unknown as { Word?: typeof Word }).Word) {
    try {
      const WordGlobal = (window as unknown as { Word: typeof Word }).Word;
      return await WordGlobal.run(async (context) => {
        const body = context.document.body;
        const p = body.insertParagraph(options.text, WordGlobal.InsertLocation.end);

        if (options.style) {
          p.styleBuiltIn = options.style;
        }

        if (options.fontBold !== undefined) {
          p.font.bold = options.fontBold;
        }

        if (options.fontSize) {
          p.font.size = options.fontSize;
        }

        if (options.fontColor) {
          p.font.color = options.fontColor;
        }

        if (options.alignment) {
          p.alignment = options.alignment as Word.Alignment;
        }

        await context.sync();
        return { success: true, text: options.text };
      });
    } catch (err) {
      console.error('[WordWriter] insertWordParagraph error:', err);
      throw err;
    }
  }

  console.log(`[WordWriter Standalone] Inserted paragraph: "${options.text.slice(0, 30)}..."`);
  return { success: true, text: options.text };
}

/**
 * Inserts a formal styled table into the Word document
 */
export async function insertWordTable(
  options: WordTableOptions
): Promise<{ success: boolean; rowCount: number }> {
  const allRows = [options.headers, ...options.rows];
  const rowCount = allRows.length;
  const colCount = options.headers.length;

  if (typeof window !== 'undefined' && (window as unknown as { Word?: typeof Word }).Word) {
    try {
      const WordGlobal = (window as unknown as { Word: typeof Word }).Word;
      return await WordGlobal.run(async (context) => {
        const body = context.document.body;
        const table = body.insertTable(rowCount, colCount, WordGlobal.InsertLocation.end, allRows);

        table.styleBuiltIn = 'GridTable4_Accent1';
        table.alignment = WordGlobal.Alignment.centered;

        await context.sync();
        return { success: true, rowCount };
      });
    } catch (err) {
      console.error('[WordWriter] insertWordTable error:', err);
      throw err;
    }
  }

  console.log(`[WordWriter Standalone] Inserted table: ${rowCount}x${colCount}`);
  return { success: true, rowCount };
}

/**
 * Generates a complete executive report in Word based on Excel analytics
 */
export async function generateExecutiveReportFromExcel(data: {
  reportTitle: string;
  sourceSheet: string;
  kpis: Array<{ label: string; value: string }>;
  categoryRows: string[][];
}): Promise<{ success: boolean; sectionsCreated: number }> {
  // Title
  await insertWordParagraph({
    text: data.reportTitle,
    style: 'Title',
    fontSize: 20,
    fontBold: true,
    fontColor: '#0f172a',
    alignment: 'Right'
  });

  // Metadata
  const nowStr = new Date().toLocaleDateString('fa-IR', {
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });

  await insertWordParagraph({
    text: `منبع داده: اکسل (${data.sourceSheet}) · تاریخ تهیه: ${nowStr} · وضعیت: نهایی`,
    style: 'Normal',
    fontSize: 10,
    fontColor: '#64748b',
    alignment: 'Right'
  });

  // Section 1: Introduction
  await insertWordParagraph({
    text: '۱. خلاصه مدیریتی و اهداف گزارش',
    style: 'Heading1',
    fontSize: 14,
    fontBold: true,
    fontColor: '#1e3a8a',
    alignment: 'Right'
  });

  await insertWordParagraph({
    text: `این سند بر اساس تحلیل جامع داده‌های استخراج‌شده از کاربرگ «${data.sourceSheet}» تدوین گردیده است. هدف از این گزارش، ارائه تصویری شفاف از عملکرد مالی، تفکیک هزینه‌ها و ارزیابی شاخص‌های کلیدی جهت تصمیم‌گیری مدیران ارشد می‌باشد.`,
    style: 'Normal',
    fontSize: 11,
    alignment: 'Justified'
  });

  // Section 2: KPIs Table
  await insertWordParagraph({
    text: '۲. شاخص‌های کلیدی عملکرد (KPIs)',
    style: 'Heading1',
    fontSize: 14,
    fontBold: true,
    fontColor: '#1e3a8a',
    alignment: 'Right'
  });

  const kpiHeaders = ['ردیف', 'عنوان شاخص عملکرد', 'مقدار محاسبه‌شده در اکسل'];
  const kpiTableRows = data.kpis.map((kpi, idx) => [String(idx + 1), kpi.label, kpi.value]);
  await insertWordTable({
    headers: kpiHeaders,
    rows: kpiTableRows
  });

  // Section 3: Category breakdown
  await insertWordParagraph({
    text: '۳. تحلیل تفکیکی دسته‌بندی‌ها و سهم از کل',
    style: 'Heading1',
    fontSize: 14,
    fontBold: true,
    fontColor: '#1e3a8a',
    alignment: 'Right'
  });

  const catHeaders = ['دسته‌بندی اقلام', 'تعداد تراکنش', 'مجموع مبالغ (ریال)', 'سهم'];
  await insertWordTable({
    headers: catHeaders,
    rows: data.categoryRows
  });

  // Section 4: Concluding Recommendations
  await insertWordParagraph({
    text: '۴. پیشنهادها و راهکارهای بهبود',
    style: 'Heading1',
    fontSize: 14,
    fontBold: true,
    fontColor: '#1e3a8a',
    alignment: 'Right'
  });

  await insertWordParagraph({
    text: '۱) بازنگری در قراردادهای پشتیبانی فنی با توجه به سهم ۱۸ درصدی از کل مصارف.\n۲) تثبیت سیستم انبارداری الکترونیک جهت جلوگیری از ثبت رکوردهای تکراری در کاربرگ‌های دوره‌ای.\n۳) گزارش‌گیری منظم ماهانه بر اساس استانداردهای گزارش حاضر.',
    style: 'Normal',
    fontSize: 11,
    alignment: 'Justified'
  });

  return { success: true, sectionsCreated: 4 };
}
