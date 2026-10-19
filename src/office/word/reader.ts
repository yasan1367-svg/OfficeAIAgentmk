/// <reference types="office-js" />
/**
 * Office AI Agent - Microsoft Word Reader & Document Inspection (Phase 6)
 */

export interface WordHeadingItem {
  text: string;
  level: number;
  paragraphIndex: number;
}

export interface WordInspectionReport {
  timestamp: number;
  documentTitle: string;
  paragraphCount: number;
  wordCount: number;
  tableCount: number;
  headings: WordHeadingItem[];
  sampleParagraphs: string[];
  styleBreakdown: Record<string, number>;
  proofingIssues: Array<{
    type: 'zwnj' | 'punctuation' | 'repetition';
    description: string;
    originalSnippet: string;
    suggestedSnippet: string;
    paragraphIndex: number;
  }>;
}

/**
 * Inspects active Word document structure, headings, tables and paragraphs
 */
export async function inspectActiveWordDocument(): Promise<WordInspectionReport> {
  if (typeof window !== 'undefined' && (window as unknown as { Word?: typeof Word }).Word) {
    try {
      const WordGlobal = (window as unknown as { Word: typeof Word }).Word;
      return await WordGlobal.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        const tables = body.tables;

        paragraphs.load('items/text, items/style, items/styleBuiltIn');
        tables.load('items/length');

        await context.sync();

        const rawParagraphs = paragraphs.items.map((p) => p.text.trim()).filter(Boolean);
        const totalWords = rawParagraphs.reduce((sum, p) => sum + p.split(/\s+/).length, 0);

        const headings: WordHeadingItem[] = [];
        const styleBreakdown: Record<string, number> = {};

        paragraphs.items.forEach((p, idx) => {
          const style = String(p.style || p.styleBuiltIn || 'Normal');
          styleBreakdown[style] = (styleBreakdown[style] || 0) + 1;

          if (style.includes('Heading 1') || style.includes('Heading1') || style.includes('Title')) {
            headings.push({ text: p.text.trim(), level: 1, paragraphIndex: idx });
          } else if (style.includes('Heading 2') || style.includes('Heading2')) {
            headings.push({ text: p.text.trim(), level: 2, paragraphIndex: idx });
          }
        });

        // Run orthography / proofing scanner
        const proofingIssues = scanPersianProofingIssues(rawParagraphs);

        return {
          timestamp: Date.now(),
          documentTitle: getSafeWordDocTitle(),
          paragraphCount: rawParagraphs.length,
          wordCount: totalWords,
          tableCount: tables.items ? tables.items.length : 0,
          headings,
          sampleParagraphs: rawParagraphs.slice(0, 10),
          styleBreakdown,
          proofingIssues
        };
      });
    } catch (err) {
      console.warn('[WordReader] Error during Word.run inspection, fallback:', err);
    }
  }

  // Standalone mock document for executive business report
  return generateStandaloneWordReport();
}

/**
 * Scans Persian text for common typographical and orthographical issues
 */
export function scanPersianProofingIssues(paragraphs: string[]) {
  const issues: WordInspectionReport['proofingIssues'] = [];

  const patterns = [
    {
      regex: /(می|نمی)\s+([آ-ی]+)/g,
      type: 'zwnj' as const,
      desc: 'عدم استفاده از نیم‌فاصله پس از پیشوند «می / نمی‌»',
      fix: '$1\u200C$2'
    },
    {
      regex: /([آ-ی]+)\s+(ها|های|تر|ترین|ام|ات|اش|مان|تان|شان)(?=\s|[،.؛:])/g,
      type: 'zwnj' as const,
      desc: 'عدم رعایت نیم‌فاصله در پسوندها',
      fix: '$1\u200C$2'
    },
    {
      regex: /\s+([،.؛:!؟])/g,
      type: 'punctuation' as const,
      desc: 'فاصله اضافه قبل از علامت نگارشی',
      fix: '$1'
    },
    {
      regex: /([،.؛:!؟])([آ-یA-Za-z0-9])/g,
      type: 'punctuation' as const,
      desc: 'عدم وجود فاصله استاندارد بعد از علامت نگارشی',
      fix: '$1 $2'
    }
  ];

  paragraphs.forEach((pText, pIdx) => {
    for (const rule of patterns) {
      let match;
      const re = new RegExp(rule.regex);
      while ((match = re.exec(pText)) !== null) {
        if (issues.length >= 10) break;
        const originalSnippet = match[0];
        const suggestedSnippet = originalSnippet.replace(rule.regex, rule.fix);

        if (originalSnippet !== suggestedSnippet) {
          issues.push({
            type: rule.type,
            description: rule.desc,
            originalSnippet,
            suggestedSnippet,
            paragraphIndex: pIdx + 1
          });
        }
      }
    }
  });

  return issues;
}

function getSafeWordDocTitle(): string {
  if (typeof window !== 'undefined') {
    const officeGlobal = (window as unknown as { Office?: { context?: { document?: { url?: string } } } }).Office;
    const url = officeGlobal?.context?.document?.url;
    if (url) {
      const parts = url.split(/[\\/]/);
      return parts[parts.length - 1];
    }
  }
  return 'گزارش_رسمی_مدیریتی_۱۴۰۳.docx';
}

function generateStandaloneWordReport(): WordInspectionReport {
  const sampleParagraphs = [
    'گزارش رسمی عملکرد و تحلیل هزینه‌های فصلی شرکت',
    'مقدمه و اهداف راهبردی',
    'در راستای بهبود فرآیندهای مالی و نظارت بر هزینه‌ها ،این گزارش به بررسی عملکرد سه ماهه می‌ پردازد. بر اساس داده های استخراج شده از سیستم حسابداری، روند مصارف با بودجه مصوب تطبیق داده شده است.',
    'یافته‌های کلیدی و تحلیل شاخص‌ها',
    'مجموع تعهدات مالی در این دوره با کاهش ۸ درصدی در بخش هزینه‌های جاری همراه بوده است که نشان دهنده کارآمدی اقدامات اصلاحی می‌باشد.',
    'برنامه‌های بهبود و پیشنهادهای آتی',
    'پیشنهاد می شود برای دوره‌های آینده، سامانه‌های خودکار کنترل بودجه مستقر گردند تا از بروز مغایرت‌های مالی جلوگیری به عمل آید.'
  ];

  const headings: WordHeadingItem[] = [
    { text: 'گزارش رسمی عملکرد و تحلیل هزینه‌های فصلی شرکت', level: 1, paragraphIndex: 0 },
    { text: 'مقدمه و اهداف راهبردی', level: 2, paragraphIndex: 1 },
    { text: 'یافته‌های کلیدی و تحلیل شاخص‌ها', level: 2, paragraphIndex: 3 },
    { text: 'برنامه‌های بهبود و پیشنهادهای آتی', level: 2, paragraphIndex: 5 }
  ];

  return {
    timestamp: Date.now(),
    documentTitle: 'گزارش_رسمی_مدیریتی_۱۴۰۳.docx',
    paragraphCount: 7,
    wordCount: 118,
    tableCount: 1,
    headings,
    sampleParagraphs,
    styleBreakdown: { Title: 1, 'Heading 2': 3, Normal: 3 },
    proofingIssues: [
      {
        type: 'zwnj',
        description: 'عدم استفاده از نیم‌فاصله پس از پیشوند «می»',
        originalSnippet: 'می‌ پردازد',
        suggestedSnippet: 'می‌پردازد',
        paragraphIndex: 3
      },
      {
        type: 'zwnj',
        description: 'عدم رعایت نیم‌فاصله در پسوند «ها»',
        originalSnippet: 'داده های',
        suggestedSnippet: 'داده‌های',
        paragraphIndex: 3
      },
      {
        type: 'punctuation',
        description: 'فاصله اضافه قبل از ویرگول',
        originalSnippet: 'هزینه‌ها ،این',
        suggestedSnippet: 'هزینه‌ها، این',
        paragraphIndex: 3
      },
      {
        type: 'zwnj',
        description: 'عدم رعایت نیم‌فاصله در ترکیب «نشان‌دهنده»',
        originalSnippet: 'نشان دهنده',
        suggestedSnippet: 'نشان‌دهنده',
        paragraphIndex: 5
      },
      {
        type: 'zwnj',
        description: 'عدم رعایت نیم‌فاصله در پیشوند «می‌شود»',
        originalSnippet: 'پیشنهاد می شود',
        suggestedSnippet: 'پیشنهاد می‌شود',
        paragraphIndex: 7
      }
    ]
  };
}
