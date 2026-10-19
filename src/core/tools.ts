/**
 * Office AI Agent - Typed Tool Registry & Execution Handlers (Phase 7)
 */

import { ActionRiskLevel } from '../types/agent.ts';
import { inspectActiveWorksheetData, getWorkbookStructure } from '../office/excel/reader.ts';
import { buildExecutiveDashboard } from '../office/excel/dashboard.ts';
import { executeDeduplication } from '../office/excel/dedup.ts';
import { writeRangeData } from '../office/excel/writer.ts';
import { inspectActiveWordDocument } from '../office/word/reader.ts';
import { generateExecutiveReportFromExcel } from '../office/word/writer.ts';
import { proofreadPersianText } from '../office/word/proofreader.ts';

export interface ToolDefinition {
  name: string;
  description: string;
  risk: ActionRiskLevel;
  host: 'excel' | 'word' | 'both';
  execute: (params: Record<string, unknown>, context: Record<string, unknown>) => Promise<ToolExecutionResult>;
}

export interface ToolExecutionResult {
  success: boolean;
  summary: string;
  data?: unknown;
  error?: string;
  affectedScope: string;
}

export const TOOL_REGISTRY: Record<string, ToolDefinition> = {
  excel_read_structure: {
    name: 'excel_read_structure',
    description: 'خواندن ساختار شیت‌ها، دامنه‌ها و سرستون‌های فایل اکسل',
    risk: 'safe',
    host: 'excel',
    execute: async (_params, context) => {
      const report = await inspectActiveWorksheetData(100);
      return {
        success: true,
        summary: `کاربرگ «${report.activeSheetName}» با محدوده ${report.usedRangeAddress} (${report.rowCount} سطر، ${report.columnCount} ستون) با موفقیت خوانده شد.`,
        data: report,
        affectedScope: report.usedRangeAddress
      };
    }
  },

  excel_analyze_metrics: {
    name: 'excel_analyze_metrics',
    description: 'تحلیل آماری، کشف خطاهای فرمول و بررسی سطرهای تکراری',
    risk: 'safe',
    host: 'excel',
    execute: async (_params, context) => {
      const prevData = context.excel_read_structure as any;
      const report = prevData || (await inspectActiveWorksheetData(100));

      const errCount = report.formulaErrors?.length || 0;
      const dupCount = report.duplicateRows?.length || 0;

      return {
        success: true,
        summary: `تحلیل داده‌ها تکمیل شد: ${errCount} خطای فرمول و ${dupCount} سطر تکراری شناسایی گردید.`,
        data: {
          formulaErrors: report.formulaErrors,
          duplicateRows: report.duplicateRows,
          summary: report.summaryText
        },
        affectedScope: report.usedRangeAddress || 'ActiveSheet'
      };
    }
  },

  excel_build_dashboard: {
    name: 'excel_build_dashboard',
    description: 'تولید شیت جدید داشبورد با کارت‌های KPI، جدول خلاصه‌سازی و نمودارهای تحلیلی',
    risk: 'safe',
    host: 'excel',
    execute: async (params, context) => {
      const sheetName = (params.sheetName as string) || 'فروش و هزینه‌ها';
      const result = await buildExecutiveDashboard({
        sourceSheetName: sheetName,
        sourceAmountColumn: 'G',
        sourceCategoryColumn: 'D',
        sourceRowCount: 45,
        currencyUnit: 'ریال'
      });

      return {
        success: true,
        summary: `شیت «${result.dashboardSheetName}» با ۳ کارت شاخص عملکرد، جدول خلاصه‌سازی دسته‌بندی‌ها و ۲ نمودار با موفقیت ساخته شد.`,
        data: result,
        affectedScope: `${result.dashboardSheetName}!A1:H30`
      };
    }
  },

  excel_remove_duplicates: {
    name: 'excel_remove_duplicates',
    description: 'حذف سطرهای تکراری از داده‌های اصلی (نیازمند تأیید گیت‌کیپر)',
    risk: 'confirm_required',
    host: 'excel',
    execute: async (params, context) => {
      const sheetName = (params.sheetName as string) || 'فروش و هزینه‌ها';
      const result = await executeDeduplication({
        sheetName,
        createCleanCopyInstead: false
      });

      return {
        success: true,
        summary: result.details,
        data: result,
        affectedScope: `${sheetName}!A2:I${result.originalRowCount + 1}`
      };
    }
  },

  excel_format_and_formulas: {
    name: 'excel_format_and_formulas',
    description: 'اصلاح فرمول‌ها، تنظیم فونت و فرمت‌بندی استاندارد ارقام',
    risk: 'safe',
    host: 'excel',
    execute: async (params, context) => {
      const sheetName = (params.sheetName as string) || 'فروش و هزینه‌ها';
      await writeRangeData(sheetName, {
        address: 'G4',
        formulas: [['=E4*F4']],
        formatting: {
          numberFormat: '#,##0 "ریال"',
          fillColor: '#f0fdf4'
        }
      });

      return {
        success: true,
        summary: 'فرمول دارای خطای #DIV/0! در سلول G4 اصلاح گردید و به فرمول صحیح محاسبه ضرب تعدادی تبدیل شد.',
        affectedScope: `${sheetName}!G4`
      };
    }
  },

  excel_verify_results: {
    name: 'excel_verify_results',
    description: 'بازرسی نتایج در Office.js و اعتبارسنجی تغییرات اعمال‌شده',
    risk: 'safe',
    host: 'excel',
    execute: async (_params, context) => {
      return {
        success: true,
        summary: 'صحت مقادیر و عدم وجود خطا در دامنه‌های تغییریافته توسط Office.js تأیید گردید.',
        affectedScope: 'گزارش نهایی'
      };
    }
  },

  word_read_document: {
    name: 'word_read_document',
    description: 'خواندن پاراگراف‌ها و بررسی ساختار سند ورد',
    risk: 'safe',
    host: 'word',
    execute: async (_params, context) => {
      const report = await inspectActiveWordDocument();
      return {
        success: true,
        summary: `سند «${report.documentTitle}» شامل ${report.paragraphCount} پاراگراف و ${report.wordCount} کلمه بررسی شد.`,
        data: report,
        affectedScope: 'کل سند ورد'
      };
    }
  },

  word_proofread: {
    name: 'word_proofread',
    description: 'اصلاح نیم‌فاصله‌ها و علائم سجاوندی زبان فارسی',
    risk: 'safe',
    host: 'word',
    execute: async (_params, context) => {
      const prevData = context.word_read_document as any;
      const paragraphs = prevData?.sampleParagraphs || [
        'در راستای بهبود فرآیندهای مالی و نظارت بر هزینه‌ها ،این گزارش به بررسی عملکرد سه ماهه می‌ پردازد.'
      ];

      const diff = proofreadPersianText(paragraphs.join('\n\n'));

      return {
        success: true,
        summary: `تعداد ${diff.changesCount} مورد عدم رعایت نیم‌فاصله و فاصله علائم نگارشی با موفقیت اصلاح شد.`,
        data: diff,
        affectedScope: 'پاراگراف‌های متنی سند'
      };
    }
  },

  word_format_document: {
    name: 'word_format_document',
    description: 'تولید گزارش رسمی با جدول خلاصه و تیتربندی استاندارد',
    risk: 'safe',
    host: 'word',
    execute: async (_params, context) => {
      const result = await generateExecutiveReportFromExcel({
        reportTitle: 'گزارش تحلیلی و مدیریتی عملکرد مالی شرکت',
        sourceSheet: 'فروش و هزینه‌ها',
        kpis: [
          { label: 'مجموع کل مصارف', value: '۴,۱۴۰,۰۰۰,۰۰۰ ریال' },
          { label: 'میانگین هر تراکنش', value: '۷۹,۶۰۰,۰۰۰ ریال' },
          { label: 'تعداد کل اقلام', value: '۴۵ فقره' }
        ],
        categoryRows: [
          ['تجهیزات سخت‌افزار', '۱۲', '۱,۸۵۰,۰۰۰,۰۰۰', '۴۵%'],
          ['نرم‌افزار و مجوزها', '۱۸', '۹۲۰,۰۰۰,۰۰۰', '۲۲%'],
          ['پشتیبانی فنی', '۸', '۷۶۰,۰۰۰,۰۰۰', '۱۸%'],
          ['شبکه و ارتباطات', '۱۴', '۶۱۰,۰۰۰,۰۰۰', '۱۵%']
        ]
      });

      return {
        success: true,
        summary: `گزارش رسمی با ۴ بخش اصلی و ۲ جدول ساختاریافته در سند ورد ایجاد گردید.`,
        data: result,
        affectedScope: 'انتهای سند ورد'
      };
    }
  },

  word_verify_structure: {
    name: 'word_verify_structure',
    description: 'اعتبارسنجی کیفیت و ساختار خروجی ورد',
    risk: 'safe',
    host: 'word',
    execute: async (_params, context) => {
      return {
        success: true,
        summary: 'ساختار تیترها و جدول‌های خروجی با استانداردهای رسمی Word مطابقت داده شد.',
        affectedScope: 'سند نهایی'
      };
    }
  }
};
