/**
 * Office AI Agent - Excel Deep Inspector Component (Phase 4)
 */

import React, { useState } from 'react';
import {
  FileSpreadsheet,
  AlertTriangle,
  Copy,
  Table,
  CheckCircle2,
  RefreshCw,
  Search,
  Sparkles,
  ArrowRight,
  Database,
  SlidersHorizontal,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import { ExcelInspectionReport, ColumnDataType } from '../types/office.ts';

const KNOWN_FORMULA_ERRORS = [
  '#N/A',
  '#VALUE!',
  '#REF!',
  '#DIV/0!',
  '#NUM!',
  '#NAME?',
  '#NULL!',
  '#SPILL!',
  '#CALC!'
];

interface ExcelInspectorProps {
  report: ExcelInspectionReport | null;
  isLoading: boolean;
  onRefreshInspection: () => void;
  onAskAiAboutReport: (prompt: string) => void;
}

export const ExcelInspector: React.FC<ExcelInspectorProps> = ({
  report,
  isLoading,
  onRefreshInspection,
  onAskAiAboutReport
}) => {
  const [activeTab, setActiveTab] = useState<'overview' | 'columns' | 'data'>('overview');
  const [isErrorDetailsExpanded, setIsErrorDetailsExpanded] = useState(true);

  if (isLoading) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center p-8 text-slate-500 space-y-3">
        <RefreshCw className="w-8 h-8 animate-spin text-emerald-600" />
        <div className="text-xs font-bold text-slate-800">در حال ارزیابی سلول‌ها و ساختار کاربرگ فعال...</div>
        <p className="text-[11px] text-slate-400 text-center max-w-xs">
          در حال خواندن محدوده پرشده (UsedRange)، تشخیص انواع داده، فرمول‌های نامعتبر و سطرهای تکراری.
        </p>
      </div>
    );
  }

  if (!report) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center p-8 text-slate-500 space-y-3 text-center">
        <FileSpreadsheet className="w-10 h-10 text-slate-300 stroke-1" />
        <div className="text-xs font-bold text-slate-800">هنوز ارزیابی کاربرگ انجام نشده است</div>
        <p className="text-[11px] text-slate-400 max-w-xs">
          برای استخراج خودکار سرستون‌ها، فرمول‌ها و مشکلات کیفی داده‌ها، دکمه زیر را کلیک کنید.
        </p>
        <button
          onClick={onRefreshInspection}
          className="px-4 py-2 bg-emerald-700 text-white rounded-xl text-xs font-bold hover:bg-emerald-800 transition-colors shadow-xs flex items-center gap-2"
        >
          <Search className="w-4 h-4" />
          شروع اسکن هوشمند کاربرگ
        </button>
      </div>
    );
  }

  const getTypeBadge = (type: ColumnDataType) => {
    switch (type) {
      case 'currency':
        return { text: 'ارزی / مالی', color: 'text-emerald-700 bg-emerald-50 border-emerald-200' };
      case 'number':
        return { text: 'عددی', color: 'text-blue-700 bg-blue-50 border-blue-200' };
      case 'date':
        return { text: 'تاریخ', color: 'text-purple-700 bg-purple-50 border-purple-200' };
      case 'formula':
        return { text: 'فرمول‌دار', color: 'text-amber-800 bg-amber-50 border-amber-200' };
      default:
        return { text: 'متنی', color: 'text-slate-700 bg-slate-100 border-slate-200' };
    }
  };

  return (
    <div className="flex-1 flex flex-col h-full overflow-hidden bg-slate-50 font-sans text-xs">
      {/* نوار ابزار بالای بازرس */}
      <div className="bg-white border-b border-slate-200 px-4 py-2.5 flex items-center justify-between shrink-0">
        <div className="flex items-center gap-2">
          <Database className="w-4 h-4 text-emerald-700" />
          <span className="font-bold text-slate-900">
            کاربرگ: <span className="font-mono text-emerald-800">{report.activeSheetName}</span>
          </span>
          <span aria-hidden="true" className="text-slate-300">·</span>
          <span className="font-mono text-[11px] text-slate-500">{report.usedRangeAddress}</span>
        </div>

        <button
          onClick={onRefreshInspection}
          title="اسکن مجدد"
          className="p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors flex items-center gap-1.5 font-medium text-[11px]"
        >
          <RefreshCw className="w-3.5 h-3.5" />
          اسکن مجدد
        </button>
      </div>

      {/* ناوبری داخلی بازرس */}
      <div className="bg-slate-100/90 p-1 border-b border-slate-200 shrink-0">
        <div className="grid grid-cols-3 gap-1 text-[11px] font-medium">
          <button
            onClick={() => setActiveTab('overview')}
            className={`py-1.5 px-2 rounded-md transition-all text-center ${
              activeTab === 'overview'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            خلاصه و کیفیت داده
          </button>
          <button
            onClick={() => setActiveTab('columns')}
            className={`py-1.5 px-2 rounded-md transition-all text-center ${
              activeTab === 'columns'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            ستون‌ها و تایپ‌ها ({report.columnCount})
          </button>
          <button
            onClick={() => setActiveTab('data')}
            className={`py-1.5 px-2 rounded-md transition-all text-center ${
              activeTab === 'data'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            پیش‌نمایش داده‌ها ({report.sampleRows.length})
          </button>
        </div>
      </div>

      {/* محتوای تب فعال */}
      <div className="flex-1 overflow-y-auto p-3.5 space-y-3.5">
        {activeTab === 'overview' && (
          <div className="space-y-3.5">
            {/* کارت‌های آماری ابعاد کاربرگ */}
            <div className="grid grid-cols-3 gap-2">
              <div className="p-3 bg-white border border-slate-200 rounded-xl space-y-0.5">
                <span className="text-[11px] text-slate-500">سطرهای پرشده</span>
                <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
                  {report.rowCount.toLocaleString('fa-IR')}
                </div>
              </div>
              <div className="p-3 bg-white border border-slate-200 rounded-xl space-y-0.5">
                <span className="text-[11px] text-slate-500">تعداد ستون‌ها</span>
                <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
                  {report.columnCount.toLocaleString('fa-IR')}
                </div>
              </div>
              <div className="p-3 bg-white border border-slate-200 rounded-xl space-y-0.5">
                <span className="text-[11px] text-slate-500">سلول‌های خالی</span>
                <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
                  {report.emptyCellCount.toLocaleString('fa-IR')}
                </div>
              </div>
            </div>

            {/* کارت وضعیت سلامت داده‌ها (Data Hygiene) */}
            <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs space-y-3">
              <div className="flex items-center justify-between">
                <span className="font-bold text-slate-900 flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                  ارزیابی سلامت و اعتبارسنجی داده‌ها
                </span>
                <span className="text-[11px] font-mono text-slate-400">
                  {new Date(report.timestamp).toLocaleTimeString('fa-IR')}
                </span>
              </div>

              <p className="text-slate-600 leading-relaxed text-[11px]">
                {report.summaryText}
              </p>

              {/* خطاهای فرمول شناسایی‌شده */}
              {report.formulaErrors.length > 0 && (
                <div className="p-3 bg-red-50/70 border border-red-200 rounded-xl space-y-2">
                  <div
                    onClick={() => setIsErrorDetailsExpanded(!isErrorDetailsExpanded)}
                    className="flex items-center justify-between cursor-pointer"
                  >
                    <span className="font-bold text-red-900 flex items-center gap-1.5">
                      <AlertTriangle className="w-4 h-4 text-red-600" />
                      {report.formulaErrors.length.toLocaleString('fa-IR')} خطای فرمول نیازمند اصلاح
                    </span>
                    {isErrorDetailsExpanded ? (
                      <ChevronUp className="w-4 h-4 text-red-500" />
                    ) : (
                      <ChevronDown className="w-4 h-4 text-red-500" />
                    )}
                  </div>

                  {isErrorDetailsExpanded && (
                    <div className="space-y-1.5 pt-1">
                      {report.formulaErrors.map((err, i) => (
                        <div
                          key={i}
                          className="bg-white p-2 rounded-lg border border-red-100 flex items-center justify-between font-mono text-[11px]"
                        >
                          <span className="text-red-700 font-bold">{err.cellAddress}</span>
                          <span className="bg-red-100 text-red-800 px-1.5 py-0.5 rounded font-bold">
                            {err.errorType}
                          </span>
                          <span className="text-slate-500 truncate max-w-[120px]">{err.formula}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* سطرهای تکراری */}
              {report.duplicateRows.length > 0 && (
                <div className="p-3 bg-amber-50/70 border border-amber-200 rounded-xl space-y-2">
                  <span className="font-bold text-amber-900 flex items-center gap-1.5">
                    <Copy className="w-4 h-4 text-amber-600" />
                    {report.duplicateRows.length.toLocaleString('fa-IR')} سطر تکراری شناسایی شد
                  </span>
                  <div className="space-y-1 pt-0.5">
                    {report.duplicateRows.map((dup, i) => (
                      <div
                        key={i}
                        className="bg-white p-2 rounded-lg border border-amber-100 flex items-center justify-between text-[11px]"
                      >
                        <span className="text-amber-900">
                          ردیف <strong className="font-mono">{dup.duplicateRow}</strong> کپی از ردیف{' '}
                          <strong className="font-mono">{dup.originalRow}</strong>
                        </span>
                        <span className="font-mono text-slate-500 text-[10px] truncate max-w-[130px]">
                          {dup.keyValues.join(' | ')}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* دکمه‌های اقدام سریع با AI بر اساس یافته‌ها */}
            <div className="bg-emerald-50/60 border border-emerald-200 rounded-xl p-3.5 space-y-2.5">
              <div className="font-bold text-emerald-950 flex items-center gap-1.5">
                <Sparkles className="w-4 h-4 text-emerald-700" />
                اقدامات هوشمند پیشنهادی Agent بر اساس یافته‌ها:
              </div>

              <div className="space-y-1.5">
                {report.formulaErrors.length > 0 && (
                  <button
                    onClick={() =>
                      onAskAiAboutReport(
                        `در کاربرگ «${report.activeSheetName}»، خطاهای فرمول شناسایی‌شده (${report.formulaErrors.map((e) => e.cellAddress).join(', ')}) را بررسی کن و علت خطا و فرمول اصلاح‌شده را ارائه بده.`
                      )
                    }
                    className="w-full text-right p-2 rounded-lg bg-white border border-emerald-200 text-emerald-900 hover:bg-emerald-100/50 transition-colors flex items-center justify-between"
                  >
                    <span>رفع و اصلاح خودکار خطاهای فرمول</span>
                    <ArrowRight className="w-3.5 h-3.5 rotate-180 text-emerald-600" />
                  </button>
                )}

                {report.duplicateRows.length > 0 && (
                  <button
                    onClick={() =>
                      onAskAiAboutReport(
                        `داده‌های تکراری کاربرگ «${report.activeSheetName}» (مانند ردیف ${report.duplicateRows[0].duplicateRow}) را قبل از حذف به من گزارش بده و پس از تأیید پاکسازی کن.`
                      )
                    }
                    className="w-full text-right p-2 rounded-lg bg-white border border-emerald-200 text-emerald-900 hover:bg-emerald-100/50 transition-colors flex items-center justify-between"
                  >
                    <span>گزارش و حذف ردیف‌های تکراری با تأییدیه</span>
                    <ArrowRight className="w-3.5 h-3.5 rotate-180 text-emerald-600" />
                  </button>
                )}

                <button
                  onClick={() =>
                    onAskAiAboutReport(
                      `بر اساس داده‌های ارزیابی‌شده کاربرگ «${report.activeSheetName}» با سرستون‌های (${report.headers.join(', ')})، یک شیت داشبورد مدیریتی کامل با KPI و ۲ نمودار تحلیلی بساز.`
                    )
                  }
                  className="w-full text-right p-2 rounded-lg bg-emerald-700 text-white hover:bg-emerald-800 transition-colors flex items-center justify-between font-bold shadow-xs"
                >
                  <span>ساخت داشبورد مدیریتی با KPI و نمودار در شیت جدید</span>
                  <ArrowRight className="w-3.5 h-3.5 rotate-180" />
                </button>
              </div>
            </div>
          </div>
        )}

        {activeTab === 'columns' && (
          <div className="space-y-2">
            {report.columns.map((col) => {
              const badge = getTypeBadge(col.inferredType);
              return (
                <div
                  key={col.index}
                  className="p-3 bg-white border border-slate-200 rounded-xl space-y-1.5"
                >
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-800 flex items-center gap-1.5">
                      <span className="font-mono text-slate-400 text-[11px]">{col.index + 1}.</span>
                      <span>{col.name}</span>
                    </span>
                    <span
                      className={`text-[10px] font-medium px-2 py-0.5 rounded-full border ${badge.color}`}
                    >
                      {badge.text}
                    </span>
                  </div>

                  <div className="text-[11px] text-slate-500 flex items-center justify-between pt-1 border-t border-slate-100">
                    <span className="truncate max-w-[200px]">
                      نمونه‌ها:{' '}
                      <span className="font-mono text-slate-700">
                        {col.sampleValues.slice(0, 3).join(' ، ')}
                      </span>
                    </span>
                    {col.emptyCount > 0 && (
                      <span className="text-amber-700 font-medium">
                        {col.emptyCount} سلول خالی
                      </span>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {activeTab === 'data' && (
          <div className="bg-white border border-slate-200 rounded-xl overflow-x-auto shadow-xs">
            <table className="w-full text-right border-collapse">
              <thead>
                <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 text-[11px]">
                  <th className="p-2 border-l border-slate-200 w-10 text-center font-mono">#</th>
                  {report.headers.map((h, i) => (
                    <th key={i} className="p-2 border-l border-slate-200 whitespace-nowrap font-bold">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-[11px]">
                {report.sampleRows.map((row, rIdx) => (
                  <tr key={rIdx} className="hover:bg-slate-50/80 transition-colors">
                    <td className="p-2 border-l border-slate-100 text-center font-mono text-slate-400 tabular-nums">
                      {rIdx + 1}
                    </td>
                    {report.headers.map((h, cIdx) => {
                      const val = String(row[h] ?? '');
                      const isError = KNOWN_FORMULA_ERRORS.some((err) => val.includes(err));

                      return (
                        <td
                          key={cIdx}
                          className={`p-2 border-l border-slate-100 whitespace-nowrap font-mono tabular-nums ${
                            isError ? 'bg-red-50 text-red-700 font-bold' : 'text-slate-700'
                          }`}
                        >
                          {val || <span className="text-slate-300">خالی</span>}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};
