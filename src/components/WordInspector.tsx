/**
 * Office AI Agent - Word Document Inspector & Proofreader Component (Phase 6)
 */

import React, { useState } from 'react';
import {
  FileText,
  CheckCircle2,
  AlertTriangle,
  RefreshCw,
  Sparkles,
  BookOpen,
  ArrowRight,
  ListTree,
  Wand2,
  Table as TableIcon
} from 'lucide-react';
import { WordInspectionReport } from '../office/word/reader.ts';
import { proofreadPersianText, ProofreadingDiff } from '../office/word/proofreader.ts';

interface WordInspectorProps {
  report: WordInspectionReport | null;
  isLoading: boolean;
  onRefreshInspection: () => void;
  onAskAiAboutReport: (prompt: string) => void;
  onGenerateReportFromExcel: () => void;
}

export const WordInspector: React.FC<WordInspectorProps> = ({
  report,
  isLoading,
  onRefreshInspection,
  onAskAiAboutReport,
  onGenerateReportFromExcel
}) => {
  const [activeTab, setActiveTab] = useState<'overview' | 'proofing' | 'structure'>('overview');
  const [proofreadResult, setProofreadResult] = useState<ProofreadingDiff | null>(null);

  if (isLoading) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center p-8 text-slate-500 space-y-3">
        <RefreshCw className="w-8 h-8 animate-spin text-blue-600" />
        <div className="text-xs font-bold text-slate-800">در حال ارزیابی ساختار سند و نگارش ورد...</div>
        <p className="text-[11px] text-slate-400 text-center max-w-xs">
          در حال شمارش پاراگراف‌ها، استخراج تیترها، بررسی نیم‌فاصله‌ها و علائم نگارشی.
        </p>
      </div>
    );
  }

  if (!report) {
    return (
      <div className="flex-1 flex flex-col items-center justify-center p-8 text-slate-500 space-y-3 text-center">
        <FileText className="w-10 h-10 text-slate-300 stroke-1" />
        <div className="text-xs font-bold text-slate-800">ارزیابی سند ورد انجام نشده است</div>
        <button
          onClick={onRefreshInspection}
          className="px-4 py-2 bg-blue-700 text-white rounded-xl text-xs font-bold hover:bg-blue-800 transition-colors shadow-xs flex items-center gap-2"
        >
          <RefreshCw className="w-4 h-4" />
          شروع ارزیابی و ویراستاری سند
        </button>
      </div>
    );
  }

  const handleApplyAutoProofread = () => {
    const combinedText = report.sampleParagraphs.join('\n\n');
    const diff = proofreadPersianText(combinedText);
    setProofreadResult(diff);
  };

  const readingTimeMin = Math.max(1, Math.round(report.wordCount / 180));

  return (
    <div className="flex-1 flex flex-col h-full overflow-hidden bg-slate-50 font-sans text-xs">
      {/* هدر بالای بازرس ورد */}
      <div className="bg-white border-b border-slate-200 px-4 py-2.5 flex items-center justify-between shrink-0">
        <div className="flex items-center gap-2">
          <FileText className="w-4 h-4 text-blue-700" />
          <span className="font-bold text-slate-900 truncate max-w-[180px]">
            {report.documentTitle}
          </span>
          <span aria-hidden="true" className="text-slate-300">·</span>
          <span className="text-[11px] text-slate-500 font-mono">
            {report.wordCount.toLocaleString('fa-IR')} کلمه
          </span>
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

      {/* ناوبری داخلی */}
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
            خلاصه و ویراستاری
          </button>
          <button
            onClick={() => setActiveTab('proofing')}
            className={`py-1.5 px-2 rounded-md transition-all text-center ${
              activeTab === 'proofing'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            ایرادات نگارشی ({report.proofingIssues.length})
          </button>
          <button
            onClick={() => setActiveTab('structure')}
            className={`py-1.5 px-2 rounded-md transition-all text-center ${
              activeTab === 'structure'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            ساختار تیترها ({report.headings.length})
          </button>
        </div>
      </div>

      {/* محتوا */}
      <div className="flex-1 overflow-y-auto p-3.5 space-y-3.5">
        {activeTab === 'overview' && (
          <div className="space-y-3.5">
            {/* کارت‌های آماری ابعاد سند */}
            <div className="grid grid-cols-3 gap-2">
              <div className="p-3 bg-white border border-slate-200 rounded-xl space-y-0.5">
                <span className="text-[11px] text-slate-500">پاراگراف‌ها</span>
                <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
                  {report.paragraphCount.toLocaleString('fa-IR')}
                </div>
              </div>
              <div className="p-3 bg-white border border-slate-200 rounded-xl space-y-0.5">
                <span className="text-[11px] text-slate-500">تعداد جدول‌ها</span>
                <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
                  {report.tableCount.toLocaleString('fa-IR')}
                </div>
              </div>
              <div className="p-3 bg-white border border-slate-200 rounded-xl space-y-0.5">
                <span className="text-[11px] text-slate-500">زمان خواندن</span>
                <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
                  {readingTimeMin.toLocaleString('fa-IR')} دقیقه
                </div>
              </div>
            </div>

            {/* کارت ارزیابی نگارشی و نیم‌فاصله‌ها */}
            <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs space-y-3">
              <div className="flex items-center justify-between">
                <span className="font-bold text-slate-900 flex items-center gap-2">
                  <Wand2 className="w-4 h-4 text-blue-700" />
                  ارزیابی نگارش، نیم‌فاصله و علائم سجاوندی
                </span>
                <span className="text-[11px] font-mono text-slate-400">
                  {new Date(report.timestamp).toLocaleTimeString('fa-IR')}
                </span>
              </div>

              <p className="text-slate-600 leading-relaxed text-[11px]">
                {report.proofingIssues.length > 0
                  ? `تعداد ${report.proofingIssues.length.toLocaleString('fa-IR')} مورد عدم رعایت نیم‌فاصله یا فاصله نامناسب علائم نگارشی شناسایی شد.`
                  : 'متن سند از نظر نیم‌فاصله‌ها و قواعد ویرایشی زبان فارسی کاملاً مرتب و استاندارد است.'}
              </p>

              <div className="flex gap-2">
                <button
                  onClick={handleApplyAutoProofread}
                  className="flex-1 py-2 px-3 bg-blue-700 text-white rounded-xl font-bold hover:bg-blue-800 transition-colors shadow-xs flex items-center justify-center gap-1.5"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  اصلاح خودکار نیم‌فاصله‌ها
                </button>

                <button
                  onClick={() =>
                    onAskAiAboutReport(
                      `این سند ورد را از نظر لحن رسمی، ساختار پاراگراف‌ها و خلاصه‌سازی ارزیابی کن و یک نسخه بازنویسی‌شده استاندارد ارائه بده.`
                    )
                  }
                  className="py-2 px-3 border border-slate-300 text-slate-700 rounded-xl font-bold hover:bg-slate-100 transition-colors"
                >
                  تحلیل جامع با AI
                </button>
              </div>

              {/* نتیجه اصلاح خودکار */}
              {proofreadResult && (
                <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl space-y-2 text-[11px] animate-in fade-in">
                  <div className="font-bold text-emerald-900 flex items-center gap-1.5">
                    <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                    تعداد {proofreadResult.changesCount} اصلاح نگارشی با موفقیت اعمال گردید.
                  </div>
                  <div className="space-y-1">
                    {proofreadResult.fixedSnippets.map((fix, idx) => (
                      <div key={idx} className="flex items-center gap-2 text-slate-700">
                        <span className="line-through text-red-600 bg-red-50 px-1 rounded">{fix.before}</span>
                        <ArrowRight className="w-3 h-3 text-slate-400 rotate-180" />
                        <span className="font-bold text-emerald-800 bg-emerald-100/70 px-1 rounded">{fix.after}</span>
                        <span className="text-slate-400 text-[10px]">({fix.reason})</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* کارت تولید گزارش رسمی از اکسل */}
            <div className="bg-blue-50/70 border border-blue-200 rounded-xl p-4 space-y-2.5">
              <div className="font-bold text-blue-950 flex items-center gap-2">
                <TableIcon className="w-4 h-4 text-blue-700" />
                تولید گزارش رسمی Word بر اساس داده‌های Excel:
              </div>
              <p className="text-blue-900 leading-relaxed text-[11px]">
                انتقال خودکار شاخص‌های مالی KPI، جدول تفکیک هزینه‌ها و نمودارهای کاربرگ اکسل به یک سند رسمی با تیتربندی استاندارد، مقدمه و نتیجه‌گیری.
              </p>
              <button
                onClick={onGenerateReportFromExcel}
                className="w-full py-2 px-3 bg-blue-800 hover:bg-blue-900 text-white rounded-xl font-bold transition-colors shadow-xs flex items-center justify-between"
              >
                <span>درج گزارش مدیریتی با جدول و تحلیل در سند ورد</span>
                <ArrowRight className="w-3.5 h-3.5 rotate-180" />
              </button>
            </div>
          </div>
        )}

        {activeTab === 'proofing' && (
          <div className="space-y-2">
            {report.proofingIssues.map((issue, idx) => (
              <div
                key={idx}
                className="p-3 bg-white border border-slate-200 rounded-xl space-y-1.5"
              >
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-800 flex items-center gap-1.5">
                    <AlertTriangle className="w-3.5 h-3.5 text-amber-600" />
                    <span>پاراگراف {issue.paragraphIndex}</span>
                  </span>
                  <span className="text-[10px] text-amber-800 bg-amber-50 px-2 py-0.5 rounded-full border border-amber-200">
                    {issue.description}
                  </span>
                </div>

                <div className="flex items-center gap-2 text-[11px] pt-1 border-t border-slate-100">
                  <span className="text-red-700 bg-red-50 px-1.5 py-0.5 rounded font-mono">
                    {issue.originalSnippet}
                  </span>
                  <ArrowRight className="w-3 h-3 text-slate-400 rotate-180" />
                  <span className="text-emerald-800 font-bold bg-emerald-50 px-1.5 py-0.5 rounded font-mono">
                    {issue.suggestedSnippet}
                  </span>
                </div>
              </div>
            ))}
          </div>
        )}

        {activeTab === 'structure' && (
          <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
            <div className="font-bold text-slate-900 flex items-center gap-2">
              <ListTree className="w-4 h-4 text-blue-700" />
              سلسله‌مراتب تیترهای سند (Headings Hierarchy)
            </div>

            <div className="space-y-2 pt-1">
              {report.headings.map((h, i) => (
                <div
                  key={i}
                  className={`p-2 rounded-lg border text-slate-800 flex items-center gap-2 ${
                    h.level === 1
                      ? 'bg-blue-50/50 border-blue-200 font-bold mr-0'
                      : 'bg-slate-50 border-slate-200 mr-4'
                  }`}
                >
                  <span className="font-mono text-[10px] text-blue-600">
                    H{h.level}
                  </span>
                  <span>{h.text}</span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
