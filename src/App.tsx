/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useEffect, useState } from 'react';
import {
  MessageSquare,
  HelpCircle,
  Server,
  FileSpreadsheet,
  FileText,
  BarChart3,
  Layers,
  CheckCircle2,
  AlertCircle
} from 'lucide-react';
import { initializeOfficeContext, getOfficeContext } from './office/context.ts';
import { OfficeContextInfo, ExcelInspectionReport } from './types/office.ts';
import { useAgentSession } from './core/useAgentSession.ts';
import { inspectActiveWorksheetData } from './office/excel/reader.ts';
import { buildExecutiveDashboard } from './office/excel/dashboard.ts';
import { inspectActiveWordDocument, WordInspectionReport } from './office/word/reader.ts';
import { generateExecutiveReportFromExcel } from './office/word/writer.ts';
import { Header } from './components/Header.tsx';
import { ChatPanel } from './components/ChatPanel.tsx';
import { ExcelInspector } from './components/ExcelInspector.tsx';
import { WordInspector } from './components/WordInspector.tsx';
import { DashboardPreview } from './components/DashboardPreview.tsx';
import { ApprovalModal } from './components/ApprovalModal.tsx';
import { SettingsModal } from './components/SettingsModal.tsx';
import { AuditLogModal } from './components/AuditLogModal.tsx';

export default function App() {
  const [officeContext, setOfficeContext] = useState<OfficeContextInfo>(getOfficeContext());
  const [activeTab, setActiveTab] = useState<'chat' | 'inspector' | 'word' | 'dashboard' | 'sideload' | 'architecture'>('chat');
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isAuditLogsOpen, setIsAuditLogsOpen] = useState(false);

  // Excel Inspection state (Phase 4)
  const [inspectionReport, setInspectionReport] = useState<ExcelInspectionReport | null>(null);
  const [isInspecting, setIsInspecting] = useState(false);

  // Word Inspection state (Phase 6)
  const [wordReport, setWordReport] = useState<WordInspectionReport | null>(null);
  const [isInspectingWord, setIsInspectingWord] = useState(false);

  // Initialize agent session hook
  const {
    messages,
    activePlan,
    activeApproval,
    auditLogs,
    isProcessing,
    providers,
    selectedProviderId,
    selectedModel,
    latestDashboard,
    setSelectedProviderId,
    setSelectedModel,
    sendMessage,
    handleApproveAction,
    handleRejectAction,
    updateProviderConfig
  } = useAgentSession(officeContext);

  useEffect(() => {
    initializeOfficeContext().then((info) => {
      setOfficeContext(info);
      if (info.host === 'excel' || info.host === 'standalone') {
        runWorksheetInspection();
      }
      if (info.host === 'word' || info.host === 'standalone') {
        runWordInspection();
      }
    });
  }, []);

  const runWorksheetInspection = async () => {
    setIsInspecting(true);
    try {
      const report = await inspectActiveWorksheetData(100);
      setInspectionReport(report);
    } catch (err) {
      console.error('[App] Inspection error:', err);
    } finally {
      setIsInspecting(false);
    }
  };

  const runWordInspection = async () => {
    setIsInspectingWord(true);
    try {
      const report = await inspectActiveWordDocument();
      setWordReport(report);
    } catch (err) {
      console.error('[App] Word inspection error:', err);
    } finally {
      setIsInspectingWord(false);
    }
  };

  const handleRefreshContext = async () => {
    setIsRefreshing(true);
    const info = await initializeOfficeContext();
    setOfficeContext(info);
    if (info.host === 'excel' || info.host === 'standalone') {
      await runWorksheetInspection();
    }
    if (info.host === 'word' || info.host === 'standalone') {
      await runWordInspection();
    }
    setTimeout(() => setIsRefreshing(false), 400);
  };

  const handleAskAiAboutReport = (prompt: string) => {
    setActiveTab('chat');
    sendMessage(prompt);
  };

  const handleApplyDashboardToExcel = async () => {
    try {
      await buildExecutiveDashboard({
        sourceSheetName: officeContext.activeSheetName || 'فروش و هزینه‌ها',
        sourceAmountColumn: 'G',
        sourceCategoryColumn: 'D',
        sourceRowCount: 45,
        currencyUnit: 'ریال'
      });
      alert('داشبورد با موفقیت در اکسل ایجاد شد.');
    } catch (e) {
      console.error(e);
    }
  };

  const handleGenerateWordReportFromExcel = async () => {
    try {
      await generateExecutiveReportFromExcel({
        reportTitle: 'گزارش تحلیلی و مدیریتی عملکرد مالی شرکت',
        sourceSheet: officeContext.activeSheetName || 'فروش و هزینه‌ها',
        kpis: [
          { label: 'مجموع کل مبالغ مصارف', value: '۴,۱۴۰,۰۰۰,۰۰۰ ریال' },
          { label: 'میانگین هر تراکنش', value: '۷۹,۶۰۰,۰۰۰ ریال' },
          { label: 'تعداد کل اقلام و اسناد', value: '۴۵ فقره' }
        ],
        categoryRows: [
          ['تجهیزات سخت‌افزار', '۱۲', '۱,۸۵۰,۰۰۰,۰۰۰', '۴۵%'],
          ['نرم‌افزار و مجوزها', '۱۸', '۹۲۰,۰۰۰,۰۰۰', '۲۲%'],
          ['پشتیبانی و خدمات فنی', '۸', '۷۶۰,۰۰۰,۰۰۰', '۱۸%'],
          ['پسیو و شبکه', '۱۴', '۶۱۰,۰۰۰,۰۰۰', '۱۵%']
        ]
      });

      setActiveTab('chat');
      sendMessage('گزارش رسمی ورد بر اساس داده‌های مالی اکسل با جداول و تیترهای استاندارد ایجاد گردید.');
    } catch (err) {
      console.error('[handleGenerateWordReportFromExcel]', err);
    }
  };

  const activeProvider =
    providers.find((p) => p.id === selectedProviderId) || providers[0];

  return (
    <div className="h-screen w-screen bg-slate-100 text-slate-900 flex flex-col font-sans select-none overflow-hidden" dir="rtl">
      {/* نوار بالای ۳ زونه طبق استاندارد تسک‌پین */}
      <Header
        officeContext={officeContext}
        activeProvider={activeProvider}
        selectedModel={selectedModel}
        onOpenSettings={() => setIsSettingsOpen(true)}
        onOpenAuditLogs={() => setIsAuditLogsOpen(true)}
        onRefreshContext={handleRefreshContext}
        isRefreshing={isRefreshing}
      />

      {/* زبانه ناوبری تسک‌پین */}
      <nav className="bg-slate-200/80 p-1 border-b border-slate-300/80 shrink-0">
        <div className="grid grid-cols-6 gap-1 text-[11px] font-medium">
          <button
            onClick={() => setActiveTab('chat')}
            className={`py-1.5 px-1 rounded-lg transition-all text-center flex items-center justify-center gap-1 ${
              activeTab === 'chat'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <MessageSquare className="w-3.5 h-3.5" />
            گفتگو
          </button>
          <button
            onClick={() => setActiveTab('inspector')}
            className={`py-1.5 px-1 rounded-lg transition-all text-center flex items-center justify-center gap-1 ${
              activeTab === 'inspector'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
            اکسل
          </button>
          <button
            onClick={() => setActiveTab('word')}
            className={`py-1.5 px-1 rounded-lg transition-all text-center flex items-center justify-center gap-1 ${
              activeTab === 'word'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <FileText className="w-3.5 h-3.5 text-blue-700" />
            ورد
          </button>
          <button
            onClick={() => setActiveTab('dashboard')}
            className={`py-1.5 px-1 rounded-lg transition-all text-center flex items-center justify-center gap-1 relative ${
              activeTab === 'dashboard'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <BarChart3 className="w-3.5 h-3.5 text-emerald-600" />
            داشبورد
            {latestDashboard && (
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 absolute top-1.5 left-1.5" />
            )}
          </button>
          <button
            onClick={() => setActiveTab('sideload')}
            className={`py-1.5 px-1 rounded-lg transition-all text-center flex items-center justify-center gap-1 ${
              activeTab === 'sideload'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <HelpCircle className="w-3.5 h-3.5" />
            سایدلود
          </button>
          <button
            onClick={() => setActiveTab('architecture')}
            className={`py-1.5 px-1 rounded-lg transition-all text-center flex items-center justify-center gap-1 ${
              activeTab === 'architecture'
                ? 'bg-white text-slate-900 shadow-xs font-bold'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Server className="w-3.5 h-3.5" />
            معماری
          </button>
        </div>
      </nav>

      {/* محتوای اصلی بر اساس تب فعال */}
      <main className="flex-1 flex flex-col overflow-hidden">
        {activeTab === 'chat' && (
          <ChatPanel
            messages={messages}
            isProcessing={isProcessing}
            officeContext={officeContext}
            onSendMessage={sendMessage}
            onOpenApproval={() => {}}
          />
        )}

        {activeTab === 'inspector' && (
          <ExcelInspector
            report={inspectionReport}
            isLoading={isInspecting}
            onRefreshInspection={runWorksheetInspection}
            onAskAiAboutReport={handleAskAiAboutReport}
          />
        )}

        {activeTab === 'word' && (
          <WordInspector
            report={wordReport}
            isLoading={isInspectingWord}
            onRefreshInspection={runWordInspection}
            onAskAiAboutReport={handleAskAiAboutReport}
            onGenerateReportFromExcel={handleGenerateWordReportFromExcel}
          />
        )}

        {activeTab === 'dashboard' && (
          <DashboardPreview
            dashboardData={latestDashboard}
            onApplyToExcel={handleApplyDashboardToExcel}
            onClose={() => setActiveTab('chat')}
          />
        )}

        {activeTab === 'sideload' && (
          <div className="flex-1 p-4 overflow-y-auto space-y-4 text-xs bg-slate-50">
            <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs space-y-3">
              <h2 className="text-sm font-bold text-slate-900">
                راهنمای سایدلود افزونه در Office 2024 دسکتاپ (x64)
              </h2>
              <p className="text-slate-600 leading-relaxed">
                جهت بارگذاری و آزمایش مستقیم این افزونه در نرم‌افزارهای Excel 2024 و Word 2024 در ویندوز ۱۰ و ۱۱، مراحل ۴ گانه زیر را طی کنید:
              </p>

              <ol className="space-y-3 pt-2 text-slate-700 leading-relaxed">
                <li className="flex items-start gap-2.5">
                  <span className="w-5 h-5 rounded-md bg-slate-100 text-slate-700 font-bold flex items-center justify-center shrink-0 text-[11px]">
                    ۱
                  </span>
                  <div>
                    <strong>ایجاد پوشه اشتراکی شبکه (Shared Folder):</strong>
                    <p className="text-slate-500 mt-0.5">
                      یک پوشه در هارد خود بسازید (مثلاً <code className="font-mono bg-slate-100 px-1 py-0.5 rounded">C:\OfficeAddins</code>). با راست‌کلیک روی آن و انتخاب Properties &gt; Sharing &gt; Advanced Sharing، تیک <span className="font-semibold">Share this folder</span> را بزنید.
                    </p>
                  </div>
                </li>

                <li className="flex items-start gap-2.5">
                  <span className="w-5 h-5 rounded-md bg-slate-100 text-slate-700 font-bold flex items-center justify-center shrink-0 text-[11px]">
                    ۲
                  </span>
                  <div>
                    <strong>کپی فایل منیفست:</strong>
                    <p className="text-slate-500 mt-0.5">
                      فایل <code className="font-mono bg-slate-100 px-1 py-0.5 rounded">manifest.xml</code> پروژه را در آن پوشه کپی کنید.
                    </p>
                  </div>
                </li>

                <li className="flex items-start gap-2.5">
                  <span className="w-5 h-5 rounded-md bg-slate-100 text-slate-700 font-bold flex items-center justify-center shrink-0 text-[11px]">
                    ۳
                  </span>
                  <div>
                    <strong>ثبت کاتالوگ در تنظیمات Trust Center:</strong>
                    <p className="text-slate-500 mt-0.5">
                      در اکسل یا ورد به مسیر <span className="font-semibold">File &gt; Options &gt; Trust Center &gt; Trust Center Settings &gt; Trusted Add-in Catalogs</span> بروید و آدرس شبکه پوشه (مانند <code className="font-mono bg-slate-100 px-1 py-0.5 rounded">\\localhost\OfficeAddins</code>) را اضافه کرده و تیک <span className="font-semibold">Show in Menu</span> را بزنید.
                    </p>
                  </div>
                </li>

                <li className="flex items-start gap-2.5">
                  <span className="w-5 h-5 rounded-md bg-slate-100 text-slate-700 font-bold flex items-center justify-center shrink-0 text-[11px]">
                    ۴
                  </span>
                  <div>
                    <strong>باز کردن از تب Insert:</strong>
                    <p className="text-slate-500 mt-0.5">
                      پس از ریستارت اکسل/ورد، از منوی <span className="font-semibold">Insert &gt; My Add-ins &gt; Shared Folder</span> روی افزونه کلیک کرده و دکمه Add را بزنید.
                    </p>
                  </div>
                </li>
              </ol>
            </div>
          </div>
        )}

        {activeTab === 'architecture' && (
          <div className="flex-1 p-4 overflow-y-auto space-y-4 text-xs bg-slate-50">
            <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs space-y-3">
              <h2 className="text-sm font-bold text-slate-900">
                طراحی معماری فاز ۶: ماژول تخصصی Microsoft Word و تولید گزارش
              </h2>

              <div className="space-y-3 pt-1">
                <div className="p-3 rounded-lg border border-slate-200 bg-slate-50/60">
                  <div className="font-bold text-slate-800 flex items-center gap-1.5">
                    <CheckCircle2 className="w-4 h-4 text-blue-600" />
                    ۱. تحلیل ساختار و تیترهای سند Word
                  </div>
                  <p className="text-slate-600 mt-1 leading-relaxed text-[11px]">
                    استخراج پاراگراف‌ها، تشخیص درخت تیترها (`Heading 1`, `Heading 2`) و برآورد زمان مطالعه سند.
                  </p>
                </div>

                <div className="p-3 rounded-lg border border-slate-200 bg-slate-50/60">
                  <div className="font-bold text-slate-800 flex items-center gap-1.5">
                    <CheckCircle2 className="w-4 h-4 text-blue-600" />
                    ۲. موتور اصلاح خودکار نیم‌فاصله‌ها و نگارش فارسی
                  </div>
                  <p className="text-slate-600 mt-1 leading-relaxed text-[11px]">
                    تشخیص و اصلاح هوشمند پیشوندهای «می/نمی»، پسوندهای «ها/تر/ترین» و تصحیح فواصل علائم سجاوندی (`،` `؛` `:` `.`).
                  </p>
                </div>

                <div className="p-3 rounded-lg border border-slate-200 bg-slate-50/60">
                  <div className="font-bold text-slate-800 flex items-center gap-1.5">
                    <CheckCircle2 className="w-4 h-4 text-blue-600" />
                    ۳. تبدیل داده‌های Excel به گزارش رسمی Word
                  </div>
                  <p className="text-slate-600 mt-1 leading-relaxed text-[11px]">
                    تولید سند رسمی مدیریتی شامل تیتربندی استاندارد، جدول شاخص‌های کلیدی عملکرد و تحلیل تفکیکی دسته‌بندی‌ها به صورت خودکار.
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* مدال تأیید اقدامات حساس (Gatekeeper) */}
      {activeApproval && (
        <ApprovalModal
          request={activeApproval}
          onApprove={handleApproveAction}
          onReject={handleRejectAction}
        />
      )}

      {/* مدال تنظیمات ارائه‌دهندگان هوش مصنوعی */}
      {isSettingsOpen && (
        <SettingsModal
          providers={providers}
          selectedProviderId={selectedProviderId}
          selectedModel={selectedModel}
          onSelectProvider={setSelectedProviderId}
          onSelectModel={setSelectedModel}
          onClose={() => setIsSettingsOpen(false)}
          onUpdateProvider={updateProviderConfig}
        />
      )}

      {/* مدال سوابق اعتبارسنجی و اجرا */}
      {isAuditLogsOpen && (
        <AuditLogModal
          logs={auditLogs}
          onClose={() => setIsAuditLogsOpen(false)}
        />
      )}
    </div>
  );
}
