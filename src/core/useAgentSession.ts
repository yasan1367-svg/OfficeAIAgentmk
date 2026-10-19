/**
 * Office AI Agent - Core Session & State Hook (Phase 7 Orchestrator Integration)
 */

import { useState, useCallback, useEffect, useRef } from 'react';
import {
  ChatMessage,
  ExecutionPlan,
  PlanStep,
  ApprovalRequest,
  AuditLogEntry,
  ActionRiskLevel
} from '../types/agent.ts';
import { ProviderConfig, ProviderType } from '../types/provider.ts';
import { OfficeContextInfo } from '../types/office.ts';
import { AgentOrchestrator } from './orchestrator.ts';
import { TOOL_REGISTRY } from './tools.ts';
import { DashboardResult } from '../office/excel/dashboard.ts';

const DEFAULT_PROVIDERS: ProviderConfig[] = [
  {
    id: 'houshyar',
    name: 'Houshyar24 (هوشیار ۲۴)',
    baseUrl: 'https://api.houshyar24.ir/v1',
    defaultModel: 'gpt-4o',
    availableModels: ['gpt-4o', 'gpt-4o-mini', 'claude-3-5-sonnet'],
    isConfigured: true,
    requiresProxy: true,
    apiKeyMasked: 'sk-hy••••••••9841'
  },
  {
    id: 'openai',
    name: 'OpenAI Direct',
    baseUrl: 'https://api.openai.com/v1',
    defaultModel: 'gpt-4o',
    availableModels: ['gpt-4o', 'gpt-4o-mini'],
    isConfigured: false,
    requiresProxy: true
  },
  {
    id: 'gemini',
    name: 'Google Gemini',
    baseUrl: 'https://generativelanguage.googleapis.com',
    defaultModel: 'gemini-3.8-flash',
    availableModels: ['gemini-3.8-flash', 'gemini-3.1-flash-lite'],
    isConfigured: true,
    requiresProxy: true,
    apiKeyMasked: 'AIza••••••••8392'
  }
];

export function useAgentSession(officeContext: OfficeContextInfo) {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [activePlan, setActivePlan] = useState<ExecutionPlan | null>(null);
  const [activeApproval, setActiveApproval] = useState<ApprovalRequest | null>(null);
  const [auditLogs, setAuditLogs] = useState<AuditLogEntry[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [providers, setProviders] = useState<ProviderConfig[]>(DEFAULT_PROVIDERS);
  const [selectedProviderId, setSelectedProviderId] = useState<ProviderType>('houshyar');
  const [selectedModel, setSelectedModel] = useState<string>('gpt-4o');
  const [latestDashboard, setLatestDashboard] = useState<DashboardResult | null>(null);

  const orchestratorRef = useRef<AgentOrchestrator>(new AgentOrchestrator());

  // Load server-configured providers on mount
  useEffect(() => {
    fetch('/api/agent/providers')
      .then((res) => res.json())
      .then((data) => {
        if (data.providers && Array.isArray(data.providers)) {
          setProviders(data.providers);
        }
      })
      .catch((err) => {
        console.warn('Could not load providers from server:', err);
      });
  }, []);

  // Initialize welcome message when host is detected
  useEffect(() => {
    const hostLabel =
      officeContext.host === 'excel'
        ? 'مایکروسافت اکسل ۲۰۲۴'
        : officeContext.host === 'word'
        ? 'مایکروسافت ورد ۲۰۲۴'
        : 'محیط توسعه آفیس ۲۰۲۴';

    const initialGreeting: ChatMessage = {
      id: 'msg-welcome',
      role: 'assistant',
      timestamp: Date.now(),
      content: `سلام! من **Office AI Agent** هستم؛ دستیار اجرایی و تحلیلگر شما در ${hostLabel}.\n\nشما می‌توانید به زبان طبیعی از من بخواهید کارهایی مانند:\n- «ساخت داشبورد مدیریتی با نمودار و KPI در شیت جدید»\n- «بررسی داده‌های پرت، تکراری و رفع خطاهای فرمول»\n- «اصلاح ساختار، تیترها و نگارش گزارش» (در Word)\n- «خلاصه‌سازی اطلاعات مالی و مقایسه روندها»\nرا انجام دهم. برای عملیات حساس، همواره پیش از اجرا از شما تأیید خواهم گرفت.`
    };

    setMessages([initialGreeting]);
  }, [officeContext.host]);

  // Generate an intelligent plan based on the user's natural prompt
  const analyzePromptAndCreatePlan = useCallback(
    (promptText: string): ExecutionPlan => {
      const lower = promptText.toLowerCase();
      const planId = `plan-${Date.now()}`;
      const isExcel = officeContext.host === 'excel' || officeContext.host === 'standalone';

      const isDestructive =
        lower.includes('حذف') ||
        lower.includes('delete') ||
        lower.includes('پاک کردن') ||
        lower.includes('جایگزین') ||
        lower.includes('overwrite');

      const isDashboard =
        lower.includes('داشبورد') ||
        lower.includes('dashboard') ||
        lower.includes('kpi') ||
        lower.includes('نمودار') ||
        lower.includes('chart');

      const steps: PlanStep[] = [];

      if (isExcel) {
        steps.push({
          id: `${planId}-step-1`,
          order: 1,
          title: 'بررسی ساختار کاربرگ و خواندن محدوده داده‌ها',
          description: 'پیمایش شیت فعال، استخراج سرستون‌ها و نمونه‌گیری از سطرها بدون دستکاری فایل',
          toolName: 'excel_read_structure',
          targetScope: 'کاربرگ فعال (Active Sheet)',
          risk: 'safe',
          status: 'pending'
        });

        steps.push({
          id: `${planId}-step-2`,
          order: 2,
          title: 'تحلیل داده‌ها، اعتبارسنجی فرمول‌ها و تشخیص ناهماهنگی‌ها',
          description: 'محاسبه مقادیر کلیدی، شناسایی ردیف‌های نامعتبر یا خطاها',
          toolName: 'excel_analyze_metrics',
          targetScope: 'داده‌های جدول اصلی',
          risk: 'safe',
          status: 'pending'
        });

        if (isDestructive) {
          steps.push({
            id: `${planId}-step-3`,
            order: 3,
            title: 'حذف ردیف‌های تکراری و پالایش اطلاعات',
            description: 'پاکسازی رکوردهای تکراری بر اساس شناسه و مقادیر ستون‌های مبدأ',
            toolName: 'excel_remove_duplicates',
            targetScope: 'محدوده سلول‌های داده اصلی',
            risk: 'confirm_required',
            status: 'pending'
          });
        } else if (isDashboard) {
          steps.push({
            id: `${planId}-step-3`,
            order: 3,
            title: 'ایجاد شیت جدید «داشبورد_مدیریتی» با کارت‌های KPI',
            description: 'تولید کارت‌های شاخص عملکرد و جدول خلاصه دسته‌بندی‌ها با SUMIFS',
            toolName: 'excel_build_dashboard',
            targetScope: 'شیت جدید: داشبورد_مدیریتی',
            risk: 'safe',
            status: 'pending'
          });
        } else {
          steps.push({
            id: `${planId}-step-3`,
            order: 3,
            title: 'اعمال فرمول‌ها و فرمت‌بندی ارقام و تاریخ‌ها',
            description: 'تنظیم فونت، راست‌چین کردن متون و اعمال فرمت اعداد با جداکننده هزارگان',
            toolName: 'excel_format_and_formulas',
            targetScope: 'ستون‌های مبلغ و تاریخ',
            risk: 'safe',
            status: 'pending'
          });
        }

        steps.push({
          id: `${planId}-step-final`,
          order: steps.length + 1,
          title: 'بازرسی نتایج در Office.js و اعتبارسنجی تغییرات',
          description: 'اطمینان از صحت مقادیر درج‌شده و ارائه گزارش نهایی به کاربر',
          toolName: 'excel_verify_results',
          targetScope: 'گزارش نهایی سند',
          risk: 'safe',
          status: 'pending'
        });
      } else {
        steps.push({
          id: `${planId}-step-1`,
          order: 1,
          title: 'خواندن پاراگراف‌ها و بررسی ساختار سند',
          description: 'استخراج تیترها، بدنه متن و جدول‌های موجود در سند ورد',
          toolName: 'word_read_document',
          targetScope: 'کل سند ورد فعال',
          risk: 'safe',
          status: 'pending'
        });

        steps.push({
          id: `${planId}-step-2`,
          order: 2,
          title: 'تحلیل نگارشی، علائم ویرایشی و یکپارچگی سبک',
          description: 'بررسی نیم‌فاصله‌ها، قواعد نگارش فارسی و اصلاح ساختار تیترها',
          toolName: 'word_proofread',
          targetScope: 'پاراگراف‌های متنی',
          risk: 'safe',
          status: 'pending'
        });

        steps.push({
          id: `${planId}-step-3`,
          order: 3,
          title: 'درج جدول خلاصه نکات کلیدی و قالب‌بندی تیترها',
          description: 'اعمال استایل Heading 1 / Heading 2 و ساخت جدول ساختاریافته',
          toolName: 'word_format_document',
          targetScope: 'ابتدای سند و جداول گزارش',
          risk: 'safe',
          status: 'pending'
        });

        steps.push({
          id: `${planId}-step-final`,
          order: steps.length + 1,
          title: 'اعتبارسنجی و تأیید کیفیت خروجی ورد',
          description: 'بررسی عدم حذف ناخواسته محتوا و گزارش نتیجه',
          toolName: 'word_verify_structure',
          targetScope: 'گزارش نهایی',
          risk: 'safe',
          status: 'pending'
        });
      }

      const hasConfirmStep = steps.some((s) => s.risk !== 'safe');

      return {
        id: planId,
        userPrompt: promptText,
        createdAt: Date.now(),
        steps,
        currentStepIndex: 0,
        requiresUserConfirmation: hasConfirmStep,
        status: hasConfirmStep ? 'awaiting_approval' : 'in_progress'
      };
    },
    [officeContext.host]
  );

  // Send a new prompt
  const sendMessage = useCallback(
    async (promptText: string) => {
      if (!promptText.trim() || isProcessing) return;

      const userMsg: ChatMessage = {
        id: `msg-${Date.now()}`,
        role: 'user',
        timestamp: Date.now(),
        content: promptText
      };

      setMessages((prev) => [...prev, userMsg]);
      setIsProcessing(true);

      const plan = analyzePromptAndCreatePlan(promptText);
      setActivePlan(plan);

      // Call server proxy for live LLM response
      let llmTextResponse = '';
      try {
        const response = await fetch('/api/agent/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            provider: selectedProviderId,
            model: selectedModel,
            messages: [
              {
                role: 'system',
                content:
                  'شما دستیار ارشد هوش مصنوعی و مجری عملیات در آفیس ۲۰۲۴ (اکسل و ورد) هستید. به زبان فارسی رسمی، دقیق و گام‌به‌گام به کاربر توضیح دهید چه مراحلی اجرا می‌شود و چه توصیه‌هایی برای ساختار سند وجود دارد.'
              },
              ...messages.map((m) => ({ role: m.role, content: m.content })),
              { role: 'user', content: promptText }
            ]
          })
        });

        const data = await response.json();
        if (data.success && data.result?.text) {
          llmTextResponse = data.result.text;
        }
      } catch (err) {
        console.warn('[Proxy Chat Network Error]', err);
      }

      const assistantMsg: ChatMessage = {
        id: `msg-${Date.now() + 1}`,
        role: 'assistant',
        timestamp: Date.now(),
        content:
          llmTextResponse ||
          `برنامه اجرایی مرحله‌به‌مرحله برای درخواست شما تنظیم گردید. در حال اجرای خودکار ابزارها و ثبت سوابق اعتبارسنجی...`,
        plan
      };

      setMessages((prev) => [...prev, assistantMsg]);

      // Execute through orchestrator
      await orchestratorRef.current.executePlan(plan, {
        onStepProgress: (stepId, status, summary) => {
          setActivePlan((prev) => {
            if (!prev) return prev;
            const updated = prev.steps.map((s) =>
              s.id === stepId ? { ...s, status, resultSummary: summary || s.resultSummary } : s
            );
            return { ...prev, steps: updated };
          });
        },
        onAuditLog: (entry) => {
          setAuditLogs((prev) => [entry, ...prev]);
        },
        onRequestApproval: (request) => {
          setActiveApproval(request);
          setIsProcessing(false);
        },
        onPlanComplete: (completedPlan, finalReport) => {
          setActivePlan(completedPlan);
          setIsProcessing(false);

          if (promptText.includes('داشبورد')) {
            // Set mock/real dashboard for UI preview
            setLatestDashboard({
              success: true,
              dashboardSheetName: 'داشبورد_مدیریتی',
              kpiMetrics: [
                { label: 'مجموع کل مبالغ', value: '۴,۱۴۰,۰۰۰,۰۰۰ ریال', subtext: 'کل هزینه‌ها' },
                { label: 'میانگین هر تراکنش', value: '۷۹,۶۰۰,۰۰۰ ریال', subtext: 'میانگین اقلام' },
                { label: 'تعداد کل رکوردها', value: '۴۵ مورد', subtext: 'تراکنش‌های فعال' }
              ],
              summaryTableAddress: 'داشبورد_مدیریتی!B9:E14',
              chartsCreated: ['نمودار ستونی مقایسه مبالغ', 'نمودار دایره‌ای سهم تراکنش‌ها']
            });
          }

          setMessages((prev) => [
            ...prev,
            {
              id: `msg-done-${Date.now()}`,
              role: 'assistant',
              timestamp: Date.now(),
              content: finalReport
            }
          ]);
        },
        onPlanError: (_failedPlan, error) => {
          setIsProcessing(false);
          setMessages((prev) => [
            ...prev,
            {
              id: `msg-err-${Date.now()}`,
              role: 'assistant',
              timestamp: Date.now(),
              content: `خطا در اجرای یکی از ابزارهای سیستم:\n${error}\nعملیات متوقف گردید تا از دستکاری اشتباه فایل جلوگیری شود.`
            }
          ]);
        }
      });
    },
    [isProcessing, analyzePromptAndCreatePlan, selectedProviderId, selectedModel, messages]
  );

  // User approved sensitive action
  const handleApproveAction = useCallback(async () => {
    if (!activeApproval || !activePlan) return;

    const approvedStepId = activeApproval.stepId;
    setActiveApproval(null);
    setIsProcessing(true);

    const step = activePlan.steps.find((s) => s.id === approvedStepId);
    if (step) {
      const tool = TOOL_REGISTRY[step.toolName];
      if (tool) {
        try {
          const res = await tool.execute({}, {});
          step.status = 'completed';
          step.resultSummary = res.summary;

          const auditEntry: AuditLogEntry = {
            id: `audit-${Date.now()}`,
            timestamp: Date.now(),
            toolName: step.toolName,
            risk: step.risk,
            status: 'success',
            details: res.summary,
            affectedRangeOrTarget: res.affectedScope
          };

          setAuditLogs((prev) => [auditEntry, ...prev]);

          // Resume remaining steps
          await orchestratorRef.current.executePlan(activePlan, {
            onStepProgress: (stepId, status, summary) => {
              setActivePlan((prev) => {
                if (!prev) return prev;
                const updated = prev.steps.map((s) =>
                  s.id === stepId ? { ...s, status, resultSummary: summary || s.resultSummary } : s
                );
                return { ...prev, steps: updated };
              });
            },
            onAuditLog: (entry) => {
              setAuditLogs((prev) => [entry, ...prev]);
            },
            onRequestApproval: (request) => {
              setActiveApproval(request);
              setIsProcessing(false);
            },
            onPlanComplete: (completedPlan, finalReport) => {
              setActivePlan(completedPlan);
              setIsProcessing(false);
              setMessages((prev) => [
                ...prev,
                {
                  id: `msg-done-${Date.now()}`,
                  role: 'assistant',
                  timestamp: Date.now(),
                  content: finalReport
                }
              ]);
            },
            onPlanError: (_failedPlan, error) => {
              setIsProcessing(false);
              setMessages((prev) => [
                ...prev,
                {
                  id: `msg-err-${Date.now()}`,
                  role: 'assistant',
                  timestamp: Date.now(),
                  content: `خطا در ادامه فرآیند: ${error}`
                }
              ]);
            }
          });
        } catch (err) {
          console.error('[handleApproveAction] tool execution error:', err);
          setIsProcessing(false);
        }
      }
    }
  }, [activeApproval, activePlan]);

  // User rejected sensitive action
  const handleRejectAction = useCallback(() => {
    if (!activeApproval || !activePlan) return;

    const rejectedStepId = activeApproval.stepId;
    setActiveApproval(null);

    const updatedSteps = activePlan.steps.map((s) => {
      if (s.id === rejectedStepId) {
        return {
          ...s,
          status: 'cancelled' as const,
          executionError: 'عملیات توسط کاربر لغو گردید.'
        };
      }
      return s;
    });

    const auditEntry: AuditLogEntry = {
      id: `audit-${Date.now()}`,
      timestamp: Date.now(),
      toolName: activeApproval.title,
      risk: activeApproval.risk,
      status: 'rejected',
      details: 'عملیات حساس توسط کاربر لغو شد و هیچ تغییری روی فایل اعمال نگردید.',
      affectedRangeOrTarget: activeApproval.affectedScope
    };

    setAuditLogs((prev) => [auditEntry, ...prev]);

    setActivePlan({
      ...activePlan,
      steps: updatedSteps,
      status: 'failed'
    });

    setMessages((prev) => [
      ...prev,
      {
        id: `msg-rej-${Date.now()}`,
        role: 'assistant',
        timestamp: Date.now(),
        content: `عملیات طبق دستور شما لغو شد. هیچ داده‌ای از فایل حذف یا بازنویسی نشد.`
      }
    ]);
  }, [activeApproval, activePlan]);

  // Update provider config
  const updateProviderConfig = useCallback(
    (id: ProviderType, updates: Partial<ProviderConfig>) => {
      setProviders((prev) =>
        prev.map((p) => (p.id === id ? { ...p, ...updates } : p))
      );
    },
    []
  );

  return {
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
  };
}
