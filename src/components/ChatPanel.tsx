/**
 * Office AI Agent - Conversational Chat Panel (Phase 2)
 */

import React, { useState, useRef, useEffect } from 'react';
import {
  Send,
  Sparkles,
  Bot,
  User,
  Loader2,
  FileSpreadsheet,
  FileText,
  Trash2,
  ArrowDown
} from 'lucide-react';
import { ChatMessage, ExecutionPlan } from '../types/agent.ts';
import { OfficeContextInfo } from '../types/office.ts';
import { PlanView } from './PlanView.tsx';

interface ChatPanelProps {
  messages: ChatMessage[];
  isProcessing: boolean;
  officeContext: OfficeContextInfo;
  onSendMessage: (text: string) => void;
  onOpenApproval?: () => void;
}

export const ChatPanel: React.FC<ChatPanelProps> = ({
  messages,
  isProcessing,
  officeContext,
  onSendMessage,
  onOpenApproval
}) => {
  const [inputText, setInputText] = useState('');
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const isExcel = officeContext.host === 'excel' || officeContext.host === 'standalone';

  const quickSuggestions = isExcel
    ? [
        'یک داشبورد حرفه‌ای از اطلاعات این فایل بساز',
        'داده‌های تکراری را پیدا کن، اما قبل از حذف به من گزارش بده',
        'یک گزارش مدیریتی از هزینه‌ها و روند ماهانه تهیه کن',
        'فرمول‌های دارای خطا و مقادیر نامعتبر را بررسی کن'
      ]
    : [
        'این سند را از نظر نگارش، ساختار و تیترها اصلاح کن',
        'از این فایل یک گزارش رسمی با جدول نکات کلیدی بساز',
        'این سند را بدون حذف نکات مهم خلاصه کن'
      ];

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages, isProcessing]);

  const handleSubmit = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!inputText.trim() || isProcessing) return;
    onSendMessage(inputText.trim());
    setInputText('');
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto';
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  };

  const handleTextareaChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInputText(e.target.value);
    e.target.style.height = 'auto';
    e.target.style.height = `${Math.min(e.target.scrollHeight, 120)}px`;
  };

  return (
    <div className="flex-1 flex flex-col h-full overflow-hidden bg-slate-50/50">
      {/* پیام‌های چت */}
      <div className="flex-1 overflow-y-auto p-3.5 space-y-4">
        {messages.map((msg) => {
          const isUser = msg.role === 'user';

          return (
            <div
              key={msg.id}
              className={`flex items-start gap-2.5 max-w-[92%] ${
                isUser ? 'mr-auto flex-row-reverse' : 'ml-auto'
              }`}
            >
              {/* آیکون فرستنده */}
              <div
                className={`w-7 h-7 rounded-lg flex items-center justify-center shrink-0 shadow-2xs text-xs ${
                  isUser
                    ? 'bg-slate-800 text-white'
                    : isExcel
                    ? 'bg-emerald-700 text-white'
                    : 'bg-blue-700 text-white'
                }`}
              >
                {isUser ? <User className="w-4 h-4" /> : <Bot className="w-4 h-4" />}
              </div>

              {/* حباب پیام */}
              <div className="space-y-2 flex-1">
                <div
                  className={`p-3 rounded-2xl text-xs leading-relaxed transition-all shadow-xs ${
                    isUser
                      ? 'bg-slate-900 text-white rounded-tr-xs'
                      : 'bg-white text-slate-800 border border-slate-200/90 rounded-tl-xs'
                  }`}
                >
                  <p className="whitespace-pre-wrap font-sans">{msg.content}</p>
                </div>

                {/* نمایش پلن در صورت وجود */}
                {msg.plan && (
                  <PlanView plan={msg.plan} onOpenApproval={onOpenApproval} />
                )}

                <div className="text-[10px] text-slate-400 px-1 font-mono tabular-nums">
                  {new Date(msg.timestamp).toLocaleTimeString('fa-IR', {
                    hour: '2-digit',
                    minute: '2-digit'
                  })}
                </div>
              </div>
            </div>
          );
        })}

        {isProcessing && (
          <div className="flex items-center gap-2.5 text-xs text-slate-500 py-1">
            <div className="w-7 h-7 rounded-lg bg-emerald-700 text-white flex items-center justify-center shadow-xs">
              <Loader2 className="w-4 h-4 animate-spin" />
            </div>
            <div className="bg-white border border-slate-200 px-3 py-2 rounded-2xl shadow-xs flex items-center gap-2">
              <Sparkles className="w-3.5 h-3.5 text-emerald-600 animate-pulse" />
              <span>در حال ارزیابی درخواست، انتخاب ابزار و آماده‌سازی برنامه...</span>
            </div>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {/* پیشنهادات سریع برای کاربر */}
      {messages.length <= 2 && (
        <div className="px-3 pb-2 pt-1 border-t border-slate-200/60 bg-white/70">
          <div className="text-[11px] font-bold text-slate-500 mb-1.5 flex items-center gap-1.5">
            <Sparkles className="w-3 h-3 text-emerald-600" />
            پیشنهادات سریع برای شروع:
          </div>
          <div className="flex flex-wrap gap-1.5">
            {quickSuggestions.map((sug, i) => (
              <button
                key={i}
                onClick={() => onSendMessage(sug)}
                disabled={isProcessing}
                className="text-[11px] text-slate-700 bg-slate-100 hover:bg-slate-200 hover:text-slate-900 px-2.5 py-1 rounded-md transition-colors text-right leading-tight max-w-full truncate"
              >
                {sug}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* فیلد ورودی متن */}
      <div className="p-3 bg-white border-t border-slate-200 shrink-0">
        <form onSubmit={handleSubmit} className="flex items-end gap-2">
          <div className="flex-1 relative">
            <textarea
              ref={textareaRef}
              rows={1}
              value={inputText}
              onChange={handleTextareaChange}
              onKeyDown={handleKeyDown}
              disabled={isProcessing}
              placeholder={
                isExcel
                  ? 'دستور خود را برای اکسل وارد کنید (مثلاً: داشبورد فروش بساز)...'
                  : 'دستور خود را برای ورد بنویسید (مثلاً: نگارش را اصلاح کن)...'
              }
              className="w-full resize-none p-2.5 text-xs text-slate-800 placeholder:text-slate-400 bg-slate-50 border border-slate-300 rounded-xl focus:outline-hidden focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 focus:bg-white transition-all max-h-32"
            />
          </div>

          <button
            type="submit"
            disabled={!inputText.trim() || isProcessing}
            title="ارسال دستور (Enter)"
            className="p-2.5 rounded-xl bg-emerald-700 text-white hover:bg-emerald-800 disabled:opacity-40 disabled:hover:bg-emerald-700 transition-colors shadow-xs shrink-0"
          >
            <Send className="w-4 h-4 rotate-180" />
          </button>
        </form>

        <div className="flex items-center justify-between mt-1.5 text-[10px] text-slate-400 px-1">
          <span>کلید Enter برای ارسال، Shift + Enter برای خط جدید</span>
          <span className="font-mono">RTL / LTR خودکار</span>
        </div>
      </div>
    </div>
  );
};
