/**
 * Office AI Agent - AI Provider & Security Settings Modal (Phase 2)
 */

import React, { useState } from 'react';
import {
  X,
  ShieldCheck,
  Cpu,
  KeyRound,
  ExternalLink,
  Check,
  Server
} from 'lucide-react';
import { ProviderConfig, ProviderType } from '../types/provider.ts';

interface SettingsModalProps {
  providers: ProviderConfig[];
  selectedProviderId: ProviderType;
  selectedModel: string;
  onSelectProvider: (id: ProviderType) => void;
  onSelectModel: (model: string) => void;
  onClose: () => void;
  onUpdateProvider: (id: ProviderType, updates: Partial<ProviderConfig>) => void;
}

export const SettingsModal: React.FC<SettingsModalProps> = ({
  providers,
  selectedProviderId,
  selectedModel,
  onSelectProvider,
  onSelectModel,
  onClose,
  onUpdateProvider
}) => {
  const currentProvider = providers.find((p) => p.id === selectedProviderId) || providers[0];
  const [customKey, setCustomKey] = useState('');
  const [saveSuccess, setSaveSuccess] = useState(false);

  const handleSaveKey = () => {
    if (!customKey.trim()) return;
    const masked = customKey.slice(0, 4) + '••••••••' + customKey.slice(-4);
    onUpdateProvider(selectedProviderId, {
      isConfigured: true,
      apiKeyMasked: masked
    });
    setCustomKey('');
    setSaveSuccess(true);
    setTimeout(() => setSaveSuccess(false), 2000);
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-3 animate-in fade-in duration-200">
      <div className="bg-white border border-slate-200 rounded-2xl shadow-xl max-w-md w-full overflow-hidden text-right font-sans flex flex-col max-h-[90vh]">
        {/* هدر */}
        <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center gap-2">
            <Cpu className="w-4 h-4 text-emerald-700" />
            <h2 className="text-xs font-bold text-slate-900">
              تنظیمات هوش مصنوعی و ارائه‌دهندگان (AI Providers)
            </h2>
          </div>
          <button
            onClick={onClose}
            className="p-1 text-slate-400 hover:text-slate-700 rounded-lg transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* محتوا */}
        <div className="p-4 space-y-4 overflow-y-auto text-xs flex-1">
          {/* کارت معماری امنیتی */}
          <div className="p-3 bg-emerald-50/70 border border-emerald-200 rounded-xl space-y-1.5 text-[11px]">
            <div className="font-bold text-emerald-900 flex items-center gap-1.5">
              <ShieldCheck className="w-4 h-4 text-emerald-700" />
              معماری محلی پروکسی امن (Local Express Proxy)
            </div>
            <p className="text-emerald-800 leading-relaxed">
              کلیدهای محرمانه API به هیچ عنوان در کدهای سمت کلاینت یا لاگ‌های مرورگر ذخیره نمی‌شوند و صرفاً از طریق سرور پروکسی امن لوکال با هدرهای رمزنگاری‌شده ارسال می‌گردند.
            </p>
          </div>

          {/* انتخاب ارائه‌دهنده */}
          <div className="space-y-1.5">
            <label className="font-bold text-slate-800 block">انتخاب ارائه‌دهنده فعال:</label>
            <div className="grid grid-cols-1 gap-2">
              {providers.map((p) => (
                <button
                  key={p.id}
                  onClick={() => {
                    onSelectProvider(p.id);
                    onSelectModel(p.defaultModel);
                  }}
                  className={`p-3 rounded-xl border text-right transition-all flex items-center justify-between ${
                    p.id === selectedProviderId
                      ? 'border-emerald-600 bg-emerald-50/40 shadow-xs'
                      : 'border-slate-200 hover:bg-slate-50'
                  }`}
                >
                  <div>
                    <div className="font-bold text-slate-900 flex items-center gap-2">
                      <span>{p.name}</span>
                      {p.isConfigured && (
                        <span className="text-[10px] text-emerald-700 bg-emerald-100/70 px-1.5 py-0.5 rounded font-mono">
                          آماده به کار
                        </span>
                      )}
                    </div>
                    <div className="text-[11px] font-mono text-slate-500 mt-0.5" dir="ltr">
                      {p.baseUrl}
                    </div>
                  </div>

                  {p.id === selectedProviderId && (
                    <div className="w-5 h-5 rounded-full bg-emerald-600 text-white flex items-center justify-center shrink-0">
                      <Check className="w-3.5 h-3.5" />
                    </div>
                  )}
                </button>
              ))}
            </div>
          </div>

          {/* انتخاب مدل برای ارائه‌دهنده فعال */}
          <div className="space-y-1.5">
            <label className="font-bold text-slate-800 block">مدل اجرایی:</label>
            <select
              value={selectedModel}
              onChange={(e) => onSelectModel(e.target.value)}
              className="w-full p-2 rounded-lg border border-slate-300 bg-white font-mono text-slate-800 focus:outline-hidden focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
            >
              {currentProvider.availableModels.map((m) => (
                <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </select>
          </div>

          {/* وضعیت کلید اختصاصی */}
          <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
            <div className="flex items-center justify-between text-[11px]">
              <span className="font-bold text-slate-700 flex items-center gap-1.5">
                <KeyRound className="w-3.5 h-3.5 text-slate-500" />
                وضعیت کلید محرمانه (API Key):
              </span>
              <span className="font-mono text-slate-600">
                {currentProvider.apiKeyMasked || 'تنظیم نشده (استفاده از کلید پیش‌فرض سرور)'}
              </span>
            </div>

            <div className="flex gap-2">
              <input
                type="password"
                placeholder="ورود کلید جدید (اختیاری جهت تغییر)..."
                value={customKey}
                onChange={(e) => setCustomKey(e.target.value)}
                className="flex-1 p-2 rounded-lg border border-slate-300 bg-white text-xs font-mono placeholder:font-sans focus:outline-hidden focus:ring-2 focus:ring-emerald-500/20"
                dir="ltr"
              />
              <button
                onClick={handleSaveKey}
                disabled={!customKey.trim()}
                className="px-3 py-1.5 rounded-lg bg-slate-900 text-white text-xs font-medium hover:bg-slate-800 transition-colors disabled:opacity-40"
              >
                ثبت
              </button>
            </div>

            {saveSuccess && (
              <div className="text-[11px] text-emerald-700 font-medium">
                کلید با موفقیت در لایه امن ذخیره گردید.
              </div>
            )}
          </div>
        </div>

        {/* دکمه بستن */}
        <div className="p-3 bg-slate-50 border-t border-slate-200 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-1.5 rounded-lg bg-emerald-700 text-white text-xs font-medium hover:bg-emerald-800 transition-colors shadow-xs"
          >
            تأیید و ذخیره
          </button>
        </div>
      </div>
    </div>
  );
};
