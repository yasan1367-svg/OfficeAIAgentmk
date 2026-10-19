/**
 * Office AI Agent - Top Navigation Header (3-Zone Contract)
 */

import React from 'react';
import {
  FileSpreadsheet,
  FileText,
  Terminal,
  Settings2,
  History,
  Shield,
  RefreshCw
} from 'lucide-react';
import { OfficeContextInfo } from '../types/office.ts';
import { ProviderConfig, ProviderType } from '../types/provider.ts';

interface HeaderProps {
  officeContext: OfficeContextInfo;
  activeProvider: ProviderConfig;
  selectedModel: string;
  onOpenSettings: () => void;
  onOpenAuditLogs: () => void;
  onRefreshContext: () => void;
  isRefreshing: boolean;
}

export const Header: React.FC<HeaderProps> = ({
  officeContext,
  activeProvider,
  selectedModel,
  onOpenSettings,
  onOpenAuditLogs,
  onRefreshContext,
  isRefreshing
}) => {
  const getHostInfo = () => {
    if (officeContext.host === 'excel') {
      return {
        label: 'Excel 2024',
        icon: <FileSpreadsheet className="w-4 h-4 text-emerald-600" />,
        color: 'text-emerald-700'
      };
    }
    if (officeContext.host === 'word') {
      return {
        label: 'Word 2024',
        icon: <FileText className="w-4 h-4 text-blue-600" />,
        color: 'text-blue-700'
      };
    }
    return {
      label: 'Office Dev Preview',
      icon: <Terminal className="w-4 h-4 text-slate-500" />,
      color: 'text-slate-600'
    };
  };

  const hostInfo = getHostInfo();

  return (
    <header className="bg-white border-b border-slate-200 px-3.5 py-2.5 shrink-0">
      <div className="flex items-center justify-between gap-3">
        {/* Zone 1: Single text element wordmark */}
        <div className="flex items-center gap-2 shrink-0">
          <div className="w-7 h-7 rounded-lg bg-emerald-700 text-white flex items-center justify-center font-bold text-xs shadow-xs">
            OA
          </div>
          <span className="font-bold text-sm text-slate-900 tracking-tight whitespace-nowrap">
            Office AI Agent
          </span>
        </div>

        {/* Zone 2: Document Context & Active Model Info */}
        <div className="hidden sm:flex items-center gap-2 text-xs text-slate-500 truncate">
          <span className="flex items-center gap-1.5 font-medium text-slate-700 truncate max-w-[130px]">
            {hostInfo.icon}
            <span className="truncate">{officeContext.documentTitle}</span>
          </span>
          <span aria-hidden="true" className="text-slate-300">·</span>
          <span className="font-mono text-[11px] text-slate-600 truncate">
            {activeProvider.name.split(' ')[0]} ({selectedModel})
          </span>
        </div>

        {/* Zone 3: Actions (Audit Logs, Settings, Refresh) */}
        <div className="flex items-center gap-1 shrink-0">
          <button
            onClick={onOpenAuditLogs}
            title="سوابق اجرای عملیات و اعتبارسنجی"
            className="p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors"
          >
            <History className="w-4 h-4" />
          </button>

          <button
            onClick={onOpenSettings}
            title="تنظیمات ارائه‌دهندگان هوش مصنوعی و مدل‌ها"
            className="p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors"
          >
            <Settings2 className="w-4 h-4" />
          </button>

          <button
            onClick={onRefreshContext}
            disabled={isRefreshing}
            title="بررسی مجدد اتصال Office.js"
            className="p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors"
          >
            <RefreshCw className={`w-4 h-4 ${isRefreshing ? 'animate-spin text-emerald-600' : ''}`} />
          </button>
        </div>
      </div>

      {/* Mobile context line */}
      <div className="flex sm:hidden items-center justify-between mt-2 pt-1.5 border-t border-slate-100 text-xs">
        <span className="flex items-center gap-1.5 text-slate-700 truncate max-w-[200px]">
          {hostInfo.icon}
          <span className="truncate">{officeContext.documentTitle}</span>
        </span>
        <span className="text-[11px] font-mono text-slate-500">
          {activeProvider.name.split(' ')[0]}
        </span>
      </div>
    </header>
  );
};
