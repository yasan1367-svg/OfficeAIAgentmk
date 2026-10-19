/**
 * Office AI Agent - Audit & Execution History Modal (Phase 2)
 */

import React from 'react';
import {
  X,
  History,
  CheckCircle2,
  XCircle,
  AlertTriangle,
  FileSpreadsheet
} from 'lucide-react';
import { AuditLogEntry } from '../types/agent.ts';

interface AuditLogModalProps {
  logs: AuditLogEntry[];
  onClose: () => void;
}

export const AuditLogModal: React.FC<AuditLogModalProps> = ({ logs, onClose }) => {
  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-3 animate-in fade-in duration-200">
      <div className="bg-white border border-slate-200 rounded-2xl shadow-xl max-w-md w-full overflow-hidden text-right font-sans flex flex-col max-h-[85vh]">
        {/* هدر */}
        <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center gap-2">
            <History className="w-4 h-4 text-emerald-700" />
            <h2 className="text-xs font-bold text-slate-900">
              سوابق اجرای عملیات و اعتبارسنجی (Audit Logs)
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
        <div className="p-4 space-y-3 overflow-y-auto text-xs flex-1">
          {logs.length === 0 ? (
            <div className="text-center py-8 text-slate-400 space-y-2">
              <History className="w-8 h-8 mx-auto stroke-1" />
              <p>هنوز عملیاتی روی فایل اجرا یا ثبت نشده است.</p>
            </div>
          ) : (
            logs.map((log) => {
              const dateStr = new Date(log.timestamp).toLocaleTimeString('fa-IR', {
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
              });

              return (
                <div
                  key={log.id}
                  className="p-3 rounded-xl border border-slate-200 bg-slate-50/50 space-y-1.5"
                >
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-1.5 font-bold text-slate-800">
                      {log.status === 'success' ? (
                        <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                      ) : log.status === 'rejected' ? (
                        <AlertTriangle className="w-4 h-4 text-amber-600" />
                      ) : (
                        <XCircle className="w-4 h-4 text-red-600" />
                      )}
                      <span>{log.toolName}</span>
                    </div>

                    <span className="font-mono text-[10px] text-slate-400 tabular-nums">
                      {dateStr}
                    </span>
                  </div>

                  <p className="text-slate-600 text-[11px] leading-relaxed">
                    {log.details}
                  </p>

                  {log.affectedRangeOrTarget && (
                    <div className="text-[10px] text-slate-500 font-mono pt-1 border-t border-slate-100 flex items-center justify-between">
                      <span>محدوده: {log.affectedRangeOrTarget}</span>
                      <span className="uppercase text-slate-400">ریسک: {log.risk}</span>
                    </div>
                  )}
                </div>
              );
            })
          )}
        </div>

        {/* پابرگ */}
        <div className="p-3 bg-slate-50 border-t border-slate-200 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-1.5 rounded-lg bg-slate-900 text-white text-xs font-medium hover:bg-slate-800 transition-colors"
          >
            بستن
          </button>
        </div>
      </div>
    </div>
  );
};
