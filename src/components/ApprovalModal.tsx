/**
 * Office AI Agent - Sensitive Action Approval Gatekeeper Modal (Phase 2)
 */

import React from 'react';
import {
  AlertTriangle,
  ShieldAlert,
  CheckCircle,
  XCircle,
  FileSpreadsheet,
  FileWarning
} from 'lucide-react';
import { ApprovalRequest } from '../types/agent.ts';

interface ApprovalModalProps {
  request: ApprovalRequest;
  onApprove: () => void;
  onReject: () => void;
}

export const ApprovalModal: React.FC<ApprovalModalProps> = ({
  request,
  onApprove,
  onReject
}) => {
  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-3 animate-in fade-in duration-200">
      <div className="bg-white border border-slate-200 rounded-2xl shadow-xl max-w-sm w-full overflow-hidden text-right font-sans">
        {/* هدر مدال هشدار */}
        <div className="bg-amber-500/10 border-b border-amber-200/80 px-4 py-3 flex items-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-amber-500 text-white flex items-center justify-center shrink-0 shadow-xs">
            <AlertTriangle className="w-5 h-5" />
          </div>
          <div>
            <h3 className="text-xs font-bold text-amber-900 leading-tight">
              درخواست تأیید عملیات حساس (Gatekeeper)
            </h3>
            <p className="text-[11px] text-amber-700 mt-0.5">
              جهت حفاظت از فایل اصلی، مجوز کاربر الزامی است
            </p>
          </div>
        </div>

        {/* بدنه اطلاعات مدال */}
        <div className="p-4 space-y-3.5 text-xs">
          <div>
            <div className="font-bold text-slate-900 text-sm">{request.title}</div>
            <p className="text-slate-600 mt-1 leading-relaxed">
              {request.description}
            </p>
          </div>

          {/* جزئیات محدوده اثر */}
          <div className="bg-slate-50 border border-slate-200/80 rounded-xl p-3 space-y-2">
            <div className="flex items-center justify-between text-slate-700">
              <span className="text-slate-500">محدوده تحت تأثیر:</span>
              <span className="font-mono font-medium text-slate-900">{request.affectedScope}</span>
            </div>

            {request.itemCount !== undefined && (
              <div className="flex items-center justify-between text-slate-700">
                <span className="text-slate-500">تعداد رکوردهای ارزیابی‌شده:</span>
                <span className="font-mono font-bold text-slate-900 tabular-nums">
                  {request.itemCount} مورد
                </span>
              </div>
            )}

            <div className="flex items-center justify-between text-slate-700">
              <span className="text-slate-500">قابلیت بازگشت (Undo):</span>
              <span className={`font-semibold ${request.isReversible ? 'text-emerald-700' : 'text-red-700'}`}>
                {request.isReversible ? 'قابل بازگشت' : 'غیرقابل بازگشت مستقیم'}
              </span>
            </div>
          </div>

          {/* اقدامات پیشنهادی */}
          <div className="space-y-1.5">
            <div className="text-[11px] font-bold text-slate-700">تغییرات مدنظر در سند:</div>
            <ul className="space-y-1 text-slate-600 text-[11px]">
              {request.proposedChanges.map((change, idx) => (
                <li key={idx} className="flex items-start gap-1.5">
                  <span className="w-1 h-1 rounded-full bg-amber-500 mt-1.5 shrink-0" />
                  <span>{change}</span>
                </li>
              ))}
            </ul>
          </div>
        </div>

        {/* دکمه‌های اقدام */}
        <div className="p-3 bg-slate-50 border-t border-slate-100 flex items-center justify-end gap-2">
          <button
            onClick={onReject}
            className="px-3 py-1.5 rounded-lg border border-slate-300 text-slate-700 font-medium hover:bg-slate-100 transition-colors text-xs"
          >
            لغو و انصراف
          </button>
          <button
            onClick={onApprove}
            className="px-3.5 py-1.5 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white font-medium shadow-xs transition-colors text-xs flex items-center gap-1.5"
          >
            <CheckCircle className="w-3.5 h-3.5" />
            تأیید و اجرای عملیات
          </button>
        </div>
      </div>
    </div>
  );
};
