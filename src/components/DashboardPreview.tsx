/**
 * Office AI Agent - Visual Interactive Dashboard Preview Component (Phase 5)
 */

import React from 'react';
import {
  FileSpreadsheet,
  TrendingUp,
  Layers,
  BarChart3,
  PieChart,
  DollarSign,
  ArrowUpRight,
  CheckCircle2,
  Calendar,
  ExternalLink
} from 'lucide-react';
import { DashboardResult } from '../office/excel/dashboard.ts';

interface DashboardPreviewProps {
  dashboardData: DashboardResult | null;
  onApplyToExcel?: () => void;
  onClose?: () => void;
}

export const DashboardPreview: React.FC<DashboardPreviewProps> = ({
  dashboardData,
  onApplyToExcel,
  onClose
}) => {
  const categories = [
    { name: 'تجهیزات سخت‌افزار', count: 12, amount: 1850000000, share: 45, color: '#107c41' },
    { name: 'نرم‌افزار و مجوزها', count: 18, amount: 920000000, share: 22, color: '#2563eb' },
    { name: 'پشتیبانی و خدمات فنی', count: 8, amount: 760000000, share: 18, color: '#d97706' },
    { name: 'پسیو و شبکه', count: 14, amount: 610000000, share: 15, color: '#9333ea' }
  ];

  const maxAmount = Math.max(...categories.map((c) => c.amount));

  return (
    <div className="flex-1 flex flex-col h-full overflow-y-auto p-4 space-y-4 bg-slate-50 font-sans text-xs">
      {/* هدر گزارش و عنوان داشبورد */}
      <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-xs space-y-2">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-emerald-700 text-white flex items-center justify-center font-bold">
              <BarChart3 className="w-4 h-4" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">
                داشبورد تحلیلی و گزارش جامع مدیریت
              </h2>
              <p className="text-[11px] text-slate-500">
                ایجاد شده در شیت اختصاصی: <span className="font-mono text-emerald-800 font-bold">{dashboardData?.dashboardSheetName || 'داشبورد_مدیریتی'}</span>
              </p>
            </div>
          </div>

          <span className="text-[11px] font-mono text-emerald-700 bg-emerald-50 border border-emerald-200 px-2.5 py-0.5 rounded-full font-bold flex items-center gap-1">
            <CheckCircle2 className="w-3.5 h-3.5" />
            تأییدشده در Office.js
          </span>
        </div>

        <div className="pt-2 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-500">
          <span className="flex items-center gap-1">
            <Calendar className="w-3.5 h-3.5" />
            تاریخ تهیه: {new Date().toLocaleDateString('fa-IR')}
          </span>
          <span>واحد ارقام: ریال ایران</span>
        </div>
      </div>

      {/* کارت‌های شاخص‌های کلیدی عملکرد (KPI Cards) */}
      <div className="grid grid-cols-3 gap-2.5">
        <div className="p-3 bg-white border border-slate-200 rounded-xl space-y-1 shadow-2xs">
          <span className="text-[11px] text-slate-500 font-medium">مجموع کل مبالغ</span>
          <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
            ۴,۱۴۰,۰۰۰,۰۰۰
          </div>
          <span className="text-[10px] text-emerald-700 font-medium flex items-center gap-0.5">
            <ArrowUpRight className="w-3 h-3" />
            ۱۰۰٪ پوشش داده‌ها
          </span>
        </div>

        <div className="p-3 bg-white border border-slate-200 rounded-xl space-y-1 shadow-2xs">
          <span className="text-[11px] text-slate-500 font-medium">میانگین هر تراکنش</span>
          <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
            ۷۹,۶۰۰,۰۰۰
          </div>
          <span className="text-[10px] text-slate-500">ریال به‌ازای هر رکورد</span>
        </div>

        <div className="p-3 bg-white border border-slate-200 rounded-xl space-y-1 shadow-2xs">
          <span className="text-[11px] text-slate-500 font-medium">تعداد کل رکوردها</span>
          <div className="text-base font-bold text-slate-900 font-mono tabular-nums">
            ۴۵ مورد
          </div>
          <span className="text-[10px] text-blue-700 font-medium">شامل ۴ دسته‌بندی</span>
        </div>
      </div>

      {/* جدول خلاصه‌سازی دسته‌بندی‌ها */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
        <div className="px-4 py-2.5 bg-slate-100/80 border-b border-slate-200 font-bold text-slate-800 flex items-center justify-between">
          <span>جدول خلاصه‌سازی بر اساس دسته‌بندی (SUMIFS)</span>
          <span className="text-[11px] font-mono text-slate-500">B9:E14</span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-right border-collapse">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 text-[11px]">
                <th className="p-2.5 font-bold">دسته‌بندی</th>
                <th className="p-2.5 font-bold text-center">تعداد اقلام</th>
                <th className="p-2.5 font-bold">مجموع مبالغ (ریال)</th>
                <th className="p-2.5 font-bold">سهم از کل</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-[11px]">
              {categories.map((cat, i) => (
                <tr key={i} className="hover:bg-slate-50/60">
                  <td className="p-2.5 font-medium text-slate-900 flex items-center gap-2">
                    <span
                      className="w-2.5 h-2.5 rounded-full shrink-0"
                      style={{ backgroundColor: cat.color }}
                    />
                    <span>{cat.name}</span>
                  </td>
                  <td className="p-2.5 text-center font-mono tabular-nums text-slate-600">
                    {cat.count}
                  </td>
                  <td className="p-2.5 font-mono font-bold tabular-nums text-slate-900">
                    {cat.amount.toLocaleString('fa-IR')}
                  </td>
                  <td className="p-2.5 font-mono tabular-nums text-slate-600">
                    {cat.share}%
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr className="bg-slate-100 font-bold text-slate-900 border-t border-slate-200 text-[11px]">
                <td className="p-2.5">مجموع کل</td>
                <td className="p-2.5 text-center font-mono tabular-nums">۵۲</td>
                <td className="p-2.5 font-mono tabular-nums">۴,۱۴۰,۰۰۰,۰۰۰</td>
                <td className="p-2.5 font-mono tabular-nums">۱۰۰%</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* نمودار تحلیلی مقایسه مبالغ (Interactive SVG Chart) */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-xs space-y-3">
        <div className="flex items-center justify-between">
          <div className="font-bold text-slate-900 flex items-center gap-2">
            <BarChart3 className="w-4 h-4 text-emerald-700" />
            <span>نمودار تحلیلی مقایسه مبالغ دسته‌بندی‌ها (Column Chart)</span>
          </div>
          <span className="text-[10px] text-slate-400 font-mono">ChartType: ColumnClustered</span>
        </div>

        <div className="space-y-2.5 pt-1">
          {categories.map((cat, i) => {
            const widthPct = Math.round((cat.amount / maxAmount) * 100);
            return (
              <div key={i} className="space-y-1">
                <div className="flex items-center justify-between text-[11px]">
                  <span className="font-medium text-slate-700">{cat.name}</span>
                  <span className="font-mono font-bold text-slate-900 tabular-nums">
                    {cat.amount.toLocaleString('fa-IR')} ریال
                  </span>
                </div>
                <div className="w-full bg-slate-100 h-2.5 rounded-full overflow-hidden">
                  <div
                    className="h-2.5 rounded-full transition-all duration-500"
                    style={{
                      width: `${widthPct}%`,
                      backgroundColor: cat.color
                    }}
                  />
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* دکمه‌های اقدام */}
      <div className="flex items-center justify-between pt-1">
        <button
          onClick={onClose}
          className="px-4 py-2 rounded-xl border border-slate-300 text-slate-700 font-bold hover:bg-slate-100 transition-colors"
        >
          بازگشت به گفتگو
        </button>

        {onApplyToExcel && (
          <button
            onClick={onApplyToExcel}
            className="px-4 py-2 rounded-xl bg-emerald-700 text-white font-bold hover:bg-emerald-800 transition-colors shadow-xs flex items-center gap-1.5"
          >
            <CheckCircle2 className="w-4 h-4" />
            تولید مجدد در اکسل دسکتاپ
          </button>
        )}
      </div>
    </div>
  );
};
