/**
 * Office AI Agent - Plan View Component (Phase 2)
 */

import React from 'react';
import {
  CheckCircle2,
  Clock,
  AlertTriangle,
  ShieldAlert,
  Loader2,
  XCircle,
  Layers,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import { ExecutionPlan, PlanStep, ActionRiskLevel } from '../types/agent.ts';

interface PlanViewProps {
  plan: ExecutionPlan;
  onOpenApproval?: () => void;
}

export const PlanView: React.FC<PlanViewProps> = ({ plan, onOpenApproval }) => {
  const [isExpanded, setIsExpanded] = React.useState(true);

  const getRiskLabel = (risk: ActionRiskLevel) => {
    switch (risk) {
      case 'safe':
        return {
          text: 'ایمن (غیرمخرب)',
          badgeColor: 'text-emerald-700 bg-emerald-50 border-emerald-200'
        };
      case 'confirm_required':
        return {
          text: 'نیازمند تأیید',
          badgeColor: 'text-amber-800 bg-amber-50 border-amber-200'
        };
      case 'destructive':
        return {
          text: 'تخریب / بازنویسی',
          badgeColor: 'text-red-700 bg-red-50 border-red-200'
        };
    }
  };

  const getStatusIcon = (step: PlanStep) => {
    switch (step.status) {
      case 'completed':
        return <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />;
      case 'running':
        return <Loader2 className="w-4 h-4 text-blue-600 animate-spin shrink-0" />;
      case 'awaiting_approval':
        return <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />;
      case 'cancelled':
        return <XCircle className="w-4 h-4 text-slate-400 shrink-0" />;
      default:
        return <Clock className="w-4 h-4 text-slate-400 shrink-0" />;
    }
  };

  const completedCount = plan.steps.filter((s) => s.status === 'completed').length;
  const totalCount = plan.steps.length;
  const progressPercent = Math.round((completedCount / totalCount) * 100);

  return (
    <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs my-2.5">
      {/* هدر پلن */}
      <div
        onClick={() => setIsExpanded(!isExpanded)}
        className="px-3.5 py-2.5 bg-slate-50/80 border-b border-slate-100 flex items-center justify-between cursor-pointer hover:bg-slate-100/60 transition-colors"
      >
        <div className="flex items-center gap-2">
          <Layers className="w-4 h-4 text-slate-700" />
          <span className="text-xs font-bold text-slate-900">
            برنامه اجرایی مرحله‌به‌مرحله ({completedCount} از {totalCount})
          </span>
        </div>

        <div className="flex items-center gap-2">
          <span className="text-[11px] font-mono tabular-nums text-slate-500">
            {progressPercent}%
          </span>
          {isExpanded ? (
            <ChevronUp className="w-3.5 h-3.5 text-slate-400" />
          ) : (
            <ChevronDown className="w-3.5 h-3.5 text-slate-400" />
          )}
        </div>
      </div>

      {/* نوار پیشرفت */}
      <div className="w-full bg-slate-100 h-1">
        <div
          className="bg-emerald-600 h-1 transition-all duration-300"
          style={{ width: `${progressPercent}%` }}
        />
      </div>

      {/* فهرست مراحل */}
      {isExpanded && (
        <div className="p-3 space-y-2.5">
          {plan.steps.map((step) => {
            const riskInfo = getRiskLabel(step.risk);
            const isNeedApproval = step.risk !== 'safe' && step.status !== 'completed';

            return (
              <div
                key={step.id}
                className={`p-2.5 rounded-lg border text-xs transition-colors ${
                  step.status === 'completed'
                    ? 'bg-emerald-50/30 border-emerald-100'
                    : isNeedApproval
                    ? 'bg-amber-50/40 border-amber-200'
                    : 'bg-white border-slate-100'
                }`}
              >
                <div className="flex items-start justify-between gap-2">
                  <div className="flex items-start gap-2">
                    {getStatusIcon(step)}
                    <div>
                      <div className="font-semibold text-slate-800 flex items-center gap-1.5">
                        <span>{step.order}.</span>
                        <span>{step.title}</span>
                      </div>
                      <p className="text-slate-500 text-[11px] mt-0.5 leading-relaxed">
                        {step.description}
                      </p>
                    </div>
                  </div>

                  <span
                    className={`text-[10px] font-medium px-2 py-0.5 rounded-full border whitespace-nowrap shrink-0 ${riskInfo.badgeColor}`}
                  >
                    {riskInfo.text}
                  </span>
                </div>

                {/* محدوده هدف و گزارش مرحله */}
                <div className="mt-2 pt-2 border-t border-slate-100/80 flex items-center justify-between text-[11px] text-slate-500">
                  <span className="truncate max-w-[200px]">
                    محدوده: <span className="font-mono text-slate-700">{step.targetScope}</span>
                  </span>

                  {step.resultSummary && (
                    <span className="text-emerald-700 font-medium">{step.resultSummary}</span>
                  )}

                  {isNeedApproval && onOpenApproval && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onOpenApproval();
                      }}
                      className="text-amber-800 font-bold hover:underline"
                    >
                      مشاهده و تأیید
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};
