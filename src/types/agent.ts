/**
 * Office AI Agent - Agent Core Execution & Planning Types
 */

export type ActionRiskLevel = 'safe' | 'confirm_required' | 'destructive';

export type StepStatus = 'pending' | 'awaiting_approval' | 'running' | 'completed' | 'failed' | 'cancelled';

export interface PlanStep {
  id: string;
  order: number;
  title: string;
  description: string;
  toolName: string;
  targetScope: string;
  risk: ActionRiskLevel;
  status: StepStatus;
  executionError?: string;
  resultSummary?: string;
}

export interface ExecutionPlan {
  id: string;
  userPrompt: string;
  createdAt: number;
  steps: PlanStep[];
  currentStepIndex: number;
  requiresUserConfirmation: boolean;
  status: 'planning' | 'awaiting_approval' | 'in_progress' | 'completed' | 'failed';
}

export interface ApprovalRequest {
  id: string;
  stepId: string;
  title: string;
  description: string;
  risk: ActionRiskLevel;
  affectedScope: string;
  itemCount?: number;
  isReversible: boolean;
  proposedChanges: string[];
}

export interface AuditLogEntry {
  id: string;
  timestamp: number;
  toolName: string;
  risk: ActionRiskLevel;
  status: 'success' | 'failed' | 'rejected';
  details: string;
  affectedRangeOrTarget?: string;
}

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: number;
  plan?: ExecutionPlan;
  isStreaming?: boolean;
}
