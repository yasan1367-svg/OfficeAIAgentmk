/**
 * Office AI Agent - Base AI Model Adapter Interface
 */

import { AgentModelRequest, AgentModelResponse, ProviderType } from '../types/provider.ts';

export interface AIModelAdapter {
  providerType: ProviderType;
  sendMessage(request: AgentModelRequest, customApiKey?: string): Promise<AgentModelResponse>;
  checkHealth(customApiKey?: string): Promise<{ isHealthy: boolean; latencyMs?: number; error?: string }>;
}
