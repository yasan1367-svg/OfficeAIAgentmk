/**
 * Office AI Agent - Multi-Provider Configuration Types
 */

export type ProviderType = 'houshyar' | 'openai' | 'gemini';

export interface ProviderConfig {
  id: ProviderType;
  name: string;
  baseUrl: string;
  defaultModel: string;
  availableModels: string[];
  isConfigured: boolean;
  requiresProxy: boolean;
  apiKeyMasked?: string;
}

export interface ModelMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

export interface AgentModelRequest {
  provider: ProviderType;
  model: string;
  messages: ModelMessage[];
  temperature?: number;
  maxTokens?: number;
}

export interface AgentModelResponse {
  text: string;
  usage?: {
    promptTokens: number;
    completionTokens: number;
    totalTokens: number;
  };
  provider: ProviderType;
  model: string;
}
