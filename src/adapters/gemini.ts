/**
 * Office AI Agent - Google Gemini Adapter (@google/genai SDK)
 */

import { GoogleGenAI } from '@google/genai';
import { AIModelAdapter } from './base.ts';
import { AgentModelRequest, AgentModelResponse, ProviderType } from '../types/provider.ts';

export class GeminiAdapter implements AIModelAdapter {
  providerType: ProviderType = 'gemini';

  async sendMessage(
    request: AgentModelRequest,
    customApiKey?: string
  ): Promise<AgentModelResponse> {
    const apiKey = customApiKey || process.env.GEMINI_API_KEY || '';
    if (!apiKey) {
      throw new Error('کلید دسترسی Gemini API موجود نیست. لطفاً متغیر محیطی GEMINI_API_KEY را بررسی کنید.');
    }

    const ai = new GoogleGenAI({ apiKey });
    const modelName = request.model || 'gemini-3.8-flash';

    // Separate system message if provided
    const systemMsg = request.messages.find((m) => m.role === 'system');
    const userOrAssistantMessages = request.messages.filter((m) => m.role !== 'system');

    // Combine contents
    const contents = userOrAssistantMessages.map((m) => ({
      role: m.role === 'assistant' ? 'model' : 'user',
      parts: [{ text: m.content }]
    }));

    // If only user message or simple list
    const response = await ai.models.generateContent({
      model: modelName,
      contents: contents.length > 0 ? contents : [{ role: 'user', parts: [{ text: 'سلام' }] }],
      config: {
        systemInstruction: systemMsg ? systemMsg.content : undefined,
        temperature: request.temperature ?? 0.3
      }
    });

    const text = response.text || '';

    return {
      text,
      provider: 'gemini',
      model: modelName,
      usage: response.usageMetadata
        ? {
            promptTokens: response.usageMetadata.promptTokenCount || 0,
            completionTokens: response.usageMetadata.candidatesTokenCount || 0,
            totalTokens: response.usageMetadata.totalTokenCount || 0
          }
        : undefined
    };
  }

  async checkHealth(customApiKey?: string): Promise<{ isHealthy: boolean; latencyMs?: number; error?: string }> {
    const startTime = Date.now();
    try {
      await this.sendMessage(
        {
          provider: 'gemini',
          model: 'gemini-3.8-flash',
          messages: [{ role: 'user', content: 'test ping' }],
          maxTokens: 5
        },
        customApiKey
      );
      return { isHealthy: true, latencyMs: Date.now() - startTime };
    } catch (err) {
      return { isHealthy: false, error: err instanceof Error ? err.message : String(err) };
    }
  }
}
