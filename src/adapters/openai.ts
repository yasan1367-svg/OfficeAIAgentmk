/**
 * Office AI Agent - Direct OpenAI Standard Adapter
 */

import { AIModelAdapter } from './base.ts';
import { AgentModelRequest, AgentModelResponse, ProviderType } from '../types/provider.ts';

export class OpenAIAdapter implements AIModelAdapter {
  providerType: ProviderType = 'openai';
  private baseUrl: string;

  constructor(baseUrl: string = 'https://api.openai.com/v1') {
    this.baseUrl = baseUrl;
  }

  async sendMessage(
    request: AgentModelRequest,
    customApiKey?: string
  ): Promise<AgentModelResponse> {
    const apiKey = customApiKey || process.env.OPENAI_API_KEY || '';
    if (!apiKey) {
      throw new Error('کلید دسترسی OpenAI تنظیم نشده است. لطفاً در تنظیمات کلید معتبر وارد نمایید.');
    }

    const endpoint = `${this.baseUrl}/chat/completions`;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);

    try {
      const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${apiKey}`
        },
        body: JSON.stringify({
          model: request.model || 'gpt-4o',
          messages: request.messages,
          temperature: request.temperature ?? 0.3,
          max_tokens: request.maxTokens ?? 2048
        }),
        signal: controller.signal
      });

      clearTimeout(timeout);

      if (!response.ok) {
        const errorBody = await response.text();
        throw new Error(`OpenAI API error [${response.status}]: ${errorBody}`);
      }

      const data = await response.json();
      return {
        text: data.choices?.[0]?.message?.content || '',
        provider: 'openai',
        model: request.model,
        usage: data.usage
          ? {
              promptTokens: data.usage.prompt_tokens || 0,
              completionTokens: data.usage.completion_tokens || 0,
              totalTokens: data.usage.total_tokens || 0
            }
          : undefined
      };
    } catch (err: unknown) {
      clearTimeout(timeout);
      if (err instanceof Error && err.name === 'AbortError') {
        throw new Error('زمان ارتباط با OpenAI به پایان رسید (Timeout).');
      }
      throw err;
    }
  }

  async checkHealth(customApiKey?: string): Promise<{ isHealthy: boolean; latencyMs?: number; error?: string }> {
    const startTime = Date.now();
    try {
      await this.sendMessage(
        {
          provider: 'openai',
          model: 'gpt-4o-mini',
          messages: [{ role: 'user', content: 'ping' }],
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
