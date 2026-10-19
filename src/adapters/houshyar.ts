/**
 * Office AI Agent - Houshyar24 (OpenAI-compatible) Adapter
 */

import { AIModelAdapter } from './base.ts';
import { AgentModelRequest, AgentModelResponse, ProviderType } from '../types/provider.ts';

export class HoushyarAdapter implements AIModelAdapter {
  providerType: ProviderType = 'houshyar';
  private defaultBaseUrl: string;

  constructor(baseUrl: string = 'https://api.houshyar24.ir/v1') {
    this.defaultBaseUrl = baseUrl;
  }

  async sendMessage(
    request: AgentModelRequest,
    customApiKey?: string
  ): Promise<AgentModelResponse> {
    const apiKey = customApiKey || process.env.HOUSHYAR_API_KEY || '';
    const endpoint = `${this.defaultBaseUrl}/chat/completions`;

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000); // 30s timeout

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
        throw new Error(
          `Houshyar24 API error [${response.status}]: ${errorBody || response.statusText}`
        );
      }

      const data = await response.json();
      const content =
        data.choices?.[0]?.message?.content ||
        'پاسخی از مدل دریافت نشد.';

      return {
        text: content,
        provider: 'houshyar',
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
        throw new Error('زمان پاسخ‌دهی به سرویس هوشیار ۲۴ بیش از ۳۰ ثانیه به طول انجامید (Timeout).');
      }
      throw err;
    }
  }

  async checkHealth(customApiKey?: string): Promise<{ isHealthy: boolean; latencyMs?: number; error?: string }> {
    const startTime = Date.now();
    try {
      // Light verification with small model request
      const res = await this.sendMessage(
        {
          provider: 'houshyar',
          model: 'gpt-4o-mini',
          messages: [{ role: 'user', content: 'ping' }],
          maxTokens: 5
        },
        customApiKey
      );

      return {
        isHealthy: true,
        latencyMs: Date.now() - startTime
      };
    } catch (err) {
      return {
        isHealthy: false,
        error: err instanceof Error ? err.message : String(err)
      };
    }
  }
}
