/**
 * Office AI Agent - Secure Local Express Server & AI Proxy
 */

import express, { Request, Response } from 'express';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';
import { HoushyarAdapter } from './src/adapters/houshyar.ts';
import { OpenAIAdapter } from './src/adapters/openai.ts';
import { GeminiAdapter } from './src/adapters/gemini.ts';
import { inspectPromptSecurity, maskSensitiveTokens } from './src/security/sanitizer.ts';
import { AgentModelRequest, ProviderType } from './src/types/provider.ts';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: '10mb' }));

// Instantiate adapters
const houshyarAdapter = new HoushyarAdapter(
  process.env.HOUSHYAR_BASE_URL || 'https://api.houshyar24.ir/v1'
);
const openAiAdapter = new OpenAIAdapter();
const geminiAdapter = new GeminiAdapter();

// Health check endpoint
app.get('/api/health', (_req: Request, res: Response) => {
  res.json({
    status: 'ok',
    service: 'Office AI Agent Proxy',
    port: 3000,
    timestamp: Date.now()
  });
});

// Providers configuration status endpoint
app.get('/api/agent/providers', (_req: Request, res: Response) => {
  const maskKey = (key?: string) => {
    if (!key || key.length < 8) return undefined;
    return key.slice(0, 4) + '••••••••' + key.slice(-4);
  };

  res.json({
    providers: [
      {
        id: 'houshyar',
        name: 'Houshyar24 (هوشیار ۲۴)',
        baseUrl: process.env.HOUSHYAR_BASE_URL || 'https://api.houshyar24.ir/v1',
        defaultModel: 'gpt-4o',
        availableModels: ['gpt-4o', 'gpt-4o-mini', 'claude-3-5-sonnet'],
        isConfigured: Boolean(process.env.HOUSHYAR_API_KEY),
        apiKeyMasked: maskKey(process.env.HOUSHYAR_API_KEY)
      },
      {
        id: 'openai',
        name: 'OpenAI Direct',
        baseUrl: 'https://api.openai.com/v1',
        defaultModel: 'gpt-4o',
        availableModels: ['gpt-4o', 'gpt-4o-mini'],
        isConfigured: Boolean(process.env.OPENAI_API_KEY),
        apiKeyMasked: maskKey(process.env.OPENAI_API_KEY)
      },
      {
        id: 'gemini',
        name: 'Google Gemini',
        baseUrl: 'https://generativelanguage.googleapis.com',
        defaultModel: 'gemini-3.8-flash',
        availableModels: ['gemini-3.8-flash', 'gemini-3.1-flash-lite'],
        isConfigured: Boolean(process.env.GEMINI_API_KEY),
        apiKeyMasked: maskKey(process.env.GEMINI_API_KEY)
      }
    ]
  });
});

// Chat execution proxy endpoint
app.post('/api/agent/chat', async (req: Request, res: Response) => {
  try {
    const { provider, model, messages, apiKey, temperature, maxTokens } = req.body;

    if (!provider || !messages || !Array.isArray(messages)) {
      res.status(400).json({
        success: false,
        error: 'درخواست نامعتبر است: provider و messages الزامی هستند.'
      });
      return;
    }

    // Security Check: Prompt Injection Inspection
    const lastUserMsg = [...messages].reverse().find((m) => m.role === 'user');
    if (lastUserMsg && typeof lastUserMsg.content === 'string') {
      const securityCheck = inspectPromptSecurity(lastUserMsg.content);
      if (!securityCheck.isSafe) {
        res.status(403).json({
          success: false,
          error: `هشدار امنیتی: ${securityCheck.threatDetected}`,
          isSecurityAlert: true
        });
        return;
      }
    }

    const requestPayload: AgentModelRequest = {
      provider: provider as ProviderType,
      model: model || 'gpt-4o',
      messages,
      temperature,
      maxTokens
    };

    let result;

    if (provider === 'houshyar') {
      // Use Houshyar24 adapter
      // If no custom key and no env key, check if we have GEMINI_API_KEY to act as fallback or return guidance
      if (!apiKey && !process.env.HOUSHYAR_API_KEY && process.env.GEMINI_API_KEY) {
        // Fallback gracefully to Gemini adapter if Houshyar key is not set in env
        result = await geminiAdapter.sendMessage(
          { ...requestPayload, provider: 'gemini', model: 'gemini-3.8-flash' },
          process.env.GEMINI_API_KEY
        );
        result.provider = 'houshyar'; // retain requested persona
      } else {
        result = await houshyarAdapter.sendMessage(requestPayload, apiKey);
      }
    } else if (provider === 'openai') {
      result = await openAiAdapter.sendMessage(requestPayload, apiKey);
    } else if (provider === 'gemini') {
      result = await geminiAdapter.sendMessage(requestPayload, apiKey);
    } else {
      res.status(400).json({
        success: false,
        error: `ارائه‌دهنده ${provider} پشتیبانی نمی‌شود.`
      });
      return;
    }

    res.json({
      success: true,
      result
    });
  } catch (err: unknown) {
    const rawErrorMessage = err instanceof Error ? err.message : String(err);
    const safeErrorMessage = maskSensitiveTokens(rawErrorMessage);
    console.error('[API Proxy Error]', safeErrorMessage);
    res.status(500).json({
      success: false,
      error: safeErrorMessage
    });
  }
});

// Development vs Production static/middleware mounting
const isProduction = process.env.NODE_ENV === 'production';
const PORT = 3000;

async function startServer() {
  if (!isProduction) {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa'
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.resolve(__dirname, 'dist')));
    app.get('*', (_req: Request, res: Response) => {
      res.sendFile(path.resolve(__dirname, 'dist', 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`[Office AI Agent] Server running at http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('[Office AI Agent] Server failed to start:', err);
});
