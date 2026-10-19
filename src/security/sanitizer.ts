/**
 * Office AI Agent - Advanced Security, Token Masking & Injection Defense (Phase 8 Hardened)
 */

const ADVANCED_INJECTION_PATTERNS = [
  /ignore\s+(all\s+)?(previous|prior|above)\s+instructions/i,
  /system\s+override/i,
  /you\s+are\s+now\s+in\s+developer\s+mode/i,
  /disregard\s+all\s+(rules|policies)/i,
  /send\s+(all\s+)?files?\s+to/i,
  /jailbreak/i,
  /pretend\s+you\s+are\s+an?\s+(unrestricted|unfiltered)/i,
  /DAN\s+mode/i,
  /دستورات\s+(قبلی|بالا)\s+را\s+نادیده\s+بگیر/i,
  /حذف\s+تمام\s+فایل‌ها/i,
  /ارسال\s+اطلاعات\s+به/i,
  /حالت\s+(توسعه‌دهنده|بدون\s+محدودیت)/i,
  /قوانین\s+را\s+دور\s+بزن/i,
  /فرمان\s+سیستمی\s+جدید/i
];

const SECRET_TOKEN_PATTERNS = [
  /sk-[a-zA-Z0-9_-]{20,}/g,
  /AIza[a-zA-Z0-9_-]{35}/g,
  /Bearer\s+[a-zA-Z0-9._-]{20,}/gi,
  /password\s*[:=]\s*["']?[^"'}\s]+/gi
];

export interface SanitizationResult {
  isSafe: boolean;
  sanitizedText: string;
  threatDetected?: string;
}

/**
 * Inspects untrusted text for malicious injection directives and disguised jailbreak commands
 */
export function inspectPromptSecurity(text: string): SanitizationResult {
  // Strip zero-width invisible characters often used to conceal injection attacks
  const normalizedText = text.replace(/[\u200B-\u200D\uFEFF]/g, '');

  for (const pattern of ADVANCED_INJECTION_PATTERNS) {
    if (pattern.test(normalizedText)) {
      return {
        isSafe: false,
        sanitizedText: text,
        threatDetected: `دستور غیرمجاز یا مشکوک به تزریق پرامپت شناسایی شد (${pattern.source}).`
      };
    }
  }

  return {
    isSafe: true,
    sanitizedText: text
  };
}

/**
 * Masks any API keys, tokens or secrets from strings, error messages or logs
 */
export function maskSensitiveTokens(content: string): string {
  let masked = content;
  SECRET_TOKEN_PATTERNS.forEach((pattern) => {
    masked = masked.replace(pattern, (match) => {
      if (match.length <= 8) return '••••••••';
      return match.slice(0, 4) + '••••••••' + match.slice(-4);
    });
  });
  return masked;
}

/**
 * Wraps file cell contents or paragraph texts in untrusted data boundaries
 * Enforces strict demarcation between instructions and document data.
 */
export function wrapUntrustedDocumentData(
  sourceName: string,
  rawContent: string
): string {
  const sanitizedContent = rawContent.replace(/<\/?untrusted_document_data>/gi, '');
  return [
    `<!-- BEGIN UNTRUSTED DATA: ${sourceName} -->`,
    `<untrusted_document_data source="${sourceName}">`,
    sanitizedContent,
    `</untrusted_document_data>`,
    `<!-- END UNTRUSTED DATA: ${sourceName} -->`
  ].join('\n');
}

/**
 * Sanitizes spreadsheet cell strings to prevent DDE/Formula Injection attacks
 */
export function sanitizeCellFormulaInjection(value: string): string {
  if (typeof value !== 'string') return value;
  const trimmed = value.trim();
  // If text starts with =, +, -, @, prefix with a single quote to neutralize formula execution in external exports
  if (/^[=+\-@]/.test(trimmed) && !/^=[A-Z]+[0-9]*\(/.test(trimmed)) {
    return `'${trimmed}`;
  }
  return value;
}
