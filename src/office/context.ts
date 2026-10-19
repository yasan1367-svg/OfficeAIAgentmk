/// <reference types="office-js" />
/**
 * Office AI Agent - Office.js Runtime Initialization & Context Bridge
 */
import { OfficeContextInfo, OfficeHost } from '../types/office';

let cachedContext: OfficeContextInfo | null = null;

/**
 * Initializes Office.js runtime and detects active host application
 */
export async function initializeOfficeContext(): Promise<OfficeContextInfo> {
  if (cachedContext) {
    return cachedContext;
  }

  // Check if Office global is loaded
  if (typeof window !== 'undefined' && (window as unknown as { Office?: typeof Office }).Office) {
    try {
      const officeGlobal = (window as unknown as { Office: typeof Office }).Office;
      const info = await officeGlobal.onReady();

      let host: OfficeHost = 'standalone';
      if (info.host === officeGlobal.HostType.Excel) {
        host = 'excel';
      } else if (info.host === officeGlobal.HostType.Word) {
        host = 'word';
      }

      cachedContext = {
        isOfficeReady: true,
        host,
        platform: info.platform ? String(info.platform) : 'Office 2024 Desktop (x64)',
        documentTitle: getDocumentTitleFromOffice(host),
      };

      return cachedContext;
    } catch (err) {
      console.warn('Office.js initialization returned an error, falling back to standalone preview mode:', err);
      cachedContext = {
        isOfficeReady: false,
        host: 'standalone',
        error: err instanceof Error ? err.message : String(err),
      };
      return cachedContext;
    }
  }

  // Running outside Office (e.g. browser preview or dev test)
  cachedContext = {
    isOfficeReady: false,
    host: 'standalone',
    platform: 'Web Preview (Standalone)',
    documentTitle: 'حالت پیش‌نمایش توسعه (خارج از آفیس دسکتاپ)',
  };

  return cachedContext;
}

/**
 * Attempt to read document title safely based on host
 */
function getDocumentTitleFromOffice(host: OfficeHost): string {
  try {
    if (typeof window === 'undefined') return 'Office Document';
    const officeGlobal = (window as unknown as { Office?: { context?: { document?: { url?: string } } } }).Office;
    const url = officeGlobal?.context?.document?.url;
    if (url) {
      const parts = url.split(/[\\/]/);
      return parts[parts.length - 1] || (host === 'excel' ? 'Book1.xlsx' : 'Document1.docx');
    }
  } catch {
    // Ignore and return fallback
  }
  return host === 'excel' ? 'کاربرگ اکسل فعال (Office 2024)' : 'سند ورد فعال (Office 2024)';
}

/**
 * Get current cached Office context synchronously
 */
export function getOfficeContext(): OfficeContextInfo {
  return (
    cachedContext || {
      isOfficeReady: false,
      host: 'standalone',
      platform: 'Pending Initialization',
      documentTitle: 'در حال بررسی محیط...',
    }
  );
}
