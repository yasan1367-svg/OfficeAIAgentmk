/**
 * Office AI Agent - Persian Typography & Proofreading Engine (Phase 6)
 */

export interface ProofreadingDiff {
  original: string;
  corrected: string;
  changesCount: number;
  fixedSnippets: Array<{ before: string; after: string; reason: string }>;
}

/**
 * Applies Persian orthography, ZWNJ (نیم‌فاصله) and typography rules
 */
export function proofreadPersianText(rawText: string): ProofreadingDiff {
  let text = rawText;
  const fixedSnippets: ProofreadingDiff['fixedSnippets'] = [];

  // Rule 1: Prefixes 'می' and 'نمی'
  const prefixRegex = /(می|نمی)(?:\s+|\u200C\s+)([آ-ی]+)/g;
  let match;
  while ((match = prefixRegex.exec(text)) !== null) {
    const before = match[0];
    const after = `${match[1]}\u200C${match[2]}`;
    if (!fixedSnippets.some((f) => f.before === before)) {
      fixedSnippets.push({
        before,
        after,
        reason: 'اصلاح نیم‌فاصله پیشوند «می / نمی‌»'
      });
    }
  }
  text = text.replace(prefixRegex, '$1\u200C$2');

  // Rule 2: Suffixes 'ها', 'های', 'تر', 'ترین'
  const suffixRegex = /([آ-ی]+)\s+(ها|های|تر|ترین)(?=\s|[،.؛:!؟]|$)/g;
  while ((match = suffixRegex.exec(text)) !== null) {
    const before = match[0];
    const after = `${match[1]}\u200C${match[2]}`;
    if (!fixedSnippets.some((f) => f.before === before)) {
      fixedSnippets.push({
        before,
        after,
        reason: 'اصلاح نیم‌فاصله پسوندها'
      });
    }
  }
  text = text.replace(suffixRegex, '$1\u200C$2');

  // Rule 3: Spacing before punctuation
  const spaceBeforePunctRegex = /\s+([،.؛:!؟])/g;
  while ((match = spaceBeforePunctRegex.exec(text)) !== null) {
    const before = match[0];
    const after = match[1];
    if (!fixedSnippets.some((f) => f.before === before)) {
      fixedSnippets.push({
        before,
        after,
        reason: 'حذف فاصله اضافه قبل از علامت نگارشی'
      });
    }
  }
  text = text.replace(spaceBeforePunctRegex, '$1');

  // Rule 4: Spacing after punctuation
  const spaceAfterPunctRegex = /([،.؛:!؟])([آ-یA-Za-z0-9])/g;
  while ((match = spaceAfterPunctRegex.exec(text)) !== null) {
    const before = match[0];
    const after = `${match[1]} ${match[2]}`;
    if (!fixedSnippets.some((f) => f.before === before)) {
      fixedSnippets.push({
        before,
        after,
        reason: 'درج فاصله استاندارد بعد از علامت نگارشی'
      });
    }
  }
  text = text.replace(spaceAfterPunctRegex, '$1 $2');

  return {
    original: rawText,
    corrected: text,
    changesCount: fixedSnippets.length,
    fixedSnippets
  };
}
