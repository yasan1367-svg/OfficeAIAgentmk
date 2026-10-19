/// <reference types="office-js" />
/**
 * Office AI Agent - Excel Reader & Deep Inspection Module (Phase 4)
 */

import {
  ExcelInspectionReport,
  ExcelWorkbookStructure,
  ColumnMetadata,
  ColumnDataType,
  FormulaErrorFinding,
  DuplicateRowFinding
} from '../../types/office.ts';

const KNOWN_FORMULA_ERRORS = [
  '#N/A',
  '#VALUE!',
  '#REF!',
  '#DIV/0!',
  '#NUM!',
  '#NAME?',
  '#NULL!',
  '#SPILL!',
  '#CALC!'
];

/**
 * Inspects all sheets and tables in the active Excel workbook
 */
export async function getWorkbookStructure(): Promise<ExcelWorkbookStructure> {
  if (typeof window !== 'undefined' && (window as unknown as { Excel?: typeof Excel }).Excel) {
    try {
      const ExcelGlobal = (window as unknown as { Excel: typeof Excel }).Excel;
      return await ExcelGlobal.run(async (context) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        const activeSheet = workbook.worksheets.getActiveWorksheet();
        const tables = workbook.tables;

        worksheets.load('items/name, items/id, items/visibility');
        activeSheet.load('name');
        tables.load('items/name, items/id, items/worksheet/name');

        await context.sync();

        const sheetSummaries = worksheets.items.map((ws) => ({
          name: ws.name,
          id: ws.id,
          visibility: ws.visibility
        }));

        const tableSummaries = tables.items.map((tb) => ({
          name: tb.name,
          id: tb.id,
          sheetName: tb.worksheet?.name || 'Unknown',
          headerNames: [],
          rowCount: 0
        }));

        return {
          documentName: getSafeDocumentName(),
          activeSheetName: activeSheet.name,
          sheets: sheetSummaries,
          tables: tableSummaries
        };
      });
    } catch (err) {
      console.warn('[ExcelReader] getWorkbookStructure error, using fallback:', err);
    }
  }

  // Standalone fallback
  return {
    documentName: 'فایل مالی_شرکت_۱۴۰۳.xlsx',
    activeSheetName: 'فروش و هزینه‌ها',
    sheets: [
      { name: 'فروش و هزینه‌ها', id: 'ws-1', visibility: 'Visible' },
      { name: 'موجودی انبار', id: 'ws-2', visibility: 'Visible' },
      { name: 'خلاصه بودجه', id: 'ws-3', visibility: 'Hidden' }
    ],
    tables: [
      {
        name: 'Table_Transactions',
        id: 'tbl-1',
        sheetName: 'فروش و هزینه‌ها',
        headerNames: ['شناسه', 'تاریخ', 'دسته‌بندی', 'مبلغ'],
        rowCount: 45
      }
    ]
  };
}

/**
 * Inspects active worksheet data, columns, headers, formula errors and duplicates
 */
export async function inspectActiveWorksheetData(
  sampleLimit: number = 100
): Promise<ExcelInspectionReport> {
  if (typeof window !== 'undefined' && (window as unknown as { Excel?: typeof Excel }).Excel) {
    try {
      const ExcelGlobal = (window as unknown as { Excel: typeof Excel }).Excel;
      return await ExcelGlobal.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load('name');
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load('address, rowCount, columnCount, values, formulas, numberFormat');

        await context.sync();

        if (usedRange.isNullObject || usedRange.rowCount === 0) {
          return createEmptyReport(sheet.name);
        }

        const rowCount = usedRange.rowCount;
        const columnCount = usedRange.columnCount;
        const rawValues = usedRange.values as unknown[][];
        const rawFormulas = usedRange.formulas as unknown[][];
        const rawFormats = usedRange.numberFormat as unknown[][];

        return analyzeRangeData(
          sheet.name,
          usedRange.address,
          rowCount,
          columnCount,
          rawValues,
          rawFormulas,
          rawFormats,
          sampleLimit
        );
      });
    } catch (err) {
      console.warn('[ExcelReader] Error running Excel.run inspection, fallback:', err);
    }
  }

  // Standalone mode: realistic business dataset for Iranian enterprise
  return generateStandaloneMockReport();
}

/**
 * Core analysis algorithm for cells, types, errors and duplicates
 */
function analyzeRangeData(
  sheetName: string,
  address: string,
  rowCount: number,
  columnCount: number,
  rawValues: unknown[][],
  rawFormulas: unknown[][],
  rawFormats: unknown[][],
  sampleLimit: number
): ExcelInspectionReport {
  const isLargeDataset = rowCount > 5000;
  const headers = (rawValues[0] || []).map((v, i) => (v != null ? String(v).trim() : `ستون_${i + 1}`));

  const formulaErrors: FormulaErrorFinding[] = [];
  const duplicateRows: DuplicateRowFinding[] = [];
  let emptyCellCount = 0;

  // Track duplicates using composite hash
  const seenRowHashes = new Map<string, number>();

  // Columns metadata accumulator
  const columns: ColumnMetadata[] = headers.map((name, colIdx) => ({
    name,
    index: colIdx,
    inferredType: 'text',
    sampleValues: [],
    emptyCount: 0
  }));

  // Analyze rows (row 0 is header)
  const maxInspectRows = Math.min(rowCount, isLargeDataset ? 1000 : rowCount);

  for (let r = 1; r < maxInspectRows; r++) {
    const rowVals = rawValues[r] || [];
    const rowForms = rawFormulas[r] || [];
    const rowFmts = rawFormats[r] || [];

    // Check for duplicate row
    const nonNullRowVals = rowVals.map((v) => (v != null ? String(v).trim() : ''));
    const isRowEmpty = nonNullRowVals.every((v) => v === '');

    if (!isRowEmpty) {
      const compositeKey = nonNullRowVals.join('___');
      if (seenRowHashes.has(compositeKey)) {
        duplicateRows.push({
          originalRow: seenRowHashes.get(compositeKey)! + 1,
          duplicateRow: r + 1,
          keyValues: nonNullRowVals.slice(0, 3)
        });
      } else {
        seenRowHashes.set(compositeKey, r);
      }
    }

    for (let c = 0; c < columnCount; c++) {
      const val = rowVals[c];
      const formula = rowForms[c];
      const fmt = rowFmts[c];

      // Empty cell check
      if (val == null || val === '') {
        emptyCellCount++;
        columns[c].emptyCount++;
        continue;
      }

      // Check formula errors
      const valStr = String(val).trim();
      const formulaStr = String(formula || '').trim();

      for (const errCode of KNOWN_FORMULA_ERRORS) {
        if (valStr.includes(errCode) || formulaStr.includes(errCode)) {
          formulaErrors.push({
            cellAddress: `${sheetName}!${getColumnLetter(c)}${r + 1}`,
            errorType: errCode,
            formula: formulaStr,
            row: r + 1,
            column: c + 1
          });
          break;
        }
      }

      // Collect sample values
      if (columns[c].sampleValues.length < 5) {
        columns[c].sampleValues.push(val as string | number);
      }
    }
  }

  // Infer data type for each column
  columns.forEach((col) => {
    col.inferredType = inferColumnDataType(col.sampleValues);
  });

  // Construct sample rows for preview
  const sampleRows: Array<Record<string, unknown>> = [];
  const previewRowCount = Math.min(rowCount - 1, 15);
  for (let r = 1; r <= previewRowCount; r++) {
    const rowObj: Record<string, unknown> = {};
    headers.forEach((h, c) => {
      rowObj[h] = rawValues[r]?.[c] ?? '';
    });
    sampleRows.push(rowObj);
  }

  // Generate natural language summary
  const summaryParts: string[] = [
    `کاربرگ «${sheetName}» با محدوده ${address} ارزیابی شد.`,
    `تعداد سطرها: ${rowCount.toLocaleString('fa-IR')} | تعداد ستون‌ها: ${columnCount.toLocaleString('fa-IR')}.`
  ];

  if (formulaErrors.length > 0) {
    summaryParts.push(
      `⚠️ تعداد ${formulaErrors.length.toLocaleString('fa-IR')} خطای فرمول (${formulaErrors.map((e) => e.errorType).slice(0, 3).join(', ')}) شناسایی شد.`
    );
  } else {
    summaryParts.push('✅ هیچ خطای فرمولی در محدوده ارزیابی‌شده وجود ندارد.');
  }

  if (duplicateRows.length > 0) {
    summaryParts.push(
      `⚠️ تعداد ${duplicateRows.length.toLocaleString('fa-IR')} سطر تکراری شناسایی گردید.`
    );
  } else {
    summaryParts.push('✅ سطرهای تکراری در داده‌ها یافت نشد.');
  }

  if (emptyCellCount > 0) {
    summaryParts.push(`تعداد ${emptyCellCount.toLocaleString('fa-IR')} سلول خالی در محدوده موجود است.`);
  }

  return {
    timestamp: Date.now(),
    documentTitle: getSafeDocumentName(),
    activeSheetName: sheetName,
    usedRangeAddress: address,
    rowCount,
    columnCount,
    headers,
    columns,
    formulaErrors,
    duplicateRows,
    emptyCellCount,
    sampleRows,
    isLargeDataset,
    summaryText: summaryParts.join(' ')
  };
}

/**
 * Infer column data type based on sample values
 */
function inferColumnDataType(samples: (string | number)[]): ColumnDataType {
  if (samples.length === 0) return 'text';

  let numberCount = 0;
  let dateCount = 0;
  let currencyCount = 0;

  for (const s of samples) {
    if (typeof s === 'number') {
      numberCount++;
      continue;
    }
    const str = String(s).trim();
    if (/^[\d,.\s]+$/.test(str) && !isNaN(Number(str.replace(/,/g, '')))) {
      numberCount++;
    } else if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(str) || /^\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(str)) {
      dateCount++;
    } else if (str.includes('ریال') || str.includes('تومان') || str.includes('$') || str.includes('€')) {
      currencyCount++;
    }
  }

  const threshold = samples.length * 0.7;
  if (currencyCount >= threshold) return 'currency';
  if (dateCount >= threshold) return 'date';
  if (numberCount >= threshold) return 'number';
  return 'text';
}

function getColumnLetter(colIndex: number): string {
  let letter = '';
  let temp = colIndex;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function createEmptyReport(sheetName: string): ExcelInspectionReport {
  return {
    timestamp: Date.now(),
    documentTitle: getSafeDocumentName(),
    activeSheetName: sheetName,
    usedRangeAddress: `${sheetName}!A1`,
    rowCount: 0,
    columnCount: 0,
    headers: [],
    columns: [],
    formulaErrors: [],
    duplicateRows: [],
    emptyCellCount: 0,
    sampleRows: [],
    isLargeDataset: false,
    summaryText: `کاربرگ «${sheetName}» در حال حاضر خالی است و هیچ داده یا سلول پرشده‌ای در آن یافت نشد.`
  };
}

function getSafeDocumentName(): string {
  if (typeof window !== 'undefined') {
    const officeGlobal = (window as unknown as { Office?: { context?: { document?: { url?: string } } } }).Office;
    const url = officeGlobal?.context?.document?.url;
    if (url) {
      const parts = url.split(/[\\/]/);
      return parts[parts.length - 1];
    }
  }
  return 'دفترکل_مالی_شرکت.xlsx';
}

/**
 * Realistic business dataset for Iranian enterprise sales and finance
 */
function generateStandaloneMockReport(): ExcelInspectionReport {
  const headers = [
    'شناسه',
    'تاریخ',
    'شرح کالا / خدمات',
    'دسته‌بندی',
    'تعداد',
    'قیمت واحد (ریال)',
    'مبلغ کل (ریال)',
    'مرکز هزینه',
    'وضعیت تسویه'
  ];

  const columns: ColumnMetadata[] = [
    { name: 'شناسه', index: 0, inferredType: 'text', sampleValues: ['TRX-101', 'TRX-102', 'TRX-103'], emptyCount: 0 },
    { name: 'تاریخ', index: 1, inferredType: 'date', sampleValues: ['1403/07/15', '1403/07/16', '1403/07/18'], emptyCount: 0 },
    { name: 'شرح کالا / خدمات', index: 2, inferredType: 'text', sampleValues: ['سرور HP Gen10', 'لایسنس مایکروسافت', 'روتر سیسکو'], emptyCount: 0 },
    { name: 'دسته‌بندی', index: 3, inferredType: 'text', sampleValues: ['تجهیزات IT', 'نرم‌افزار', 'شبکه'], emptyCount: 0 },
    { name: 'تعداد', index: 4, inferredType: 'number', sampleValues: [2, 10, 4], emptyCount: 0 },
    { name: 'قیمت واحد (ریال)', index: 5, inferredType: 'currency', sampleValues: [450000000, 18000000, 120000000], emptyCount: 0 },
    { name: 'مبلغ کل (ریال)', index: 6, inferredType: 'formula', sampleValues: [900000000, 180000000, 480000000], emptyCount: 0 },
    { name: 'مرکز هزینه', index: 7, inferredType: 'text', sampleValues: ['فناوری اطلاعات', 'فنی', 'اداری'], emptyCount: 1 },
    { name: 'وضعیت تسویه', index: 8, inferredType: 'text', sampleValues: ['پرداخت شده', 'معوق', 'در انتظار'], emptyCount: 0 }
  ];

  const sampleRows = [
    {
      'شناسه': 'TRX-101',
      'تاریخ': '1403/07/15',
      'شرح کالا / خدمات': 'سرور اچ‌پی Gen10 Plus',
      'دسته‌بندی': 'تجهیزات سخت‌افزار',
      'تعداد': 2,
      'قیمت واحد (ریال)': '450,000,000',
      'مبلغ کل (ریال)': '900,000,000',
      'مرکز هزینه': 'مرکز داده تهران',
      'وضعیت تسویه': 'تسویه کامل'
    },
    {
      'شناسه': 'TRX-102',
      'تاریخ': '1403/07/16',
      'شرح کالا / خدمات': 'لایسنس آفیس ۲۰۲۴ اینترپرایز',
      'دسته‌بندی': 'نرم‌افزار و مجوزها',
      'تعداد': 15,
      'قیمت واحد (ریال)': '18,500,000',
      'مبلغ کل (ریال)': '277,500,000',
      'مرکز هزینه': 'فناوری اطلاعات',
      'وضعیت تسویه': 'تسویه کامل'
    },
    {
      'شناسه': 'TRX-103',
      'تاریخ': '1403/07/18',
      'شرح کالا / خدمات': 'خدمات پشتیبانی شبکه ماه مهر',
      'دسته‌بندی': 'خدمات فنی',
      'تعداد': 1,
      'قیمت واحد (ریال)': '85,000,000',
      'مبلغ کل (ریال)': '#DIV/0!',
      'مرکز هزینه': 'پشتیبانی فنی',
      'وضعیت تسویه': 'در انتظار بررسی'
    },
    {
      'شناسه': 'TRX-104',
      'تاریخ': '1403/07/20',
      'شرح کالا / خدمات': 'کابل شبکه Cat6 Legrand',
      'دسته‌بندی': 'پسیو شبکه',
      'تعداد': 6,
      'قیمت واحد (ریال)': '14,000,000',
      'مبلغ کل (ریال)': '84,000,000',
      'مرکز هزینه': '',
      'وضعیت تسویه': 'پرداخت شده'
    },
    {
      'شناسه': 'TRX-101',
      'تاریخ': '1403/07/15',
      'شرح کالا / خدمات': 'سرور اچ‌پی Gen10 Plus',
      'دسته‌بندی': 'تجهیزات سخت‌افزار',
      'تعداد': 2,
      'قیمت واحد (ریال)': '450,000,000',
      'مبلغ کل (ریال)': '900,000,000',
      'مرکز هزینه': 'مرکز داده تهران',
      'وضعیت تسویه': 'تسویه کامل'
    }
  ];

  return {
    timestamp: Date.now(),
    documentTitle: 'دفترکل_مالی_شرکت.xlsx',
    activeSheetName: 'فروش و هزینه‌ها',
    usedRangeAddress: 'فروش و هزینه‌ها!A1:I46',
    rowCount: 46,
    columnCount: 9,
    headers,
    columns,
    formulaErrors: [
      {
        cellAddress: 'فروش و هزینه‌ها!G4',
        errorType: '#DIV/0!',
        formula: '=E4/0',
        row: 4,
        column: 7
      }
    ],
    duplicateRows: [
      {
        originalRow: 2,
        duplicateRow: 6,
        keyValues: ['TRX-101', '1403/07/15', 'سرور اچ‌پی Gen10 Plus']
      }
    ],
    emptyCellCount: 1,
    sampleRows,
    isLargeDataset: false,
    summaryText:
      'کاربرگ «فروش و هزینه‌ها» با محدوده A1:I46 شامل ۴۶ سطر و ۹ ستون ارزیابی شد. ۱ خطای فرمول (#DIV/0!) در سلول G4 و ۱ سطر تکراری (ردیف ۶ نسبت به ردیف ۲) و ۱ سلول خالی در ستون مرکز هزینه شناسایی گردید.'
  };
}
