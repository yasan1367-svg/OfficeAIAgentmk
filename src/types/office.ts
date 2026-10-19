/**
 * Office AI Agent - Office Integration Types (Phase 4 Detailed Excel Inspection)
 */

export type OfficeHost = 'excel' | 'word' | 'standalone';

export interface OfficeContextInfo {
  isOfficeReady: boolean;
  host: OfficeHost;
  platform?: string;
  documentTitle?: string;
  activeSheetName?: string;
  error?: string;
}

export interface ExcelSheetSummary {
  name: string;
  id: string;
  visibility: string;
  usedRangeAddress?: string;
  rowCount?: number;
  columnCount?: number;
}

export interface ExcelTableSummary {
  name: string;
  id: string;
  sheetName: string;
  headerNames: string[];
  rowCount: number;
}

export interface ExcelWorkbookStructure {
  documentName: string;
  activeSheetName: string;
  sheets: ExcelSheetSummary[];
  tables: ExcelTableSummary[];
}

export type ColumnDataType = 'number' | 'text' | 'date' | 'currency' | 'formula' | 'mixed';

export interface ColumnMetadata {
  name: string;
  index: number;
  inferredType: ColumnDataType;
  sampleValues: (string | number)[];
  emptyCount: number;
}

export interface FormulaErrorFinding {
  cellAddress: string;
  errorType: string;
  formula: string;
  row: number;
  column: number;
}

export interface DuplicateRowFinding {
  originalRow: number;
  duplicateRow: number;
  keyValues: string[];
}

export interface ExcelInspectionReport {
  timestamp: number;
  documentTitle: string;
  activeSheetName: string;
  usedRangeAddress: string;
  rowCount: number;
  columnCount: number;
  headers: string[];
  columns: ColumnMetadata[];
  formulaErrors: FormulaErrorFinding[];
  duplicateRows: DuplicateRowFinding[];
  emptyCellCount: number;
  sampleRows: Array<Record<string, unknown>>;
  isLargeDataset: boolean;
  summaryText: string;
}

export interface WordDocumentStructure {
  documentName: string;
  paragraphCount: number;
  headings: Array<{ text: string; level: number }>;
  tableCount: number;
}
