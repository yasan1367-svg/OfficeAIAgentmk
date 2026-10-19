/// <reference types="office-js" />
/**
 * Office AI Agent - Excel Writer & Cell Formatting Module (Phase 5)
 */

export interface CellFormatOptions {
  numberFormat?: string;
  fillColor?: string;
  fontBold?: boolean;
  fontSize?: number;
  fontColor?: string;
  fontName?: string;
  horizontalAlignment?: 'Left' | 'Center' | 'Right' | 'General';
  border?: boolean;
}

export interface RangeWriteOptions {
  address: string;
  values?: (string | number | boolean)[][];
  formulas?: string[][];
  formatting?: CellFormatOptions;
}

/**
 * Writes values and formulas to an Excel range with optional formatting in a single batch
 */
export async function writeRangeData(
  sheetName: string,
  options: RangeWriteOptions
): Promise<{ success: boolean; address: string; rowCount: number }> {
  if (typeof window !== 'undefined' && (window as unknown as { Excel?: typeof Excel }).Excel) {
    try {
      const ExcelGlobal = (window as unknown as { Excel: typeof Excel }).Excel;
      return await ExcelGlobal.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const range = sheet.getRange(options.address);

        if (options.values) {
          range.values = options.values;
        }

        if (options.formulas) {
          range.formulas = options.formulas;
        }

        if (options.formatting) {
          applyFormattingToRange(range, options.formatting);
        }

        range.format.autofitColumns();
        await context.sync();

        return {
          success: true,
          address: `${sheetName}!${options.address}`,
          rowCount: options.values ? options.values.length : 1
        };
      });
    } catch (err) {
      console.error('[ExcelWriter] writeRangeData error:', err);
      throw new Error(`خطا در نوشتن داده‌ها در اکسل: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  // Standalone simulated write
  console.log(`[ExcelWriter Standalone] Written to ${sheetName}!${options.address}`);
  return {
    success: true,
    address: `${sheetName}!${options.address}`,
    rowCount: options.values ? options.values.length : 1
  };
}

/**
 * Creates a new styled Excel Table from a range
 */
export async function createTableFromRange(
  sheetName: string,
  rangeAddress: string,
  tableName: string,
  hasHeaders: boolean = true
): Promise<{ success: boolean; tableName: string }> {
  if (typeof window !== 'undefined' && (window as unknown as { Excel?: typeof Excel }).Excel) {
    try {
      const ExcelGlobal = (window as unknown as { Excel: typeof Excel }).Excel;
      return await ExcelGlobal.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const range = sheet.getRange(rangeAddress);
        const table = sheet.tables.add(range, hasHeaders);
        table.name = tableName;
        table.style = 'TableStyleMedium9'; // Professional green/emerald theme
        await context.sync();
        return { success: true, tableName };
      });
    } catch (err) {
      console.error('[ExcelWriter] createTableFromRange error:', err);
      throw err;
    }
  }

  return { success: true, tableName };
}

/**
 * Creates a brand new worksheet safely
 */
export async function createNewWorksheet(
  sheetName: string
): Promise<{ success: boolean; sheetName: string; isNew: boolean }> {
  if (typeof window !== 'undefined' && (window as unknown as { Excel?: typeof Excel }).Excel) {
    try {
      const ExcelGlobal = (window as unknown as { Excel: typeof Excel }).Excel;
      return await ExcelGlobal.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load('items/name');
        await context.sync();

        // Check if sheet name already exists
        let finalName = sheetName;
        let counter = 1;
        const existingNames = new Set(sheets.items.map((s) => s.name.toLowerCase()));

        while (existingNames.has(finalName.toLowerCase())) {
          finalName = `${sheetName}_${counter}`;
          counter++;
        }

        const newSheet = sheets.add(finalName);
        newSheet.activate();
        await context.sync();

        return {
          success: true,
          sheetName: finalName,
          isNew: true
        };
      });
    } catch (err) {
      console.error('[ExcelWriter] createNewWorksheet error:', err);
      throw err;
    }
  }

  return { success: true, sheetName, isNew: true };
}

function applyFormattingToRange(range: Excel.Range, fmt: CellFormatOptions) {
  if (fmt.numberFormat) {
    range.numberFormat = [[fmt.numberFormat]];
  }

  if (fmt.fillColor) {
    range.format.fill.color = fmt.fillColor;
  }

  if (fmt.fontBold !== undefined) {
    range.format.font.bold = fmt.fontBold;
  }

  if (fmt.fontSize) {
    range.format.font.size = fmt.fontSize;
  }

  if (fmt.fontColor) {
    range.format.font.color = fmt.fontColor;
  }

  if (fmt.fontName) {
    range.format.font.name = fmt.fontName;
  }

  if (fmt.horizontalAlignment) {
    range.format.horizontalAlignment = fmt.horizontalAlignment;
  }
}
