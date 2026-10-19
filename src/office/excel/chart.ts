/// <reference types="office-js" />
/**
 * Office AI Agent - Excel Chart Generation Module (Phase 5)
 */

export interface ChartCreateOptions {
  sheetName: string;
  sourceDataRange: string;
  title: string;
  chartType: 'Column' | 'Line' | 'Pie' | 'Bar';
  position: {
    top: number;
    left: number;
    width: number;
    height: number;
  };
}

/**
 * Creates and formats an Excel chart from source data
 */
export async function createWorksheetChart(
  options: ChartCreateOptions
): Promise<{ success: boolean; chartTitle: string }> {
  if (typeof window !== 'undefined' && (window as unknown as { Excel?: typeof Excel }).Excel) {
    try {
      const ExcelGlobal = (window as unknown as { Excel: typeof Excel }).Excel;
      return await ExcelGlobal.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(options.sheetName);
        const sourceRange = sheet.getRange(options.sourceDataRange);

        let apiChartType = ExcelGlobal.ChartType.columnClustered;
        if (options.chartType === 'Line') {
          apiChartType = ExcelGlobal.ChartType.line;
        } else if (options.chartType === 'Pie') {
          apiChartType = ExcelGlobal.ChartType.pie;
        } else if (options.chartType === 'Bar') {
          apiChartType = ExcelGlobal.ChartType.barClustered;
        }

        const chart = sheet.charts.add(apiChartType, sourceRange, ExcelGlobal.ChartSeriesBy.auto);

        chart.title.text = options.title;
        chart.title.format.font.size = 11;
        chart.title.format.font.bold = true;

        chart.top = options.position.top;
        chart.left = options.position.left;
        chart.width = options.position.width;
        chart.height = options.position.height;

        await context.sync();

        return {
          success: true,
          chartTitle: options.title
        };
      });
    } catch (err) {
      console.error('[ExcelChart] Error creating chart:', err);
      throw new Error(`خطا در ایجاد نمودار: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  console.log(`[ExcelChart Standalone] Created ${options.chartType} chart: "${options.title}"`);
  return { success: true, chartTitle: options.title };
}
