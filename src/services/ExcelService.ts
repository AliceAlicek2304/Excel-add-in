export interface ExcelDataRow {
  [key: string]: any;
}

export interface ExcelContext {
  data: ExcelDataRow[];
  usedRangeAddress: string;
  activeCellAddress: string;
  allSheetNames: string[];
}

export const getSurroundingData = async (): Promise<ExcelContext> => {
  return await Excel.run(async (context: Excel.RequestContext) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const activeCell = context.workbook.getActiveCell();
    const allSheets = context.workbook.worksheets;
    
    usedRange.load("values, address, rowCount, columnCount");
    activeCell.load("address");
    allSheets.load("items/name");
    
    await context.sync();

    const usedRangeAddress = usedRange.address;
    const activeCellAddress = activeCell.address;
    const allSheetNames = allSheets.items.map(s => s.name);
    const values = usedRange.values;
    
    if (values.length === 0) {
      return { data: [], usedRangeAddress, activeCellAddress, allSheetNames };
    }

    // Limit data for AI context (first 50 rows)
    const limitedValues = values.length > 50 ? values.slice(0, 50) : values;

    // Check headers
    const firstRow = limitedValues[0];
    const hasHeaders = firstRow.every((cell: any) => typeof cell === 'string' && cell.trim() !== '');

    let data: ExcelDataRow[] = [];
    if (hasHeaders && limitedValues.length > 1) {
      const headers = firstRow;
      data = limitedValues.slice(1).map((row: any[]) => {
        const obj: ExcelDataRow = {};
        headers.forEach((header: any, index: number) => {
          obj[header.toString() || `Column${index}`] = row[index];
        });
        return obj;
      });
    } else {
      data = limitedValues.map((row: any[]) => {
        const obj: ExcelDataRow = {};
        row.forEach((cell: any, index: number) => {
          const colLetter = String.fromCharCode(65 + index); // Simplified A, B, C...
          obj[colLetter] = cell;
        });
        return obj;
      });
    }

    return { data, usedRangeAddress, activeCellAddress, allSheetNames };
  });
};

export const writeToActiveCell = async (value: string) => {
  await Excel.run(async (context: Excel.RequestContext) => {
    const range = context.workbook.getSelectedRange();
    if (value.startsWith("=")) {
      range.formulas = [[value]];
    } else {
      range.values = [[value]];
    }
    await context.sync();
  });
};

export const writeToCell = async (cellAddress: string, value: string) => {
  await Excel.run(async (context: Excel.RequestContext) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(cellAddress);
    
    if (value.startsWith("=")) {
      range.formulas = [[value]];
    } else {
      range.values = [[value]];
    }
    await context.sync();
  });
};

export const writeArrayToRange = async (values: string[]) => {
  await Excel.run(async (context: Excel.RequestContext) => {
    const range = context.workbook.getSelectedRange();
    
    // Convert array to 2D array format (vertical)
    const cellValues = values.map(v => [v]);
    
    // Expand range to fit all values
    const expandedRange = range.getResizedRange(values.length - 1, 0);
    
    // Check if values are formulas or plain values
    const hasFormula = values.some(v => v.startsWith("="));
    if (hasFormula) {
      expandedRange.formulas = cellValues;
    } else {
      expandedRange.values = cellValues;
    }
    
    await context.sync();
  });
};
export const createChart = async (type: string, rangeAddress: string, title: string = "AI Generated Chart") => {
  await Excel.run(async (context: Excel.RequestContext) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeAddress);
    
    let chartType: Excel.ChartType;
    switch (type.toLowerCase()) {
      case 'pie':
        chartType = Excel.ChartType.pie;
        break;
      case 'line':
        chartType = Excel.ChartType.line;
        break;
      case 'column':
      default:
        chartType = Excel.ChartType.columnClustered;
        break;
    }

    const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);
    chart.title.text = title;
    chart.legend.position = Excel.ChartLegendPosition.right;
    chart.legend.format.fill.setSolidColor("white");
    
    await context.sync();
  });
};
export const consolidateAllSheets = async (cellAddress: string, chartTypeStr: string) => {
  await Excel.run(async (context: Excel.RequestContext) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    // 1. Create or get "Result" sheet
    let summarySheet = context.workbook.worksheets.getItemOrNullObject("Result");
    await context.sync();

    if (summarySheet.isNullObject) {
      summarySheet = context.workbook.worksheets.add("Result");
    } else {
      summarySheet.activate();
      summarySheet.getUsedRange().clear();
    }

    const data: any[][] = [["Tên Sheet", `Giá trị (${cellAddress})`]];
    
    // 2. Collect data
    for (const sheet of sheets.items) {
      if (sheet.name === "Result") continue;
      
      const range = sheet.getRange(cellAddress);
      range.load("values");
      await context.sync();
      
      data.push([sheet.name, range.values[0][0]]);
    }

    // 3. Write data to summary sheet
    const targetRange = summarySheet.getRangeByIndexes(0, 0, data.length, 2);
    targetRange.values = data;
    summarySheet.activate();

    // 4. Create chart
    let chartType: Excel.ChartType;
    switch (chartTypeStr.toLowerCase()) {
      case 'pie': chartType = Excel.ChartType.pie; break;
      case 'line': chartType = Excel.ChartType.line; break;
      default: chartType = Excel.ChartType.columnClustered; break;
    }

    const chart = summarySheet.charts.add(chartType, targetRange, Excel.ChartSeriesBy.auto);
    chart.title.text = `Tổng hợp ${cellAddress} từ tất cả các Sheet`;
    
    await context.sync();
  });
};
