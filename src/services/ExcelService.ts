export interface ExcelDataRow {
  [key: string]: any;
}

export interface ExcelContext {
  data: ExcelDataRow[];
  usedRangeAddress: string;
  activeCellAddress: string;
}

export const getSurroundingData = async (): Promise<ExcelContext> => {
  return await Excel.run(async (context: Excel.RequestContext) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const activeCell = context.workbook.getActiveCell();
    
    usedRange.load("values, address, rowCount, columnCount");
    activeCell.load("address");
    
    await context.sync();

    const usedRangeAddress = usedRange.address;
    const activeCellAddress = activeCell.address;
    const values = usedRange.values;
    
    if (values.length === 0) {
      return { data: [], usedRangeAddress, activeCellAddress };
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

    return { data, usedRangeAddress, activeCellAddress };
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
