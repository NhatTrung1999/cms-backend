import type { Column, Worksheet } from 'exceljs';

/**
 * Makes every column as wide as its longest cell (times 1.2)
 * and draws a thin border around every cell of the sheet.
 */
export const autoFitColumnsAndAddBorders = (sheet: Worksheet) => {
  sheet.columns.forEach((column) => {
    let maxLength = 0;
    if (typeof column.eachCell === 'function') {
      column.eachCell({ includeEmpty: true }, (cell) => {
        const cellValue = cell.value ? String(cell.value) : '';
        maxLength = Math.max(maxLength, cellValue.length);
      });
    }
    column.width = maxLength * 1.2;
  });
  sheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });
  });
};

/** Writes `rows` under `columns` and makes the sheet readable (auto-fit widths, borders). */
export const fillSheet = (
  sheet: Worksheet,
  columns: Partial<Column>[],
  rows: object[],
) => {
  sheet.columns = columns;
  rows.forEach((item) => sheet.addRow(item));
  autoFitColumnsAndAddBorders(sheet);
};
