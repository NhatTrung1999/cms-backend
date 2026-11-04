import * as ExcelJS from 'exceljs';
import * as fs from 'fs';

/**
 * Reads the first sheet of an uploaded Excel file and returns its data rows as objects keyed by the header
 * (spaces in a header become '_'). Throws when the file is missing, has no sheet or lacks a required column.
 * Only rows for which `isRowAccepted` is truthy are returned; `cellsAsStrings` turns every cell value into a string.
 */
export const readImportRows = async <T = any>(
  file: Express.Multer.File,
  requiredHeaders: string[],
  isRowAccepted: (rowData: T) => unknown,
  options: { cellsAsStrings?: boolean } = {},
): Promise<T[]> => {
  if (!file.path || !fs.existsSync(file.path)) {
    throw new Error('File path not found!');
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file.path);
  const worksheet = workbook.getWorksheet(1);
  if (!worksheet) {
    throw new Error('No worksheet found in the Excel file');
  }

  const headerRow = worksheet.getRow(1);
  const headers: string[] = [];
  headerRow.eachCell((cell) => {
    if (cell.value) {
      headers.push(cell.value.toString().trim().replace(/\s+/g, '_'));
    }
  });

  const missingHeaders = requiredHeaders.filter((h) => !headers.includes(h));

  if (missingHeaders.length > 0) {
    throw new Error(
      `Excel file format is invalid! Missing columns: ${missingHeaders.join(', ')}`,
    );
  }

  const data: T[] = [];

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (rowNumber === 1 && row.cellCount > 0) return;

    const rowData: Record<string, unknown> = {};
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const headerCell = headerRow.getCell(colNumber);
      const headerValue = headerCell.value;
      if (headerValue !== null && headerValue !== undefined) {
        const key = headerValue.toString().trim().replace(/\s+/g, '_');
        rowData[key] = options.cellsAsStrings ? String(cell.value) : cell.value;
      }
    });
    if (isRowAccepted(rowData as T)) {
      data.push(rowData as T);
    }
  });

  return data;
};

/**
 * Runs `upsert` on every row, one after the other, and counts the rows it inserted and the rows it updated
 * (a row it skipped counts for neither).
 */
export const countUpserts = async <T>(
  rows: T[],
  upsert: (row: T) => Promise<'insert' | 'update' | 'skip'>,
) => {
  let insertCount = 0;
  let updateCount = 0;
  for (const row of rows) {
    const action = await upsert(row);
    if (action === 'update') {
      updateCount++;
    } else if (action === 'insert') {
      insertCount++;
    }
  }
  return { insertCount, updateCount };
};

/** Result message of the import methods. */
export const buildImportMessage = (
  insertCount: number,
  updateCount: number,
  totalRows: number,
) =>
  `Processed successfully! Inserted: ${insertCount} records, Updated: ${updateCount} records. Total rows processed: ${totalRows}.`;
