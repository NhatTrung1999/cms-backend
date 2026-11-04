import { ConfigService } from '@nestjs/config';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';

/** Folder of the exported Excel files (EXCEL_STORAGE_PATH), created when it does not exist. */
export const getExcelRootFolder = (configService: ConfigService): string => {
  const rootFolder: string = configService.get(
    'EXCEL_STORAGE_PATH',
    'D:/FileExcel',
  );
  if (!fs.existsSync(rootFolder)) {
    fs.mkdirSync(rootFolder, { recursive: true });
  }
  return rootFolder;
};

/** Path of an exported file: <rootFolder>/<module>/<fileName> (the module folder is created when needed). */
export const getExcelFilePath = (
  rootFolder: string,
  module: string,
  fileName: string,
) => {
  const folder = path.join(rootFolder, module);
  if (!fs.existsSync(folder)) {
    fs.mkdirSync(folder, { recursive: true });
  }
  return path.join(folder, fileName);
};

/** The file records of the File Management page live in CMW_File_Management: Status 0 = waiting, 1 = done, 2 = error. */
export const insertFileRecord = async (
  eip: Sequelize,
  record: {
    id: string;
    module: string;
    fileName: string;
    filePath: string;
    factory: string;
    userID: string;
  },
) => {
  await eip.query(
    `INSERT INTO CMW_File_Management
      (
        ID,
        Module,
        [File_Name],
        [Path],
        [Status],
        CreatedAt,
        CreatedFactory,
        CreatedDate
      )
      VALUES
      (
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?
      )`,
    {
      replacements: [
        record.id,
        record.module,
        record.fileName,
        record.filePath,
        '0',
        record.userID,
        record.factory,
        new Date().toISOString(),
      ],
      type: QueryTypes.INSERT,
    },
  );
};

export const findFileRecord = (eip: Sequelize, id: string) =>
  eip.query(
    `SELECT * FROM CMW_File_Management WHERE ID = ? ORDER BY CreatedDate`,
    { replacements: [id], type: QueryTypes.SELECT },
  );

/** Marks the file as done and returns its record. */
export const markFileDone = async (eip: Sequelize, id: string) => {
  try {
    await eip.query(
      `UPDATE CMW_File_Management
        SET
          [Status] = '1'
        WHERE ID = ?`,
      { replacements: [id], type: QueryTypes.UPDATE },
    );
    return await findFileRecord(eip, id);
  } catch (error: any) {
    throw new Error(error);
  }
};

export const markFileFailed = async (eip: Sequelize, id: string) => {
  await eip.query(
    `UPDATE CMW_File_Management SET [Status] = '2' WHERE ID = ?`,
    { replacements: [id], type: QueryTypes.UPDATE },
  );
};

/**
 * Builds one Excel file in the background (`fill` writes the sheet), saves it to `filePath`, marks its record
 * as done and tells the page. When anything fails the record is marked as failed (when the database allows it)
 * and the page is told too. It never throws: nothing awaits it, so an error would stop the whole backend.
 */
export const runFileExport = async ({
  eip,
  gateway,
  id,
  filePath,
  sheetName,
  fill,
}: {
  eip: Sequelize;
  gateway: { broadcastEvent: (event: string, data: string) => void };
  id: string;
  filePath: string;
  sheetName: string;
  fill: (sheet: ExcelJS.Worksheet) => Promise<void>;
}) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(sheetName);

    await fill(sheet);

    await workbook.xlsx.writeFile(filePath);

    const statusDone = await markFileDone(eip, id);
    if (statusDone.length) {
      gateway.broadcastEvent('file-excel-done', 'Export file excel success!');
    }
  } catch (error) {
    console.error('Export file excel failed:', error);
    try {
      await markFileFailed(eip, id);
    } catch (updateError) {
      // e.g. the database cannot be reached: the file stays "waiting", but the backend keeps running
      console.error('Could not mark the file as failed:', updateError);
    }
    gateway.broadcastEvent('file-excel-error', 'Error export file excel!');
  }
};
