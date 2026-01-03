import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { buildQueryHRModule } from 'src/helper/hrmodule';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';

@Injectable()
export class HrService {
  constructor(
    @Inject('LYV_HRIS') private readonly LYV_HRIS: Sequelize,
    @Inject('LHG_HRIS') private readonly LHG_HRIS: Sequelize,
    @Inject('LVL_HRIS') private readonly LVL_HRIS: Sequelize,
    @Inject('LYM_HRIS') private readonly LYM_HRIS: Sequelize,
    @Inject('JAZ_HRIS') private readonly JAZ_HRIS: Sequelize,
    @Inject('JZS_HRIS') private readonly JZS_HRIS: Sequelize,
  ) {}

  async findAll(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'ID',
    sortOrder: string = 'asc',
  ) {
    let db: Sequelize;
    switch (factory.trim().toUpperCase()) {
      case 'LYV':
        db = this.LYV_HRIS;
        break;
      case 'LHG':
        db = this.LHG_HRIS;
        break;
      case 'LVL':
        db = this.LVL_HRIS;
        break;
      case 'LYM':
        db = this.LYM_HRIS;
        break;
      case 'JAZ':
        db = this.JAZ_HRIS;
        break;
      case 'JZS':
        db = this.JZS_HRIS;
        break;
      default:
        throw new Error('Invalid factory code');
    }
    const offset = (page - 1) * limit;
    const { query, countQuery } = buildQueryHRModule(dateFrom, dateTo, factory);
    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

    const [dataResults, countResults] = await Promise.all([
      db.query(query, { replacements, type: QueryTypes.SELECT }),
      db.query(countQuery, {
        type: QueryTypes.SELECT,
        replacements,
      }),
    ]);
    //ID, FullName, Department, PermanentAddress, CurrentAddress, TransportationMode, NumberOfWorkingDays
    // const results = await db.query(`SELECT * FROM users`, {
    //   type: QueryTypes.SELECT,
    // });
    let data = dataResults;

    data.sort((a, b) => {
      const aValue = a[sortField];
      const bValue = b[sortField];
      if (sortOrder === 'asc') {
        return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
      } else {
        return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
      }
    });

    data = data.slice(offset, offset + limit);

    const total = (countResults[0] as { total: number })?.total || 0;

    const hasMore = offset + data.length < total;
    return { data, page, limit, total, hasMore };
  }

  async importFromExcel(
    file: Express.Multer.File,
    userid: string,
    factory: string,
  ) {
    try {
      let updateCount = 0;
      if (!file.path || !fs.existsSync(file.path)) {
        throw new Error('File path not found!');
      }

      let db: Sequelize;
      switch (factory.trim().toUpperCase()) {
        case 'LYV':
          db = this.LYV_HRIS;
          break;
        case 'LHG':
          db = this.LHG_HRIS;
          break;
        case 'LVL':
          db = this.LVL_HRIS;
          break;
        case 'LYM':
          db = this.LYM_HRIS;
          break;
        case 'JAZ':
          db = this.JAZ_HRIS;
          break;
        case 'JZS':
          db = this.JZS_HRIS;
          break;
        default:
          throw new Error('Invalid factory code');
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
      const requiredHeaders = [
        'ID',
        'Full_Name',
        'Department',
        'Permanent_Address',
        'Current_Address',
        'Transportation_Method',
      ];

      const missingHeaders = requiredHeaders.filter(
        (h) => !headers.includes(h),
      );

      if (missingHeaders.length > 0) {
        throw new Error(
          `Excel file format is invalid! Missing columns: ${missingHeaders.join(', ')}`,
        );
      }

      const data: {
        ID: string;
        Full_Name: string;
        Department: string;
        Permanent_Address: string;
        Current_Address: string;
        Transportation_Method: string;
      }[] = [];

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber === 1 && row.cellCount > 0) return;

        const rowData: any = {};
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const headerCell = headerRow.getCell(colNumber);
          const headerValue = headerCell.value;
          if (headerValue !== null && headerValue !== undefined) {
            const key = headerValue.toString().trim().replace(/\s+/g, '_');
            rowData[key] = cell.value;
          }
        });
        if (
          rowData?.ID &&
          rowData?.Full_Name &&
          rowData?.Department &&
          rowData?.Permanent_Address &&
          rowData?.Current_Address &&
          rowData?.Transportation_Method
        ) {
          data.push(rowData);
        }
      });

      for (let item of data) {
        const records: { total: number }[] = await db.query(
          `SELECT COUNT(*) total
          FROM users
          WHERE userId = ?`,
          {
            replacements: [String(item.ID)],
            type: QueryTypes.SELECT,
          },
        );

        if (records[0].total > 0) {
          await db.query(
            `UPDATE users
            SET
              Address_Live = ?,
              Vehicle = ?,
              UpdatedBy = ?,
              UpdatedAt = GETDATE()
            WHERE userId = ?`,
            {
              replacements: [
                item.Current_Address,
                item.Transportation_Method,
                userid,
                String(item.ID),
              ],
              type: QueryTypes.SELECT,
            },
          );
          updateCount++;
        }
      }
      const message = `Updated: ${updateCount} records. Total rows processed: ${data.length}.`;
      return { message };
    } catch (error) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }
}
