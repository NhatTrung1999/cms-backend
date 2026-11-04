import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes, Transaction } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import {
  CreateLogCat1And4,
  CreateLogCat5,
  CreateLogCat6Accommodation,
  CreateLogCat6BusinessTravel,
  CreateLogCat7,
  CreateLogCat9And12,
} from './dto/create-logcat.dto';
import * as ExcelJS from 'exceljs';
import { PassThrough } from 'stream';
import { buildLogFilter } from './logcat.helper';
import {
  CAT1_AND_4_LOG_STREAM_COLUMNS,
  LOG_EXPORT_COLUMNS,
  LOG_TABLE,
  LogTable,
} from './logcat.config';
import { LOG_INSERT, insertLogRows } from './logcat.inserts';

@Injectable()
export class LogcatService {
  constructor(@Inject('EIP') private readonly EIP: Sequelize) {}

  /** Runs `work` in a transaction on EIP: commit on success, rollback + error on failure. */
  private async withTransaction(
    work: (transaction: Transaction) => Promise<void>,
    options: { logError?: boolean } = {},
  ) {
    const transaction = await this.EIP.transaction();

    try {
      await work(transaction);
      await transaction.commit();
    } catch (error) {
      if (options.logError) console.log(error);
      await transaction.rollback();
      throw new InternalServerErrorException('Error!');
    }
  }

  /** One page of a log table, filtered by date range and factory, sorted in memory. */
  private async getPagedLog(
    table: LogTable,
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    const offset = (page - 1) * limit;

    const { where, replacements } = buildLogFilter(dateFrom, dateTo, factory);
    const orderBy = table.newestFirst ? 'ORDER BY CreatedAt DESC' : '';

    const dataResults = await this.EIP.query(
      `
      SELECT *,
             COUNT(ID) OVER() AS total
      FROM ${table.name}
      ${where}
      ${orderBy}
      `,
      {
        type: QueryTypes.SELECT,
        replacements,
      },
    );

    let data = dataResults as any[];

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

    const cleanData = data.map(({ total, ...rest }) => rest);

    const total = dataResults.length
      ? Number((dataResults[0] as any).total)
      : 0;

    const hasMore = offset + data.length < total;

    return { data: cleanData, page, limit, total, hasMore };
  }

  /** Whole log table (filtered by date range and factory) as an .xlsx buffer. */
  private async exportLogExcel(
    table: LogTable,
    columns: Partial<ExcelJS.Column>[],
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    try {
      const { where, replacements } = buildLogFilter(dateFrom, dateTo, factory);

      const dataResults = await this.EIP.query(
        `
          SELECT *
          FROM ${table.name}
          ${where}
        `,
        {
          type: QueryTypes.SELECT,
          replacements,
        },
      );
      const data = dataResults as any[];

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('LogCat');

      worksheet.columns = columns;

      worksheet.getRow(1).font = { bold: true };

      worksheet.addRows(data);

      worksheet.columns.forEach((column) => {
        let maxColumnLength = 0;

        if (typeof column.eachCell === 'function') {
          column.eachCell({ includeEmpty: true }, (cell) => {
            const cellValue = cell.value ? cell.value.toString() : '';

            if (cellValue.length > maxColumnLength) {
              maxColumnLength = cellValue.length;
            }
          });
        }

        const newWidth = maxColumnLength + 2;
        column.width = newWidth < 10 ? 10 : newWidth > 50 ? 50 : newWidth;
      });

      worksheet.eachRow((row) => {
        row.eachCell((cell) => {
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      return buffer;
    } catch (error) {
      throw new InternalServerErrorException(error);
    }
  }

  // Logging CAT1&4
  async getLogCat1And4(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    return this.getPagedLog(
      LOG_TABLE.cat1And4,
      dateFrom,
      dateTo,
      factory,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  async createLogCat1And4(
    data: CreateLogCat1And4[],
    factory: string,
    userid: string,
  ) {
    await this.withTransaction((transaction) =>
      insertLogRows(this.EIP, transaction, LOG_INSERT.cat1And4, data, {
        factory,
        userid,
      }),
    );

    return {
      success: true,
      message: 'Create log CAT1&4 success!',
    };
  }

  async getLogCat1And4All(dateFrom: string, dateTo: string, factory: string) {
    const { where, replacements } = buildLogFilter(dateFrom, dateTo, factory);

    const dataResults = await this.EIP.query(
      `
      SELECT *
      FROM CMW_Category_1_And_4_Log
      ${where}
      ORDER BY CreatedAt DESC
      `,
      {
        type: QueryTypes.SELECT,
        replacements,
      },
    );

    return dataResults as any[];
  }

  async exportExcelCat1And4(dateFrom: string, dateTo: string, factory: string) {
    try {
      const { where, replacements } = buildLogFilter(dateFrom, dateTo, factory);

      const passThrough = new PassThrough();

      const bufferPromise: Promise<Buffer> = new Promise((resolve, reject) => {
        const chunks: Buffer[] = [];
        passThrough.on('data', (chunk) => chunks.push(chunk));
        passThrough.on('end', () => resolve(Buffer.concat(chunks)));
        passThrough.on('error', reject);
      });

      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        stream: passThrough,
        useStyles: true,
        useSharedStrings: false,
      });

      const worksheet = workbook.addWorksheet('LogCat');

      worksheet.columns = CAT1_AND_4_LOG_STREAM_COLUMNS.map((column) => ({
        ...column,
      }));

      worksheet.getRow(1).font = { bold: true };
      worksheet.getRow(1).commit();

      const CHUNK_SIZE = 10000;
      let offset = 0;

      while (true) {
        const rows = await this.EIP.query(
          `SELECT * FROM CMW_Category_1_And_4_Log
           ${where}
           ORDER BY Id
           OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`,
          {
            type: QueryTypes.SELECT,
            replacements: [...replacements, offset, CHUNK_SIZE],
          },
        );

        if (rows.length === 0) break;

        for (const row of rows) {
          worksheet.addRow(row).commit();
        }

        offset += CHUNK_SIZE;
      }

      await worksheet.commit();
      await workbook.commit();

      const buffer = await bufferPromise;

      return buffer;
    } catch (error) {
      throw new InternalServerErrorException(error);
    }
  }

  // Logging CAT5
  async getLogCat5(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    return this.getPagedLog(
      LOG_TABLE.cat5,
      dateFrom,
      dateTo,
      factory,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  async createLogCat5(data: CreateLogCat5[], factory: string, userid: string) {
    await this.withTransaction((transaction) =>
      insertLogRows(this.EIP, transaction, LOG_INSERT.cat5, data, {
        factory,
        userid,
      }),
    );

    return {
      success: true,
      message: 'Create log CAT5 success!',
    };
  }

  async exportExcelCat5(dateFrom: string, dateTo: string, factory: string) {
    return this.exportLogExcel(
      LOG_TABLE.cat5,
      LOG_EXPORT_COLUMNS.cat5,
      dateFrom,
      dateTo,
      factory,
    );
  }

  // Logging CAT6
  async getLogCat6BusinessTravel(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    return this.getPagedLog(
      LOG_TABLE.cat6BusinessTravel,
      dateFrom,
      dateTo,
      factory,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  async createLogCat6BusinessTravel(
    data: CreateLogCat6BusinessTravel[],
    factory: string,
    userid: string,
  ) {
    await this.withTransaction((transaction) =>
      insertLogRows(
        this.EIP,
        transaction,
        LOG_INSERT.cat6BusinessTravel,
        data,
        { factory, userid },
      ),
    );

    return {
      success: true,
      message: 'Create log CAT6 Business Travel!',
    };
  }

  async exportExcelCat6BusinessTravel(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    return this.exportLogExcel(
      LOG_TABLE.cat6BusinessTravel,
      LOG_EXPORT_COLUMNS.cat6BusinessTravel,
      dateFrom,
      dateTo,
      factory,
    );
  }

  async getLogCat6Accommodation(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    return this.getPagedLog(
      LOG_TABLE.cat6Accommodation,
      dateFrom,
      dateTo,
      factory,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  async createLogCat6Accommodation(
    data: CreateLogCat6Accommodation[],
    factory: string,
    userid: string,
  ) {
    await this.withTransaction((transaction) =>
      insertLogRows(this.EIP, transaction, LOG_INSERT.cat6Accommodation, data, {
        factory,
        userid,
      }),
    );

    return {
      success: true,
      message: 'Create log CAT6 Accommodation!',
    };
  }

  async exportExcelCat6Accommodation(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    return this.exportLogExcel(
      LOG_TABLE.cat6Accommodation,
      LOG_EXPORT_COLUMNS.cat6Accommodation,
      dateFrom,
      dateTo,
      factory,
    );
  }

  // Logging CAT7
  async getLogCat7(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    return this.getPagedLog(
      LOG_TABLE.cat7,
      dateFrom,
      dateTo,
      factory,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  async createLogCat7(data: CreateLogCat7[], factory: string, userid: string) {
    await this.withTransaction((transaction) =>
      insertLogRows(this.EIP, transaction, LOG_INSERT.cat7, data, {
        factory,
        userid,
      }),
    );

    return {
      success: true,
      message: 'Create log CAT7 success!',
    };
  }

  async exportExcelCat7(dateFrom: string, dateTo: string, factory: string) {
    return this.exportLogExcel(
      LOG_TABLE.cat7,
      LOG_EXPORT_COLUMNS.cat7,
      dateFrom,
      dateTo,
      factory,
    );
  }

  // Logging CAT9 & 12
  async getLogCat9And12(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    return this.getPagedLog(
      LOG_TABLE.cat9And12,
      dateFrom,
      dateTo,
      factory,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  async createLogCat9And12(
    data: CreateLogCat9And12[],
    factory: string,
    userid: string,
  ) {
    await this.withTransaction(
      (transaction) =>
        insertLogRows(this.EIP, transaction, LOG_INSERT.cat9And12, data, {
          factory,
          userid,
        }),
      { logError: true },
    );

    return {
      success: true,
      message: 'Create log CAT9&12 success!',
    };
  }

  async exportExcelCat9And12(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    return this.exportLogExcel(
      LOG_TABLE.cat9And12,
      LOG_EXPORT_COLUMNS.cat9And12,
      dateFrom,
      dateTo,
      factory,
    );
  }
}
