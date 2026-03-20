import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import {
  CreateLogCat1And4,
  CreateLogCat5,
  CreateLogCat6Accommodation,
  CreateLogCat6BusinessTravel,
  CreateLogCat7,
  CreateLogCat9And12,
} from './dto/create-logcat.dto';
import { v4 as uuidv4 } from 'uuid';
import * as ExcelJS from 'exceljs';
import { PassThrough } from 'stream';

@Injectable()
export class LogcatService {
  constructor(@Inject('EIP') private readonly EIP: Sequelize) {}

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
    const offset = (page - 1) * limit;

    let where: string = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ' AND Fac LIKE ?';
      replacements.push(`%${factory}%`);
    }

    const dataResults = await this.EIP.query(
      `
    SELECT *,
           COUNT(ID) OVER() AS total
    FROM CMW_Category_1_And_4_Log
    ${where}
    ORDER BY CreatedAt DESC
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

  async createLogCat1And4(
    data: CreateLogCat1And4[],
    factory: string,
    userid: string,
  ) {
    const transaction = await this.EIP.transaction();

    try {
      for (const item of data) {
        let fac: string = '';

        switch (item.Factory.trim().toLowerCase()) {
          case '樂億 - LYV'.trim().toLowerCase():
            fac = 'LYV';
            break;
          case '樂億II - LHG'.trim().toLowerCase():
            fac = 'LHG';
            break;
          case '億春B - LVL'.trim().toLowerCase():
            fac = 'LVL';
            break;
          case '昌億 - LYM'.trim().toLowerCase():
            fac = 'LYM';
            break;
          case '億福 - LYF'.trim().toLowerCase():
            fac = 'LYF';
            break;
          case 'Jiazhi-1'.trim().toLowerCase():
            fac = 'JAZ';
            break;
          case 'Jiazhi-2'.trim().toLowerCase():
            fac = 'JZS';
            break;
          default:
            break;
        }

        await this.EIP.query(
          `
          INSERT INTO CMW_Category_1_And_4_Log
          (
            ID, [System], Corporation, Factory, Fac, Department, DocKey,
            SPeriodData, EPeriodData, ActivityType, DataType, DocType,
            UndDoc, DocFlow, DocDate, DocDate2, DocNo, UndDocNo,
            CustVenName, InvoiceNo, TransType, Departure, Destination,
            PortType, StPort, ThPort, EndPort, Product, Quity, Amount,
            ActivityData, ActivityUnit, Unit, UnitWeight, Memo,
            CreateDateTime, Creator, CreatedUser, CreatedFactory, CreatedAt, ActivitySource
          )
          VALUES
          (
            :id, :system, :corporation, :factory, :fac, :department, :docKey,
            :sPeriod, :ePeriod, :activityType, :dataType, :docType,
            :undDoc, :docFlow, :docDate, :docDate2, :docNo, :undDocNo,
            :custVenName, :invoiceNo, :transType, :departure, :destination,
            :portType, :stPort, :thPort, :endPort, :product, :quity, :amount,
            :activityData, :activityUnit, :unit, :unitWeight, :memo,
            :createDateTime, :creator, :createdUser, :createdFactory, GETDATE(), :activitySource
          );
          `,
          {
            type: QueryTypes.INSERT,
            transaction,
            replacements: {
              id: uuidv4(),
              system: item.System,
              corporation: item.Corporation,
              factory: item.Factory,
              fac: fac,
              department: item.Department,
              docKey: item.DocKey,
              sPeriod: item.SPeriodData,
              ePeriod: item.EPeriodData,
              activityType: item.ActivityType,
              dataType: item.DataType,
              docType: item.DocType,
              undDoc: item.UndDoc,
              docFlow: item.DocFlow,
              docDate: item.DocDate,
              docDate2: item.DocDate2,
              docNo: item.DocNo,
              undDocNo: item.UndDocNo,
              custVenName: item.CustVenName,
              invoiceNo: item.InvoiceNo,
              transType: item.TransType,
              departure: item.Departure,
              destination: item.Destination,
              portType: item.PortType,
              stPort: item.StPort,
              thPort: item.ThPort,
              endPort: item.EndPort,
              product: item.Product,
              quity: item.Quity,
              amount: item.Amount,
              activityData: item.ActivityData,
              activityUnit: item.ActivityUnit,
              unit: item.Unit,
              unitWeight: item.UnitWeight,
              memo: item.Memo,
              createDateTime: item.CreateDateTime,
              creator: item.Creator,
              createdUser: userid,
              createdFactory: factory,
              activitySource: item.ActivitySource,
            },
          },
        );
      }

      await transaction.commit();

      return {
        success: true,
        message: 'Create log CAT1&4 success!',
      };
    } catch (error) {
      await transaction.rollback();
      throw new InternalServerErrorException('Error!');
    }
  }

  // async exportExcelCat1And4(dateFrom: string, dateTo: string, factory: string) {
  //   try {
  //     let where: string = 'WHERE 1=1';
  //     const replacements: any[] = [];

  //     if (dateFrom && dateTo) {
  //       where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
  //       replacements.push(dateFrom, dateTo);
  //     }

  //     if (factory) {
  //       where += ' AND Fac LIKE ?';
  //       replacements.push(`%${factory}%`);
  //     }

  //     const dataResults = await this.EIP.query(
  //       `
  //         SELECT *
  //         FROM CMW_Category_1_And_4_Log
  //         ${where}
  //       `,
  //       {
  //         type: QueryTypes.SELECT,
  //         replacements,
  //       },
  //     );

  //     let data = dataResults as any[];

  //     const workbook = new ExcelJS.Workbook();
  //     const worksheet = workbook.addWorksheet('LogCat');

  //     worksheet.columns = [
  //       { header: 'System', key: 'System' },
  //       { header: 'Corporation', key: 'Corporation' },
  //       { header: 'Factory', key: 'Factory' },
  //       { header: 'Department', key: 'Department' },
  //       { header: 'DocKey', key: 'DocKey' },
  //       { header: 'SPeriodData', key: 'SPeriodData' },
  //       { header: 'EPeriodData', key: 'EPeriodData' },
  //       { header: 'ActivityType', key: 'ActivityType' },
  //       { header: 'DataType', key: 'DataType' },
  //       { header: 'DocType', key: 'DocType' },
  //       { header: 'UndDoc', key: 'UndDoc' },
  //       { header: 'DocFlow', key: 'DocFlow' },
  //       { header: 'DocDate', key: 'DocDate' },
  //       { header: 'DocDate2', key: 'DocDate2' },
  //       { header: 'DocNo', key: 'DocNo' },
  //       { header: 'UndDocNo', key: 'UndDocNo' },
  //       { header: 'CustVenName', key: 'CustVenName' },
  //       { header: 'InvoiceNo', key: 'InvoiceNo' },
  //       { header: 'TransType', key: 'TransType' },
  //       { header: 'Departure', key: 'Departure' },
  //       { header: 'Destination', key: 'Destination' },
  //       { header: 'PortType', key: 'PortType' },
  //       { header: 'StPort', key: 'StPort' },
  //       { header: 'ThPort', key: 'ThPort' },
  //       { header: 'EndPort', key: 'EndPort' },
  //       { header: 'Product', key: 'Product' },
  //       { header: 'Quity', key: 'Quity' },
  //       { header: 'Amount', key: 'Amount' },
  //       { header: 'ActivityData', key: 'ActivityData' },
  //       { header: 'ActivityUnit', key: 'ActivityUnit' },
  //       { header: 'Unit', key: 'Unit' },
  //       { header: 'UnitWeight', key: 'UnitWeight' },
  //       { header: 'Memo', key: 'Memo' },
  //       { header: 'CreateDateTime', key: 'CreateDateTime' },
  //       { header: 'Creator', key: 'Creator' },
  //       { header: 'CreatedUser', key: 'CreatedUser' },
  //       { header: 'CreatedFactory', key: 'CreatedFactory' },
  //       { header: 'CreatedAt', key: 'CreatedAt' },
  //     ];

  //     worksheet.getRow(1).font = { bold: true };

  //     worksheet.addRows(data);

  //     worksheet.columns.forEach((column) => {
  //       let maxColumnLength = 0;

  //       if (typeof column.eachCell === 'function') {
  //         column.eachCell({ includeEmpty: true }, (cell) => {
  //           const cellValue = cell.value ? cell.value.toString() : '';

  //           if (cellValue.length > maxColumnLength) {
  //             maxColumnLength = cellValue.length;
  //           }
  //         });
  //       }

  //       const newWidth = maxColumnLength + 2;
  //       column.width = newWidth < 10 ? 10 : newWidth > 50 ? 50 : newWidth;
  //     });

  //     worksheet.eachRow((row) => {
  //       row.eachCell((cell) => {
  //         cell.border = {
  //           top: { style: 'thin' },
  //           left: { style: 'thin' },
  //           bottom: { style: 'thin' },
  //           right: { style: 'thin' },
  //         };
  //       });
  //     });

  //     const buffer = await workbook.xlsx.writeBuffer();
  //     return buffer;
  //   } catch (error) {
  //     throw new InternalServerErrorException(error);
  //   }
  // }

  // async exportExcelCat1And4(dateFrom: string, dateTo: string, factory: string) {
  //   try {
  //     let where: string = 'WHERE 1=1';
  //     const replacements: any[] = [];

  //     if (dateFrom && dateTo) {
  //       where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
  //       replacements.push(dateFrom, dateTo);
  //     }

  //     if (factory) {
  //       where += ' AND Fac LIKE ?';
  //       replacements.push(`%${factory}%`);
  //     }

  //     const passThrough = new PassThrough();

  //     const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
  //       stream: passThrough,
  //       useStyles: true,
  //       useSharedStrings: false,
  //     });

  //     const worksheet = workbook.addWorksheet('LogCat');

  //     worksheet.columns = [
  //       { header: 'System', key: 'System', width: 15 },
  //       { header: 'Corporation', key: 'Corporation', width: 15 },
  //       { header: 'Factory', key: 'Factory', width: 15 },
  //       { header: 'Department', key: 'Department', width: 15 },
  //       { header: 'DocKey', key: 'DocKey', width: 15 },
  //       { header: 'SPeriodData', key: 'SPeriodData', width: 15 },
  //       { header: 'EPeriodData', key: 'EPeriodData', width: 15 },
  //       { header: 'ActivityType', key: 'ActivityType', width: 15 },
  //       { header: 'DataType', key: 'DataType', width: 15 },
  //       { header: 'DocType', key: 'DocType', width: 15 },
  //       { header: 'UndDoc', key: 'UndDoc', width: 15 },
  //       { header: 'DocFlow', key: 'DocFlow', width: 15 },
  //       { header: 'DocDate', key: 'DocDate', width: 15 },
  //       { header: 'DocDate2', key: 'DocDate2', width: 15 },
  //       { header: 'DocNo', key: 'DocNo', width: 15 },
  //       { header: 'UndDocNo', key: 'UndDocNo', width: 15 },
  //       { header: 'CustVenName', key: 'CustVenName', width: 25 },
  //       { header: 'InvoiceNo', key: 'InvoiceNo', width: 15 },
  //       { header: 'TransType', key: 'TransType', width: 15 },
  //       { header: 'Departure', key: 'Departure', width: 15 },
  //       { header: 'Destination', key: 'Destination', width: 15 },
  //       { header: 'PortType', key: 'PortType', width: 15 },
  //       { header: 'StPort', key: 'StPort', width: 15 },
  //       { header: 'ThPort', key: 'ThPort', width: 15 },
  //       { header: 'EndPort', key: 'EndPort', width: 15 },
  //       { header: 'Product', key: 'Product', width: 20 },
  //       { header: 'Quity', key: 'Quity', width: 12 },
  //       { header: 'Amount', key: 'Amount', width: 12 },
  //       { header: 'ActivityData', key: 'ActivityData', width: 15 },
  //       { header: 'ActivityUnit', key: 'ActivityUnit', width: 15 },
  //       { header: 'Unit', key: 'Unit', width: 10 },
  //       { header: 'UnitWeight', key: 'UnitWeight', width: 12 },
  //       { header: 'Memo', key: 'Memo', width: 20 },
  //       { header: 'CreateDateTime', key: 'CreateDateTime', width: 20 },
  //       { header: 'Creator', key: 'Creator', width: 15 },
  //       { header: 'CreatedUser', key: 'CreatedUser', width: 15 },
  //       { header: 'CreatedFactory', key: 'CreatedFactory', width: 15 },
  //       { header: 'CreatedAt', key: 'CreatedAt', width: 20 },
  //     ];

  //     worksheet.getRow(1).font = { bold: true };
  //     worksheet.getRow(1).commit();

  //     const CHUNK_SIZE = 10000;
  //     let offset = 0;

  //     while (true) {
  //       const rows = (await this.EIP.query(
  //         `SELECT * FROM CMW_Category_1_And_4_Log
  //          ${where}
  //          ORDER BY (SELECT NULL) OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`,
  //         {
  //           type: QueryTypes.SELECT,
  //           replacements: [...replacements, offset, CHUNK_SIZE],
  //         },
  //       )) as any[];

  //       if (rows.length === 0) break;

  //       for (const row of rows) {
  //         worksheet.addRow(row).commit();
  //       }

  //       offset += CHUNK_SIZE;
  //     }

  //     await worksheet.commit();
  //     await workbook.commit();

  //     const buffer: Buffer = await new Promise((resolve, reject) => {
  //       const chunks: Buffer[] = [];
  //       passThrough.on('data', (chunk) => chunks.push(chunk));
  //       passThrough.on('end', () => resolve(Buffer.concat(chunks)));
  //       passThrough.on('error', reject);
  //     });

  //     return buffer;
  //   } catch (error) {
  //     throw new InternalServerErrorException(error);
  //   }
  // }

  async exportExcelCat1And4(dateFrom: string, dateTo: string, factory: string) {
    try {
      let where: string = 'WHERE 1=1';
      const replacements: any[] = [];

      if (dateFrom && dateTo) {
        where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
        replacements.push(dateFrom, dateTo);
      }

      if (factory) {
        where += ' AND Fac LIKE ?';
        replacements.push(`%${factory}%`);
      }

      const passThrough = new PassThrough();

      // ✅ Gắn listener TRƯỚC — không bỏ lỡ event 'end'
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

      worksheet.columns = [
        { header: 'System', key: 'System', width: 15 },
        { header: 'Corporation', key: 'Corporation', width: 15 },
        { header: 'Factory', key: 'Factory', width: 15 },
        { header: 'Department', key: 'Department', width: 15 },
        { header: 'DocKey', key: 'DocKey', width: 15 },
        { header: 'SPeriodData', key: 'SPeriodData', width: 15 },
        { header: 'EPeriodData', key: 'EPeriodData', width: 15 },
        { header: 'ActivityType', key: 'ActivityType', width: 15 },
        { header: 'DataType', key: 'DataType', width: 15 },
        { header: 'DocType', key: 'DocType', width: 15 },
        { header: 'UndDoc', key: 'UndDoc', width: 15 },
        { header: 'DocFlow', key: 'DocFlow', width: 15 },
        { header: 'DocDate', key: 'DocDate', width: 15 },
        { header: 'DocDate2', key: 'DocDate2', width: 15 },
        { header: 'DocNo', key: 'DocNo', width: 15 },
        { header: 'UndDocNo', key: 'UndDocNo', width: 15 },
        { header: 'CustVenName', key: 'CustVenName', width: 25 },
        { header: 'InvoiceNo', key: 'InvoiceNo', width: 15 },
        { header: 'TransType', key: 'TransType', width: 15 },
        { header: 'Departure', key: 'Departure', width: 15 },
        { header: 'Destination', key: 'Destination', width: 15 },
        { header: 'PortType', key: 'PortType', width: 15 },
        { header: 'StPort', key: 'StPort', width: 15 },
        { header: 'ThPort', key: 'ThPort', width: 15 },
        { header: 'EndPort', key: 'EndPort', width: 15 },
        { header: 'Product', key: 'Product', width: 20 },
        { header: 'Quity', key: 'Quity', width: 12 },
        { header: 'Amount', key: 'Amount', width: 12 },
        { header: 'ActivityData', key: 'ActivityData', width: 15 },
        { header: 'ActivityUnit', key: 'ActivityUnit', width: 15 },
        { header: 'Unit', key: 'Unit', width: 10 },
        { header: 'UnitWeight', key: 'UnitWeight', width: 12 },
        { header: 'Memo', key: 'Memo', width: 20 },
        { header: 'CreateDateTime', key: 'CreateDateTime', width: 20 },
        { header: 'Creator', key: 'Creator', width: 15 },
        { header: 'CreatedUser', key: 'CreatedUser', width: 15 },
        { header: 'CreatedFactory', key: 'CreatedFactory', width: 15 },
        { header: 'CreatedAt', key: 'CreatedAt', width: 20 },
        { header: 'ActivitySource', key: 'ActivitySource', width: 20 },
      ];

      worksheet.getRow(1).font = { bold: true };
      worksheet.getRow(1).commit();

      const CHUNK_SIZE = 10000;
      let offset = 0;

      while (true) {
        const rows = (await this.EIP.query(
          `SELECT * FROM CMW_Category_1_And_4_Log
           ${where}
           ORDER BY (SELECT NULL) OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`,
          {
            type: QueryTypes.SELECT,
            replacements: [...replacements, offset, CHUNK_SIZE],
          },
        )) as any[];

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
    const offset = (page - 1) * limit;

    let where: string = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ' AND Fac LIKE ?';
      replacements.push(`%${factory}%`);
    }

    const dataResults = await this.EIP.query(
      `
    SELECT *,
           COUNT(ID) OVER() AS total
    FROM CMW_Category_5_Log
    ${where}
    ORDER BY CreatedAt DESC
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

  async createLogCat5(data: CreateLogCat5[], factory: string, userid: string) {
    const transaction = await this.EIP.transaction();

    try {
      for (const item of data) {
        let fac: string = '';

        switch (item.Factory.trim().toLowerCase()) {
          case '樂億 - LYV'.trim().toLowerCase():
            fac = 'LYV';
            break;
          case '樂億II - LHG'.trim().toLowerCase():
            fac = 'LHG';
            break;
          case '億春B - LVL'.trim().toLowerCase():
            fac = 'LVL';
            break;
          case '昌億 - LYM'.trim().toLowerCase():
            fac = 'LYM';
            break;
          case '億福 - LYF'.trim().toLowerCase():
            fac = 'LYF';
            break;
          case 'Jiazhi-1'.trim().toLowerCase():
            fac = 'JAZ';
            break;
          case 'Jiazhi-2'.trim().toLowerCase():
            fac = 'JZS';
            break;
          default:
            break;
        }

        await this.EIP.query(
          `
          INSERT INTO CMW_Category_5_Log
          (
            ID, [System], Corporation, Factory, Fac, Department, DocKey,
            SPeriodData, EPeriodData, ActivityType, DataType, DocType,
            UndDoc, DocFlow, DocDate, DocDate2, DocNo, UndDocNo,
            CustVenName, InvoiceNo, TransType, Departure, Destination,
            PortType, StPort, ThPort, EndPort, Product, Quity, Amount,
            ActivityData, ActivityUnit, Unit, UnitWeight, Memo,
            CreateDateTime, Creator, CreatedUser, CreatedFactory, CreatedAt, ActivitySource
          )
          VALUES
          (
            :id, :system, :corporation, :factory, :fac, :department, :docKey,
            :sPeriod, :ePeriod, :activityType, :dataType, :docType,
            :undDoc, :docFlow, :docDate, :docDate2, :docNo, :undDocNo,
            :custVenName, :invoiceNo, :transType, :departure, :destination,
            :portType, :stPort, :thPort, :endPort, :product, :quity, :amount,
            :activityData, :activityUnit, :unit, :unitWeight, :memo,
            :createDateTime, :creator, :createdUser, :createdFactory, GETDATE(), :activitySource
          )
          `,
          {
            type: QueryTypes.INSERT,
            transaction,
            replacements: {
              id: uuidv4(),
              system: item.System,
              corporation: item.Corporation,
              factory: item.Factory,
              fac: fac,
              department: item.Department,
              docKey: item.DocKey,
              sPeriod: item.SPeriodData,
              ePeriod: item.EPeriodData,
              activityType: item.ActivityType,
              dataType: item.DataType,
              docType: item.DocType,
              undDoc: item.UndDoc,
              docFlow: item.DocFlow,
              docDate: item.DocDate,
              docDate2: item.DocDate2,
              docNo: item.DocNo,
              undDocNo: item.UndDocNo,
              custVenName: item.CustVenName,
              invoiceNo: item.InvoiceNo,
              transType: item.TransType,
              departure: item.Departure,
              destination: item.Destination,
              portType: item.PortType,
              stPort: item.StPort,
              thPort: item.ThPort,
              endPort: item.EndPort,
              product: item.Product,
              quity: item.Quity,
              amount: item.Amount,
              activityData: item.ActivityData,
              activityUnit: item.ActivityUnit,
              unit: item.Unit,
              unitWeight: item.UnitWeight,
              memo: item.Memo,
              createDateTime: item.CreateDateTime,
              creator: item.Creator,
              createdUser: userid,
              createdFactory: factory,
              activitySource: item.ActivitySource,
            },
          },
        );
      }

      await transaction.commit();

      return {
        success: true,
        message: 'Create log CAT5 success!',
      };
    } catch (error) {
      await transaction.rollback();
      throw new InternalServerErrorException('Error!');
    }
  }

  async exportExcelCat5(dateFrom: string, dateTo: string, factory: string) {
    try {
      let where: string = 'WHERE 1=1';
      const replacements: any[] = [];
      if (dateFrom && dateTo) {
        where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
        replacements.push(dateFrom, dateTo);
      }

      if (factory) {
        where += ' AND Fac LIKE ?';
        replacements.push(`%${factory}%`);
      }

      const dataResults = await this.EIP.query(
        `
          SELECT *
          FROM CMW_Category_5_Log
          ${where}
        `,
        {
          type: QueryTypes.SELECT,
          replacements,
        },
      );
      let data = dataResults as any[];

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('LogCat');

      worksheet.columns = [
        { header: 'System', key: 'System' },
        { header: 'Corporation', key: 'Corporation' },
        { header: 'Factory', key: 'Factory' },
        { header: 'Fac', key: 'Fac' },
        { header: 'Department', key: 'Department' },
        { header: 'DocKey', key: 'DocKey' },
        { header: 'SPeriodData', key: 'SPeriodData' },
        { header: 'EPeriodData', key: 'EPeriodData' },
        { header: 'ActivityType', key: 'ActivityType' },
        { header: 'DataType', key: 'DataType' },
        { header: 'DocType', key: 'DocType' },
        { header: 'UndDoc', key: 'UndDoc' },
        { header: 'DocFlow', key: 'DocFlow' },
        { header: 'DocDate', key: 'DocDate' },
        { header: 'DocDate2', key: 'DocDate2' },
        { header: 'DocNo', key: 'DocNo' },
        { header: 'UndDocNo', key: 'UndDocNo' },
        { header: 'CustVenName', key: 'CustVenName' },
        { header: 'InvoiceNo', key: 'InvoiceNo' },
        { header: 'TransType', key: 'TransType' },
        { header: 'Departure', key: 'Departure' },
        { header: 'Destination', key: 'Destination' },
        { header: 'PortType', key: 'PortType' },
        { header: 'StPort', key: 'StPort' },
        { header: 'ThPort', key: 'ThPort' },
        { header: 'EndPort', key: 'EndPort' },
        { header: 'Product', key: 'Product' },
        { header: 'Quity', key: 'Quity' },
        { header: 'Amount', key: 'Amount' },
        { header: 'ActivityData', key: 'ActivityData' },
        { header: 'ActivityUnit', key: 'ActivityUnit' },
        { header: 'Unit', key: 'Unit' },
        { header: 'UnitWeight', key: 'UnitWeight' },
        { header: 'Memo', key: 'Memo' },
        { header: 'CreateDateTime', key: 'CreateDateTime' },
        { header: 'Creator', key: 'Creator' },
        { header: 'CreatedUser', key: 'CreatedUser' },
        { header: 'CreatedFactory', key: 'CreatedFactory' },
        { header: 'CreatedAt', key: 'CreatedAt' },
        { header: 'ActivitySource', key: 'ActivitySource' },
      ];

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
    const offset = (page - 1) * limit;

    let where: string = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ' AND Fac LIKE ?';
      replacements.push(`%${factory}%`);
    }

    const dataResults = await this.EIP.query(
      `
        SELECT *,
              COUNT(ID) OVER() AS total
        FROM CMW_Category_6_Business_Travel
        ${where}
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

  async createLogCat6BusinessTravel(
    data: CreateLogCat6BusinessTravel[],
    factory: string,
    userid: string,
  ) {
    const transaction = await this.EIP.transaction();

    try {
      for (const item of data) {
        let fac: string = '';

        switch (item.Factory.trim().toLowerCase()) {
          case '樂億 - LYV'.trim().toLowerCase():
            fac = 'LYV';
            break;
          case '樂億II - LHG'.trim().toLowerCase():
            fac = 'LHG';
            break;
          case '億春B - LVL'.trim().toLowerCase():
            fac = 'LVL';
            break;
          case '昌億 - LYM'.trim().toLowerCase():
            fac = 'LYM';
            break;
          case '億福 - LYF'.trim().toLowerCase():
            fac = 'LYF';
            break;
          case 'Jiazhi-1'.trim().toLowerCase():
            fac = 'JAZ';
            break;
          case 'Jiazhi-2'.trim().toLowerCase():
            fac = 'JZS';
            break;
          default:
            break;
        }

        await this.EIP.query(
          `
          INSERT INTO CMW_Category_6_Business_Travel
          (
            ID, [System], Corporation, Factory, Fac, Department, DocKey, ActivitySource,
            SPeriodData, EPeriodData, ActivityType, DataType, DocType, DocDate, DocDate2,
            DocNo, UndDocNo, TransType, Departure, Destination, Memo,
            CreateDateTime, Creator, CreatedUser, CreatedFactory, CreatedAt
          )
          VALUES
          (
            :id, :system, :corporation, :factory, :fac, :department, :docKey, :activitySource,
            :sPeriodData, :ePeriodData, :activityType, :dataType, :docType, :docDate, :docDate2,
            :docNo, :undDocNo, :transType, :departure, :destination, :memo,
            :createDateTime, :creator, :createdUser, :createdFactory, GETDATE()
          )
          `,
          {
            type: QueryTypes.INSERT,
            transaction,
            replacements: {
              id: uuidv4(),
              system: item.System,
              corporation: item.Corporation,
              factory: item.Factory,
              fac,
              department: item.Department,
              docKey: item.DocKey,
              activitySource: item.ActivitySource,
              sPeriodData: item.SPeriodData,
              ePeriodData: item.EPeriodData,
              activityType: item.ActivityType,
              dataType: item.DataType,
              docType: item.DocType,
              docDate: item.DocDate,
              docDate2: item.DocDate2,
              docNo: item.DocNo,
              undDocNo: item.UndDocNo,
              transType: item.TransType,
              departure: item.Departure,
              destination: item.Destination,
              memo: item.Memo,
              createDateTime: item.CreateDateTime,
              creator: item.Creator,
              createdUser: userid,
              createdFactory: factory,
            },
          },
        );
      }

      await transaction.commit();

      return {
        success: true,
        message: 'Create log CAT6 Business Travel!',
      };
    } catch (error) {
      await transaction.rollback();
      throw new InternalServerErrorException('Error!');
    }
  }

  async exportExcelCat6BusinessTravel(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    try {
      let where: string = 'WHERE 1=1';
      const replacements: any[] = [];
      if (dateFrom && dateTo) {
        where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
        replacements.push(dateFrom, dateTo);
      }

      if (factory) {
        where += ' AND Fac LIKE ?';
        replacements.push(`%${factory}%`);
      }

      const dataResults = await this.EIP.query(
        `
          SELECT *
          FROM CMW_Category_6_Business_Travel
          ${where}
        `,
        {
          type: QueryTypes.SELECT,
          replacements,
        },
      );
      let data = dataResults as any[];

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('LogCat');

      worksheet.columns = [
        { header: 'System', key: 'System' },
        { header: 'Corporation', key: 'Corporation' },
        { header: 'Factory', key: 'Factory' },
        { header: 'Fac', key: 'Fac' },
        { header: 'Department', key: 'Department' },
        { header: 'DocKey', key: 'DocKey' },
        { header: 'ActivitySource', key: 'ActivitySource' },
        { header: 'SPeriodData', key: 'SPeriodData' },
        { header: 'EPeriodData', key: 'EPeriodData' },
        { header: 'ActivityType', key: 'ActivityType' },
        { header: 'DataType', key: 'DataType' },
        { header: 'DocType', key: 'DocType' },
        { header: 'DocDate', key: 'DocDate' },
        { header: 'DocDate2', key: 'DocDate2' },
        { header: 'DocNo', key: 'DocNo' },
        { header: 'UndDocNo', key: 'UndDocNo' },
        { header: 'TransType', key: 'TransType' },
        { header: 'Departure', key: 'Departure' },
        { header: 'Destination', key: 'Destination' },
        { header: 'Memo', key: 'Memo' },
        { header: 'CreateDateTime', key: 'CreateDateTime' },
        { header: 'Creator', key: 'Creator' },
        { header: 'CreatedUser', key: 'CreatedUser' },
        { header: 'CreatedFactory', key: 'CreatedFactory' },
        { header: 'CreatedAt', key: 'CreatedAt' },
      ];

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

  async getLogCat6Accommodation(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    const offset = (page - 1) * limit;

    let where: string = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ' AND Fac LIKE ?';
      replacements.push(`%${factory}%`);
    }

    const dataResults = await this.EIP.query(
      `
        SELECT *,
              COUNT(ID) OVER() AS total
        FROM CMW_Category_6_Accommodation
        ${where}
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

  async createLogCat6Accommodation(
    data: CreateLogCat6Accommodation[],
    factory: string,
    userid: string,
  ) {
    const transaction = await this.EIP.transaction();

    try {
      for (const item of data) {
        let fac: string = '';

        switch (item.Factory.trim().toLowerCase()) {
          case '樂億 - LYV'.trim().toLowerCase():
            fac = 'LYV';
            break;
          case '樂億II - LHG'.trim().toLowerCase():
            fac = 'LHG';
            break;
          case '億春B - LVL'.trim().toLowerCase():
            fac = 'LVL';
            break;
          case '昌億 - LYM'.trim().toLowerCase():
            fac = 'LYM';
            break;
          case '億福 - LYF'.trim().toLowerCase():
            fac = 'LYF';
            break;
          case 'Jiazhi-1'.trim().toLowerCase():
            fac = 'JAZ';
            break;
          case 'Jiazhi-2'.trim().toLowerCase():
            fac = 'JZS';
            break;
          default:
            break;
        }

        await this.EIP.query(
          `
          INSERT INTO CMW_Category_6_Accommodation
          (
            ID, [System], Corporation, Factory, Fac, Department, DocKey, ActivitySource,
            SPeriodData, EPeriodData, ActivityType, DataType, DocType, DocDate, DocDate2,
            DocNo, UndDocNo, TransType, ActivityData, Memo,
            CreateDateTime, Creator, CreatedUser, CreatedFactory, CreatedAt
          )
          VALUES
          (
            :id, :system, :corporation, :factory, :fac, :department, :docKey, :activitySource,
            :sPeriodData, :ePeriodData, :activityType, :dataType, :docType, :docDate, :docDate2,
            :docNo, :undDocNo, :transType, :activityData, :memo,
            :createDateTime, :creator, :createdUser, :createdFactory, GETDATE()
          )
          `,
          {
            type: QueryTypes.INSERT,
            transaction,
            replacements: {
              id: uuidv4(),
              system: item.System,
              corporation: item.Corporation,
              factory: item.Factory,
              fac,
              department: item.Department,
              docKey: item.DocKey,
              activitySource: item.ActivitySource,
              sPeriodData: item.SPeriodData,
              ePeriodData: item.EPeriodData,
              activityType: item.ActivityType,
              dataType: item.DataType,
              docType: item.DocKey,
              docDate: item.DocDate,
              docDate2: item.DocDate2,
              docNo: item.DocNo,
              undDocNo: item.UndDocNo,
              transType: item.TransType,
              activityData: item.ActivityData,
              memo: item.Memo,
              createDateTime: item.CreateDateTime,
              creator: item.Creator,
              createdUser: userid,
              createdFactory: factory,
            },
          },
        );
      }

      await transaction.commit();

      return {
        success: true,
        message: 'Create log CAT6 Accommodatio!',
      };
    } catch (error) {
      await transaction.rollback();
      throw new InternalServerErrorException('Error!');
    }
  }

  async exportExcelCat6Accommodation(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    try {
      let where: string = 'WHERE 1=1';
      const replacements: any[] = [];
      if (dateFrom && dateTo) {
        where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
        replacements.push(dateFrom, dateTo);
      }

      if (factory) {
        where += ' AND Fac LIKE ?';
        replacements.push(`%${factory}%`);
      }

      const dataResults = await this.EIP.query(
        `
          SELECT *
          FROM CMW_Category_6_Accommodation
          ${where}
        `,
        {
          type: QueryTypes.SELECT,
          replacements,
        },
      );
      let data = dataResults as any[];

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('LogCat');

      worksheet.columns = [
        { header: 'System', key: 'System' },
        { header: 'Corporation', key: 'Corporation' },
        { header: 'Factory', key: 'Factory' },
        { header: 'Fac', key: 'Fac' },
        { header: 'Department', key: 'Department' },
        { header: 'DocKey', key: 'DocKey' },
        { header: 'ActivitySource', key: 'ActivitySource' },
        { header: 'SPeriodData', key: 'SPeriodData' },
        { header: 'EPeriodData', key: 'EPeriodData' },
        { header: 'ActivityType', key: 'ActivityType' },
        { header: 'DataType', key: 'DataType' },
        { header: 'DocType', key: 'DocType' },
        { header: 'DocDate', key: 'DocDate' },
        { header: 'DocDate2', key: 'DocDate2' },
        { header: 'DocNo', key: 'DocNo' },
        { header: 'UndDocNo', key: 'UndDocNo' },
        { header: 'TransType', key: 'TransType' },
        { header: 'ActivityData', key: 'ActivityData' },
        { header: 'Memo', key: 'Memo' },
        { header: 'CreateDateTime', key: 'CreateDateTime' },
        { header: 'Creator', key: 'Creator' },
        { header: 'CreatedUser', key: 'CreatedUser' },
        { header: 'CreatedFactory', key: 'CreatedFactory' },
        { header: 'CreatedAt', key: 'CreatedAt' },
      ];

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
    const offset = (page - 1) * limit;

    let where: string = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ' AND Fac LIKE ?';
      replacements.push(`%${factory}%`);
    }

    const dataResults = await this.EIP.query(
      `
    SELECT *,
           COUNT(ID) OVER() AS total
    FROM CMW_Category_7_Log
    ${where}
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

  async createLogCat7(data: CreateLogCat7[], factory: string, userid: string) {
    const transaction = await this.EIP.transaction();

    try {
      for (const item of data) {
        let fac: string = '';

        switch (item.Factory.trim().toLowerCase()) {
          case '樂億 - LYV'.trim().toLowerCase():
            fac = 'LYV';
            break;
          case '樂億II - LHG'.trim().toLowerCase():
            fac = 'LHG';
            break;
          case '億春B - LVL'.trim().toLowerCase():
            fac = 'LVL';
            break;
          case '昌億 - LYM'.trim().toLowerCase():
            fac = 'LYM';
            break;
          case '億福 - LYF'.trim().toLowerCase():
            fac = 'LYF';
            break;
          case 'Jiazhi-1'.trim().toLowerCase():
            fac = 'JAZ';
            break;
          case 'Jiazhi-2'.trim().toLowerCase():
            fac = 'JZS';
            break;
          default:
            break;
        }

        await this.EIP.query(
          `
          INSERT INTO CMW_Category_7_Log
          (
            ID,
            [System],
            Corporation,
            Factory,
            Fac,
            Department,
            DocKey,
            SPeriodData,
            EPeriodData,
            ActivityType,
            DataType,
            DocType,
            DocDate,
            DocDate2,
            UndDocNo,
            TransType,
            Departure,
            Destination,
            Attendance,
            Memo,
            CreateDateTime,
            Creator,
            CreatedUser,
            CreatedFactory,
            CreatedAt
          )
          VALUES
          (
            :id,
            :system,
            :corporation,
            :factory,
            :fac,
            :department,
            :docKey,
            :sPeriodData,
            :ePeriodData,
            :activityType,
            :dataType,
            :docType,
            :docDate,
            :docDate2,
            :undDocNo,
            :transType,
            :departure,
            :destination,
            :attendance,
            :memo,
            :createDateTime,
            :creator,
            :createdUser,
            :createdFactory,
            GETDATE()
          )
          `,
          {
            type: QueryTypes.INSERT,
            transaction,
            replacements: {
              id: uuidv4(),
              system: item.System,
              corporation: item.Corporation,
              factory: item.Factory,
              fac: fac,
              department: item.Department,
              docKey: item.DocKey,
              sPeriodData: item.SPeriodData,
              ePeriodData: item.EPeriodData,
              activityType: item.ActivityType,
              dataType: item.DataType,
              docType: item.DocType,
              docDate: item.DocDate,
              docDate2: item.DocDate2,
              undDocNo: item.UndDocNo,
              transType: item.TransType,
              departure: item.Departure,
              destination: item.Destination,
              attendance: item.Attendance,
              memo: item.Memo,
              createDateTime: item.CreateDateTime,
              creator: item.Creator,
              createdUser: userid,
              createdFactory: factory,
            },
          },
        );
      }

      await transaction.commit();

      return {
        success: true,
        message: 'Create log CAT7 success!',
      };
    } catch (error) {
      await transaction.rollback();
      throw new InternalServerErrorException('Error!');
    }
  }

  async exportExcelCat7(dateFrom: string, dateTo: string, factory: string) {
    try {
      let where: string = 'WHERE 1=1';
      const replacements: any[] = [];

      if (dateFrom && dateTo) {
        where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
        replacements.push(dateFrom, dateTo);
      }

      if (factory) {
        where += ' AND Fac LIKE ?';
        replacements.push(`%${factory}%`);
      }

      const dataResults = await this.EIP.query(
        `
          SELECT *
          FROM CMW_Category_7_Log
          ${where}
        `,
        {
          type: QueryTypes.SELECT,
          replacements,
        },
      );

      let data = dataResults as any[];

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('LogCat');

      worksheet.columns = [
        { header: 'System', key: 'System' },
        { header: 'Corporation', key: 'Corporation' },
        { header: 'Factory', key: 'Factory' },
        { header: 'Department', key: 'Department' },
        { header: 'DocKey', key: 'DocKey' },
        { header: 'SPeriodData', key: 'SPeriodData' },
        { header: 'EPeriodData', key: 'EPeriodData' },
        { header: 'ActivityType', key: 'ActivityType' },
        { header: 'DataType', key: 'DataType' },
        { header: 'DocType', key: 'DocType' },
        { header: 'DocDate', key: 'DocDate' },
        { header: 'DocDate2', key: 'DocDate2' },
        { header: 'UndDocNo', key: 'UndDocNo' },
        { header: 'TransType', key: 'TransType' },
        { header: 'Departure', key: 'Departure' },
        { header: 'Destination', key: 'Destination' },
        { header: 'Attendance', key: 'Attendance' },
        { header: 'Memo', key: 'Memo' },
        { header: 'CreateDateTime', key: 'CreateDateTime' },
        { header: 'Creator', key: 'Creator' },
        { header: 'CreatedUser', key: 'CreatedUser' },
        { header: 'CreatedFactory', key: 'CreatedFactory' },
        { header: 'CreatedAt', key: 'CreatedAt' },
      ];

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
    const offset = (page - 1) * limit;

    let where: string = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ' AND Fac LIKE ?';
      replacements.push(`%${factory}%`);
    }

    const dataResults = await this.EIP.query(
      `
    SELECT *,
           COUNT(ID) OVER() AS total
    FROM CMW_Category_9_And_12_Log
    ${where}
    ORDER BY CreatedAt DESC
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

  async createLogCat9And12(
    data: CreateLogCat9And12[],
    factory: string,
    userid: string,
  ) {
    const transaction = await this.EIP.transaction();

    try {
      for (const item of data) {
        let fac: string = '';

        switch (item.Factory.trim().toLowerCase()) {
          case '樂億 - LYV'.trim().toLowerCase():
            fac = 'LYV';
            break;
          case '樂億II - LHG'.trim().toLowerCase():
            fac = 'LHG';
            break;
          case '億春B - LVL'.trim().toLowerCase():
            fac = 'LVL';
            break;
          case '昌億 - LYM'.trim().toLowerCase():
            fac = 'LYM';
            break;
          case '億福 - LYF'.trim().toLowerCase():
            fac = 'LYF';
            break;
          case 'Jiazhi-1'.trim().toLowerCase():
            fac = 'JAZ';
            break;
          case 'Jiazhi-2'.trim().toLowerCase():
            fac = 'JZS';
            break;
          default:
            break;
        }

        await this.EIP.query(
          `
          INSERT INTO CMW_Category_9_And_12_Log
          (
            ID,
            [System],
            Corporation,
            Factory,
            Fac,
            Department,
            DocKey,
            SPeriodData,
            EPeriodData,
            ActivityType,
            DataType,
            DocType,
            UndDoc,
            DocFlow,
            DocDate,
            DocDate2,
            DocNo,
            UndDocNo,
            CustVenName,
            InvoiceNo,
            TransType,
            Departure,
            Destination,
            PortType,
            StPort,
            ThPort,
            EndPort,
            Product,
            Quity,
            Amount,
            ActivityData,
            ActivityUnit,
            Unit,
            UnitWeight,
            Memo,
            CreateDateTime,
            Creator,
            ActivitySource,
            CreatedUser,
            CreatedFactory,
            CreatedAt,
            ActivitySource
          )
          VALUES
          (
            :id,
            :system,
            :corporation,
            :factory,
            :fac,
            :department,
            :docKey,
            :sPeriodData,
            :ePeriodData,
            :activityType,
            :dataType,
            :docType,
            :undDoc,
            :docFlow,
            :docDate,
            :docDate2,
            :docNo,
            :undDocNo,
            :custVenName,
            :invoiceNo,
            :transType,
            :departure,
            :destination,
            :portType,
            :stPort,
            :thPort,
            :endPort,
            :product,
            :quity,
            :amount,
            :activityData,
            :activityUnit,
            :unit,
            :unitWeight,
            :memo,
            :createDateTime,
            :creator,
            :activitySource,
            :createdUser,
            :createdFactory,
            GETDATE()
          )
          `,
          {
            type: QueryTypes.INSERT,
            transaction,
            replacements: {
              id: uuidv4(),
              system: item.System,
              corporation: item.Corporation,
              factory: item.Factory,
              fac: fac,
              department: item.Department,
              docKey: item.DocKey,
              sPeriodData: item.SPeriodData,
              ePeriodData: item.EPeriodData,
              activityType: item.ActivityType,
              dataType: item.DataType,
              docType: item.DocType,
              undDoc: item.UndDoc,
              docFlow: item.DocFlow,
              docDate: item.DocDate,
              docDate2: item.DocDate2,
              docNo: item.DocNo,
              undDocNo: item.UndDocNo,
              custVenName: item.CustVenName,
              invoiceNo: item.InvoiceNo,
              transType: item.TransType,
              departure: item.Departure,
              destination: item.Destination,
              portType: item.PortType,
              stPort: item.StPort,
              thPort: item.ThPort,
              endPort: item.EndPort,
              product: item.Product,
              quity: item.Quity,
              amount: item.Amount,
              activityData: item.ActivityData,
              activityUnit: item.ActivityUnit,
              unit: item.Unit,
              unitWeight: item.UnitWeight,
              memo: item.Memo,
              createDateTime: item.CreateDateTime,
              creator: item.Creator,
              activitySource: item.ActivitySource,
              createdUser: userid,
              createdFactory: factory,
            },
          },
        );
      }

      await transaction.commit();

      return {
        success: true,
        message: 'Create log CAT9&12 success!',
      };
    } catch (error) {
      console.log(error);
      await transaction.rollback();
      throw new InternalServerErrorException('Error!');
    }
  }

  async exportExcelCat9And12(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    try {
      let where: string = 'WHERE 1=1';
      const replacements: any[] = [];

      if (dateFrom && dateTo) {
        where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
        replacements.push(dateFrom, dateTo);
      }

      if (factory) {
        where += ' AND Fac LIKE ?';
        replacements.push(`%${factory}%`);
      }

      const dataResults = await this.EIP.query(
        `
          SELECT *
          FROM CMW_Category_9_And_12_Log
          ${where}
        `,
        {
          type: QueryTypes.SELECT,
          replacements,
        },
      );

      let data = dataResults as any[];

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('LogCat');

      worksheet.columns = [
        { header: 'System', key: 'System' },
        { header: 'Corporation', key: 'Corporation' },
        { header: 'Factory', key: 'Factory' },
        { header: 'Department', key: 'Department' },
        { header: 'DocKey', key: 'DocKey' },
        { header: 'SPeriodData', key: 'SPeriodData' },
        { header: 'EPeriodData', key: 'EPeriodData' },
        { header: 'ActivityType', key: 'ActivityType' },
        { header: 'DataType', key: 'DataType' },
        { header: 'DocType', key: 'DocType' },
        { header: 'UndDoc', key: 'UndDoc' },
        { header: 'DocFlow', key: 'DocFlow' },
        { header: 'DocDate', key: 'DocDate' },
        { header: 'DocDate2', key: 'DocDate2' },
        { header: 'DocNo', key: 'DocNo' },
        { header: 'UndDocNo', key: 'UndDocNo' },
        { header: 'CustVenName', key: 'CustVenName' },
        { header: 'InvoiceNo', key: 'InvoiceNo' },
        { header: 'TransType', key: 'TransType' },
        { header: 'Departure', key: 'Departure' },
        { header: 'Destination', key: 'Destination' },
        { header: 'PortType', key: 'PortType' },
        { header: 'StPort', key: 'StPort' },
        { header: 'ThPort', key: 'ThPort' },
        { header: 'EndPort', key: 'EndPort' },
        { header: 'Product', key: 'Product' },
        { header: 'Quity', key: 'Quity' },
        { header: 'Amount', key: 'Amount' },
        { header: 'ActivityData', key: 'ActivityData' },
        { header: 'ActivityUnit', key: 'ActivityUnit' },
        { header: 'Unit', key: 'Unit' },
        { header: 'UnitWeight', key: 'UnitWeight' },
        { header: 'Memo', key: 'Memo' },
        { header: 'CreateDateTime', key: 'CreateDateTime' },
        { header: 'Creator', key: 'Creator' },
        { header: 'ActivitySource', key: 'ActivitySource' },
        { header: 'CreatedUser', key: 'CreatedUser' },
        { header: 'CreatedFactory', key: 'CreatedFactory' },
        { header: 'CreatedAt', key: 'CreatedAt' },
      ];

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
}
