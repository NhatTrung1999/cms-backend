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
  CreateLogCat7,
  CreateLogCat9And12,
} from './dto/create-logcat.dto';
import { v4 as uuidv4 } from 'uuid';
import * as ExcelJS from 'exceljs';

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
            CreateDateTime, Creator, CreatedUser, CreatedFactory, CreatedAt
          )
          VALUES
          (
            :id, :system, :corporation, :factory, :fac, :department, :docKey,
            :sPeriod, :ePeriod, :activityType, :dataType, :docType,
            :undDoc, :docFlow, :docDate, :docDate2, :docNo, :undDocNo,
            :custVenName, :invoiceNo, :transType, :departure, :destination,
            :portType, :stPort, :thPort, :endPort, :product, :quity, :amount,
            :activityData, :activityUnit, :unit, :unitWeight, :memo,
            :createDateTime, :creator, :createdUser, :createdFactory, GETDATE()
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

      const dataResults = await this.EIP.query(
        `
          SELECT *
          FROM CMW_Category_1_And_4_Log
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
            CreateDateTime, Creator, CreatedUser, CreatedFactory, CreatedAt
          )
          VALUES
          (
            :id, :system, :corporation, :factory, :fac, :department, :docKey,
            :sPeriod, :ePeriod, :activityType, :dataType, :docType,
            :undDoc, :docFlow, :docDate, :docDate2, :docNo, :undDocNo,
            :custVenName, :invoiceNo, :transType, :departure, :destination,
            :portType, :stPort, :thPort, :endPort, :product, :quity, :amount,
            :activityData, :activityUnit, :unit, :unitWeight, :memo,
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
              docType: item.DocKey,
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
