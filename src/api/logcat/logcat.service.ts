import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import {
  CreateLogCat5,
  CreateLogCat7,
  CreateLogCat9And12,
} from './dto/create-logcat.dto';
import { v4 as uuidv4 } from 'uuid';

@Injectable()
export class LogcatService {
  constructor(@Inject('EIP') private readonly EIP: Sequelize) {}

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
      where += ' AND Factory LIKE ?';
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
        await this.EIP.query(
          `
          INSERT INTO CMW_Category_5_Log
          (
            ID, [System], Corporation, Factory, Department, DocKey,
            SPeriodData, EPeriodData, ActivityType, DataType, DocType,
            UndDoc, DocFlow, DocDate, DocDate2, DocNo, UndDocNo,
            CustVenName, InvoiceNo, TransType, Departure, Destination,
            PortType, StPort, ThPort, EndPort, Product, Quity, Amount,
            ActivityData, ActivityUnit, Unit, UnitWeight, Memo,
            CreateDateTime, Creator, CreatedUser, CreatedFactory, CreatedAt
          )
          VALUES
          (
            :id, :system, :corporation, :factory, :department, :docKey,
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
      where += ' AND Factory LIKE ?';
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
        await this.EIP.query(
          `
          INSERT INTO CMW_Category_7_Log
          (
            ID,
            [System],
            Corporation,
            Factory,
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
      where += ' AND Factory LIKE ?';
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
        await this.EIP.query(
          `
          INSERT INTO CMW_Category_9_And_12_Log
          (
            ID,
            [System],
            Corporation,
            Factory,
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
}
