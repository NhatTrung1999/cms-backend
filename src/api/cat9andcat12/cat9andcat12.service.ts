import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { orderByClause } from 'src/helper/sql.helper';
import { QueryTypes } from 'sequelize';
import { IDataPortCode } from 'src/types/cat9andcat12';
import { v4 as uuidv4 } from 'uuid';
import {
  buildImportMessage,
  countUpserts,
  readImportRows,
} from 'src/helper/excel-import.helper';
import { getCat9AndCat12CMSRows } from 'src/helper/cat9andcat12.cms';
import { addPortsToRows, ShipmentRow } from 'src/helper/shipment.helper';
import { FACTORY_LIST, FactoryCode } from 'src/helper/factory.helper';
import { sortByField } from 'src/helper/sort.helper';
import { getOwn } from 'src/helper/object.helper';
import {
  buildCat9AndCat12AllCountQuery,
  buildCat9AndCat12AllPageQuery,
  buildCat9AndCat12CountQuery,
  buildCat9AndCat12Filter,
  buildCat9AndCat12PageQuery,
} from 'src/helper/cat9andcat12.queries';

const ALL_PAGE_FACTORIES: FactoryCode[] = [
  'LYV',
  'LHG',
  'LYM',
  'LVL',
  'LYF',
  'JAZ',
  'JZS',
];

type PortCodeRow = {
  Customer_Number: string;
  Transport_Method: string;
  Port_Code: string;
};

@Injectable()
export class Cat9andcat12Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
    @Inject('LHG_ERP') private readonly LHG_ERP: Sequelize,
    @Inject('LVL_ERP') private readonly LVL_ERP: Sequelize,
    @Inject('LYM_ERP') private readonly LYM_ERP: Sequelize,
    @Inject('LYF_ERP') private readonly LYF_ERP: Sequelize,
    @Inject('JAZ_ERP') private readonly JAZ_ERP: Sequelize,
    @Inject('JZS_ERP') private readonly JZS_ERP: Sequelize,
  ) {}

  async getData(
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'No',
    sortOrder: string = 'asc',
  ) {
    const db = FACTORY_LIST.includes(factory)
      ? this.getDbByFactory(factory)
      : null;
    if (!db) {
      return await this.getAllDataFactory(
        dateFrom,
        dateTo,
        factory,
        ry,
        page,
        limit,
        sortField,
        sortOrder,
      );
    }

    return await this.getDataFactory(
      db,
      dateFrom,
      dateTo,
      factory,
      ry,
      page,
      limit,
      sortField,
      sortOrder,
    );
  }

  async getPortCode(
    sortField: string = 'CustomerNumber',
    sortOrder: string = 'asc',
  ): Promise<IDataPortCode[]> {
    const records: IDataPortCode[] = await this.EIP.query(
      `SELECT *
      FROM CMW_PortCode
      ORDER BY ${orderByClause(sortField, sortOrder)}
      `,
      { type: QueryTypes.SELECT },
    );
    return records;
  }

  private async getDataFactory(
    db: Sequelize,
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    try {
      const offset = (page - 1) * limit;

      const { where, where1, replacements } = buildCat9AndCat12Filter(
        dateFrom,
        dateTo,
        ry,
      );

      const query = buildCat9AndCat12PageQuery(where, where1);

      const countQuery = buildCat9AndCat12CountQuery(where, where1);

      const [dataResults, countResults] = await Promise.all([
        db.query(query, {
          replacements,
          type: QueryTypes.SELECT,
        }),
        db.query(countQuery, {
          replacements,
          type: QueryTypes.SELECT,
        }),
      ]);

      let data = dataResults as ShipmentRow[];
      data = await addPortsToRows(data, this.EIP, factory);

      sortByField(data, sortField, sortOrder);

      data = data.slice(offset, offset + limit);

      const total = (countResults[0] as { total: number })?.total || 0;

      const hasMore = offset + data.length < total;

      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  private async getAllDataFactory(
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    const offset = (page - 1) * limit;

    const { where, where1, replacements } = buildCat9AndCat12Filter(
      dateFrom,
      dateTo,
      ry,
    );
    const getQuery = (factoryName: string) =>
      buildCat9AndCat12AllPageQuery(factoryName, where, where1);
    const getCountQuery = (factoryName: string) =>
      buildCat9AndCat12AllCountQuery(factoryName, where, where1);

    // The order decides the order of the rows of the "ALL" page.
    const connects = ALL_PAGE_FACTORIES.map((factoryName) => ({
      conn: this.getDbByFactory(factoryName)!,
      factoryName,
    }));
    const [dataResults, countResults] = await Promise.all([
      Promise.all(
        connects.map(({ conn, factoryName }) => {
          const query = getQuery(factoryName);
          return conn.query<ShipmentRow>(query, {
            type: QueryTypes.SELECT,
            replacements,
          });
        }),
      ),
      Promise.all(
        connects.map(({ conn, factoryName }) => {
          const countQuery = getCountQuery(factoryName);
          return conn.query(countQuery, {
            type: QueryTypes.SELECT,
            replacements,
          });
        }),
      ),
    ]);

    const allData = dataResults.flat();

    let data: ShipmentRow[] = allData.map((item, index) => ({
      ...item,
      No: index + 1,
    }));

    // No factory is passed: the port rules follow the Factory_Name of each row (like the ALL Excel export and the CMS payload).
    data = await addPortsToRows(data, this.EIP);

    sortByField(data, sortField, sortOrder);

    data = data.slice(offset, offset + limit);

    const total = countResults.reduce((sum, result) => {
      return sum + ((result[0] as { total: number })?.total || 0);
    }, 0);
    const hasMore = offset + data.length < total;

    return { data, page, limit, total, hasMore };
  }

  async importExcelPortCode(
    file: Express.Multer.File,
    userid: string,
    factory: string,
  ) {
    try {
      const data = await readImportRows<PortCodeRow>(
        file,
        ['Customer_Number', 'Transport_Method', 'Port_Code'],
        (rowData) =>
          rowData?.Customer_Number ||
          rowData?.Transport_Method ||
          rowData?.Port_Code,
        { cellsAsStrings: true },
      );

      const { insertCount, updateCount } = await countUpserts(data, (item) =>
        this.upsertPortCode(item, userid, factory),
      );
      const records = await this.EIP.query(
        `SELECT *
          FROM CMW_PortCode`,
        { type: QueryTypes.SELECT },
      );
      const message = buildImportMessage(insertCount, updateCount, data.length);
      return { message, records };
    } catch (error: any) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }

  private async upsertPortCode(
    item: PortCodeRow,
    userid: string,
    factory: string,
  ): Promise<'insert' | 'update'> {
    const id = uuidv4();
    const records: { total: number }[] = await this.EIP.query(
      `SELECT COUNT(*) total
        FROM CMW_PortCode
        WHERE CustomerNumber = ? AND TransportMethod = ?`,
      {
        replacements: [item.Customer_Number, item.Transport_Method],
        type: QueryTypes.SELECT,
      },
    );

    if (records[0].total > 0) {
      await this.EIP.query(
        `UPDATE CMW_PortCode
        SET
              PortCode = ?,
              UpdatedAt = ?,
              UpdatedFactory = ?,
              UpdatedDate = GETDATE()
        WHERE CustomerNumber = ? AND TransportMethod = ?`,
        {
          replacements: [
            item.Port_Code,
            userid,
            factory,
            item.Customer_Number,
            item.Transport_Method,
          ],
          type: QueryTypes.UPDATE,
        },
      );
      return 'update';
    }

    await this.EIP.query(
      `INSERT INTO CMW_PortCode
              (
                    Id,
                    CustomerNumber,
                    PortCode,
                    TransportMethod,
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
                    GETDATE()
              )`,
      {
        replacements: [
          id,
          item.Customer_Number,
          item.Port_Code,
          item.Transport_Method,
          userid,
          factory,
        ],
        type: QueryTypes.INSERT,
      },
    );
    return 'insert';
  }

  async autoSentCMS(
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
    dockeyCMS: string,
  ) {
    try {
      if (factory.trim().toUpperCase() === 'ALL') {
        return this.autoSentCMSAllFactories(dateFrom, dateTo, ry, dockeyCMS);
      }

      return this.getCMSByFactory(factory, dateFrom, dateTo, ry, dockeyCMS);
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  private async autoSentCMSAllFactories(
    dateFrom: string,
    dateTo: string,
    ry: string,
    dockeyCMS: string,
  ) {
    const results = await Promise.all(
      FACTORY_LIST.map((factory) =>
        this.getCMSByFactory(factory, dateFrom, dateTo, ry, dockeyCMS),
      ),
    );

    return results.flat();
  }

  private async getCMSByFactory(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    ry: string,
    dockeyCMS: string,
  ) {
    const db = this.getDbByFactory(factory);
    if (!db) return [];

    return getCat9AndCat12CMSRows(
      db,
      this.EIP,
      factory,
      dateFrom,
      dateTo,
      ry,
      dockeyCMS,
    );
  }

  private getDbByFactory(factory: FactoryCode): Sequelize | null {
    const dbMap: Record<FactoryCode, Sequelize> = {
      LYV: this.LYV_ERP,
      LHG: this.LHG_ERP,
      LVL: this.LVL_ERP,
      LYM: this.LYM_ERP,
      LYF: this.LYF_ERP,
      JAZ: this.JAZ_ERP,
      JZS: this.JZS_ERP,
    };

    return getOwn(dbMap, factory) ?? null;
  }

  // logging
  async getLoggingCat9AndCat12(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    const offset = (page - 1) * limit;

    const [dataResults, countResults] = await Promise.all([
      this.EIP.query(
        `SELECT *
        FROM CMW_Category_9_And_12_Log`,
        { type: QueryTypes.SELECT },
      ),
      this.EIP.query(
        `SELECT COUNT(*) AS total
        FROM CMW_Category_9_And_12_Log`,
        {
          type: QueryTypes.SELECT,
        },
      ),
    ]);

    let data = dataResults;
    sortByField(data, sortField, sortOrder);

    data = data.slice(offset, offset + limit);

    const total = (countResults[0] as { total: number })?.total || 0;

    const hasMore = offset + data.length < total;

    return { data, page, limit, total, hasMore };
  }
}
