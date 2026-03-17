import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import {
  buildQuery,
  buildQueryAutoSentCMS,
  buildQueryAutoSentCmsJAZ,
  buildQueryAutoSentCmsJZS,
  buildQueryAutoSentCmsLHG,
  buildQueryAutoSentCmsLVL,
  buildQueryAutoSentCmsLYM,
  buildQueryAutoSentCmsLYV,
  buildQueryCustomExport,
} from 'src/helper/cat7.helper';
import dayjs from 'dayjs';
import { ACTIVITY_TYPES, ActivityType } from 'src/types/cat7';
import { FACTORY_LIST, FactoryCode } from 'src/helper/factory.helper';
dayjs().format();

type BuildQueryFn = (
  dateFrom: string,
  dateTo: string,
  eip: Sequelize,
) => Promise<string>;

@Injectable()
export class Cat7Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_HRIS') private readonly LYV_HRIS: Sequelize,
    @Inject('LHG_HRIS') private readonly LHG_HRIS: Sequelize,
    @Inject('LVL_HRIS') private readonly LVL_HRIS: Sequelize,
    @Inject('LYM_HRIS') private readonly LYM_HRIS: Sequelize,
    @Inject('JAZ_HRIS') private readonly JAZ_HRIS: Sequelize,
    @Inject('JZS_HRIS') private readonly JZS_HRIS: Sequelize,
  ) {}

  private get factoryDbMap(): Record<string, Sequelize> {
    return {
      LYV: this.LYV_HRIS,
      LHG: this.LHG_HRIS,
      LYM: this.LYM_HRIS,
      LVL: this.LVL_HRIS,
      JAZ: this.JAZ_HRIS,
      JZS: this.JZS_HRIS,
    };
  }

  private readonly buildQueryMap: Record<FactoryCode, BuildQueryFn> = {
    LYV: buildQueryAutoSentCmsLYV,
    LHG: buildQueryAutoSentCmsLHG,
    LVL: buildQueryAutoSentCmsLVL,
    LYM: buildQueryAutoSentCmsLYM,
    JAZ: buildQueryAutoSentCmsJAZ,
    JZS: buildQueryAutoSentCmsJZS,
  };

  private readonly sortFieldMap: Record<string, string> = {
    Staff_ID: 'u.userId',
    Residential_Address: 'u.Address_Live',
    Main_Transportation_Type: 'u.Vehicle',
    Number_of_Working_Days: 'e.Number_of_Working_Days',
  };

  private buildReplacements(
    dateFrom: string,
    dateTo: string,
  ): {
    dateToExclusive: string;
    hasDate: boolean;
  } {
    return {
      dateToExclusive: dateTo
        ? dayjs(dateTo).add(1, 'day').format('YYYY-MM-DD')
        : '',
      hasDate: !!(dateFrom && dateTo),
    };
  }

  async getDataCat7(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Staff_ID',
    sortOrder: string = 'asc',
  ) {
    const db = this.factoryDbMap[factory];

    if (!db) {
      return this.getAllDataFactory(
        dateFrom,
        dateTo,
        page,
        limit,
        sortField,
        sortOrder,
      );
    }

    return this.getDataFactory(
      db,
      dateFrom,
      dateTo,
      page,
      factory,
      limit,
      sortField,
      sortOrder,
    );
  }

  private async getDataFactory(
    db: Sequelize,
    dateFrom: string,
    dateTo: string,
    page: number,
    factory: string,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    const offset = (page - 1) * limit;
    const { dateToExclusive, hasDate } = this.buildReplacements(
      dateFrom,
      dateTo,
    );

    const { query, countQuery } = buildQuery(
      factory,
      dateFrom,
      dateToExclusive,
    );

    const safeSortField = this.sortFieldMap[sortField] ?? 'u.userId';
    const safeSortOrder = sortOrder === 'asc' ? 'ASC' : 'DESC';
    const sortedQuery = query.replace(
      'ORDER BY u.userId',
      `ORDER BY ${safeSortField} ${safeSortOrder}`,
    );

    const dataReplacements = hasDate
      ? [dateFrom, dateToExclusive, offset, limit]
      : [offset, limit];
    const countReplacements = hasDate ? [dateFrom, dateToExclusive] : [];

    const [dataResults, countResults] = await Promise.all([
      db.query(sortedQuery, {
        replacements: dataReplacements,
        type: QueryTypes.SELECT,
      }),
      db.query(countQuery, {
        replacements: countReplacements,
        type: QueryTypes.SELECT,
      }),
    ]);

    const total = (countResults[0] as { total: number })?.total || 0;
    const hasMore = offset + dataResults.length < total;

    return { data: dataResults, page, limit, total, hasMore };
  }

  private async getAllDataFactory(
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    try {
      const offset = (page - 1) * limit;
      const { dateToExclusive, hasDate } = this.buildReplacements(
        dateFrom,
        dateTo,
      );

      const factories = Object.entries(this.factoryDbMap);

      const countReplacements = hasDate ? [dateFrom, dateToExclusive] : [];

      const results = await Promise.all(
        factories.map(([factoryName, conn]) => {
          const { query, countQuery } = buildQuery(
            factoryName,
            dateFrom,
            dateToExclusive,
          );

          const safeSortField = this.sortFieldMap[sortField] ?? 'u.userId';
          const safeSortOrder = sortOrder === 'asc' ? 'ASC' : 'DESC';
          const sortedQuery = query.replace(
            'ORDER BY u.userId',
            `ORDER BY ${safeSortField} ${safeSortOrder}`,
          );

          const dataReplacements = hasDate
            ? [dateFrom, dateToExclusive, 0, 999999]
            : [0, 999999];

          return Promise.all([
            conn.query(sortedQuery, {
              replacements: dataReplacements,
              type: QueryTypes.SELECT,
            }),
            conn.query(countQuery, {
              replacements: countReplacements,
              type: QueryTypes.SELECT,
            }),
          ]);
        }),
      );

      const allData = results.flatMap(([data]) => data);
      const allCount = results.reduce((sum, [, count]) => {
        return sum + ((count[0] as { total: number })?.total || 0);
      }, 0);

      allData.sort((a, b) => {
        const aVal = a[sortField];
        const bVal = b[sortField];
        if (sortOrder === 'asc') return aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
        else return aVal > bVal ? -1 : aVal < bVal ? 1 : 0;
      });

      const data = allData.slice(offset, offset + limit);
      const hasMore = offset + data.length < allCount;

      return { data, page, limit, total: allCount, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  // custom export
  // async getCustomExport(
  //   dateFrom: string,
  //   dateTo: string,
  //   factory: string,
  //   page: number = 1,
  //   limit: number = 20,
  //   sortField: string = 'ID',
  //   sortOrder: string = 'asc',
  // ) {
  //   let db: Sequelize;
  //   switch (factory) {
  //     case 'LYV':
  //       db = this.LYV_HRIS;
  //       break;
  //     case 'LHG':
  //       db = this.LHG_HRIS;
  //       break;
  //     case 'LYM':
  //       db = this.LYM_HRIS;
  //       break;
  //     case 'LVL':
  //       db = this.LVL_HRIS;
  //       break;
  //     case 'JAZ':
  //       db = this.JAZ_HRIS;
  //       break;
  //     case 'JZS':
  //       db = this.JZS_HRIS;
  //       break;
  //     default:
  //       return await this.getAllDataCustomExport(
  //         dateFrom,
  //         dateTo,
  //         page,
  //         limit,
  //         sortField,
  //         sortOrder,
  //       );
  //   }
  //   return await this.getDataCustomExport(
  //     db,
  //     dateFrom,
  //     dateTo,
  //     page,
  //     factory,
  //     limit,
  //     sortField,
  //     sortOrder,
  //   );
  // }

  // private async getDataCustomExport(
  //   db: Sequelize,
  //   dateFrom: string,
  //   dateTo: string,
  //   page: number,
  //   factory: string,
  //   limit: number,
  //   sortField: string,
  //   sortOrder: string,
  // ) {
  //   // console.log(sortField, sortOrder, page, limit);
  //   const offset = (page - 1) * limit;

  //   const { query, countQuery } = buildQueryCustomExport(
  //     dateFrom,
  //     dateTo,
  //     factory,
  //   );

  //   const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

  //   const [dataResults, countResults] = await Promise.all([
  //     db.query(query, { replacements, type: QueryTypes.SELECT }),
  //     db.query(countQuery, {
  //       replacements,
  //       type: QueryTypes.SELECT,
  //     }),
  //   ]);
  //   let data = dataResults;
  //   data.sort((a, b) => {
  //     const aValue = a[sortField];
  //     const bValue = b[sortField];
  //     if (sortOrder === 'asc') {
  //       return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
  //     } else {
  //       return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
  //     }
  //   });

  //   data = data.slice(offset, offset + limit);

  //   const total = (countResults[0] as { total: number })?.total || 0;

  //   const hasMore = offset + data.length < total;

  //   return { data, page, limit, total, hasMore };
  // }

  // private async getAllDataCustomExport(
  //   dateFrom: string,
  //   dateTo: string,
  //   page: number,
  //   limit: number,
  //   sortField: string,
  //   sortOrder: string,
  // ) {
  //   try {
  //     const offset = (page - 1) * limit;

  //     const connects: { factoryName: string; conn: Sequelize }[] = [
  //       { factoryName: 'LYV', conn: this.LYV_HRIS },
  //       { factoryName: 'LHG', conn: this.LHG_HRIS },
  //       { factoryName: 'LVL', conn: this.LVL_HRIS },
  //       { factoryName: 'LYM', conn: this.LYM_HRIS },
  //       { factoryName: 'JAZ', conn: this.JAZ_HRIS },
  //       { factoryName: 'JZS', conn: this.JZS_HRIS },
  //     ];

  //     const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

  //     const [dataResults, countResults] = await Promise.all([
  //       Promise.all(
  //         connects.map(({ factoryName, conn }) => {
  //           const { query } = buildQueryCustomExport(
  //             dateFrom,
  //             dateTo,
  //             factoryName,
  //           );
  //           return conn.query(query, { type: QueryTypes.SELECT, replacements });
  //         }),
  //       ),
  //       Promise.all(
  //         connects.map(({ factoryName, conn }) => {
  //           const { countQuery } = buildQueryCustomExport(
  //             dateFrom,
  //             dateTo,
  //             factoryName,
  //           );
  //           return conn.query(countQuery, {
  //             type: QueryTypes.SELECT,
  //             replacements,
  //           });
  //         }),
  //       ),
  //     ]);

  //     const allData = dataResults.flat();

  //     let data = allData.map((item, index) => ({ ...item, No: index + 1 }));

  //     data.sort((a, b) => {
  //       const aValue = a[sortField];
  //       const bValue = b[sortField];
  //       if (sortOrder === 'asc') {
  //         return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
  //       } else {
  //         return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
  //       }
  //     });

  //     data = data.slice(offset, offset + limit);

  //     const total = countResults.reduce((sum, result) => {
  //       return sum + ((result[0] as { total: number })?.total || 0);
  //     }, 0);
  //     const hasMore = offset + data.length < total;

  //     return { data, page, limit, total, hasMore };
  //   } catch (error: any) {
  //     throw new InternalServerErrorException(error);
  //   }
  // }
  private readonly sortFieldMapCustomExport: Record<string, string> = {
    No: 'a.userId',
    ID: 'a.userId',
    Department: 'c.Department_Name',
    Full_Name: 'a.fullName',
    Current_Address: 'a.Address_Live',
    Transportation_Mode: 'a.Vehicle',
    Bus_Route: 'a.Bus_Route',
    Pickup_Point: 'a.PickupDropoffStation',
    Number_of_Working_Days: 'e.Number_of_Working_Days',
  };
  async getCustomExport(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'ID',
    sortOrder: string = 'asc',
  ) {
    const db = this.factoryDbMap[factory];

    if (!db) {
      return this.getAllDataCustomExport(
        dateFrom,
        dateTo,
        page,
        limit,
        sortField,
        sortOrder,
      );
    }

    return this.getDataCustomExport(
      db,
      dateFrom,
      dateTo,
      page,
      factory,
      limit,
      sortField,
      sortOrder,
    );
  }
  private async getDataCustomExport(
    db: Sequelize,
    dateFrom: string,
    dateTo: string,
    page: number,
    factory: string,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    const offset = (page - 1) * limit;
    const { dateToExclusive, hasDate } = this.buildReplacements(
      dateFrom,
      dateTo,
    );

    const { query, countQuery } = buildQueryCustomExport(
      dateFrom,
      dateToExclusive,
      factory,
    );

    const safeSortField =
      this.sortFieldMapCustomExport[sortField] ?? 'a.userId';
    const safeSortOrder = sortOrder === 'asc' ? 'ASC' : 'DESC';
    const sortedQuery = query.replace(
      'ORDER BY a.userId',
      `ORDER BY ${safeSortField} ${safeSortOrder}`,
    );

    // ✅ Sort + pagination trong SQL
    const dataReplacements = hasDate
      ? [dateFrom, dateToExclusive, offset, limit]
      : [offset, limit];
    const countReplacements = hasDate ? [dateFrom, dateToExclusive] : [];

    const [dataResults, countResults] = await Promise.all([
      db.query(sortedQuery, {
        replacements: dataReplacements,
        type: QueryTypes.SELECT,
      }),
      db.query(countQuery, {
        replacements: countReplacements,
        type: QueryTypes.SELECT,
      }),
    ]);

    const total = (countResults[0] as { total: number })?.total || 0;
    const hasMore = offset + dataResults.length < total;

    return { data: dataResults, page, limit, total, hasMore };
  }

  private async getAllDataCustomExport(
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    try {
      const offset = (page - 1) * limit;
      const { dateToExclusive, hasDate } = this.buildReplacements(
        dateFrom,
        dateTo,
      );

      const safeSortField =
        this.sortFieldMapCustomExport[sortField] ?? 'a.userId';
      const safeSortOrder = sortOrder === 'asc' ? 'ASC' : 'DESC';
      const countReplacements = hasDate ? [dateFrom, dateToExclusive] : [];

      // ✅ Gộp data + count vào 1 Promise.all duy nhất
      const results = await Promise.all(
        Object.entries(this.factoryDbMap).map(([factoryName, conn]) => {
          const { query, countQuery } = buildQueryCustomExport(
            dateFrom,
            dateToExclusive,
            factoryName,
          );

          const sortedQuery = query.replace(
            'ORDER BY a.userId',
            `ORDER BY ${safeSortField} ${safeSortOrder}`,
          );
          const dataReplacements = hasDate
            ? [dateFrom, dateToExclusive, 0, 999999]
            : [0, 999999];

          return Promise.all([
            conn.query(sortedQuery, {
              replacements: dataReplacements,
              type: QueryTypes.SELECT,
            }),
            conn.query(countQuery, {
              replacements: countReplacements,
              type: QueryTypes.SELECT,
            }),
          ]);
        }),
      );

      // ✅ Merge + đánh lại số No sau khi gộp tất cả factory
      const allData = results
        .flatMap(([data]) => data)
        .map((item, index) => ({ ...item, No: index + 1 }));

      allData.sort((a, b) => {
        const aVal = a[sortField];
        const bVal = b[sortField];
        if (sortOrder === 'asc') return aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
        else return aVal > bVal ? -1 : aVal < bVal ? 1 : 0;
      });

      const total = results.reduce(
        (sum, [, count]) => sum + ((count[0] as { total: number })?.total || 0),
        0,
      );
      const data = allData.slice(offset, offset + limit);
      const hasMore = offset + data.length < total;

      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  // auto send cms

  async autoSentCMS(dateFrom: string, dateTo: string, factory: string) {
    try {
      if (factory.trim().toUpperCase() === 'ALL') {
        return this.autoSentCMSAllFactories(dateFrom, dateTo);
      }

      return this.getCMSByFactory(factory as FactoryCode, dateFrom, dateTo);
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  private async autoSentCMSAllFactories(dateFrom: string, dateTo: string) {
    const results = await Promise.all(
      FACTORY_LIST.map((factory) =>
        this.getCMSByFactory(factory, dateFrom, dateTo),
      ),
    );

    return results.flat();
  }

  private async getCMSByFactory(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
  ) {
    const db = this.getDbByFactory(factory);
    if (!db) return [];

    const buildQuery = this.buildQueryMap[factory];
    if (!buildQuery) return [];

    const query = await buildQuery(dateFrom, dateTo, this.EIP);

    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

    const data = await db.query<any>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    return data.flatMap((item) => this.mapToCMSFormat(item, dateFrom, dateTo));
  }

  private getDbByFactory(factory: FactoryCode): Sequelize | null {
    const dbMap: Record<FactoryCode, Sequelize> = {
      LYV: this.LYV_HRIS,
      LHG: this.LHG_HRIS,
      LVL: this.LVL_HRIS,
      LYM: this.LYM_HRIS,
      JAZ: this.JAZ_HRIS,
      JZS: this.JZS_HRIS,
    };

    return dbMap[factory] ?? null;
  }

  private mapToCMSFormat(item: any, dateFrom: string, dateTo: string) {
    const staffId = item.Staff_ID ?? '';
    const Main_transportation_type = item.Main_transportation_type ?? '';
    const Number_of_working_days = item.Number_of_working_days ?? '';
    const Factory_address = item.Factory_address ?? '';
    const Residential_address = item.Residential_address ?? Factory_address;
    const FactoryName = item.Factory_Name ?? '';
    const DepartmentName = item.Department_Name ?? '';

    // if (DepartmentName.length > 20) {
    //   return [];
    // }

    // console.log(DepartmentName.length);
    let dockey = '';
    let transType = '';

    switch (Main_transportation_type.trim().toLowerCase()) {
      case 'Walking'.trim().toLowerCase(): //步行
        dockey = '3.3.1';
        transType = '步行';
        break;
      case 'Bicycle'.trim().toLowerCase(): //自行車
        dockey = '3.3.2';
        transType = '自行車';
        break;
      case 'Electric motorcycle'.trim().toLowerCase(): //電動機車
        dockey = '3.3.3';
        transType = '電動機車';
        break;
      case 'Motorcycle'.trim().toLowerCase(): //機車
        dockey = '3.3.4';
        transType = '機車';
        break;
      case 'Electric car'.trim().toLowerCase(): //電動汽車
        dockey = '3.3.5';
        transType = '電動汽車';
        break;
      case 'Car'.trim().toLowerCase(): //汽車
        dockey = '3.3.6';
        transType = '汽車';
        break;
      case 'Bus'.trim().toLowerCase(): //公車/客運
        dockey = '3.3.7';
        transType = '公車/客運';
        break;
      case 'Company shuttle bus'.trim().toLowerCase(): //員工接駁車
        dockey = '3.3.8';
        transType = '員工接駁車';
        break;
      case 'Subway'.trim().toLowerCase(): //地鐵
        dockey = '3.3.9';
        transType = '地鐵';
        break;
      default:
        break;
    }

    return ACTIVITY_TYPES.map((activityType: ActivityType) => ({
      // System: 'CMS Web', // Default
      // Corporation: 'LAI YIH', // Default
      // Factory: FactoryName,
      // Department: DepartmentName,
      // DocKey: staffId,
      // SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      // EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      // ActivityType: activityType, // Default
      // DataType: '2', // Default
      // DocType: '員工通勤', // Default
      // DocDate: dayjs().format('YYYY/MM/DD'),
      // DocDate2: dayjs().format('YYYY/MM/DD'),
      // UndDocNo: staffId,
      // TransType: Main_transportation_type,
      // Departure: Factory_address ? Factory_address : '',
      // Destination: Residential_address ? Residential_address : '',
      // Attendance: Number_of_working_days.toString(),
      // Memo: '',
      // CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      // Creator: '',
      System: 'CMS Web',
      Corporation: 'LAI YIH',
      Factory: FactoryName,
      Department: DepartmentName.slice(0, 20),
      DocKey: dockey,
      SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
      EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
      ActivityType: activityType,
      DataType: '2',
      DocType: 'CMS Web',
      DocDate: dayjs('2025/12/31').format('YYYY/MM/DD'),
      DocDate2: dayjs('2025/12/31').format('YYYY/MM/DD'),
      UndDocNo: staffId,
      TransType: transType,
      Departure: Residential_address,
      Destination: Factory_address,
      Attendance: Number_of_working_days.toString(),
      Memo: '',
      CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
      Creator: '',
    }));
  }
  //logging
  async getLoggingCat7(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'System',
    sortOrder: string = 'asc',
  ) {
    const offset = (page - 1) * limit;
    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

    const [dataResults, countResults] = await Promise.all([
      this.EIP.query(
        `SELECT *
        FROM CMW_Category_7_Log`,
        { type: QueryTypes.SELECT },
      ),
      this.EIP.query(
        `SELECT COUNT(*) AS total
        FROM CMW_Category_7_Log`,
        {
          type: QueryTypes.SELECT,
        },
      ),
    ]);

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
    // const records: any[] = await this.EIP.query(
    //   `SELECT *
    //     FROM CMW_Category_7_Log
    //     ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
    //         `,
    //   { type: QueryTypes.SELECT },
    // );
    // return records;
  }
}
