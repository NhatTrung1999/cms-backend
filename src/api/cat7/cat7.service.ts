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
dayjs().format();

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

  async getDataCat7(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Staff_ID',
    sortOrder: string = 'asc',
  ) {
    let db: Sequelize;
    switch (factory) {
      case 'LYV':
        db = this.LYV_HRIS;
        break;
      case 'LHG':
        db = this.LHG_HRIS;
        break;
      case 'LYM':
        db = this.LYM_HRIS;
        break;
      case 'LVL':
        db = this.LVL_HRIS;
        break;
      case 'JAZ':
        db = this.JAZ_HRIS;
        break;
      case 'JZS':
        db = this.JZS_HRIS;
        break;
      default:
        return await this.getAllDataFactory(
          dateFrom,
          dateTo,
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
      page,
      factory,
      limit,
      sortField,
      sortOrder,
    );
  }

  // private buildQuery(factory: string, dateFrom?: string, dateTo?: string) {
  //   const isLYM = factory === 'LYM';

  //   const baseWhere = !isLYM
  //     ? "WHERE 1=1 AND Work_Or_Not<>'2' AND u.Vehicle IS NOT NULL AND dwt.Working_Time > 0"
  //     : 'WHERE 1=1 AND u.Vehicle IS NOT NULL AND dwt.workhours > 0';

  //   const dateFilter = !isLYM
  //     ? 'AND CONVERT(DATE, dwt.Check_Day) BETWEEN ? AND ?'
  //     : 'AND CONVERT(DATE ,dwt.CDate) BETWEEN ? AND ?';

  //   const where = dateFrom && dateTo ? `${baseWhere} ${dateFilter}` : baseWhere;

  //   const workCol = !isLYM ? 'WORKING_TIME' : 'workhours';

  //   const table = !isLYM ? 'Data_Work_Time' : 'HR_Attendance';

  //   const join = !isLYM
  //     ? 'u.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS = dwt.Person_Serial_Key COLLATE Chinese_Taiwan_Stroke_CI_AS'
  //     : 'u.userId = dwt.UserNo';

  //   const query = `SELECT u.userId                       AS Staff_ID
  //                         ,CASE
  //                               WHEN u.Address_Live IS NULL THEN u.Bus_Route
  //                               ELSE u.Address_Live
  //                         END                            AS Residential_address
  //                         ,u.Vehicle                      AS Main_transportation_type
  //                         ,'API Calculation'              AS km
  //                         ,COUNT(${workCol})  AS Number_of_working_days
  //                         ,'API Calculation'              AS PKT_p_km
  //                   FROM   ${table}                 AS dwt
  //                         LEFT JOIN users                AS u
  //                               ON  ${join}
  //                   ${where}
  //                   GROUP BY
  //                           u.userId
  //                           ,u.Address_Live
  //                           ,u.Vehicle
  //                           ,u.Bus_Route`;

  //   // console.log(query);
  //   const countQuery = `SELECT COUNT(*) AS total
  //                       FROM   (
  //                                 SELECT u.userId                       AS Staff_ID
  //                                       ,CASE
  //                                             WHEN u.Address_Live IS NULL THEN u.Bus_Route
  //                                             ELSE u.Address_Live
  //                                       END                            AS Residential_address
  //                                       ,u.Vehicle                      AS Main_transportation_type
  //                                       ,'API Calculation'              AS km
  //                                       ,COUNT(${workCol})  AS Number_of_working_days
  //                                       ,'API Calculation'              AS PKT_p_km
  //                                 FROM   ${table}                 AS dwt
  //                                       LEFT JOIN users                AS u
  //                                             ON  ${join}
  //                                 ${where}
  //                                 GROUP BY
  //                                         u.userId
  //                                         ,u.Address_Live
  //                                         ,u.Vehicle
  //                                         ,u.Bus_Route
  //                       ) AS Sub`;
  //   return { query, countQuery };
  // }

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

    const { query, countQuery } = buildQuery(factory, dateFrom, dateTo);

    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

    const [dataResults, countResults] = await Promise.all([
      db.query(query, { replacements, type: QueryTypes.SELECT }),
      db.query(countQuery, {
        type: QueryTypes.SELECT,
        replacements,
      }),
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

      const connects: { factoryName: string; conn: Sequelize }[] = [
        { factoryName: 'LYV', conn: this.LYV_HRIS },
        { factoryName: 'LHG', conn: this.LHG_HRIS },
        { factoryName: 'LVL', conn: this.LVL_HRIS },
        { factoryName: 'LYM', conn: this.LYM_HRIS },
        { factoryName: 'JAZ', conn: this.JAZ_HRIS },
        { factoryName: 'JZS', conn: this.JZS_HRIS },
      ];

      const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

      const [dataResults, countResults] = await Promise.all([
        Promise.all(
          connects.map(({ factoryName, conn }) => {
            const { query } = buildQuery(factoryName, dateFrom, dateTo);
            return conn.query(query, { type: QueryTypes.SELECT, replacements });
          }),
        ),
        Promise.all(
          connects.map(({ factoryName, conn }) => {
            const { countQuery } = buildQuery(factoryName, dateFrom, dateTo);
            return conn.query(countQuery, {
              type: QueryTypes.SELECT,
              replacements,
            });
          }),
        ),
      ]);

      let data = dataResults.flat();

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

      const total = countResults.reduce((sum, result) => {
        return sum + ((result[0] as { total: number })?.total || 0);
      }, 0);
      const hasMore = offset + data.length < total;

      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  // custom export
  async getCustomExport(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'ID',
    sortOrder: string = 'asc',
  ) {
    // console.log(dateFrom, dateTo, factory, page, limit, sortField, sortOrder);
    let db: Sequelize;
    switch (factory) {
      case 'LYV':
        db = this.LYV_HRIS;
        break;
      case 'LHG':
        db = this.LHG_HRIS;
        break;
      case 'LYM':
        db = this.LYM_HRIS;
        break;
      case 'LVL':
        db = this.LVL_HRIS;
        break;
      case 'JAZ':
        db = this.JAZ_HRIS;
        break;
      case 'JZS':
        db = this.JZS_HRIS;
        break;
      default:
        return await this.getAllDataCustomExport(
          dateFrom,
          dateTo,
          page,
          limit,
          sortField,
          sortOrder,
        );
    }
    return await this.getDataCustomExport(
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
    // console.log(sortField, sortOrder, page, limit);
    const offset = (page - 1) * limit;

    const { query, countQuery } = buildQueryCustomExport(
      dateFrom,
      dateTo,
      factory,
    );

    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

    const [dataResults, countResults] = await Promise.all([
      db.query(query, { replacements, type: QueryTypes.SELECT }),
      db.query(countQuery, {
        replacements,
        type: QueryTypes.SELECT,
      }),
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

      const connects: { factoryName: string; conn: Sequelize }[] = [
        { factoryName: 'LYV', conn: this.LYV_HRIS },
        { factoryName: 'LHG', conn: this.LHG_HRIS },
        { factoryName: 'LVL', conn: this.LVL_HRIS },
        { factoryName: 'LYM', conn: this.LYM_HRIS },
        { factoryName: 'JAZ', conn: this.JAZ_HRIS },
        { factoryName: 'JZS', conn: this.JZS_HRIS },
      ];

      const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

      const [dataResults, countResults] = await Promise.all([
        Promise.all(
          connects.map(({ factoryName, conn }) => {
            const { query } = buildQueryCustomExport(
              dateFrom,
              dateTo,
              factoryName,
            );
            return conn.query(query, { type: QueryTypes.SELECT, replacements });
          }),
        ),
        Promise.all(
          connects.map(({ factoryName, conn }) => {
            const { countQuery } = buildQueryCustomExport(
              dateFrom,
              dateTo,
              factoryName,
            );
            return conn.query(countQuery, {
              type: QueryTypes.SELECT,
              replacements,
            });
          }),
        ),
      ]);

      const allData = dataResults.flat();

      let data = allData.map((item, index) => ({ ...item, No: index + 1 }));

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

      const total = countResults.reduce((sum, result) => {
        return sum + ((result[0] as { total: number })?.total || 0);
      }, 0);
      const hasMore = offset + data.length < total;

      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  // auto send cms

  async autoSentCMS(dateFrom: string, dateTo: string) {
    try {
      const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];
      const [dataLYV, dataLHG, dataLVL, dataLYM, dataJAZ, dataJZS] = await Promise.all([
        await this.LYV_HRIS.query(
          await buildQueryAutoSentCmsLYV(dateFrom, dateTo, this.EIP),
          {
            type: QueryTypes.SELECT,
            replacements,
          },
        ),
        await this.LHG_HRIS.query(
          await buildQueryAutoSentCmsLHG(dateFrom, dateTo, this.EIP),
          {
            type: QueryTypes.SELECT,
            replacements,
          },
        ),
        await this.LVL_HRIS.query(
          await buildQueryAutoSentCmsLVL(dateFrom, dateTo, this.EIP),
          {
            type: QueryTypes.SELECT,
            replacements,
          },
        ),
        await this.LYM_HRIS.query(
          await buildQueryAutoSentCmsLYM(dateFrom, dateTo, this.EIP),
          {
            type: QueryTypes.SELECT,
            replacements,
          },
        ),
        await this.JAZ_HRIS.query(
          await buildQueryAutoSentCmsJAZ(dateFrom, dateTo, this.EIP),
          {
            type: QueryTypes.SELECT,
            replacements,
          },
        ),
        await this.JZS_HRIS.query(
          await buildQueryAutoSentCmsJZS(dateFrom, dateTo, this.EIP),
          {
            type: QueryTypes.SELECT,
            replacements,
          },
        ),
      ]);

      // console.log(123);

      const data = [
        // ...dataLYV,
        // ...dataLHG,
        // ...dataLVL,
        ...dataLYM,
        // ...dataJAZ,
        // ...dataJZS,
      ].flat();
      // const data = [...dataLYV].flat();
      // console.log(data);

      const formatData = data.map((item: any) => {
        const staffId = item.Staff_ID;
        const Residential_address = item.Residential_address;
        const Main_transportation_type = item.Main_transportation_type;
        const Number_of_working_days = item.Number_of_working_days;
        const Factory_address = item.Factory_address;
        const FactoryName = item.Factory_Name;
        const DepartmentName = item.Department_Name;

        return {
          System: 'CMS Web', // Default
          Corporation: 'LAI YIH', // Default
          Factory: FactoryName,
          Department: DepartmentName,
          DocKey: staffId,
          SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
          EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
          ActivityType: '3.3 員工通勤', // Default
          DataType: '2', // Default
          DocType: '員工通勤', // Default
          DocDate: dayjs().format('YYYY/MM/DD'),
          DocDate2: dayjs().format('YYYY/MM/DD'),
          UndDocNo: staffId,
          TransType: Main_transportation_type,
          Departure: Factory_address,
          Destination: Residential_address,
          Attendance: Number_of_working_days.toString(),
          Memo: '',
          CreateDateTime: dayjs().format('YYYY/MM/DD HH:mm:ss'),
          Creator: '',
        };
      });

      // console.log(formatData);
      return formatData;
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }
}

