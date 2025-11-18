import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { buildQuery } from 'src/helper/cat7.helper';

@Injectable()
export class Cat7Service {
  constructor(
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

      const connects: { facotryName: string; conn: Sequelize }[] = [
        { facotryName: 'LYV', conn: this.LYV_HRIS },
        { facotryName: 'LHG', conn: this.LHG_HRIS },
        { facotryName: 'LVL', conn: this.LVL_HRIS },
        { facotryName: 'LYM', conn: this.LYM_HRIS },
        { facotryName: 'JAZ', conn: this.JAZ_HRIS },
        { facotryName: 'JZS', conn: this.JZS_HRIS },
      ];

      const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

      const [dataResults, countResults] = await Promise.all([
        Promise.all(
          connects.map(({ facotryName, conn }) => {
            const { query } = buildQuery(facotryName, dateFrom, dateTo);
            return conn.query(query, { type: QueryTypes.SELECT, replacements });
          }),
        ),
        Promise.all(
          connects.map(({ facotryName, conn }) => {
            const { countQuery } = buildQuery(
              facotryName,
              dateFrom,
              dateTo,
            );
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
}
