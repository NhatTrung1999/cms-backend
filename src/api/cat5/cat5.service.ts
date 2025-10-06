import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class Cat5Service {
  constructor(
    @Inject('LYV_WMS') private readonly LYV_WMS: Sequelize,
    @Inject('LHG_WMS') private readonly LHG_WMS: Sequelize,
    @Inject('LYM_WMS') private readonly LYM_WMS: Sequelize,
    @Inject('LVL_WMS') private readonly LVL_WMS: Sequelize,
    @Inject('JAZ_WMS') private readonly JAZ_WMS: Sequelize,
    @Inject('JZS_WMS') private readonly JZS_WMS: Sequelize,
  ) {}

  async getDataWMS(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'Waste_disposal_date',
    sortOrder: string = 'asc',
  ) {
    let db: Sequelize;
    switch (factory) {
      case 'LYV':
        db = this.LYV_WMS;
        break;
      case 'LHG':
        db = this.LHG_WMS;
        break;
      case 'LYM':
        db = this.LYM_WMS;
        break;
      case 'LVL':
        db = this.LVL_WMS;
        break;
      case 'JAZ':
        db = this.JAZ_WMS;
        break;
      case 'JZS':
        db = this.JZS_WMS;
        break;
      default:
        return 'data All';
    }

    return await this.getDataFactory(
      db,
      dateFrom,
      dateTo,
      page,
      limit,
      sortField,
      sortOrder,
    );
    // try {
    //   const offset = (page - 1) * limit;

    //   let where = 'WHERE 1=1';
    //   const replacements: any[] = [];

    // if (dateTo) {
    //   where += ` AND CONVERT(VARCHAR, dwo.WASTE_DATE, 23) = ?`;
    //   replacements.push(dateTo);
    // }

    // const query = `SELECT dwo.WASTE_DATE AS Waste_disposal_date
    //                       ,dtv.TREATMENT_VENDOR_NAME          AS Vender_Name
    //                       ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' AS Waste_collection_address
    //                       ,CAST('0' AS INT)                   AS Transportation_Distance_km
    //                       ,CASE
    //                             WHEN dwo.HAZARDOUS<>'N/A' THEN 'hazardous waste'
    //                             WHEN dwo.NON_HAZARDOUS<>'N/A' THEN 'Non-hazardous waste'
    //                             ELSE NULL
    //                       END                                AS The_type_of_waste
    //                       ,CASE
    //                             WHEN dwo.HAZARDOUS<>'N/A' THEN dwo.HAZARDOUS
    //                             WHEN dwo.NON_HAZARDOUS<>'N/A' THEN dwo.NON_HAZARDOUS
    //                             ELSE NULL
    //                       END                                AS Waste_type
    //                       ,dtm.TREATMENT_METHOD_ENGLISH_NAME  AS Waste_Treatment_method
    //                       ,dwo.QUANTITY                       AS Weight_of_waste_treated_Unit_kg
    //                       ,CAST('0' AS INT)                   AS TKT_Ton_km
    //                   FROM   dbo.DATA_WASTE_OUTPUT_CUSTOMER dwo
    //                         LEFT JOIN dbo.DATA_TREATMENT_VENDOR dtv
    //                               ON  dtv.TREATMENT_VENDOR_ID = dwo.TREATMENT_SUPPLIER
    //                         LEFT JOIN dbo.DATA_TREATMENT_METHOD dtm
    //                               ON  dtm.TREATMENT_METHOD_ID = dwo.TREATMENT_METHOD_ID
    //                         LEFT JOIN dbo.DATA_EMERET_WASTE_CATEGORY dewc
    //                               ON  dewc.EMERET_WASTE_ID = dwo.EMERET_WASTE_ID
    //                         LEFT JOIN dbo.DATA_WASTE_CATEGORY dwc
    //                               ON  dwc.WASTE_CODE = dwo.WASTE_CODE
    //                         LEFT JOIN dbo.DATA_TREATMENT_DISTANCE td
    //                               ON  td.CODE = dwo.LOCATION_CODE
    //                      ${where}`;
    //   // ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
    //   // OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY`;

    // const countQuery = `SELECT COUNT(*) AS total
    //                   FROM dbo.DATA_WASTE_OUTPUT_CUSTOMER dwo
    //                   LEFT JOIN dbo.DATA_TREATMENT_VENDOR dtv
    //                         ON dtv.TREATMENT_VENDOR_ID = dwo.TREATMENT_SUPPLIER
    //                   LEFT JOIN dbo.DATA_TREATMENT_METHOD dtm
    //                         ON dtm.TREATMENT_METHOD_ID = dwo.TREATMENT_METHOD_ID
    //                   LEFT JOIN dbo.DATA_EMERET_WASTE_CATEGORY dewc
    //                         ON dewc.EMERET_WASTE_ID = dwo.EMERET_WASTE_ID
    //                   LEFT JOIN dbo.DATA_WASTE_CATEGORY dwc
    //                         ON dwc.WASTE_CODE = dwo.WASTE_CODE
    //                   LEFT JOIN dbo.DATA_TREATMENT_DISTANCE td
    //                         ON td.CODE = dwo.LOCATION_CODE
    //                   ${where}`;
    //   const connects = [
    //     this.LYV_WMS,
    //     this.LHG_WMS,
    //     this.LYM_WMS,
    //     this.LVL_WMS,
    //     this.JAZ_WMS,
    //     this.JZS_WMS,
    //   ];

    //   const [dataResults, countResults] = await Promise.all([
    //     Promise.all(
    //       connects.map((conn) => {
    //         return conn.query(query, { type: QueryTypes.SELECT, replacements });
    //       }),
    //     ),
    //     Promise.all(
    //       connects.map((conn) => {
    //         return conn.query(countQuery, {
    //           type: QueryTypes.SELECT,
    //           replacements,
    //         });
    //       }),
    //     ),
    //   ]);

    //   let data = dataResults.flat();

    // data.sort((a, b) => {
    //   const aValue = a[sortField];
    //   const bValue = b[sortField];
    //   if (sortOrder === 'asc') {
    //     return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
    //   } else {
    //     return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
    //   }
    // });

    // data = data.slice(offset, offset + limit);

    // const total = countResults.reduce((sum, result) => {
    //   return sum + ((result[0] as { total: number })?.total || 0);
    // }, 0);
    // const hasMore = offset + data.length < total;

    // return { data, page, limit, total, hasMore };
    // } catch (error: any) {
    //   throw new InternalServerErrorException(error);
    // }
  }

  private async getDataFactory(
    db: Sequelize,
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    // console.log(dateFrom, dateTo, factory, page, limit, sortField, sortOrder);
    const offset = (page - 1) * limit;
    let where = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateTo && dateFrom) {
      where += ` AND dwo.WASTE_DATE BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    const query = `SELECT dwo.WASTE_DATE AS Waste_disposal_date
                        ,dtv.TREATMENT_VENDOR_NAME          AS Vender_Name
                        ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' AS Waste_collection_address
                        ,CAST('0' AS INT)                   AS Transportation_Distance_km
                        ,CASE
                              WHEN dwo.HAZARDOUS<>'N/A' THEN 'hazardous waste'
                              WHEN dwo.NON_HAZARDOUS<>'N/A' THEN 'Non-hazardous waste'
                              ELSE NULL
                        END                                AS The_type_of_waste
                        ,CASE
                              WHEN dwo.HAZARDOUS<>'N/A' THEN dwo.HAZARDOUS
                              WHEN dwo.NON_HAZARDOUS<>'N/A' THEN dwo.NON_HAZARDOUS
                              ELSE NULL
                        END                                AS Waste_type
                        ,dtm.TREATMENT_METHOD_ENGLISH_NAME  AS Waste_Treatment_method
                        ,dwo.QUANTITY                       AS Weight_of_waste_treated_Unit_kg
                        ,CAST('0' AS INT)                   AS TKT_Ton_km
                    FROM   dbo.DATA_WASTE_OUTPUT_CUSTOMER dwo
                          LEFT JOIN dbo.DATA_TREATMENT_VENDOR dtv
                                ON  dtv.TREATMENT_VENDOR_ID = dwo.TREATMENT_SUPPLIER
                          LEFT JOIN dbo.DATA_TREATMENT_METHOD dtm
                                ON  dtm.TREATMENT_METHOD_ID = dwo.TREATMENT_METHOD_ID
                          LEFT JOIN dbo.DATA_EMERET_WASTE_CATEGORY dewc
                                ON  dewc.EMERET_WASTE_ID = dwo.EMERET_WASTE_ID
                          LEFT JOIN dbo.DATA_WASTE_CATEGORY dwc
                                ON  dwc.WASTE_CODE = dwo.WASTE_CODE
                          LEFT JOIN dbo.DATA_TREATMENT_DISTANCE td
                                ON  td.CODE = dwo.LOCATION_CODE
                    ${where}`;
    // ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
    // OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY`;

    // console.log(query);
    const countQuery = `SELECT COUNT(*) AS total
                        FROM dbo.DATA_WASTE_OUTPUT_CUSTOMER dwo
                        LEFT JOIN dbo.DATA_TREATMENT_VENDOR dtv
                              ON dtv.TREATMENT_VENDOR_ID = dwo.TREATMENT_SUPPLIER
                        LEFT JOIN dbo.DATA_TREATMENT_METHOD dtm
                              ON dtm.TREATMENT_METHOD_ID = dwo.TREATMENT_METHOD_ID
                        LEFT JOIN dbo.DATA_EMERET_WASTE_CATEGORY dewc
                              ON dewc.EMERET_WASTE_ID = dwo.EMERET_WASTE_ID
                        LEFT JOIN dbo.DATA_WASTE_CATEGORY dwc
                              ON dwc.WASTE_CODE = dwo.WASTE_CODE
                        LEFT JOIN dbo.DATA_TREATMENT_DISTANCE td
                              ON td.CODE = dwo.LOCATION_CODE
                        ${where}`;

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

    // const total = totalResult[0]?.total || 0;
    // const hasMore = offset + data.length < total;
    // return { data, page, limit, total, hasMore };
  }
}
