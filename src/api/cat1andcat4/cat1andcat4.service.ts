import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

@Injectable()
export class Cat1andcat4Service {
  constructor(
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
    @Inject('LHG_ERP') private readonly LHG_ERP: Sequelize,
    @Inject('LVL_ERP') private readonly LVL_ERP: Sequelize,
    @Inject('LYM_ERP') private readonly LYM_ERP: Sequelize,
  ) {}

  async getDataCat1AndCat4(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'No',
    sortOrder: string = 'asc',
  ) {
    let db: Sequelize;
    switch (factory) {
      case 'LYV':
        db = this.LYV_ERP;
        break;
      case 'LHG':
        // db = this.LHG_ERP;
        return 'LHG coming soon...';
        // break;
      case 'LYM':
        // db = this.LYM_ERP;
        return 'LYM coming soon...';
      // break;
      case 'LVL':
        // db = this.LVL_ERP;
        return 'LHG coming soon...';
      // break;
      default:
        return 'Coming soon...';
      // return await this.getAllFactoryData(
      //   dateFrom,
      //   dateTo,
      //   page,
      //   limit,
      //   sortField,
      //   sortOrder,
      // );
    }

    return await this.getAFactoryData(
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

    //   if (dateFrom && dateTo) {
    //     where += ` AND CONVERT(VARCHAR, c.USERDate, 23) BETWEEN ? AND ?`;
    //     replacements.push(dateFrom, dateTo);
    //   }
    //   // console.log(dateFrom, dateTo, factory, page, limit, sortField, sortOrder);
    //   const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY c.USERDate) AS INT) AS [No]
    //                       ,c.USERDate        AS [Date]
    //                       ,c.CGNO            AS Purchase_Order
    //                       ,c2.CLBH           AS Material_No
    //                       ,CAST('0' AS INT)     AS [Weight]
    //                       ,z.SupplierCode AS Supplier_Code
    //                       ,z.ThirdCountryLandTransport AS Thirdcountry_Land_Transport
    //                       ,z.PortofDeparture AS Port_Of_Departure
    //                       ,z.PortofArrival AS Port_Of_Arrival
    //                       ,z.Transportationmethod AS Factory_Domestic_Land_Transport
    //                       ,CAST('0' AS INT)     AS Land_Transport_Distance
    //                       ,z.SeaTransportDistance AS Sea_Transport_Distance
    //                       ,CAST('0' AS INT) AS Air_Transport_Distance
    //                       ,CAST('0' AS INT) AS Land_Transport_Ton_Kilometers
    //                       ,CAST('0' AS INT) AS Sea_Transport_Ton_Kilometers
    //                       ,CAST('0' AS INT) AS Air_Transport_Ton_Kilometers
    //                 FROM   CGZL              AS c
    //                       INNER JOIN CGZLS  AS c2
    //                             ON  c2.CGNO = c.CGNO
    //                       LEFT JOIN zszl    AS z
    //                             ON  z.zsdh = c.CGNO
    //                 ${where}`;
    //   const countQuery = `SELECT COUNT(*) total
    //                     FROM   (
    //                               SELECT CAST(ROW_NUMBER() OVER(ORDER BY c.USERDate) AS INT) AS [No]
    //                                     ,c.USERDate              AS [Date]
    //                                     ,c.CGNO                  AS Purchase_Order
    //                                     ,c2.CLBH                 AS Material_No
    //                                     ,z.SupplierCode          AS Supplier_Code
    //                                     ,z.ThirdCountryLandTransport AS Thirdcountry_Land_Transport
    //                                     ,z.PortofDeparture       AS Port_Of_Departure
    //                                     ,z.PortofArrival         AS Port_Of_Arrival
    //                                     ,z.Transportationmethod  AS Factory_Domestic_Land_Transport
    //                                     ,z.SeaTransportDistance  AS Sea_Transport_Distance
    //                               FROM   CGZL                    AS c
    //                                       INNER JOIN CGZLS        AS c2
    //                                           ON  c2.CGNO = c.CGNO
    //                                       LEFT JOIN zszl          AS z
    //                                           ON  z.zsdh = c.CGNO
    //                               ${where}
    //                     ) AS Sub`;

    //   const [dataResults, countResults] = await Promise.all([
    //     this.LYV_ERP.query(query, {
    //       replacements,
    //       type: QueryTypes.SELECT,
    //     }),
    //     this.LYV_ERP.query(countQuery, {
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
    // } catch (error: any) {
    //   throw new InternalServerErrorException(error);
    // }
  }

  private async getAFactoryData(
    db: Sequelize,
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    try {
      const offset = (page - 1) * limit;
      let where = 'WHERE 1=1';
      const replacements: any[] = [];
      if (dateFrom && dateTo) {
        where += ` AND CONVERT(VARCHAR, c.USERDate, 23) BETWEEN ? AND ?`;
        replacements.push(dateFrom, dateTo);
      }
      // console.log(dateFrom, dateTo, factory, page, limit, sortField, sortOrder);
      const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY c.USERDate) AS INT) AS [No]
                            ,c.USERDate                   AS [Date]
                            ,c.CGNO                       AS Purchase_Order
                            ,c2.CLBH                      AS Material_No
                            ,CAST('0' AS INT)             AS [Weight]
                            ,'SupplierCode'              AS Supplier_Code
                            ,'ThirdCountryLandTransport'  AS Thirdcountry_Land_Transport
                            ,'PortofDeparture'            AS Port_Of_Departure
                            ,'PortofArrival'              AS Port_Of_Arrival
                            ,'Transportationmethod'       AS Factory_Domestic_Land_Transport
                            ,CAST('0' AS INT)             AS Land_Transport_Distance
                            ,'SeaTransportDistance'       AS Sea_Transport_Distance
                            ,CAST('0' AS INT)             AS Air_Transport_Distance
                            ,CAST('0' AS INT)             AS Land_Transport_Ton_Kilometers
                            ,CAST('0' AS INT)             AS Sea_Transport_Ton_Kilometers
                            ,CAST('0' AS INT)             AS Air_Transport_Ton_Kilometers
                    FROM   CGZL              AS c
                          INNER JOIN CGZLS  AS c2
                                ON  c2.CGNO = c.CGNO
                          LEFT JOIN zszl    AS z
                                ON  z.zsdh = c.CGNO
                    ${where}`;
      const countQuery = `SELECT COUNT(*) total
                        FROM   (
                                  SELECT CAST(ROW_NUMBER() OVER(ORDER BY c.USERDate) AS INT) AS [No]
                                        ,c.USERDate                   AS [Date]
                                        ,c.CGNO                       AS Purchase_Order
                                        ,c2.CLBH                      AS Material_No
                                        ,CAST('0' AS INT)             AS [Weight]
                                        ,'SupplierCode'              AS Supplier_Code
                                        ,'ThirdCountryLandTransport'  AS Thirdcountry_Land_Transport
                                        ,'PortofDeparture'            AS Port_Of_Departure
                                        ,'PortofArrival'              AS Port_Of_Arrival
                                        ,'Transportationmethod'       AS Factory_Domestic_Land_Transport
                                        ,CAST('0' AS INT)             AS Land_Transport_Distance
                                        ,'SeaTransportDistance'       AS Sea_Transport_Distance
                                        ,CAST('0' AS INT)             AS Air_Transport_Distance
                                        ,CAST('0' AS INT)             AS Land_Transport_Ton_Kilometers
                                        ,CAST('0' AS INT)             AS Sea_Transport_Ton_Kilometers
                                        ,CAST('0' AS INT)             AS Air_Transport_Ton_Kilometers
                    FROM   CGZL              AS c
                          INNER JOIN CGZLS  AS c2
                                ON  c2.CGNO = c.CGNO
                          LEFT JOIN zszl    AS z
                                ON  z.zsdh = c.CGNO
                                  ${where}
                        ) AS Sub`;
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
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  private async getAllFactoryData(
    dateFrom: string,
    dateTo: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    try {
      const offset = (page - 1) * limit;
      let where = 'WHERE 1=1';
      const replacements: any[] = [];
      if (dateFrom && dateTo) {
        where += ` AND CONVERT(VARCHAR, c.USERDate, 23) BETWEEN ? AND ?`;
        replacements.push(dateFrom, dateTo);
      }
      // console.log(dateFrom, dateTo, factory, page, limit, sortField, sortOrder);
      const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY c.USERDate) AS INT) AS [No]
                            ,c.USERDate                   AS [Date]
                            ,c.CGNO                       AS Purchase_Order
                            ,c2.CLBH                      AS Material_No
                            ,CAST('0' AS INT)             AS [Weight]
                            ,'SupplierCode'              AS Supplier_Code
                            ,'ThirdCountryLandTransport'  AS Thirdcountry_Land_Transport
                            ,'PortofDeparture'            AS Port_Of_Departure
                            ,'PortofArrival'              AS Port_Of_Arrival
                            ,'Transportationmethod'       AS Factory_Domestic_Land_Transport
                            ,CAST('0' AS INT)             AS Land_Transport_Distance
                            ,'SeaTransportDistance'       AS Sea_Transport_Distance
                            ,CAST('0' AS INT)             AS Air_Transport_Distance
                            ,CAST('0' AS INT)             AS Land_Transport_Ton_Kilometers
                            ,CAST('0' AS INT)             AS Sea_Transport_Ton_Kilometers
                            ,CAST('0' AS INT)             AS Air_Transport_Ton_Kilometers
                    FROM   CGZL              AS c
                          INNER JOIN CGZLS  AS c2
                                ON  c2.CGNO = c.CGNO
                          LEFT JOIN zszl    AS z
                                ON  z.zsdh = c.CGNO
                    ${where}`;
      const countQuery = `SELECT COUNT(*) total
                        FROM   (
                                  SELECT CAST(ROW_NUMBER() OVER(ORDER BY c.USERDate) AS INT) AS [No]
                                          ,c.USERDate                   AS [Date]
                                          ,c.CGNO                       AS Purchase_Order
                                          ,c2.CLBH                      AS Material_No
                                          ,CAST('0' AS INT)             AS [Weight]
                                          ,'SupplierCode'              AS Supplier_Code
                                          ,'ThirdCountryLandTransport'  AS Thirdcountry_Land_Transport
                                          ,'PortofDeparture'            AS Port_Of_Departure
                                          ,'PortofArrival'              AS Port_Of_Arrival
                                          ,'Transportationmethod'       AS Factory_Domestic_Land_Transport
                                          ,CAST('0' AS INT)             AS Land_Transport_Distance
                                          ,'SeaTransportDistance'       AS Sea_Transport_Distance
                                          ,CAST('0' AS INT)             AS Air_Transport_Distance
                                          ,CAST('0' AS INT)             AS Land_Transport_Ton_Kilometers
                                          ,CAST('0' AS INT)             AS Sea_Transport_Ton_Kilometers
                                          ,CAST('0' AS INT)             AS Air_Transport_Ton_Kilometers
                                  FROM   CGZL              AS c
                                        INNER JOIN CGZLS  AS c2
                                              ON  c2.CGNO = c.CGNO
                                        LEFT JOIN zszl    AS z
                                              ON  z.zsdh = c.CGNO
                                  ${where}
                        ) AS Sub`;
      const connects = [this.LYV_ERP, this.LHG_ERP, this.LYM_ERP, this.LVL_ERP];
      const [dataResults, countResults] = await Promise.all([
        Promise.all(
          connects.map((conn) => {
            return conn.query(query, {
              type: QueryTypes.SELECT,
              replacements,
            });
          }),
        ),
        Promise.all(
          connects.map((conn) => {
            return conn.query(countQuery, {
              type: QueryTypes.SELECT,
              replacements,
            });
          }),
        ),
      ]);
      // console.log(dataResults,countResults);
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
