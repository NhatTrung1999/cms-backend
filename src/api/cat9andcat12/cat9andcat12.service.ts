import {
  Inject,
  Injectable,
  UnauthorizedException,
  InternalServerErrorException,
} from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import { IDataCat9AndCat12, IDataPortCode } from 'src/types/cat9andcat12';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import { v4 as uuidv4 } from 'uuid';
import dayjs from 'dayjs';
import { buildQueryAutoSentCMS } from 'src/helper/cat9andcat12.helper';
dayjs().format();

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
        db = this.LHG_ERP;
        break;
      case 'LYM':
        db = this.LYM_ERP;
        break;
      case 'LVL':
        db = this.LVL_ERP;
        break;
      case 'LYF':
        db = this.LYF_ERP;
        break;
      case 'JAZ':
        db = this.JAZ_ERP;
        break;
      case 'JZS':
        db = this.JZS_ERP;
        break;
      default:
        return await this.getAllDataFactory(
          dateFrom,
          dateTo,
          factory,
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
      ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
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
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    try {
      const offset = (page - 1) * limit;

      let where = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
      let where1 = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
      const replacements: any[] = [];

      if (dateFrom && dateTo) {
        where += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
        where1 += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
        replacements.push(dateFrom, dateTo, dateFrom, dateTo);
      }

      const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                    FROM   (
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,sb.Booking_No AS Booking_No
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,p.GW                    AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'SEA'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_M im
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = im.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = im.INV_NO
                                    LEFT JOIN (
                                              SELECT INV_NO
                                                    ,RYNO
                                                    ,SUM(PAIRS)     Qty
                                                    ,SUM(GW)        GW
                                              FROM   PACKING
                                              GROUP BY
                                                    INV_NO
                                                    ,RYNO
                                          ) p
                                          ON  p.INV_NO = id.INV_NO
                                              AND p.RYNO = id.RYNO
                                    LEFT JOIN YWDD y
                                          ON  y.DDBH = id.RYNO
                                    LEFT JOIN DE_ORDERM do
                                          ON  do.ORDERNO = y.YSBH
                                    LEFT JOIN B_GradeOrder bg
                                          ON  bg.ORDER_B = y.YSBH
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'SEA'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where} AND NOT EXISTS (
                                                    SELECT 1
                                                    FROM   INVOICE_SAMPLE is1
                                                    WHERE  is1.Inv_No = im.Inv_No
                                                )
                              UNION
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,sb.Booking_No AS Booking_No
                                    ,is1.Inv_No              AS Invoice_Number
                                    ,'SAMPLE SHOE'           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,is1.GW                  AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'AIR'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_SAMPLE is1
                                    LEFT JOIN INVOICE_M im
                                          ON  im.Inv_No = is1.Inv_No
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = is1.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = is1.Inv_No
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'AIR'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where1}
                          ) AS Cat9AndCat12`;

      const countQuery = `SELECT COUNT(*) AS total
                        FROM   (
                                  SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                    FROM   (
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,p.GW                    AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'SEA'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_M im
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = im.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = im.INV_NO
                                    LEFT JOIN (
                                              SELECT INV_NO
                                                    ,RYNO
                                                    ,SUM(PAIRS)     Qty
                                                    ,SUM(GW)        GW
                                              FROM   PACKING
                                              GROUP BY
                                                    INV_NO
                                                    ,RYNO
                                          ) p
                                          ON  p.INV_NO = id.INV_NO
                                              AND p.RYNO = id.RYNO
                                    LEFT JOIN YWDD y
                                          ON  y.DDBH = id.RYNO
                                    LEFT JOIN DE_ORDERM do
                                          ON  do.ORDERNO = y.YSBH
                                    LEFT JOIN B_GradeOrder bg
                                          ON  bg.ORDER_B = y.YSBH
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'SEA'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where} AND NOT EXISTS (
                                                    SELECT 1
                                                    FROM   INVOICE_SAMPLE is1
                                                    WHERE  is1.Inv_No = im.Inv_No
                                                )
                              UNION
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,is1.Inv_No              AS Invoice_Number
                                    ,'SAMPLE SHOE'           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,is1.GW                  AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'AIR'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_SAMPLE is1
                                    LEFT JOIN INVOICE_M im
                                          ON  im.Inv_No = is1.Inv_No
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = is1.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = is1.Inv_No
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'AIR'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where1}
                          ) AS Cat9AndCat12
                        ) AS sub`;

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

      // console.log(dataResults);

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

  private async getAllDataFactory(
    dateFrom: string,
    dateTo: string,
    factory: string,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    const offset = (page - 1) * limit;

    let where = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
    let where1 = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
      where1 += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo, dateFrom, dateTo);
    }

    const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                    FROM   (
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,p.GW                    AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'SEA'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_M im
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = im.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = im.INV_NO
                                    LEFT JOIN (
                                              SELECT INV_NO
                                                    ,RYNO
                                                    ,SUM(PAIRS)     Qty
                                                    ,SUM(GW)        GW
                                              FROM   PACKING
                                              GROUP BY
                                                    INV_NO
                                                    ,RYNO
                                          ) p
                                          ON  p.INV_NO = id.INV_NO
                                              AND p.RYNO = id.RYNO
                                    LEFT JOIN YWDD y
                                          ON  y.DDBH = id.RYNO
                                    LEFT JOIN DE_ORDERM do
                                          ON  do.ORDERNO = y.YSBH
                                    LEFT JOIN B_GradeOrder bg
                                          ON  bg.ORDER_B = y.YSBH
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'SEA'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where} AND NOT EXISTS (
                                                    SELECT 1
                                                    FROM   INVOICE_SAMPLE is1
                                                    WHERE  is1.Inv_No = im.Inv_No
                                                )
                              UNION
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,is1.Inv_No              AS Invoice_Number
                                    ,'SAMPLE SHOE'           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,is1.GW                  AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'AIR'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_SAMPLE is1
                                    LEFT JOIN INVOICE_M im
                                          ON  im.Inv_No = is1.Inv_No
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = is1.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = is1.Inv_No
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'AIR'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where1}
                          ) AS Cat9AndCat12`;

    const countQuery = `SELECT COUNT(*) AS total
                        FROM   (
                                  SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                    FROM   (
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,p.GW                    AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'SEA'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_M im
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = im.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = im.INV_NO
                                    LEFT JOIN (
                                              SELECT INV_NO
                                                    ,RYNO
                                                    ,SUM(PAIRS)     Qty
                                                    ,SUM(GW)        GW
                                              FROM   PACKING
                                              GROUP BY
                                                    INV_NO
                                                    ,RYNO
                                          ) p
                                          ON  p.INV_NO = id.INV_NO
                                              AND p.RYNO = id.RYNO
                                    LEFT JOIN YWDD y
                                          ON  y.DDBH = id.RYNO
                                    LEFT JOIN DE_ORDERM do
                                          ON  do.ORDERNO = y.YSBH
                                    LEFT JOIN B_GradeOrder bg
                                          ON  bg.ORDER_B = y.YSBH
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'SEA'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where} AND NOT EXISTS (
                                                    SELECT 1
                                                    FROM   INVOICE_SAMPLE is1
                                                    WHERE  is1.Inv_No = im.Inv_No
                                                )
                              UNION
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,is1.Inv_No              AS Invoice_Number
                                    ,'SAMPLE SHOE'           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,is1.GW                  AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CASE 
                                          WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE 
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='YF' THEN 'IDSRG'
                                                                                    WHEN SUBSTRING(
                                                                                              im.INV_NO
                                                                                            ,CHARINDEX('/' ,im.INV_NO)+1
                                                                                            ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
                                                                                            - CHARINDEX('/' ,im.INV_NO)- 1
                                                                                          )='TY' THEN 'MMRGN'
                                                                                    ELSE 'VNCLP'
                                                                                END
                                          ELSE CASE 
                                                    WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
                                                    ELSE 'VNCLP'
                                              END
                                    END                     AS Port_Of_Departure
                                    ,pc.PortCode             AS Port_Of_Arrival
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,'AIR'                   AS Transport_Method
                                    ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
                                    ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
                              FROM   INVOICE_SAMPLE is1
                                    LEFT JOIN INVOICE_M im
                                          ON  im.Inv_No = is1.Inv_No
                                    LEFT JOIN INVOICE_D     AS id
                                          ON  id.INV_NO = is1.INV_NO
                                    LEFT JOIN Ship_Booking  AS sb
                                          ON  sb.INV_NO = is1.Inv_No
                                    LEFT JOIN (
                                              SELECT CustomerNumber
                                                    ,PortCode
                                                    ,TransportMethod
                                              FROM   CMW.CMW.dbo.CMW_PortCode
                                              WHERE  TransportMethod = 'AIR'
                                          ) pc
                                          ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                              ${where1}
                          ) AS Cat9AndCat12
                        ) AS sub`;
    const connects = [
      this.LYV_ERP,
      this.LHG_ERP,
      this.LYM_ERP,
      this.LVL_ERP,
      this.LYF_ERP,
      this.JAZ_ERP,
      this.JZS_ERP,
    ];
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
  }

  async importExcelPortCode(file: Express.Multer.File) {
    try {
      let insertCount = 0;
      let updateCount = 0;
      if (!file.path || !fs.existsSync(file.path)) {
        throw new Error('File path not found!');
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(file.path);
      const worksheet = workbook.getWorksheet(1);
      if (!worksheet) {
        throw new Error('No worksheet found in the Excel file');
      }

      const headerRow = worksheet.getRow(1);
      // console.log(headerRow);
      const headers: string[] = [];
      headerRow.eachCell((cell) => {
        if (cell.value) {
          headers.push(cell.value.toString().trim().replace(/\s+/g, '_'));
        }
      });

      const requiredHeaders = [
        'Customer_Number',
        'Transport_Method',
        'Port_Code',
      ];
      const missingHeaders = requiredHeaders.filter(
        (h) => !headers.includes(h),
      );

      if (missingHeaders.length > 0) {
        throw new Error(
          `Excel file format is invalid! Missing columns: ${missingHeaders.join(', ')}`,
        );
      }

      const data: {
        Customer_Number: string;
        Transport_Method: string;
        Port_Code: string;
      }[] = [];

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber === 1 && row.cellCount > 0) return;

        const rowData: any = {};
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const headerCell = headerRow.getCell(colNumber);
          const headerValue = headerCell.value;
          if (headerValue !== null && headerValue !== undefined) {
            const key = headerValue.toString().trim().replace(/\s+/g, '_');
            rowData[key] = cell.value;
            // console.log(cell.value);
          }
        });
        if (
          rowData?.Customer_Number &&
          rowData?.Transport_Method &&
          rowData?.Port_Code
        ) {
          data.push(rowData);
        }
      });
      // console.log(data);
      // return { data, length: data.length };

      for (let item of data) {
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
                'admin',
                'LYV',
                item.Customer_Number,
                item.Transport_Method,
              ],
              type: QueryTypes.SELECT,
            },
          );
          updateCount++;
        } else {
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
                'admin',
                'LYV',
              ],
              type: QueryTypes.INSERT,
            },
          );
          insertCount++;
        }
      }
      const records: any = await this.EIP.query(
        `SELECT *
          FROM CMW_PortCode`,
        { type: QueryTypes.SELECT },
      );
      const message = `Processed successfully! Inserted: ${insertCount} records, Updated: ${updateCount} records. Total rows processed: ${data.length}.`;
      return { message, records };
    } catch (error: any) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }

  async autoSentCMS(dateFrom: string, dateTo: string) {
    // let where = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
    // let where1 = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
    // const replacements: any[] = [];

    // if (dateFrom && dateTo) {
    //   where += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
    //   where1 += ` AND CONVERT(VARCHAR ,im.INV_DATE ,23) BETWEEN ? AND ?`;
    //   replacements.push(dateFrom, dateTo, dateFrom, dateTo);
    // }

    // const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
    //                       ,*
    //                 FROM   (
    //                                 SELECT CONVERT(VARCHAR(10), im.INV_DATE, 111)             AS [Date]
    //                                 ,CONVERT(VARCHAR(10), sb.ExFty_Date, 111)           AS Shipment_Date
    //                                 ,sb.Booking_No AS Booking_No
    //                                 ,im.INV_NO               AS Invoice_Number
    //                                 --,id.STYLE_NAME           AS Article_Name
    //                                 --,id.ARTICLE              AS Article_ID
    //                                 ,p.Qty                   AS Quantity
    //                                 ,p.GW                    AS Gross_Weight
    //                                 ,im.CUSTID               AS Customer_ID
    //                                 ,'Truck'                 AS Local_Land_Transportation
    //                                 ,CASE
    //                                       WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE
    //                                                                                 WHEN SUBSTRING(
    //                                                                                           im.INV_NO
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO)+1
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
    //                                                                                         - CHARINDEX('/' ,im.INV_NO)- 1
    //                                                                                       ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
    //                                                                                 WHEN SUBSTRING(
    //                                                                                           im.INV_NO
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO)+1
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
    //                                                                                         - CHARINDEX('/' ,im.INV_NO)- 1
    //                                                                                       )='YF' THEN 'IDSRG'
    //                                                                                 WHEN SUBSTRING(
    //                                                                                           im.INV_NO
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO)+1
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
    //                                                                                         - CHARINDEX('/' ,im.INV_NO)- 1
    //                                                                                       )='TY' THEN 'MMRGN'
    //                                                                                 ELSE 'VNCLP'
    //                                                                             END
    //                                       ELSE CASE
    //                                                 WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
    //                                                 ELSE 'VNCLP'
    //                                           END
    //                                 END                     AS Port_Of_Departure
    //                                 ,pc.PortCode             AS Port_Of_Arrival
    //                                 ,CAST('0' AS INT)        AS Land_Transport_Distance
    //                                 ,CAST('0' AS INT)        AS Sea_Transport_Distance
    //                                 ,CAST('0' AS INT)        AS Air_Transport_Distance
    //                                 ,'SEA'                   AS Transport_Method
    //                                 ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
    //                                 ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
    //                                 ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
    //                           FROM   INVOICE_M im
    //                                 LEFT JOIN INVOICE_D     AS id
    //                                       ON  id.INV_NO = im.INV_NO
    //                                 LEFT JOIN Ship_Booking  AS sb
    //                                       ON  sb.INV_NO = im.INV_NO
    //                                 LEFT JOIN (
    //                                           SELECT INV_NO
    //                                                 ,SUM(PAIRS)     Qty
    //                                                 ,SUM(GW)        GW
    //                                           FROM   PACKING
    //                                           GROUP BY
    //                                                 INV_NO
    //                                       ) p
    //                                       ON  p.INV_NO = id.INV_NO
    //                                 LEFT JOIN YWDD y
    //                                       ON  y.DDBH = id.RYNO
    //                                 LEFT JOIN DE_ORDERM do
    //                                       ON  do.ORDERNO = y.YSBH
    //                                 LEFT JOIN B_GradeOrder bg
    //                                       ON  bg.ORDER_B = y.YSBH
    //                                 LEFT JOIN (
    //                                           SELECT CustomerNumber
    //                                                 ,PortCode
    //                                                 ,TransportMethod
    //                                           FROM   CMW.CMW.dbo.CMW_PortCode
    //                                           WHERE  TransportMethod = 'SEA'
    //                                       ) pc
    //                                       ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
    //                           ${where} AND NOT EXISTS (
    //                                                 SELECT 1
    //                                                 FROM   INVOICE_SAMPLE is1
    //                                                 WHERE  is1.Inv_No = im.Inv_No
    //                                             )
    //                           UNION
    //                           SELECT CONVERT(VARCHAR(10), im.INV_DATE, 111)             AS [Date]
    //                                 ,CONVERT(VARCHAR(10), sb.ExFty_Date, 111)           AS Shipment_Date
    //                                 ,sb.Booking_No AS Booking_No
    //                                 ,is1.Inv_No              AS Invoice_Number
    //                                 --,'SAMPLE SHOE'           AS Article_Name
    //                                 --,id.ARTICLE              AS Article_ID
    //                                 ,is1.Qty                 AS Quantity
    //                                 ,is1.GW                  AS Gross_Weight
    //                                 ,im.CUSTID               AS Customer_ID
    //                                 ,'Truck'                 AS Local_Land_Transportation
    //                                 ,CASE
    //                                       WHEN CHARINDEX('/' ,im.INV_NO)>0 THEN CASE
    //                                                                                 WHEN SUBSTRING(
    //                                                                                           im.INV_NO
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO)+1
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
    //                                                                                         - CHARINDEX('/' ,im.INV_NO)- 1
    //                                                                                       ) IN ('LT' ,'LT2' ,'TX') THEN 'VNCLP'
    //                                                                                 WHEN SUBSTRING(
    //                                                                                           im.INV_NO
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO)+1
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
    //                                                                                         - CHARINDEX('/' ,im.INV_NO)- 1
    //                                                                                       )='YF' THEN 'IDSRG'
    //                                                                                 WHEN SUBSTRING(
    //                                                                                           im.INV_NO
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO)+1
    //                                                                                         ,CHARINDEX('/' ,im.INV_NO ,CHARINDEX('/' ,im.INV_NO)+1)
    //                                                                                         - CHARINDEX('/' ,im.INV_NO)- 1
    //                                                                                       )='TY' THEN 'MMRGN'
    //                                                                                 ELSE 'VNCLP'
    //                                                                             END
    //                                       ELSE CASE
    //                                                 WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,3)='LYV' THEN 'SGN'
    //                                                 ELSE 'VNCLP'
    //                                           END
    //                                 END                     AS Port_Of_Departure
    //                                 ,pc.PortCode             AS Port_Of_Arrival
    //                                 ,CAST('0' AS INT)        AS Land_Transport_Distance
    //                                 ,CAST('0' AS INT)        AS Sea_Transport_Distance
    //                                 ,CAST('0' AS INT)        AS Air_Transport_Distance
    //                                 ,'AIR'                   AS Transport_Method
    //                                 ,CAST('0' AS INT)        AS Land_Transport_Ton_Kilometers
    //                                 ,CAST('0' AS INT)        AS Sea_Transport_Ton_Kilometers
    //                                 ,CAST('0' AS INT)        AS Air_Transport_Ton_Kilometers
    //                           FROM   INVOICE_SAMPLE is1
    //                                 LEFT JOIN INVOICE_M im
    //                                       ON  im.Inv_No = is1.Inv_No
    //                                 LEFT JOIN INVOICE_D     AS id
    //                                       ON  id.INV_NO = is1.INV_NO
    //                                 LEFT JOIN Ship_Booking  AS sb
    //                                       ON  sb.INV_NO = is1.Inv_No
    //                                 LEFT JOIN (
    //                                           SELECT CustomerNumber
    //                                                 ,PortCode
    //                                                 ,TransportMethod
    //                                           FROM   CMW.CMW.dbo.CMW_PortCode
    //                                           WHERE  TransportMethod = 'AIR'
    //                                       ) pc
    //                                       ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
    //                           ${where1}
    //                       ) AS Cat9AndCat12`;
    const replacements =
      dateFrom && dateTo ? [dateFrom, dateTo, dateFrom, dateTo] : [];

    const connects = [
      // { factory: 'LYV', db: this.LYV_ERP },
      // { factory: 'LHG', db: this.LHG_ERP },
      // { factory: 'LYM', db: this.LYM_ERP },
      { factory: 'LVL', db: this.LVL_ERP },
      // { factory: 'LYF', db: this.LYF_ERP },
      // { factory: 'JAZ', db: this.JAZ_ERP },
      // { factory: 'JZS', db: this.JZS_ERP },
    ];

    const dataResults = await Promise.all(
      connects.map(async (conn) => {
        const query = await buildQueryAutoSentCMS(
          dateFrom,
          dateTo,
          conn.factory,
          this.EIP
        );
        return conn.db.query(query, {
          type: QueryTypes.SELECT,
          replacements,
        });
      }),
    );

    // console.log(dataResults);
    const allData = dataResults.flat();

    let data = allData.map((item, index) => ({ ...item, No: index + 1 }));

    const formatData = data.map((itemData: any) => {
      const date = itemData.Date;
      const factory = itemData.Factory;
      const factoryAddress = itemData.Factory_address;
      const shipmentDate = itemData.Shipment_Date;
      const bookingNo = itemData.Booking_No;
      const customerID = itemData.Customer_ID;
      const invoiceNumber = itemData.Invoice_Number;
      const transportMethod = itemData.Transport_Method; // AIR OR SEA
      const portOfArrival = itemData.Port_Of_Arrival; // DEPARTURE AND END PORT
      const portOfDeparture = itemData.Port_Of_Departure; // ST PORT
      const quantity = itemData.Quantity;
      const grossWeight = itemData.Gross_Weight; // ACTIVITY DATA
      const createdDate = dayjs().format('YYYY-MM-DD HH:mm:ss');

      return {
        System: 'ERP', // DEFAULT
        Corporation: 'LAI YIH', // DEFAULT
        Factory: factory,
        Department: 'Shipping', // DEFAULT
        DocKey: 'S3.C9', // DEFAULT
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: '3.2 下游貨物運輸及配送', // DEFAULT
        DataType: '1', // DEFAULT
        DocType: '應收憑單', // DEFAULT
        UndDoc: '銷貨單', // DEFAULT
        DocFlow: '銷貨相關流程', // DEFAULT
        DocDate: date,
        DocDate2: shipmentDate,
        DocNo: bookingNo, // DEFAULT
        UndDocNo: '', // DEFAULT
        CustVenName: customerID,
        InvoiceNo: invoiceNumber,
        TransType: transportMethod === 'AIR' ? '空運' : '海運',
        Departure: factoryAddress,
        Destination: portOfArrival,
        PortType: transportMethod === 'AIR' ? '空港' : '海港',
        StPort: portOfDeparture,
        ThPort: '', // DEFAULT
        EndPort: portOfArrival,
        Product: '', // DEFAULT
        Quity: quantity,
        Amount: '', // DEFAULT
        ActivityData: +Number(grossWeight).toFixed(2),
        ActivityUnit: '噸', // DEFAULT
        Unit: '噸', // DEFAULT
        UnitWeight: '', // DEFAULT
        Memo: '', // DEFAULT
        CreateDateTime: createdDate,
        Creator: '',
      };
    });

    return formatData;
  }
}
