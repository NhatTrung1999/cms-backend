import {
  Inject,
  Injectable,
  UnauthorizedException,
  InternalServerErrorException,
} from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import {
  ACTIVITY_TYPES,
  ActivityType,
  IDataCat9AndCat12,
  IDataPortCode,
} from 'src/types/cat9andcat12';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import { v4 as uuidv4 } from 'uuid';
import dayjs from 'dayjs';
import { buildQueryAutoSentCMS } from 'src/helper/cat9andcat12.helper';
import { FACTORY_LIST, FactoryCode } from 'src/helper/factory.helper';
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
        where += ` AND CONVERT(VARCHAR ,sb.ExFty_Date ,23) BETWEEN ? AND ?`;
        where1 += ` AND CONVERT(VARCHAR ,sb.ExFty_Date ,23) BETWEEN ? AND ?`;
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
                                    ,ISNULL(p.GW ,0)         AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,ISNULL(
                                        ISNULL(
                                            bg.SHPIDS
                                          ,CASE 
                                                WHEN (do.ShipMode='Air')
                                            AND (do.Shipmode_1 IS NULL) THEN '10 AC'
                                                WHEN(do.ShipMode='Air Expres')
                                            AND (do.Shipmode_1 IS NULL) THEN '20 CC'
                                                WHEN(do.ShipMode='Ocean')
                                            AND (do.Shipmode_1 IS NULL) THEN '11 SC'
                                                WHEN do.ShipMode_1 IS NULL THEN ''
                                                ELSE do.ShipMode_1 END
                                        )
                                      ,y.ShipMode
                                    )                       AS Transport_Method
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
                                              FROM   PACKING_D
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
                                    ,ISNULL(id.ARTICLE, 'SAMPLE SHOE')              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,ISNULL(is1.GW ,0)       AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
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
                                    ,ISNULL(p.GW ,0)         AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,ISNULL(
                                        ISNULL(
                                            bg.SHPIDS
                                          ,CASE 
                                                WHEN (do.ShipMode='Air')
                                            AND (do.Shipmode_1 IS NULL) THEN 'AIR'
                                                WHEN(do.ShipMode='Air Expres')
                                            AND (do.Shipmode_1 IS NULL) THEN 'AIR'
                                                WHEN(do.ShipMode='Ocean')
                                            AND (do.Shipmode_1 IS NULL) THEN 'SEA'
                                                WHEN do.Shipmode_1 LIKE '%SC%' THEN 'SEA'
                                                WHEN do.Shipmode_1 LIKE '%AC%' THEN 'AIR'
                                                WHEN do.Shipmode_1 LIKE '%CC%' THEN 'AIR'
                                                WHEN do.ShipMode_1 IS NULL THEN '' 
                                                ELSE 'N/A' END
                                        )
                                      ,y.ShipMode
                                    )                       AS Transport_Method
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
                                              FROM   PACKING_D
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
                                    ,ISNULL(id.ARTICLE, 'SAMPLE SHOE')              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,ISNULL(is1.GW ,0)       AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
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

      let data: any = dataResults;
      data = await Promise.all(
        data.map(async (item) => {
          if (
            item?.Article_ID.trim().toLowerCase() ===
            'SAMPLE SHOE'.trim().toLowerCase()
          ) {
            item.Transport_Method = 'AIR';
            item.Port_Of_Departure = 'SGN';
          } else {
            switch (item?.Transport_Method.trim().toLowerCase()) {
              case '11 SC'.trim().toLowerCase():
              case 'Ocean'.trim().toLowerCase():
              case 'Sea Air Collect'.trim().toLowerCase():
              case 'SC'.trim().toLowerCase():
              case 'Sea/Air'.trim().toLowerCase():
              case 'Partial'.trim().toLowerCase():
                item.Transport_Method = 'SEA';
                break;
              case '20 CC'.trim().toLowerCase():
              case 'AC'.trim().toLowerCase():
              case '10 AC'.trim().toLowerCase():
              case 'CC'.trim().toLowerCase():
              case 'Air PP'.trim().toLowerCase():
              case 'Air Express'.trim().toLowerCase():
              case 'AE'.trim().toLowerCase():
                item.Transport_Method = 'AIR';
                break;
              case 'RAIL'.trim().toLowerCase():
              case 'Rail Collect'.trim().toLowerCase():
              case 'Rail Colle'.trim().toLowerCase():
                item.Transport_Method = 'Land';
                break;
              default:
                item.Transport_Method = 'SEA';
                break;
            }
            switch (factory.trim().toLowerCase()) {
              case 'LYV'.trim().toLowerCase():
              case 'LVL'.trim().toLowerCase():
              case 'LHG'.trim().toLowerCase():
                if (
                  item.Transport_Method.trim().toLowerCase() ===
                  'SEA'.trim().toLowerCase()
                ) {
                  item.Port_Of_Departure = 'VNCLP';
                }
                if (
                  item.Transport_Method.trim().toLowerCase() ===
                  'Land'.trim().toLowerCase()
                ) {
                  item.Port_Of_Departure = 'SOTH';
                }
                if (
                  item.Transport_Method.trim().toLowerCase() ===
                  'AIR'.trim().toLowerCase()
                ) {
                  item.Port_Of_Departure = 'SGN';
                }
                break;
              case 'LYM'.trim().toLowerCase():
                if (
                  item.Transport_Method.trim().toLowerCase() ===
                  'SEA'.trim().toLowerCase()
                ) {
                  item.Port_Of_Departure = 'MMRGN';
                } else {
                  item.Port_Of_Departure = 'RGN';
                }
                break;
              case 'LYF'.trim().toLowerCase():
                if (
                  item.Transport_Method.trim().toLowerCase() ===
                  'SEA'.trim().toLowerCase()
                ) {
                  item.Port_Of_Departure = 'IDSRG';
                } else {
                  item.Port_Of_Departure = 'CGK';
                }
                break;
              default:
                if (
                  item.Transport_Method.trim().toLowerCase() ===
                  'SEA'.trim().toLowerCase()
                ) {
                  item.Port_Of_Departure = 'VNCLP';
                }
                if (
                  item.Transport_Method.trim().toLowerCase() ===
                  'Land'.trim().toLowerCase()
                ) {
                  item.Port_Of_Departure = 'SOTH';
                }
                if (
                  item.Transport_Method.trim().toLowerCase() ===
                  'AIR'.trim().toLowerCase()
                ) {
                  item.Port_Of_Departure = 'SGN';
                }
                break;
            }
          }

          const queryPoA = `SELECT PortCode
                        FROM CMW_PortCode
                        WHERE TransportMethod LIKE ? AND CustomerNumber LIKE ?`;
          const result: any = await this.EIP.query(queryPoA, {
            type: QueryTypes.SELECT,
            replacements: [
              `%${item.Transport_Method}%`,
              `%${item.Customer_ID}%`,
            ],
          });

          return {
            ...item,
            Transport_Method: item.Transport_Method,
            Port_Of_Departure: item.Port_Of_Departure,
            Port_Of_Arrival: result[0]?.PortCode.trim().toUpperCase() || 'N/A',
          };
        }),
      );

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

    const getQuery = (
      factoryName: string,
    ) => `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                          ,'${factoryName}' AS Factory_Name
                    FROM   (
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,sb.Booking_No AS Booking_No
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,ISNULL(p.GW ,0)         AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,ISNULL(
                                        ISNULL(
                                            bg.SHPIDS
                                          ,CASE 
                                                WHEN (do.ShipMode='Air')
                                            AND (do.Shipmode_1 IS NULL) THEN '10 AC'
                                                WHEN(do.ShipMode='Air Expres')
                                            AND (do.Shipmode_1 IS NULL) THEN '20 CC'
                                                WHEN(do.ShipMode='Ocean')
                                            AND (do.Shipmode_1 IS NULL) THEN '11 SC'
                                                WHEN do.ShipMode_1 IS NULL THEN ''
                                                ELSE do.ShipMode_1 END
                                        )
                                      ,y.ShipMode
                                    )                       AS Transport_Method
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
                                              FROM   PACKING_D
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
                                    ,ISNULL(id.ARTICLE, 'SAMPLE SHOE')              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,ISNULL(is1.GW ,0)       AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
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
                              ${where1}
                          ) AS Cat9AndCat1`;
    const getCountQuery = (factoryName: string) => `SELECT COUNT(*) AS total
                        FROM   (
                                  SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                          ,'${factoryName}' AS Factory_Name
                    FROM   (
                              SELECT im.INV_DATE             AS [Date]
                                    ,sb.ExFty_Date           AS Shipment_Date
                                    ,im.INV_NO               AS Invoice_Number
                                    ,id.STYLE_NAME           AS Article_Name
                                    ,id.ARTICLE              AS Article_ID
                                    ,p.Qty                   AS Quantity
                                    ,ISNULL(p.GW ,0)         AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
                                    ,CAST('0' AS INT)        AS Land_Transport_Distance
                                    ,CAST('0' AS INT)        AS Sea_Transport_Distance
                                    ,CAST('0' AS INT)        AS Air_Transport_Distance
                                    ,ISNULL(
                                        ISNULL(
                                            bg.SHPIDS
                                          ,CASE 
                                                WHEN (do.ShipMode='Air')
                                            AND (do.Shipmode_1 IS NULL) THEN 'AIR'
                                                WHEN(do.ShipMode='Air Expres')
                                            AND (do.Shipmode_1 IS NULL) THEN 'AIR'
                                                WHEN(do.ShipMode='Ocean')
                                            AND (do.Shipmode_1 IS NULL) THEN 'SEA'
                                                WHEN do.Shipmode_1 LIKE '%SC%' THEN 'SEA'
                                                WHEN do.Shipmode_1 LIKE '%AC%' THEN 'AIR'
                                                WHEN do.Shipmode_1 LIKE '%CC%' THEN 'AIR'
                                                WHEN do.ShipMode_1 IS NULL THEN '' 
                                                ELSE 'N/A' END
                                        )
                                      ,y.ShipMode
                                    )                       AS Transport_Method
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
                                              FROM   PACKING_D
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
                                    ,ISNULL(id.ARTICLE, 'SAMPLE SHOE')              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,ISNULL(is1.GW ,0)       AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
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
                              ${where1}
                          ) AS Cat9AndCat12
                        ) AS sub`;

    let where = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
    let where1 = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR ,sb.ExFty_Date ,23) BETWEEN ? AND ?`;
      where1 += ` AND CONVERT(VARCHAR ,sb.ExFty_Date ,23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo, dateFrom, dateTo);
    }
    const connects = [
      { conn: this.LYV_ERP, factoryName: 'LYV' },
      { conn: this.LHG_ERP, factoryName: 'LHG' },
      { conn: this.LYM_ERP, factoryName: 'LYM' },
      { conn: this.LVL_ERP, factoryName: 'LVL' },
      { conn: this.LYF_ERP, factoryName: 'LYF' },
      { conn: this.JAZ_ERP, factoryName: 'JAZ' },
      { conn: this.JZS_ERP, factoryName: 'JZS' },
    ];
    const [dataResults, countResults] = await Promise.all([
      Promise.all(
        connects.map(({ conn, factoryName }) => {
          const query = getQuery(factoryName);
          return conn.query(query, {
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

    let data: any = allData.map((item, index) => ({ ...item, No: index + 1 }));

    data = await Promise.all(
      data.map(async (item) => {
        if (
          item?.Article_ID.trim().toLowerCase() ===
          'SAMPLE SHOE'.trim().toLowerCase()
        ) {
          item.Transport_Method = 'AIR';
          item.Port_Of_Departure = 'SGN';
        } else {
          switch (item?.Transport_Method.trim().toLowerCase()) {
            case '11 SC'.trim().toLowerCase():
            case 'Ocean'.trim().toLowerCase():
            case 'Sea Air Collect'.trim().toLowerCase():
            case 'SC'.trim().toLowerCase():
            case 'Sea/Air'.trim().toLowerCase():
            case 'Partial'.trim().toLowerCase():
              item.Transport_Method = 'SEA';
              break;
            case '20 CC'.trim().toLowerCase():
            case 'AC'.trim().toLowerCase():
            case '10 AC'.trim().toLowerCase():
            case 'CC'.trim().toLowerCase():
            case 'Air PP'.trim().toLowerCase():
            case 'Air Express'.trim().toLowerCase():
            case 'AE'.trim().toLowerCase():
              item.Transport_Method = 'AIR';
              break;
            case 'RAIL'.trim().toLowerCase():
            case 'Rail Collect'.trim().toLowerCase():
            case 'Rail Colle'.trim().toLowerCase():
              item.Transport_Method = 'Land';
              break;
            default:
              item.Transport_Method = 'SEA';
              break;
          }
          switch (factory.trim().toLowerCase()) {
            case 'LYV'.trim().toLowerCase():
            case 'LVL'.trim().toLowerCase():
            case 'LHG'.trim().toLowerCase():
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                item.Port_Of_Departure = 'VNCLP';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'Land'.trim().toLowerCase()
              ) {
                item.Port_Of_Departure = 'SOTH';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'AIR'.trim().toLowerCase()
              ) {
                item.Port_Of_Departure = 'SGN';
              }
              break;
            case 'LYM'.trim().toLowerCase():
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                item.Port_Of_Departure = 'MMRGN';
              } else {
                item.Port_Of_Departure = 'RGN';
              }
              break;
            case 'LYF'.trim().toLowerCase():
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                item.Port_Of_Departure = 'IDSRG';
              } else {
                item.Port_Of_Departure = 'CGK';
              }
              break;
            default:
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                item.Port_Of_Departure = 'VNCLP';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'Land'.trim().toLowerCase()
              ) {
                item.Port_Of_Departure = 'SOTH';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'AIR'.trim().toLowerCase()
              ) {
                item.Port_Of_Departure = 'SGN';
              }
              break;
          }
        }

        const queryPoA = `SELECT PortCode
                        FROM CMW_PortCode
                        WHERE TransportMethod LIKE ? AND CustomerNumber LIKE ?`;
        const result: any = await this.EIP.query(queryPoA, {
          type: QueryTypes.SELECT,
          replacements: [`%${item.Transport_Method}%`, `%${item.Customer_ID}%`],
        });

        return {
          ...item,
          Transport_Method: item.Transport_Method,
          Port_Of_Departure: item.Port_Of_Departure,
          Port_Of_Arrival: result[0]?.PortCode.trim().toUpperCase() || 'N/A',
        };
      }),
    );

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

  async importExcelPortCode(
    file: Express.Multer.File,
    userid: string,
    factory: string,
  ) {
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

      // console.log(headers);

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
            rowData[key] = String(cell.value);
          }
        });
        if (
          rowData?.Customer_Number ||
          rowData?.Transport_Method ||
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
                userid,
                factory,
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
                userid,
                factory,
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

  async autoSentCMS(
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    try {
      if (factory.trim().toUpperCase() === 'ALL') {
        return this.autoSentCMSAllFactories(dateFrom, dateTo, dockeyCMS);
      }

      return this.getCMSByFactory(
        factory as FactoryCode,
        dateFrom,
        dateTo,
        dockeyCMS,
      );
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
    // const replacements =
    //   dateFrom && dateTo ? [dateFrom, dateTo, dateFrom, dateTo] : [];

    // const connects = [
    //   { factory: 'LYV', db: this.LYV_ERP },
    //   { factory: 'LHG', db: this.LHG_ERP },
    //   { factory: 'LYM', db: this.LYM_ERP },
    //   { factory: 'LVL', db: this.LVL_ERP },
    //   { factory: 'LYF', db: this.LYF_ERP },
    //   { factory: 'JAZ', db: this.JAZ_ERP },
    //   { factory: 'JZS', db: this.JZS_ERP },
    // ];

    // const dataResults = await Promise.all(
    //   connects.map(async (conn) => {
    //     const query = await buildQueryAutoSentCMS(
    //       dateFrom,
    //       dateTo,
    //       conn.factory,
    //       this.EIP,
    //     );
    //     return conn.db.query(query, {
    //       type: QueryTypes.SELECT,
    //       replacements,
    //     });
    //   }),
    // );

    // // console.log(dataResults);
    // const allData = dataResults.flat();

    // let data = allData.map((item, index) => ({ ...item, No: index + 1 }));

    // const formatData = data.map((itemData: any) => {
    // const date = itemData.Date;
    // const factory = itemData.Factory;
    // const factoryAddress = itemData.Factory_address;
    // const shipmentDate = itemData.Shipment_Date;
    // const bookingNo = itemData.Booking_No;
    // const customerID = itemData.Customer_ID;
    // const invoiceNumber = itemData.Invoice_Number;
    // const transportMethod = itemData.Transport_Method; // AIR OR SEA
    // const portOfArrival = itemData.Port_Of_Arrival; // DEPARTURE AND END PORT
    // const portOfDeparture = itemData.Port_Of_Departure; // ST PORT
    // const quantity = itemData.Quantity;
    // const grossWeight = itemData.Gross_Weight; // ACTIVITY DATA
    // const createdDate = dayjs().format('YYYY-MM-DD HH:mm:ss');

    //   return {
    //     System: 'ERP', // DEFAULT
    //     Corporation: 'LAI YIH', // DEFAULT
    //     Factory: factory,
    //     Department: 'Shipping', // DEFAULT
    //     DocKey: 'S3.C9', // DEFAULT
    //     SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
    //     EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
    //     ActivityType: '3.2 下游貨物運輸及配送', // DEFAULT
    //     DataType: '1', // DEFAULT
    //     DocType: '應收憑單', // DEFAULT
    //     UndDoc: '銷貨單', // DEFAULT
    //     DocFlow: '銷貨相關流程', // DEFAULT
    //     DocDate: date,
    //     DocDate2: shipmentDate,
    //     DocNo: bookingNo, // DEFAULT
    //     UndDocNo: '', // DEFAULT
    //     CustVenName: customerID,
    //     InvoiceNo: invoiceNumber,
    //     TransType: transportMethod === 'AIR' ? '空運' : '海運',
    //     Departure: factoryAddress,
    //     Destination: portOfArrival,
    //     PortType: transportMethod === 'AIR' ? '空港' : '海港',
    //     StPort: portOfDeparture,
    //     ThPort: '', // DEFAULT
    //     EndPort: portOfArrival,
    //     Product: '', // DEFAULT
    //     Quity: quantity,
    //     Amount: '', // DEFAULT
    //     ActivityData: +Number(grossWeight).toFixed(2),
    //     ActivityUnit: '噸', // DEFAULT
    //     Unit: '噸', // DEFAULT
    //     UnitWeight: '', // DEFAULT
    //     Memo: '', // DEFAULT
    //     CreateDateTime: createdDate,
    //     Creator: '',
    //   };
    // });

    // return formatData;
  }

  private async autoSentCMSAllFactories(
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const results = await Promise.all(
      FACTORY_LIST.map((factory) =>
        this.getCMSByFactory(factory, dateFrom, dateTo, dockeyCMS),
      ),
    );

    return results.flat();
  }

  private async getCMSByFactory(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDbByFactory(factory);
    if (!db) return [];

    const query = await buildQueryAutoSentCMS(
      dateFrom,
      dateTo,
      factory,
      this.EIP,
    );

    const replacements =
      dateFrom && dateTo ? [dateFrom, dateTo, dateFrom, dateTo] : [];

    let data = await db.query<any>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    data = await Promise.all(
      data.map(async (item: any) => {
        let portOfDeparture = '';
        if (
          item?.Article_ID.trim().toLowerCase() ===
          'SAMPLE SHOE'.trim().toLowerCase()
        ) {
          item.Transport_Method = 'AIR';
          portOfDeparture = 'SGN';
        } else {
          switch (item?.Transport_Method.trim().toLowerCase()) {
            case '11 SC'.trim().toLowerCase():
            case 'Ocean'.trim().toLowerCase():
            case 'Sea Air Collect'.trim().toLowerCase():
            case 'SC'.trim().toLowerCase():
            case 'Sea/Air'.trim().toLowerCase():
            case 'Partial'.trim().toLowerCase():
              item.Transport_Method = 'SEA';
              break;
            case '20 CC'.trim().toLowerCase():
            case 'AC'.trim().toLowerCase():
            case '10 AC'.trim().toLowerCase():
            case 'CC'.trim().toLowerCase():
            case 'Air PP'.trim().toLowerCase():
            case 'Air Express'.trim().toLowerCase():
            case 'AE'.trim().toLowerCase():
              item.Transport_Method = 'AIR';
              break;
            case 'RAIL'.trim().toLowerCase():
            case 'Rail Collect'.trim().toLowerCase():
            case 'Rail Colle'.trim().toLowerCase():
              item.Transport_Method = 'Land';
              break;
            default:
              item.Transport_Method = 'SEA';
              break;
          }

          switch (item.Factory_Name.trim().toLowerCase()) {
            case 'LYV'.trim().toLowerCase():
            case 'LVL'.trim().toLowerCase():
            case 'LHG'.trim().toLowerCase():
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                portOfDeparture = 'VNCLP';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'Land'.trim().toLowerCase()
              ) {
                portOfDeparture = 'SOTH';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'AIR'.trim().toLowerCase()
              ) {
                portOfDeparture = 'SGN';
              }
              break;
            case 'LYM'.trim().toLowerCase():
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                portOfDeparture = 'MMRGN';
              } else {
                portOfDeparture = 'RGN';
              }
              break;
            case 'LYF'.trim().toLowerCase():
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                portOfDeparture = 'IDSRG';
              } else {
                portOfDeparture = 'CGK';
              }
              break;
            default:
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'SEA'.trim().toLowerCase()
              ) {
                portOfDeparture = 'VNCLP';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'Land'.trim().toLowerCase()
              ) {
                portOfDeparture = 'SOTH';
              }
              if (
                item.Transport_Method.trim().toLowerCase() ===
                'AIR'.trim().toLowerCase()
              ) {
                portOfDeparture = 'SGN';
              }
              break;
          }
        }

        const queryPoA = `SELECT PortCode
                        FROM CMW_PortCode
                        WHERE TransportMethod LIKE ? AND CustomerNumber LIKE ?`;
        const result: any = await this.EIP.query(queryPoA, {
          type: QueryTypes.SELECT,
          replacements: [`%${item.Transport_Method}%`, `%${item.Customer_ID}%`],
        });
        // console.log(result[0].PortCode);
        return {
          ...item,
          Transport_Method: item.Transport_Method,
          Port_Of_Departure: portOfDeparture,
          Port_Of_Arrival: result[0]?.PortCode.trim().toUpperCase() || 'N/A',
        };
      }),
    );
    return data.flatMap((item) =>
      this.mapToCMSFormat(item, dateFrom, dateTo, dockeyCMS),
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

    return dbMap[factory] ?? null;
  }

  private mapToCMSFormat(
    item: any,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const date = item.Date;
    const factory = item.Factory;
    const factoryAddress = item.Factory_address;
    const shipmentDate = item.Shipment_Date;
    const bookingNo = item.Booking_No;
    const customerID = item.Customer_ID;
    const invoiceNumber = item.Invoice_Number;
    const transportMethod = item.Transport_Method; // AIR OR SEA
    const portOfArrival = item.Port_Of_Arrival; // DEPARTURE AND END PORT
    const portOfDeparture = item.Port_Of_Departure; // ST PORT
    const quantity = item.Quantity;
    const grossWeight = item.Gross_Weight; // ACTIVITY DATA
    const createdDate = dayjs().format('YYYY-MM-DD HH:mm:ss');

    let dockey = '';
    let transtType = '';
    let portType = '';

    switch (transportMethod.trim().toLowerCase()) {
      case 'SEA'.trim().toLowerCase():
        dockey = '3.2.01';
        transtType = '海運';
        portType = '海港';
        break;
      case 'AIR'.trim().toLowerCase():
        dockey = '3.2.02';
        transtType = '空運';
        portType = '空港';
        break;
      case 'LAND'.trim().toLowerCase():
        dockey = '3.2.03';
        transtType = '陸運';
        portType = '';
        break;
      default:
        dockey = '';
        break;
    }

    return ACTIVITY_TYPES.filter((item) => item === dockeyCMS).map(
      (activityType: ActivityType) => ({
        System: 'ERP', // DEFAULT
        Corporation: 'LAI YIH', // DEFAULT
        Factory: factory,
        Department: 'Shipping', // DEFAULT
        // DocKey:
        //   transportMethod.trim().toLowerCase() !== 'AIR'.trim().toLowerCase()
        //     ? '3.2.01'
        //     : '3.2.02', // DEFAULT
        DocKey: activityType.trim() === '5.3'.trim() ? '5.3' : dockey, // DEFAULT
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: activityType, // DEFAULT
        DataType: activityType.trim() === '3.2' ? '1' : '999', // DEFAULT
        DocType: 'Invoice', // DEFAULT 應收憑單
        UndDoc: '銷貨單', // DEFAULT
        DocFlow: '銷貨相關流程', // DEFAULT
        DocDate: dayjs(date).format('YYYY/MM/DD'),
        DocDate2: dayjs(shipmentDate).format('YYYY/MM/DD'),
        DocNo: bookingNo, // DEFAULT
        UndDocNo: '', // DEFAULT
        CustVenName: customerID,
        InvoiceNo: invoiceNumber,
        TransType: transtType,
        Departure: factoryAddress,
        Destination: portOfArrival,
        PortType: portType,
        StPort: portOfDeparture,
        ThPort: '', // DEFAULT
        EndPort: portOfArrival,
        Product: '', // DEFAULT
        Quity: quantity,
        Amount: '', // DEFAULT
        ActivityData: +Number(grossWeight).toFixed(2) || 0,
        ActivityUnit: 'Kg', // DEFAULT 噸
        Unit: '', // DEFAULT
        UnitWeight: '', // DEFAULT
        Memo: '', // DEFAULT
        CreateDateTime: createdDate,
        Creator: '',
        ActivitySource: 'Incineration (not energy recovery)',
      }),
    );
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
    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

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
