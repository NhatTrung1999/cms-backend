import {
  Inject,
  Injectable,
  UnauthorizedException,
  InternalServerErrorException,
} from '@nestjs/common';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import { IDataPortCode } from 'src/types/cat9andcat12';
import * as ExcelJS from 'exceljs';

@Injectable()
export class Cat9andcat12Service {
  constructor(
    @Inject('ERP') private readonly ERP: Sequelize,
    @Inject('EIP') private readonly EIP: Sequelize,
  ) {}

  // async getData(date, userID) {

  //   console.log(date, userID);

  //   if (!userID) throw new UnauthorizedException();

  //   try {
  //     let where = ' 1 = 1 ';

  //     if (date !== '') {
  //     }

  //     const payload: any = await this.EIP.query<any>(
  //       `
  //         SELECT *
  //         FROM CMS_File_Management
  //         WHERE CreatedAt = '${userID}' ${where}
  //       `,
  //       {
  //         replacements: [],
  //         type: QueryTypes.SELECT,
  //       },
  //     );
  //     return payload;
  //   } catch (error: any) {
  //     throw new InternalServerErrorException(error);
  //   }
  // }

  // async getData(date: string, page: number = 1, limit: number = 20) {
  //   try {
  //     const offset = (page - 1) * limit;
  //     let query = `
  //       SELECT im.INV_DATE,im.INV_NO,id.STYLE_NAME,p.Qty,p.GW,im.CUSTID,
  //           'Truck' LocalLandTransportation,sb.Place_Delivery,im.TO_WHERE Country
  //           ,ISNULL(ISNULL(bg.SHPIDS,CASE WHEN (do.ShipMode='Air') AND (do.Shipmode_1 IS NULL) THEN '10 AC'
  //                         WHEN (do.ShipMode='Air Expres') AND (do.Shipmode_1 IS NULL) THEN '20 CC'
  //                                           WHEN (do.ShipMode='Ocean') AND (do.Shipmode_1 IS NULL) THEN '11 SC'
  //                                           WHEN do.ShipMode_1 IS NULL THEN ''
  //                                           ELSE do.ShipMode_1 END),y.ShipMode) TransportMethod

  //       FROM INVOICE_M im
  //       LEFT JOIN Ship_Booking sb ON sb.INV_NO = im.INV_NO
  //       LEFT JOIN INVOICE_D AS id ON id.INV_NO = im.INV_NO
  //       LEFT JOIN (SELECT INV_NO,RYNO,SUM(PAIRS)Qty,SUM (GW) GW
  //                 FROM PACKING
  //                 GROUP BY INV_NO,RYNO) p ON p.INV_NO=id.INV_NO AND p.RYNO=id.RYNO
  //       LEFT JOIN YWDD y ON y.DDBH=id.RYNO
  //       LEFT JOIN DE_ORDERM do ON do.ORDERNO=y.YSBH
  //       LEFT JOIN B_GradeOrder bg ON bg.ORDER_B=y.YSBH
  //       WHERE 1=1 ${date ? `AND CONVERT(VARCHAR, im.INV_DATE, 23) = '${date}'` : ''}
  //       ORDER BY im.INV_DATE
  //       OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY
  //     `;
  //     // const replacements: any[] = [];

  //     // if (date) {
  //     //   query += ` AND CONVERT(VARCHAR, im.INV_DATE, 23) = ?`;
  //     //   // replacements.push(date);
  //     // }

  //     // query += ` ORDER BY im.INV_DATE
  //     //           OFFSET ? ROWS FETCH NEXT ? ROWS ONLY`;
  //     // replacements.push(offset, limit)

  //     const res = await this.ERP.query(query, {
  //       replacements: [],
  //       type: QueryTypes.SELECT,
  //     });

  //     return res;
  //   } catch (error) {
  //     throw new InternalServerErrorException(error);
  //   }
  // }

  // async getDataTest(date: string, page: number = 1, limit: number = 20) {
  //   try {
  //     const offset = (page - 1) * limit;

  //     let where = 'WHERE 1=1';
  //     const replacements: any[] = [];

  //     if (date) {
  //       where += ` AND CONVERT(VARCHAR, im.INV_DATE, 23) = ?`;
  //       replacements.push(date);
  //     }

  //     const query = `SELECT im.INV_DATE,im.INV_NO,id.STYLE_NAME,p.Qty,p.GW,im.CUSTID,
  //           'Truck' LocalLandTransportation,sb.Place_Delivery,im.TO_WHERE Country
  //           ,ISNULL(ISNULL(bg.SHPIDS,CASE WHEN (do.ShipMode='Air') AND (do.Shipmode_1 IS NULL) THEN '10 AC'
  //                         WHEN (do.ShipMode='Air Expres') AND (do.Shipmode_1 IS NULL) THEN '20 CC'
  //                                           WHEN (do.ShipMode='Ocean') AND (do.Shipmode_1 IS NULL) THEN '11 SC'
  //                                           WHEN do.ShipMode_1 IS NULL THEN ''
  //                                           ELSE do.ShipMode_1 END),y.ShipMode) TransportMethod

  //       FROM INVOICE_M im
  //       LEFT JOIN Ship_Booking sb ON sb.INV_NO = im.INV_NO
  //       LEFT JOIN INVOICE_D AS id ON id.INV_NO = im.INV_NO
  //       LEFT JOIN (SELECT INV_NO,RYNO,SUM(PAIRS)Qty,SUM (GW) GW
  //                 FROM PACKING
  //                 GROUP BY INV_NO,RYNO) p ON p.INV_NO=id.INV_NO AND p.RYNO=id.RYNO
  //       LEFT JOIN YWDD y ON y.DDBH=id.RYNO
  //       LEFT JOIN DE_ORDERM do ON do.ORDERNO=y.YSBH
  //       LEFT JOIN B_GradeOrder bg ON bg.ORDER_B=y.YSBH
  //       ${where}
  //       ORDER BY im.INV_DATE
  //       OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY`;
  //     const data: any = await this.ERP.query(query, {
  //       replacements,
  //       type: QueryTypes.SELECT,
  //     });
  //     return data;
  //   } catch (error: any) {
  //     throw new InternalServerErrorException(error);
  //   }
  // }

  async getData(
    date: string,
    page: number = 1,
    limit: number = 20,
    sortField: string = 'No',
    sortOrder: string = 'asc',
  ) {
    try {
      const offset = (page - 1) * limit;

      let where = 'WHERE 1=1';
      const replacements: any[] = [];

      if (date) {
        where += ` AND CONVERT(VARCHAR, im.INV_DATE, 23) = ?`;
        replacements.push(date);
      }

      const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY im.INV_DATE) AS INT) AS [No]
                            ,im.INV_DATE          AS [Date]
                            ,im.INV_NO            AS Invoice_Number
                            ,id.STYLE_NAME        AS Article_Name
                            ,p.Qty                AS Quantity
                            ,p.GW                 AS Gross_Weight
                            ,im.CUSTID            AS Customer_ID
                            ,'Truck'              AS Local_Land_Transportation
                            ,sb.Place_Delivery    AS Port_Of_Departure
                            ,im.TO_WHERE          AS Port_Of_Arrival
                            ,CAST('0' AS INT)     AS Land_Transport_Distance
                            ,CAST('0' AS INT)     AS Sea_Transport_Distance
                            ,CAST('0' AS INT)     AS Air_Transport_Distance
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
                            )                    AS Transport_Method
                            ,CAST('0' AS INT)     AS Land_Transport_Ton_Kilometers
                            ,CAST('0' AS INT)     AS Sea_Transport_Ton_Kilometers
                            ,CAST('0' AS INT)     AS Air_Transport_Ton_Kilometers
                      FROM   INVOICE_M im
                            LEFT JOIN Ship_Booking sb
                                  ON  sb.INV_NO = im.INV_NO
                            LEFT JOIN INVOICE_D  AS id
                                  ON  id.INV_NO = im.INV_NO
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
                      ${where}
                      ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
                      OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY`;
      const data = await this.ERP.query(query, {
        replacements,
        type: QueryTypes.SELECT,
      });

      const totalResult: { total: number }[] = await this.ERP.query(
        `SELECT COUNT(im.INV_NO) total
            FROM   INVOICE_M im
                  LEFT JOIN Ship_Booking sb
                        ON  sb.INV_NO = im.INV_NO
                  LEFT JOIN INVOICE_D  AS id
                        ON  id.INV_NO = im.INV_NO
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
            ${where}`,
        { replacements, type: QueryTypes.SELECT },
      );
      const total = totalResult[0]?.total || 0;
      const hasMore = offset + data.length < total;

      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  async getPortCode(): Promise<IDataPortCode[]> {
    const records: IDataPortCode[] = await this.EIP.query(
      `SELECT *
      FROM CMS_PortCode
      ORDER BY CreatedDate
      `,
      { type: QueryTypes.SELECT },
    );
    return records;
  }

//   async importExcelPortCode(buffer: Buffer) {
//     const workbook = new ExcelJS.Workbook();
//     await workbook.xlsx.load(buffer.buffer);

//     const worksheet = workbook.worksheets[0];
//     console.log(worksheet);
//     return 'Import Excel Port Code';
//   }
}
