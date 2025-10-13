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
import * as fs from 'fs';
import { v4 as uuidv4 } from 'uuid';

@Injectable()
export class Cat9andcat12Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
    @Inject('LHG_ERP') private readonly LHG_ERP: Sequelize,
    @Inject('LVL_ERP') private readonly LVL_ERP: Sequelize,
    @Inject('LYM_ERP') private readonly LYM_ERP: Sequelize,
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
      default:
        return 'data All';
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
      FROM CMS_PortCode
      ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
      `,
      { type: QueryTypes.SELECT },
    );
    return records;
  }

  async getDataFactory(
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

      let where = 'WHERE 1=1';
      const replacements: any[] = [];

      if (dateFrom && dateTo) {
        where += ` AND CONVERT(VARCHAR, im.INV_DATE, 23) BETWEEN ? AND ?`;
        replacements.push(dateFrom, dateTo);
      }

      const query = `SELECT CAST(ROW_NUMBER() OVER(ORDER BY im.INV_DATE) AS INT) AS [No]
                            ,im.INV_DATE          AS [Date]
                            ,im.INV_NO            AS Invoice_Number
                            ,id.STYLE_NAME        AS Article_Name
                            ,p.Qty                AS Quantity
                            ,p.GW                 AS Gross_Weight
                            ,im.CUSTID            AS Customer_ID
                            ,'Truck'              AS Local_Land_Transportation
                            ,CASE 
                                  WHEN ISNULL(LEFT(LTRIM(RTRIM(im.INV_NO)) ,2) ,'')='' THEN ''
                                  WHEN LEFT(LTRIM(RTRIM(im.INV_NO)) ,2)<>'AM' THEN 'VNCLP'
                                  ELSE 'MMRGN'
                            END                  AS Port_Of_Departure
                            ,pc.PortCode          AS Port_Of_Arrival
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
                            LEFT JOIN ${factory === 'LYM' || factory === 'LVL' ? 'EIPDB' : 'EIP'}.EIP.dbo.CMS_PortCode pc
                                  ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                      ${where}`;
      // ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
      // OFFSET ${offset} ROWS FETCH NEXT ${limit} ROWS ONLY`;
      // const data = await db.query(query, {
      //   replacements,
      //   type: QueryTypes.SELECT,
      // });

      const countQuery = `SELECT COUNT(im.INV_NO) total
                          FROM   INVOICE_M im
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
                                LEFT JOIN ${factory === 'LYM' || factory === 'LVL' ? 'EIPDB' : 'EIP'}.EIP.dbo.CMS_PortCode pc
                                      ON  pc.CustomerNumber COLLATE Chinese_Taiwan_Stroke_CI_AS = im.CUSTID
                            ${where}`;

      // const totalResult: { total: number }[] = await db.query(
      // `SELECT COUNT(im.INV_NO) total
      //     FROM   INVOICE_M im
      //           LEFT JOIN Ship_Booking sb
      //                 ON  sb.INV_NO = im.INV_NO
      //           LEFT JOIN INVOICE_D  AS id
      //                 ON  id.INV_NO = im.INV_NO
      //           LEFT JOIN (
      //                     SELECT INV_NO
      //                           ,RYNO
      //                           ,SUM(PAIRS)     Qty
      //                           ,SUM(GW)        GW
      //                     FROM   PACKING
      //                     GROUP BY
      //                           INV_NO
      //                           ,RYNO
      //                 ) p
      //                 ON  p.INV_NO = id.INV_NO
      //                     AND p.RYNO = id.RYNO
      //           LEFT JOIN YWDD y
      //                 ON  y.DDBH = id.RYNO
      //           LEFT JOIN DE_ORDERM do
      //                 ON  do.ORDERNO = y.YSBH
      //           LEFT JOIN B_GradeOrder bg
      //                 ON  bg.ORDER_B = y.YSBH
      //       ${where}`,
      //   { replacements, type: QueryTypes.SELECT },
      // );

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
      const headers: string[] = [];
      headerRow.eachCell((cell) => {
        if (cell.value) {
          headers.push(cell.value.toString().trim().replace(/\s+/g, '_'));
        }
      });

      const requiredHeaders = ['Customer_Number', 'Port_Code'];
      const missingHeaders = requiredHeaders.filter(
        (h) => !headers.includes(h),
      );

      if (missingHeaders.length > 0) {
        throw new Error(
          `Excel file format is invalid! Missing columns: ${missingHeaders.join(', ')}`,
        );
      }

      const data: { Customer_Number: string; Port_Code: string }[] = [];

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
        if (rowData?.Customer_Number && rowData?.Port_Code) {
          data.push(rowData);
        }
      });

      for (let item of data) {
        const id = uuidv4();
        const records: { total: number }[] = await this.EIP.query(
          `SELECT COUNT(*) total
            FROM CMS_PortCode
            WHERE CustomerNumber = ?`,
          {
            replacements: [item.Customer_Number],
            type: QueryTypes.SELECT,
          },
        );

        if (records[0].total > 0) {
          await this.EIP.query(
            `UPDATE CMS_PortCode
            SET
                  PortCode = ?,
                  UpdatedAt = ?,
                  UpdatedFactory = ?,
                  UpdatedDate = GETDATE()
            WHERE CustomerNumber = ?`,
            {
              replacements: [
                item.Port_Code,
                'admin',
                'LYV',
                item.Customer_Number,
              ],
              type: QueryTypes.SELECT,
            },
          );
          updateCount++;
        } else {
          await this.EIP.query(
            `INSERT INTO CMS_PortCode
                    (
                          Id,
                          CustomerNumber,
                          PortCode,
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
                          GETDATE()
                    )`,
            {
              replacements: [
                id,
                item.Customer_Number,
                item.Port_Code,
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
          FROM CMS_PortCode`,
        { type: QueryTypes.SELECT },
      );
      const message = `Processed successfully! Inserted: ${insertCount} records, Updated: ${updateCount} records. Total rows processed: ${data.length}.`;
      return { message, records };
    } catch (error: any) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }
}
