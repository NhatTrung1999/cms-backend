import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import { buildQuery } from 'src/helper/cat1andcat4.helper';
import { IDataPortCodeCat1AndCat4 } from 'src/types/cat1andcat4';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import { v4 as uuidv4 } from 'uuid';

@Injectable()
export class Cat1andcat4Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
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
    switch (factory.trim()) {
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
        return await this.getAllFactoryData(
          dateFrom,
          dateTo,
          factory,
          page,
          limit,
          sortField,
          sortOrder,
        );
    }

    return await this.getAFactoryData(
      db,
      dateFrom,
      dateTo,
      factory,
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

  async getPortCode(
    sortField: string = 'SupplierID',
    sortOrder: string = 'asc',
  ): Promise<IDataPortCodeCat1AndCat4[]> {
    const records: IDataPortCodeCat1AndCat4[] = await this.EIP.query(
      `SELECT *
        FROM CMW_PortCode_Cat1_4
        ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
        `,
      { type: QueryTypes.SELECT },
    );
    return records;
  }

  private async getAFactoryData(
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

      const { query, countQuery } = await buildQuery(
        dateFrom,
        dateTo,
        factory,
        this.EIP,
      );
      const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];
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

  private async getAllFactoryData(
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

      const { query, countQuery } = await buildQuery(
        dateFrom,
        dateTo,
        factory,
        this.EIP,
      );
      const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

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
      const headers: string[] = [];
      headerRow.eachCell((cell) => {
        if (cell.value) {
          headers.push(cell.value.toString().trim().replace(/\s+/g, '_'));
        }
      });

      const requiredHeaders = [
        'Supplier_ID',
        'Transport_Method',
        'Port_Code',
        'Factory_Code',
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
        Supplier_ID: string;
        Transport_Method: string;
        Port_Code: string;
        Factory_Code: string;
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
          }
        });
        if (
          rowData?.Supplier_ID &&
          rowData?.Transport_Method &&
          rowData?.Port_Code &&
          rowData?.Factory_Code
        ) {
          data.push(rowData);
        }
      });

      for (let item of data) {
        const id = uuidv4();
        const records: { total: number }[] = await this.EIP.query(
          `SELECT COUNT(*) total
            FROM CMW_PortCode_Cat1_4
            WHERE SupplierID = ? AND TransportMethod = ? AND PortCode = ? AND FactoryCode = ?`,
          {
            replacements: [
              item.Supplier_ID,
              item.Transport_Method,
              item.Port_Code,
              item.Factory_Code,
            ],
            type: QueryTypes.SELECT,
          },
        );

        if (records[0].total > 0) {
          await this.EIP.query(
            `UPDATE CMW_PortCode_Cat1_4
            SET
                  PortCode = ?,
                  FactoryCode = ?,
                  UpdatedBy = ?,
                  UpdatedFactory = ?,
                  UpdatedDate = GETDATE()
            WHERE SupplierID = ? AND TransportMethod = ?`,
            {
              replacements: [
                item.Port_Code,
                item.Factory_Code,
                userid,
                factory,
                item.Supplier_ID,
                item.Transport_Method,
              ],
              type: QueryTypes.SELECT,
            },
          );
          updateCount++;
        } else {
          await this.EIP.query(
            `INSERT INTO CMW_PortCode_Cat1_4
                    (
                          Id,
                          SupplierID,
                          PortCode,
                          FactoryCode,
                          TransportMethod,
                          CreatedBy,
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
                          ?,
                          GETDATE()
                    )`,
            {
              replacements: [
                id,
                item.Supplier_ID,
                item.Port_Code,
                item.Factory_Code,
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
          FROM CMW_PortCode_Cat1_4`,
        { type: QueryTypes.SELECT },
      );
      const message = `Processed successfully! Inserted: ${insertCount} records, Updated: ${updateCount} records. Total rows processed: ${data.length}.`;
      return { message, records };
    } catch (error: any) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }
}
