import {
  Inject,
  Injectable,
  InternalServerErrorException,
  NotFoundException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import {
  buildQuery,
  buildQueryAutoSentCMS,
  buildQueryTest,
} from 'src/helper/cat1andcat4.helper';
import {
  ACTIVITY_TYPES,
  ActivityType,
  IDataPortCodeCat1AndCat4,
} from 'src/types/cat1andcat4';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import { v4 as uuidv4 } from 'uuid';
import dayjs from 'dayjs';
import { FACTORY_LIST, FactoryCode } from 'src/helper/factory.helper';
import { PassThrough } from 'stream';
dayjs().format();

@Injectable()
export class Cat1andcat4Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
    @Inject('LHG_ERP') private readonly LHG_ERP: Sequelize,
    @Inject('LVL_ERP') private readonly LVL_ERP: Sequelize,
    @Inject('LYM_ERP') private readonly LYM_ERP: Sequelize,
  ) {}

  private getVerificationComparisonKey(row: any) {
    return [row.DocNo, row.Product]
      .map((value) => String(value ?? '').trim().toLowerCase())
      .join('|');
  }

  private async applyTaxFreeZoneOverrides(data: any[], factory: string) {
    if (data.length === 0) {
      return data;
    }

    const supplierCodes = Array.from(
      new Set(
        data
          .map((item: any) => String(item.SupplierCode ?? '').trim())
          .filter(Boolean),
      ),
    );

    if (supplierCodes.length === 0) {
      return data;
    }

    const supplierPlaceholders = supplierCodes
      .map((_, index) => `:supplierCode${index}`)
      .join(', ');
    const replacements = supplierCodes.reduce<Record<string, string>>(
      (acc, supplierCode, index) => {
        acc[`supplierCode${index}`] = supplierCode;
        return acc;
      },
      { factory },
    );

    const taxFreeZoneRows: Array<{
      SupplierID: string;
      TaxFreeZoneAddress: string;
    }> = await this.EIP.query(
      `SELECT SupplierID, TaxFreeZoneAddress
       FROM CMW_TAX_FREE_ZONE_ADDRESS
       WHERE Factory = :factory
         AND SupplierID IN (${supplierPlaceholders})`,
      {
        replacements,
        type: QueryTypes.SELECT,
      },
    );

    const taxFreeZoneMap = new Map<string, string>();
    for (const row of taxFreeZoneRows) {
      const supplierCode = String(row.SupplierID ?? '').trim();
      if (!supplierCode || taxFreeZoneMap.has(supplierCode)) {
        continue;
      }

      taxFreeZoneMap.set(supplierCode, row.TaxFreeZoneAddress);
    }

    return data.map((item: any) => {
      const supplierCode = String(item.SupplierCode ?? '').trim();
      const taxFreeZoneAddress = taxFreeZoneMap.get(supplierCode);

      if (!taxFreeZoneAddress) {
        return item;
      }

      return {
        ...item,
        TransportationMethod: 'Land',
        Departure: taxFreeZoneAddress,
      };
    });
  }

  async getDataCat1AndCat4(
    dateFrom: string,
    dateTo: string,
    factory: string,
    usage: boolean,
    unitWeight: boolean,
    weight: boolean,
    departure: boolean,
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
        return await this.getAllFactoryDataTest(
          dateFrom,
          dateTo,
          usage,
          unitWeight,
          weight,
          departure,
          page,
          limit,
          sortField,
          sortOrder,
        );
    }

    return await this.getAFactoryDataTest(
      db,
      dateFrom,
      dateTo,
      factory,
      usage,
      unitWeight,
      weight,
      departure,
      page,
      limit,
      sortField,
      sortOrder,
    );
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

  async getTaxFreeZoneAddress(
    sortField: string = 'No',
    sortOrder: string = 'asc',
  ): Promise<IDataPortCodeCat1AndCat4[]> {
    const records: IDataPortCodeCat1AndCat4[] = await this.EIP.query(
      `SELECT ROW_NUMBER() OVER (ORDER BY ID) AS No,*
        FROM CMW_TAX_FREE_ZONE_ADDRESS
        ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
        `,
      { type: QueryTypes.SELECT },
    );
    return records;
  }

  async update(
    id: string,
    updateDto: {
      TaxFreeZoneAddress: string;
    },
    factory: string,
    userid: string,
  ) {
    try {
      const query = `
          UPDATE CMW_TAX_FREE_ZONE_ADDRESS
          SET
            TaxFreeZoneAddress = :taxFreeZoneAddress,
            UpdatedBy = :updatedBy,
            UpdatedFactory = :updatedFactory,
            UpdatedAt = GETDATE()
          OUTPUT
            INSERTED.ID AS ID,
            INSERTED.TaxFreeZoneAddress AS TaxFreeZoneAddress,
            INSERTED.UpdatedBy AS UpdatedBy,
            INSERTED.UpdatedAt AS UpdatedAt
          WHERE ID = :id
        `;

      const results = await this.EIP.query(query, {
        replacements: {
          taxFreeZoneAddress: updateDto.TaxFreeZoneAddress,
          updatedBy: userid,
          updatedFactory: factory,
          id: id,
        },
        type: QueryTypes.SELECT,
      });
      if (!results || results.length === 0) {
        throw new NotFoundException('Update failed!');
      }
      return results[0];
    } catch (error) {
      if (error instanceof NotFoundException) {
        throw error;
      }
      throw new InternalServerErrorException(error.message);
    }
  }

  // private async getAFactoryData(
  //   db: Sequelize,
  //   dateFrom: string,
  //   dateTo: string,
  //   factory: string,
  //   page: number,
  //   limit: number,
  //   sortField: string,
  //   sortOrder: string,
  // ) {
  //   try {
  //     const offset = (page - 1) * limit;

  //     const { query, countQuery } = await buildQuery(
  //       dateFrom,
  //       dateTo,
  //       factory,
  //       this.EIP,
  //     );
  //     const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];
  //     const [dataResults, countResults] = await Promise.all([
  //       db.query(query, {
  //         replacements,
  //         type: QueryTypes.SELECT,
  //       }),
  //       db.query(countQuery, {
  //         replacements,
  //         type: QueryTypes.SELECT,
  //       }),
  //     ]);
  //     // console.log(dataResults);
  //     let data = dataResults;
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
  //     const total = (countResults[0] as { total: number })?.total || 0;
  //     const hasMore = offset + data.length < total;
  //     return { data, page, limit, total, hasMore };
  //   } catch (error: any) {
  //     throw new InternalServerErrorException(error);
  //   }
  // }

  // private async getAllFactoryData(
  //   dateFrom: string,
  //   dateTo: string,
  //   factory: string,
  //   page: number,
  //   limit: number,
  //   sortField: string,
  //   sortOrder: string,
  // ) {
  //   try {
  //     const offset = (page - 1) * limit;

  //     const { query, countQuery } = await buildQuery(
  //       dateFrom,
  //       dateTo,
  //       factory,
  //       this.EIP,
  //     );
  //     const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

  //     const connects = [this.LYV_ERP, this.LHG_ERP, this.LYM_ERP, this.LVL_ERP];
  //     const [dataResults, countResults] = await Promise.all([
  //       Promise.all(
  //         connects.map((conn) => {
  //           return conn.query(query, {
  //             type: QueryTypes.SELECT,
  //             replacements,
  //           });
  //         }),
  //       ),
  //       Promise.all(
  //         connects.map((conn) => {
  //           return conn.query(countQuery, {
  //             type: QueryTypes.SELECT,
  //             replacements,
  //           });
  //         }),
  //       ),
  //     ]);
  //     // console.log(dataResults,countResults);
  //     let data = dataResults.flat();
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

  async importExcelTaxFreeZoneAddress(
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
        'No',
        'Factory',
        'Supplier_ID',
        'Country',
        'Tax_Free_Zone_Address',
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
        Factory: string;
        Supplier_ID: string;
        Country: string;
        Tax_Free_Zone_Address: string;
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
          rowData?.Factory ||
          rowData?.Supplier_ID ||
          rowData?.Country ||
          rowData?.Tax_Free_Zone_Address
        ) {
          data.push(rowData);
        }
      });

      for (let item of data) {
        const id = uuidv4();
        const records: { total: number }[] = await this.EIP.query(
          `SELECT COUNT(*) total
            FROM CMW_TAX_FREE_ZONE_ADDRESS 
            WHERE Factory = ? AND SupplierID = ?`,
          {
            replacements: [item.Factory, item.Supplier_ID],
            type: QueryTypes.SELECT,
          },
        );

        if (records[0].total > 0) {
          await this.EIP.query(
            `UPDATE CMW_TAX_FREE_ZONE_ADDRESS
              SET
                Country = ?,
                TaxFreeZoneAddress = ?,
                UpdatedBy = ?,
                UpdatedFactory = ?,
                UpdatedAt = GETDATE()
              WHERE Factory = ? AND SupplierID = ?`,
            {
              replacements: [
                item.Country,
                item.Tax_Free_Zone_Address,
                userid,
                factory,
                item.Factory,
                item.Supplier_ID,
              ],
              type: QueryTypes.SELECT,
            },
          );
          updateCount++;
        } else {
          await this.EIP.query(
            `INSERT INTO CMW_TAX_FREE_ZONE_ADDRESS
              (
                ID,
                Factory,
                SupplierID,
                Country,
                TaxFreeZoneAddress,
                CreatedBy,
                CreatedFactory,
                CreatedAt
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
                item.Factory,
                item.Supplier_ID,
                item.Country,
                item.Tax_Free_Zone_Address,
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
          FROM CMW_TAX_FREE_ZONE_ADDRESS`,
        { type: QueryTypes.SELECT },
      );
      const message = `Processed successfully! Inserted: ${insertCount} records, Updated: ${updateCount} records. Total rows processed: ${data.length}.`;
      return { message, records };
    } catch (error: any) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }

  private async getAFactoryDataTest(
    db: Sequelize,
    dateFrom: string,
    dateTo: string,
    factory: string,
    usage: boolean,
    unitWeight: boolean,
    weight: boolean,
    departure: boolean,
    page: number,
    limit: number,
    sortField: string = 'No',
    sortOrder: string = 'asc',
  ) {
    try {
      const startedAt = Date.now();
      const offset = (page - 1) * limit;

      const buildQueryStartedAt = Date.now();
      const query = await buildQueryTest(
        sortField,
        sortOrder,
        factory,
        usage,
        unitWeight,
        weight,
        departure,
        this.EIP,
      );
      const buildQueryDurationMs = Date.now() - buildQueryStartedAt;

      const replacements = {
        startDate: dateFrom,
        endDate: dayjs(dateTo)
          .add(1, 'day')
          .startOf('day')
          .format('YYYY-MM-DD'),
        offset,
        limit,
      };

      const queryStartedAt = Date.now();
      let data: any[] = await db.query(query, {
        replacements,
        type: QueryTypes.SELECT,
      });
      const queryDurationMs = Date.now() - queryStartedAt;

      const overrideStartedAt = Date.now();
      data = await this.applyTaxFreeZoneOverrides(data, factory);
      const overrideDurationMs = Date.now() - overrideStartedAt;

      const total = data.length > 0 ? data[0].TotalRowsCount : 0;
      const hasMore = offset + data.length < total;
      const totalDurationMs = Date.now() - startedAt;
      console.log(
        `[Cat1AndCat4][getAFactoryDataTest] factory=${factory} page=${page} limit=${limit} rows=${data.length} total=${total} buildQueryMs=${buildQueryDurationMs} queryMs=${queryDurationMs} overrideMs=${overrideDurationMs} totalMs=${totalDurationMs}`,
      );
      return { data, page, limit, total, hasMore };
    } catch (error: any) {
      throw error;
    }
  }

  private async getAllFactoryDataTest(
    dateFrom: string,
    dateTo: string,
    usage: boolean,
    unitWeight: boolean,
    weight: boolean,
    departure: boolean,
    page: number,
    limit: number,
    sortField: string,
    sortOrder: string,
  ) {
    try {
      const startedAt = Date.now();
      const offset = (page - 1) * limit;
      const connects: { factoryName: string; conn: Sequelize }[] = [
        { factoryName: 'LYV', conn: this.LYV_ERP },
        { factoryName: 'LHG', conn: this.LHG_ERP },
        { factoryName: 'LVL', conn: this.LVL_ERP },
        { factoryName: 'LYM', conn: this.LYM_ERP },
      ];
      const replacements = {
        startDate: dateFrom,
        endDate: dayjs(dateTo)
          .add(1, 'day')
          .startOf('day')
          .format('YYYY-MM-DD'),
      };

      const results = await Promise.all(
        connects.map(async ({ factoryName, conn }) => {
          const factoryStartedAt = Date.now();
          const buildQueryStartedAt = Date.now();
          const query = await buildQueryTest(
            sortField,
            sortOrder,
            factoryName,
            usage,
            unitWeight,
            weight,
            departure,
            this.EIP,
            true,
          );
          const buildQueryDurationMs = Date.now() - buildQueryStartedAt;

          const queryStartedAt = Date.now();
          let rows: any[] = await conn.query(query, {
            type: QueryTypes.SELECT,
            replacements,
          });
          const queryDurationMs = Date.now() - queryStartedAt;

          const overrideStartedAt = Date.now();
          rows = await this.applyTaxFreeZoneOverrides(rows, factoryName);
          const overrideDurationMs = Date.now() - overrideStartedAt;
          const totalDurationMs = Date.now() - factoryStartedAt;
          console.log(
            `[Cat1AndCat4][getAllFactoryDataTest] factory=${factoryName} rows=${rows.length} buildQueryMs=${buildQueryDurationMs} queryMs=${queryDurationMs} overrideMs=${overrideDurationMs} totalMs=${totalDurationMs}`,
          );

          return rows;
        }),
      );

      let allData: any[] = results.flat();

      allData.sort((a, b) => {
        if (sortOrder === 'asc') {
          return a[sortField] > b[sortField] ? 1 : -1;
        }
        return a[sortField] < b[sortField] ? 1 : -1;
      });

      const total = allData.length;
      const data = allData
        .slice(offset, offset + limit)
        .map((row, index) => ({ ...row, No: offset + index + 1 }));
      const hasMore = offset + data.length < total;
      const totalDurationMs = Date.now() - startedAt;
      console.log(
        `[Cat1AndCat4][getAllFactoryDataTest] page=${page} limit=${limit} mergedRows=${allData.length} returnedRows=${data.length} totalMs=${totalDurationMs}`,
      );
      return { data, page, limit, total, hasMore };
    } catch (error) {}
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
  }

  async getVerificationReport(payload: {
    dateFrom: string;
    dateTo: string;
    factory: string;
    category: string;
    status: string;
    page: number;
    limit: number;
  }) {
    const {
      dateFrom,
      dateTo,
      factory,
      category,
      status,
      page = 1,
      limit = 50,
    } = payload;

    const normalizedCategory = category?.trim().toUpperCase() === 'CAT4' ? 'CAT4' : 'CAT1';
    const dockeyCMS = normalizedCategory === 'CAT4' ? '3.1' : '4.1';
    const activityType = dockeyCMS;
    const normalizedStatus = String(status ?? 'ALL').trim().toUpperCase();
    const safePage = page > 0 ? page : 1;
    const safeLimit = limit > 0 ? limit : 50;

    const previewRows = (await this.autoSentCMS(
      dateFrom,
      dateTo,
      factory,
      dockeyCMS,
    )) as any[];

    let where = 'WHERE 1=1';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ' AND Fac LIKE ?';
      replacements.push(`%${factory}%`);
    }

    where += ' AND ActivityType = ?';
    replacements.push(activityType);

    const loggingRows = (await this.EIP.query(
      `
      SELECT *
      FROM CMW_Category_1_And_4_Log
      ${where}
      ORDER BY CreatedAt DESC
      `,
      {
        type: QueryTypes.SELECT,
        replacements,
      },
    )) as any[];

    const previewKeys = new Set(
      previewRows.map((row) => this.getVerificationComparisonKey(row)),
    );
    const loggingKeys = new Set(
      loggingRows.map((row) => this.getVerificationComparisonKey(row)),
    );

    const matchedCount = previewRows.filter((row) =>
      loggingKeys.has(this.getVerificationComparisonKey(row)),
    ).length;
    const missingCount = previewRows.length - matchedCount;
    const extraCount = loggingRows.filter(
      (row) => !previewKeys.has(this.getVerificationComparisonKey(row)),
    ).length;

    const rowsWithStatus = previewRows.map((row) => {
      const isMatched = loggingKeys.has(this.getVerificationComparisonKey(row));
      return {
        ...row,
        VerificationStatus: isMatched ? 'MATCHED' : 'MISSING',
      };
    });

    const filteredRows =
      normalizedStatus === 'MATCHED'
        ? rowsWithStatus.filter((row) => row.VerificationStatus === 'MATCHED')
        : normalizedStatus === 'MISSING'
          ? rowsWithStatus.filter((row) => row.VerificationStatus === 'MISSING')
          : rowsWithStatus;

    const total = filteredRows.length;
    const offset = (safePage - 1) * safeLimit;
    const pagedRows = filteredRows.slice(offset, offset + safeLimit);

    return {
      summary: {
        previewCount: previewRows.length,
        loggingCount: loggingRows.length,
        matchedCount,
        missingCount,
        extraCount,
      },
      rows: pagedRows,
      page: safePage,
      limit: safeLimit,
      total,
      hasMore: offset + pagedRows.length < total,
      category: normalizedCategory,
      status: normalizedStatus,
    };
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

    const startedAt = Date.now();
    const buildQueryStartedAt = Date.now();
    const query = await buildQueryAutoSentCMS(factory, this.EIP);
    const buildQueryDurationMs = Date.now() - buildQueryStartedAt;

    const replacements =
      dateFrom && dateTo
        ? {
            startDate: dateFrom,
            endDate: dayjs(dateTo)
              .add(1, 'day')
              .startOf('day')
              .format('YYYY-MM-DD'),
          }
        : {};

    const queryStartedAt = Date.now();
    let data = await db.query<any>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });
    const queryDurationMs = Date.now() - queryStartedAt;

    const overrideStartedAt = Date.now();
    data = await this.applyTaxFreeZoneOverrides(data, factory);
    const overrideDurationMs = Date.now() - overrideStartedAt;

    const mapped = data.flatMap((item) =>
      this.mapToCMSFormat(item, dateFrom, dateTo, dockeyCMS),
    );
    const totalDurationMs = Date.now() - startedAt;
    console.log(
      `[Cat1AndCat4][getCMSByFactory] factory=${factory} dockeyCMS=${dockeyCMS} sourceRows=${data.length} mappedRows=${mapped.length} buildQueryMs=${buildQueryDurationMs} queryMs=${queryDurationMs} overrideMs=${overrideDurationMs} totalMs=${totalDurationMs}`,
    );

    return mapped;
  }

  private getDbByFactory(factory: FactoryCode): Sequelize | null {
    const dbMap: Record<FactoryCode, Sequelize> = {
      LYV: this.LYV_ERP,
      LHG: this.LHG_ERP,
      LVL: this.LVL_ERP,
      LYM: this.LYM_ERP,
    };

    return dbMap[factory] ?? null;
  }

  private mapToCMSFormat(
    item: any,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const factory = item.FactoryCode ?? '';
    // const docKey = `${item.MatID ?? ''}${item.ReceivedNo ?? ''}`;
    const matId = item.MatID ?? '';
    const docKey = matId.substring(0, 1);
    if (docKey.startsWith('W')) {
      return [];
    }
    const docDate = item.RKDate ? dayjs(item.RKDate).format('YYYY/MM/DD') : '';
    // const docDate2 = item.PurDate
    //   ? dayjs(item.PurDate).format('YYYY/MM/DD')
    //   : '';
    // const docNo = item.PurNo ?? '';
    const custVenName = item.SupplierCode ?? '';
    const transportationMethod = item.TransportationMethod ?? '';
    const departure = item.Departure ?? '';
    const portOfDeparture = item.PortOfDeparture ?? '';
    const portOfArrival = item.PortOfArrival ?? '';
    const destination = item.Destination ?? '';
    const activityData = item.WeightUnitkg ?? 0;
    const qtyReceive = item.QtyReceive ?? 0;
    const receivedNo = item.ReceivedNo ?? '';

    let transType = '';
    let portType = '';

    switch (transportationMethod.trim().toLowerCase()) {
      case 'SEA'.trim().toLowerCase():
      case 'SEA+LAND'.trim().toLowerCase():
        transType = '海運';
        portType = '海港';
        break;
      case 'AIR'.trim().toLowerCase():
        transType = '空運';
        portType = '空港';
        break;
      case 'LAND'.trim().toLowerCase():
        transType = '陸運';
        portType = '';
        break;
      default:
        break;
    }

    return ACTIVITY_TYPES.filter((item) => item === dockeyCMS).map(
      (activityType: ActivityType) => ({
        System: 'CMS Web',
        Corporation: 'Lai Yih',
        Factory: factory,
        Department: '',
        DocKey:
          activityType.trim() === '4.1'
            ? `${activityType.trim()}${docKey}`
            : activityType.trim(),
        SPeriodData: dayjs(dateFrom).format('YYYY/MM/DD'),
        EPeriodData: dayjs(dateTo).format('YYYY/MM/DD'),
        ActivityType: activityType.trim(),
        DataType: activityType.trim() === '3.1' ? '1' : '999',
        DocType: '入庫單',
        UndDoc: '',
        DocFlow: '',
        DocDate: docDate,
        DocDate2: docDate,
        DocNo: receivedNo.trim(),
        UndDocNo: '',
        CustVenName: custVenName,
        InvoiceNo: '',
        TransType: transType,
        Departure: departure,
        Destination: destination,
        PortType: portType,
        StPort: portOfDeparture,
        ThPort: '',
        EndPort: portOfArrival,
        Product: matId.trim(),
        Quity: qtyReceive,
        Amount: '',
        ActivityData: activityData,
        ActivityUnit: 'KG',
        Unit: '',
        UnitWeight: '',
        Memo: '',
        CreateDateTime: '',
        Creator: '',
        ActivitySource: matId.charAt(0),
      }),
    );
  }

  async exportPreviewPayload(
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    try {
      // Lấy data giống autoSentCMS
      let data: any[];
      if (factory.trim().toUpperCase() === 'ALL') {
        const results = await Promise.all(
          FACTORY_LIST.map((f) =>
            this.getCMSByFactory(f, dateFrom, dateTo, dockeyCMS),
          ),
        );
        data = results.flat();
      } else {
        data = await this.getCMSByFactory(
          factory as FactoryCode,
          dateFrom,
          dateTo,
          dockeyCMS,
        );
      }

      // Xuất Excel
      const passThrough = new PassThrough();

      const bufferPromise: Promise<Buffer> = new Promise((resolve, reject) => {
        const chunks: Buffer[] = [];
        passThrough.on('data', (chunk) => chunks.push(chunk));
        passThrough.on('end', () => resolve(Buffer.concat(chunks)));
        passThrough.on('error', reject);
      });

      const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
        stream: passThrough,
        useStyles: true,
        useSharedStrings: false,
      });

      const worksheet = workbook.addWorksheet('PreviewPayload');

      worksheet.columns = [
        { header: 'System', key: 'System', width: 15 },
        { header: 'Corporation', key: 'Corporation', width: 15 },
        { header: 'Factory', key: 'Factory', width: 15 },
        { header: 'Department', key: 'Department', width: 15 },
        { header: 'DocKey', key: 'DocKey', width: 15 },
        { header: 'SPeriodData', key: 'SPeriodData', width: 15 },
        { header: 'EPeriodData', key: 'EPeriodData', width: 15 },
        { header: 'ActivityType', key: 'ActivityType', width: 15 },
        { header: 'DataType', key: 'DataType', width: 15 },
        { header: 'DocType', key: 'DocType', width: 15 },
        { header: 'UndDoc', key: 'UndDoc', width: 15 },
        { header: 'DocFlow', key: 'DocFlow', width: 15 },
        { header: 'DocDate', key: 'DocDate', width: 15 },
        { header: 'DocDate2', key: 'DocDate2', width: 15 },
        { header: 'DocNo', key: 'DocNo', width: 15 },
        { header: 'UndDocNo', key: 'UndDocNo', width: 15 },
        { header: 'CustVenName', key: 'CustVenName', width: 25 },
        { header: 'InvoiceNo', key: 'InvoiceNo', width: 15 },
        { header: 'TransType', key: 'TransType', width: 15 },
        { header: 'Departure', key: 'Departure', width: 25 },
        { header: 'Destination', key: 'Destination', width: 15 },
        { header: 'PortType', key: 'PortType', width: 15 },
        { header: 'StPort', key: 'StPort', width: 15 },
        { header: 'ThPort', key: 'ThPort', width: 15 },
        { header: 'EndPort', key: 'EndPort', width: 15 },
        { header: 'Product', key: 'Product', width: 20 },
        { header: 'Quity', key: 'Quity', width: 12 },
        { header: 'Amount', key: 'Amount', width: 12 },
        { header: 'ActivityData', key: 'ActivityData', width: 15 },
        { header: 'ActivityUnit', key: 'ActivityUnit', width: 15 },
        { header: 'Unit', key: 'Unit', width: 10 },
        { header: 'UnitWeight', key: 'UnitWeight', width: 12 },
        { header: 'Memo', key: 'Memo', width: 20 },
        { header: 'CreateDateTime', key: 'CreateDateTime', width: 20 },
        { header: 'Creator', key: 'Creator', width: 15 },
        { header: 'ActivitySource', key: 'ActivitySource', width: 20 },
      ];

      worksheet.getRow(1).font = { bold: true };
      worksheet.getRow(1).commit();

      for (const row of data) {
        worksheet.addRow(row).commit();
      }

      await worksheet.commit();
      await workbook.commit();

      const buffer = await bufferPromise;
      return buffer;
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }
}
