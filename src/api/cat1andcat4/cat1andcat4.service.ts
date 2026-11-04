import {
  Inject,
  Injectable,
  InternalServerErrorException,
  NotFoundException,
} from '@nestjs/common';
import { QueryTypes } from 'sequelize';
import { orderByClause } from 'src/helper/sql.helper';
import { Sequelize } from 'sequelize-typescript';
import {
  buildQuery,
  buildQueryAutoSentCMS,
} from 'src/helper/cat1andcat4.helper';
import {
  IDataPortCodeCat1AndCat4,
  IDataStyleAutoFill,
} from 'src/types/cat1andcat4';
import { mapToCMSFormatCat1And4 } from 'src/helper/cat1andcat4.cms';
import { CMS_PAYLOAD_COLUMNS } from 'src/helper/cms-payload.columns';
import * as ExcelJS from 'exceljs';
import { v4 as uuidv4 } from 'uuid';
import { FACTORY_LIST, FactoryCode } from 'src/helper/factory.helper';
import { getDayAfter } from 'src/helper/date.helper';
import { getOwn } from 'src/helper/object.helper';
import { memoizeByKey } from 'src/helper/lookup.helper';
import { PassThrough } from 'stream';
import {
  buildImportMessage,
  countUpserts,
  readImportRows,
} from 'src/helper/excel-import.helper';

// The factories that have a Cat1&4 database (the order is the order of the "ALL" query).
const CAT1_AND_CAT4_FACTORIES: FactoryCode[] = ['LYV', 'LHG', 'LVL', 'LYM'];

type PortCodeRow = {
  Supplier_ID: string;
  Transport_Method: string;
  Port_Code: string;
  Factory_Code: string;
};

type TaxFreeZoneRow = {
  Factory: string;
  Supplier_ID: string;
  Country: string;
  Tax_Free_Zone_Address: string;
};

@Injectable()
export class Cat1andcat4Service {
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
    @Inject('LHG_ERP') private readonly LHG_ERP: Sequelize,
    @Inject('LVL_ERP') private readonly LVL_ERP: Sequelize,
    @Inject('LYM_ERP') private readonly LYM_ERP: Sequelize,
  ) {}

  private getVerificationComparisonKey(row: {
    DocNo?: unknown;
    Product?: unknown;
  }) {
    return [row.DocNo, row.Product]
      .map((value) =>
        String(value ?? '')
          .trim()
          .toLowerCase(),
      )
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

  private async getStyleAutoFillByPrefix(prefix: string) {
    if (!prefix) {
      return null;
    }

    const [record]: Array<{ Style: string }> = await this.EIP.query(
      `SELECT TOP 1 Style
       FROM CMW_StyleAutoFill
       WHERE PrefixOfMatCode = ?`,
      {
        replacements: [prefix],
        type: QueryTypes.SELECT,
      },
    );

    return record?.Style ?? null;
  }

  private async upsertStyleAutoFill(
    item: { Prefix_of_Mat_Code: string; Style: string },
    userid: string,
    factory: string,
  ): Promise<'insert' | 'update' | 'skip'> {
    const prefix = String(item.Prefix_of_Mat_Code ?? '').trim();
    if (!prefix) {
      return 'skip';
    }

    const [record]: Array<{ total: number }> = await this.EIP.query(
      `SELECT COUNT(*) AS total
       FROM CMW_StyleAutoFill
       WHERE PrefixOfMatCode = ?`,
      {
        replacements: [prefix],
        type: QueryTypes.SELECT,
      },
    );

    if ((record?.total ?? 0) > 0) {
      await this.EIP.query(
        `UPDATE CMW_StyleAutoFill
         SET
           Style = ?,
           UpdatedBy = ?,
           UpdatedFactory = ?,
           UpdatedAt = GETDATE()
         WHERE PrefixOfMatCode = ?`,
        {
          replacements: [item.Style, userid, factory, prefix],
          type: QueryTypes.UPDATE,
        },
      );

      return 'update';
    }

    await this.EIP.query(
      `INSERT INTO CMW_StyleAutoFill
       (
         Id,
         PrefixOfMatCode,
         Style,
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
         GETDATE()
       )`,
      {
        replacements: [uuidv4(), prefix, item.Style, userid, factory],
        type: QueryTypes.INSERT,
      },
    );

    return 'insert';
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
    const trimmedFactory = factory.trim();
    const db = CAT1_AND_CAT4_FACTORIES.includes(trimmedFactory)
      ? this.getDbByFactory(trimmedFactory)
      : null;
    if (!db) {
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
        ORDER BY ${orderByClause(sortField, sortOrder)}
        `,
      { type: QueryTypes.SELECT },
    );
    return records;
  }

  async getStyleAutoFill(
    sortField: string = 'No',
    sortOrder: string = 'asc',
  ): Promise<IDataStyleAutoFill[]> {
    const records: IDataStyleAutoFill[] = await this.EIP.query(
      `SELECT *,ROW_NUMBER() OVER(ORDER BY PrefixOfMatCode) AS [No]
        FROM CMW_StyleAutoFill
        ORDER BY ${orderByClause(sortField, sortOrder)}
        `,
      { type: QueryTypes.SELECT },
    );
    return records;
  }

  async deleteStyleAutoFill(id: string) {
    try {
      const [record]: Array<{ total: number }> = await this.EIP.query(
        `SELECT COUNT(*) AS total
         FROM CMW_StyleAutoFill
         WHERE Id = ?`,
        {
          replacements: [id],
          type: QueryTypes.SELECT,
        },
      );

      if ((record?.total ?? 0) === 0) {
        throw new NotFoundException('Delete failed!');
      }

      await this.EIP.query(`DELETE FROM CMW_StyleAutoFill WHERE Id = ?`, {
        replacements: [id],
        type: QueryTypes.DELETE,
      });

      return { message: 'Deleted successfully!', Id: id };
    } catch (error: any) {
      if (error instanceof NotFoundException) {
        throw error;
      }
      throw new InternalServerErrorException(error.message);
    }
  }

  async getTaxFreeZoneAddress(
    sortField: string = 'No',
    sortOrder: string = 'asc',
  ): Promise<IDataPortCodeCat1AndCat4[]> {
    const records: IDataPortCodeCat1AndCat4[] = await this.EIP.query(
      `SELECT ROW_NUMBER() OVER (ORDER BY ID) AS No,*
        FROM CMW_TAX_FREE_ZONE_ADDRESS
        ORDER BY ${orderByClause(sortField, sortOrder)}
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
    } catch (error: any) {
      if (error instanceof NotFoundException) {
        throw error;
      }
      throw new InternalServerErrorException(error.message);
    }
  }

  async importExcelPortCode(
    file: Express.Multer.File,
    userid: string,
    factory: string,
  ) {
    try {
      const data = await readImportRows<PortCodeRow>(
        file,
        ['Supplier_ID', 'Transport_Method', 'Port_Code', 'Factory_Code'],
        (rowData) =>
          rowData?.Supplier_ID &&
          rowData?.Transport_Method &&
          rowData?.Port_Code &&
          rowData?.Factory_Code,
      );

      const { insertCount, updateCount } = await countUpserts(data, (item) =>
        this.upsertPortCode(item, userid, factory),
      );
      const records = await this.EIP.query(
        `SELECT *
          FROM CMW_PortCode_Cat1_4`,
        { type: QueryTypes.SELECT },
      );
      const message = buildImportMessage(insertCount, updateCount, data.length);
      return { message, records };
    } catch (error: any) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }

  private async upsertPortCode(
    item: PortCodeRow,
    userid: string,
    factory: string,
  ): Promise<'insert' | 'update'> {
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
          type: QueryTypes.UPDATE,
        },
      );
      return 'update';
    }

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
    return 'insert';
  }

  async importExcelStyleAutoFill(
    file: Express.Multer.File,
    userid: string,
    factory: string,
  ) {
    try {
      const data = await readImportRows<{
        No: string;
        Prefix_of_Mat_Code: string;
        Style: string;
      }>(
        file,
        ['No', 'Prefix_of_Mat_Code', 'Style'],
        (rowData) =>
          rowData?.No && rowData?.Prefix_of_Mat_Code && rowData?.Style,
      );

      const { insertCount, updateCount } = await countUpserts(data, (item) =>
        this.upsertStyleAutoFill(item, userid, factory),
      );
      const records = await this.EIP.query(
        `SELECT *, ROW_NUMBER() OVER(ORDER BY PrefixOfMatCode) AS [No]
          FROM CMW_StyleAutoFill`,
        { type: QueryTypes.SELECT },
      );
      const message = buildImportMessage(insertCount, updateCount, data.length);
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
      const data = await readImportRows<TaxFreeZoneRow>(
        file,
        ['No', 'Factory', 'Supplier_ID', 'Country', 'Tax_Free_Zone_Address'],
        (rowData) =>
          rowData?.Factory ||
          rowData?.Supplier_ID ||
          rowData?.Country ||
          rowData?.Tax_Free_Zone_Address,
      );

      const { insertCount, updateCount } = await countUpserts(data, (item) =>
        this.upsertTaxFreeZone(item, userid, factory),
      );
      const records = await this.EIP.query(
        `SELECT *
          FROM CMW_TAX_FREE_ZONE_ADDRESS`,
        { type: QueryTypes.SELECT },
      );
      const message = buildImportMessage(insertCount, updateCount, data.length);
      return { message, records };
    } catch (error: any) {
      throw new InternalServerErrorException(`${error.message}`);
    }
  }

  private async upsertTaxFreeZone(
    item: TaxFreeZoneRow,
    userid: string,
    factory: string,
  ): Promise<'insert' | 'update'> {
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
          type: QueryTypes.UPDATE,
        },
      );
      return 'update';
    }

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
    return 'insert';
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
    const offset = (page - 1) * limit;

    const query = await buildQuery(
      sortField,
      sortOrder,
      factory,
      usage,
      unitWeight,
      weight,
      departure,
      this.EIP,
    );

    const replacements = {
      startDate: dateFrom,
      endDate: getDayAfter(dateTo),
      offset,
      limit,
    };

    let data: any[] = await db.query(query, {
      replacements,
      type: QueryTypes.SELECT,
    });

    data = await this.applyTaxFreeZoneOverrides(data, factory);

    const styleOfPrefix = memoizeByKey((prefix) =>
      this.getStyleAutoFillByPrefix(prefix),
    );
    data = await Promise.all(
      data.map(async (item: any) => {
        const prefix = String(item?.MatID ?? '')
          .trim()
          .charAt(0);
        const styleAutoFill = await styleOfPrefix(prefix);
        return {
          ...item,
          Style:
            (item.Style === '' ? styleAutoFill : item?.Style) ||
            (item?.Style ?? styleAutoFill ?? 'N/A'),
        };
      }),
    );

    const total = data.length > 0 ? data[0].TotalRowsCount : 0;
    const hasMore = offset + data.length < total;

    return { data, page, limit, total, hasMore };
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
      const offset = (page - 1) * limit;
      const connects = CAT1_AND_CAT4_FACTORIES.map((factoryName) => ({
        factoryName,
        conn: this.getDbByFactory(factoryName)!,
      }));
      const replacements = {
        startDate: dateFrom,
        endDate: getDayAfter(dateTo),
      };

      const results = await Promise.all(
        connects.map(async ({ factoryName, conn }) => {
          const query = await buildQuery(
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
          let rows: any[] = await conn.query(query, {
            type: QueryTypes.SELECT,
            replacements,
          });

          rows = await this.applyTaxFreeZoneOverrides(rows, factoryName);

          return rows;
        }),
      );

      const allData: any[] = results.flat();

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
      return { data, page, limit, total, hasMore };
    } catch (error) {
      // The error is only logged: as before, the caller gets `undefined` (an empty response).
      console.error('getAllFactoryDataTest failed:', error);
    }
  }

  async autoSentCMS(
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    try {
      return this.getCMSRows(factory, dateFrom, dateTo, dockeyCMS);
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  async getVerificationReport(payload: {
    previewDateFrom: string;
    previewDateTo: string;
    loggingDateFrom: string;
    loggingDateTo: string;
    factory: string;
    category: string;
    status: string;
    page: number;
    limit: number;
  }) {
    const {
      previewDateFrom,
      previewDateTo,
      loggingDateFrom,
      loggingDateTo,
      factory,
      category,
      status,
      page = 1,
      limit = 50,
    } = payload;

    const normalizedCategory =
      category?.trim().toUpperCase() === 'CAT4' ? 'CAT4' : 'CAT1';
    const dockeyCMS = normalizedCategory === 'CAT4' ? '3.1' : '4.1';
    const activityType = dockeyCMS;
    const normalizedStatus = String(status ?? 'ALL')
      .trim()
      .toUpperCase();
    const safePage = page > 0 ? page : 1;
    const safeLimit = limit > 0 ? limit : 50;

    const previewRows = await this.autoSentCMS(
      previewDateFrom,
      previewDateTo,
      factory,
      dockeyCMS,
    );

    let where = 'WHERE 1=1';
    const replacements: any[] = [];

    if (loggingDateFrom && loggingDateTo) {
      where += ' AND CONVERT(VARCHAR, CreatedAt, 23) BETWEEN ? AND ?';
      replacements.push(loggingDateFrom, loggingDateTo);
    }

    if (factory) {
      where += ' AND Fac LIKE ?';
      replacements.push(`%${factory}%`);
    }

    where += ' AND ActivityType = ?';
    replacements.push(activityType);

    const loggingRows = await this.EIP.query(
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
    );

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
    const extraRows = loggingRows.filter(
      (row) => !previewKeys.has(this.getVerificationComparisonKey(row)),
    );
    const extraCount = extraRows.length;

    const previewRowsWithStatus = previewRows.map((row) => {
      const isMatched = loggingKeys.has(this.getVerificationComparisonKey(row));
      return {
        ...row,
        VerificationStatus: isMatched ? 'MATCHED' : 'MISSING',
      };
    });

    const extraRowsWithStatus = extraRows.map((row) => ({
      ...row,
      VerificationStatus: 'EXTRA',
    }));

    const rowsWithStatus = [...previewRowsWithStatus, ...extraRowsWithStatus];

    const filteredRows = ['MATCHED', 'MISSING', 'EXTRA'].includes(
      normalizedStatus,
    )
      ? rowsWithStatus.filter(
          (row) => row.VerificationStatus === normalizedStatus,
        )
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

  /**
   * CMS payload rows of every factory (in FACTORY_LIST order) when factory is 'ALL', otherwise of that one factory.
   * Not async on purpose: a bad `factory` throws right here, inside the caller's try.
   */
  private getCMSRows(
    factory: string,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ): Promise<any[]> {
    if (factory.trim().toUpperCase() === 'ALL') {
      return Promise.all(
        FACTORY_LIST.map((f) =>
          this.getCMSByFactory(f, dateFrom, dateTo, dockeyCMS),
        ),
      ).then((results) => results.flat());
    }

    return this.getCMSByFactory(factory, dateFrom, dateTo, dockeyCMS);
  }

  private async getCMSByFactory(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDbByFactory(factory);
    if (!db) return [];

    const query = await buildQueryAutoSentCMS(factory, this.EIP);

    const replacements =
      dateFrom && dateTo
        ? {
            startDate: dateFrom,
            endDate: getDayAfter(dateTo),
          }
        : {};

    let data = await db.query<any>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    data = await this.applyTaxFreeZoneOverrides(data, factory);

    const mapped = data.flatMap((item) =>
      mapToCMSFormatCat1And4(item, dateFrom, dateTo, dockeyCMS),
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

    return getOwn(dbMap, factory) ?? null;
  }

  async exportPreviewPayload(
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    try {
      // Lấy data giống autoSentCMS
      const data = await this.getCMSRows(factory, dateFrom, dateTo, dockeyCMS);

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

      worksheet.columns = CMS_PAYLOAD_COLUMNS;

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
