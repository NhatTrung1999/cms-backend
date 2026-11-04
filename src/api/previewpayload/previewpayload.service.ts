import { Inject, Injectable } from '@nestjs/common';
import { EventsGateway } from '../events/events.gateway';
import { v4 as uuidv4 } from 'uuid';
import { Sequelize } from 'sequelize-typescript';
import { QueryTypes } from 'sequelize';
import { ConfigService } from '@nestjs/config';
import * as ExcelJS from 'exceljs';
import { FACTORY_LIST, FactoryCode } from 'src/helper/factory.helper';
import { buildQueryAutoSentCMS } from 'src/helper/cat1andcat4.helper';
import { buildReplacements } from 'src/helper/date.helper';
import {
  buildQueryAutoSentCmsJAZ,
  buildQueryAutoSentCmsJZS,
  buildQueryAutoSentCmsLHG,
  buildQueryAutoSentCmsLVL,
  buildQueryAutoSentCmsLYM,
  buildQueryAutoSentCmsLYV,
} from 'src/helper/cat7.cms-queries';
import { BuildQueryFn } from 'src/types/cat7';
import { buildCat6PreviewQuery } from './previewpayload.queries';
import { buildCat6Filter } from 'src/helper/cat6.queries';
import { getDayAfter } from 'src/helper/date.helper';
import { getOwn } from 'src/helper/object.helper';
import {
  getExcelFilePath,
  getExcelRootFolder,
  insertFileRecord,
  runFileExport,
} from 'src/helper/file-export.helper';
import { applyTaxFreeZoneToRows } from 'src/helper/cat1andcat4.taxfree';
import { mapToCMSFormatCat7 } from './previewpayload.mappers';
import { mapToCMSFormatCat1And4 } from 'src/helper/cat1andcat4.cms';
import { getCat5CMSRows } from 'src/helper/cat5.cms';
import { getCat9AndCat12CMSRows } from 'src/helper/cat9andcat12.cms';
import {
  mapCat6RowToAccommodationPreviewPayload,
  mapCat6RowToPreviewPayload,
  toCat6ActiveRows,
} from './previewpayload.cat6';
import {
  PREVIEW_CAT1_AND_CAT4_COLUMNS,
  PREVIEW_CAT5_COLUMNS,
  PREVIEW_CAT6_ACCOMMODATION_COLUMNS,
  PREVIEW_CAT6_COLUMNS,
  PREVIEW_CAT7_COLUMNS,
  PREVIEW_CAT9_AND_CAT12_COLUMNS,
} from './previewpayload.columns';

type DbSystem = 'ERP' | 'WMS' | 'HRIS';

@Injectable()
export class PreviewpayloadService {
  private rootFolder: string;
  constructor(
    @Inject('EIP') private readonly EIP: Sequelize,
    private readonly eventsGateway: EventsGateway,
    private configService: ConfigService,
    //cat1&4, cat9&12
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
    @Inject('LHG_ERP') private readonly LHG_ERP: Sequelize,
    @Inject('LVL_ERP') private readonly LVL_ERP: Sequelize,
    @Inject('LYM_ERP') private readonly LYM_ERP: Sequelize,
    @Inject('LYF_ERP') private readonly LYF_ERP: Sequelize,
    @Inject('JAZ_ERP') private readonly JAZ_ERP: Sequelize,
    @Inject('JZS_ERP') private readonly JZS_ERP: Sequelize,
    //cat5
    @Inject('LYV_WMS') private readonly LYV_WMS: Sequelize,
    @Inject('LHG_WMS') private readonly LHG_WMS: Sequelize,
    @Inject('LYM_WMS') private readonly LYM_WMS: Sequelize,
    @Inject('LVL_WMS') private readonly LVL_WMS: Sequelize,
    @Inject('JAZ_WMS') private readonly JAZ_WMS: Sequelize,
    @Inject('JZS_WMS') private readonly JZS_WMS: Sequelize,
    //cat7
    @Inject('LYV_HRIS') private readonly LYV_HRIS: Sequelize,
    @Inject('LHG_HRIS') private readonly LHG_HRIS: Sequelize,
    @Inject('LVL_HRIS') private readonly LVL_HRIS: Sequelize,
    @Inject('LYM_HRIS') private readonly LYM_HRIS: Sequelize,
    @Inject('JAZ_HRIS') private readonly JAZ_HRIS: Sequelize,
    @Inject('JZS_HRIS') private readonly JZS_HRIS: Sequelize,
    @Inject('UOF') private readonly UOF: Sequelize,
  ) {
    this.rootFolder = getExcelRootFolder(this.configService);
  }

  /** Databases of the factories, per system (ERP: cat1&4, cat9&12 / WMS: cat5 / HRIS: cat7). */
  private get databases(): Record<DbSystem, Record<string, Sequelize>> {
    return {
      ERP: {
        LYV: this.LYV_ERP,
        LHG: this.LHG_ERP,
        LVL: this.LVL_ERP,
        LYM: this.LYM_ERP,
        LYF: this.LYF_ERP,
        JAZ: this.JAZ_ERP,
        JZS: this.JZS_ERP,
      },
      WMS: {
        LYV: this.LYV_WMS,
        LHG: this.LHG_WMS,
        LVL: this.LVL_WMS,
        LYM: this.LYM_WMS,
        JAZ: this.JAZ_WMS,
        JZS: this.JZS_WMS,
      },
      HRIS: {
        LYV: this.LYV_HRIS,
        LHG: this.LHG_HRIS,
        LVL: this.LVL_HRIS,
        LYM: this.LYM_HRIS,
        JAZ: this.JAZ_HRIS,
        JZS: this.JZS_HRIS,
      },
    };
  }

  private getDb(system: DbSystem, factory: FactoryCode): Sequelize | null {
    return getOwn(this.databases[system], factory) ?? null;
  }

  /** Rows of every factory (in FACTORY_LIST order) when factory is 'ALL', otherwise of that one factory. */
  private async collectByFactory(
    factory: string,
    fetchOne: (factory: FactoryCode) => Promise<any[]>,
  ): Promise<any[]> {
    if (factory.trim().toUpperCase() === 'ALL') {
      const results = await Promise.all(FACTORY_LIST.map((f) => fetchOne(f)));
      return results.flat();
    }
    return fetchOne(factory);
  }

  private writePreviewSheet(
    sheet: ExcelJS.Worksheet,
    columns: Partial<ExcelJS.Column>[],
    rows: any[],
  ) {
    sheet.columns = columns;

    sheet.getRow(1).font = { bold: true };

    for (const row of rows) {
      sheet.addRow(row);
    }
  }

  private readonly buildQueryMap: Record<FactoryCode, BuildQueryFn> = {
    LYV: buildQueryAutoSentCmsLYV,
    LHG: buildQueryAutoSentCmsLHG,
    LVL: buildQueryAutoSentCmsLVL,
    LYM: buildQueryAutoSentCmsLYM,
    JAZ: buildQueryAutoSentCmsJAZ,
    JZS: buildQueryAutoSentCmsJZS,
  };

  private async getCMSByFactory(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDb('ERP', factory);
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

    data = await applyTaxFreeZoneToRows(data, factory, this.EIP);

    return data.flatMap((item) =>
      mapToCMSFormatCat1And4(item, dateFrom, dateTo, dockeyCMS),
    );
  }

  // cat9 & 12
  private async getCMSByFactoryCat9(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    ry: string,
    dockeyCMS: string,
  ) {
    const db = this.getDb('ERP', factory);
    if (!db) return [];

    return getCat9AndCat12CMSRows(
      db,
      this.EIP,
      factory,
      dateFrom,
      dateTo,
      ry,
      dockeyCMS,
    );
  }

  private async getCMSByFactoryCat5(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDb('WMS', factory);
    if (!db) return [];

    return getCat5CMSRows(db, this.EIP, factory, dateFrom, dateTo, dockeyCMS);
  }

  private async getCMSByFactoryCat7(
    factory: FactoryCode,
    dateFrom: string,
    dateTo: string,
    dockeyCMS: string,
  ) {
    const db = this.getDb('HRIS', factory);
    if (!db) return [];

    const buildQuery = getOwn(this.buildQueryMap, factory);
    if (!buildQuery) return [];

    const { dateToExclusive, hasDate } = buildReplacements(dateFrom, dateTo);

    const query = await buildQuery(dateFrom, dateTo, this.EIP);

    const replacements = hasDate ? [dateFrom, dateToExclusive] : [];

    const data = await db.query<any>(query, {
      type: QueryTypes.SELECT,
      replacements,
    });

    return data.flatMap((item) =>
      mapToCMSFormatCat7(item, dateFrom, dateTo, dockeyCMS),
    );
  }

  async exportExcelCat6(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    const { where, replacements: finalReplacements } = buildCat6Filter(
      dateFrom,
      dateTo,
      factory && factory.trim().toUpperCase() !== 'ALL' ? factory : null,
    );

    const query = buildCat6PreviewQuery(where);

    const rawRows = await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements: finalReplacements,
    });

    const transformedRows = toCat6ActiveRows(rawRows);

    if (dockeyCMS === '3.5.5') {
      const accommodationPayloadRows = transformedRows
        .filter((row) => Number(row.Number_of_nights_stayed ?? 0) > 0)
        .map((row) =>
          mapCat6RowToAccommodationPreviewPayload(
            row,
            factory,
            dateFrom,
            dateTo,
          ),
        );

      this.writePreviewSheet(
        sheet,
        PREVIEW_CAT6_ACCOMMODATION_COLUMNS,
        accommodationPayloadRows,
      );
      return;
    }

    const payloadRows = transformedRows
      .filter((row) => Number(row.Number_of_nights_stayed ?? 0) > 0)
      .flatMap((row) =>
        mapCat6RowToPreviewPayload(row, factory, dateFrom, dateTo),
      );

    this.writePreviewSheet(sheet, PREVIEW_CAT6_COLUMNS, payloadRows);
  }
  // Preview Excel
  async previewExcel(
    module: string,
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
    userID: string,
    dockeyCMS: string,
  ) {
    const id = uuidv4();
    const moduleName = module.trim();
    const fileName = `Preview-Payload-${factory}-${moduleName}-${dateFrom}-${dateTo}-${dockeyCMS}.xlsx`;
    const filePath = getExcelFilePath(this.rootFolder, moduleName, fileName);

    await insertFileRecord(this.EIP, {
      id,
      module: moduleName,
      fileName,
      filePath,
      factory,
      userID,
    });

    // Runs in the background: nothing awaits it, so an error that escapes must be caught here
    // (an unhandled rejection would stop the whole backend).
    this.createFileExcel(
      id,
      moduleName,
      filePath,
      dateFrom,
      dateTo,
      factory,
      ry,
      dockeyCMS,
    ).catch((error) => console.error('createFileExcel failed:', error));

    return true;
  }

  async createFileExcel(
    id: string,
    module: string,
    filePath: string,
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
    dockeyCMS: string,
  ) {
    await runFileExport({
      eip: this.EIP,
      gateway: this.eventsGateway,
      id,
      filePath,
      sheetName: module,
      fill: async (sheet) => {
        switch (module.trim().toLowerCase()) {
          case 'cat1andcat4':
            await this.exportExcelCat1And4(
              sheet,
              dateFrom,
              dateTo,
              factory,
              dockeyCMS,
            );
            break;
          case 'cat5':
            await this.exportExcelCat5(
              sheet,
              dateFrom,
              dateTo,
              factory,
              dockeyCMS,
            );
            break;
          case 'cat6':
            await this.exportExcelCat6(
              sheet,
              dateFrom,
              dateTo,
              factory,
              dockeyCMS,
            );
            break;
          case 'cat7':
            await this.exportExcelCat7(
              sheet,
              dateFrom,
              dateTo,
              factory,
              dockeyCMS,
            );
            break;
          case 'cat9andcat12':
            await this.exportExcelCat9And12(
              sheet,
              dateFrom,
              dateTo,
              factory,
              ry,
              dockeyCMS,
            );
            break;
          default:
            throw new Error(`Module "${module}" không hợp lệ`);
        }
      },
    });
  }

  async exportExcelCat1And4(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    const data = await this.collectByFactory(factory, (f) =>
      this.getCMSByFactory(f, dateFrom, dateTo, dockeyCMS),
    );

    // Set columns giống exportPreviewPayload
    this.writePreviewSheet(sheet, PREVIEW_CAT1_AND_CAT4_COLUMNS, data);
  }

  async exportExcelCat5(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    const data = await this.collectByFactory(factory, (f) =>
      this.getCMSByFactoryCat5(f, dateFrom, dateTo, dockeyCMS),
    );

    this.writePreviewSheet(sheet, PREVIEW_CAT5_COLUMNS, data);
  }

  async exportExcelCat7(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dockeyCMS: string,
  ) {
    const data = await this.collectByFactory(factory, (f) =>
      this.getCMSByFactoryCat7(f, dateFrom, dateTo, dockeyCMS),
    );

    this.writePreviewSheet(sheet, PREVIEW_CAT7_COLUMNS, data);
  }

  async exportExcelCat9And12(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
    dockeyCMS: string,
  ) {
    const data = await this.collectByFactory(factory, (f) =>
      this.getCMSByFactoryCat9(f, dateFrom, dateTo, ry, dockeyCMS),
    );

    this.writePreviewSheet(sheet, PREVIEW_CAT9_AND_CAT12_COLUMNS, data);
  }
}
