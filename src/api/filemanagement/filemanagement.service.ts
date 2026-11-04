import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { orderByClause } from 'src/helper/sql.helper';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import * as ExcelJS from 'exceljs';
import { v4 as uuidv4 } from 'uuid';
import { EventsGateway } from '../events/events.gateway';
import { getADataExcelFactoryCat5 } from 'src/helper/cat5.helper';
import {
  findFileRecord,
  getExcelFilePath,
  getExcelRootFolder,
  insertFileRecord,
  runFileExport,
} from 'src/helper/file-export.helper';
import { getADataExcelFactoryCat9AndCat12 } from 'src/helper/cat9andcat12.helper';
import { getDayAfter } from 'src/helper/date.helper';
import {
  buildCat9AndCat12Filter,
  EXPORT_DATE_COLUMN,
} from 'src/helper/cat9andcat12.queries';
import { getADataExcelFactoryCat1AndCat4 } from 'src/helper/cat1andcat4.helper';
import { buildQuery, getADataExcelFactoryCat7 } from 'src/helper/cat7.helper';
import {
  buildQueryCustomExport,
  getADataCustomExport,
} from 'src/helper/cat7.customexport';

import { buildQuery as buildQueryCat1AndCat4 } from 'src/helper/cat1andcat4.helper';
import {
  autoFitColumnsAndAddBorders,
  fillSheet,
} from 'src/helper/excel.helper';
import { addPortsToRows } from 'src/helper/shipment.helper';
import { buildCat6Filter } from 'src/helper/cat6.queries';
import { applyTaxFreeZoneToRows } from 'src/helper/cat1andcat4.taxfree';
import {
  buildCat6LegsFromPlaces,
  extractCat6PlacesAndTransports,
  toCat6ExportRows,
} from 'src/helper/cat6.helper';
import {
  buildAllFactoryCat5Query,
  buildAllFactoryCat9AndCat12Query,
  buildCat6TripQuery,
} from './filemanagement.queries';
import {
  CAT1_AND_CAT4_COLUMNS,
  CAT5_COLUMNS,
  CAT6_ACCOMMODATION_COLUMNS,
  CAT6_COLUMNS,
  CAT7_COLUMNS,
  CAT9_AND_CAT12_COLUMNS,
} from './filemanagement.columns';

type FactoryConnection = { factoryName: string; conn: Sequelize };

@Injectable()
export class FilemanagementService {
  private rootFolder: string;
  constructor(
    private configService: ConfigService,
    @Inject('EIP') private readonly EIP: Sequelize,
    // cat 9 & 12 and cat 1 & 4
    @Inject('LYV_ERP') private readonly LYV_ERP: Sequelize,
    @Inject('LHG_ERP') private readonly LHG_ERP: Sequelize,
    @Inject('LVL_ERP') private readonly LVL_ERP: Sequelize,
    @Inject('LYM_ERP') private readonly LYM_ERP: Sequelize,
    @Inject('LYF_ERP') private readonly LYF_ERP: Sequelize,
    @Inject('JAZ_ERP') private readonly JAZ_ERP: Sequelize,
    @Inject('JZS_ERP') private readonly JZS_ERP: Sequelize,
    private readonly eventsGateway: EventsGateway,
    //cat 5
    @Inject('LYV_WMS') private readonly LYV_WMS: Sequelize,
    @Inject('LHG_WMS') private readonly LHG_WMS: Sequelize,
    @Inject('LYM_WMS') private readonly LYM_WMS: Sequelize,
    @Inject('LVL_WMS') private readonly LVL_WMS: Sequelize,
    @Inject('JAZ_WMS') private readonly JAZ_WMS: Sequelize,
    @Inject('JZS_WMS') private readonly JZS_WMS: Sequelize,
    // cat 7
    @Inject('LYV_HRIS') private readonly LYV_HRIS: Sequelize,
    @Inject('LHG_HRIS') private readonly LHG_HRIS: Sequelize,
    @Inject('LVL_HRIS') private readonly LVL_HRIS: Sequelize,
    @Inject('LYM_HRIS') private readonly LYM_HRIS: Sequelize,
    @Inject('JAZ_HRIS') private readonly JAZ_HRIS: Sequelize,
    @Inject('JZS_HRIS') private readonly JZS_HRIS: Sequelize,
    // cat 6
    @Inject('UOF') private readonly UOF: Sequelize,
  ) {
    this.rootFolder = getExcelRootFolder(this.configService);
  }

  // Connections per export, in the order the "all factories" export reads them
  // (this order is also the row order of the Excel file).
  private get erpConnections(): FactoryConnection[] {
    return [
      { factoryName: 'LYV', conn: this.LYV_ERP },
      { factoryName: 'LHG', conn: this.LHG_ERP },
      { factoryName: 'LYM', conn: this.LYM_ERP },
      { factoryName: 'LVL', conn: this.LVL_ERP },
      { factoryName: 'LYF', conn: this.LYF_ERP },
      { factoryName: 'JAZ', conn: this.JAZ_ERP },
      { factoryName: 'JZS', conn: this.JZS_ERP },
    ];
  }

  /** Cat 1 & 4 only exists in four factories. */
  private get cat1And4Connections(): FactoryConnection[] {
    return [
      { factoryName: 'LYV', conn: this.LYV_ERP },
      { factoryName: 'LHG', conn: this.LHG_ERP },
      { factoryName: 'LVL', conn: this.LVL_ERP },
      { factoryName: 'LYM', conn: this.LYM_ERP },
    ];
  }

  private get wmsConnections(): FactoryConnection[] {
    return [
      { factoryName: 'LYV', conn: this.LYV_WMS },
      { factoryName: 'LHG', conn: this.LHG_WMS },
      { factoryName: 'LYM', conn: this.LYM_WMS },
      { factoryName: 'LVL', conn: this.LVL_WMS },
      { factoryName: 'JAZ', conn: this.JAZ_WMS },
      { factoryName: 'JZS', conn: this.JZS_WMS },
    ];
  }

  private get hrisConnections(): FactoryConnection[] {
    return [
      { factoryName: 'LYV', conn: this.LYV_HRIS },
      { factoryName: 'LHG', conn: this.LHG_HRIS },
      { factoryName: 'LVL', conn: this.LVL_HRIS },
      { factoryName: 'LYM', conn: this.LYM_HRIS },
      { factoryName: 'JAZ', conn: this.JAZ_HRIS },
      { factoryName: 'JZS', conn: this.JZS_HRIS },
    ];
  }

  /** The connection of one factory, or undefined when it is not in the list (e.g. 'ALL'). */
  private findConnection(connections: FactoryConnection[], factory: string) {
    const factoryName = factory.trim();
    return connections.find((item) => item.factoryName === factoryName)?.conn;
  }

  async getData(
    module: string,
    file_name: string,
    userID: string,
    sortField: string = 'CreatedDate',
    sortOrder: string = 'asc',
  ) {
    try {
      let where = ' AND 1 = 1 ';
      const replacements: string[] = [userID];

      if (module !== '') {
        where += ` AND Module LIKE ?`;
        replacements.push(`%${module}%`);
      }

      if (file_name !== '') {
        where += ` AND File_Name LIKE ?`;
        replacements.push(`%${file_name}%`);
      }

      const payload: any = await this.EIP.query<any>(
        `
          SELECT *
          FROM CMW_File_Management
          WHERE CreatedAt = ? ${where}
          ORDER BY ${orderByClause(sortField, sortOrder)}
        `,
        {
          replacements,
          type: QueryTypes.SELECT,
        },
      );
      return payload;
    } catch (error: any) {
      throw new InternalServerErrorException(error);
    }
  }

  async getFileById(id: string) {
    const payload: any = await this.EIP.query(
      `SELECT *
        FROM CMW_File_Management
        WHERE ID = ?`,
      {
        replacements: [id],
        type: QueryTypes.SELECT,
      },
    );
    if (payload.length === 0) return false;
    return payload[0];
  }

  async generateFileExcel(
    module: string,
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
    userID: string,
    fields?: string[],
    usage?: boolean,
    unitWeight?: boolean,
    weight?: boolean,
    departure?: boolean,
  ) {
    const id = uuidv4();
    const fileName = `${factory}-${module}-${dateFrom}-${dateTo}.xlsx`;
    const filePath = getExcelFilePath(this.rootFolder, module, fileName);

    const statusWaiting = await this.fileExcelWaiting(
      id,
      module,
      fileName,
      filePath,
      factory,
      userID,
    );
    if (statusWaiting.length === 0) return false;

    // Runs in the background: nothing awaits it, so an error that escapes must be caught here
    // (an unhandled rejection would stop the whole backend).
    this.createFileExcel(
      id,
      module,
      filePath,
      dateFrom,
      dateTo,
      factory,
      ry,
      fields,
      usage,
      unitWeight,
      weight,
      departure,
    ).catch((error) => console.error('createFileExcel failed:', error));

    return true;
  }

  async fileExcelWaiting(
    id: string,
    module: string,
    fileName: string,
    filePath: string,
    factory: string,
    userID: string,
  ) {
    try {
      await insertFileRecord(this.EIP, {
        id,
        module,
        fileName,
        filePath,
        factory,
        userID,
      });
      return await findFileRecord(this.EIP, id);
    } catch (error: any) {
      throw new Error(error);
    }
  }

  async createFileExcel(
    id: string,
    module: string,
    filePath: string,
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
    fields?: string[],
    usage?: boolean,
    unitWeight?: boolean,
    weight?: boolean,
    departure?: boolean,
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
            await this.fileExcelCat1AndCat4(
              sheet,
              dateFrom,
              dateTo,
              factory,
              usage,
              unitWeight,
              weight,
              departure,
            );
            break;
          case 'cat5':
            await this.fileExcelCat5(sheet, dateFrom, dateTo, factory);
            break;
          case 'cat6':
          case 'cat6businesstravel':
            await this.fileExcelCat6(sheet, dateFrom, dateTo, factory);
            break;
          case 'cat6accommodation':
            await this.fileExcelCat6Accommodation(
              sheet,
              dateFrom,
              dateTo,
              factory,
            );
            break;
          case 'cat7':
            await this.fileExcelCat7(sheet, dateFrom, dateTo, factory);
            break;
          case 'cat9andcat12':
            await this.fileExcelCat9AndCat12(
              sheet,
              dateFrom,
              dateTo,
              factory,
              ry,
            );
            break;
          case 'customexport':
            await this.fileCustomExport(
              sheet,
              dateFrom,
              dateTo,
              factory,
              fields,
            );
            break;
          default:
            throw new Error(`Unknown excel file module: ${module}`);
        }
      },
    });
  }

  private async getCat6TransformedRows(
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    // Like the preview: 'ALL' (or no factory) means no factory filter.
    const { where, replacements: finalReplacements } = buildCat6Filter(
      dateFrom,
      dateTo,
      factory && factory.trim().toUpperCase() !== 'ALL' ? factory : null,
    );

    const query = buildCat6TripQuery(where);

    const rawRows = await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements: finalReplacements,
    });

    return toCat6ExportRows(rawRows);
  }

  async fileExcelCat6(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    const transformed = await this.getCat6TransformedRows(
      dateFrom,
      dateTo,
      factory,
    );
    const withNightsStayed = transformed.filter(
      (row) => Number(row.Number_of_nights_stayed ?? 0) > 0,
    );

    const legRows = withNightsStayed.flatMap((row) => {
      const { Routes, Document_Number, ...tripFields } = row;
      const documentNumberBase = Document_Number ?? '';
      const { places, transports } = extractCat6PlacesAndTransports(Routes);
      const legs = buildCat6LegsFromPlaces(places, transports);

      if (legs.length === 0) {
        return [
          {
            ...tripFields,
            Document_Number: `${documentNumberBase}-1`,
            Departure: '',
            Destination: '',
            Transport: '',
          },
        ];
      }

      return legs.map((leg, index) => ({
        ...tripFields,
        Document_Number: `${documentNumberBase}-${index + 1}`,
        Departure: leg.dep,
        Destination: leg.dest,
        Transport: leg.transType,
      }));
    });

    fillSheet(sheet, CAT6_COLUMNS, legRows);
  }

  async fileExcelCat6Accommodation(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    const transformed = await this.getCat6TransformedRows(
      dateFrom,
      dateTo,
      factory,
    );
    const withNightsStayed = transformed.filter(
      (row) => Number(row.Number_of_nights_stayed ?? 0) > 0,
    );

    fillSheet(sheet, CAT6_ACCOMMODATION_COLUMNS, withNightsStayed);
  }

  async fileExcelCat9AndCat12(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry: string,
  ) {
    const db = this.findConnection(this.erpConnections, factory);
    if (!db) {
      return await this.getAllExcelFactoryCat9AndCat12Data(
        sheet,
        dateFrom,
        dateTo,
        factory,
        ry,
      );
    }

    await getADataExcelFactoryCat9AndCat12(
      sheet,
      db,
      dateFrom,
      dateTo,
      factory,
      ry,
      this.EIP,
    );
  }

  private async getAllExcelFactoryCat9AndCat12Data(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    ry,
    dbEIP?: Sequelize,
  ) {
    const { where, where1, replacements } = buildCat9AndCat12Filter(
      dateFrom,
      dateTo,
      ry,
      EXPORT_DATE_COLUMN,
    );
    const dataResults = await Promise.all(
      this.erpConnections.map(({ conn, factoryName }) => {
        const query = buildAllFactoryCat9AndCat12Query(
          factoryName,
          where,
          where1,
        );
        return conn.query(query, {
          type: QueryTypes.SELECT,
          replacements,
        });
      }),
    );

    const allData = dataResults.flat();

    let data: any = allData.map((item, index) => ({ ...item, No: index + 1 }));

    data = await addPortsToRows(data, dbEIP);

    fillSheet(sheet, CAT9_AND_CAT12_COLUMNS, data);
  }

  async fileExcelCat5(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    const db = this.findConnection(this.wmsConnections, factory);
    if (!db) {
      return await this.getAllDataExcelFactoryCat5(sheet, dateFrom, dateTo);
    }

    await getADataExcelFactoryCat5(sheet, db, dateFrom, dateTo);
  }

  private async getAllDataExcelFactoryCat5(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
  ) {
    let where = 'WHERE 1=1 AND dwc.DISABLED = 0 AND dwo.QUANTITY<>0';
    const replacements: any[] = [];
    if (dateTo && dateFrom) {
      where += ` AND dwo.WASTE_DATE BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }
    const query = buildAllFactoryCat5Query(where);
    const dataResults = await Promise.all(
      this.wmsConnections.map(({ conn }) => {
        return conn.query(query, {
          type: QueryTypes.SELECT,
          replacements,
        });
      }),
    );
    const data = dataResults.flat();
    fillSheet(sheet, CAT5_COLUMNS, data);
  }

  async fileExcelCat1AndCat4(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    usage?: boolean,
    unitWeight?: boolean,
    weight?: boolean,
    departure?: boolean,
  ) {
    const db = this.findConnection(this.cat1And4Connections, factory);
    if (!db) {
      return await this.getAllDataExcelFactoryCat1AndCat4(
        sheet,
        dateFrom,
        dateTo,
        usage,
        unitWeight,
        weight,
        departure,
      );
    }

    await getADataExcelFactoryCat1AndCat4(
      sheet,
      db,
      dateFrom,
      dateTo,
      factory,
      usage,
      unitWeight,
      weight,
      departure,
      this.EIP,
    );
  }

  private async getAllDataExcelFactoryCat1AndCat4(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    usage?: boolean,
    unitWeight?: boolean,
    weight?: boolean,
    departure?: boolean,
  ) {
    const replacements = {
      startDate: dateFrom,
      endDate: getDayAfter(dateTo),
      offset: 1,
      limit: 20,
    };
    const dataResults = await Promise.all(
      this.cat1And4Connections.map(async ({ factoryName, conn }) => {
        const query = await buildQueryCat1AndCat4(
          'No',
          'asc',
          factoryName,
          usage,
          unitWeight,
          weight,
          departure,
          this.EIP,
        );
        let rows: any[] = await conn.query(query, {
          type: QueryTypes.SELECT,
          replacements,
        });

        rows = await applyTaxFreeZoneToRows(rows, factoryName, this.EIP);

        return rows;
      }),
    );
    const data = dataResults.flat();
    fillSheet(sheet, CAT1_AND_CAT4_COLUMNS, data);
  }

  async fileExcelCat7(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    const db = this.findConnection(this.hrisConnections, factory);
    if (!db) {
      return await this.getAllDataExcelFactoryCat7(sheet, dateFrom, dateTo);
    }

    await getADataExcelFactoryCat7(sheet, db, dateFrom, dateTo, factory);
  }

  private async getAllDataExcelFactoryCat7(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
  ) {
    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

    const dataResults = await Promise.all(
      this.hrisConnections.map(({ factoryName, conn }) => {
        const { query } = buildQuery(factoryName, dateFrom, dateTo);
        return conn.query(query, { type: QueryTypes.SELECT, replacements });
      }),
    );

    const data = dataResults.flat();
    fillSheet(sheet, CAT7_COLUMNS, data);
  }

  private async fileCustomExport(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    fields?: string[],
  ) {
    const db = this.findConnection(this.hrisConnections, factory);
    if (!db) {
      return await this.getAllDataCustomExport(sheet, dateFrom, dateTo, fields);
    }
    await getADataCustomExport(sheet, db, dateFrom, dateTo, factory, fields);
  }

  private async getAllDataCustomExport(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    fields?: string[],
  ) {
    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];
    const dataResults = await Promise.all(
      this.hrisConnections.map(({ factoryName, conn }) => {
        const { query } = buildQueryCustomExport(dateFrom, dateTo, factoryName);
        return conn.query(query, { type: QueryTypes.SELECT, replacements });
      }),
    );

    const allData = dataResults.flat();

    const data = allData.map((item, index) => ({ ...item, No: index + 1 }));

    if (!data || data.length === 0) return;

    const allKeys = Object.keys(data[0]);
    const selectedFields = fields && fields.length > 0 ? fields : allKeys;

    sheet.columns = selectedFields.map((key) => ({
      header: key.replace(/_/g, ' '),
      key,
    }));

    data.forEach((item) => {
      const rowData: Record<string, any> = {};
      selectedFields.forEach((key) => {
        rowData[key] = item[key] ?? '';
      });
      sheet.addRow(rowData);
    });

    autoFitColumnsAndAddBorders(sheet);
  }
}
