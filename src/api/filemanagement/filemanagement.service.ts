import {
  Inject,
  Injectable,
  InternalServerErrorException,
} from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';
import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { v4 as uuidv4 } from 'uuid';
import { EventsGateway } from '../events/events.gateway';
import { getADataExcelFactoryCat5 } from 'src/helper/cat5.helper';
import { getADataExcelFactoryCat9AndCat12 } from 'src/helper/cat9andcat12.helper';
import { getADataExcelFactoryCat1AndCat4 } from 'src/helper/cat1andcat4.helper';
import {
  buildQuery,
  buildQueryCustomExport,
  getADataCustomExport,
  getADataExcelFactoryCat7,
} from 'src/helper/cat7.helper';
import { getCat6AirportName } from 'src/helper/cat6-airport.helper';
import {
  appendCat6DocumentSuffix,
  getCat6AssistedIds,
} from 'src/helper/cat6-assisted.helper';

import { buildQueryTest as buildQueryCat1AndCat4 } from 'src/helper/cat1andcat4.helper';
import dayjs from 'dayjs';

type Cat6RouteItem = {
  AddressName: string;
  Transport: string;
  AddressDetail: string;
  isAirport: boolean;
  From: string;
  To: string;
};

type Cat6AccommodationItem = {
  id?: string;
  isSameAsAbove?: boolean;
  type?: string;
  nights?: number | string;
  addressID?: string;
};

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
    this.rootFolder = this.configService.get(
      'EXCEL_STORAGE_PATH',
      'D:/FileExcel',
    );
    if (!fs.existsSync(this.rootFolder)) {
      fs.mkdirSync(this.rootFolder, { recursive: true });
    }
  }

  private parseJsonArray<T = unknown>(value: unknown): T[] {
    if (Array.isArray(value)) return value as T[];
    if (typeof value !== 'string' || !value.trim()) return [];
    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? (parsed as T[]) : [];
    } catch {
      return [];
    }
  }

  private formatCat6TripType(type: unknown): string | null {
    if (!type || typeof type !== 'string') return null;
    const t = type.toLowerCase().trim();
    if (t === 'round' || t === 'roundtrip' || t === 'round trip')
      return 'Round trip';
    return 'One-way';
  }

  private formatCat6BusinessTripType(type: unknown): string | null {
    if (!type || typeof type !== 'string') return null;
    const map: Record<string, string> = {
      factory_domestic: 'Domestic business trip within the group',
      outside_domestic: 'Domestic business trip to third party entities',
      factory_oversea: 'Overseas business trip within the group',
      outside_oversea: 'Overseas business trip to third party entities',
    };
    return map[type.toLowerCase()] ?? type;
  }

  private formatCat6PlacesAndTransports(routesValue: unknown) {
    const routes = this.parseJsonArray<Cat6RouteItem>(routesValue);
    const places: string[] = [];
    const transports: string[] = [];

    const pushPlace = (value?: string) => {
      const place = getCat6AirportName(value);
      if (!place) return;
      if (places[places.length - 1] === place) return;
      places.push(place);
    };

    const isFlightRoute = (route: Cat6RouteItem) => {
      const transport = route.Transport?.trim().toLowerCase() ?? '';
      return route.isAirport === true || transport === 'flight';
    };

    let index = 0;
    while (index < routes.length) {
      const route = routes[index];
      const transport = route.Transport?.trim() ?? '';
      const nextRoute = routes[index + 1];

      if (isFlightRoute(route)) {
        const flightStart = route.From?.trim() ?? '';
        let flightEnd = route.To?.trim() ?? '';
        let cursor = index;

        while (cursor + 1 < routes.length && isFlightRoute(routes[cursor + 1])) {
          cursor += 1;
          flightEnd = routes[cursor].To?.trim() ?? flightEnd;
        }

        pushPlace(flightStart);
        pushPlace(flightEnd);
        transports.push('Flight');
        index = cursor + 1;
        continue;
      }

      if (!transport && places.length === 0 && nextRoute && isFlightRoute(nextRoute)) {
        pushPlace(nextRoute.From);
        index += 1;
        continue;
      }

      pushPlace(route.AddressDetail ?? route.AddressName);
      if (transport) {
        transports.push(transport);
      }
      index += 1;
    }

    const transportLimit = Math.max(places.length - 1, 0);
    const normalizedTransports = transports.slice(0, transportLimit);
    while (normalizedTransports.length < transportLimit) {
      normalizedTransports.push('');
    }

    return {
      ...places.reduce<Record<string, string>>((acc, place, index) => {
        acc[`Place${index + 1}`] = place;
        return acc;
      }, {}),
      ...normalizedTransports.reduce<Record<string, string>>(
        (acc, transport, index) => {
          acc[`Transport_${index + 1}`] = transport;
          return acc;
        },
        {},
      ),
    };
  }

  private getCat6AccommodationNights(accommodationValue: unknown) {
    const accommodations =
      this.parseJsonArray<Cat6AccommodationItem>(accommodationValue);
    return accommodations.reduce((sum, item) => {
      const nights = Number(item?.nights ?? 0);
      return sum + (Number.isFinite(nights) ? nights : 0);
    }, 0);
  }

  private getCat6NumberOfPeople(assistedIdsValue: unknown) {
    return getCat6AssistedIds(assistedIdsValue).length;
  }

  private expandCat6RowsByAssistedIds(
    row: Record<string, any>,
  ): Record<string, any>[] {
    const assistedIds = getCat6AssistedIds(row.AssisstedIDs);

    if (assistedIds.length === 0) {
      return [
        {
          ...row,
          DOC_NBR: row.DOC_NBR ?? '',
          EffectiveStaffID: row.UserCreate ?? '',
          Number_of_People: 0,
        },
      ];
    }

    if (assistedIds.length === 1) {
      return [
        {
          ...row,
          DOC_NBR: row.DOC_NBR ?? '',
          EffectiveStaffID: assistedIds[0],
          Number_of_People: 1,
        },
      ];
    }

    return assistedIds.map((assistedId, index) => ({
      ...row,
      DOC_NBR: appendCat6DocumentSuffix(String(row.DOC_NBR ?? ''), index),
      EffectiveStaffID: assistedId,
      Number_of_People: 1,
    }));
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

      if (module !== '') {
        where += ` AND Module LIKE '%${module}%'`;
      }

      if (file_name !== '') {
        where += ` AND File_Name LIKE '%${file_name}%'`;
      }

      const payload: any = await this.EIP.query<any>(
        `
          SELECT *
          FROM CMW_File_Management
          WHERE CreatedAt = '${userID}' ${where}
          ORDER BY ${sortField} ${sortOrder === 'asc' ? 'ASC' : 'DESC'}
        `,
        {
          replacements: [],
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
    userID: string,
    fields?: string[],
    usage?: boolean,
    unitWeight?: boolean,
    weight?: boolean,
    departure?: boolean,
  ) {
    const id = uuidv4();
    const fileName = `${factory}-${module}-${dateFrom}-${dateTo}.xlsx`;
    const folder = path.join(this.rootFolder, module);
    if (!fs.existsSync(folder)) {
      fs.mkdirSync(folder, { recursive: true });
    }
    const filePath = path.join(folder, fileName);

    await this.createFileExcel(
      id,
      module,
      filePath,
      dateFrom,
      dateTo,
      factory,
      fields,
      usage,
      unitWeight,
      weight,
      departure,
    );

    let statusWaiting = await this.fileExcelWaiting(
      id,
      module,
      fileName,
      filePath,
      factory,
      userID,
    );
    if (statusWaiting.length === 0) return false;
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
      await this.EIP.query(
        `INSERT INTO CMW_File_Management
          (
            ID,
            Module,
            [File_Name],
            [Path],
            [Status],
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
            ?,
            ?
          )`,
        {
          replacements: [
            id,
            module,
            fileName,
            filePath,
            '0',
            userID,
            factory,
            new Date().toISOString(),
          ],
          type: QueryTypes.INSERT,
        },
      );
      const result = await this.EIP.query(
        `SELECT * FROM CMW_File_Management WHERE ID = ? ORDER BY CreatedDate`,
        { replacements: [id], type: QueryTypes.SELECT },
      );
      return result;
    } catch (error) {
      throw new Error(error);
    }
  }

  async fileExcelDone(id: string) {
    try {
      await this.EIP.query(
        `UPDATE CMW_File_Management
          SET
            [Status] = '1'
          WHERE ID = ?`,
        { replacements: [id], type: QueryTypes.UPDATE },
      );
      const result = await this.EIP.query(
        `SELECT * FROM CMW_File_Management WHERE ID = ? ORDER BY CreatedDate`,
        { replacements: [id], type: QueryTypes.SELECT },
      );
      return result;
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
    fields?: string[],
    usage?: boolean,
    unitWeight?: boolean,
    weight?: boolean,
    departure?: boolean,
  ) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(module);

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
        await this.fileExcelCat6(sheet, dateFrom, dateTo, factory);
        break;
      case 'cat7':
        await this.fileExcelCat7(sheet, dateFrom, dateTo, factory);
        break;
      case 'cat9andcat12':
        await this.fileExcelCat9AndCat12(sheet, dateFrom, dateTo, factory);
        break;
      case 'customexport':
        await this.fileCustomExport(sheet, dateFrom, dateTo, factory, fields);
        break;
      default:
        console.log(`Unknown excel file module: ${module}`);
        break;
    }

    await workbook.xlsx
      .writeFile(filePath)
      .then(async () => {
        setTimeout(async () => {
          const statusDone = await this.fileExcelDone(id);
          if (statusDone.length) {
            return await this.eventsGateway.broadcastEvent(
              'file-excel-done',
              'Export file excel success!',
            );
          }
        }, 60000);
      })
      .catch((error: any) => {
        this.eventsGateway.broadcastEvent(
          'file-excel-error',
          'Error export file excel!',
        );
    });
  }

  async fileExcelCat6(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    let where = `WHERE  chb.BPMStatus = 'F'
                        AND ISNULL(chb.Factory_User ,'')<>''`;
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR ,chb.CreatedAt ,23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo);
    }

    if (factory) {
      where += ` AND chb.Factory_User LIKE ?`;
      replacements.push(`%${factory}%`);
    }

    const query = `
      SELECT chb.*
                          --,twt.Current_doc
                          ,COALESCE(
                              vwd.GROUP_NAME
                              ,(
                                  SELECT TOP 1 vwd2.GROUP_NAME
                                  FROM   TB_EB_USER teu2
                                          OUTER APPLY (
                                      SELECT teed2.GROUP_ID
                                      FROM   TB_EB_EMPL_DEP AS teed2
                                      WHERE  teed2.USER_GUID = teu2.USER_GUID
                                              AND teed2.ORDERS = 0
                                  ) teed2
                                  LEFT JOIN vwDepartment_Factory vwd2
                                              ON  vwd2.GROUP_ID = teed2.GROUP_ID
                                  WHERE  teu2.ACCOUNT = chb.Factory_User+chb.UserCreate
                                          AND vwd2.GROUP_NAME IS NOT NULL
                              )
                              ,(
                                  SELECT TOP 1 vwd3.GROUP_NAME
                                  FROM   TB_EB_USER teu3
                                          OUTER APPLY (
                                      SELECT teed3.GROUP_ID
                                      FROM   TB_EB_EMPL_DEP AS teed3
                                      WHERE  teed3.USER_GUID = teu3.USER_GUID
                                              AND teed3.ORDERS = 0
                                  ) teed3
                                  LEFT JOIN vwDepartment_Factory vwd3
                                              ON  vwd3.GROUP_ID = teed3.GROUP_ID
                                  WHERE  teu3.ACCOUNT = chb.Departure+chb.UserCreate
                                          AND vwd3.GROUP_NAME IS NOT NULL
                              )
                              ,(
                                  SELECT TOP 1 vwd4.GROUP_NAME
                                  FROM   CDS_FMEval_Employee cfe
                                          JOIN TB_EB_USER teu4
                                              ON  teu4.ACCOUNT = cfe.BPMAccount
                                          OUTER APPLY (
                                      SELECT teed4.GROUP_ID
                                      FROM   TB_EB_EMPL_DEP AS teed4
                                      WHERE  teed4.USER_GUID = teu4.USER_GUID
                                              AND teed4.ORDERS = 0
                                  ) teed4
                                  LEFT JOIN vwDepartment_Factory vwd4
                                              ON  vwd4.GROUP_ID = teed4.GROUP_ID
                                  WHERE  cfe.EmpID = chb.UserCreate
                                          AND vwd4.GROUP_NAME IS NOT NULL
                              )
                          )                         AS Dept
                          ,COUNT(chb.TripID) OVER()  AS TotalRow
                    FROM   CDS_HRBUSS_BusTripData    AS chb
                          LEFT JOIN TB_EB_USER      AS teu
                                ON  (
                                        teu.ACCOUNT LIKE chb.Factory_User+'[0-9][0-9][0-9][0-9][0-9]'
                                        AND teu.ACCOUNT LIKE '%'+chb.UserCreate+'%'
                                    )
                                    OR (
                                        teu.ACCOUNT NOT LIKE chb.Factory_User+'[0-9][0-9][0-9][0-9][0-9]'
                                        AND teu.ACCOUNT=chb.UserCreate
                                    )
                          OUTER APPLY (
                        SELECT teed2.GROUP_ID
                        FROM   TB_EB_EMPL_DEP AS teed2
                        WHERE  teed2.USER_GUID = teu.USER_GUID
                              AND teed2.ORDERS = 0
                    )                                AS teed
                    LEFT JOIN vwDepartment_Factory   AS vwd
                                ON  vwd.GROUP_ID = teed.GROUP_ID
                          LEFT JOIN tb_wkf_task     AS twt
                                ON  twt.DOC_NBR = chb.DOC_NBR
                                    AND (
                                            twt.DOC_NBR LIKE 'LYV-HR-BT%'
                                            OR twt.DOC_NBR LIKE 'LHG-SUGG%'
                                            OR twt.DOC_NBR LIKE 'LVL-HR-BTF%'
                                            OR twt.DOC_NBR LIKE 'LVL-ODBT%'
                                            OR twt.DOC_NBR LIKE 'LYM-HR-BT%'
                                        )
      ${where}
      ORDER BY CreatedAt ASC
    `;

    const rawRows = (await this.UOF.query(query, {
      type: QueryTypes.SELECT,
      replacements,
    })) as Record<string, any>[];

    const transformed = rawRows.flatMap((row) =>
      this.expandCat6RowsByAssistedIds(row).map((expandedRow) => ({
        Document_Date: expandedRow.CreatedAt
          ? dayjs(expandedRow.CreatedAt).format('YYYY-MM-DD')
          : '',
        Document_Number: expandedRow.DOC_NBR ?? '',
        Staff_ID: expandedRow.EffectiveStaffID ?? '',
        Dept: expandedRow.Dept ?? expandedRow.Department ?? '',
        Round_trip_One_way: this.formatCat6TripType(expandedRow.TypeTravel) ?? '',
        Start_Time: expandedRow.DateStart
          ? dayjs(expandedRow.DateStart).format('YYYY-MM-DD')
          : '',
        End_Time: expandedRow.DateEnd
          ? dayjs(expandedRow.DateEnd).format('YYYY-MM-DD')
          : '',
        Business_Trip_Type:
          this.formatCat6BusinessTripType(expandedRow.Factory) ?? '',
        ...this.formatCat6PlacesAndTransports(expandedRow.Routes),
        Number_of_nights_stayed: this.getCat6AccommodationNights(
          expandedRow.Accommodation,
        ),
        Number_of_People: expandedRow.Number_of_People,
      })),
    );

    const placeCount = Math.max(
      1,
      ...transformed.map(
        (row) => Object.keys(row).filter((key) => /^Place\d+$/.test(key)).length,
      ),
    );
    const transportCount = Math.max(
      0,
      placeCount - 1,
      ...transformed.map(
        (row) =>
          Object.keys(row).filter((key) => /^Transport_\d+$/.test(key)).length,
      ),
    );

    sheet.columns = [
      { header: 'Document Date', key: 'Document_Date' },
      { header: 'Document Number', key: 'Document_Number' },
      { header: 'Staff ID', key: 'Staff_ID' },
      { header: 'Dept', key: 'Dept' },
      { header: 'Round trip / One way', key: 'Round_trip_One_way' },
      { header: 'Start Time', key: 'Start_Time' },
      { header: 'End Time', key: 'End_Time' },
      { header: 'Business Trip Type', key: 'Business_Trip_Type' },
      ...Array.from({ length: placeCount }, (_, index) => ({
        header: `Place ${index + 1}`,
        key: `Place${index + 1}`,
      })),
      ...Array.from({ length: transportCount }, (_, index) => ({
        header: `Transport ${index + 1}`,
        key: `Transport_${index + 1}`,
      })),
      { header: 'Number of nights stayed', key: 'Number_of_nights_stayed' },
      { header: 'Number of people', key: 'Number_of_People' },
    ];

    transformed.forEach((item) => sheet.addRow(item));
    sheet.columns.forEach((column) => {
      let maxLength = 0;
      if (typeof column.eachCell === 'function') {
        column.eachCell({ includeEmpty: true }, (cell) => {
          const cellValue = cell.value ? String(cell.value) : '';
          maxLength = Math.max(maxLength, cellValue.length);
        });
      }
      column.width = maxLength * 1.2;
    });
    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    });
  }

  async fileExcelCat9AndCat12(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
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
        return await this.getAllExcelFactoryCat9AndCat12Data(
          sheet,
          dateFrom,
          dateTo,
          factory,
        );
    }

    await getADataExcelFactoryCat9AndCat12(
      sheet,
      db,
      dateFrom,
      dateTo,
      factory,
      this.EIP,
    );
  }

  private async getAllExcelFactoryCat9AndCat12Data(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    dbEIP?: Sequelize,
  ) {
    const getQuery = (
      factoryName: string,
    ) => `SELECT CAST(ROW_NUMBER() OVER(ORDER BY [Date]) AS INT) AS [No]
                          ,*
                          ,'${factoryName}' AS Factory_Name
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
                                    ,pc.PortCode             AS Port_Of_Arrival
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
                                    ,ISNULL(id.ARTICLE, 'SAMPLE SHOE')              AS Article_ID
                                    ,is1.Qty                 AS Quantity
                                    ,is1.GW                  AS Gross_Weight
                                    ,im.CUSTID               AS Customer_ID
                                    ,'Truck'                 AS Local_Land_Transportation
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

    let where = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
    let where1 = 'WHERE 1=1 AND sb.CFMID IS NOT NULL';
    const replacements: any[] = [];

    if (dateFrom && dateTo) {
      where += ` AND CONVERT(VARCHAR ,sb.ExFty_Date,23) BETWEEN ? AND ?`;
      where1 += ` AND CONVERT(VARCHAR ,sb.ExFty_Date,23) BETWEEN ? AND ?`;
      replacements.push(dateFrom, dateTo, dateFrom, dateTo);
    }
    // const query = ``;
    const connects = [
      { conn: this.LYV_ERP, factoryName: 'LYV' },
      { conn: this.LHG_ERP, factoryName: 'LHG' },
      { conn: this.LYM_ERP, factoryName: 'LYM' },
      { conn: this.LVL_ERP, factoryName: 'LVL' },
      { conn: this.LYF_ERP, factoryName: 'LYF' },
      { conn: this.JAZ_ERP, factoryName: 'JAZ' },
      { conn: this.JZS_ERP, factoryName: 'JZS' },
    ];
    const dataResults = await Promise.all(
      connects.map(({ conn, factoryName }) => {
        const query = getQuery(factoryName);
        return conn.query(query, {
          type: QueryTypes.SELECT,
          replacements,
        });
      }),
    );

    const allData = dataResults.flat();

    let data: any = allData.map((item, index) => ({ ...item, No: index + 1 }));

    data = await Promise.all(
      data.map(async (item) => {
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
        const result: any = await dbEIP?.query(queryPoA, {
          type: QueryTypes.SELECT,
          replacements: [`%${item.Transport_Method}%`, `%${item.Customer_ID}%`],
        });

        return {
          ...item,
          Transport_Method: item.Transport_Method,
          Port_Of_Departure: portOfDeparture,
          Port_Of_Arrival: result[0]?.PortCode.trim().toUpperCase() || 'N/A',
        };
      }),
    );

    sheet.columns = [
      {
        header: 'No.',
        key: 'No',
      },
      {
        header: 'Date',
        key: 'Date',
      },
      {
        header: 'Shipment Date',
        key: 'Shipment_Date',
      },
      {
        header: 'Invoice Number',
        key: 'Invoice_Number',
      },
      {
        header: 'Article Name',
        key: 'Article_Name',
      },
      {
        header: 'Article ID',
        key: 'Article_ID',
      },
      {
        header: 'Quantity',
        key: 'Quantity',
      },
      {
        header: 'Gross Weight',
        key: 'Gross_Weight',
      },
      {
        header: 'Customer ID',
        key: 'Customer_ID',
      },
      {
        header: 'Local land transportation',
        key: 'Local_Land_Transportation',
      },
      {
        header: 'Port of Departure',
        key: 'Port_Of_Departure',
      },
      {
        header: 'Port of Arrival',
        key: 'Port_Of_Arrival',
      },
      {
        header: 'Land Transport Distance',
        key: 'Land_Transport_Distance',
      },
      {
        header: 'Sea Transport Distance',
        key: 'Sea_Transport_Distance',
      },
      {
        header: 'Air Transport Distance',
        key: 'Air_Transport_Distance',
      },
      {
        header: 'Transport Method',
        key: 'Transport_Method',
      },
      {
        header: 'Land Transport Ton-Kilometers',
        key: 'Land_Transport_Ton_Kilometers',
      },
      {
        header: 'Sea Transport Ton-Kilometers',
        key: 'Sea_Transport_Ton_Kilometers',
      },
      {
        header: 'Air Transport Ton-Kilometers',
        key: 'Air_Transport_Ton_Kilometers',
      },
    ];
    data.forEach((item) => sheet.addRow(item));
    sheet.columns.forEach((column) => {
      let maxLength = 0;
      if (typeof column.eachCell === 'function') {
        column.eachCell({ includeEmpty: true }, (cell) => {
          const cellValue = cell.value ? String(cell.value) : '';
          maxLength = Math.max(maxLength, cellValue.length);
        });
      }
      column.width = maxLength * 1.2;
    });
    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    });
  }

  async fileExcelCat5(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    // console.log(dateFrom, dateTo, factory);
    let db: Sequelize;
    switch (factory.trim()) {
      case 'LYV':
        db = this.LYV_WMS;
        break;
      case 'LHG':
        db = this.LHG_WMS;
        break;
      case 'LYM':
        db = this.LYM_WMS;
        break;
      case 'LVL':
        db = this.LVL_WMS;
        break;
      case 'JAZ':
        db = this.JAZ_WMS;
        break;
      case 'JZS':
        db = this.JZS_WMS;
        break;
      default:
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
    const query = `SELECT dwo.WASTE_DATE                     AS Waste_disposal_date
                          ,dwc.CONSOLIDATED_WASTE_CODE AS Consolidated_Waste
                          ,dwc.WASTE_CODE AS Waste_Code
                          ,dtv.TREATMENT_VENDOR_NAME          AS Vendor_Name
                          ,dtv.TREATMENT_VENDOR_ID            AS Vendor_ID
                          ,td.ADDRESS+'('+CONVERT(VARCHAR(5) ,td.DISTANCE)+'km)' AS Waste_collection_address
                          ,dwc.LOCATION_CODE AS Location_Code
                          ,CAST('0' AS INT)                   AS Transportation_Distance_km
                          ,CASE 
                                WHEN dwo.HAZARDOUS<>'N/A' THEN 'hazardous waste'
                                WHEN dwo.NON_HAZARDOUS<>'N/A' THEN 'Non-hazardous waste'
                                ELSE NULL
                          END                                AS The_type_of_waste
                          ,CASE 
                                WHEN dwo.HAZARDOUS<>'N/A' THEN dwo.HAZARDOUS
                                WHEN dwo.NON_HAZARDOUS<>'N/A' THEN dwo.NON_HAZARDOUS
                                ELSE NULL
                          END                                AS Waste_type
                          ,dtm.TREATMENT_METHOD_ENGLISH_NAME  AS Waste_Treatment_method
                          ,dtm.TREATMENT_METHOD_ID            AS Treatment_Method_ID
                          ,dwo.QUANTITY                       AS Weight_of_waste_treated_Unit_kg
                          ,CAST('0' AS INT)                   AS TKT_Ton_km
                    FROM   dbo.DATA_WASTE_OUTPUT_CUSTOMER dwo
                          LEFT JOIN dbo.DATA_TREATMENT_VENDOR dtv
                                ON  dtv.TREATMENT_VENDOR_ID = dwo.TREATMENT_SUPPLIER
                          LEFT JOIN dbo.DATA_TREATMENT_METHOD dtm
                                ON  dtm.TREATMENT_METHOD_ID = dwo.TREATMENT_METHOD_ID
                          LEFT JOIN dbo.DATA_EMERET_WASTE_CATEGORY dewc
                                ON  dewc.EMERET_WASTE_ID = dwo.EMERET_WASTE_ID
                          LEFT JOIN dbo.DATA_WASTE_CATEGORY dwc
                                ON  dwc.WASTE_CODE = dwo.WASTE_CODE
                          LEFT JOIN dbo.DATA_TREATMENT_DISTANCE td
                                ON  td.CODE = dwo.LOCATION_CODE
                    ${where}`;
    const connects = [
      this.LYV_WMS,
      this.LHG_WMS,
      this.LYM_WMS,
      this.LVL_WMS,
      this.JAZ_WMS,
      this.JZS_WMS,
    ];
    const dataResults = await Promise.all(
      connects.map((conn) => {
        return conn.query(query, {
          type: QueryTypes.SELECT,
          replacements,
        });
      }),
    );
    let data = dataResults.flat();
    sheet.columns = [
      {
        header: 'Waste disposal date',
        key: 'Waste_disposal_date',
      },
      {
        header: 'Consolidated Waste',
        key: 'Consolidated_Waste',
      },
      {
        header: 'Waste Code',
        key: 'Waste_Code',
      },
      {
        header: 'Vendor Name',
        key: 'Vendor_Name',
      },
      {
        header: 'Vendor ID',
        key: 'Vendor_ID',
      },
      {
        header: 'Waste collection address',
        key: 'Waste_collection_address',
      },
      {
        header: 'Location Code',
        key: 'Location_Code',
      },
      {
        header: 'Transportation Distance (km)',
        key: 'Transportation_Distance_km',
      },
      {
        header: '*The type of waste',
        key: 'The_type_of_waste',
      },
      {
        header: '*Waste type',
        key: 'Waste_type',
      },
      {
        header: '*Waste Treatment method',
        key: 'Waste_Treatment_method',
      },
      {
        header: 'Treatment Method ID',
        key: 'Treatment_Method_ID',
      },
      {
        header: '*Weight of waste treated (Unit：kg)',
        key: 'Weight_of_waste_treated_Unit_kg',
      },
      {
        header: 'TKT (Ton-km)',
        key: 'TKT_Ton_km',
      },
    ];
    data.forEach((item) => sheet.addRow(item));
    sheet.columns.forEach((column) => {
      let maxLength = 0;
      if (typeof column.eachCell === 'function') {
        column.eachCell({ includeEmpty: true }, (cell) => {
          const cellValue = cell.value ? String(cell.value) : '';
          maxLength = Math.max(maxLength, cellValue.length);
        });
      }
      column.width = maxLength * 1.2;
    });
    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    });
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
    const connects: { factoryName: string; conn: Sequelize }[] = [
      { factoryName: 'LYV', conn: this.LYV_ERP },
      { factoryName: 'LHG', conn: this.LHG_ERP },
      { factoryName: 'LVL', conn: this.LVL_ERP },
      { factoryName: 'LYM', conn: this.LYM_ERP },
    ];
    const replacements = {
      startDate: dateFrom,
      endDate: dayjs(dateTo).add(1, 'day').startOf('day').format('YYYY-MM-DD'),
      offset: 1,
      limit: 20,
    };
    const dataResults = await Promise.all(
      connects.map(async ({ factoryName, conn }) => {
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

        rows = await Promise.all(
          rows.map(async (item: any) => {
            const taxFreeZone: any[] = await this.EIP.query(
              `SELECT TaxFreeZoneAddress
              FROM CMW_TAX_FREE_ZONE_ADDRESS
              WHERE SupplierID = ? AND Factory = ?`,
              {
                replacements: [item.SupplierCode.trim(), factoryName],
                type: QueryTypes.SELECT,
              },
            );
            return {
              ...item,
              TransportationMethod:
                taxFreeZone.length > 0 ? 'Land' : item.TransportationMethod,
              Departure:
                taxFreeZone.length > 0
                  ? taxFreeZone[0].TaxFreeZoneAddress
                  : item.Departure,
            };
          }),
        );

        return rows;
      }),
    );
    let data = dataResults.flat();
    sheet.columns = [
      { header: 'No.', key: 'No' },
      { header: 'Factory Code', key: 'FactoryCode' },
      { header: 'Pur Date', key: 'PurDate' },
      { header: 'RK Date', key: 'RKDate' },
      { header: 'Purchase Order', key: 'PurNo' },
      { header: 'Received No.', key: 'ReceivedNo' },
      { header: 'Material No.', key: 'MatID' },
      { header: 'Qty.(Usage)', key: 'QtyUsage' },
      { header: 'Qty.(receive)', key: 'QtyReceive' },
      { header: 'Unit weight', key: 'UnitWeight' },
      { header: 'Weight (Unit：KG)', key: 'WeightUnitkg' },
      { header: 'Supplier Code', key: 'SupplierCode' },
      { header: 'Style', key: 'Style' },
      { header: 'Transportation Method', key: 'TransportationMethod' },
      { header: 'Departure', key: 'Departure' },
      {
        header: 'Third-country Land Transport (A)',
        key: 'ThirdCountryLandTransport',
      },
      { header: 'Port of Departure', key: 'PortOfDeparture' },
      { header: 'Port of Arrival', key: 'PortOfArrival' },
      {
        header: 'Factory (Domestic Land Transport)(B)',
        key: 'FactoryDomesticLandTransport',
      },
      { header: 'Destination', key: 'Destination' },
      {
        header: 'Land Transport Distance (A+B)',
        key: 'LandTransportDistance',
      },
      { header: 'Sea Transport Distance', key: 'SeaTransportDistance' },
      { header: 'Air Transport Distance', key: 'AirTransportDistance' },
      {
        header: 'Land Transport Ton-Kilometer',
        key: 'LandTransportTonKilometers',
      },
      {
        header: 'Sea Transport Ton-Kilometers',
        key: 'SeaTransportTonKilometers',
      },
      {
        header: '	Air Transport Ton-Kilomete',
        key: 'AirTransportTonKilometers',
      },
    ];
    data.forEach((item) => sheet.addRow(item));
    sheet.columns.forEach((column) => {
      let maxLength = 0;
      if (typeof column.eachCell === 'function') {
        column.eachCell({ includeEmpty: true }, (cell) => {
          const cellValue = cell.value ? String(cell.value) : '';
          maxLength = Math.max(maxLength, cellValue.length);
        });
      }
      column.width = maxLength * 1.2;
    });
    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    });
  }

  async fileExcelCat7(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
  ) {
    let db: Sequelize;
    switch (factory.trim()) {
      case 'LYV':
        db = this.LYV_HRIS;
        break;
      case 'LHG':
        db = this.LHG_HRIS;
        break;
      case 'LYM':
        db = this.LYM_HRIS;
        break;
      case 'LVL':
        db = this.LVL_HRIS;
        break;
      case 'JAZ':
        db = this.JAZ_HRIS;
        break;
      case 'JZS':
        db = this.JZS_HRIS;
        break;
      default:
        return await this.getAllDataExcelFactoryCat7(sheet, dateFrom, dateTo);
    }

    await getADataExcelFactoryCat7(sheet, db, dateFrom, dateTo, factory);
  }

  private async getAllDataExcelFactoryCat7(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
  ) {
    const connects: { facotryName: string; conn: Sequelize }[] = [
      { facotryName: 'LYV', conn: this.LYV_HRIS },
      { facotryName: 'LHG', conn: this.LHG_HRIS },
      { facotryName: 'LVL', conn: this.LVL_HRIS },
      { facotryName: 'LYM', conn: this.LYM_HRIS },
      { facotryName: 'JAZ', conn: this.JAZ_HRIS },
      { facotryName: 'JZS', conn: this.JZS_HRIS },
    ];

    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];

    const dataResults = await Promise.all(
      connects.map(({ facotryName, conn }) => {
        const { query } = buildQuery(facotryName, dateFrom, dateTo);
        return conn.query(query, { type: QueryTypes.SELECT, replacements });
      }),
    );

    let data = dataResults.flat();
    sheet.columns = [
      { header: 'Staff ID', key: 'Staff_ID' },
      { header: 'Residential address', key: 'Residential_address' },
      {
        header: 'Main transportation type',
        key: 'Main_transportation_type',
      },
      { header: 'km', key: 'km' },
      { header: 'Number of working days', key: 'Number_of_working_days' },
      { header: 'PKT (p-km)', key: 'PKT_p_km' },
    ];
    data.forEach((item) => sheet.addRow(item));

    sheet.columns.forEach((column) => {
      let maxLength = 0;
      if (typeof column.eachCell === 'function') {
        column.eachCell({ includeEmpty: true }, (cell) => {
          const cellValue = cell.value ? String(cell.value) : '';
          maxLength = Math.max(maxLength, cellValue.length);
        });
      }
      column.width = maxLength * 1.2;
    });

    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    });
  }

  private async fileCustomExport(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    factory: string,
    fields?: string[],
  ) {
    let db: Sequelize;
    switch (factory.trim()) {
      case 'LYV':
        db = this.LYV_HRIS;
        break;
      case 'LHG':
        db = this.LHG_HRIS;
        break;
      case 'LYM':
        db = this.LYM_HRIS;
        break;
      case 'LVL':
        db = this.LVL_HRIS;
        break;
      case 'JAZ':
        db = this.JAZ_HRIS;
        break;
      case 'JZS':
        db = this.JZS_HRIS;
        break;
      default:
        return await this.getAllDataCustomExport(
          sheet,
          dateFrom,
          dateTo,
          fields,
        );
    }
    await getADataCustomExport(sheet, db, dateFrom, dateTo, factory, fields);
  }

  private async getAllDataCustomExport(
    sheet: ExcelJS.Worksheet,
    dateFrom: string,
    dateTo: string,
    fields?: string[],
  ) {
    const connects: { facotryName: string; conn: Sequelize }[] = [
      { facotryName: 'LYV', conn: this.LYV_HRIS },
      { facotryName: 'LHG', conn: this.LHG_HRIS },
      { facotryName: 'LVL', conn: this.LVL_HRIS },
      { facotryName: 'LYM', conn: this.LYM_HRIS },
      { facotryName: 'JAZ', conn: this.JAZ_HRIS },
      { facotryName: 'JZS', conn: this.JZS_HRIS },
    ];
    const replacements = dateFrom && dateTo ? [dateFrom, dateTo] : [];
    const dataResults = await Promise.all(
      connects.map(({ facotryName, conn }) => {
        const { query } = buildQueryCustomExport(dateFrom, dateTo, facotryName);
        return conn.query(query, { type: QueryTypes.SELECT, replacements });
      }),
    );

    const allData = dataResults.flat();

    let data = allData.map((item, index) => ({ ...item, No: index + 1 }));

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

    // console.log(data);

    sheet.columns.forEach((column) => {
      let maxLength = 0;
      if (typeof column.eachCell === 'function') {
        column.eachCell({ includeEmpty: true }, (cell) => {
          const cellValue = cell.value ? String(cell.value) : '';
          maxLength = Math.max(maxLength, cellValue.length);
        });
      }
      column.width = maxLength * 1.2;
    });

    sheet.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    });
  }
}
